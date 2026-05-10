import asyncio
import base64
import io
import os
import re
from urllib.parse import quote

import fitz  # PyMuPDF
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
import anthropic
from PIL import Image as PILImage
from pydantic import BaseModel
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE

app = FastAPI(title="Paper Summarizer API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

client       = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))
async_client = anthropic.AsyncAnthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))

_BLACK = RGBColor(0x00, 0x00, 0x00)
_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
_NAVY  = RGBColor(0x1F, 0x49, 0x7D)
_LGRAY = RGBColor(0xF2, 0xF2, 0xF2)
_MGRAY = RGBColor(0x60, 0x60, 0x60)

# ── Prompts ───────────────────────────────────────────────────────────────────

SYSTEM_PROMPT = """あなたは医学論文の抄読会資料を作成する専門家です。
論文のIMRAD構造（Introduction, Methods, Results, Discussion）を正確に把握したうえで、
指定された5項目を日本語のMarkdown形式で出力してください。

出力ルール：
- 各項目は ## 見出しで始める
- 数値・統計は原文のまま記載する（改変しない）
- 情報が論文中に明示されていない場合は「記載なし」と書く
- 推測や補足は加えない"""


def build_prompt(paper_text: str) -> str:
    return f"""以下の医学論文テキストを読み、次の5項目をMarkdown形式で出力してください。

---

## 1. PICO

| 項目 | 内容 |
|------|------|
| **P** (Patient/Population) | 対象患者・集団 |
| **I** (Intervention) | 介入・曝露 |
| **C** (Comparison) | 比較対照 |
| **O** (Outcome) | 主要アウトカム |

上記の表を論文の内容で埋めてください。

---

## 2. この研究の新規性と重要性

（既存研究との差別化ポイント、なぜ今この研究が必要か）

---

## 3. 主要な結果（数値を含む）

（主要アウトカムの結果を、p値・信頼区間・効果量などの統計値とともに記載）

---

## 4. この論文の限界（Limitations）

（著者が認めている限界点を箇条書きで列挙）

---

## 5. 臨床現場でどう活かせるか（Clinical Implications）

（実際の診療・ケアへの応用可能性）

---

論文テキスト：
{paper_text[:15000]}
"""

# ── PDF utilities ─────────────────────────────────────────────────────────────

def extract_text(pdf_bytes: bytes) -> str:
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    return "\n".join(page.get_text() for page in doc)


# ── Image utilities ───────────────────────────────────────────────────────────

def _autocrop_white(img: PILImage.Image, threshold: int = 248, padding: int = 8) -> PILImage.Image:
    """Crop near-white margins to maximize figure area on slide."""
    mask = img.convert("L").point(lambda p: 255 if p < threshold else 0)
    bbox = mask.getbbox()
    if not bbox:
        return img
    w, h = img.size
    return img.crop((
        max(0, bbox[0] - padding),
        max(0, bbox[1] - padding),
        min(w, bbox[2] + padding),
        min(h, bbox[3] + padding),
    ))


def _is_real_figure(img: PILImage.Image) -> bool:
    """
    Return True if the image looks like a graph / photo / table.
    Reject text-heavy document fragments (mostly white + black with no midtones/color).
    """
    w, h = img.size

    # Very tall-and-narrow strips are usually column fragments, not figures
    if w / h < 0.25:
        return False

    # Use an 80×80 thumbnail for fast pixel analysis
    thumb = img.convert("RGB").resize((80, 80), PILImage.LANCZOS)
    gray  = thumb.convert("L")
    hist  = gray.histogram()          # 256 bins, total = 6400
    total = 80 * 80

    white_ratio = sum(hist[230:]) / total   # near-white  (230-255)
    dark_ratio  = sum(hist[:25])  / total   # near-black  (0-24)
    mid_ratio   = sum(hist[25:230]) / total  # midtones

    # Signature of a text/document image: lots of white, some black, barely any midtone
    if white_ratio > 0.82 and dark_ratio > 0.01 and mid_ratio < 0.18:
        # Allow if the image has significant colour (e.g. colour bar charts)
        def _wmean(ch_hist):
            return sum(i * c for i, c in enumerate(ch_hist)) / total

        r_m = _wmean(thumb.getchannel("R").histogram())
        g_m = _wmean(thumb.getchannel("G").histogram())
        b_m = _wmean(thumb.getchannel("B").histogram())
        color_spread = max(abs(r_m - g_m), abs(g_m - b_m), abs(r_m - b_m))

        if color_spread < 12:   # Essentially greyscale → text, reject
            return False

    return True


def _load_and_normalize(raw: bytes, max_dim: int = 1600) -> PILImage.Image | None:
    """Open raw image bytes, convert to RGB, downscale if too large."""
    try:
        img = PILImage.open(io.BytesIO(raw))
        if max(img.width, img.height) > max_dim:
            ratio = max_dim / max(img.width, img.height)
            img = img.resize(
                (int(img.width * ratio), int(img.height * ratio)),
                PILImage.LANCZOS,
            )
        if img.mode == "P":
            img = img.convert("RGBA")
        if img.mode in ("RGBA", "LA"):
            bg = PILImage.new("RGB", img.size, (255, 255, 255))
            bg.paste(img, mask=img.split()[-1])
            return bg
        if img.mode != "RGB":
            return img.convert("RGB")
        return img
    except Exception:
        return None


def _to_png_bytes(img: PILImage.Image) -> bytes:
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


_CAPTION_RE = re.compile(
    r'^(Fig\.?|Figure|Table|Tab\.?|FIG\.?|TABLE|FIGURE|Suppl\.?\s*Fig)',
    re.IGNORECASE,
)


def _find_caption(page, xref: int, text_blocks: list) -> str:
    """Return the Figure/Table legend nearest to this image on the page."""
    try:
        rects = page.get_image_rects(xref)
        if not rects:
            return ""
        ir = rects[0]   # image rect on the page

        best, best_dist = "", float("inf")
        for blk in text_blocks:
            if len(blk) < 5:
                continue
            bx0, by0, bx1, by1, text = blk[0], blk[1], blk[2], blk[3], blk[4]
            if not isinstance(text, str) or not text.strip():
                continue
            text = text.strip()
            if not _CAPTION_RE.match(text):
                continue

            dist_below = by0 - ir.y1   # positive → block is below image
            dist_above = ir.y0 - by1   # positive → block is above image

            in_below = -5 <= dist_below <= 120
            in_above = -5 <= dist_above <= 80
            if not (in_below or in_above):
                continue

            # Require horizontal overlap or near-alignment
            h_overlap = min(float(bx1), float(ir.x1)) - max(float(bx0), float(ir.x0))
            if h_overlap < -30:
                continue

            dist = dist_below if in_below else dist_above
            if dist < best_dist:
                best_dist = dist
                best = text[:400]

        return best
    except Exception:
        return ""


def extract_figures(pdf_bytes: bytes, max_figures: int = 10) -> list[dict]:
    doc  = fitz.open(stream=pdf_bytes, filetype="pdf")
    figs: list[dict] = []
    seen: set[int]   = set()

    for page_num, page in enumerate(doc, 1):
        text_blocks = page.get_text("blocks")
        page_w = page.rect.width
        page_h = page.rect.height
        page_area = page_w * page_h
        page_aspect = page_w / page_h if page_h > 0 else 1.0

        for img_info in page.get_images(full=True):
            xref = img_info[0]
            if xref in seen:
                continue
            seen.add(xref)

            try:
                # ── ① Page-level image guard (most important filter) ───────────
                # Check how much of the page this image covers.
                # Anything > 55 % is almost certainly a page scan, not a figure.
                try:
                    img_rects = page.get_image_rects(xref)
                    if img_rects:
                        ir = img_rects[0]
                        coverage = (ir.width * ir.height) / page_area
                        if coverage > 0.55:
                            continue
                        # Also reject very tall-narrow images (portrait page slice)
                        if ir.height > 0 and (ir.width / ir.height) < 0.30:
                            continue
                except Exception:
                    pass   # fall through to pixel-dimension check below

                # ── ② Embedded image dimension filter ────────────────────────
                base_img = doc.extract_image(xref)
                w, h = base_img["width"], base_img["height"]

                # Skip tiny images (icons, logos)
                if min(w, h) < 100 or w * h < 15_000:
                    continue

                # If pixel aspect ratio matches the page aspect ratio AND the
                # image is large → almost certainly a page scan
                img_aspect = w / h if h > 0 else 1.0
                if abs(img_aspect - page_aspect) < 0.08 and min(w, h) > 700:
                    continue

                # ── ③ Load & quality-filter ───────────────────────────────────
                img = _load_and_normalize(base_img["image"])
                if img is None:
                    continue

                if not _is_real_figure(img):
                    continue

                # ── ④ Crop white margins ──────────────────────────────────────
                img = _autocrop_white(img)

                # ── ⑤ Caption lookup ──────────────────────────────────────────
                caption = _find_caption(page, xref, text_blocks)

                figs.append({
                    "page_num":   page_num,
                    "figure_num": len(figs) + 1,
                    "png_bytes":  _to_png_bytes(img),
                    "caption":    caption,
                })
                if len(figs) >= max_figures:
                    return figs
            except Exception:
                continue

    return figs


async def _analyze_figure_async(png_bytes: bytes, caption: str = "") -> dict:
    """
    Multimodal analysis: returns {"title": str, "description": str}.
    Caption (detected from PDF) is fed as a hint so the model can reproduce
    the original Fig/Table numbering in the title.
    """
    b64 = base64.standard_b64encode(png_bytes).decode()
    caption_hint = f"\n\nPDF内のキャプション（参考）：{caption}" if caption else ""

    prompt = (
        "この図表を分析し、以下の形式で日本語のみで出力してください。\n\n"
        "タイトル: [50字以内。PDF内にFig/Table番号があればそれを含める。"
        "例：Fig 1. カプラン・マイヤー生存曲線、Table 2. 患者背景の比較]\n"
        "解説: [150字以内。グラフの種類・表示データ・主要な傾向または統計的結果]"
        + caption_hint
    )

    try:
        resp = await async_client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=500,
            messages=[{
                "role": "user",
                "content": [
                    {"type": "image", "source": {
                        "type": "base64", "media_type": "image/png", "data": b64,
                    }},
                    {"type": "text", "text": prompt},
                ],
            }],
        )
        raw = resp.content[0].text

        title, desc = "", ""
        for line in raw.splitlines():
            if re.match(r'タイトル[:：]', line):
                title = re.sub(r'^タイトル[:：]\s*', '', line).strip()[:60]
            elif re.match(r'解説[:：]', line):
                desc = re.sub(r'^解説[:：]\s*', '', line).strip()

        # Fallback: if parsing failed, use the whole response as description
        if not title:
            title = (caption.split('.')[0] + '.').strip() if caption else "図表"
        if not desc:
            desc = raw.strip()[:300]

        return {"title": title, "description": desc}
    except Exception:
        return {"title": caption.split('\n')[0][:60] if caption else "図表", "description": "図表の解析に失敗しました。"}

# ── PPTX helpers ──────────────────────────────────────────────────────────────

def _strip_md(text: str) -> str:
    text = re.sub(r'\*\*(.+?)\*\*', r'\1', text)
    text = re.sub(r'\*(.+?)\*', r'\1', text)
    text = re.sub(r'^- ', '• ', text, flags=re.MULTILINE)
    return text.strip()


def _parse_sections(md: str) -> list[dict]:
    parts = re.split(r'(?m)^## ', md)
    sections = []
    for part in parts:
        if not part.strip():
            continue
        nl      = part.find('\n')
        title   = part[:nl].strip() if nl != -1 else part.strip()
        content = part[nl + 1:].strip() if nl != -1 else ''
        content = re.sub(r'^\s*---\s*$', '', content, flags=re.MULTILINE)
        content = re.sub(r'\n{3,}', '\n\n', content).strip()
        sections.append({'title': title, 'content': content})
    return sections


def _parse_md_table(content: str) -> list[list[str]]:
    rows = []
    for line in content.splitlines():
        line = line.strip()
        if not (line.startswith('|') and line.endswith('|')):
            continue
        if re.match(r'^\|[\s\-|:]+\|$', line):
            continue
        cells = [c.strip() for c in line[1:-1].split('|')]
        cells = [re.sub(r'\*\*(.+?)\*\*', r'\1', c) for c in cells]
        cells = [re.sub(r'\*(.+?)\*', r'\1', c) for c in cells]
        rows.append(cells)
    return rows


def _add_rect(slide, left, top, width, height, rgb: RGBColor):
    s = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, left, top, width, height)
    s.fill.solid()
    s.fill.fore_color.rgb = rgb
    s.line.fill.background()
    return s


def _add_title_bar(slide, slide_w, title: str, slide_num: int, total: int):
    bar = _add_rect(slide, 0, 0, slide_w, Inches(1.05), _NAVY)
    tf  = bar.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    r = p.add_run()
    r.text = title
    r.font.size = Pt(22)
    r.font.bold = True
    r.font.color.rgb = _WHITE

    nb = slide.shapes.add_textbox(slide_w - Inches(1.3), Inches(0.05), Inches(1.2), Inches(0.45))
    nt = nb.text_frame.paragraphs[0]
    nt.alignment = PP_ALIGN.RIGHT
    nr = nt.add_run()
    nr.text = f"{slide_num} / {total}"
    nr.font.size = Pt(11)
    nr.font.color.rgb = _WHITE


def _add_text_content(slide, content: str, top, slide_w, avail_h, font_pt: int = 15):
    pad = Inches(0.45)
    tb  = slide.shapes.add_textbox(pad, top, slide_w - pad * 2, avail_h)
    tf  = tb.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    first = True
    for line in _strip_md(content).split('\n'):
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        r = p.add_run()
        r.text = line
        r.font.size = Pt(font_pt)
        r.font.color.rgb = _BLACK


def _add_table_content(slide, rows: list[list[str]], top, slide_w, max_h):
    if not rows:
        return
    n_rows, n_cols = len(rows), len(rows[0])
    tbl_h = min(max_h, Inches(0.56) * n_rows)
    pad   = Inches(0.45)
    tbl_w = slide_w - pad * 2

    tbl = slide.shapes.add_table(n_rows, n_cols, pad, top, tbl_w, tbl_h).table
    if n_cols == 2:
        tbl.columns[0].width = int(tbl_w * 0.38)
        tbl.columns[1].width = int(tbl_w * 0.62)

    for r_idx, row in enumerate(rows):
        for c_idx, cell_text in enumerate(row[:n_cols]):
            cell = tbl.cell(r_idx, c_idx)
            cell.text_frame.word_wrap = True
            p   = cell.text_frame.paragraphs[0]
            run = p.add_run()
            run.text = cell_text
            run.font.size = Pt(14)
            run.font.color.rgb = _BLACK
            if r_idx == 0:
                run.font.bold = True
                cell.fill.solid()
                cell.fill.fore_color.rgb = _LGRAY


def _add_section_content(slide, content: str, slide_w, slide_h):
    top     = Inches(1.2)
    avail_h = slide_h - top - Inches(0.15)
    has_tbl = bool(re.search(r'^\|.+\|$', content, re.MULTILINE))

    if has_tbl:
        rows      = _parse_md_table(content)
        non_table = re.sub(r'^\|.*\|$', '', content, flags=re.MULTILINE)
        non_table = re.sub(r'\n{2,}', '\n', non_table).strip()
        tbl_h     = min(Inches(4.8), Inches(0.56) * len(rows))
        _add_table_content(slide, rows, top, slide_w, tbl_h)
        if non_table:
            _add_text_content(slide, non_table, top + tbl_h + Inches(0.15), slide_w, avail_h - tbl_h, 14)
    else:
        _add_text_content(slide, content, top, slide_w, avail_h)


def _compute_fit(orig_w: int, orig_h: int, max_w: int, max_h: int) -> tuple[int, int]:
    if orig_w <= 0 or orig_h <= 0:
        return max_w, max_h
    scale = min(max_w / orig_w, max_h / orig_h)
    return int(orig_w * scale), int(orig_h * scale)


def _add_figure_slide(prs, slide, png_bytes: bytes, title: str, description: str,
                      slide_num: int, total: int):
    """Left 55%: image stretched to fill the full area. Right 43%: description."""
    w, h = prs.slide_width, prs.slide_height
    _add_rect(slide, 0, 0, w, h, _WHITE)
    _add_title_bar(slide, w, title, slide_num, total)

    content_top = Inches(1.15)
    content_h   = h - content_top - Inches(0.05)

    # ── Left 55 %: stretch image to fill every pixel of the left panel ───────
    img_area_w = int(w * 0.55)
    slide.shapes.add_picture(
        io.BytesIO(png_bytes),
        0, content_top,           # flush to left edge
        img_area_w, content_h,    # explicit width + height → fills area completely
    )

    # ── Right 43 %: description ───────────────────────────────────────────────
    txt_left = img_area_w + Inches(0.18)
    txt_w    = w - txt_left - Inches(0.1)
    tb = slide.shapes.add_textbox(txt_left, content_top + Inches(0.1), txt_w, content_h - Inches(0.2))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = description
    r.font.size = Pt(16)
    r.font.color.rgb = _BLACK


def _new_prs():
    prs = Presentation()
    prs.slide_width  = Inches(13.333)
    prs.slide_height = Inches(7.5)
    return prs, prs.slide_layouts[6]


def _add_title_slide(prs, blank, source_filename: str):
    w, h = prs.slide_width, prs.slide_height
    ts   = prs.slides.add_slide(blank)
    _add_rect(ts, 0, 0, w, h, _WHITE)

    tb = ts.shapes.add_textbox(Inches(1.5), Inches(2.2), w - Inches(3), Inches(2.2))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = "抄読会資料"
    r.font.size = Pt(44)
    r.font.bold = True
    r.font.color.rgb = _NAVY

    if source_filename:
        p2 = tf.add_paragraph()
        p2.alignment = PP_ALIGN.CENTER
        r2 = p2.add_run()
        r2.text = source_filename.replace('.pdf', '')
        r2.font.size = Pt(20)
        r2.font.color.rgb = _MGRAY

    _add_rect(ts, Inches(1.5), Inches(4.7), w - Inches(3), Pt(3), _NAVY)


def _add_section_slides(prs, blank, sections: list[dict], total: int):
    w, h = prs.slide_width, prs.slide_height
    for i, sec in enumerate(sections, 1):
        sl = prs.slides.add_slide(blank)
        _add_rect(sl, 0, 0, w, h, _WHITE)
        _add_title_bar(sl, w, sec['title'], i, total)
        _add_section_content(sl, sec['content'], w, h)


def generate_pptx(markdown: str, source_filename: str = "") -> bytes:
    sections = _parse_sections(markdown)
    total    = 1 + len(sections)
    prs, blank = _new_prs()
    _add_title_slide(prs, blank, source_filename)
    _add_section_slides(prs, blank, sections, total)
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


def generate_pptx_with_figures(markdown: str, figures: list[dict], source_filename: str = "") -> bytes:
    sections   = _parse_sections(markdown)
    n_sec, n_fig = len(sections), len(figures)
    total      = 1 + n_sec + n_fig
    prs, blank = _new_prs()
    _add_title_slide(prs, blank, source_filename)
    _add_section_slides(prs, blank, sections, total)

    for j, fig in enumerate(figures, 1):
        sl = prs.slides.add_slide(blank)
        _add_figure_slide(
            prs, sl,
            fig["png_bytes"],
            fig["title"],
            fig["description"],
            1 + n_sec + j, total,
        )

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()

# ── API endpoints ──────────────────────────────────────────────────────────────

@app.get("/health")
def health():
    return {"status": "ok", "message": "バックエンド正常稼働中"}


@app.post("/summarize")
async def summarize(file: UploadFile = File(...)):
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="PDFファイルのみ対応しています")
    pdf_bytes = await file.read()
    try:
        paper_text = extract_text(pdf_bytes)
    except Exception as e:
        raise HTTPException(status_code=422, detail=f"PDF読み込みエラー: {e}")
    if not paper_text.strip():
        raise HTTPException(status_code=422, detail="PDFからテキストを抽出できませんでした（スキャンPDFは非対応）")
    try:
        response = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=4096,
            system=SYSTEM_PROMPT,
            messages=[{"role": "user", "content": build_prompt(paper_text)}],
        )
        summary = response.content[0].text
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"AI生成エラー: {e}")
    return {"filename": file.filename, "summary": summary}


class ExportRequest(BaseModel):
    summary:  str
    filename: str = "論文"


@app.post("/export/pptx")
async def export_pptx(req: ExportRequest):
    try:
        pptx_bytes = generate_pptx(req.summary, req.filename)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"PPTX生成エラー: {e}")
    stem = req.filename.replace('.pdf', '')
    dl   = f"{stem}_抄読会資料.pptx"
    return StreamingResponse(
        io.BytesIO(pptx_bytes),
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{quote(dl)}"},
    )


@app.post("/export/pptx-with-figures")
async def export_pptx_with_figures(
    file:     UploadFile = File(...),
    summary:  str        = Form(...),
    filename: str        = Form("論文"),
):
    pdf_bytes = await file.read()

    try:
        raw_figures = extract_figures(pdf_bytes)
    except Exception as e:
        raise HTTPException(status_code=422, detail=f"図表抽出エラー: {e}")

    if not raw_figures:
        raise HTTPException(
            status_code=422,
            detail="PDFから抽出可能な図表が見つかりませんでした（ベクター形式の図は非対応です）",
        )

    # Parallel vision analysis — pass caption extracted from PDF
    analyses = await asyncio.gather(
        *[_analyze_figure_async(f["png_bytes"], f["caption"]) for f in raw_figures]
    )
    analyzed = [{**fig, **ana} for fig, ana in zip(raw_figures, analyses)]

    try:
        pptx_bytes = generate_pptx_with_figures(summary, analyzed, filename)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"PPTX生成エラー: {e}")

    stem = filename.replace('.pdf', '')
    dl   = f"{stem}_抄読会資料_図表付.pptx"
    return StreamingResponse(
        io.BytesIO(pptx_bytes),
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{quote(dl)}"},
    )
