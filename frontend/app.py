import streamlit as st
import requests

BACKEND_URL = "http://backend:8000"
PPTX_MIME   = "application/vnd.openxmlformats-officedocument.presentationml.presentation"

st.set_page_config(page_title="抄読会資料ジェネレーター", page_icon="📄", layout="wide")
st.title("📄 医学論文 抄読会資料ジェネレーター")
st.caption("PDFをアップロードすると、PICO・新規性・結果・限界・臨床応用を自動生成します。")

for k in ["summary", "filename", "pdf_bytes", "pptx_bytes", "fig_pptx_bytes", "progress_pptx_bytes"]:
    if k not in st.session_state:
        st.session_state[k] = None

uploaded = st.file_uploader("論文PDFをアップロード", type=["pdf"])

if uploaded:
    # Reset state when a new file is loaded
    if st.session_state.filename != uploaded.name:
        for k in ["summary", "pptx_bytes", "fig_pptx_bytes"]:
            st.session_state[k] = None
        st.session_state.filename  = uploaded.name
        st.session_state.pdf_bytes = uploaded.getvalue()

    st.info(f"ファイル: {uploaded.name}")

    if st.button("抄読会資料を生成する", type="primary"):
        with st.spinner("AIが論文を解析中です…（30〜60秒かかります）"):
            try:
                res = requests.post(
                    f"{BACKEND_URL}/summarize",
                    files={"file": (uploaded.name, uploaded.getvalue(), "application/pdf")},
                    timeout=120,
                )
                res.raise_for_status()
                summary = res.json()["summary"]

                pptx_res = requests.post(
                    f"{BACKEND_URL}/export/pptx",
                    json={"summary": summary, "filename": uploaded.name},
                    timeout=30,
                )
                pptx_res.raise_for_status()

                st.session_state.summary    = summary
                st.session_state.pptx_bytes = pptx_res.content
                st.session_state.pdf_bytes  = uploaded.getvalue()
                st.session_state.filename   = uploaded.name
                st.session_state.fig_pptx_bytes = None

            except requests.exceptions.ConnectionError:
                st.error("バックエンドに接続できません。")
            except requests.exceptions.HTTPError as e:
                detail = e.response.json().get("detail", str(e)) if e.response else str(e)
                st.error(f"エラー: {detail}")
            except Exception as e:
                st.error(f"予期しないエラー: {e}")

# ── Results (persisted via session_state) ─────────────────────────────────────
if st.session_state.summary:
    st.success("生成完了！")
    st.divider()
    st.markdown(st.session_state.summary)
    st.divider()

    col1, col2 = st.columns(2)
    with col1:
        st.download_button(
            label="📄 Markdownとしてダウンロード",
            data=st.session_state.summary,
            file_name=st.session_state.filename.replace(".pdf", "_抄読会資料.md"),
            mime="text/markdown",
        )
    with col2:
        st.download_button(
            label="📊 PowerPointとしてダウンロード (.pptx)",
            data=st.session_state.pptx_bytes,
            file_name=st.session_state.filename.replace(".pdf", "_抄読会資料.pptx"),
            mime=PPTX_MIME,
        )

    st.divider()
    st.subheader("🖼️ 図表入りPowerPoint")
    st.caption(
        "PDFから図表を自動抽出し、AIによる解説を付けたスライドを追加します。"
        "図表の数に応じて1〜3分かかります（最大10枚を並列解析）。"
    )

    if st.session_state.fig_pptx_bytes:
        st.success("図表入りPPTXの生成が完了しました！")
        st.download_button(
            label="🖼️ 図表入りPowerPointをダウンロード (.pptx)",
            data=st.session_state.fig_pptx_bytes,
            file_name=st.session_state.filename.replace(".pdf", "_抄読会資料_図表付.pptx"),
            mime=PPTX_MIME,
        )
    else:
        if st.button("🖼️ 図表入りPPTXを生成する", type="secondary"):
            with st.spinner("図表を抽出・解析中です…（1〜3分かかる場合があります）"):
                try:
                    fig_res = requests.post(
                        f"{BACKEND_URL}/export/pptx-with-figures",
                        files={"file": (
                            st.session_state.filename,
                            st.session_state.pdf_bytes,
                            "application/pdf",
                        )},
                        data={
                            "summary":  st.session_state.summary,
                            "filename": st.session_state.filename,
                        },
                        timeout=300,
                    )
                    fig_res.raise_for_status()
                    st.session_state.fig_pptx_bytes = fig_res.content
                    st.rerun()
                except requests.exceptions.HTTPError as e:
                    detail = e.response.json().get("detail", str(e)) if e.response else str(e)
                    st.error(f"エラー: {detail}")
                except Exception as e:
                    st.error(f"予期しないエラー: {e}")

# ── Project tools ─────────────────────────────────────────────────────────────
st.divider()
st.subheader("🗂️ プロジェクト管理ツール")

pt_col1, pt_col2, pt_col3 = st.columns(3)

with pt_col1:
    st.markdown("**📝 Obsidian 論文要約保存**")
    if st.session_state.summary:
        if st.button("論文要約をObsidianに保存", type="secondary", key="btn_obs_note"):
            with st.spinner("Obsidianに書き込み中…"):
                try:
                    res = requests.post(
                        f"{BACKEND_URL}/export/obsidian-note",
                        json={
                            "summary":  st.session_state.summary,
                            "filename": st.session_state.filename or "論文",
                        },
                        timeout=10,
                    )
                    res.raise_for_status()
                    st.success(f"保存完了: {res.json().get('note_name', '')}")
                except requests.exceptions.HTTPError as e:
                    detail = e.response.json().get("detail", str(e)) if e.response else str(e)
                    st.error(f"エラー: {detail}")
                except Exception as e:
                    st.error(f"エラー: {e}")
    else:
        st.caption("論文要約を生成してから使用できます")

with pt_col2:
    st.markdown("**🔄 Obsidian 開発ログ同期**")
    if st.button("dev_log をObsidianに同期", type="secondary", key="btn_sync_devlog"):
        with st.spinner("同期中…"):
            try:
                res = requests.post(f"{BACKEND_URL}/sync/devlog", timeout=10)
                res.raise_for_status()
                st.success("同期完了")
            except requests.exceptions.HTTPError as e:
                detail = e.response.json().get("detail", str(e)) if e.response else str(e)
                st.error(f"エラー: {detail}")
            except Exception as e:
                st.error(f"エラー: {e}")

with pt_col3:
    st.markdown("**📊 進捗報告PPTX**")
    if st.session_state.progress_pptx_bytes:
        st.success("生成完了！")
        st.download_button(
            label="📥 進捗報告PPTXをダウンロード",
            data=st.session_state.progress_pptx_bytes,
            file_name="進捗報告.pptx",
            mime=PPTX_MIME,
            key="dl_progress_pptx",
        )
        if st.button("再生成", key="btn_regen_progress"):
            st.session_state.progress_pptx_bytes = None
            st.rerun()
    else:
        if st.button("進捗報告PPTXを生成", type="secondary", key="btn_gen_progress"):
            with st.spinner("AIが開発ログを解析してスライドを生成中…（30〜60秒）"):
                try:
                    res = requests.post(
                        f"{BACKEND_URL}/export/progress-pptx", timeout=120
                    )
                    res.raise_for_status()
                    st.session_state.progress_pptx_bytes = res.content
                    st.rerun()
                except requests.exceptions.HTTPError as e:
                    detail = e.response.json().get("detail", str(e)) if e.response else str(e)
                    st.error(f"エラー: {detail}")
                except Exception as e:
                    st.error(f"エラー: {e}")
