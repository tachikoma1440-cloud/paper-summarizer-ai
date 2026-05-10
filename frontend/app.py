import streamlit as st
import requests

BACKEND_URL = "http://backend:8000"
PPTX_MIME   = "application/vnd.openxmlformats-officedocument.presentationml.presentation"

st.set_page_config(page_title="抄読会資料ジェネレーター", page_icon="📄", layout="wide")
st.title("📄 医学論文 抄読会資料ジェネレーター")
st.caption("PDFをアップロードすると、PICO・新規性・結果・限界・臨床応用を自動生成します。")

for k in ["summary", "filename", "pdf_bytes", "pptx_bytes", "fig_pptx_bytes"]:
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
