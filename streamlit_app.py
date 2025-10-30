import os
import tempfile
from io import BytesIO

import streamlit as st

from PPT_Language_Change import (
    translate_presentation,
    LANG_OPTIONS,
    TONE_OPTIONS,
)


def ensure_env_from_secrets() -> None:
    """Populate os.environ from Streamlit secrets if present."""
    if "OPENAI_API_KEY" in st.secrets:
        os.environ.setdefault("OPENAI_API_KEY", st.secrets["OPENAI_API_KEY"])
    if "DEEPSEEK_API_KEY" in st.secrets:
        os.environ.setdefault("DEEPSEEK_API_KEY", st.secrets["DEEPSEEK_API_KEY"])


def save_uploaded_to_tmp(uploaded_file) -> str:
    """Save the uploaded PPTX to a temp file and return its path."""
    suffix = os.path.splitext(uploaded_file.name)[1] or ".pptx"
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(uploaded_file.getbuffer())
        return tmp.name


st.set_page_config(page_title="PPT Translator", page_icon="🗂️", layout="centered")
st.title("PPT Translator (Streamlit)")
st.caption("서식 보존 · 톤 선택 · 중국어는 DeepSeek 지원")

ensure_env_from_secrets()

uploaded = st.file_uploader("번역할 PPTX 파일 업로드", type=["pptx"])

col1, col2 = st.columns(2)
with col1:
    target_lang = st.selectbox("번역 대상 언어", options=LANG_OPTIONS, index=0)
with col2:
    tone = st.selectbox("번역 톤", options=TONE_OPTIONS, index=0)

use_deepseek = False
if "Chinese" in target_lang:
    use_deepseek = st.checkbox("중국어 번역 시 DeepSeek 사용 (권장)", value=True)

font_scale = st.slider("번역 후 폰트 크기 배율(%)", min_value=50, max_value=300, value=100, step=5)

run = st.button("번역 시작")

if run:
    if not uploaded:
        st.warning("PPTX 파일을 업로드하세요.")
        st.stop()

    try:
        src_path = save_uploaded_to_tmp(uploaded)
        with st.status("번역 중...", expanded=True) as status:
            st.write("파일 처리 및 모델 호출 중...")
            prog = st.progress(0)

            def on_progress(done, total, msg):
                try:
                    if total > 0:
                        pct = int(max(0, min(100, (done / total) * 100)))
                        prog.progress(pct)
                    if msg:
                        status.update(label=f"번역 중... {msg}")
                except Exception:
                    pass

            out_path = translate_presentation(
                src_path,
                target_lang=target_lang,
                tone=tone,
                use_deepseek=use_deepseek,
                font_scale_percent=font_scale,
                on_progress=on_progress,
            )
            status.update(label="번역 완료", state="complete")

        with open(out_path, "rb") as f:
            out_bytes = f.read()
        out_name = os.path.basename(out_path)
        st.success("번역이 완료되었습니다. 아래 버튼으로 다운로드하세요.")
        st.download_button(
            label="번역된 PPTX 다운로드",
            data=out_bytes,
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
    except Exception as e:
        st.error(f"오류가 발생했습니다: {e}")

