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

# 커스텀 프롬프트 입력 영역 (커스텀 프롬프트 선택 시에만 표시)
custom_prompt = ""
if tone == "커스텀 프롬프트":
    with st.container():
        st.markdown("---")
        template_example = """#역할
전문 [B언어] 번역가로서, 사용자가 입력한 모든 [A언어] 문장을 정확하고 자연스러운 [B언어]로 번역합니다.

##주요 특징
정확성: 프레젠테이션, 보고서, 비즈니스 문서 등에 적합한 공식적이고 세련된 표현 사용
원어민이 봤을 때 절대 어색하지 않은 번역

문맥 고려: 문장의 의미와 뉘앙스를 세밀하게 분석하여 적절한 표현으로 번역
의미가 모호하거나 여러 해석이 가능한 경우, 사용자에게 반드시 확인 후 번역

자연스러움 유지: 원문의 의도와 어조를 유지하되, [B언어]에서 자연스럽게 들리도록 문장 구조 조정 가능

브랜드의 표기: [브랜드명]은 [영어 브랜드명]을 사용하며 [B언어]로 번역하지 않고 [영어 브랜드명] 유지

스타일 조정 가능: 사용자의 피드백에 따라 격식체, 반격식체, 발표체 등 스타일을 즉시 조정

###제한 사항
번역 이외의 불필요한 설명 금지
창의적 재해석 없이 원문에 충실한 번역 수행

####검토
번역 완료 후 재 검토하여 원어민이 봤을 때 어색한 부분이 있는지 검토하여 재 수정 하여 최종 번역본 출력

참고: [A언어]는 자동으로 원본 언어(한국어)로, [B언어]는 대상 언어로 치환됩니다.
마커 [[P#]]와 [[R#]]는 절대 변경하지 마세요."""
        
        custom_prompt = st.text_area(
            "커스텀 프롬프트 입력:",
            value=template_example,
            height=400,
            help="💡 팁: [A언어]는 원본 언어(한국어), [B언어]는 대상 언어로 자동 치환됩니다. [[P#]]와 [[R#]] 마커는 반드시 유지하세요.",
        )
        if not custom_prompt.strip():
            st.warning("커스텀 프롬프트를 입력하세요.")

use_deepseek = False
if "Chinese" in target_lang:
    use_deepseek = st.checkbox("중국어 번역 시 DeepSeek 사용 (권장)", value=True)

font_scale = st.slider("번역 후 폰트 크기 배율(%)", min_value=50, max_value=300, value=100, step=5)

with st.expander("고유 명사/용어집 (선택)"):
    st.markdown("입력 형식: 한 줄당 `원문 - 번역어` 형태. 예: `리쥬부스터 - rejuvuster`")
    glossary_text = st.text_area(
        "용어집",
        value="",
        height=120,
        placeholder="피더란 - PYDERIN\n리쥬부스터 - rejuvuster",
    )

def parse_glossary(text: str) -> dict:
    result = {}
    for raw in (text or "").splitlines():
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        # split on first hyphen
        if "-" in line:
            src, tgt = line.split("-", 1)
            src = src.strip()
            tgt = tgt.strip()
            if src and tgt:
                result[src] = tgt
    return result

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

            glossary = parse_glossary(glossary_text)
            out_path = translate_presentation(
                src_path,
                target_lang=target_lang,
                tone=tone,
                use_deepseek=use_deepseek,
                font_scale_percent=font_scale,
                on_progress=on_progress,
                glossary=glossary if glossary else None,
                custom_prompt=custom_prompt if tone == "커스텀 프롬프트" else "",
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

