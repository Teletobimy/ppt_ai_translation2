## PPT Translation (Streamlit)

웹에서 PPTX 파일을 업로드하고, 서식을 최대한 보존하며 원하는 언어/톤으로 번역합니다. 중국어 번역의 경우 DeepSeek를 선택적으로 사용합니다.

### 로컬 실행

1) Python 3.10+ 환경 준비 후 의존성 설치:

```bash
pip install -r requirements.txt
```

2) 환경 변수로 키 설정 (또는 Streamlit secrets 사용):

```bash
set OPENAI_API_KEY=YOUR_OPENAI_API_KEY
# 선택: 중국어 번역 시 DeepSeek 사용 시
# set DEEPSEEK_API_KEY=YOUR_DEEPSEEK_API_KEY
```

3) 앱 실행:

```bash
streamlit run streamlit_app.py
```

### Streamlit Community Cloud 배포

1) 이 저장소를 GitHub에 푸시합니다.

2) Streamlit Cloud에서 New app → GitHub 저장소 선택 → `streamlit_app.py`를 엔트리 파일로 지정합니다.

3) Secrets에 다음 키를 추가합니다:

- `OPENAI_API_KEY`: OpenAI API 키 (필수)
- `DEEPSEEK_API_KEY`: DeepSeek API 키 (선택, 중국어 번역 시 사용)

예시 값은 `.streamlit/secrets.toml.example`를 참고하세요. 실제 `secrets.toml` 파일은 커밋하지 마세요.

### 참고

- 원본 데스크톱용 UI(Tkinter)는 `PPT_Language_Change.py`에 남아 있으며, 웹 앱에서는 `streamlit_app.py`의 업로드/옵션 UI를 사용합니다.
- 키는 코드에 하드코딩되지 않고 환경 변수(또는 Streamlit secrets)를 통해 주입됩니다.

