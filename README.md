# YGA Slide Formatter

영어 문제 원문을 붙여 넣으면 PPTX 슬라이드를 생성하는 Streamlit 앱입니다.

## Local Run

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Google Login Setup

이 앱은 Streamlit OIDC 로그인으로 Google 계정을 인증합니다. 공식 문서 기준으로 `st.login()`, `st.user`, `st.logout()`과 `.streamlit/secrets.toml`의 `[auth]` 설정을 사용합니다.

필수 환경변수:

```bash
YGA_REQUIRE_LOGIN=1
STREAMLIT_AUTH_REDIRECT_URI=http://localhost:8501/oauth2callback
STREAMLIT_AUTH_COOKIE_SECRET=long-random-secret
STREAMLIT_AUTH_GOOGLE_CLIENT_ID=your-google-client-id
STREAMLIT_AUTH_GOOGLE_CLIENT_SECRET=your-google-client-secret
STREAMLIT_AUTH_GOOGLE_SERVER_METADATA_URL=https://accounts.google.com/.well-known/openid-configuration
```

로컬에서는 위 값을 `.env`에 두고 실행하면 되고, Render에서는 같은 키를 환경변수로 넣으면 됩니다. `bootstrap_streamlit_auth.py`가 시작 전에 `.streamlit/secrets.toml`을 자동 생성합니다.

## Allowed Emails

허용 사용자 제한은 두 방식 중 하나 또는 둘 다 사용할 수 있습니다.

```bash
YGA_ALLOWED_EMAILS=teacher@example.com,admin@example.com
YGA_ALLOWED_EMAIL_DOMAINS=school.ac.kr,example.org
```

- `YGA_ALLOWED_EMAILS`: 정확히 허용할 이메일 목록
- `YGA_ALLOWED_EMAIL_DOMAINS`: 도메인 단위 허용 목록

둘 다 비워 두면 로그인만 성공한 모든 Google 계정을 허용합니다.

## Google Cloud Setup

Google Auth Platform에서 `Web application` 클라이언트를 만들고, `Authorized redirect URIs`에 아래 주소를 등록해야 합니다.

- 로컬: `http://localhost:8501/oauth2callback`
- Render: `https://<your-render-domain>.onrender.com/oauth2callback`

Google 앱이 `Testing` 상태면 테스트 사용자만 로그인할 수 있습니다. 공개 사용 전에는 `Audience`에서 필요한 계정을 추가하거나, `Published`로 전환하세요.

## Render Deploy

이 저장소에는 [`render.yaml`](/Users/hanhyemin/Desktop/make_yga/render.yaml)이 포함되어 있습니다.

1. GitHub 저장소를 Render에 연결합니다.
2. Blueprint를 생성합니다.
3. 아래 환경변수를 Render에 입력합니다.
   - `OPENAI_API_KEY`
   - `STREAMLIT_AUTH_REDIRECT_URI`
   - `STREAMLIT_AUTH_COOKIE_SECRET`
   - `STREAMLIT_AUTH_GOOGLE_CLIENT_ID`
   - `STREAMLIT_AUTH_GOOGLE_CLIENT_SECRET`
   - 필요시 `YGA_ALLOWED_EMAILS` 또는 `YGA_ALLOWED_EMAIL_DOMAINS`
4. 배포 후 Render 주소를 Google 클라이언트의 redirect URI에도 추가합니다.
