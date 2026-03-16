# YGA Slide Formatter

영어 교재 원문을 붙여 넣으면, 문제 유형을 분석해 수업용 PPTX 슬라이드로 자동 변환하는 Streamlit 기반 웹 앱입니다.

이 프로젝트는 단순한 텍스트 복붙 도구가 아니라, 실제 수업 자료 제작 흐름에 맞춰 아래 작업을 한 번에 처리합니다.

- 문제 묶음 분리
- 문제 유형 판정
- 지문 / 문제 / 선지 구조화
- 슬라이드 페이지네이션
- PPTX 파일 생성 및 다운로드

## What It Does

YGA Slide Formatter는 이런 상황을 위한 도구입니다.

- 영어 문제를 교재 PDF나 문서에서 복사해 왔는데 PPT로 다시 옮기기 번거로운 경우
- 여러 문제를 한 번에 넣고 순서대로 슬라이드를 만들고 싶은 경우
- 문제 유형에 따라 지문과 선지 배치를 다르게 적용하고 싶은 경우
- 웹에서 바로 로그인 후 사용하고 싶은 경우

## Core Features

- 여러 문제를 한 번에 입력해도 문제 단위로 분리
- AI를 이용한 문제 유형 판정
- 문제 유형별 PPT 조합 로직 분기
- 문제 번호 자동 재부여
- 문제 본문 줄바꿈 정리
- Google 로그인 기반 접근 제어
- 허용 이메일 / 도메인 제한 지원
- Render 배포 지원

## Tech Stack

| Area | Stack |
| --- | --- |
| UI | Streamlit |
| PPT 생성 | Python + PPTX/XML 조합 로직 |
| 문제 분석 | OpenAI API |
| 인증 | Streamlit OIDC + Google Login |
| 배포 | Render |

## Project Structure

```text
make_yga/
├─ app.py                     # Streamlit UI 진입점
├─ slide_formatter.py         # 문제 파싱 / 페이지 분할 / PPT 조합
├─ ai_parser.py               # OpenAI 기반 문제 유형 판정
├─ auth_support.py            # 로그인 / 허용 이메일 검사 / auth secrets 생성
├─ bootstrap_streamlit_auth.py# 실행 전 Streamlit auth 설정 준비
├─ runtime_logging.py         # 런타임 로그 유틸
├─ render.yaml                # Render 배포 설정
├─ requirements.txt           # Python 의존성
└─ .env.example               # 환경변수 예시
```

## Quick Start

### 1. Install

```bash
pip install -r requirements.txt
```

### 2. Configure

로컬 실행 전 [`.env.example`](.env.example)를 참고해 `.env`를 채웁니다.

최소 예시는 아래와 같습니다.

```bash
OPENAI_API_KEY=sk-your-key-here
YGA_REQUIRE_LOGIN=1
STREAMLIT_AUTH_REDIRECT_URI=http://localhost:8501/oauth2callback
STREAMLIT_AUTH_COOKIE_SECRET=replace-with-a-long-random-secret
STREAMLIT_AUTH_GOOGLE_CLIENT_ID=google-client-id
STREAMLIT_AUTH_GOOGLE_CLIENT_SECRET=google-client-secret
STREAMLIT_AUTH_GOOGLE_SERVER_METADATA_URL=https://accounts.google.com/.well-known/openid-configuration
YGA_ALLOWED_EMAILS=teacher@example.com
YGA_ENABLE_AI_PARSER=1
YGA_AI_MODEL=gpt-4.1-mini
YGA_LOG_TO_FILES=1
```

### 3. Run

```bash
streamlit run app.py
```

기본 주소:

```text
http://localhost:8501
```

## Authentication

이 앱은 Google 로그인을 통해 사용자 접근을 제한할 수 있습니다.

인증 흐름은 다음과 같습니다.

1. 사용자가 앱 접속
2. 로그인 안 된 상태면 Google 로그인 화면 표시
3. 로그인 성공 후 허용 이메일 / 도메인 검사
4. 허용된 사용자만 PPT 생성 기능 접근

### Required Auth Variables

| Variable | Description |
| --- | --- |
| `YGA_REQUIRE_LOGIN` | 로그인 강제 여부 |
| `STREAMLIT_AUTH_REDIRECT_URI` | 로그인 후 돌아올 콜백 주소 |
| `STREAMLIT_AUTH_COOKIE_SECRET` | 로그인 쿠키 서명용 랜덤 비밀값 |
| `STREAMLIT_AUTH_GOOGLE_CLIENT_ID` | Google OAuth Client ID |
| `STREAMLIT_AUTH_GOOGLE_CLIENT_SECRET` | Google OAuth Client Secret |
| `STREAMLIT_AUTH_GOOGLE_SERVER_METADATA_URL` | Google OIDC 메타데이터 URL |

## Allowed Users

허용 사용자 제한은 두 가지 방식으로 설정할 수 있습니다.

### Exact Email Allowlist

```bash
YGA_ALLOWED_EMAILS=teacher@example.com,admin@example.com
```

### Domain Allowlist

```bash
YGA_ALLOWED_EMAIL_DOMAINS=school.ac.kr,example.org
```

둘 다 비워 두면 로그인만 성공한 모든 Google 계정을 허용합니다.

## Google Cloud Setup

Google Cloud Console에서 `OAuth client ID`를 `Web application` 타입으로 생성해야 합니다.

### Authorized JavaScript Origins

```text
http://localhost:8501
https://yga-slide-formatter.onrender.com
```

### Authorized Redirect URIs

```text
http://localhost:8501/oauth2callback
https://yga-slide-formatter.onrender.com/oauth2callback
```

참고:

- 개인 Gmail 계정도 사용할 수 있도록 보통 `외부` 앱으로 만듭니다.
- 앱이 `Testing` 상태면 테스트 사용자로 등록된 계정만 로그인할 수 있습니다.

## Render Deployment

이 저장소에는 `render.yaml`이 포함되어 있어 Blueprint 기반으로 바로 배포할 수 있습니다.

### Deploy Order

1. GitHub 저장소를 Render에 연결
2. Blueprint 생성
3. Render `Environment`에 필요한 값 입력
4. 배포 완료 후 접속
5. 필요 시 `Manual Sync` 또는 `Manual Deploy`

### Required Render Environment Variables

```bash
OPENAI_API_KEY=...
YGA_REQUIRE_LOGIN=1
STREAMLIT_AUTH_REDIRECT_URI=https://yga-slide-formatter.onrender.com/oauth2callback
STREAMLIT_AUTH_COOKIE_SECRET=...
STREAMLIT_AUTH_GOOGLE_CLIENT_ID=...
STREAMLIT_AUTH_GOOGLE_CLIENT_SECRET=...
YGA_ALLOWED_EMAILS=teacher@example.com
```

기본 배포 주소:

```text
https://yga-slide-formatter.onrender.com
```

## Typical Workflow

1. 웹앱 접속
2. Google 로그인
3. 원문 텍스트 입력
4. `제출`
5. 문제 수 / 예상 슬라이드 수 확인
6. `PPT 생성`
7. `PPTX 다운로드`

## Notes

- 이 앱은 OpenAI API를 사용하므로 API 키 관리가 중요합니다.
- `.env`는 Git에 올리지 않습니다.
- Render 배포 환경에서는 `Environment` 값을 사용합니다.
- 로그 파일 저장은 환경변수로 끌 수 있습니다.
- 문제 유형 판정은 AI가 하고, 최종 PPT 조합은 코드가 담당합니다.

## Troubleshooting

### 로그인 버튼 클릭 시 400 오류

아래 항목이 정확히 일치하는지 먼저 확인하세요.

- Google OAuth redirect URI
- Render `STREAMLIT_AUTH_REDIRECT_URI`
- 배포 주소

### 로그인은 되는데 앱에 못 들어감

다음을 확인하세요.

- `YGA_ALLOWED_EMAILS`
- `YGA_ALLOWED_EMAIL_DOMAINS`
- Google 계정 이메일

### Render에서 최신 코드가 반영되지 않음

다음을 실행하세요.

- `Manual Sync`
- 또는 `Manual Deploy`

## License

내부 사용 목적 기준으로 관리 중입니다. 별도 오픈소스 라이선스가 필요하면 추가로 명시하세요.
