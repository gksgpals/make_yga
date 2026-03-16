# 영어교재 PPT 자동 생성기

텍스트 문제지를 입력하면 수업용 `.pptx`를 생성합니다.  
입력 방식은 텍스트 직접 입력 또는 `.txt` 업로드만 지원합니다.

## 설치

```bash
git clone <your-repo-url>
cd make_yga
pip install -r requirements.txt
```

## Streamlit 앱 실행

```bash
python3 -m streamlit run app.py
```

또는 아래처럼 실행해도 자동으로 Streamlit 서버로 재실행됩니다.

```bash
python3 app.py
```

사용 순서:

1. 텍스트를 붙여 넣거나 `.txt` 파일 업로드
2. `PPT 생성` 클릭
3. `PPTX 다운로드` 클릭

## 웹 배포(Render 권장)

비전공자도 브라우저에서 바로 쓸 수 있게 하려면 Render 같은 공개 웹 호스팅에 배포하는 구성이 가장 단순합니다.

이 저장소에는 배포 준비 파일이 포함돼 있습니다.

- `render.yaml`: Render Web Service 설정
- `.python-version`: 배포 Python 버전 고정
- `.env.example`: 필요한 환경변수 예시

배포 순서:

1. 이 프로젝트를 GitHub 저장소로 올립니다.
2. Render에서 `New +` -> `Blueprint` 또는 `Web Service`로 저장소를 연결합니다.
3. 환경변수 `OPENAI_API_KEY`를 Render 대시보드에서 설정합니다.
4. 배포가 끝나면 Render가 발급한 공개 URL로 접속합니다.

Render 시작 명령:

```bash
streamlit run app.py --server.port $PORT --server.address 0.0.0.0 --server.headless true
```

배포 환경변수 기본값:

- `PYTHON_VERSION=3.11.11`
- `YGA_ENABLE_AI_PARSER=1`
- `YGA_AI_MODEL=gpt-4.1-mini`
- `YGA_LOG_TO_FILES=0`

`YGA_LOG_TO_FILES=0` 이면 배포 서버에서는 로컬 파일 로그를 남기지 않고 콘솔 로그만 사용합니다.

## CLI 사용(선택)

```bash
python3 slide_formatter.py input.txt
python3 slide_formatter.py input.txt -o output.txt
python3 slide_formatter.py input.txt --pptx class_slides.pptx
python3 slide_formatter.py input.txt --pptx class_slides.pptx --pdf class_slides.pdf
cat input.txt | python3 slide_formatter.py
```

`--pdf` 사용 시 `LibreOffice(soffice)`가 필요합니다.

## 입력 포맷

문제 번호/헤더 패턴(`1-A`, `2-B` 등) 또는 충분한 빈 줄 구분이 있으면 자동 분리됩니다.

예시:

```text
1-A 다음 글의 목적으로 가장 적절한 것은?
Dear students,
...
문제. 윗글의 목적으로 가장 적절한 것은?
① to announce a club expansion
② to cancel all clubs
③ to request uniforms
④ to change class hours
⑤ to open a cafeteria
```

라벨 포맷도 지원합니다.

```text
[TITLE] Reading Practice
[PASSAGE]
Dear students, ...
[QUESTION] What is the main purpose?
[A] ...
[B] ...
[C] ...
[D] ...
```

인식 라벨:

- `[TITLE]`, `TITLE:`
- `[PASSAGE]`, `[BODY]`, `[본문]`
- `[QUESTION]`, `[Q]`, `[문제]`
- `[CHOICE]`, `[선지]`, `[A]...[E]`, `1)`, `A)`, `①`

## 스타일 기본값

- 배경: 검정(`#000000`)
- 글자: 흰색(`#FFFFFF`)
- 폰트: 굴림(없으면 시스템 대체)
- 줄간격: 180%
