# 주일예배 PPT 생성기 (웹 버전)

**서비스 링크: https://sjh-worship-gen.onrender.com/**

## 배포 방법 (Railway)

### 1단계: 환경변수 준비
서비스 계정 JSON 키 파일을 한 줄로 변환:
```bash
# Mac/Linux
cat your-key.json | tr -d '\n'

# Windows PowerShell
(Get-Content your-key.json -Raw) -replace "`n","\n"
```

### 2단계: Railway 배포
1. https://railway.app 접속 → GitHub 로그인
2. "New Project" → "Deploy from GitHub repo"
3. 이 프로젝트 레포 선택
4. "Variables" 탭에서 환경변수 추가:
   - `GOOGLE_SERVICE_ACCOUNT_JSON` = 위에서 변환한 JSON 한 줄
   - `DRIVE_ROOT_FOLDER_ID` = `1QQiCwz9TQiz90haWQlaojHfSAknSsbJc`
5. Deploy!

### 로컬 테스트
```bash
# .env 파일 생성
cp .env.example .env
# .env 파일에 실제 값 입력 후:

docker compose up --build
# 브라우저에서 http://localhost:8000 접속
```

## 프로젝트 구조
```
worship-web/
├── backend/
│   ├── main.py           # FastAPI 서버
│   ├── worship_core.py   # PPT 생성 핵심 로직
│   ├── drive_client.py   # Google Drive 파일 접근
│   └── requirements.txt
├── frontend/
│   └── index.html        # 웹 UI (단일 파일)
├── Dockerfile
├── docker-compose.yml
└── .env.example
```
