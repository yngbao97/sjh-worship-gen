FROM python:3.12-slim

# LibreOffice + 한글 폰트 설치
RUN apt-get update && apt-get install -y \
    libreoffice \
    libreoffice-impress \
    fonts-noto-cjk \
    fonts-noto-cjk-extra \
    --no-install-recommends \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Python 의존성 설치
COPY backend/requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 백엔드 코드 복사
COPY backend/ ./

# 프론트엔드 복사 (static 폴더로)
COPY frontend/ ./static/

EXPOSE 8000

CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000"]
