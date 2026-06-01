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

# FastAPI에서 프론트엔드 정적 파일 서빙하도록 main.py 수정
RUN python -c "
import re, pathlib
code = pathlib.Path('main.py').read_text()
if 'StaticFiles' not in code:
    insert = '''
from fastapi.staticfiles import StaticFiles
from fastapi.responses import HTMLResponse
import os as _os
'''
    mount = '''
# 프론트엔드 정적 파일 서빙
_static_dir = _os.path.join(_os.path.dirname(__file__), 'static')
if _os.path.exists(_static_dir):
    app.mount('/static', StaticFiles(directory=_static_dir), name='static')

@app.get('/', response_class=HTMLResponse)
def root():
    idx = _os.path.join(_static_dir, 'index.html')
    return open(idx, encoding='utf-8').read()
'''
    code = code.replace('app = FastAPI(title=\"주일예배 PPT 생성기\")', 
                        insert + 'app = FastAPI(title=\"주일예배 PPT 생성기\")' + mount)
    pathlib.Path('main.py').write_text(code)
"

EXPOSE 8000

CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000"]
