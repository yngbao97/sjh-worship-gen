"""
Google Drive 파일 접근 클라이언트
서비스 계정 JSON으로 인증, 찬송가/교독문/템플릿 파일 다운로드
"""

import os
import io
import re
import tempfile
from pathlib import Path
from functools import lru_cache
from typing import Optional

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

SCOPES = ['https://www.googleapis.com/auth/drive.readonly']

# 환경변수에서 루트 폴더 ID 읽기
ROOT_FOLDER_ID = os.environ.get('DRIVE_ROOT_FOLDER_ID', '')

# 기본 템플릿 파일명
DEFAULT_TEMPLATE_1BU = '20250000-주일예배 샘플(1부).pptx'
DEFAULT_TEMPLATE_2BU = '20250000-주일예배 샘플(2부).pptx'

HYMN_RANGES = [(1,100),(101,200),(201,300),(301,400),(401,500),(501,600),(601,645)]


def _get_service():
    """Google Drive 서비스 객체 반환"""
    import json

    creds_json = os.environ.get('GOOGLE_SERVICE_ACCOUNT_JSON', '')
    if not creds_json:
        raise RuntimeError("GOOGLE_SERVICE_ACCOUNT_JSON 환경변수가 설정되지 않았습니다.")

    creds_info = json.loads(creds_json)
    credentials = service_account.Credentials.from_service_account_info(
        creds_info, scopes=SCOPES
    )
    return build('drive', 'v3', credentials=credentials)


def _list_folder(service, folder_id: str) -> list:
    """폴더 내 파일/폴더 목록 반환"""
    results = []
    page_token = None
    while True:
        resp = service.files().list(
            q=f"'{folder_id}' in parents and trashed=false",
            fields="nextPageToken, files(id, name, mimeType)",
            pageToken=page_token,
            pageSize=1000,
            includeItemsFromAllDrives=True,
            supportsAllDrives=True,
        ).execute()
        results.extend(resp.get('files', []))
        page_token = resp.get('nextPageToken')
        if not page_token:
            break
    return results


def _download_file(service, file_id: str, dest_path: str):
    """Drive 파일을 로컬 경로로 다운로드"""
    request = service.files().get_media(fileId=file_id, supportsAllDrives=True)
    with open(dest_path, 'wb') as f:
        downloader = MediaIoBaseDownload(f, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()


def _find_folder_by_name(service, parent_id: str, name: str) -> Optional[str]:
    """parent 폴더 안에서 이름으로 폴더 ID 찾기"""
    items = _list_folder(service, parent_id)
    for item in items:
        if item['name'] == name and item['mimeType'] == 'application/vnd.google-apps.folder':
            return item['id']
    return None


def _find_file_by_name(service, parent_id: str, name: str) -> Optional[str]:
    """parent 폴더 안에서 이름으로 파일 ID 찾기"""
    items = _list_folder(service, parent_id)
    for item in items:
        if item['name'] == name and item['mimeType'] != 'application/vnd.google-apps.folder':
            return item['id']
    return None


def _get_db_ppt_folder_id(service) -> Optional[str]:
    """'자막 Database (PPT)' 폴더 ID 탐색"""
    # 루트 → 자막 Database (2024-...) → 자막 Database (PPT)
    root_items = _list_folder(service, ROOT_FOLDER_ID)
    db_folder_id = None
    for item in root_items:
        if '자막 Database' in item['name'] and item['mimeType'] == 'application/vnd.google-apps.folder':
            db_folder_id = item['id']
            break
    if not db_folder_id:
        return None

    db_items = _list_folder(service, db_folder_id)
    for item in db_items:
        if item['name'] == '자막 Database (PPT)' and item['mimeType'] == 'application/vnd.google-apps.folder':
            return item['id']
    return None


def download_hymn(number: int, dest_dir: str) -> Optional[str]:
    """
    찬송가 번호로 Drive에서 파일 다운로드.
    반환: 로컬 파일 경로 (없으면 None)
    """
    service = _get_service()
    db_ppt_id = _get_db_ppt_folder_id(service)
    if not db_ppt_id:
        return None

    hymn_folder_id = _find_folder_by_name(service, db_ppt_id, '찬송가')
    if not hymn_folder_id:
        return None

    # 서브폴더명 탐색 (예: '1~100장')
    for start, end in HYMN_RANGES:
        if start <= number <= end:
            subfolder_name = f'{start}~{end}장'
            break
    else:
        return None

    sub_id = _find_folder_by_name(service, hymn_folder_id, subfolder_name)
    if not sub_id:
        # '장' 없는 버전도 시도
        subfolder_name_no_jang = subfolder_name.replace('장', '')
        sub_id = _find_folder_by_name(service, hymn_folder_id, subfolder_name_no_jang)
        if not sub_id:
            return None

    # 파일명 후보
    for fname in [f'{number}장.pptx', f'{number}.pptx', f'찬송가{number}장.pptx']:
        file_id = _find_file_by_name(service, sub_id, fname)
        if file_id:
            dest_path = str(Path(dest_dir) / fname)
            _download_file(service, file_id, dest_path)
            return dest_path

    return None


def download_responsory(number: int, dest_dir: str) -> Optional[str]:
    """
    교독문 번호로 Drive에서 파일 다운로드.
    반환: 로컬 파일 경로 (없으면 None)
    """
    service = _get_service()
    db_ppt_id = _get_db_ppt_folder_id(service)
    if not db_ppt_id:
        return None

    resp_folder_id = _find_folder_by_name(service, db_ppt_id, '교독문')
    if not resp_folder_id:
        return None

    for fname in [
        f'교독문 {number}번.pptx',
        f'교독문_{number}번.pptx',
        f'교독문{number}번.pptx',
        f'교독문 {number:03d}번.pptx',
        f'{number}번.pptx',
    ]:
        file_id = _find_file_by_name(service, resp_folder_id, fname)
        if file_id:
            dest_path = str(Path(dest_dir) / fname)
            _download_file(service, file_id, dest_path)
            return dest_path

    # 폴더 전체 순회해서 번호 포함 파일 탐색
    items = _list_folder(service, resp_folder_id)
    for item in items:
        if item['mimeType'] == 'application/vnd.google-apps.folder':
            continue
        if re.search(rf'(?<!\d){number}(?!\d)', Path(item['name']).stem):
            dest_path = str(Path(dest_dir) / item['name'])
            _download_file(service, item['id'], dest_path)
            return dest_path

    return None


def download_template(part: str, dest_dir: str, custom_file_id: str = None) -> Optional[str]:
    """
    템플릿 파일 다운로드.
    part: '1bu' 또는 '2bu'
    custom_file_id: 사용자 지정 Drive 파일 ID (없으면 기본 파일명으로 탐색)
    반환: 로컬 파일 경로
    """
    service = _get_service()

    if custom_file_id:
        fname = f'template_{part}.pptx'
        dest_path = str(Path(dest_dir) / fname)
        _download_file(service, custom_file_id, dest_path)
        return dest_path

    # 기본: 루트 폴더에서 파일명으로 탐색
    default_name = DEFAULT_TEMPLATE_1BU if part == '1bu' else DEFAULT_TEMPLATE_2BU
    import unicodedata
    all_items = _list_folder(service, ROOT_FOLDER_ID)
    norm_default = unicodedata.normalize('NFC', default_name)
    file_id = next(
        (i['id'] for i in all_items
         if unicodedata.normalize('NFC', i['name']) == norm_default),
        None
    )
    if file_id:
        dest_path = str(Path(dest_dir) / default_name)
        _download_file(service, file_id, dest_path)
        return dest_path

    return None
