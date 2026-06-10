"""
Godpia 온라인 성경에서 본문을 자동으로 가져오는 클라이언트.
POST https://www.godpia.com/read/reading_body.asp
"""

import re
import requests
from bs4 import BeautifulSoup
from typing import Optional

GODPIA_URL = 'https://www.godpia.com/read/reading_body.asp'
GODPIA_HEADERS = {
    'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
    'Origin': 'https://www.godpia.com',
    'Referer': 'https://www.godpia.com/read/reading.asp',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
}

# 한국어 책 이름 → 약어 (설교제목 슬라이드용)
BOOK_ABBR: dict[str, str] = {
    '창세기': '창', '출애굽기': '출', '레위기': '레', '민수기': '민',
    '신명기': '신', '여호수아': '수', '사사기': '삿', '룻기': '룻',
    '사무엘상': '삼상', '사무엘하': '삼하', '열왕기상': '왕상', '열왕기하': '왕하',
    '역대상': '대상', '역대하': '대하', '에스라': '스', '느헤미야': '느',
    '에스더': '에', '욥기': '욥', '시편': '시', '잠언': '잠',
    '전도서': '전', '아가': '아', '이사야': '사', '예레미야': '렘',
    '예레미야애가': '애', '에스겔': '겔', '다니엘': '단', '호세아': '호',
    '요엘': '욜', '아모스': '암', '오바댜': '옵', '요나': '욘',
    '미가': '미', '나훔': '나', '하박국': '합', '스바냐': '습',
    '학개': '학', '스가랴': '슥', '말라기': '말',
    '마태복음': '마', '마가복음': '막', '누가복음': '눅', '요한복음': '요',
    '사도행전': '행', '로마서': '롬', '고린도전서': '고전', '고린도후서': '고후',
    '갈라디아서': '갈', '에베소서': '엡', '빌립보서': '빌', '골로새서': '골',
    '데살로니가전서': '살전', '데살로니가후서': '살후', '디모데전서': '딤전',
    '디모데후서': '딤후', '디도서': '딛', '빌레몬서': '몬', '히브리서': '히',
    '야고보서': '약', '베드로전서': '벧전', '베드로후서': '벧후',
    '요한일서': '요일', '요한이서': '요이', '요한삼서': '요삼',
    '유다서': '유', '요한계시록': '계',
}

def book_name_to_abbr(book_name: str) -> str:
    """책 전체 이름 → 약어. 매핑 없으면 원본 반환."""
    return BOOK_ABBR.get(book_name.strip(), book_name.strip())

# 한국어 책 이름 → Godpia vol 코드
# 사용자가 직접 확인한 코드로 수정 가능
BOOK_CODES: dict[str, str] = {
    # 구약
    '창세기': 'gen', '출애굽기': 'exo', '레위기': 'lev', '민수기': 'num',
    '신명기': 'deu', '여호수아': 'jos', '사사기': 'jdg', '룻기': 'rut',
    '사무엘상': '1sa', '사무엘하': '2sa', '열왕기상': '1ki', '열왕기하': '2ki',
    '역대상': '1ch', '역대하': '2ch', '에스라': 'ezr', '느헤미야': 'neh',
    '에스더': 'est', '욥기': 'job', '시편': 'psa', '잠언': 'pro',
    '전도서': 'ecc', '아가': 'sng', '이사야': 'isa', '예레미야': 'jer',
    '예레미야애가': 'lam', '에스겔': 'ezk', '다니엘': 'dan', '호세아': 'hos',
    '요엘': 'jol', '아모스': 'amo', '오바댜': 'oba', '요나': 'jnh',
    '미가': 'mic', '나훔': 'nam', '하박국': 'hab', '스바냐': 'zep',
    '학개': 'hag', '스가랴': 'zec', '말라기': 'mal',
    # 신약
    '마태복음': 'mat', '마가복음': 'mrk', '누가복음': 'luk', '요한복음': 'jhn',
    '사도행전': 'act', '로마서': 'rom', '고린도전서': '1co', '고린도후서': '2co',
    '갈라디아서': 'gal', '에베소서': 'eph', '빌립보서': 'php', '골로새서': 'col',
    '데살로니가전서': '1th', '데살로니가후서': '2th', '디모데전서': '1ti',
    '디모데후서': '2ti', '디도서': 'tit', '빌레몬서': 'phm', '히브리서': 'heb',
    '야고보서': 'jas', '베드로전서': '1pe', '베드로후서': '2pe',
    '요한1서': '1jn', '요한2서': '2jn', '요한3서': '3jn',
    '요한일서': '1jn', '요한이서': '2jn', '요한삼서': '3jn',
    '유다서': 'jud', '요한계시록': 'rev',
}


def _book_name_to_code(book_name: str) -> Optional[str]:
    """책 이름(전체 또는 약어)으로 Godpia vol 코드 반환."""
    name = book_name.strip()
    if name in BOOK_CODES:
        return BOOK_CODES[name]
    # 부분 일치 (예: '창' → '창세기')
    for full, code in BOOK_CODES.items():
        if full.startswith(name) or name in full:
            return code
    return None


def parse_bible_label(bible_label: str) -> Optional[str]:
    """
    봉독 구절 표기에서 책 이름 추출 후 Godpia vol 코드 반환.
    예: '창세기 3장 16-18절' → 'gen'
        '요한복음 3:16'      → 'joh'
    """
    # 숫자, 장, 절, 콜론, 하이픈, 공백 제거 → 책 이름만 추출
    book_name = re.split(r'\s*\d', bible_label.strip())[0].strip()
    return _book_name_to_code(book_name)


def parse_bible_ref(bible_text: str) -> tuple[Optional[str], int, Optional[int], Optional[int]]:
    """
    성경 본문(약어 구절 참조)에서 vol코드, 장, 시작절, 끝절 파싱.
    예: '창 3:16-18' → ('gen', 3, 16, 18)
        '창 3:16'    → ('gen', 3, 16, 16)
        '창 3'       → ('gen', 3, None, None)
    반환: (vol_code, chapter, start_verse, end_verse)
    """
    text = bible_text.strip()
    # 책 약어: 숫자 앞까지
    m = re.match(r'^([가-힣]+)\s+(\d+)(?::(\d+)(?:-(\d+))?)?', text)
    if not m:
        return None, 0, None, None
    book_abbr = m.group(1)
    chap = int(m.group(2))
    start = int(m.group(3)) if m.group(3) else None
    end = int(m.group(4)) if m.group(4) else (start if start else None)
    vol = _book_name_to_code(book_abbr)
    return vol, chap, start, end


def fetch_bible_verses(vol: str, chap: int, ver: str = 'gae') -> list[dict]:
    """
    Godpia에서 특정 장 전체를 가져와 절 목록 반환.
    반환: [{'sec': 1, 'text': '본문...'}, ...]
    """
    data = f'vol={vol}&chap={chap}&ver={ver}'
    resp = requests.post(GODPIA_URL, data=data, headers=GODPIA_HEADERS, timeout=10)
    resp.raise_for_status()
    resp.encoding = 'utf-8'

    soup = BeautifulSoup(resp.text, 'html.parser')
    verses = []
    for li in soup.select('li[datavol], li[dataVol]'):
        sec_attr = li.get('datasec') or li.get('dataSec') or li.get('dataSec'.lower())
        if not sec_attr:
            # id 속성에서 추출: gae_gen_3_1
            id_val = li.get('id', '')
            parts = id_val.rsplit('_', 1)
            sec_attr = parts[-1] if len(parts) == 2 else None
        if not sec_attr:
            continue
        cont = li.select_one('.bible-read-cont')
        if not cont:
            continue
        # dic-link span 텍스트 포함, 불필요한 공백 정리
        text = cont.get_text(separator='', strip=True)
        text = re.sub(r'\s+', ' ', text).strip()
        verses.append({'sec': int(sec_attr), 'text': text})
    return verses


def fetch_bible_text(
    vol: str,
    book_name: str,
    chap: int,
    start_verse: Optional[int] = None,
    end_verse: Optional[int] = None,
    ver: str = 'gae',
) -> list[dict]:
    """
    Godpia에서 성경 본문을 가져와 apply_bible_text 호환 형식으로 반환.

    vol:         Godpia vol 코드 (예: 'gen')
    book_name:   슬라이드 레이블용 책 이름 (예: '창세기')
    chap:        장 번호
    start_verse: 시작 절 (None이면 장 전체)
    end_verse:   마지막 절 (None이면 start_verse까지)
    반환: [{'label': '창세기 3:16', 'lines': ['본문텍스트']}, ...]
    """
    all_verses = fetch_bible_verses(vol, chap, ver)

    if start_verse is not None:
        all_verses = [v for v in all_verses if start_verse <= v['sec'] <= (end_verse or start_verse)]

    return [
        {'label': f'{book_name} {chap}:{v["sec"]}', 'lines': [v['text']]}
        for v in all_verses
    ]
