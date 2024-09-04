"""
북마크를 추출하여 현재 북마크와 이전 북마크의 리스트를 저장합니다.
[현재북마크 레벨, 현재북마크 제목, 현재북마크 페이지수, 상위레벨 북마크 리스트]
"""

import fitz


def extract_bookmark(pdf_path):
    """pdf에서 북마크를 추출합니다"""
    doc = fitz.open(pdf_path)
    toc = doc.get_toc(simple=False)
    return _parse_toc(toc)


def _parse_toc(toc):
    """추출한 북마크의 정보들을 취합하여 리스트로 저장합니다"""
    bookmarks = []

    for item in toc:
        level, title, page, _ = item
        bookmarks.append({
            "level": level,
            "title": title,
            "page": page,
            "parent": None
        })

    for i in range(1, len(bookmarks)):
        for j in range(i-1, -1, -1):
            if bookmarks[j]['level'] < bookmarks[i]['level']:
                bookmarks[i]['parent'] = bookmarks[j]
                break

    return bookmarks
