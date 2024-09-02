import fitz


def extract_bookmark(pdf_path):
    doc = fitz.open(pdf_path)
    toc = doc.get_toc(simple=False)
    return _parse_toc(toc)


def _parse_toc(toc):
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
