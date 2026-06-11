import re
_ALL_PAGE_ALIASES = {
    "all", "allpage", "allpages", "all page", "all pages",
    "every", "everypage", "everypages", "every page", "every pages", "*",
}

def _parse_page_range(s, max_pages):
    s = str(s or "").strip().lower()
    compact = re.sub(r'[\s_\-]+', '', s)
    if not s or s in _ALL_PAGE_ALIASES or compact in _ALL_PAGE_ALIASES:
        return list(range(max_pages))
    pages = set()
    for part in re.split(r'[,;]', s):
        part = part.strip()
        if not part:
            continue
        part_compact = re.sub(r'[\s_\-]+', '', part)
        if part in _ALL_PAGE_ALIASES or part_compact in _ALL_PAGE_ALIASES:
            pages.update(range(max_pages))
            continue
        m = re.match(r'^(\d+)\s*-\s*(\d+)$', part)
        if m:
            a, b = int(m.group(1)), int(m.group(2))
            if a > b:
                a, b = b, a
            pages.update(range(max(0, a - 1), min(max_pages, b)))
        elif part.isdigit():
            p = int(part) - 1
            if 0 <= p < max_pages:
                pages.add(p)
    return sorted(pages)

n = 5
date_rules = [
    {"pages": "1", "date": "1-1-2025"},
    {"pages": "2", "date": "2-2-2025"}
]

page_date = {}
for rule in date_rules:
    pages_str = str(rule.get("pages", "")).strip()
    date_str  = str(rule.get("date",  "")).strip()
    if not date_str: continue
    if pages_str:
        for p in _parse_page_range(pages_str, n):
            page_date[p] = date_str
    else:
        for p in range(n):
            page_date[p] = date_str

print("Rules:", date_rules)
print("page_date:", page_date)

date_rules2 = [
    {"pages": "1", "date": "1-1-2025"},
    {"pages": "all", "date": "2-2-2025"}
]

page_date2 = {}
for rule in date_rules2:
    pages_str = str(rule.get("pages", "")).strip()
    date_str  = str(rule.get("date",  "")).strip()
    if not date_str: continue
    if pages_str:
        for p in _parse_page_range(pages_str, n):
            page_date2[p] = date_str
    else:
        for p in range(n):
            page_date2[p] = date_str

print("Rules 2:", date_rules2)
print("page_date 2:", page_date2)
