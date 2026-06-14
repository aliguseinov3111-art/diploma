import sys, re, random
sys.stdout.reconfigure(encoding='utf-8')

vkr_path  = r'd:\diploma\АП\_vkr_text.txt'
html_path = r'd:\diploma\АП\netlify_deploy\index.html'

raw = open(vkr_path, encoding='utf-8').read()

# ── 1. Split into paragraphs (same as patch_fulltext.py) ────────────────────
segments = re.split(r'(?= / (?:Рис|Таб|Рисунок|Таблица))', raw)
paragraphs = []
for seg in segments:
    if len(seg) > 2000:
        parts = re.split(r'(?<=[.!?])\s{1,2}(?=[А-ЯA-Z1-9\(])', seg)
        chunk = ''
        for part in parts:
            chunk += part + ' '
            if len(chunk) >= 800:
                paragraphs.append(chunk.strip())
                chunk = ''
        if chunk.strip():
            paragraphs.append(chunk.strip())
    else:
        if seg.strip():
            paragraphs.append(seg.strip())

# ── 2. Group into pages ──────────────────────────────────────────────────────
pages = []
cur_page = []
cur_len  = 0
PAGE_SIZE = 3000

for para in paragraphs:
    cur_page.append(para)
    cur_len += len(para)
    if cur_len >= PAGE_SIZE:
        pages.append(cur_page)
        cur_page = []
        cur_len  = 0
if cur_page:
    pages.append(cur_page)

# ── 3. Build HTML with ~28% plagiat spans ────────────────────────────────────
def escape(text):
    return (text.replace('&', '&amp;')
                .replace('<', '&lt;')
                .replace('>', '&gt;'))

rng = random.Random(42)
TARGET = 0.28   # fraction of sentences to mark red

def para_to_html(para):
    # Split at sentence boundaries (capital letter after punctuation)
    sents = re.split(r'(?<=[.!?])\s+(?=[А-ЯA-Z«\(])', para)
    if len(sents) == 1:
        cls = 'plagiat' if rng.random() < TARGET else 'normal'
        return f'<p><span class="{cls}">{escape(para)}</span></p>'

    parts_html = ''
    i = 0
    while i < len(sents):
        if rng.random() < TARGET:
            # red run: 1–3 consecutive sentences
            run = min(rng.randint(1, 3), len(sents) - i)
            text = ' '.join(sents[i:i+run])
            parts_html += f'<span class="plagiat">{escape(text)} </span>'
            i += run
        else:
            parts_html += f'<span class="normal">{escape(sents[i])} </span>'
            i += 1

    return f'<p>{parts_html.rstrip()}</p>'

html_pages = ''
for i, page_paras in enumerate(pages):
    page_num = i + 1
    html_pages += f'<div class="ap-text-page" data-page="{page_num}">'
    for para in page_paras:
        html_pages += para_to_html(para)
    html_pages += '</div>'
    if i < len(pages) - 1:
        html_pages += (f'<div class="ap-page-divider">'
                       f'<span>Страница {page_num}</span>'
                       f'<hr>'
                       f'<span>Страница {page_num + 1}</span>'
                       f'</div>')

new_doc = f'<div class="ap-text-document">{html_pages}</div>'

# ── 4. Find current ap-text-document boundaries (dynamic, no hardcoded offset)
html = open(html_path, encoding='utf-8').read()

doc_start = html.find('<div class="ap-text-document">')
if doc_start == -1:
    print("ERROR: ap-text-document not found in index.html")
    sys.exit(1)

# Count balanced <div> tags to find the matching </div>
pos   = doc_start + len('<div class="ap-text-document">')
depth = 1
while depth > 0:
    next_open  = html.find('<div', pos)
    next_close = html.find('</div>', pos)
    if next_close == -1:
        print("ERROR: unbalanced <div> — file may be corrupt")
        sys.exit(1)
    if next_open != -1 and next_open < next_close:
        depth += 1
        pos = next_open + 4
    else:
        depth -= 1
        pos = next_close + 6

doc_end = pos   # position right after the closing </div>

# ── 5. Write result ──────────────────────────────────────────────────────────
html = html[:doc_start] + new_doc + html[doc_end:]
open(html_path, 'w', encoding='utf-8').write(html)

# Quick stats
total_chars = sum(len(p) for p in paragraphs)
print(f'Страниц: {len(pages)}, параграфов: {len(paragraphs)}')
print(f'Готово: {html_path}')
print(f'Новый размер HTML: {len(html)//1024} КБ')
