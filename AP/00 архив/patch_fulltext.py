import sys, re
sys.stdout.reconfigure(encoding='utf-8')

vkr_path  = r'd:\diploma\АП\_vkr_text.txt'
html_path = r'd:\diploma\АП\netlify_deploy\index.html'

# ── 1. Read VKR text ────────────────────────────────────────────────────────
raw = open(vkr_path, encoding='utf-8').read()

# ── 2. Split into paragraphs ─────────────────────────────────────────────────
# Natural split points: figure/table markers, and sentence-ending periods
# followed by a capital letter
segments = re.split(r'(?= / (?:Рис|Таб|Рисунок|Таблица))', raw)

paragraphs = []
for seg in segments:
    # Further split long segments at sentence boundaries
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

# ── 3. Group paragraphs into pages (~3 000 chars each) ──────────────────────
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

print(f'Параграфов: {len(paragraphs)}, страниц: {len(pages)}')

# ── 4. Build replacement HTML ─────────────────────────────────────────────────
def escape(text):
    return (text.replace('&', '&amp;')
                .replace('<', '&lt;')
                .replace('>', '&gt;'))

html_pages = ''
for i, page_paras in enumerate(pages):
    page_num = i + 1
    html_pages += f'<div class="ap-text-page" data-page="{page_num}">'
    for para in page_paras:
        html_pages += f'<p><span class="normal">{escape(para)}</span></p>'
    html_pages += '</div>'
    if i < len(pages) - 1:
        html_pages += (f'<div class="ap-page-divider">'
                       f'<span>Страница {page_num}</span>'
                       f'<hr>'
                       f'<span>Страница {page_num + 1}</span>'
                       f'</div>')

new_doc = f'<div class="ap-text-document">{html_pages}</div>'

# ── 5. Replace in HTML ────────────────────────────────────────────────────────
html = open(html_path, encoding='utf-8').read()

doc_start = html.find('<div class="ap-text-document">')
doc_end   = 478222   # pre-computed closing position

html = html[:doc_start] + new_doc + html[doc_end:]
open(html_path, 'w', encoding='utf-8').write(html)

print(f'Готово: {html_path}')
print(f'Новый размер HTML: {len(html)//1024} КБ')
