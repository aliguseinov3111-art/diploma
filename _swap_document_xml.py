"""
Swap word/document.xml: keep everything from the original (images, styles,
media, relationships) and replace only the document body with the rewritten
version from the final file.
"""
import zipfile
import sys
import io

sys.stdout.reconfigure(encoding='utf-8')

SRC_ALL   = '02 ВКР Али.docx'          # original — all binary / style / media
SRC_BODY  = '02_VKR_Ali_final.docx'    # has rewritten text
OUT       = '02_VKR_Ali_final.docx'    # overwrite final with clean version

# 1. Read the rewritten document.xml from the final
with zipfile.ZipFile(SRC_BODY, 'r') as z:
    new_doc_xml = z.read('word/document.xml')
print(f'Rewritten document.xml: {len(new_doc_xml):,} bytes')

# 2. Build new docx: original zip, but with document.xml swapped
buf = io.BytesIO()
with zipfile.ZipFile(SRC_ALL, 'r') as z_orig:
    with zipfile.ZipFile(buf, 'w', compression=zipfile.ZIP_DEFLATED) as z_out:
        for info in z_orig.infolist():
            if info.filename == 'word/document.xml':
                z_out.writestr(info, new_doc_xml)
                print(f'  swapped: {info.filename}')
            else:
                z_out.writestr(info, z_orig.read(info.filename))

# 3. Write output
data = buf.getvalue()
with open(OUT, 'wb') as f:
    f.write(data)

print(f'\nSaved → {OUT}  ({len(data)/1024/1024:.2f} MB)')
