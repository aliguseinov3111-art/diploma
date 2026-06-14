import sys, re
sys.stdout.reconfigure(encoding='utf-8')

path_file = r'C:\Users\Honor\Downloads\pdf_path.txt'
target = open(path_file, encoding='utf-8-sig').read().strip()
print(f"Reading: {target}")

with open(target, 'rb') as f:
    data = f.read()

chunks = re.findall(b'BT.*?ET', data, re.DOTALL)
texts = []
for c in chunks:
    parts = re.findall(b'\(([^)]{2,})\)', c)
    for p in parts:
        try:
            t = p.decode('cp1251', errors='ignore').strip()
            if t and len(t) > 1:
                texts.append(t)
        except:
            pass

if texts:
    print('\n'.join(texts[:300]))
else:
    print("No text in BT/ET streams — PDF may be image-based or use ToUnicode CMap")
    print(f"File size: {len(data)} bytes")
