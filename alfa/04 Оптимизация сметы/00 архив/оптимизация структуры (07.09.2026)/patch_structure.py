# -*- coding: utf-8 -*-
"""
Оптимизация структуры листа «Смета ЕР (общ) » — хирургический патч zip-XML.
НЕ через openpyxl.save (режет x14-dataValidation) и НЕ через COM (бьёт _FilterDatabase).

Три правки, только xl/worksheets/sheet1.xml:
  1. Закрепление областей — строки 1..14 (пустая 1 + шапка 2..13 + строка столбцов 14).
  2. Группировка строк — 2-уровневый структурный план:
       уровень 0  — 10 строк «Раздел №N» (+ шапка 1..14, итоги 1687..1691);
       уровень 1  — строки-заголовки подразделов и строки-подытоги (SUM/ROUND-роллап по N);
       уровень 2  — позиции (N = ROUND(I*L,2)).
     summaryBelow=0 — кнопка [-] на строке-заголовке (итог над блоком).
     Лист открывается развёрнутым; сворачивает пользователь кнопками [1][2][3].
  3. N и O: 3222 формулы вида
       =ROUND('Смета ЕР (общ) '!$I17*'Смета ЕР (общ) '!$L17,2)
     -> =ROUND(I17*L17,2)   (N),  =ROUND(I17*M17,2)  (O)
     Самоссылка на свой же лист по полному имени убрана; расчёт идентичен.

Всё остальное (x14-проверки на «Освещение»/«Освещение Рефлекта», _xlnm._FilterDatabase,
calcChain, sharedStrings, styles, прочие 40 листов) — байт-в-байт без изменений.
"""
import sys, re, zipfile, shutil, os

SRC = r"d:\diploma\alfa\04 Оптимизация сметы\ТС ЕР исходник.xlsx"
DST = sys.argv[1] if len(sys.argv) > 1 else SRC  # по умолчанию — на месте

SHEET_QUAL = "'Смета ЕР (общ) '!"
SECTION_ROWS = {15, 109, 253, 573, 639, 751, 879, 988, 1077, 1661}  # старты Р1..Р10
BODY_FIRST, BODY_LAST = 15, 1685

def classify_levels(sheet_xml):
    """r -> outlineLevel (1|2) для строк тела; уровень 0 не пишем."""
    # формула столбца N для каждой строки
    n_formula = {}
    for m in re.finditer(r'<c r="N(\d+)"[^>]*>(?:<f[^>]*>([^<]*)</f>)?', sheet_xml):
        n_formula[int(m.group(1))] = (m.group(2) or "")
    levels = {}
    detail_re = re.compile(r"^ROUND\('Смета ЕР \(общ\) '!\$I\d+\*'Смета ЕР \(общ\) '!\$[LM]\d+,2\)$")
    for r in range(BODY_FIRST, BODY_LAST + 1):
        if r in SECTION_ROWS:
            continue  # уровень 0
        f = n_formula.get(r, "")
        if detail_re.match(f):
            levels[r] = 2          # позиция
        else:
            levels[r] = 1          # заголовок подраздела / подытог / метка
    return levels

def patch_sheet(xml):
    rep = {"freeze": 0, "outlinePr": 0, "rows": 0, "N": 0, "O": 0}

    # --- 2a. outlinePr в <sheetPr> ---
    old = '<sheetPr><tabColor rgb="FF00B050"/><pageSetUpPr fitToPage="1"/></sheetPr>'
    new = '<sheetPr><tabColor rgb="FF00B050"/><outlinePr summaryBelow="0" summaryRight="0"/><pageSetUpPr fitToPage="1"/></sheetPr>'
    assert xml.count(old) == 1, "sheetPr не найден в ожидаемом виде"
    xml = xml.replace(old, new, 1); rep["outlinePr"] = 1

    # --- 1. закрепление областей ---
    old = '<sheetView showGridLines="0" tabSelected="1" zoomScale="54" zoomScaleNormal="112" workbookViewId="0"/>'
    new = ('<sheetView showGridLines="0" tabSelected="1" zoomScale="54" zoomScaleNormal="112" workbookViewId="0">'
           '<pane ySplit="14" topLeftCell="A15" activePane="bottomLeft" state="frozen"/>'
           '<selection pane="bottomLeft" activeCell="A15" sqref="A15"/>'
           '</sheetView>')
    assert xml.count(old) == 1, "sheetView не найден в ожидаемом виде"
    xml = xml.replace(old, new, 1); rep["freeze"] = 1

    # --- 2b. outlineLevel на строках тела ---
    levels = classify_levels(xml)
    for r, lvl in levels.items():
        a = f'<row r="{r}" spans='
        b = f'<row r="{r}" outlineLevel="{lvl}" spans='
        if a in xml:
            xml = xml.replace(a, b, 1); rep["rows"] += 1
        else:
            raise AssertionError(f'<row r="{r}" spans=...> не найден')

    # --- 3. N/O: убрать самоквалификацию ---
    def n_sub(m):
        cr, attrs, ri, rl = m.group(1), m.group(2), m.group(3), m.group(4)
        assert ri == rl == cr, f"N{cr}: несовпадение строк {ri}/{rl}"
        rep["N"] += 1
        return f'<c r="N{cr}"{attrs}><f>ROUND(I{cr}*L{cr},2)</f>'
    def o_sub(m):
        cr, attrs, ri, rm = m.group(1), m.group(2), m.group(3), m.group(4)
        assert ri == rm == cr, f"O{cr}: несовпадение строк {ri}/{rm}"
        rep["O"] += 1
        return f'<c r="O{cr}"{attrs}><f>ROUND(I{cr}*M{cr},2)</f>'
    xml = re.sub(
        r'<c r="N(\d+)"([^>]*)><f>ROUND\(\'Смета ЕР \(общ\) \'!\$I(\d+)\*\'Смета ЕР \(общ\) \'!\$L(\d+),2\)</f>',
        n_sub, xml)
    xml = re.sub(
        r'<c r="O(\d+)"([^>]*)><f>ROUND\(\'Смета ЕР \(общ\) \'!\$I(\d+)\*\'Смета ЕР \(общ\) \'!\$M(\d+),2\)</f>',
        o_sub, xml)

    # остаточная самоквалификация допустима только в столбце E (CONCATENATE)
    leftover = xml.count(SHEET_QUAL)
    e_expected = 1532 * 3 + 139 * 3  # 3-уровневые + 4-уровневые (C — без квалификации)
    return xml, rep, leftover, e_expected

def main():
    with zipfile.ZipFile(SRC) as z:
        names = z.namelist()
        infos = {i.filename: i for i in z.infolist()}
        data = {n: z.read(n) for n in names}

    xml = data["xl/worksheets/sheet1.xml"].decode("utf-8")
    new_xml, rep, leftover, e_expected = patch_sheet(xml)
    data["xl/worksheets/sheet1.xml"] = new_xml.encode("utf-8")

    print("=== применено ===")
    print(f"  закрепление областей : {rep['freeze']}  (ySplit=14, topLeftCell=A15)")
    print(f"  <outlinePr>          : {rep['outlinePr']}  (summaryBelow=0)")
    print(f"  outlineLevel на строках: {rep['rows']}")
    lv = classify_levels(xml)
    print(f"      уровень 1 (заголовки/подытоги): {sum(1 for v in lv.values() if v == 1)}")
    print(f"      уровень 2 (позиции)          : {sum(1 for v in lv.values() if v == 2)}")
    print(f"      уровень 0 (Разделы)          : {len(SECTION_ROWS)}")
    print(f"  N-формулы упрощены     : {rep['N']}  (ждали 1611)")
    print(f"  O-формулы упрощены     : {rep['O']}  (ждали 1611)")
    print(f"  остаток '{SHEET_QUAL}' : {leftover}  (ждали {e_expected} — только столбец E)")
    assert rep["N"] == 1611 and rep["O"] == 1611, "число N/O-правок не 1611"
    assert leftover == e_expected, "неожиданный остаток самоквалификации"

    # запись: копия -> временный -> замена
    tmp = DST + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for n in names:
            zi = zipfile.ZipInfo(n, date_time=infos[n].date_time)
            zi.compress_type = zipfile.ZIP_DEFLATED
            zi.external_attr = infos[n].external_attr
            z.writestr(zi, data[n])
    os.replace(tmp, DST)
    print(f"\nзаписано: {DST}")

if __name__ == "__main__":
    main()
