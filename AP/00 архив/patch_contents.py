import sys, re
sys.stdout.reconfigure(encoding='utf-8')

html_path = r'd:\diploma\АП\netlify_deploy\index.html'

# New table of contents from _vkr_text.txt
toc_lines = [
    ("ОГЛАВЛЕНИЕ", True),
    ("ВВЕДЕНИЕ\t3", False),
    ("Глава 1. Обзор состояния вопроса\t6", False),
    ("1.2 Обзор состояния вопроса в отечественных нормах.\t14", False),
    ("1.3. Обзор состояния вопроса в зарубежных нормах\t16", False),
    ("1.4. Кривые предельного загружения\t21", False),
    ("1.5. Критическая нагрузка\t29", False),
    ("1.6. Обоснование расчетной схемы.\t31", False),
    ("1.7 Статический расчет.\t32", False),
    ("1.8. Конструктивный расчет.\t43", False),
    ("1.9. Выводы по главе 1.\t43", False),
    ("Глава 2. Экспериментальная часть и совершенствование методики расчета ферм из ЛСТК.\t44", False),
    ("2.1. Численное моделирование образца для эксперимента\t44", False),
    ("2.2. Натурные испытания образцов\t46", False),
    ("2.3. Совершенствование методики расчета несущих элементов рамы из ЛСТК\t55", False),
    ("2.4. Алгоритм расчета в ANSYS Workbench/Multiphysics.\t58", False),
    ("2.5. Статический расчет фермы с помощью объемных конечных элементов.\t59", False),
    ("2.6. Результаты статического расчета.\t63", False),
    ("2.7. Выводы по главе 2\t66", False),
    ("Глава 3. Численное моделирование в ПК ANSYS фермы из ЛСТК на устойчивость и модифицированной фермы\t67", False),
    ("3.1. Расчет на устойчивость фермы из ЛСТК\t67", False),
    ("3.2. Сравнение результатов расчета объемной и стержневой моделей\t69", False),
    ("3.3. Расчет модифицированной фермы\t70", False),
    ("3.4. Экономические показатели\t74", False),
    ("3.5. Выводы по главе 3\t76", False),
    ("ЗАКЛЮЧЕНИЕ\t78", False),
    ("БИБЛИОГРАФИЧЕСКИЙ СПИСОК\t80", False),
    ("ПРИЛОЖЕНИЯ\t83", False),
    ("ПРИЛОЖЕНИЕ А. Результаты статического расчета фермы\t83", False),
]

# Build replacement HTML paragraphs
new_paras = ""
for text, is_header in toc_lines:
    # Format: replace tab with spaces for display
    display = text.replace("\t", "  ")
    cls = "normal"
    new_paras += f'<p><span class="{cls}">{display} </span></p>'

html = open(html_path, encoding='utf-8').read()

# Find the СОДЕРЖАНИЕ block start: <p><span ...>СОДЕРЖАНИЕ
start_marker = html.find('<p><span class="plagiat">СОДЕРЖАНИЕ')
if start_marker == -1:
    start_marker = html.find('<p><span class="normal">СОДЕРЖАНИЕ')

# Find where the ToC ends — right before the ВВЕДЕНИЕ body paragraph
# The body ВВЕДЕНИЕ is: <p><span class="plagiat">ВВЕДЕНИЕ <!----></span></p>
# (different from "Введение4" in the ToC)
# Look for ВВЕДЕНИЕ as a standalone heading after the литература entry
end_marker_str = '<p><span class="plagiat">ВВЕДЕНИЕ <!----></span></p>'
end_marker = html.find(end_marker_str, start_marker)

if start_marker == -1 or end_marker == -1:
    print(f"Маркеры не найдены: start={start_marker}, end={end_marker}")
else:
    old_block = html[start_marker:end_marker]
    html = html[:start_marker] + new_paras + html[end_marker:]
    open(html_path, 'w', encoding='utf-8').write(html)
    print("Готово! СОДЕРЖАНИЕ заменено на ОГЛАВЛЕНИЕ из VKR.")
    print(f"Удалено символов: {len(old_block)}, добавлено: {len(new_paras)}")
