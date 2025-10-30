import os
import win32com.client as win32

INPUT = r"Карточка 20501 с 01.01.2025-24.10.2025.xls"   # исходный .xls
OUTPUT = r"output.xls"                  # куда сохранить (можно .xls или .xlsx)
SHEET_NAME_OR_INDEX = 1                 # 1-й лист (можно указать имя строки, напр. "Лист1")
TARGET_PHRASE = "Выплата заработной платы по ведомости"
REPLACEMENT = "ФИО <...><...>"

# форматы Excel: 56=.xls, 51=.xlsx
FILEFORMAT = 56 if OUTPUT.lower().endswith(".xls") else 51

excel = win32.gencache.EnsureDispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

try:
    wb = excel.Workbooks.Open(os.path.abspath(INPUT))
    ws = wb.Worksheets(SHEET_NAME_OR_INDEX)

    # найти последнюю строку во 2-м столбце (B)
    xlUp = -4162
    last_row = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row

    changed = 0
    # перебор строк: проверяем столбец B (2), пишем в C (3)
    for r in range(1, last_row + 1):
        v = ws.Cells(r, 2).Value
        if v is not None and TARGET_PHRASE in str(v):
            ws.Cells(r, 3).Value = REPLACEMENT
            changed += 1

    # сохранить КОПИЮ в новый файл, сохранив весь формат
    wb.SaveAs(os.path.abspath(OUTPUT), FileFormat=FILEFORMAT)
    # Альтернатива: wb.SaveCopyAs(...) чтобы не трогать исходник вообще

    print(f"Готово. Изменено строк: {changed}. Файл: {OUTPUT}")
finally:
    # аккуратно закрываем Excel
    wb.Close(SaveChanges=False)
    excel.Quit()