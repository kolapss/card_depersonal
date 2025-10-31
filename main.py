import os
import sys
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import win32com.client as win32

APP_TITLE = "Обезличивание платежных карточек (Excel)"
TARGET_PHRASE_DEFAULT = "Выплата заработной платы по ведомости"
REPLACEMENT_DEFAULT = "ФИО <...><...>"
SHEET_DEFAULT = "1"  # можно указать "1" (первый лист) или имя листа

def safe_save_as_name(src_path, suffix="__anon"):
    folder, base = os.path.split(src_path)
    name, ext = os.path.splitext(base)
    out = os.path.join(folder, f"{name}{suffix}{ext or '.xls'}")
    if not os.path.exists(out):
        return out
    # если файл есть — добавляем счетчик
    i = 1
    while True:
        cand = os.path.join(folder, f"{name}{suffix}_{i}{ext or '.xls'}")
        if not os.path.exists(cand):
            return cand
        i += 1

def fileformat_for_ext(path):
    # Excel constants: 56=.xls, 51=.xlsx
    ext = os.path.splitext(path)[1].lower()
    if ext == ".xlsx":
        return 51
    return 56

def process_with_excel(excel, input_path, sheet_selector, target_phrase, replacement, logger):
    """
    Обрабатывает один файл через уже запущенный Excel (COM), чтобы не терять формат.
    Возвращает (output_path, changed_count).
    """
    try:
        wb = excel.Workbooks.Open(os.path.abspath(input_path))
        # выбор листа: по номеру (1-based) или по имени
        if sheet_selector.strip().isdigit():
            ws = wb.Worksheets(int(sheet_selector))
        else:
            ws = wb.Worksheets(sheet_selector)

        xlUp = -4162
        last_row = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row  # столбец B
        changed = 0
        for r in range(1, last_row + 1):
            v = ws.Cells(r, 2).Value  # B
            if v is not None and target_phrase in str(v):
                ws.Cells(r, 3).Value = replacement  # C
                changed += 1

        output_path = safe_save_as_name(input_path, suffix="__anon")
        fileformat = fileformat_for_ext(output_path)
        wb.SaveAs(os.path.abspath(output_path), FileFormat=fileformat)
        return output_path, changed
    finally:
        # закрываем без сохранения изменений в исходнике
        try:
            wb.Close(SaveChanges=False)
        except Exception as e:
            logger(f"⚠️ Не удалось корректно закрыть книгу: {e}")

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("800x520")
        self.minsize(720, 480)

        # Верхний блок — пояснение
        info = (
            "Выберите один или несколько Excel-файлов (.xls/.xlsx).\n"
            "Программа найдёт строки, где во 2-м столбце (B) есть фраза:\n"
            f"«{TARGET_PHRASE_DEFAULT}» и запишет в 3-й столбец (C): «{REPLACEMENT_DEFAULT}».\n"
            "Результат сохраняется рядом с исходником с суффиксом __anon, формат и стили сохраняются."
        )
        self.lbl = tk.Label(self, text=info, justify="left", anchor="w")
        self.lbl.pack(fill="x", padx=12, pady=(12, 8))

        # Параметры
        frm_params = tk.Frame(self)
        frm_params.pack(fill="x", padx=12, pady=4)

        tk.Label(frm_params, text="Лист (номер или имя):").grid(row=0, column=0, sticky="w")
        self.ent_sheet = tk.Entry(frm_params, width=18)
        self.ent_sheet.insert(0, SHEET_DEFAULT)
        self.ent_sheet.grid(row=0, column=1, padx=(6, 18), sticky="w")

        tk.Label(frm_params, text="Искомая фраза:").grid(row=0, column=2, sticky="w")
        self.ent_phrase = tk.Entry(frm_params)
        self.ent_phrase.insert(0, TARGET_PHRASE_DEFAULT)
        self.ent_phrase.grid(row=0, column=3, padx=(6, 18), sticky="we")

        tk.Label(frm_params, text="Замена (в столбец C):").grid(row=0, column=4, sticky="w")
        self.ent_repl = tk.Entry(frm_params)
        self.ent_repl.insert(0, REPLACEMENT_DEFAULT)
        self.ent_repl.grid(row=0, column=5, padx=(6, 0), sticky="we")

        frm_params.columnconfigure(3, weight=1)
        frm_params.columnconfigure(5, weight=1)

        # Список файлов + кнопки
        frm_files = tk.Frame(self)
        frm_files.pack(fill="both", expand=True, padx=12, pady=8)

        self.btn_add = ttk.Button(frm_files, text="Выбрать файлы…", command=self.add_files)
        self.btn_add.pack(anchor="w")

        self.lst = tk.Listbox(frm_files, height=10, selectmode="extended")
        self.lst.pack(fill="both", expand=True, pady=6)

        frm_actions = tk.Frame(self)
        frm_actions.pack(fill="x", padx=12, pady=8)

        self.btn_clear = ttk.Button(frm_actions, text="Удалить выбранные", command=self.remove_selected)
        self.btn_clear.pack(side="left")

        self.btn_run = ttk.Button(frm_actions, text="Обезличить", command=self.run_processing)
        self.btn_run.pack(side="right")

        # Лог
        self.txt = tk.Text(self, height=10, state="disabled")
        self.txt.pack(fill="both", expand=False, padx=12, pady=(0, 12))

        # Прогресс
        self.pb = ttk.Progressbar(self, mode="determinate")
        self.pb.pack(fill="x", padx=12, pady=(0, 12))

    def log(self, msg):
        self.txt.configure(state="normal")
        self.txt.insert("end", msg + "\n")
        self.txt.see("end")
        self.txt.configure(state="disabled")
        self.update_idletasks()

    def add_files(self):
        paths = filedialog.askopenfilenames(
            title="Выберите Excel-файлы",
            filetypes=[("Excel files", "*.xls *.xlsx"), ("All files", "*.*")]
        )
        for p in paths:
            if p and p not in self.lst.get(0, "end"):
                self.lst.insert("end", p)

    def remove_selected(self):
        # удаляем с конца, чтобы индексы не смещались
        for idx in reversed(self.lst.curselection()):
            self.lst.delete(idx)

    def run_processing(self):
        files = list(self.lst.get(0, "end"))
        if not files:
            messagebox.showinfo(APP_TITLE, "Сначала выберите файлы.")
            return

        phrase = self.ent_phrase.get().strip()
        repl = self.ent_repl.get().strip()
        sheet_selector = self.ent_sheet.get().strip() or "1"

        # Проверка Excel/pywin32
        try:
            excel = win32.gencache.EnsureDispatch("Excel.Application")
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Не удалось запустить Excel (COM).\n\n{e}\n\n"
                                            "Убедитесь, что установлен Excel и пакет pywin32.")
            return

        excel.Visible = False
        excel.DisplayAlerts = False

        self.pb.configure(maximum=len(files), value=0)
        self.log("=== Старт обработки ===")
        self.btn_run.config(state="disabled")

        total_changed = 0
        ok = 0
        fail = 0

        try:
            for i, path in enumerate(files, start=1):
                self.log(f"Обработка {i}/{len(files)}: {path}")
                try:
                    out_path, changed = process_with_excel(
                        excel, path, sheet_selector, phrase, repl, self.log
                    )
                    self.log(f"  → Изменено строк: {changed}")
                    self.log(f"  → Сохранено: {out_path}\n")
                    total_changed += changed
                    ok += 1
                except Exception as e:
                    self.log(f"  ✖ Ошибка: {e}\n")
                    fail += 1

                self.pb.configure(value=i)
                self.update_idletasks()
                time.sleep(0.05)
        finally:
            try:
                excel.Quit()
            except Exception:
                pass
            self.btn_run.config(state="normal")

        self.log("=== Готово ===")
        self.log(f"Успешно: {ok}, ошибок: {fail}, всего замен: {total_changed}")

        if fail == 0:
            messagebox.showinfo(APP_TITLE, f"Готово.\nУспешно: {ok}\nВсего замен: {total_changed}")
        else:
            messagebox.showwarning(APP_TITLE, f"Завершено с ошибками.\nУспешно: {ok}\nОшибок: {fail}")

def main():
    # Проверка на соответствие разрядности Office/Python – частая причина ошибок COM
    # Не строго обязательно, но полезно предупредить.
    if sys.maxsize > 2**32:
        arch = "64-bit"
    else:
        arch = "32-bit"
    print(f"Python: {arch}. Для стабильной работы желательно, чтобы разрядность Excel совпадала.")

    app = App()
    app.mainloop()

if __name__ == "__main__":
    main()