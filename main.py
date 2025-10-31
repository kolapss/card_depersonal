# -*- coding: utf-8 -*-
import os
import sys
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText
import win32com.client as win32

APP_TITLE = "Обезличивание платежных карточек (Excel)"
TARGET_PHRASE = "Выплата заработной платы по ведомости"
REPLACEMENT = "ФИО <...><...>"
SHEET_DEFAULT = "1"  # "1" (первый лист) или имя листа


def safe_name_in_dir(output_dir, src_filename, suffix="__anon"):
    name, ext = os.path.splitext(src_filename)
    cand = os.path.join(output_dir, f"{name}{ext or '.xls'}")
    if not os.path.exists(cand):
        return cand

    cand = os.path.join(output_dir, f"{name}{suffix}{ext or '.xls'}")
    if not os.path.exists(cand):
        return cand

    i = 1
    while True:
        c = os.path.join(output_dir, f"{name}{suffix}_{i}{ext or '.xls'}")
        if not os.path.exists(c):
            return c
        i += 1


def fileformat_for_ext(path):
    ext = os.path.splitext(path)[1].lower()
    return 51 if ext == ".xlsx" else 56  # 51=.xlsx, 56=.xls


def process_with_excel(excel, input_path, sheet_selector, out_dir, logger):
    try:
        wb = excel.Workbooks.Open(os.path.abspath(input_path))
        if sheet_selector.strip().isdigit():
            ws = wb.Worksheets(int(sheet_selector))
        else:
            ws = wb.Worksheets(sheet_selector)

        xlUp = -4162
        last_row = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row  # B

        changed = 0
        for r in range(1, last_row + 1):
            v = ws.Cells(r, 2).Value  # B
            if v is not None and TARGET_PHRASE in str(v):
                ws.Cells(r, 3).Value = REPLACEMENT  # C
                changed += 1

        src_filename = os.path.basename(input_path)
        output_path = safe_name_in_dir(out_dir, src_filename)
        fileformat = fileformat_for_ext(output_path)

        wb.SaveAs(os.path.abspath(output_path), FileFormat=fileformat)
        return output_path, changed

    finally:
        try:
            wb.Close(SaveChanges=False)
        except Exception as e:
            logger(f"⚠️ Ошибка закрытия книги: {e}")


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("880x560")
        self.minsize(820, 520)

        info = (
            "Выберите Excel-файлы (.xls/.xlsx).\n"
            f"Во 2-м столбце ищется: «{TARGET_PHRASE}».\n"
            f"В 3-й столбец будет записано: «{REPLACEMENT}».\n"
            "Форматирование сохраняется."
        )
        tk.Label(self, text=info, justify="left").pack(fill="x", padx=12, pady=(12, 8))

        frm_params = tk.Frame(self)
        frm_params.pack(fill="x", padx=12)

        tk.Label(frm_params, text="Лист (номер или имя):").grid(row=0, column=0, sticky="w")
        self.ent_sheet = tk.Entry(frm_params, width=18)
        self.ent_sheet.insert(0, SHEET_DEFAULT)
        self.ent_sheet.grid(row=0, column=1, padx=(6, 18), sticky="w")

        frm_out = tk.Frame(self)
        frm_out.pack(fill="x", padx=12, pady=(8, 6))

        tk.Label(frm_out, text="Папка для сохранения:").pack(side="left")
        self.out_dir_var = tk.StringVar(value="")
        tk.Entry(frm_out, textvariable=self.out_dir_var).pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(frm_out, text="Выбрать…", command=self.choose_out_dir).pack(side="left")

        frm_files = tk.Frame(self)
        frm_files.pack(fill="both", expand=True, padx=12, pady=8)

        ttk.Button(frm_files, text="Выбрать файлы…", command=self.add_files).pack(anchor="w")
        self.lst = tk.Listbox(frm_files, height=10, selectmode="extended")
        self.lst.pack(fill="both", expand=True, pady=6)

        frm_actions = tk.Frame(self)
        frm_actions.pack(fill="x", padx=12, pady=8)

        ttk.Button(frm_actions, text="Удалить выбранные", command=self.remove_selected).pack(side="left")
        self.btn_run = ttk.Button(frm_actions, text="Обезличить", command=self.run_processing)
        self.btn_run.pack(side="right")

        self.txt = ScrolledText(self, height=10, state="disabled", wrap="word")
        self.txt.pack(fill="both", expand=False, padx=12, pady=(0, 12))

        self.pb = ttk.Progressbar(self, mode="determinate")
        self.pb.pack(fill="x", padx=12, pady=(0, 12))


    def choose_out_dir(self):
        path = filedialog.askdirectory()
        if path:
            self.out_dir_var.set(path)


    def log(self, msg):
        self.txt.configure(state="normal")
        self.txt.insert("end", msg + "\n")
        self.txt.see("end")
        self.txt.configure(state="disabled")
        self.update_idletasks()


    def add_files(self):
        paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xls *.xlsx")])
        for p in paths:
            if p not in self.lst.get(0, "end"):
                self.lst.insert("end", p)

        if not self.out_dir_var.get() and paths:
            base = os.path.dirname(paths[0])
            out = os.path.join(base, "anon_output")
            os.makedirs(out, exist_ok=True)
            self.out_dir_var.set(out)


    def remove_selected(self):
        for idx in reversed(self.lst.curselection()):
            self.lst.delete(idx)


    def run_processing(self):
        files = self.lst.get(0, "end")
        if not files:
            messagebox.showinfo(APP_TITLE, "Добавьте файлы.")
            return

        out_dir = self.out_dir_var.get().strip()
        if not out_dir:
            messagebox.showinfo(APP_TITLE, "Выберите папку для сохранения.")
            return

        os.makedirs(out_dir, exist_ok=True)

        sheet_selector = self.ent_sheet.get().strip() or "1"

        try:
            excel = win32.gencache.EnsureDispatch("Excel.Application")
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Ошибка запуска Excel:\n{e}")
            return

        excel.Visible = False
        excel.DisplayAlerts = False

        self.pb.config(maximum=len(files), value=0)
        self.btn_run.config(state="disabled")
        self.log("=== Запуск ===")

        ok = 0
        fail = 0
        total = 0

        try:
            for i, f in enumerate(files, start=1):
                self.log(f"{i}/{len(files)}: {f}")
                try:
                    out, changed = process_with_excel(excel, f, sheet_selector, out_dir, self.log)
                    self.log(f"  → {changed} замен")
                    self.log(f"  → {out}\n")
                    ok += 1
                    total += changed
                except Exception as e:
                    fail += 1
                    self.log(f"  ✖ Ошибка: {e}\n")

                self.pb.config(value=i)
                self.update_idletasks()
                time.sleep(0.05)
        finally:
            try:
                excel.Quit()
            except:
                pass
            self.btn_run.config(state="normal")

        self.log("=== Готово ===")
        self.log(f"Успешно: {ok}, ошибок: {fail}, замен всего: {total}")

        messagebox.showinfo(APP_TITLE, f"Готово.\nУспешно: {ok}\nОшибок: {fail}\nЗаменено: {total}")


def main():
    arch = "64-bit" if sys.maxsize > 2**32 else "32-bit"
    print(f"Python: {arch}")

    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()