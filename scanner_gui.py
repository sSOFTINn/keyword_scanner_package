# scanner_gui.py
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
from datetime import datetime
import os
import string
import queue
import csv

from scanner_core import KeywordScanner


LANGS = {
    "Українська": {
        "title": "Сканер ключових слів",
        "start": "Почати сканування",
        "progress": "Прогрес: {count} знайдено",
        "done": "Сканування завершено! Звіт збережено у scan_report.docx",
        "file": "Файл",
        "keywords": "Ключові слова / умови",
        "lang_label": "Мова / Language:",
        "mode_label": "Режим сканування:",
        "mode_fast": "Швидкий",
        "mode_deep": "Повний",
        "mode_names": "Тільки назви",
        "patterns_label": "File List для назв (*.txt):",
        "patterns_button": "Обрати...",
        "regex_label": "Regex по вмісту:",
        "regex_case": "Враховувати регістр (вміст)",
        "name_regex_label": "Regex по назві файлу:",
        "name_regex_case": "Врах. регістр (назви)",
        "log_start": "=== Початок сканування {ts} ===",
        "drives_label": "Диск для сканування:",
        "drives_all": "Усі диски",
        "preview_header": "Файл:",
        "preview_no_content": "Попередній перегляд недоступний (архів або помилка читання).",
        "preview_names_only": "Пошук тільки по назвах: вміст файлу не аналізувався.",
        "export_txt": "Експорт TXT",
        "export_csv": "Експорт CSV",
        "export_done": "Експорт виконано.",
    },
    "English": {
        "title": "Keyword Scanner",
        "start": "Start scan",
        "progress": "Progress: {count} found",
        "done": "Scan complete! Report saved as scan_report.docx",
        "file": "File",
        "keywords": "Keywords / patterns",
        "lang_label": "Мова / Language:",
        "mode_label": "Scan mode:",
        "mode_fast": "Fast",
        "mode_deep": "Deep",
        "mode_names": "Names only",
        "patterns_label": "File List for names (*.txt):",
        "patterns_button": "Browse...",
        "regex_label": "Content regex:",
        "regex_case": "Match case (content)",
        "name_regex_label": "File name regex:",
        "name_regex_case": "Match case (names)",
        "log_start": "=== Scan started at {ts} ===",
        "drives_label": "Drive to scan:",
        "drives_all": "All drives",
        "preview_header": "File:",
        "preview_no_content": "Preview not available (archive or read error).",
        "preview_names_only": "Names-only search: file content was not scanned.",
        "export_txt": "Export TXT",
        "export_csv": "Export CSV",
        "export_done": "Export complete.",
    },
}


class ScannerApp:
    def __init__(self, root):
        self.language = "Українська"
        self.root = root
        self.strings = LANGS[self.language]

        self.root.title(self.strings["title"])
        self.root.geometry("1000x600")
        self.root.resizable(True, True)

        self.patterns_path = None
        self.scan_mode = "deep"
        self.scanner: KeywordScanner | None = None
        self.selected_drives: list[str] = []

        self.available_drives = self.get_drives()
        self.result_data: dict[str, tuple[str, list[str]]] = {}

        self.result_queue: queue.Queue = queue.Queue()
        self.scanning = False

        self.create_widgets()

    # ---------- ДИСКИ ----------

    def get_drives(self):
        drives = []
        for letter in string.ascii_uppercase:
            d = f"{letter}:\\"
            if os.path.exists(d):
                drives.append(d)
        return drives

    # ---------- UI ----------

    def create_widgets(self):
        # Мова
        lang_frame = ttk.Frame(self.root)
        lang_frame.pack(pady=5, fill="x", padx=10)

        self.lang_label = ttk.Label(lang_frame, text=self.strings["lang_label"])
        self.lang_label.pack(side="left")

        self.lang_select = ttk.Combobox(
            lang_frame, values=list(LANGS.keys()), state="readonly", width=15
        )
        self.lang_select.set(self.language)
        self.lang_select.pack(side="left", padx=5)
        self.lang_select.bind("<<ComboboxSelected>>", self.change_language)

        # Диск
        drives_frame = ttk.Frame(self.root)
        drives_frame.pack(pady=5, fill="x", padx=10)

        self.drives_label = ttk.Label(drives_frame, text=self.strings["drives_label"])
        self.drives_label.pack(side="left")

        self.drive_combo = ttk.Combobox(
            drives_frame, state="readonly", width=20
        )
        values = [self.strings["drives_all"]] + self.available_drives
        self.drive_combo["values"] = values
        self.drive_combo.set(self.strings["drives_all"])
        self.drive_combo.pack(side="left", padx=5)

        # Вкладки
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="x", padx=10, pady=5)

        self.tab_content = ttk.Frame(self.notebook)
        self.tab_names = ttk.Frame(self.notebook)

        self.notebook.add(self.tab_content, text="Вміст / Content")
        self.notebook.add(self.tab_names, text="Назви / Names")

        self.create_content_tab()
        self.create_names_tab()

        # Кнопка старту + прогрес + експорт
        control_frame = ttk.Frame(self.root)
        control_frame.pack(fill="x", padx=10, pady=5)

        self.start_button = ttk.Button(
            control_frame, text=self.strings["start"], command=self.start_scan
        )
        self.start_button.pack(side="left")

        self.export_txt_button = ttk.Button(
            control_frame, text=self.strings["export_txt"], command=self.export_txt
        )
        self.export_txt_button.pack(side="left", padx=5)

        self.export_csv_button = ttk.Button(
            control_frame, text=self.strings["export_csv"], command=self.export_csv
        )
        self.export_csv_button.pack(side="left", padx=5)

        self.progress_label = ttk.Label(
            control_frame, text=self.strings["progress"].format(count=0)
        )
        self.progress_label.pack(side="right")

        self.progress = ttk.Progressbar(self.root, mode="indeterminate")
        self.progress.pack(fill="x", padx=20, pady=5)

        # Дві рухомі панелі
        self.paned = ttk.PanedWindow(self.root, orient="horizontal")
        self.paned.pack(fill="both", expand=True, padx=10, pady=5)

        # Ліва – таблиця
        left_frame = ttk.Frame(self.paned)
        self.paned.add(left_frame, weight=1)

        columns = ("name", "path")
        self.results_tree = ttk.Treeview(
            left_frame, columns=columns, show="headings", selectmode="browse"
        )
        self.results_tree.heading("name", text="Ім'я файлу")
        self.results_tree.heading("path", text="Шлях")
        self.results_tree.column("name", width=220, anchor="w")
        self.results_tree.column("path", width=500, anchor="w")

        tree_scroll_y = ttk.Scrollbar(
            left_frame, orient="vertical", command=self.results_tree.yview
        )
        self.results_tree.configure(yscrollcommand=tree_scroll_y.set)

        self.results_tree.pack(side="left", fill="both", expand=True)
        tree_scroll_y.pack(side="right", fill="y")

        self.results_tree.bind("<<TreeviewSelect>>", self.on_result_select)
        self.results_tree.bind("<Double-1>", self.on_result_double_click)

        # Права – прев'ю
        right_frame = ttk.Frame(self.paned)
        self.paned.add(right_frame, weight=1)

        self.preview_text = tk.Text(right_frame, wrap="word")
        preview_scroll_y = ttk.Scrollbar(
            right_frame, orient="vertical", command=self.preview_text.yview
        )
        self.preview_text.configure(yscrollcommand=preview_scroll_y.set)

        self.preview_text.pack(side="left", fill="both", expand=True)
        preview_scroll_y.pack(side="right", fill="y")

    def create_content_tab(self):
        mode_frame = ttk.Frame(self.tab_content)
        mode_frame.pack(pady=5, fill="x", padx=10)
        self.mode_label = ttk.Label(mode_frame, text=self.strings["mode_label"])
        self.mode_label.pack(side="left")

        self.mode_select = ttk.Combobox(
            mode_frame,
            values=[
                self.strings["mode_fast"],
                self.strings["mode_deep"],
            ],
            state="readonly",
            width=20,
        )
        self.mode_select.set(self.strings["mode_deep"])
        self.mode_select.pack(side="left", padx=5)

        patterns_frame = ttk.Frame(self.tab_content)
        patterns_frame.pack(pady=5, fill="x", padx=10)
        self.patterns_label = ttk.Label(
            patterns_frame, text=self.strings["patterns_label"]
        )
        self.patterns_label.pack(side="left")

        self.patterns_entry = ttk.Entry(patterns_frame, width=45)
        self.patterns_entry.pack(side="left", padx=5)

        self.patterns_button = ttk.Button(
            patterns_frame,
            text=self.strings["patterns_button"],
            command=self.choose_patterns_file,
        )
        self.patterns_button.pack(side="left")

        regex_frame = ttk.Frame(self.tab_content)
        regex_frame.pack(pady=5, fill="x", padx=10)
        self.regex_label = ttk.Label(regex_frame, text=self.strings["regex_label"])
        self.regex_label.pack(side="left")

        self.regex_entry = ttk.Entry(regex_frame, width=35)
        self.regex_entry.pack(side="left", padx=5)

        self.match_case_var = tk.BooleanVar(value=False)
        self.regex_case_check = ttk.Checkbutton(
            regex_frame,
            text=self.strings["regex_case"],
            variable=self.match_case_var,
        )
        self.regex_case_check.pack(side="left")

    def create_names_tab(self):
        n_frame = ttk.Frame(self.tab_names)
        n_frame.pack(pady=10, fill="x", padx=10)

        self.name_regex_label = ttk.Label(
            n_frame, text=self.strings["name_regex_label"]
        )
        self.name_regex_label.pack(side="left")

        self.name_regex_entry = ttk.Entry(n_frame, width=40)
        self.name_regex_entry.pack(side="left", padx=5)

        self.name_match_case_var = tk.BooleanVar(value=False)
        self.name_regex_case_check = ttk.Checkbutton(
            n_frame,
            text=self.strings["name_regex_case"],
            variable=self.name_match_case_var,
        )
        self.name_regex_case_check.pack(side="left")

        hint = ttk.Label(
            self.tab_names,
            text="* Тут використовується File List (файл умов) + regex тільки по назвах файлів.",
            foreground="gray",
        )
        hint.pack(anchor="w", padx=12, pady=2)

    # ---------- МОВА ----------

    def change_language(self, event=None):
        self.language = self.lang_select.get()
        self.strings = LANGS[self.language]

        self.root.title(self.strings["title"])
        self.lang_label.config(text=self.strings["lang_label"])
        self.drives_label.config(text=self.strings["drives_label"])
        self.start_button.config(text=self.strings["start"])
        self.export_txt_button.config(text=self.strings["export_txt"])
        self.export_csv_button.config(text=self.strings["export_csv"])

        self.mode_label.config(text=self.strings["mode_label"])
        self.mode_select["values"] = [
            self.strings["mode_fast"],
            self.strings["mode_deep"],
        ]
        if self.scan_mode == "fast":
            self.mode_select.set(self.strings["mode_fast"])
        else:
            self.mode_select.set(self.strings["mode_deep"])

        self.patterns_label.config(text=self.strings["patterns_label"])
        self.patterns_button.config(text=self.strings["patterns_button"])
        self.regex_label.config(text=self.strings["regex_label"])
        self.regex_case_check.config(text=self.strings["regex_case"])
        self.name_regex_label.config(text=self.strings["name_regex_label"])
        self.name_regex_case_check.config(text=self.strings["name_regex_case"])

        current_drive = self.drive_combo.get()
        values = [self.strings["drives_all"]] + self.available_drives
        self.drive_combo["values"] = values
        if current_drive in self.available_drives:
            self.drive_combo.set(current_drive)
        else:
            self.drive_combo.set(self.strings["drives_all"])

        self.progress_label.config(
            text=self.strings["progress"].format(
                count=len(self.scanner.matches) if self.scanner else 0
            )
        )

    # ---------- ДОПОМІЖНІ ----------

    def choose_patterns_file(self):
        path = filedialog.askopenfilename(
            title="Вибір File List для назв",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
        )
        if path:
            self.patterns_path = path
            self.patterns_entry.delete(0, tk.END)
            self.patterns_entry.insert(0, path)

    def get_selected_drives(self):
        value = self.drive_combo.get()
        if value == self.strings["drives_all"] or not value:
            return self.available_drives
        return [value]

    # ---------- СТАРТ СКАНУ ----------

    def start_scan(self):
        current_tab = self.notebook.index(self.notebook.select())
        self.selected_drives = self.get_selected_drives()

        # очистка попередніх результатів
        self.result_data.clear()
        for item in self.results_tree.get_children():
            self.results_tree.delete(item)
        self.preview_text.delete("1.0", tk.END)
        self.progress_label.config(
            text=self.strings["progress"].format(count=0)
        )

        # параметри
        if current_tab == 0:
            self.prepare_content_scan()
        else:
            self.prepare_names_scan()

        self.start_button.config(state="disabled")
        self.progress.start()
        self.scanning = True
        self.root.after(50, self.process_queue)

        with open("scan_log.txt", "a", encoding="utf-8") as log:
            log.write(
                self.strings["log_start"].format(
                    ts=datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                )
                + "\n"
            )
            log.write(f"Drives: {', '.join(self.selected_drives)}\n\n")

        thread = threading.Thread(target=self.run_scan, daemon=True)
        thread.start()

    def prepare_content_scan(self):
        selected_mode = self.mode_select.get()
        if selected_mode == self.strings["mode_fast"]:
            self.scan_mode = "fast"
        else:
            self.scan_mode = "deep"

        self.patterns_path = self.patterns_entry.get().strip() or None
        regex_pattern = self.regex_entry.get().strip() or None
        match_case = bool(self.match_case_var.get())

        self.scanner = KeywordScanner(
            "dictionary.txt",
            filename_patterns_path=self.patterns_path,
            mode=self.scan_mode,
            regex_pattern=regex_pattern,
            match_case=match_case,
        )

    def prepare_names_scan(self):
        self.scan_mode = "names"
        self.patterns_path = self.patterns_entry.get().strip() or None
        name_regex = self.name_regex_entry.get().strip() or None
        name_match_case = bool(self.name_match_case_var.get())

        self.scanner = KeywordScanner(
            "dictionary.txt",
            filename_patterns_path=self.patterns_path,
            mode="names",
            name_regex_pattern=name_regex,
            name_match_case=name_match_case,
        )

    # ---------- ПОТІК СКАНУ ----------

    def run_scan(self):
        try:
            self.scanner.scan_selected_drives(
                self.selected_drives,
                update_callback=self.result_queue.put,  # кидаємо в чергу
            )
            self.scanner.generate_report()
        finally:
            self.scanning = False
            self.root.after(0, self.finish_scan_ui)

    def finish_scan_ui(self):
        self.progress.stop()
        self.start_button.config(state="normal")
        messagebox.showinfo(self.strings["title"], self.strings["done"])

    def process_queue(self):
        processed = 0
        while processed < 50:
            try:
                results = self.result_queue.get_nowait()
            except queue.Empty:
                break
            self.update_log(results)
            processed += 1

        if self.scanning:
            self.root.after(50, self.process_queue)

    # ---------- ОНОВЛЕННЯ РЕЗУЛЬТАТІВ ----------

    def update_log(self, results):
        if not self.scanner:
            return

        for path, keywords in results:
            name = path.split("->")[-1].strip()
            name = os.path.basename(name)

            item_id = self.results_tree.insert(
                "", "end", values=(name, path)
            )
            self.result_data[item_id] = (path, keywords)
            self.scanner.matches.append((path, keywords))

        self.progress_label.config(
            text=self.strings["progress"].format(count=len(self.scanner.matches))
        )

    # ---------- ПРЕВ'Ю ТА ВІДКРИТТЯ ----------

    def on_result_select(self, event=None):
        item_id = self.results_tree.focus()
        if not item_id:
            return
        path, keywords = self.result_data.get(item_id, (None, None))
        if not path:
            return
        self.show_preview(path, keywords)

    def on_result_double_click(self, event=None):
        item_id = self.results_tree.focus()
        if not item_id:
            return
        path, _ = self.result_data.get(item_id, (None, None))
        if not path:
            return
        real_path = path.split("->")[0].strip()
        try:
            os.startfile(real_path)
        except Exception as e:
            messagebox.showerror("Error", f"Cannot open file:\n{real_path}\n\n{e}")

    def show_preview(self, path: str, keywords: list[str]):
        self.preview_text.delete("1.0", tk.END)
        self.preview_text.insert(tk.END, f"{self.strings['preview_header']} {path}\n\n")

        if "->" in path:
            self.preview_text.insert(
                tk.END,
                self.strings["preview_no_content"] + "\n\n"
                + "Matches: " + ", ".join(keywords),
            )
            return

        if not os.path.isfile(path):
            self.preview_text.insert(
                tk.END,
                self.strings["preview_no_content"] + "\n\n"
                + "Matches: " + ", ".join(keywords),
            )
            return

        if self.scanner and self.scanner.mode == "names":
            self.preview_text.insert(
                tk.END,
                self.strings["preview_names_only"] + "\n\n"
                + "Name matches: " + ", ".join(keywords),
            )
            return

        ext = os.path.splitext(path)[1].lower()

        try:
            if ext == ".txt":
                with open(path, "r", encoding="utf-8", errors="ignore") as f:
                    content = f.read()
            elif ext == ".docx":
                content = self.scanner.extract_text_docx(path)
            elif ext == ".pdf":
                content = self.scanner.extract_text_pdf(path)
            else:
                content = ""
        except Exception:
            content = ""

        if not content:
            self.preview_text.insert(
                tk.END,
                self.strings["preview_no_content"] + "\n\n"
                + "Matches: " + ", ".join(keywords),
            )
            return

        plain_words = [
            k
            for k in keywords
            if not k.startswith("[NAME]")
            and not k.startswith("[NAME_REGEX]")
            and not k.startswith("[REGEX]")
        ]

        lines = content.splitlines()
        shown = 0

        for idx, line in enumerate(lines, start=1):
            line_show = False

            for w in plain_words:
                if self.scanner.match_case:
                    if w in line:
                        line_show = True
                        break
                else:
                    if w.lower() in line.lower():
                        line_show = True
                        break

            if not line_show and self.scanner.regex and self.scanner.regex.search(line):
                line_show = True

            if line_show:
                self.preview_text.insert(
                    tk.END, f"{idx}: {line}\n"
                )
                shown += 1
                if shown >= 30:
                    break

        if shown == 0:
            self.preview_text.insert(
                tk.END,
                "\n(Не вдалося знайти рядки з контент-збігами, але файл містить збіги по назві / словнику / regex.)"
            )

    # ---------- ЕКСПОРТ ----------

    def export_txt(self):
        if not self.result_data:
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
            title="Зберегти результати як TXT",
        )
        if not path:
            return
        with open(path, "w", encoding="utf-8") as f:
            for item_id, (fpath, keywords) in self.result_data.items():
                name = os.path.basename(fpath.split("->")[-1].strip())
                f.write(f"{name}\t{fpath}\t{', '.join(keywords)}\n")
        messagebox.showinfo(self.strings["title"], self.strings["export_done"])

    def export_csv(self):
        if not self.result_data:
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            title="Зберегти результати як CSV",
        )
        if not path:
            return
        with open(path, "w", encoding="utf-8", newline="") as f:
            writer = csv.writer(f, delimiter=";")
            writer.writerow(["Name", "Path", "Matches"])
            for item_id, (fpath, keywords) in self.result_data.items():
                name = os.path.basename(fpath.split("->")[-1].strip())
                writer.writerow([name, fpath, ", ".join(keywords)])
        messagebox.showinfo(self.strings["title"], self.strings["export_done"])


if __name__ == "__main__":
    import sys

    if getattr(sys, "frozen", False):
        os.chdir(os.path.dirname(sys.executable))

    root = tk.Tk()
    app = ScannerApp(root)
    root.mainloop()