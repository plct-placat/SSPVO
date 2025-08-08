import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import re
import os


class ExcelProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Сравнение данных: ТАНДЕМ vs СП")
        self.root.geometry("1200x700")
        self.root.resizable(True, True)

        self.file1_path = ""
        self.file2_path = ""
        self.file3_path = ""
        self.results = []
        self.third_data = []

        self.setup_ui()

    def setup_ui(self):
        """Настройка графического интерфейса"""
        title_label = tk.Label(self.root, text="Сравнение данных из систем", font=("Arial", 14, "bold"))
        title_label.pack(pady=10)

        file_frame = tk.LabelFrame(self.root, text="Источники данных", padx=10, pady=10)
        file_frame.pack(padx=20, pady=10, fill="x")

        tk.Label(file_frame, text="1. Выбранные конкурсы ЕПГУ:").grid(row=0, column=0, sticky="w", pady=2)
        self.file1_entry = tk.Entry(file_frame, width=50)
        self.file1_entry.grid(row=0, column=1, padx=5, pady=2, sticky="ew")
        tk.Button(file_frame, text="Выбрать...", command=self.browse_file1).grid(row=0, column=2, padx=5)

        tk.Label(file_frame, text="2. КЭШ выбранных конкурсов:").grid(row=1, column=0, sticky="w", pady=2)
        self.file2_entry = tk.Entry(file_frame, width=50)
        self.file2_entry.grid(row=1, column=1, padx=5, pady=2, sticky="ew")
        tk.Button(file_frame, text="Выбрать...", command=self.browse_file2).grid(row=1, column=2, padx=5)

        tk.Label(file_frame, text="3. Заявления и Специальности из СП:").grid(row=2, column=0, sticky="w", pady=2)
        self.file3_entry = tk.Entry(file_frame, width=50)
        self.file3_entry.grid(row=2, column=1, padx=5, pady=2, sticky="ew")
        tk.Button(file_frame, text="Выбрать...", command=self.browse_file3).grid(row=2, column=2, padx=5)

        file_frame.columnconfigure(1, weight=1)

        btn_frame = tk.Frame(self.root)
        btn_frame.pack(pady=10)

        self.btn_compare = tk.Button(btn_frame, text="Сравнить данные (ТАНДЕМ vs СП)", command=self.process_and_compare,
                                     bg="#4CAF50", fg="white", font=("Arial", 11, "bold"), width=30, height=2)
        self.btn_compare.pack(pady=5)

        self.btn_save = tk.Button(btn_frame, text="Сохранить результат", command=self.save_results,
                                  bg="#2196F3", fg="white", font=("Arial", 10, "bold"), width=30, height=2)
        self.btn_save.pack(pady=5)
        self.btn_save.config(state="disabled")

        progress_frame = tk.Frame(self.root)
        progress_frame.pack(padx=20, pady=5, fill="x")
        tk.Label(progress_frame, text="Статус:").pack(anchor="w")
        self.progress = ttk.Progressbar(progress_frame, orient="horizontal", mode="determinate", length=100)
        self.progress.pack(fill="x", pady=5)
        self.progress_label = tk.Label(progress_frame, text="Готов", fg="gray")
        self.progress_label.pack(anchor="w")

        result_frame = tk.LabelFrame(self.root, text="Результаты сравнения", padx=10, pady=10)
        result_frame.pack(padx=20, pady=10, fill="both", expand=True)

        self.notebook = ttk.Notebook(result_frame)
        self.notebook.pack(fill="both", expand=True)

        self.frame_only_in_results = tk.Frame(self.notebook)
        self.frame_only_in_third = tk.Frame(self.notebook)
        self.frame_diff_fio = tk.Frame(self.notebook)
        self.frame_diff_status = tk.Frame(self.notebook)

        self.notebook.add(self.frame_only_in_results, text="Только в ТАНДЕМЕ")
        self.notebook.add(self.frame_only_in_third, text="Только в СП")
        self.notebook.add(self.frame_diff_fio, text="Разные ФИО (одинаковые заявки)")
        self.notebook.add(self.frame_diff_status, text="Разные статусы (одинаковые заявки)")

        self.setup_comparison_tables()

        self.create_menu()

    def setup_comparison_tables(self):
        """Настройка таблиц для отображения результатов"""
        cols_base = ("ID специальности (конкурса)", "ID заявления", "ФИО", "Статус")

        tree1_frame = tk.Frame(self.frame_only_in_results)
        tree1_frame.pack(fill="both", expand=True, padx=5, pady=5)
        self.tree_only_in_results = ttk.Treeview(tree1_frame, columns=cols_base, show="headings", height=10)
        self.setup_tree(self.tree_only_in_results, cols_base)
        self.add_scrollbar(tree1_frame, self.tree_only_in_results)

        tree2_frame = tk.Frame(self.frame_only_in_third)
        tree2_frame.pack(fill="both", expand=True, padx=5, pady=5)
        self.tree_only_in_third = ttk.Treeview(tree2_frame, columns=cols_base, show="headings", height=10)
        self.setup_tree(self.tree_only_in_third, cols_base)
        self.add_scrollbar(tree2_frame, self.tree_only_in_third)

        fio_cols = ["ID специальности (конкурса)", "ID заявления", "ФИО (ТАНДЕМ)", "ФИО (СП)", "Статус (ТАНДЕМ)", "Статус (СП)"]
        tree3_frame = tk.Frame(self.frame_diff_fio)
        tree3_frame.pack(fill="both", expand=True, padx=5, pady=5)
        self.tree_diff_fio = ttk.Treeview(tree3_frame, columns=fio_cols, show="headings", height=10)
        for col in fio_cols:
            self.tree_diff_fio.heading(col, text=col)
            self.tree_diff_fio.column(col, width=140, anchor="w")
        self.add_scrollbar(tree3_frame, self.tree_diff_fio)

        status_cols = ["ID специальности (конкурса)", "ID заявления", "ФИО", "Статус (ТАНДЕМ)", "Статус (СП)"]
        tree4_frame = tk.Frame(self.frame_diff_status)
        tree4_frame.pack(fill="both", expand=True, padx=5, pady=5)
        self.tree_diff_status = ttk.Treeview(tree4_frame, columns=status_cols, show="headings", height=10)
        for col in status_cols:
            self.tree_diff_status.heading(col, text=col)
            self.tree_diff_status.column(col, width=140, anchor="w")
        self.add_scrollbar(tree4_frame, self.tree_diff_status)

    def setup_tree(self, tree, columns):
        """Настройка заголовков и ширины колонок для Treeview"""
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=140, anchor="w")

    def add_scrollbar(self, parent, tree):
        """Добавление скроллбара к Treeview"""
        scrollbar = tk.Scrollbar(parent, orient="vertical", command=tree.yview)
        tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        tree.pack(fill="both", expand=True)

    def create_menu(self):
        """Создание меню приложения"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Файл", menu=file_menu)
        file_menu.add_command(label="Сохранить результат", command=self.save_results)
        file_menu.add_separator()
        file_menu.add_command(label="Выход", command=self.root.quit)

    def browse_file1(self):
        """Выбор файла 1"""
        path = filedialog.askopenfilename(title="Выбранные конкурсы ЕПГУ", filetypes=[("Excel", "*.xls *.xlsx")])
        if path:
            self.file1_path = path
            self.file1_entry.delete(0, tk.END)
            self.file1_entry.insert(0, os.path.basename(path))

    def browse_file2(self):
        """Выбор файла 2"""
        path = filedialog.askopenfilename(title="КЭШ выбранных конкурсов", filetypes=[("Excel", "*.xls *.xlsx")])
        if path:
            self.file2_path = path
            self.file2_entry.delete(0, tk.END)
            self.file2_entry.insert(0, os.path.basename(path))

    def browse_file3(self):
        """Выбор файла 3"""
        path = filedialog.askopenfilename(title="Заявления и Специальности из СП", filetypes=[("Excel", "*.xlsx")])
        if path:
            self.file3_path = path
            self.file3_entry.delete(0, tk.END)
            self.file3_entry.insert(0, os.path.basename(path))

    def extract_numbers(self, cell):
        """Извлечение чисел из ячейки"""
        if pd.isna(cell) or cell is None:
            return []
        return re.findall(r'\d+', str(cell))

    def clean_text(self, cell):
        """Очистка текста от NaN и пробелов"""
        if pd.isna(cell) or cell is None:
            return "—"
        return str(cell).strip()

    def clean_status(self, status):
        """Удаление шаблона (число) из статуса"""
        status = self.clean_text(status)
        if status == "—":
            return status
        return re.sub(r'\s*\(\d+\)\s*$', '', status).strip()

    def normalize_fio(self, fio):
        """Нормализация ФИО: приведение к нижнему регистру и удаление лишних пробелов"""
        fio = self.clean_text(fio)
        if fio == "—":
            return fio
        return re.sub(r'\s+', ' ', fio).strip().lower()

    def process_column(self, file_path, num_col, fio_col, status_col, skip_empty_num=False):
        """Обработка столбцов из файлов ТАНДЕМ"""
        try:
            df = pd.read_excel(file_path, skiprows=4)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка чтения файла:\n{str(e)}")
            return []

        def col_index(letter):
            return ord(letter.upper()) - ord('A')

        cols = {
            'num': col_index(num_col),
            'fio': col_index(fio_col),
            'status': col_index(status_col)
        }
        max_idx = max(cols.values())
        if max_idx >= df.shape[1] or df.empty:
            messagebox.showwarning("Внимание", f"В файле {os.path.basename(file_path)} недостаточно столбцов.")
            return []

        results = []
        for _, row in df.iterrows():
            try:
                num_cell = row.iloc[cols['num']]
                if skip_empty_num and (pd.isna(num_cell) or str(num_cell).strip() == ""):
                    continue

                numbers = self.extract_numbers(num_cell)
                five_digit = next((n for n in numbers if len(n) == 5 and n[0] != '0'), None)
                other_num = next((n for n in numbers if n != five_digit), None)
                if not five_digit and numbers:
                    sorted_nums = sorted(numbers, key=lambda x: (abs(len(x)-5), x[0]=='0'))
                    five_digit = sorted_nums[0]

                fio = self.clean_text(row.iloc[cols['fio']])
                status = self.clean_status(row.iloc[cols['status']])

                results.append((five_digit or "—", other_num or "—", fio, status))
            except Exception:
                continue
        return results

    def process_files(self):
        """Обработка первых двух файлов (ТАНДЕМ)"""
        self.results = []
        if not self.file1_path and not self.file2_path:
            messagebox.showwarning("Внимание", "Не выбраны файлы ТАНДЕМ (ЕПГУ или КЭШ).")
            return False

        success = False
        if self.file1_path:
            data1 = self.process_column(self.file1_path, 'C', 'D', 'J', False)
            self.results.extend(data1)
            if data1:
                success = True
        if self.file2_path:
            data2 = self.process_column(self.file2_path, 'G', 'B', 'I', True)
            self.results.extend(data2)
            if data2:
                success = True

        if not success:
            messagebox.showwarning("Ошибка", "Не удалось обработать данные из ТАНДЕМ.")
            return False

        seen = set()
        unique = []
        for r in self.results:
            key = (r[0], r[1], self.normalize_fio(r[2]))
            if key not in seen:
                seen.add(key)
                unique.append(r)
        self.results = unique
        return True

    def load_third_file(self):
        """Загрузка третьего файла (СП)"""
        if not self.file3_path:
            messagebox.showwarning("Внимание", "Не выбран файл 'Заявления и Специальности из СП'.")
            return False

        try:
            df = pd.read_excel(self.file3_path)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось прочитать файл СП:\n{str(e)}")
            return False

        if df.shape[1] <= 32:
            messagebox.showerror("Ошибка", "Файл СП не содержит столбец AG (требуется минимум 33 столбца).")
            return False

        self.third_data = []
        for _, row in df.iterrows():
            try:
                five = self.clean_text(row.iloc[27])
                other = self.clean_text(row.iloc[15])
                fio = self.clean_text(row.iloc[2])
                status = self.clean_text(row.iloc[32])
                status = self.clean_status(status)
                self.third_data.append((five, other, fio, status))
            except Exception:
                continue

        seen = set()
        unique = []
        for r in self.third_data:
            key = (r[0], r[1], self.normalize_fio(r[2]))
            if key not in seen:
                seen.add(key)
                unique.append(r)
        self.third_data = unique
        return True

    def compare_files(self):
        """Сравнение данных и вывод результатов"""
        self.clear_all_trees()
        self.set_progress(20, "Сравнение данных...")

        key_tandem_by_nums = {(r[0], r[1]): r for r in self.results}
        key_sp_by_nums = {(t[0], t[1]): t for t in self.third_data}

        tandem_keys = set(key_tandem_by_nums.keys())
        sp_keys = set(key_sp_by_nums.keys())

        only_in_tandem = [key_tandem_by_nums[k] for k in tandem_keys - sp_keys]
        only_in_sp = [key_sp_by_nums[k] for k in sp_keys - tandem_keys]

        common_keys = tandem_keys & sp_keys
        diff_fio_list = []
        diff_status_list = []

        for key in common_keys:
            r = key_tandem_by_nums[key]
            t = key_sp_by_nums[key]
            if self.normalize_fio(r[2]) != self.normalize_fio(t[2]):
                diff_fio_list.append((*r, t[2], t[3]))
            elif r[3] != t[3]:
                diff_status_list.append((*r, t[3]))

        for r in only_in_tandem:
            self.tree_only_in_results.insert("", "end", values=r)
        for t in only_in_sp:
            self.tree_only_in_third.insert("", "end", values=t)
        for item in diff_fio_list:
            self.tree_diff_fio.insert("", "end", values=(item[0], item[1], item[2], item[4], item[3], item[5]))
        for item in diff_status_list:
            self.tree_diff_status.insert("", "end", values=(item[0], item[1], item[2], item[3], item[4]))

        self.set_progress(100, "Готово")
        self.btn_save.config(state="normal")
        messagebox.showinfo("Сравнение завершено", f"Результаты:\n"
                                                   f"Только в ТАНДЕМЕ: {len(only_in_tandem)}\n"
                                                   f"Только в СП: {len(only_in_sp)}\n"
                                                   f"Разные ФИО: {len(diff_fio_list)}\n"
                                                   f"Разные статусы: {len(diff_status_list)}")

    def process_and_compare(self):
        """Обработка и сравнение всех файлов"""
        self.clear_all_trees()
        self.set_progress(0, "Начало обработки...")
        self.root.update()

        if not self.process_files():
            self.set_progress(0, "Ошибка")
            return
        self.set_progress(50, "Загрузка СП...")
        self.root.update()

        if not self.load_third_file():
            self.set_progress(0, "Ошибка")
            return
        self.set_progress(75, "Сравнение...")
        self.root.update()

        self.compare_files()

    def set_progress(self, value, text):
        """Установка значения прогресс-бара"""
        self.progress["value"] = value
        self.progress_label.config(text=text)
        self.root.update_idletasks()

    def clear_all_trees(self):
        """Очистка всех таблиц"""
        trees = [self.tree_only_in_results, self.tree_only_in_third, self.tree_diff_fio, self.tree_diff_status]
        for tree in trees:
            for item in tree.get_children():
                tree.delete(item)
        self.btn_save.config(state="disabled")
        self.set_progress(0, "Готов")

    def save_results(self):
        """Сохранение результатов в Excel"""
        if not any([
            self.tree_only_in_results.get_children(),
            self.tree_only_in_third.get_children(),
            self.tree_diff_fio.get_children(),
            self.tree_diff_status.get_children()
        ]):
            messagebox.showwarning("Внимание", "Нет данных для сохранения.")
            return

        save_path = filedialog.asksaveasfilename(
            title="Сохранить результат сравнения",
            defaultextension=".xlsx",
            filetypes=[("Excel файлы", "*.xlsx")]
        )
        if not save_path:
            return

        try:
            with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                data1 = [self.tree_only_in_results.item(child)["values"] for child in self.tree_only_in_results.get_children()]
                df1 = pd.DataFrame(data1, columns=["ID специальности (конкурса)", "ID заявления", "ФИО", "Статус"]) if data1 else pd.DataFrame(columns=["ID специальности (конкурса)", "ID заявления", "ФИО", "Статус"])
                df1.to_excel(writer, sheet_name="Только в ТАНДЕМЕ", index=False)

                data2 = [self.tree_only_in_third.item(child)["values"] for child in self.tree_only_in_third.get_children()]
                df2 = pd.DataFrame(data2, columns=["ID специальности (конкурса)", "ID заявления", "ФИО", "Статус"]) if data2 else pd.DataFrame(columns=["ID специальности (конкурса)", "ID заявления", "ФИО", "Статус"])
                df2.to_excel(writer, sheet_name="Только в СП", index=False)

                data3 = [self.tree_diff_fio.item(child)["values"] for child in self.tree_diff_fio.get_children()]
                df3 = pd.DataFrame(data3, columns=["ID специальности (конкурса)", "ID заявления", "ФИО (ТАНДЕМ)", "ФИО (СП)", "Статус (ТАНДЕМ)", "Статус (СП)"]) if data3 else pd.DataFrame(columns=["ID специальности (конкурса)", "ID заявления", "ФИО (ТАНДЕМ)", "ФИО (СП)", "Статус (ТАНДЕМ)", "Статус (СП)"])
                df3.to_excel(writer, sheet_name="Разные ФИО", index=False)

                data4 = [self.tree_diff_status.item(child)["values"] for child in self.tree_diff_status.get_children()]
                df4 = pd.DataFrame(data4, columns=["ID специальности (конкурса)", "ID заявления", "ФИО", "Статус (ТАНДЕМ)", "Статус (СП)"]) if data4 else pd.DataFrame(columns=["ID специальности (конкурса)", "ID заявления", "ФИО", "Статус (ТАНДЕМ)", "Статус (СП)"])
                df4.to_excel(writer, sheet_name="Разные статусы", index=False)

            messagebox.showinfo("Сохранено", f"Результат сохранён:\n{os.path.basename(save_path)}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить файл:\n{str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelProcessorApp(root)
    root.mainloop()