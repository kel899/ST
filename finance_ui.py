import sqlite3
import datetime
from tkinter import *
from tkinter import ttk, messagebox
from tkcalendar import DateEntry
from database import Database
from constants import DEFAULT_CATEGORIES, logging

class FinanceUI:
    def __init__(self, notebook, app):
        self.app = app  # 接收主 Application 實例
        self.master = app.master
        self.status_var = app.status_var
        self.tab = ttk.Frame(notebook)
        notebook.add(self.tab, text="財務管理")
        self.db_lock = threading.Lock()
        self.setup_finance_tab()

    def setup_finance_tab(self):
        """設置財務管理標籤頁"""
        finance_container = ttk.Frame(self.tab)
        finance_container.pack(expand=True, fill=BOTH, padx=10, pady=10)

        button_frame = ttk.Frame(finance_container)
        button_frame.pack(fill=X, pady=10)
        ttk.Button(button_frame, text="新增支出", command=self.add_expense).pack(side=LEFT, padx=10)
        ttk.Button(button_frame, text="匯出支出", command=self.export_expenses).pack(side=LEFT, padx=10)

        expense_frame = ttk.LabelFrame(finance_container, text="支出記錄", padding=10)
        expense_frame.pack(fill=BOTH, expand=True)
        self.expense_tree = ttk.Treeview(expense_frame, columns=("日期", "類別", "金額", "備註", "淨收入"), show="headings", height=15)
        for col, width in [("日期", 150), ("類別", 150), ("金額", 150), ("備註", 300), ("淨收入", 150)]:
            self.expense_tree.heading(col, text=col)
            self.expense_tree.column(col, width=width, anchor='center')
        self.expense_tree.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar = ttk.Scrollbar(expense_frame, orient=VERTICAL, command=self.expense_tree.yview)
        self.expense_tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=RIGHT, fill=Y)
        self.expense_tree.bind("<Button-3>", self.popup_expense_menu)

        self.expense_menu = Menu(self.master, tearoff=0)
        self.expense_menu.add_command(label="編輯支出", command=self.edit_expense)
        self.expense_menu.add_command(label="刪除支出", command=self.delete_expense)

        self.load_finance_data()

    def add_expense(self):
        """新增支出"""
        win = Toplevel(self.master)
        win.title("新增支出")
        win.geometry("400x300")
        ttk.Label(win, text="日期：").grid(row=0, column=0, padx=10, pady=10)
        date_entry = DateEntry(win, width=15, date_pattern="yyyy-mm-dd", font=('微軟正黑體', 12))
        date_entry.grid(row=0, column=1)
        ttk.Label(win, text="類別：").grid(row=1, column=0, padx=10, pady=10)
        category_combo = ttk.Combobox(win, values=DEFAULT_CATEGORIES, width=20, font=('微軟正黑體', 12))
        category_combo.grid(row=1, column=1)
        ttk.Label(win, text="金額：").grid(row=2, column=0, padx=10, pady=10)
        amount_entry = ttk.Entry(win, width=15, font=('微軟正黑體', 12))
        amount_entry.grid(row=2, column=1)
        ttk.Label(win, text="備註：").grid(row=3, column=0, padx=10, pady=10)
        desc_entry = ttk.Entry(win, width=20, font=('微軟正黑體', 12))
        desc_entry.grid(row=3, column=1)

        def save():
            try:
                amount = float(amount_entry.get())
                with self.db_lock:
                    Database.execute('INSERT INTO Expenses (expense_date, category, amount, description) VALUES (?, ?, ?, ?)', 
                                    (date_entry.get_date().strftime("%Y-%m-%d"), category_combo.get(), amount, desc_entry.get()))
                self.load_finance_data()
                win.destroy()
                self.status_var.set("支出已新增")
                logging.info(f"新增支出: 日期 {date_entry.get_date().strftime('%Y-%m-%d')}, 金額 ${round(amount)}")
            except ValueError:
                messagebox.showerror("錯誤", "金額必須為數字")

        ttk.Button(win, text="保存", command=save).grid(row=4, column=0, columnspan=2, pady=20)
import sqlite3
import datetime
from tkinter import *
from tkinter import ttk, messagebox
from tkcalendar import DateEntry
from database import Database
from constants import DEFAULT_CATEGORIES, logging
import threading

class FinanceUI:
    def __init__(self, notebook, app):
        self.app = app  # 接收主 Application 實例
        self.master = app.master
        self.status_var = app.status_var
        self.tab = ttk.Frame(notebook)
        notebook.add(self.tab, text="財務管理")
        self.db_lock = threading.Lock()
        self.setup_finance_tab()

    def setup_finance_tab(self):
        """設置財務管理標籤頁"""
        finance_container = ttk.Frame(self.tab)
        finance_container.pack(expand=True, fill=BOTH, padx=10, pady=10)

        button_frame = ttk.Frame(finance_container)
        button_frame.pack(fill=X, pady=10)
        ttk.Button(button_frame, text="新增支出", command=self.add_expense).pack(side=LEFT, padx=10)
        ttk.Button(button_frame, text="匯出支出", command=self.export_expenses).pack(side=LEFT, padx=10)

        expense_frame = ttk.LabelFrame(finance_container, text="支出記錄", padding=10)
        expense_frame.pack(fill=BOTH, expand=True)
        self.expense_tree = ttk.Treeview(expense_frame, columns=("日期", "類別", "金額", "備註", "淨收入"), show="headings", height=15)
        for col, width in [("日期", 150), ("類別", 150), ("金額", 150), ("備註", 300), ("淨收入", 150)]:
            self.expense_tree.heading(col, text=col)
            self.expense_tree.column(col, width=width, anchor='center')
        self.expense_tree.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar = ttk.Scrollbar(expense_frame, orient=VERTICAL, command=self.expense_tree.yview)
        self.expense_tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=RIGHT, fill=Y)
        self.expense_tree.bind("<Button-3>", self.popup_expense_menu)

        self.expense_menu = Menu(self.master, tearoff=0)
        self.expense_menu.add_command(label="編輯支出", command=self.edit_expense)
        self.expense_menu.add_command(label="刪除支出", command=self.delete_expense)

        self.load_finance_data()

    def add_expense(self):
        """新增支出"""
        win = Toplevel(self.master)
        win.title("新增支出")
        win.geometry("400x300")
        ttk.Label(win, text="日期：").grid(row=0, column=0, padx=10, pady=10)
        date_entry = DateEntry(win, width=15, date_pattern="yyyy-mm-dd", font=('微軟正黑體', 12))
        date_entry.grid(row=0, column=1)
        ttk.Label(win, text="類別：").grid(row=1, column=0, padx=10, pady=10)
        category_combo = ttk.Combobox(win, values=DEFAULT_CATEGORIES, width=20, font=('微軟正黑體', 12))
        category_combo.grid(row=1, column=1)
        ttk.Label(win, text="金額：").grid(row=2, column=0, padx=10, pady=10)
        amount_entry = ttk.Entry(win, width=15, font=('微軟正黑體', 12))
        amount_entry.grid(row=2, column=1)
        ttk.Label(win, text="備註：").grid(row=3, column=0, padx=10, pady=10)
        desc_entry = ttk.Entry(win, width=20, font=('微軟正黑體', 12))
        desc_entry.grid(row=3, column=1)

        def save():
            try:
                amount = float(amount_entry.get())
                with self.db_lock:
                    Database.execute('INSERT INTO Expenses (expense_date, category, amount, description) VALUES (?, ?, ?, ?)', 
                                    (date_entry.get_date().strftime("%Y-%m-%d"), category_combo.get(), amount, desc_entry.get()))
                self.load_finance_data()
                win.destroy()
                self.status_var.set("支出已新增")
                logging.info(f"新增支出: 日期 {date_entry.get_date().strftime('%Y-%m-%d')}, 金額 ${round(amount)}")
            except ValueError:
                messagebox.showerror("錯誤", "金額必須為數字")

        ttk.Button(win, text="保存", command=save).grid(row=4, column=0, columnspan=2, pady=20)

    def edit_expense(self):
        """編輯支出"""
        selected = self.expense_tree.selection()
        if len(selected) != 1:
            messagebox.showwarning("警告", "請選擇一筆支出")
            return
        expense_id = self.expense_tree.item(selected[0], 'values')[0]  # 假設日期作為唯一標識，實際應使用 expense_id
        with self.db_lock:
            data = Database.execute('SELECT expense_date, category, amount, description FROM Expenses WHERE expense_date = ?', 
                                   (expense_id,), fetch=True)[0]

        win = Toplevel(self.master)
        win.title("編輯支出")
        win.geometry("400x300")
        ttk.Label(win, text="日期：").grid(row=0, column=0, padx=10, pady=10)
        date_entry = DateEntry(win, width=15, date_pattern="yyyy-mm-dd", font=('微軟正黑體', 12))
        date_entry.set_date(data['expense_date'])
        date_entry.grid(row=0, column=1)
        ttk.Label(win, text="類別：").grid(row=1, column=0, padx=10, pady=10)
        category_combo = ttk.Combobox(win, values=DEFAULT_CATEGORIES, width=20, font=('微軟正黑體', 12))
        category_combo.set(data['category'])
        category_combo.grid(row=1, column=1)
        ttk.Label(win, text="金額：").grid(row=2, column=0, padx=10, pady=10)
        amount_entry = ttk.Entry(win, width=15, font=('微軟正黑體', 12))
        amount_entry.insert(0, data['amount'])
        amount_entry.grid(row=2, column=1)
        ttk.Label(win, text="備註：").grid(row=3, column=0, padx=10, pady=10)
        desc_entry = ttk.Entry(win, width=20, font=('微軟正黑體', 12))
        desc_entry.insert(0, data['description'] or "")
        desc_entry.grid(row=3, column=1)

        def save():
            try:
                amount = float(amount_entry.get())
                with self.db_lock:
                    Database.execute('UPDATE Expenses SET expense_date = ?, category = ?, amount = ?, description = ? WHERE expense_date = ?', 
                                    (date_entry.get_date().strftime("%Y-%m-%d"), category_combo.get(), amount, desc_entry.get(), expense_id))
                self.load_finance_data()
                win.destroy()
                self.status_var.set("支出已更新")
                logging.info(f"編輯支出: 日期 {date_entry.get_date().strftime('%Y-%m-%d')}, 新金額 ${round(amount)}")
            except ValueError:
                messagebox.showerror("錯誤", "金額必須為數字")

        ttk.Button(win, text="保存", command=save).grid(row=4, column=0, columnspan=2, pady=20)

    def delete_expense(self):
        """刪除支出"""
        selected = self.expense_tree.selection()
        if not selected:
            messagebox.showwarning("警告", "請選擇支出")
            return
        if messagebox.askyesno("確認", "確定刪除所選支出？"):
            dates = [self.expense_tree.item(item, 'values')[0] for item in selected]
            with self.db_lock:
                with sqlite3.connect('secret_time.db') as conn:
                    cursor = conn.cursor()
                    cursor.executemany('DELETE FROM Expenses WHERE expense_date = ?', [(date,) for date in dates])
                    conn.commit()
            self.load_finance_data()
            self.status_var.set(f"已刪除 {len(dates)} 筆支出")
            logging.info(f"刪除支出: 日期 {', '.join(dates)}")

    def popup_expense_menu(self, event):
        """支出列表右鍵選單"""
        if self.expense_tree.identify_row(event.y):
            self.expense_tree.selection_set(self.expense_tree.identify_row(event.y))
            self.expense_menu.post(event.x_root, event.y_root)

    def load_finance_data(self):
        """載入財務數據，異步實現"""
        def load_in_background():
            with self.db_lock:
                expenses = Database.execute('''
                    SELECT e.expense_date, e.category, e.amount, e.description,
                           (SELECT SUM(ct.price) - SUM(e2.amount) 
                            FROM Customer_Treatments ct 
                            LEFT JOIN Expenses e2 ON date(ct.treatment_date) = date(e2.expense_date)
                            WHERE date(ct.treatment_date) = date(e.expense_date)) as net_income
                    FROM Expenses e
                    ORDER BY e.expense_date DESC
                ''', fetch=True)
            self.master.after(0, lambda: self.update_finance_data(expenses))

        self.expense_tree.delete(*self.expense_tree.get_children())
        threading.Thread(target=load_in_background, daemon=True).start()

    def update_finance_data(self, expenses):
        """更新財務數據到 UI"""
        if not expenses:
            self.status_var.set("無支出數據")
            return
        for e in expenses:
            net_income = f"${round(e['net_income'] or 0)}" if e['net_income'] is not None else "N/A"
            self.expense_tree.insert("", "end", values=(e['expense_date'], e['category'], f"${round(e['amount'])}", 
                                                        e['description'] or "", net_income))
        self.status_var.set("財務數據已載入")

    def export_expenses(self):
        """匯出支出資料"""
        def export_task():
            try:
                with self.db_lock:
                    expenses = Database.execute('SELECT expense_date, category, amount, description FROM Expenses ORDER BY expense_date DESC', fetch=True)
                # 假設 pandas 已導入
                import pandas as pd
                df = pd.DataFrame(expenses)
                df.to_excel('expenses_export.xlsx', index=False)
                self.status_var.set("支出資料已匯出至 expenses_export.xlsx")
            except Exception as e:
                messagebox.showerror("錯誤", f"匯出失敗：{str(e)}")
                logging.error(f"支出匯出失敗：{str(e)}")
        threading.Thread(target=export_task, daemon=True).start()