import sqlite3
import datetime
from tkinter import *
from tkinter import ttk, messagebox
from tkcalendar import DateEntry
import pandas as pd
from matplotlib import font_manager
import matplotlib.pyplot as plt
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import re
import threading
import pickle
from database import Database
from business_logic import BusinessLogic
from constants import DEFAULT_CATEGORIES, logging

# 設置 Matplotlib 支援中文
font_manager.fontManager.addfont('C:/Windows/Fonts/msjh.ttc')  # 微軟正黑體
plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei']
plt.rcParams['axes.unicode_minus'] = False

class Application:
    def __init__(self, master):
        self.master = master
        self.master.title("美容工作室管理系統")
        self.master.geometry("1400x800")
        self.master.configure(bg="#E8ECEF")
        self.status_var = StringVar(value="歡迎使用美容工作室管理系統")
        self.suggestion_window = None
        self.db_lock = threading.Lock()
        # 初始化分頁相關屬性
        self.current_page = 1
        self.page_size = 50
        # 明確調用資料庫初始化
        try:
            Database.initialize_database()
            self.status_var.set("資料庫初始化完成")
        except Exception as e:
            messagebox.showerror("錯誤", f"資料庫初始化失敗：{str(e)}")
            logging.error(f"資料庫初始化失敗：{str(e)}")
            return
        self.setup_main_ui()

    def setup_main_ui(self):
        """設置主介面，包含標籤頁和狀態列"""
        self.notebook = ttk.Notebook(self.master)
        self.client_tab = ttk.Frame(self.notebook)
        self.finance_tab = ttk.Frame(self.notebook)
        self.stats_tab = ttk.Frame(self.notebook)
        
        self.notebook.add(self.client_tab, text="客戶管理")
        self.notebook.add(self.finance_tab, text="財務管理")
        self.notebook.add(self.stats_tab, text="統計分析")
        self.notebook.pack(expand=True, fill=BOTH, padx=10, pady=10)

        status_bar = ttk.Label(self.master, textvariable=self.status_var, relief=SUNKEN, anchor=W, font=('微軟正黑體', 12))
        status_bar.pack(side=BOTTOM, fill=X, padx=10, pady=5)

        self.setup_client_tab()
        self.setup_finance_tab()
        self.setup_stats_tab()

        self.master.bind("<Control-s>", lambda e: self.add_record())
        self.master.bind("<Control-d>", lambda e: self.delete_customer())

    def setup_client_tab(self):
        """設置客戶管理標籤頁"""
        self.client_container = ttk.Frame(self.client_tab)
        self.client_container.pack(expand=True, fill=BOTH, padx=10, pady=10)

        left_frame = ttk.Frame(self.client_container)
        left_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        input_frame = ttk.LabelFrame(left_frame, text="新增客戶記錄", padding=15)
        input_frame.pack(fill=X, pady=(0, 10))
        self.create_input_frame(input_frame)
        price_frame = ttk.LabelFrame(left_frame, text="療程價目表", padding=15)
        price_frame.pack(fill=BOTH, expand=True)
        self.create_price_list(price_frame)

        list_frame = ttk.LabelFrame(self.client_container, text="客戶與療程資訊", padding=15)
        list_frame.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")
        self.create_client_list_frame(list_frame)

        self.client_container.grid_columnconfigure(0, weight=1)
        self.client_container.grid_columnconfigure(1, weight=2)
        self.client_container.grid_rowconfigure(0, weight=1)

    def create_input_frame(self, parent):
        """創建客戶輸入區域"""
        ttk.Label(parent, text="客戶姓名：").grid(row=0, column=0, padx=10, pady=10, sticky=W)
        self.name_entry = ttk.Entry(parent, width=25, font=('微軟正黑體', 12))
        self.name_entry.grid(row=0, column=1, padx=10, pady=10)
        self.name_entry.bind("<KeyRelease>", self.on_name_entry_key_release)

        ttk.Label(parent, text="聯絡方式：").grid(row=1, column=0, padx=10, pady=10, sticky=W)
        self.contact_var = StringVar(value="WhatsApp")
        ttk.Radiobutton(parent, text="WhatsApp", variable=self.contact_var, value="WhatsApp").grid(row=1, column=1, padx=5, pady=5, sticky=W)
        ttk.Radiobutton(parent, text="Instagram", variable=self.contact_var, value="Instagram").grid(row=1, column=2, padx=5, pady=5, sticky=W)

        ttk.Label(parent, text="療程類型：").grid(row=2, column=0, padx=10, pady=10, sticky=W)
        self.treatment_var = StringVar(value="Einxel Plus")
        treatments = ["Einxel Plus", "熱能氣化 - 單次脫墨/疣", "熱能氣化 - 多次脫墨/疣", "熱能氣化 - 任脫套餐", "組合療程：Einxel Plus膠原修復針 6次包套", "補脫療程"]
        ttk.Combobox(parent, textvariable=self.treatment_var, values=treatments, state="readonly", width=20, font=('微軟正黑體', 12)).grid(row=2, column=1, padx=10, pady=10)
        self.treatment_var.trace('w', self.update_input_fields)

        ttk.Label(parent, text="消費金額：").grid(row=3, column=0, padx=10, pady=10, sticky=W)
        self.price_entry = ttk.Entry(parent, width=15, font=('微軟正黑體', 12), validate="key", validatecommand=(self.master.register(self.validate_price), '%P'))
        self.price_entry.grid(row=3, column=1, padx=10, pady=10)
        self.price_entry.bind("<KeyRelease>", self.on_price_entry_key_release)

        ttk.Label(parent, text="服務日期：").grid(row=4, column=0, padx=10, pady=10, sticky=W)
        self.date_entry = DateEntry(parent, width=15, date_pattern="yyyy-mm-dd", locale='zh_tw', font=('微軟正黑體', 12))
        self.date_entry.grid(row=4, column=1, padx=10, pady=10)

        ttk.Label(parent, text="Einxel Plus：").grid(row=5, column=0, padx=10, pady=10, sticky=W)
        self.neck_var = IntVar()
        self.neck_check = ttk.Checkbutton(parent, text="加做頸部 (+300)", variable=self.neck_var)
        self.neck_check.grid(row=5, column=1, padx=10, pady=10, sticky=W)
        self.package_var = IntVar()
        self.package_check = ttk.Checkbutton(parent, text="Einxel Plus Package", variable=self.package_var)
        self.package_check.grid(row=5, column=2, padx=10, pady=10, sticky=W)

        self.retouch_var = IntVar()
        self.retouch_check = ttk.Checkbutton(parent, text="補脫 (半年內免費)", variable=self.retouch_var, state='disabled')
        self.retouch_check.grid(row=6, column=1, columnspan=2, padx=10, pady=10, sticky=W)

        ttk.Button(parent, text="新增記錄", command=self.add_record).grid(row=7, column=0, columnspan=3, pady=20)

    def validate_price(self, value):
        """實時驗證價格輸入"""
        if value:
            try:
                float(value)
                self.price_entry.config(foreground="black")
                return True
            except ValueError:
                self.price_entry.config(foreground="red")
                return False
        return True

    def update_input_fields(self, *args):
        """根據療程類型更新輸入欄位狀態"""
        treatment = self.treatment_var.get()
        if treatment == "熱能氣化 - 單次脫墨/疣":
            self.neck_check.configure(state='disabled')
            self.package_check.configure(state='disabled')
            self.retouch_check.configure(state='disabled')
            self.neck_var.set(0)
            self.package_var.set(0)
            self.retouch_var.set(0)
            self.price_entry.config(state='normal')
        elif treatment == "補脫療程":
            self.neck_check.configure(state='disabled')
            self.package_check.configure(state='disabled')
            self.retouch_check.configure(state='disabled')
            self.neck_var.set(0)
            self.package_var.set(0)
            self.retouch_var.set(1)
            self.price_entry.delete(0, END)
            self.price_entry.insert(0, "0")
            self.price_entry.config(state='disabled')
        elif treatment in ["熱能氣化 - 多次脫墨/疣", "熱能氣化 - 任脫套餐"]:
            self.neck_check.configure(state='disabled')
            self.package_check.configure(state='disabled')
            self.retouch_check.configure(state='normal')
            self.neck_var.set(0)
            self.package_var.set(0)
            self.price_entry.config(state='normal')
        else:  # Einxel Plus 或 套餐
            self.neck_check.configure(state='normal' if treatment == "Einxel Plus" else 'disabled')
            self.package_check.configure(state='normal' if treatment.startswith("組合療程") else 'disabled')
            self.retouch_check.configure(state='disabled')
            self.neck_var.set(0)
            self.package_var.set(0)
            self.retouch_var.set(0)
            self.price_entry.config(state='normal')

    def create_price_list(self, parent):
        """創建療程價目表，新增捲動軸"""
        price_frame = ttk.Frame(parent)
        price_frame.pack(fill=BOTH, expand=True)
        self.price_list_tree = ttk.Treeview(price_frame, columns=("療程名稱", "Peak價格", "Non-Peak價格"), show="headings", height=10)
        self.price_list_tree.column("療程名稱", width=350, anchor='w')
        self.price_list_tree.column("Peak價格", width=100, anchor='center')
        self.price_list_tree.column("Non-Peak價格", width=100, anchor='center')
        self.price_list_tree.heading("療程名稱", text="名稱")
        self.price_list_tree.heading("Peak價格", text="PEAK價")
        self.price_list_tree.heading("Non-Peak價格", text="Non-PEAK價")
        self.price_list_tree.pack(side=LEFT, fill=BOTH, expand=True)

        scrollbar = ttk.Scrollbar(price_frame, orient=VERTICAL, command=self.price_list_tree.yview)
        self.price_list_tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=RIGHT, fill=Y)

        treatments = Database.execute('SELECT name, peak_price, non_peak_price FROM Treatments', fetch=True)
        for treatment in treatments:
            self.price_list_tree.insert("", "end", values=(
                treatment['name'],
                f"${round(treatment['peak_price'])}",
                f"${round(treatment['non_peak_price'])}" if treatment['non_peak_price'] != treatment['peak_price'] else "N/A"
            ))

    def create_client_list_frame(self, parent):
        """創建客戶列表和療程記錄區域"""
        filter_frame = ttk.Frame(parent)
        filter_frame.pack(fill=X, pady=5)
        ttk.Label(filter_frame, text="搜尋客戶：").pack(side=LEFT, padx=10)
        self.search_entry = ttk.Entry(filter_frame, width=20, font=('微軟正黑體', 12))
        self.search_entry.pack(side=LEFT, padx=10)
        ttk.Button(filter_frame, text="搜尋", command=self.search_clients).pack(side=LEFT, padx=5)
        ttk.Label(filter_frame, text="消費次數 ≥").pack(side=LEFT, padx=10)
        self.filter_treatments = ttk.Entry(filter_frame, width=10, font=('微軟正黑體', 12))
        self.filter_treatments.pack(side=LEFT, padx=10)
        ttk.Label(filter_frame, text="總金額 ≥").pack(side=LEFT, padx=10)
        self.filter_amount = ttk.Entry(filter_frame, width=10, font=('微軟正黑體', 12))
        self.filter_amount.pack(side=LEFT, padx=10)
        ttk.Button(filter_frame, text="篩選", command=self.apply_filter).pack(side=LEFT, padx=5)
        ttk.Button(filter_frame, text="從剪貼簿匯入", command=self.import_from_clipboard).pack(side=LEFT, padx=5)
        ttk.Button(filter_frame, text="匯出客戶", command=self.export_clients).pack(side=LEFT, padx=5)
        ttk.Button(filter_frame, text="查看匯入歷史", command=self.view_import_history).pack(side=LEFT, padx=5)

        date_filter_frame = ttk.Frame(parent)
        date_filter_frame.pack(fill=X, pady=5)
        ttk.Label(date_filter_frame, text="療程日期篩選：").pack(side=LEFT, padx=10)
        self.treatment_start_date = DateEntry(date_filter_frame, width=15, date_pattern="yyyy-mm-dd", locale='zh_tw', font=('微軟正黑體', 12))
        self.treatment_start_date.pack(side=LEFT, padx=10)
        ttk.Label(date_filter_frame, text="至").pack(side=LEFT, padx=5)
        self.treatment_end_date = DateEntry(date_filter_frame, width=15, date_pattern="yyyy-mm-dd", locale='zh_tw', font=('微軟正黑體', 12))
        self.treatment_end_date.pack(side=LEFT, padx=10)
        ttk.Button(date_filter_frame, text="篩選療程", command=self.apply_treatment_filter).pack(side=LEFT, padx=5)

        customer_frame = ttk.LabelFrame(parent, text="客戶列表", padding=10)
        customer_frame.pack(fill=BOTH, expand=True, pady=10)
        self.customer_tree = ttk.Treeview(customer_frame, columns=("ID", "姓名", "聯絡方式", "消費次數", "總金額", "剩餘次數", "上次消費"), show="headings", height=8)
        for col, width in [("ID", 80), ("姓名", 200), ("聯絡方式", 150), ("消費次數", 120), ("總金額", 150), ("剩餘次數", 120), ("上次消費", 150)]:
            self.customer_tree.heading(col, text=col, command=lambda c=col: self.sort_treeview(c))
            self.customer_tree.column(col, width=width, anchor='center')
        self.customer_tree.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar = ttk.Scrollbar(customer_frame, orient=VERTICAL, command=self.customer_tree.yview)
        self.customer_tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=RIGHT, fill=Y)
        self.customer_tree.bind("<Double-1>", lambda e: self.show_history())
        self.customer_tree.bind("<Button-3>", self.popup_customer_menu)

        treatment_frame = ttk.LabelFrame(parent, text="最新療程記錄", padding=10)
        treatment_frame.pack(fill=BOTH, expand=True, pady=10)
        self.treatment_tree = ttk.Treeview(treatment_frame, columns=("日期", "姓名", "項目", "時段", "頸部", "金額", "剩餘次數", "剩餘補脫"), show="headings", height=8)
        for col, width in [("日期", 150), ("姓名", 200), ("項目", 250), ("時段", 120), ("頸部", 80), ("金額", 150), ("剩餘次數", 120), ("剩餘補脫", 80)]:
            self.treatment_tree.heading(col, text=col)
            self.treatment_tree.column(col, width=width, anchor='center')
        self.treatment_tree.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar = ttk.Scrollbar(treatment_frame, orient=VERTICAL, command=self.treatment_tree.yview)
        self.treatment_tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=RIGHT, fill=Y)

        page_frame = ttk.Frame(parent)
        page_frame.pack(fill=X)
        ttk.Button(page_frame, text="上一頁", command=self.prev_page).pack(side=LEFT)
        ttk.Button(page_frame, text="下一頁", command=self.next_page).pack(side=LEFT)

        self.customer_menu = Menu(self.master, tearoff=0)
        self.customer_menu.add_command(label="編輯客戶", command=self.edit_customer)
        self.customer_menu.add_command(label="編輯療程", command=self.edit_treatment)
        self.customer_menu.add_command(label="刪除客戶", command=self.delete_customer)
        self.customer_menu.add_command(label="查看歷史", command=self.show_history)

        self.load_client_data()

    def setup_finance_tab(self):
        """設置財務管理標籤頁"""
        finance_container = ttk.Frame(self.finance_tab)
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

    def setup_stats_tab(self):
        """設置統計分析標籤頁"""
        stats_container = ttk.Frame(self.stats_tab)
        stats_container.pack(expand=True, fill=BOTH, padx=10, pady=10)

        filter_frame = ttk.LabelFrame(stats_container, text="篩選範圍", padding=10)
        filter_frame.pack(fill=X, pady=10)
        self.stats_range_var = StringVar(value="按月")
        ttk.Combobox(filter_frame, textvariable=self.stats_range_var, values=["按月", "按年"], state="readonly").pack(side=LEFT, padx=10)
        ttk.Label(filter_frame, text="選擇時間：").pack(side=LEFT, padx=10)
        self.stats_month_var = StringVar()
        self.stats_month_combo = ttk.Combobox(filter_frame, textvariable=self.stats_month_var, state="readonly", width=15, font=('微軟正黑體', 12))
        self.stats_month_combo.pack(side=LEFT, padx=10)
        self.stats_month_combo['values'] = BusinessLogic.get_available_months()
        if self.stats_month_combo['values']:
            self.stats_month_combo.current(0)
        ttk.Button(filter_frame, text="刷新統計", command=self.update_stats).pack(side=LEFT, padx=10)

        top_frame = ttk.LabelFrame(stats_container, text="客戶消費排行榜", padding=10)
        top_frame.pack(fill=X, pady=10)
        self.top_customers_tree = ttk.Treeview(top_frame, columns=("排名", "姓名", "總金額"), show="headings", height=5)
        for col, width in [("排名", 80), ("姓名", 300), ("總金額", 150)]:
            self.top_customers_tree.heading(col, text=col)
            self.top_customers_tree.column(col, width=width, anchor='center')
        self.top_customers_tree.pack(fill=X)

        self.trend_frame = ttk.LabelFrame(stats_container, text="收入與支出趨勢圖", padding=10)
        self.trend_frame.pack(fill=BOTH, expand=True, pady=10)
        self.pie_frame = ttk.LabelFrame(stats_container, text="療程類型分布圖", padding=10)
        self.pie_frame.pack(fill=BOTH, expand=True, pady=10)
        self.retention_frame = ttk.LabelFrame(stats_container, text="客戶保留率", padding=10)
        self.retention_frame.pack(fill=BOTH, expand=True, pady=10)

        self.update_stats()

    def on_name_entry_key_release(self, event):
        """處理客戶名稱輸入提示"""
        typed = self.name_entry.get().strip()
        if len(typed) >= 2:
            if self.suggestion_window is None or not self.suggestion_window.winfo_exists():
                self.show_name_suggestions(typed)

    def on_price_entry_key_release(self, event):
        """處理價格輸入時的頸部選項啟用"""
        try:
            price = float(self.price_entry.get())
            treatment = BusinessLogic.match_treatment(price)
            self.neck_check.configure(state='normal' if treatment and treatment['name'] == 'Einxel Plus膠原修復針' else 'disabled')
            self.neck_var.set(0 if not treatment or treatment['name'] != 'Einxel Plus膠原修復針' else self.neck_var.get())
        except ValueError:
            self.neck_check.configure(state='disabled')
            self.neck_var.set(0)

    def show_name_suggestions(self, prefix):
        """顯示客戶名稱建議"""
        suggestions = BusinessLogic.get_customer_names_matching(prefix)
        if suggestions:
            self.suggestion_window = Toplevel(self.master)
            self.suggestion_window.title("客戶名稱建議")
            self.suggestion_window.geometry("300x200")
            self.suggestion_window.transient(self.master)
            listbox = Listbox(self.suggestion_window, font=('微軟正黑體', 12), height=8)
            listbox.pack(fill=BOTH, expand=True, padx=10, pady=10)
            for name in suggestions:
                listbox.insert(END, name)
            listbox.bind("<Double-1>", lambda e: [self.name_entry.delete(0, END), self.name_entry.insert(0, listbox.get(listbox.curselection())), self.suggestion_window.destroy()])
            self.suggestion_window.geometry(f"+{self.name_entry.winfo_rootx()}+{self.name_entry.winfo_rooty() + 40}")
            self.status_var.set(f"找到 {len(suggestions)} 個匹配客戶")
            self.suggestion_window.protocol("WM_DELETE_WINDOW", lambda: [self.suggestion_window.destroy(), setattr(self, 'suggestion_window', None)])
        else:
            self.status_var.set("無匹配客戶名稱")

    def popup_customer_menu(self, event):
        """客戶列表右鍵選單"""
        if self.customer_tree.identify_row(event.y):
            self.customer_tree.selection_set(self.customer_tree.identify_row(event.y))
            self.customer_menu.post(event.x_root, event.y_root)

    def popup_expense_menu(self, event):
        """支出列表右鍵選單"""
        if self.expense_tree.identify_row(event.y):
            self.expense_tree.selection_set(self.expense_tree.identify_row(event.y))
            self.expense_menu.post(event.x_root, event.y_root)

    def add_record(self):
        """新增客戶記錄"""
        name, contact = self.name_entry.get().strip(), self.contact_var.get()
        try:
            price = float(self.price_entry.get())
            treatment_type = self.treatment_var.get()
            package = self.package_var.get()
            neck = self.neck_var.get()
            retouch = self.retouch_var.get()

            if treatment_type == "熱能氣化 - 單次脫墨/疣":
                treatment = Database.execute('SELECT * FROM Treatments WHERE name = ?', (treatment_type,), fetch=True)
                if not treatment:
                    raise ValueError(f"資料庫中未找到療程: {treatment_type}")
                treatment = dict(treatment[0])
                is_peak = 0
                neck = 0
                package = 0
                final_price = price
                retouch_parent_id = None
            elif treatment_type == "補脫療程":
                treatment = Database.execute('SELECT * FROM Treatments WHERE name = ?', (treatment_type,), fetch=True)
                if not treatment:
                    raise ValueError(f"資料庫中未找到療程: {treatment_type}")
                treatment = dict(treatment[0])
                customer_id = self.get_or_create_customer(name, contact)
                date = self.date_entry.get_date().strftime("%Y-%m-%d")
                parent_id = BusinessLogic.check_retouch_eligibility(customer_id, date)
                if not parent_id:
                    raise ValueError("半年內無可用首次療程或已使用過免費補脫")
                is_peak = 0
                neck = 0
                package = 0
                final_price = 0
                retouch_parent_id = parent_id
            elif treatment_type in ["熱能氣化 - 多次脫墨/疣", "熱能氣化 - 任脫套餐"]:
                if price > 0:
                    if price in [1600, 1800, 2800, 3800, 4800]:
                        treatment_name = "熱能氣化 - 10粒" if price == 1600 else \
                                        "熱能氣化 - 20粒" if price == 1800 else \
                                        "熱能氣化 - 40粒" if price == 2800 else \
                                        "熱能氣化 - 任脫 3800" if price == 3800 else \
                                        "熱能氣化 - 任脫 4800"
                        treatment = Database.execute('SELECT * FROM Treatments WHERE name = ?', (treatment_name,), fetch=True)
                        if not treatment:
                            raise ValueError(f"資料庫中未找到療程: {treatment_name}")
                        treatment = dict(treatment[0])
                        is_peak = 0
                        neck = 0
                        package = 0
                        final_price = price
                        retouch_parent_id = None
                    else:
                        raise ValueError(f"無匹配價格 {price} 的熱能氣化療程")
                elif price == 0 and retouch:
                    raise ValueError("請選擇 '補脫療程' 進行補脫操作")
                else:
                    raise ValueError("補脫必須使用 '補脫療程' 並設定價格為 0")
            else:  # Einxel Plus 或 套餐
                if package and price == 9926:
                    treatment = Database.execute('SELECT * FROM Treatments WHERE name = ?', ('組合療程：Einxel Plus膠原修復針 6次包套',), fetch=True)
                    if not treatment:
                        raise ValueError("資料庫中未找到 Einxel Plus 6次包套療程")
                    treatment = dict(treatment[0])
                elif package and price == 0:
                    treatment = Database.execute('SELECT * FROM Treatments WHERE name = ?', ('組合療程：Einxel Plus膠原修復針 6次包套',), fetch=True)
                    if not treatment:
                        raise ValueError("資料庫中未找到 Einxel Plus 6次包套療程")
                    treatment = dict(treatment[0])
                else:
                    treatment = BusinessLogic.match_treatment(price)
                    if not treatment:
                        raise ValueError(f"無匹配價格 {price} 的療程")
                    treatment = dict(treatment)
                is_peak = 1 if price == treatment['peak_price'] else 0
                neck = neck if treatment['name'] == 'Einxel Plus膠原修復針' else 0
                final_price = price + (neck * 300) if not package else price
                retouch_parent_id = None

            date = self.date_entry.get_date().strftime("%Y-%m-%d")
            customer_id = self.get_or_create_customer(name, contact)
            remaining, package_id = self.handle_package(treatment, customer_id, final_price, package)
            retouch_count = 1 if treatment['name'].startswith("熱能氣化") and not retouch_parent_id else 0

            with self.db_lock:
                Database.execute('''
                    INSERT INTO Customer_Treatments (customer_id, treatment_id, treatment_date, is_peak, neck_treatment, package_id, remaining_sessions, price, retouch_parent_id, remaining_retouch_count, import_id)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (customer_id, treatment['treatment_id'], date, is_peak, neck, package_id, remaining, final_price, retouch_parent_id, retouch_count, None))

            if retouch_parent_id:
                Database.execute('UPDATE Customer_Treatments SET remaining_retouch_count = 0 WHERE customer_treatment_id = ?', (retouch_parent_id,))

            if remaining == 0 and package:
                messagebox.showinfo("提示", "Einxel Plus 6次套餐已完成！")

            self.load_client_data()
            self.clear_input()
            self.status_var.set(f"已為 {name} 新增記錄，金額: ${round(final_price)}")
            logging.info(f"成功新增記錄: 客戶 {name}, 療程 {treatment['name']}, 金額 ${round(final_price)}")
        except ValueError as ve:
            messagebox.showerror("錯誤", str(ve))
        except Exception as e:
            messagebox.showerror("錯誤", f"新增失敗: {str(e)}")
            logging.error(f"新增記錄失敗: {str(e)}")

    def get_or_create_customer(self, name, contact):
        """獲取或創建客戶 ID，並處理重複姓名與聯絡方式"""
        duplicates = BusinessLogic.check_duplicate(name, contact)
        if duplicates:
            if messagebox.askyesno("確認", f"客戶 {name} ({contact}) 已存在，是否為同一人？"):
                return duplicates[0]['customer_id']
            else:
                win = Toplevel(self.master)
                win.title("輸入備註")
                win.geometry("300x150")
                ttk.Label(win, text="請輸入備註以區分客戶：").pack(pady=10)
                remark_entry = ttk.Entry(win, width=20, font=('微軟正黑體', 12))
                remark_entry.pack(pady=10)
                
                customer_id = [None]
                def save_remark():
                    remark = remark_entry.get().strip()
                    if not remark:
                        messagebox.showwarning("警告", "備註不能為空")
                        return
                    full_name = f"{name}({remark})"
                    existing = Database.execute('SELECT customer_id FROM Customers WHERE name = ? AND contact_method = ?', 
                                               (full_name, contact), fetch=True)
                    if existing:
                        customer_id[0] = existing[0]['customer_id']
                    else:
                        customer_id[0] = Database.execute('INSERT INTO Customers (name, contact_method, unique_mark) VALUES (?, ?, ?)', 
                                                         (full_name, contact, remark))
                    win.destroy()
                
                ttk.Button(win, text="確認", command=save_remark).pack(pady=10)
                win.wait_window()
                return customer_id[0]
        else:
            return Database.execute('INSERT INTO Customers (name, contact_method) VALUES (?, ?)', (name, contact))

    def handle_package(self, treatment, customer_id, price, package):
        """處理套裝療程邏輯，確保 treatment 是字典"""
        if not isinstance(treatment, dict):
            treatment = dict(treatment)
        if treatment['name'] != '組合療程：Einxel Plus膠原修復針 6次包套':
            return 0, None
        package_status = BusinessLogic.get_remaining_sessions(customer_id, treatment['name'])
        remaining_sessions = package_status['remaining_sessions'] if package_status else 0
        if price == 9926:
            max_package_id = Database.execute('SELECT MAX(package_id) FROM Customer_Treatments WHERE package_id IS NOT NULL', fetch=True)[0]
            new_package_id = (max_package_id['MAX(package_id)'] or 0) + 1
            return 5, new_package_id
        elif price == 0 and package and remaining_sessions > 0:
            with self.db_lock:
                Database.execute('UPDATE Customer_Treatments SET remaining_sessions = ? WHERE package_id = ?', 
                                (remaining_sessions - 1, package_status['package_id']))
            return remaining_sessions - 1, package_status['package_id']
        elif price == 0 and not package:
            raise ValueError("價格為 0 時必須勾選 Einxel Plus Package")
        else:
            raise ValueError("套裝療程價格或剩餘次數錯誤")

    def load_client_data(self):
        """載入客戶和療程數據，異步實現"""
        def load_in_background():
            with self.db_lock:
                offset = (self.current_page - 1) * self.page_size
                customers = Database.execute('''
                    SELECT c.customer_id, c.name, c.contact_method, COUNT(ct.customer_treatment_id) as count,
                           SUM(ct.price) as total, MAX(ct.remaining_sessions) as remaining,
                           MAX(ct.treatment_date) as last_date
                    FROM Customers c LEFT JOIN Customer_Treatments ct ON c.customer_id = ct.customer_id
                    GROUP BY c.customer_id, c.name, c.contact_method
                    ORDER BY last_date DESC LIMIT ? OFFSET ?
                ''', (self.page_size, offset), fetch=True)
                treatments = Database.execute('''
                    SELECT ct.treatment_date, c.name, t.name as treatment_name, ct.is_peak, ct.neck_treatment, ct.remaining_sessions,
                           ct.price as price, ct.remaining_retouch_count, t.can_add_neck, t.has_remaining_sessions
                    FROM Customer_Treatments ct 
                    JOIN Customers c ON ct.customer_id = c.customer_id
                    JOIN Treatments t ON ct.treatment_id = t.treatment_id
                    ORDER BY ct.treatment_date DESC LIMIT 20
                ''', fetch=True)
            self.master.after(0, lambda: self.update_client_data(customers, treatments))

        self.customer_tree.delete(*self.customer_tree.get_children())
        self.treatment_tree.delete(*self.treatment_tree.get_children())
        threading.Thread(target=load_in_background, daemon=True).start()

    def update_client_data(self, customers, treatments):
        if not customers:
            self.status_var.set("無客戶數據")
            return
        for c in customers:
            self.customer_tree.insert("", "end", values=(c['customer_id'], c['name'], c['contact_method'], c['count'] or 0, 
                                                         f"${round(c['total'] or 0)}", c['remaining'] or "N/A", c['last_date'] or "N/A"))
        if not treatments:
            self.status_var.set("無療程數據")
            return
        for t in treatments:
            period = "Peak" if t['is_peak'] and t['treatment_name'] == "Einxel Plus膠原修復針" and t['peak_price'] != t['non_peak_price'] else ""
            neck_status = "是" if t['neck_treatment'] and t['can_add_neck'] else ""
            remaining = t['remaining_sessions'] if t['has_remaining_sessions'] else ""
            retouch_count = f"剩餘補脫：{t['remaining_retouch_count']}" if t['treatment_name'].startswith("熱能氣化") else ""
            self.treatment_tree.insert("", "end", values=(
                t['treatment_date'], t['name'], t['treatment_name'], period, neck_status, f"${round(t['price'])}", 
                remaining, retouch_count
            ))
        self.status_var.set(f"載入第 {self.current_page} 頁客戶與療程數據")

    def edit_customer(self):
        """編輯客戶資訊"""
        selected = self.customer_tree.selection()
        if len(selected) != 1:
            messagebox.showwarning("警告", "請選擇一個客戶")
            return
        customer_id = self.customer_tree.item(selected[0], 'values')[0]
        data = Database.execute('SELECT name, contact_method, unique_mark FROM Customers WHERE customer_id = ?', (customer_id,), fetch=True)[0]

        win = Toplevel(self.master)
        win.title(f"編輯客戶 ID: {customer_id}")
        win.geometry("400x250")
        ttk.Label(win, text="姓名：").grid(row=0, column=0, padx=10, pady=10)
        name_entry = ttk.Entry(win, width=20, font=('微軟正黑體', 12))
        name_entry.insert(0, data['name'])
        name_entry.grid(row=0, column=1)
        ttk.Label(win, text="聯絡方式：").grid(row=1, column=0, padx=10, pady=10)
        contact_var = StringVar(value=data['contact_method'])
        ttk.Radiobutton(win, text="WhatsApp", variable=contact_var, value="WhatsApp").grid(row=1, column=1, sticky=W)
        ttk.Radiobutton(win, text="Instagram", variable=contact_var, value="Instagram").grid(row=1, column=2, sticky=W)
        ttk.Label(win, text="備註：").grid(row=2, column=0, padx=10, pady=10)
        remark_entry = ttk.Entry(win, width=20, font=('微軟正黑體', 12))
        remark_entry.insert(0, data['unique_mark'] or "")
        remark_entry.grid(row=2, column=1)

        def save():
            with self.db_lock:
                Database.execute('UPDATE Customers SET name = ?, contact_method = ?, unique_mark = ? WHERE customer_id = ?', 
                                (name_entry.get(), contact_var.get(), remark_entry.get(), customer_id))
            self.load_client_data()
            win.destroy()
            self.status_var.set(f"客戶 ID: {customer_id} 已更新")
            logging.info(f"編輯客戶 ID: {customer_id}, 新名稱 {name_entry.get()}")

        ttk.Button(win, text="保存", command=save).grid(row=3, column=0, columnspan=2, pady=20)

    def edit_treatment(self):
        """編輯療程記錄"""
        selected = self.customer_tree.selection()
        if len(selected) != 1:
            messagebox.showwarning("警告", "請選擇一個客戶")
            return
        customer_id = self.customer_tree.item(selected[0], 'values')[0]
        with self.db_lock:
            treatments = Database.execute('''
                SELECT ct.customer_treatment_id, ct.treatment_date, t.name, ct.is_peak, ct.neck_treatment, ct.price, 
                       ct.remaining_sessions, ct.remaining_retouch_count, t.can_add_neck, t.has_remaining_sessions
                FROM Customer_Treatments ct JOIN Treatments t ON ct.treatment_id = t.treatment_id 
                WHERE ct.customer_id = ?
            ''', (customer_id,), fetch=True)
        
        win = Toplevel(self.master)
        win.title(f"編輯客戶 {customer_id} 療程")
        win.geometry("800x500")
        tree = ttk.Treeview(win, columns=("ID", "日期", "項目", "時段", "頸部", "金額", "剩餘次數", "剩餘補脫"), show="headings")
        for col, width in [("ID", 80), ("日期", 150), ("項目", 200), ("時段", 120), ("頸部", 80), ("金額", 150), ("剩餘次數", 120), ("剩餘補脫", 80)]:
            tree.heading(col, text=col)
            tree.column(col, width=width, anchor='center')
        tree.pack(fill=BOTH, expand=True, padx=10, pady=10)
        for t in treatments:
            period = "Peak" if t['is_peak'] and t['name'] == "Einxel Plus膠原修復針" and t['peak_price'] != t['non_peak_price'] else ""
            neck_status = "是" if t['neck_treatment'] and t['can_add_neck'] else ""
            remaining = t['remaining_sessions'] if t['has_remaining_sessions'] else ""
            retouch_count = f"剩餘補脫：{t['remaining_retouch_count']}" if t['name'].startswith("熱能氣化") else ""
            tree.insert("", "end", values=(t['customer_treatment_id'], t['treatment_date'], t['name'], period, neck_status, 
                                          f"${round(t['price'])}", remaining, retouch_count))

        def edit_selected():
            sel = tree.selection()
            if len(sel) != 1:
                messagebox.showwarning("警告", "請選擇一筆療程")
                return
            tid = tree.item(sel[0], 'values')[0]
            with self.db_lock:
                data = Database.execute('''
                    SELECT treatment_date, treatment_id, is_peak, neck_treatment, price, remaining_sessions, remaining_retouch_count
                    FROM Customer_Treatments WHERE customer_treatment_id = ?
                ''', (tid,), fetch=True)[0]
                treatment = Database.execute('SELECT name, can_add_neck, has_remaining_sessions FROM Treatments WHERE treatment_id = ?', 
                                            (data['treatment_id'],), fetch=True)[0]

            edit_win = Toplevel(win)
            edit_win.title(f"編輯療程 ID: {tid}")
            edit_win.geometry("400x300")
            ttk.Label(edit_win, text="日期：").grid(row=0, column=0, padx=10, pady=10)
            date_entry = DateEntry(edit_win, width=15, date_pattern="yyyy-mm-dd", font=('微軟正黑體', 12))
            date_entry.set_date(data['treatment_date'])
            date_entry.grid(row=0, column=1)
            ttk.Label(edit_win, text="價格：").grid(row=1, column=0, padx=10, pady=10)
            price_entry = ttk.Entry(edit_win, width=15, font=('微軟正黑體', 12))
            price_entry.insert(0, data['price'])
            price_entry.grid(row=1, column=1)
            neck_var = IntVar(value=data['neck_treatment'])
            if treatment['can_add_neck']:
                ttk.Label(edit_win, text="加做頸部：").grid(row=2, column=0, padx=10, pady=10)
                ttk.Checkbutton(edit_win, text="是", variable=neck_var).grid(row=2, column=1, sticky=W)
            if treatment['has_remaining_sessions']:
                ttk.Label(edit_win, text="剩餘次數：").grid(row=3, column=0, padx=10, pady=10)
                remaining_entry = ttk.Entry(edit_win, width=15, font=('微軟正黑體', 12))
                remaining_entry.insert(0, data['remaining_sessions'] or 0)
                remaining_entry.grid(row=3, column=1)

            def save():
                new_price = float(price_entry.get())
                neck = neck_var.get() if treatment['can_add_neck'] else 0
                remaining = int(remaining_entry.get()) if treatment['has_remaining_sessions'] else data['remaining_sessions']
                is_retouch = 0
                customer_id = Database.execute('SELECT customer_id FROM Customer_Treatments WHERE customer_treatment_id = ?', (tid,), fetch=True)[0]['customer_id']
                if treatment['name'] == "補脫療程" and new_price == 0:
                    parent_id = BusinessLogic.check_retouch_eligibility(customer_id, date_entry.get_date().strftime("%Y-%m-%d"))
                    if not parent_id:
                        messagebox.showerror("錯誤", "半年內無可用首次療程或已使用過免費補脫")
                        return
                    is_retouch = 1
                    retouch_parent_id = parent_id
                    new_price = 0
                else:
                    treatment_data = BusinessLogic.match_treatment(new_price)
                    if not treatment_data and treatment['name'] != "熱能氣化 - 單次脫墨/疣":
                        messagebox.showerror("錯誤", f"無匹配價格 {new_price} 的療程")
                        return
                    is_peak = 1 if new_price == treatment_data['peak_price'] else 0
                    retouch_parent_id = None

                with self.db_lock:
                    Database.execute('''
                        UPDATE Customer_Treatments SET treatment_date = ?, price = ?, neck_treatment = ?, remaining_sessions = ?, 
                        is_peak = ?, retouch_parent_id = ? WHERE customer_treatment_id = ?
                    ''', (date_entry.get_date().strftime("%Y-%m-%d"), new_price, neck, remaining, is_peak, retouch_parent_id, tid))
                self.load_client_data()
                edit_win.destroy()
                win.destroy()
                self.status_var.set(f"療程 ID: {tid} 已更新")
                logging.info(f"編輯療程 ID: {tid}, 新價格 ${round(new_price)}")

            ttk.Button(edit_win, text="保存", command=save).grid(row=4 if treatment['can_add_neck'] or treatment['has_remaining_sessions'] else 2, column=0, columnspan=2, pady=20)

        ttk.Button(win, text="編輯選中記錄", command=edit_selected).pack(pady=10)

    def delete_customer(self):
        """刪除客戶"""
        selected = self.customer_tree.selection()
        if not selected:
            messagebox.showwarning("警告", "請選擇客戶")
            return
        if messagebox.askyesno("確認", "確定刪除所選客戶及其療程記錄？")):
            customer_ids = [self.customer_tree.item(item, 'values')[0] for item in selected]
            with self.db_lock:
                with sqlite3.connect('secret_time.db') as conn:
                    cursor = conn.cursor()
                    cursor.executemany('DELETE FROM Customer_Treatments WHERE customer_id = ?', [(cid,) for cid in customer_ids])
                    cursor.executemany('DELETE FROM Customers WHERE customer_id = ?', [(cid,) for cid in customer_ids])
                    conn.commit()
            self.load_client_data()
            self.status_var.set(f"已刪除 {len(customer_ids)} 個客戶")
            logging.info(f"刪除客戶 ID: {', '.join(customer_ids)}")

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
        expense_id = Database.execute('SELECT expense_id FROM Expenses WHERE expense_date = ? AND category = ? AND amount = ?', 
                                     (self.expense_tree.item(selected[0], 'values')[0], self.expense_tree.item(selected[0], 'values')[1], float(self.expense_tree.item(selected[0], 'values')[2].replace('$', ''))), fetch=True)[0]['expense_id']
        data = Database.execute('SELECT expense_date, category, amount, description FROM Expenses WHERE expense_id = ?', (expense_id,), fetch=True)[0]

        win = Toplevel(self.master)
        win.title(f"編輯支出 ID: {expense_id}")
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
                    Database.execute('UPDATE Expenses SET expense_date = ?, category = ?, amount = ?, description = ? WHERE expense_id = ?', 
                                    (date_entry.get_date().strftime("%Y-%m-%d"), category_combo.get(), amount, desc_entry.get(), expense_id))
                self.load_finance_data()
                win.destroy()
                self.status_var.set(f"支出 ID: {expense_id} 已更新")
                logging.info(f"編輯支出 ID: {expense_id}, 新金額 ${round(amount)}")
            except ValueError:
                messagebox.showerror("錯誤", "金額必須為數字")

        ttk.Button(win, text="保存", command=save).grid(row=4, column=0, columnspan=2, pady=20)

    def delete_expense(self):
        """刪除支出"""
        selected = self.expense_tree.selection()
        if not selected:
            messagebox.showwarning("警告", "請選擇支出")
            return
        if messagebox.askyesno("確認", "確定刪除所選支出？")):
            for item in selected:
                expense_id = Database.execute('SELECT expense_id FROM Expenses WHERE expense_date = ? AND category = ? AND amount = ?', 
                                             (self.expense_tree.item(item, 'values')[0], self.expense_tree.item(item, 'values')[1], float(self.expense_tree.item(item, 'values')[2].replace('$', ''))), fetch=True)[0]['expense_id']
                with self.db_lock:
                    Database.execute('DELETE FROM Expenses WHERE expense_id = ?', (expense_id,))
            self.load_finance_data()
            self.status_var.set(f"已刪除 {len(selected)} 筆支出")
            logging.info(f"刪除 {len(selected)} 筆支出")

    def show_history(self):
        """顯示客戶歷史記錄"""
        selected = self.customer_tree.selection()
        if len(selected) != 1:
            messagebox.showwarning("警告", "請選擇一個客戶")
            return
        customer_id = self.customer_tree.item(selected[0], 'values')[0]
        with self.db_lock:
            treatments = Database.execute('''
                SELECT ct.treatment_date, t.name as treatment_name, ct.is_peak, ct.neck_treatment,
                       ct.price as price, ct.retouch_parent_id, ct.remaining_retouch_count
                FROM Customer_Treatments ct JOIN Treatments t ON ct.treatment_id = t.treatment_id
                WHERE ct.customer_id = ?
                ORDER BY ct.treatment_date DESC
            ''', (customer_id,), fetch=True)

        win = Toplevel(self.master)
        win.title(f"客戶 {customer_id} 歷史記錄")
        win.geometry("800x400")
        tree = ttk.Treeview(win, columns=("日期", "項目", "時段", "頸部", "金額", "剩餘補脫"), show="headings")
        for col, width in [("日期", 150), ("項目", 200), ("時段", 100), ("頸部", 80), ("金額", 150), ("剩餘補脫", 80)]:
            tree.heading(col, text=col)
            tree.column(col, width=width, anchor='center')
        tree.pack(fill=BOTH, expand=True, padx=10, pady=10)
        for t in treatments:
            period = "Peak" if t['is_peak'] and t['treatment_name'] == "Einxel Plus膠原修復針" and t['peak_price'] != t['non_peak_price'] else ""
            neck_status = "是" if t['neck_treatment'] and t['treatment_name'] == "Einxel Plus膠原修復針" else ""
            retouch_count = f"剩餘補脫：{t['remaining_retouch_count']}" if t['treatment_name'].startswith("熱能氣化") else ""
            tree.insert("", "end", values=(t['treatment_date'], t['treatment_name'], period, neck_status, f"${round(t['price'])}", retouch_count))

    def export_clients(self):
        """匯出客戶資料"""
        def export_task():
            try:
                with self.db_lock:
                    customers = Database.execute('''
                        SELECT c.customer_id, c.name, c.contact_method, COUNT(ct.customer_treatment_id) as count,
                               SUM(ct.price) as total, MAX(ct.remaining_sessions) as remaining,
                               MAX(ct.treatment_date) as last_date
                        FROM Customers c LEFT JOIN Customer_Treatments ct ON c.customer_id = ct.customer_id
                        GROUP BY c.customer_id, c.name, c.contact_method
                        ORDER BY last_date DESC
                    ''', fetch=True)
                df = pd.DataFrame([(c['customer_id'], c['name'], c['contact_method'], c['count'] or 0, c['total'] or 0, c['remaining'] or 'N/A', c['last_date'] or 'N/A') 
                                 for c in customers], columns=['ID', '姓名', '聯絡方式', '消費次數', '總金額', '剩餘次數', '上次消費'])
                df.to_excel('customers_export.xlsx', index=False)
                self.status_var.set("客戶資料已匯出至 customers_export.xlsx")
            except Exception as e:
                messagebox.showerror("錯誤", f"匯出失敗：{str(e)}")
                logging.error(f"客戶匯出失敗：{str(e)}")
        threading.Thread(target=export_task, daemon=True).start()

    def export_expenses(self):
        """匯出支出數據"""
        def export_task():
            try:
                with self.db_lock:
                    expenses = Database.execute('SELECT expense_date, category, amount, description FROM Expenses ORDER BY expense_date DESC', fetch=True)
                df = pd.DataFrame([(e['expense_date'], e['category'], e['amount'], e['description'] or "") 
                                 for e in expenses], columns=['日期', '類別', '金額', '備註'])
                df.to_excel('expenses_export.xlsx', index=False)
                self.status_var.set("支出資料已匯出至 expenses_export.xlsx")
            except Exception as e:
                messagebox.showerror("錯誤", f"匯出失敗：{str(e)}")
                logging.error(f"支出匯出失敗：{str(e)}")
        threading.Thread(target=export_task, daemon=True).start()

    def import_from_clipboard(self):
        """從剪貼簿匯入客戶資料，支援數據驗證和歷史記錄"""
        win = Toplevel(self.master)
        win.title("從剪貼簿匯入客戶資料")
        win.geometry("700x500")
        ttk.Label(win, text="請貼上表格內容（格式：日期 姓名 聯絡方式 價格，例如 2025-01-02 Layla WhatsApp 9926）：").pack(pady=10)
        text_area = Text(win, height=15, width=70, font=('微軟正黑體', 12))
        text_area.pack(padx=10, pady=10)

        def validate_and_import():
            content = text_area.get("1.0", END).strip()
            if not content:
                messagebox.showwarning("警告", "請貼上內容")
                return
            lines = [line.strip() for line in content.split('\n') if line.strip()]
            errors = []
            validated_data = []
            last_date = None
            for i, line in enumerate(lines, 1):
                values = [v for v in re.split(r'\s+|,', line) if v]
                if len(values) < 3:
                    errors.append(f"第 {i} 行格式錯誤: {line}")
                    continue
                date_str, name, contact, price = (values[0] if len(values) == 4 else None, 
                                                 values[0] if len(values) == 3 else values[1], 
                                                 values[1] if len(values) == 3 else values[2], 
                                                 values[2] if len(values) == 3 else values[3])
                try:
                    if date_str:
                        for fmt in ['%Y-%m-%d', '%m/%d/%Y', '%d-%m-%Y']:
                            try:
                                last_date = datetime.datetime.strptime(date_str, fmt)
                                break
                            except ValueError:
                                continue
                        if last_date is None:
                            raise ValueError(f"無法解析日期: {date_str}")
                    price = float(price)
                    validated_data.append((last_date.strftime('%Y-%m-%d') if last_date else None, name, contact, price))
                except ValueError as ve:
                    errors.append(f"第 {i} 行錯誤: {str(ve)} - {line}")
            
            if errors:
                error_win = Toplevel(win)
                error_win.title("數據驗證錯誤")
                error_win.geometry("600x400")
                error_text = Text(error_win, height=15, width=70)
                error_text.pack(padx=10, pady=10)
                for error in errors:
                    error_text.insert(END, error + "\n")
                ttk.Button(error_win, text="修正後重新驗證", command=lambda: [error_win.destroy(), validate_and_import()]).pack(pady=10)
                return
            
            with self.db_lock:
                with sqlite3.connect('secret_time.db') as conn:
                    cursor = conn.cursor()
                    cursor.execute('CREATE TEMP TABLE temp_import AS SELECT * FROM Customer_Treatments WHERE 1=0')
                    try:
                        import_id = Database.execute('INSERT INTO Import_History (import_date, record_count, status) VALUES (?, ?, ?)', 
                                                    (datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'), len(validated_data), 'Pending'))
                        for date, name, contact, price in validated_data:
                            customer_id = self.get_or_create_customer(name, contact)
                            treatment = BusinessLogic.match_treatment(price)
                            retouch_parent_id = None
                            if not treatment:
                                if price == 0:
                                    parent_id = BusinessLogic.check_retouch_eligibility(customer_id, date)
                                    if parent_id:
                                        treatment_data = Database.execute('SELECT * FROM Treatments WHERE treatment_id = ?', 
                                                                        (Database.execute('SELECT treatment_id FROM Customer_Treatments WHERE customer_treatment_id = ?', 
                                                                                         (parent_id,), fetch=True)[0]['treatment_id'],), fetch=True)
                                        if not treatment_data:
                                            raise ValueError("資料庫中未找到對應補脫療程")
                                        treatment = dict(treatment_data[0])
                                        treatment['neck_price'] = False
                                        is_peak = 0
                                        neck = 0
                                        package = 0
                                        final_price = 0
                                        retouch_parent_id = parent_id
                                    else:
                                        raise ValueError(f"客戶 {name} 無資格進行補脫")
                                else:
                                    raise ValueError(f"無匹配價格 {price} 的療程")
                            else:
                                treatment = dict(treatment)
                                is_peak = 1 if price == treatment['peak_price'] else 0
                                neck = 1 if treatment.get('neck_price', False) and treatment['name'] == 'Einxel Plus膠原修復針' else 0
                                package = 0
                                final_price = price + (neck * 300)
                            remaining, package_id = self.handle_package(treatment, customer_id, final_price, package)
                            retouch_count = 1 if treatment['name'].startswith("熱能氣化") and not retouch_parent_id else 0
                            cursor.execute('''
                                INSERT INTO temp_import (customer_id, treatment_id, treatment_date, is_peak, neck_treatment, package_id, remaining_sessions, price, retouch_parent_id, remaining_retouch_count, import_id)
                                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                            ''', (customer_id, treatment['treatment_id'], date, is_peak, neck, package_id, remaining, final_price, retouch_parent_id, retouch_count, import_id))
                        cursor.execute('INSERT INTO Customer_Treatments SELECT * FROM temp_import')
                        conn.commit()
                        Database.execute('UPDATE Import_History SET status = ? WHERE import_id = ?', ('Success', import_id))
                    except Exception as e:
                        conn.rollback()
                        Database.execute('UPDATE Import_History SET status = ? WHERE import_id = ?', ('Failed', import_id))
                        raise
            self.load_client_data()
            win.destroy()
            self.status_var.set("已成功匯入客戶資料")
            logging.info(f"成功匯入 {len(validated_data)} 筆數據，匯入 ID: {import_id}")

        ttk.Button(win, text="確認匯入", command=validate_and_import).pack(pady=10)

    def view_import_history(self):
        """查看匯入歷史並支援一鍵刪除"""
        win = Toplevel(self.master)
        win.title("匯入歷史記錄")
        win.geometry("600x400")
        
        history_frame = ttk.Frame(win)
        history_frame.pack(fill=BOTH, expand=True, padx=10, pady=10)
        history_tree = ttk.Treeview(history_frame, columns=("匯入ID", "日期", "筆數", "狀態"), show="headings", height=10)
        for col, width in [("匯入ID", 80), ("日期", 200), ("筆數", 100), ("狀態", 100)]:
            history_tree.heading(col, text=col)
            history_tree.column(col, width=width, anchor='center')
        history_tree.pack(side=LEFT, fill=BOTH, expand=True)

        scrollbar = ttk.Scrollbar(history_frame, orient=VERTICAL, command=history_tree.yview)
        history_tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=RIGHT, fill=Y)

        with self.db_lock:
            history_data = Database.execute('SELECT import_id, import_date, record_count, status FROM Import_History ORDER BY import_date DESC', fetch=True)
        for h in history_data:
            history_tree.insert("", "end", values=(h['import_id'], h['import_date'], h['record_count'], h['status']))

        def delete_import():
            selected = history_tree.selection()
            if not selected:
                messagebox.showwarning("警告", "請選擇一筆匯入記錄")
                return
            import_id = history_tree.item(selected[0], 'values')[0]
            if messagebox.askyesno("確認", f"確定刪除匯入 ID {import_id} 的所有相關數據？")):
                with self.db_lock:
                    Database.execute('DELETE FROM Customer_Treatments WHERE import_id = ?', (import_id,))
                    Database.execute('DELETE FROM Import_History WHERE import_id = ?', (import_id,))
                history_tree.delete(selected[0])
                self.load_client_data()
                self.status_var.set(f"已刪除匯入 ID {import_id} 的數據")
                logging.info(f"刪除匯入 ID: {import_id}")

        ttk.Button(win, text="刪除選中匯入", command=delete_import).pack(pady=10)

    def load_finance_data(self):
        """載入財務數據並計算淨收入"""
        self.expense_tree.delete(*self.expense_tree.get_children())
        current_month = datetime.datetime.now().strftime('%Y-%m')
        with self.db_lock:
            revenues = Database.execute('''
                SELECT strftime('%Y-%m', ct.treatment_date) as month, SUM(ct.price) as revenue
                FROM Customer_Treatments ct
                WHERE strftime('%Y-%m', ct.treatment_date) = ? AND ct.price > 0
                GROUP BY month
            ''', (current_month,), fetch=True)
            expenses = Database.execute('''
                SELECT strftime('%Y-%m', expense_date) as month, SUM(amount) as expense
                FROM Expenses
                WHERE strftime('%Y-%m', expense_date) = ?
                GROUP BY month
            ''', (current_month,), fetch=True)
        total_revenue = revenues[0]['revenue'] if revenues else 0
        total_expense = expenses[0]['expense'] if expenses else 0
        net_income = total_revenue - total_expense

        with self.db_lock:
            expenses_data = Database.execute('SELECT expense_date, category, amount, description FROM Expenses ORDER BY expense_date DESC', fetch=True)
        for e in expenses_data:
            self.expense_tree.insert("", "end", values=(e['expense_date'], e['category'], f"${round(e['amount'])}", e['description'] or "", f"${round(net_income)}"))

    def prev_page(self):
        """上一頁"""
        if self.current_page > 1:
            self.current_page -= 1
            self.load_client_data()

    def next_page(self):
        """下一頁"""
        self.current_page += 1
        self.load_client_data()

    def update_stats(self):
        """更新統計數據，包括收入趨勢、療程分布、客戶保留率和季度趨勢"""
        range_type = self.stats_range_var.get()
        time_period = self.stats_month_var.get()
        if not time_period:
            self.status_var.set("請選擇時間")
            return

        self.top_customers_tree.delete(*self.top_customers_tree.get_children())
        query = '''
            SELECT c.name, SUM(ct.price) as total
            FROM Customers c LEFT JOIN Customer_Treatments ct ON c.customer_id = ct.customer_id
            WHERE strftime('%Y-%m', ct.treatment_date) = ? AND ct.price > 0
            GROUP BY c.customer_id, c.name ORDER BY total DESC LIMIT 10
        ''' if range_type == "按月" else '''
            SELECT c.name, SUM(ct.price) as total
            FROM Customers c LEFT JOIN Customer_Treatments ct ON c.customer_id = ct.customer_id
            WHERE strftime('%Y', ct.treatment_date) = ? AND ct.price > 0
            GROUP BY c.customer_id, c.name ORDER BY total DESC LIMIT 10
        '''
        with self.db_lock:
            top = Database.execute(query, (time_period,), fetch=True)
        for i, c in enumerate(top, 1):
            self.top_customers_tree.insert("", "end", values=(i, c['name'], f"${round(c['total'] or 0)}"))

        # 清理舊圖表
        for widget in self.trend_frame.winfo_children():
            widget.destroy()
        for widget in self.pie_frame.winfo_children():
            widget.destroy()
        for widget in self.retention_frame.winfo_children():
            widget.destroy()

        # 收入與支出趨勢
        with self.db_lock:
            revenues = Database.execute('''
                SELECT strftime('%Y-%m-%d', ct.treatment_date) as day, SUM(ct.price) as revenue
                FROM Customer_Treatments ct
                WHERE strftime('%Y-%m', ct.treatment_date) = ? AND ct.price > 0
                GROUP BY day ORDER BY day
            ''' if range_type == "按月" else '''
                SELECT strftime('%Y-%m', ct.treatment_date) as month, SUM(ct.price) as revenue
                FROM Customer_Treatments ct
                WHERE strftime('%Y', ct.treatment_date) = ? AND ct.price > 0
                GROUP BY month ORDER BY month
            ''', (time_period,), fetch=True)
            expenses = Database.execute('''
                SELECT strftime('%Y-%m-%d', expense_date) as day, SUM(amount) as expense
                FROM Expenses
                WHERE strftime('%Y-%m', expense_date) = ? 
                GROUP BY day ORDER BY day
            ''' if range_type == "按月" else '''
                SELECT strftime('%Y-%m', expense_date) as month, SUM(amount) as expense
                FROM Expenses
                WHERE strftime('%Y', expense_date) = ? 
                GROUP BY month ORDER BY month
            ''', (time_period,), fetch=True)
        revenue_days = [datetime.datetime.strptime(r['day'], '%Y-%m-%d') if range_type == "按月" else datetime.datetime.strptime(r['month'], '%Y-%m') for r in revenues]
        revenue_values = [r['revenue'] or 0 for r in revenues]
        expense_days = [datetime.datetime.strptime(e['day'], '%Y-%m-%d') if range_type == "按月" else datetime.datetime.strptime(e['month'], '%Y-%m') for e in expenses]
        expense_values = [e['expense'] or 0 for e in expenses]
        fig = Figure(figsize=(10, 5))
        ax = fig.add_subplot(111)
        ax.plot(revenue_days, revenue_values, marker='o', label='收入', color='#A7B6C2')
        ax.plot(expense_days, expense_values, marker='x', label='支出', color='#D3CEDF')
        ax.set_title(f"{time_period} 收入與支出趨勢", fontsize=14)
        ax.set_xlabel("時間", fontsize=12)
        ax.set_ylabel("金額 (元)", fontsize=12)
        ax.legend()
        ax.grid(True)
        for tick in ax.get_xticklabels():
            tick.set_rotation(45)
        canvas = FigureCanvasTkAgg(fig, master=self.trend_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=BOTH, expand=True)

        # 療程類型分布圖
        with self.db_lock:
            treatment_counts = Database.execute('''
                SELECT t.name, COUNT(*) as count
                FROM Customer_Treatments ct JOIN Treatments t ON ct.treatment_id = t.treatment_id
                WHERE strftime('%Y-%m', ct.treatment_date) = ? AND ct.price > 0
                GROUP BY t.name
            ''' if range_type == "按月" else '''
                SELECT t.name, COUNT(*) as count
                FROM Customer_Treatments ct JOIN Treatments t ON ct.treatment_id = t.treatment_id
                WHERE strftime('%Y', ct.treatment_date) = ? AND ct.price > 0
                GROUP BY t.name
            ''', (time_period,), fetch=True)
        labels = [t['name'] for t in treatment_counts]
        sizes = [t['count'] for t in treatment_counts]
        fig_pie = Figure(figsize=(5, 5))
        ax_pie = fig_pie.add_subplot(111)
        ax_pie.pie(sizes, labels=labels, autopct='%1.1f%%', startangle=90)
        ax_pie.axis('equal')
        ax_pie.set_title(f"{time_period} 療程類型分布")
        canvas_pie = FigureCanvasTkAgg(fig_pie, master=self.pie_frame)
        canvas_pie.draw()
        canvas_pie.get_tk_widget().pack(fill=BOTH, expand=True)

        # 客戶保留率分析（按年）
        if range_type == "按年":
            with self.db_lock:
                total_customers = Database.execute('SELECT COUNT(DISTINCT customer_id) as total FROM Customer_Treatments WHERE strftime("%Y", treatment_date) = ?', (time_period,), fetch=True)[0]['total']
                repeat_customers = Database.execute('''
                    SELECT COUNT(DISTINCT customer_id) as repeat
                    FROM Customer_Treatments
                    WHERE strftime("%Y", treatment_date) = ? AND customer_id IN (
                        SELECT customer_id FROM Customer_Treatments GROUP BY customer_id HAVING COUNT(*) > 1
                    )
                ''', (time_period,), fetch=True)[0]['repeat']
            retention_rate = (repeat_customers / total_customers * 100) if total_customers > 0 else 0
            retention_label = ttk.Label(self.retention_frame, text=f"客戶保留率: {retention_rate:.1f}%", font=('微軟正黑體', 12))
            retention_label.pack(pady=10)

        # 季節性趨勢分析
        with self.db_lock:
            quarterly_data = Database.execute('''
                SELECT strftime('%Y-%m', treatment_date) as quarter, SUM(ct.price) as revenue, COUNT(DISTINCT ct.customer_id) as customer_count
                FROM Customer_Treatments ct
                WHERE strftime('%Y', ct.treatment_date) = ?
                GROUP BY strftime('%Y-%m', treatment_date)
                ORDER BY quarter
            ''', (time_period if range_type == "按年" else time_period[:4],), fetch=True)
        quarters = [f"Q{(int(q[5:7])-1)//3 + 1} {q[:4]}" for q in [d['quarter'] for d in quarterly_data]]
        revenues_q = [d['revenue'] or 0 for d in quarterly_data]
        customers_q = [d['customer_count'] or 0 for d in quarterly_data]
        fig_quarter = Figure(figsize=(10, 5))
        ax_quarter = fig_quarter.add_subplot(111)
        ax_quarter.plot(quarters, revenues_q, marker='o', label='收入', color='#A7B6C2')
        ax_quarter.plot(quarters, customers_q, marker='x', label='客戶數', color='#D3CEDF')
        ax_quarter.set_title(f"{time_period} 季節性趨勢")
        ax_quarter.set_xlabel("季度")
        ax_quarter.set_ylabel("數值")
        ax_quarter.legend()
        ax_quarter.grid(True)
        for tick in ax_quarter.get_xticklabels():
            tick.set_rotation(45)
        canvas_quarter = FigureCanvasTkAgg(fig_quarter, master=self.trend_frame)
        canvas_quarter.draw()
        canvas_quarter.get_tk_widget().pack(fill=BOTH, expand=True)

        self.status_var.set(f"已更新 {time_period} 的統計數據")

    def apply_treatment_filter(self):
        """篩選療程記錄"""
        start_date = self.treatment_start_date.get_date().strftime("%Y-%m-%d")
        end_date = self.treatment_end_date.get_date().strftime("%Y-%m-%d")
        self.treatment_tree.delete(*self.treatment_tree.get_children())
        with self.db_lock:
            treatments = Database.execute('''
                SELECT ct.treatment_date, c.name, t.name as treatment_name, ct.is_peak, ct.neck_treatment, ct.remaining_sessions,
                       ct.price as price, ct.remaining_retouch_count, t.can_add_neck, t.has_remaining_sessions
                FROM Customer_Treatments ct JOIN Customers c ON ct.customer_id = c.customer_id
                JOIN Treatments t ON ct.treatment_id = t.treatment_id
                WHERE ct.treatment_date BETWEEN ? AND ?
                ORDER BY ct.treatment_date DESC
            ''', (start_date, end_date), fetch=True)
        for t in treatments:
            period = "Peak" if t['is_peak'] and t['treatment_name'] == "Einxel Plus膠原修復針" and t['peak_price'] != t['non_peak_price'] else ""
            neck_status = "是" if t['neck_treatment'] and t['can_add_neck'] else ""
            remaining = t['remaining_sessions'] if t['has_remaining_sessions'] else ""
            retouch_count = f"剩餘補脫：{t['remaining_retouch_count']}" if t['treatment_name'].startswith("熱能氣化") else ""
            self.treatment_tree.insert("", "end", values=(
                t['treatment_date'], t['name'], t['treatment_name'], period, neck_status, f"${round(t['price'])}", 
                remaining, retouch_count
            ))
        self.status_var.set(f"篩選結果：顯示 {start_date} 至 {end_date} 的療程記錄")

    def sort_treeview(self, col):
        """排序客戶列表"""
        items = [(self.customer_tree.set(item, col), item) for item in self.customer_tree.get_children('')]
        items.sort(key=lambda x: float(x[0].replace('$', '')) if col == "總金額" else x[0] if col == "上次消費" else int(x[0]) if x[0].isdigit() else x[0], reverse=True)
        for index, (val, item) in enumerate(items):
            self.customer_tree.move(item, '', index)

    def clear_input(self):
        """清除輸入欄位"""
        self.name_entry.delete(0, END)
        self.price_entry.delete(0, END)
        self.neck_var.set(0)
        self.package_var.set(0)
        self.retouch_var.set(0)

    def apply_filter(self):
        """應用客戶篩選"""
        try:
            min_treatments = int(self.filter_treatments.get() or 0)
            min_amount = float(self.filter_amount.get() or 0)
            self.customer_tree.delete(*self.customer_tree.get_children())
            with self.db_lock:
                customers = Database.execute('''
                    SELECT c.customer_id, c.name, c.contact_method, COUNT(ct.customer_treatment_id) as count,
                           SUM(ct.price) as total, MAX(ct.remaining_sessions) as remaining,
                           MAX(ct.treatment_date) as last_date
                    FROM Customers c LEFT JOIN Customer_Treatments ct ON c.customer_id = ct.customer_id
                    GROUP BY c.customer_id, c.name, c.contact_method
                    HAVING count >= ? AND total >= ?
                    ORDER BY last_date DESC
                ''', (min_treatments, min_amount), fetch=True)
            for c in customers:
                self.customer_tree.insert("", "end", values=(c['customer_id'], c['name'], c['contact_method'], c['count'] or 0, 
                                                             f"${round(c['total'] or 0)}", c['remaining'] or "N/A", c['last_date'] or "N/A"))
            self.status_var.set(f"篩選結果：消費次數 ≥ {min_treatments}, 總金額 ≥ ${round(min_amount)}")
        except ValueError:
            messagebox.showwarning("警告", "請輸入有效數字")

    def search_clients(self):
        """搜尋客戶"""
        query = self.search_entry.get().strip()
        if query:
            self.customer_tree.delete(*self.customer_tree.get_children())
            with self.db_lock:
                customers = Database.execute('''
                    SELECT c.customer_id, c.name, c.contact_method, COUNT(ct.customer_treatment_id) as count,
                           SUM(ct.price) as total, MAX(ct.remaining_sessions) as remaining,
                           MAX(ct.treatment_date) as last_date
                    FROM Customers c LEFT JOIN Customer_Treatments ct ON c.customer_id = ct.customer_id
                    WHERE c.name LIKE ?
                    GROUP BY c.customer_id, c.name, c.contact_method
                    ORDER BY last_date DESC
                ''', (f"%{query}%",), fetch=True)
            for c in customers:
                self.customer_tree.insert("", "end", values=(c['customer_id'], c['name'], c['contact_method'], c['count'] or 0, 
                                                             f"${round(c['total'] or 0)}", c['remaining'] or "N/A", c['last_date'] or "N/A"))
            self.status_var.set(f"搜尋結果：找到 {len(customers)} 個匹配客戶")
        else:
            self.load_client_data()

if __name__ == "__main__":
    root = Tk()
    app = Application(root)
    root.mainloop()