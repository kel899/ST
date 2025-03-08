import sqlite3
from tkinter import *
from tkinter import ttk, messagebox
from database import Database
from business_logic import BusinessLogic
from constants import logging
import pandas as pd
from matplotlib import font_manager
import matplotlib.pyplot as plt
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import threading

# 設置 Matplotlib 支援中文
font_manager.fontManager.addfont('C:/Windows/Fonts/msjh.ttc')  # 微軟正黑體
plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei']
plt.rcParams['axes.unicode_minus'] = False

class StatsUI:
    def __init__(self, notebook, app):
        self.app = app  # 接收主 Application 實例
        self.master = app.master
        self.status_var = app.status_var
        self.tab = ttk.Frame(notebook)
        notebook.add(self.tab, text="統計分析")
        self.db_lock = threading.Lock()
        self.setup_stats_tab()

    def setup_stats_tab(self):
        """設置統計分析標籤頁"""
        stats_container = ttk.Frame(self.tab)
        stats_container.pack(expand=True, fill=BOTH, padx=10, pady=10)

        filter_frame = ttk.Frame(stats_container)
        filter_frame.pack(fill=X, pady=5)
        ttk.Label(filter_frame, text="選擇月份：").pack(side=LEFT, padx=10)
        self.month_combo = ttk.Combobox(filter_frame, width=15, font=('微軟正黑體', 12))
        self.month_combo.pack(side=LEFT, padx=10)
        self.month_combo['values'] = BusinessLogic.get_available_months()
        if self.month_combo['values']:
            self.month_combo.set(self.month_combo['values'][0])
        self.month_combo.bind("<<ComboboxSelected>>", lambda e: self.update_stats())

        trend_frame = ttk.LabelFrame(stats_container, text="收入與支出趨勢", padding=10)
        trend_frame.pack(fill=BOTH, expand=True, pady=10)
        self.trend_fig = Figure(figsize=(10, 3), dpi=100)
        self.trend_canvas = FigureCanvasTkAgg(self.trend_fig, master=trend_frame)
        self.trend_canvas.get_tk_widget().pack(fill=BOTH, expand=True)

        customer_frame = ttk.LabelFrame(stats_container, text="客戶排行", padding=10)
        customer_frame.pack(fill=X, pady=10)
        self.top_customers_tree = ttk.Treeview(customer_frame, columns=("排名", "姓名", "聯絡方式", "消費次數", "總金額"), show="headings", height=5)
        for col, width in [("排名", 80), ("姓名", 200), ("聯絡方式", 150), ("消費次數", 120), ("總金額", 150)]:
            self.top_customers_tree.heading(col, text=col)
            self.top_customers_tree.column(col, width=width, anchor='center')
        self.top_customers_tree.pack(fill=X, expand=True)

        treatment_frame = ttk.LabelFrame(stats_container, text="療程分布", padding=10)
        treatment_frame.pack(fill=BOTH, expand=True, pady=10)
        self.pie_fig = Figure(figsize=(6, 3), dpi=100)
        self.pie_canvas = FigureCanvasTkAgg(self.pie_fig, master=treatment_frame)
        self.pie_canvas.get_tk_widget().pack(fill=BOTH, expand=True)

        self.update_stats()

    def update_stats(self):
        """更新統計數據，異步實現"""
        def load_in_background():
            month = self.month_combo.get()
            with self.db_lock:
                income_data = Database.execute('''
                    SELECT strftime('%Y-%m', treatment_date) as month, SUM(price) as total_income
                    FROM Customer_Treatments
                    GROUP BY month
                    ORDER BY month
                ''', fetch=True)
                expense_data = Database.execute('''
                    SELECT strftime('%Y-%m', expense_date) as month, SUM(amount) as total_expense
                    FROM Expenses
                    GROUP BY month
                    ORDER BY month
                ''', fetch=True)
                customer_data = Database.execute('''
                    SELECT c.name, c.contact_method, COUNT(ct.customer_treatment_id) as count, SUM(ct.price) as total
                    FROM Customers c JOIN Customer_Treatments ct ON c.customer_id = ct.customer_id
                    WHERE strftime('%Y-%m', ct.treatment_date) = ?
                    GROUP BY c.customer_id, c.name, c.contact_method
                    ORDER BY total DESC LIMIT 10
                ''', (month,), fetch=True)
                treatment_data = Database.execute('''
                    SELECT t.name, COUNT(ct.customer_treatment_id) as count
                    FROM Treatments t JOIN Customer_Treatments ct ON t.treatment_id = ct.treatment_id
                    WHERE strftime('%Y-%m', ct.treatment_date) = ?
                    GROUP BY t.treatment_id, t.name
                    ORDER BY count DESC
                ''', (month,), fetch=True)

            income_df = pd.DataFrame(income_data)
            expense_df = pd.DataFrame(expense_data)
            self.master.after(0, lambda: self.update_charts(income_df, expense_df, customer_data, treatment_data))

        threading.Thread(target=load_in_background, daemon=True).start()

    def update_charts(self, income_df, expense_df, customer_data, treatment_data):
        """更新圖表和排行榜"""
        # 趨勢圖
        self.trend_fig.clear()
        ax = self.trend_fig.add_subplot(111)
        if not income_df.empty:
            ax.plot(income_df['month'], income_df['total_income'], label='收入', marker='o')
        if not expense_df.empty:
            ax.plot(expense_df['month'], expense_df['total_expense'], label='支出', marker='o')
        ax.set_title("收入與支出趨勢")
        ax.set_xlabel("月份")
        ax.set_ylabel("金額 ($)")
        ax.legend()
        ax.tick_params(axis='x', rotation=45)
        self.trend_fig.tight_layout()
        self.trend_canvas.draw()
import sqlite3
from tkinter import *
from tkinter import ttk, messagebox
from database import Database
from business_logic import BusinessLogic
from constants import logging
import pandas as pd
from matplotlib import font_manager
import matplotlib.pyplot as plt
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import threading

# 設置 Matplotlib 支援中文
font_manager.fontManager.addfont('C:/Windows/Fonts/msjh.ttc')  # 微軟正黑體
plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei']
plt.rcParams['axes.unicode_minus'] = False

class StatsUI:
    def __init__(self, notebook, app):
        self.app = app  # 接收主 Application 實例
        self.master = app.master
        self.status_var = app.status_var
        self.tab = ttk.Frame(notebook)
        notebook.add(self.tab, text="統計分析")
        self.db_lock = threading.Lock()
        self.setup_stats_tab()

    def setup_stats_tab(self):
        """設置統計分析標籤頁"""
        stats_container = ttk.Frame(self.tab)
        stats_container.pack(expand=True, fill=BOTH, padx=10, pady=10)

        filter_frame = ttk.Frame(stats_container)
        filter_frame.pack(fill=X, pady=5)
        ttk.Label(filter_frame, text="選擇月份：").pack(side=LEFT, padx=10)
        self.month_combo = ttk.Combobox(filter_frame, width=15, font=('微軟正黑體', 12))
        self.month_combo.pack(side=LEFT, padx=10)
        self.month_combo['values'] = BusinessLogic.get_available_months()
        if self.month_combo['values']:
            self.month_combo.set(self.month_combo['values'][0])
        self.month_combo.bind("<<ComboboxSelected>>", lambda e: self.update_stats())

        trend_frame = ttk.LabelFrame(stats_container, text="收入與支出趨勢", padding=10)
        trend_frame.pack(fill=BOTH, expand=True, pady=10)
        self.trend_fig = Figure(figsize=(10, 3), dpi=100)
        self.trend_canvas = FigureCanvasTkAgg(self.trend_fig, master=trend_frame)
        self.trend_canvas.get_tk_widget().pack(fill=BOTH, expand=True)

        customer_frame = ttk.LabelFrame(stats_container, text="客戶排行", padding=10)
        customer_frame.pack(fill=X, pady=10)
        self.top_customers_tree = ttk.Treeview(customer_frame, columns=("排名", "姓名", "聯絡方式", "消費次數", "總金額"), show="headings", height=5)
        for col, width in [("排名", 80), ("姓名", 200), ("聯絡方式", 150), ("消費次數", 120), ("總金額", 150)]:
            self.top_customers_tree.heading(col, text=col)
            self.top_customers_tree.column(col, width=width, anchor='center')
        self.top_customers_tree.pack(fill=X, expand=True)

        treatment_frame = ttk.LabelFrame(stats_container, text="療程分布", padding=10)
        treatment_frame.pack(fill=BOTH, expand=True, pady=10)
        self.pie_fig = Figure(figsize=(6, 3), dpi=100)
        self.pie_canvas = FigureCanvasTkAgg(self.pie_fig, master=treatment_frame)
        self.pie_canvas.get_tk_widget().pack(fill=BOTH, expand=True)

        self.update_stats()

    def update_stats(self):
        """更新統計數據，異步實現"""
        def load_in_background():
            month = self.month_combo.get()
            with self.db_lock:
                income_data = Database.execute('''
                    SELECT strftime('%Y-%m', treatment_date) as month, SUM(price) as total_income
                    FROM Customer_Treatments
                    GROUP BY month
                    ORDER BY month
                ''', fetch=True)
                expense_data = Database.execute('''
                    SELECT strftime('%Y-%m', expense_date) as month, SUM(amount) as total_expense
                    FROM Expenses
                    GROUP BY month
                    ORDER BY month
                ''', fetch=True)
                customer_data = Database.execute('''
                    SELECT c.name, c.contact_method, COUNT(ct.customer_treatment_id) as count, SUM(ct.price) as total
                    FROM Customers c JOIN Customer_Treatments ct ON c.customer_id = ct.customer_id
                    WHERE strftime('%Y-%m', ct.treatment_date) = ?
                    GROUP BY c.customer_id, c.name, c.contact_method
                    ORDER BY total DESC LIMIT 10
                ''', (month,), fetch=True)
                treatment_data = Database.execute('''
                    SELECT t.name, COUNT(ct.customer_treatment_id) as count
                    FROM Treatments t JOIN Customer_Treatments ct ON t.treatment_id = ct.treatment_id
                    WHERE strftime('%Y-%m', ct.treatment_date) = ?
                    GROUP BY t.treatment_id, t.name
                    ORDER BY count DESC
                ''', (month,), fetch=True)

            income_df = pd.DataFrame(income_data)
            expense_df = pd.DataFrame(expense_data)
            self.master.after(0, lambda: self.update_charts(income_df, expense_df, customer_data, treatment_data))

        threading.Thread(target=load_in_background, daemon=True).start()

    def update_charts(self, income_df, expense_df, customer_data, treatment_data):
        """更新圖表和排行榜"""
        # 趨勢圖
        self.trend_fig.clear()
        ax = self.trend_fig.add_subplot(111)
        if not income_df.empty:
            ax.plot(income_df['month'], income_df['total_income'], label='收入', marker='o')
        if not expense_df.empty:
            ax.plot(expense_df['month'], expense_df['total_expense'], label='支出', marker='o')
        ax.set_title("收入與支出趨勢")
        ax.set_xlabel("月份")
        ax.set_ylabel("金額 ($)")
        ax.legend()
        ax.tick_params(axis='x', rotation=45)
        self.trend_fig.tight_layout()
        self.trend_canvas.draw()

        # 客戶排行
        self.top_customers_tree.delete(*self.top_customers_tree.get_children())
        if customer_data:
            for i, c in enumerate(customer_data, 1):
                self.top_customers_tree.insert("", "end", values=(i, c['name'], c['contact_method'], c['count'], f"${round(c['total'])}"))
        else:
            self.top_customers_tree.insert("", "end", values=("", "無數據", "", "", ""))

        # 療程分布（餅圖）
        self.pie_fig.clear()
        ax = self.pie_fig.add_subplot(111)
        if treatment_data:
            df = pd.DataFrame(treatment_data)
            ax.pie(df['count'], labels=df['name'], autopct='%1.1f%%', startangle=90)
            ax.axis('equal')
            ax.set_title(f"{self.month_combo.get()} 療程分布")
        else:
            ax.text(0.5, 0.5, '無數據', horizontalalignment='center', verticalalignment='center')
        self.pie_fig.tight_layout()
        self.pie_canvas.draw()

        self.status_var.set(f"統計數據已更新 ({self.month_combo.get()})")