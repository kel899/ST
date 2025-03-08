from tkinter import Tk, ttk, StringVar, messagebox
from client_ui import ClientUI
from finance_ui import FinanceUI
from stats_ui import StatsUI
from database import Database
from constants import logging

class Application:
    def __init__(self, master):
        self.master = master
        self.master.title("美容工作室管理系統")
        self.master.geometry("1400x800")
        self.master.configure(bg="#E8ECEF")
        self.status_var = StringVar(value="歡迎使用美容工作室管理系統")

        try:
            # 初始化資料庫
            Database.initialize_database()
            self.status_var.set("資料庫初始化完成")
        except Exception as e:
            logging.error(f"資料庫初始化失敗: {str(e)}")
            messagebox.showerror("錯誤", f"應用程式啟動失敗：{str(e)}")
            raise

        # 設置主介面
        self.notebook = ttk.Notebook(self.master)
        self.client_ui = ClientUI(self.notebook, self)
        self.finance_ui = FinanceUI(self.notebook, self)
        self.stats_ui = StatsUI(self.notebook, self)
        self.notebook.pack(expand=True, fill="both", padx=10, pady=10)

        # 狀態列
        status_bar = ttk.Label(self.master, textvariable=self.status_var, relief="sunken", anchor="w", font=('微軟正黑體', 12))
        status_bar.pack(side="bottom", fill="x", padx=10, pady=5)

        # 綁定快捷鍵
        self.master.bind("<Control-s>", lambda e: self.client_ui.add_record())
        self.master.bind("<Control-d>", lambda e: self.client_ui.delete_customer())

if __name__ == "__main__":
    root = Tk()
    app = Application(root)
    root.mainloop()