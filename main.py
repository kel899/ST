from tkinter import Tk, messagebox
from ui import Application
from database import Database
from constants import logging

def main():
    try:
        # 初始化資料庫
        Database.initialize_database()  # 修正方法名稱
        
        # 啟動應用程式
        root = Tk()
        app = Application(root)
        root.mainloop()
    except Exception as e:
        logging.error(f"應用程式啟動失敗: {str(e)}")
        messagebox.showerror("錯誤", f"應用程式啟動失敗：{str(e)}")
        raise

if __name__ == "__main__":
    main()