import sqlite3
from constants import DB_NAME, DEFAULT_TREATMENTS, logging
import os

class Database:
    @staticmethod
    def initialize_database():
        try:
            # 檢查資料庫檔案是否存在
            db_exists = os.path.exists(DB_NAME)
            with sqlite3.connect(DB_NAME) as conn:
                conn.row_factory = sqlite3.Row
                cursor = conn.cursor()
                for script in [
                    '''CREATE TABLE IF NOT EXISTS Customers (
                        customer_id INTEGER PRIMARY KEY AUTOINCREMENT, 
                        name TEXT NOT NULL, 
                        contact_method TEXT, 
                        unique_mark TEXT DEFAULT ''
                    )''',
                    '''CREATE TABLE IF NOT EXISTS Treatments (
                        treatment_id INTEGER PRIMARY KEY AUTOINCREMENT, 
                        name TEXT NOT NULL UNIQUE, 
                        peak_price REAL NOT NULL, 
                        non_peak_price REAL NOT NULL, 
                        is_combo INTEGER NOT NULL DEFAULT 0 CHECK(is_combo IN (0, 1)),
                        can_add_neck BOOLEAN DEFAULT 0,
                        has_remaining_sessions BOOLEAN DEFAULT 0
                    )''',
                    '''CREATE TABLE IF NOT EXISTS Customer_Treatments (
                        customer_treatment_id INTEGER PRIMARY KEY AUTOINCREMENT, 
                        customer_id INTEGER, 
                        treatment_id INTEGER, 
                        treatment_date TEXT NOT NULL, 
                        is_peak INTEGER CHECK(is_peak IN (0, 1)), 
                        neck_treatment INTEGER DEFAULT 0 CHECK(neck_treatment IN (0, 1)), 
                        package_id INTEGER, 
                        remaining_sessions INTEGER DEFAULT 0, 
                        price REAL DEFAULT 0, 
                        retouch_parent_id INTEGER, 
                        remaining_retouch_count INTEGER DEFAULT 1,
                        import_id INTEGER DEFAULT NULL,
                        FOREIGN KEY(customer_id) REFERENCES Customers(customer_id), 
                        FOREIGN KEY(treatment_id) REFERENCES Treatments(treatment_id), 
                        FOREIGN KEY(retouch_parent_id) REFERENCES Customer_Treatments(customer_treatment_id)
                    )''',
                    '''CREATE TABLE IF NOT EXISTS Expenses (
                        expense_id INTEGER PRIMARY KEY AUTOINCREMENT, 
                        expense_date TEXT NOT NULL, 
                        category TEXT NOT NULL, 
                        amount REAL NOT NULL, 
                        description TEXT
                    )''',
                    '''CREATE TABLE IF NOT EXISTS Import_History (
                        import_id INTEGER PRIMARY KEY AUTOINCREMENT,
                        import_date TEXT,
                        record_count INTEGER,
                        status TEXT
                    )'''
                ]:
                    cursor.execute(script)
                cursor.execute('CREATE INDEX IF NOT EXISTS idx_customer_name ON Customers(name)')
                cursor.execute('CREATE INDEX IF NOT EXISTS idx_treatment_date ON Customer_Treatments(treatment_date)')
                cursor.executemany('INSERT OR IGNORE INTO Treatments (name, peak_price, non_peak_price, is_combo) VALUES (?, ?, ?, ?)', DEFAULT_TREATMENTS)
                cursor.execute('UPDATE Customers SET contact_method = ? WHERE contact_method = ?', ('WhatsApp', 'Whatsapp'))
                # 為舊數據添加 retouch_parent_id 和 remaining_retouch_count 欄位
                cursor.execute("PRAGMA table_info(Customer_Treatments)")
                columns = [col[1] for col in cursor.fetchall()]
                if 'retouch_parent_id' not in columns:
                    cursor.execute('ALTER TABLE Customer_Treatments ADD COLUMN retouch_parent_id INTEGER')
                if 'remaining_retouch_count' not in columns:
                    cursor.execute('ALTER TABLE Customer_Treatments ADD COLUMN remaining_retouch_count INTEGER DEFAULT 1')
                if 'import_id' not in columns:
                    cursor.execute('ALTER TABLE Customer_Treatments ADD COLUMN import_id INTEGER DEFAULT NULL')
                # 遷移補脫數據
                cursor.execute('''
                    UPDATE Customer_Treatments
                    SET treatment_id = (SELECT treatment_id FROM Treatments WHERE name = '補脫療程'),
                        remaining_retouch_count = 0
                    WHERE price = 0 AND treatment_id = (SELECT treatment_id FROM Treatments WHERE name = '熱能氣化 - 單次脫墨/疣')
                ''')
                # 初始化 Treatments 表的標記欄位
                cursor.execute('UPDATE Treatments SET can_add_neck = 1 WHERE name = "Einxel Plus膠原修復針"')
                cursor.execute('UPDATE Treatments SET has_remaining_sessions = 1 WHERE name = "組合療程：Einxel Plus膠原修復針 6次包套"')
                conn.commit()
                logging.info(f"資料庫{'初始化' if not db_exists else '更新'}成功")
        except sqlite3.Error as e:
            logging.error(f"資料庫初始化失敗：{str(e)}")
            raise

    @staticmethod
    def execute(query, params=(), fetch=False):
        try:
            with sqlite3.connect(DB_NAME) as conn:
                conn.row_factory = sqlite3.Row
                cursor = conn.cursor()
                cursor.execute(query, params)
                if fetch:
                    return [dict(row) for row in cursor.fetchall()]
                conn.commit()
                return cursor.lastrowid
        except sqlite3.Error as e:
            logging.error(f"資料庫操作失敗：{str(e)}")
            raise