import sqlite3

def init_database():
    try:
        with sqlite3.connect('secret_time.db') as conn:
            cursor = conn.cursor()

            # 檢查並創建/更新 Customers 表
            cursor.execute('PRAGMA table_info(Customers)')
            if not cursor.fetchall():
                cursor.execute('''
                    CREATE TABLE Customers (
                        customer_id INTEGER PRIMARY KEY AUTOINCREMENT,
                        name TEXT NOT NULL,
                        contact_method TEXT NOT NULL
                    )
                ''')
                print("已創建 Customers 表")

            # 檢查並創建/更新 Treatments 表
            cursor.execute('PRAGMA table_info(Treatments)')
            columns = [row[1] for row in cursor.fetchall()]
            if not columns:
                cursor.execute('''
                    CREATE TABLE Treatments (
                        treatment_id INTEGER PRIMARY KEY AUTOINCREMENT,
                        name TEXT NOT NULL,
                        peak_price REAL NOT NULL,
                        non_peak_price REAL NOT NULL
                    )
                ''')
                print("已創建 Treatments 表")
            if 'can_add_neck' not in columns:
                cursor.execute('ALTER TABLE Treatments ADD COLUMN can_add_neck BOOLEAN DEFAULT 0')
                print("已添加 can_add_neck 欄位")
            if 'has_remaining_sessions' not in columns:
                cursor.execute('ALTER TABLE Treatments ADD COLUMN has_remaining_sessions BOOLEAN DEFAULT 0')
                print("已添加 has_remaining_sessions 欄位")

            # 檢查並創建/更新 Customer_Treatments 表
            cursor.execute('PRAGMA table_info(Customer_Treatments)')
            columns = [row[1] for row in cursor.fetchall()]
            if not columns:
                cursor.execute('''
                    CREATE TABLE Customer_Treatments (
                        customer_treatment_id INTEGER PRIMARY KEY AUTOINCREMENT,
                        customer_id INTEGER,
                        treatment_id INTEGER,
                        treatment_date TEXT NOT NULL,
                        is_peak INTEGER DEFAULT 0,
                        neck_treatment INTEGER DEFAULT 0,
                        package_id INTEGER,
                        remaining_sessions INTEGER,
                        price REAL NOT NULL,
                        retouch_parent_id INTEGER,
                        FOREIGN KEY (customer_id) REFERENCES Customers(customer_id),
                        FOREIGN KEY (treatment_id) REFERENCES Treatments(treatment_id)
                    )
                ''')
                print("已創建 Customer_Treatments 表")
            if 'remaining_retouch_count' not in columns:
                cursor.execute('ALTER TABLE Customer_Treatments ADD COLUMN remaining_retouch_count INTEGER DEFAULT 1')
                print("已添加 remaining_retouch_count 欄位")
            if 'import_id' not in columns:
                cursor.execute('ALTER TABLE Customer_Treatments ADD COLUMN import_id INTEGER DEFAULT NULL')
                print("已添加 import_id 欄位")

            # 檢查並創建/更新 Expenses 表
            cursor.execute('PRAGMA table_info(Expenses)')
            if not cursor.fetchall():
                cursor.execute('''
                    CREATE TABLE Expenses (
                        expense_id INTEGER PRIMARY KEY AUTOINCREMENT,
                        expense_date TEXT NOT NULL,
                        category TEXT NOT NULL,
                        amount REAL NOT NULL,
                        description TEXT
                    )
                ''')
                print("已創建 Expenses 表")

            # 創建 Import_History 表
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS Import_History (
                    import_id INTEGER PRIMARY KEY AUTOINCREMENT,
                    import_date TEXT,
                    record_count INTEGER,
                    status TEXT
                )
            ''')
            print("已創建或確認 Import_History 表")

            # 確保 Treatments 表有 '補脫療程'
            cursor.execute('SELECT COUNT(*) FROM Treatments WHERE name = ?', ('補脫療程',))
            if cursor.fetchone()[0] == 0:
                cursor.execute('INSERT INTO Treatments (name, peak_price, non_peak_price, can_add_neck, has_remaining_sessions) VALUES (?, ?, ?, ?, ?)', 
                               ('補脫療程', 0, 0, 0, 0))
                print("已添加 '補脫療程' 到 Treatments 表")

            # 遷移補脫數據
            cursor.execute('''
                UPDATE Customer_Treatments
                SET treatment_id = (SELECT treatment_id FROM Treatments WHERE name = '補脫療程'),
                    remaining_retouch_count = 0
                WHERE price = 0 AND treatment_id = (SELECT treatment_id FROM Treatments WHERE name = '熱能氣化 - 單次脫墨/疣')
            ''')
            print("已遷移補脫數據至 '補脫療程'")

            # 初始化 Treatments 表的標記欄位
            cursor.execute('UPDATE Treatments SET can_add_neck = 1 WHERE name = "Einxel Plus膠原修復針"')
            cursor.execute('UPDATE Treatments SET has_remaining_sessions = 1 WHERE name = "組合療程：Einxel Plus膠原修復針 6次包套"')
            print("已初始化 Treatments 表的標記欄位")

            conn.commit()
            print("資料庫更新完成！")
    except Exception as e:
        print(f"資料庫更新失敗：{str(e)}")

if __name__ == "__main__":
    init_database()