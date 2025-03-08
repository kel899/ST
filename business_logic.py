from database import Database
from datetime import datetime, timedelta

class BusinessLogic:
    @staticmethod
    def match_treatment(price):
        """
        根據價格匹配療程，返回所有符合條件的療程列表。
        考慮 Peak、非 Peak 價格，以及 Einxel Plus 加頸部 (+300) 的情況。
        """
        base_prices = Database.execute('''
            SELECT treatment_id, name, peak_price, non_peak_price, is_combo 
            FROM Treatments 
            WHERE peak_price = ? OR non_peak_price = ? OR peak_price + 300 = ? OR non_peak_price + 300 = ?
        ''', (price, price, price, price), fetch=True)
        if base_prices:
            treatments = []
            for treatment in base_prices:
                treatment_dict = dict(treatment)
                treatment_dict['neck_price'] = (price == treatment['peak_price'] + 300 or 
                                               price == treatment['non_peak_price'] + 300) and \
                                               treatment['name'] == 'Einxel Plus膠原修復針'
                treatments.append(treatment_dict)
            return treatments
        return None

    @staticmethod
    def check_duplicate(name, contact):
        """檢查是否有重複的客戶名稱和聯絡方式"""
        return Database.execute('SELECT customer_id, unique_mark FROM Customers WHERE name = ? AND contact_method = ?', 
                              (name, contact), fetch=True)

    @staticmethod
    def get_customer_names_matching(prefix):
        """獲取匹配前綴的客戶名稱列表"""
        return [row['name'] for row in Database.execute('SELECT DISTINCT name FROM Customers WHERE name LIKE ? ORDER BY name', 
                                                       (f"{prefix}%",), fetch=True)]

    @staticmethod
    def get_available_months():
        """獲取有記錄的月份列表"""
        return [row['month'] for row in Database.execute('''
            SELECT DISTINCT strftime('%Y-%m', treatment_date) as month FROM Customer_Treatments
            UNION
            SELECT DISTINCT strftime('%Y-%m', expense_date) as month FROM Expenses
            ORDER BY month DESC
        ''', fetch=True)]

    @staticmethod
    def get_remaining_sessions(customer_id, treatment_name):
        """檢查指定客戶的套裝療程剩餘次數"""
        if treatment_name != "組合療程：Einxel Plus膠原修復針 6次包套":
            return None
        treatments = Database.execute('''
            SELECT package_id, remaining_sessions 
            FROM Customer_Treatments 
            WHERE customer_id = ? AND treatment_id = (SELECT treatment_id FROM Treatments WHERE name = ?) 
            ORDER BY treatment_date DESC LIMIT 1
        ''', (customer_id, treatment_name), fetch=True)
        return treatments[0] if treatments else None

    @staticmethod
    def check_retouch_eligibility(customer_id, treatment_date):
        """檢查客戶是否符合補脫資格（半年內首次療程未使用補脫）"""
        treatment_date = datetime.strptime(treatment_date, '%Y-%m-%d')
        cutoff_date = (treatment_date - timedelta(days=180)).strftime('%Y-%m-%d')
        eligible_treatments = [
            '熱能氣化 - 10粒',
            '熱能氣化 - 20粒',
            '熱能氣化 - 40粒',
            '熱能氣化 - 任脫 3800',
            '熱能氣化 - 任脫 4800'
        ]
        initial_treatments = Database.execute('''
            SELECT ct.customer_treatment_id, ct.treatment_date 
            FROM Customer_Treatments ct
            JOIN Treatments t ON ct.treatment_id = t.treatment_id
            WHERE ct.customer_id = ? 
            AND t.name IN ({}) 
            AND ct.price > 0 
            AND ct.treatment_date >= ? AND ct.treatment_date < ?
        '''.format(','.join(['?'] * len(eligible_treatments))), 
        (customer_id, *eligible_treatments, cutoff_date, treatment_date.strftime('%Y-%m-%d')), fetch=True)
        
        for treatment in initial_treatments:
            retouch_exists = Database.execute('''
                SELECT COUNT(*) FROM Customer_Treatments 
                WHERE retouch_parent_id = ?
            ''', (treatment['customer_treatment_id'],), fetch=True)[0]['COUNT(*)']
            if retouch_exists == 0:
                return treatment['customer_treatment_id']
        return None