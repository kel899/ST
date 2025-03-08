from database import Database
from datetime import datetime, timedelta

class BusinessLogic:
    @staticmethod
    def match_treatment(price):
        base_prices = Database.execute('''
            SELECT treatment_id, name, peak_price, non_peak_price, is_combo 
            FROM Treatments 
            WHERE peak_price = ? OR non_peak_price = ? OR peak_price + 300 = ? OR non_peak_price + 300 = ?
        ''', (price, price, price, price), fetch=True)
        if base_prices:
            treatment = base_prices[0]
            treatment_dict = dict(treatment)
            treatment_dict['neck_price'] = (price == treatment['peak_price'] + 300 or price == treatment['non_peak_price'] + 300) and treatment['name'] == 'Einxel Plus膠原修復針'
            return treatment_dict
        return None

    @staticmethod
    def check_duplicate(name, contact):
        return Database.execute('SELECT customer_id, unique_mark FROM Customers WHERE name = ? AND contact_method = ?', (name, contact), fetch=True)

    @staticmethod
    def get_customer_names_matching(prefix):
        return [row['name'] for row in Database.execute('SELECT DISTINCT name FROM Customers WHERE name LIKE ? ORDER BY name', (f"{prefix}%",), fetch=True)]

    @staticmethod
    def get_available_months():
        return [row['month'] for row in Database.execute('''
            SELECT DISTINCT strftime(\'%Y-%m\', treatment_date) as month FROM Customer_Treatments
            UNION
            SELECT DISTINCT strftime(\'%Y-%m\', expense_date) as month FROM Expenses
            ORDER BY month DESC
        ''', fetch=True)]

    @staticmethod
    def get_remaining_sessions(customer_id, treatment_name):
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