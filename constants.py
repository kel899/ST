import logging

# 配置日誌
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s', filename='secret_time.log')

# 常量配置
DB_NAME = 'secret_time.db'  # 資料庫檔案名稱
DEFAULT_CATEGORIES = ['房租', '水電費', '物料採購', '設備維護', '行銷費用', '其他']  # 預設支出類別
DEFAULT_TREATMENTS = {
    "超聲波注水護理": {"peak_price": 580, "non_peak_price": 522, "is_combo": 0},
    "Exuviance果酸煥膚": {"peak_price": 650, "non_peak_price": 585, "is_combo": 0},
    "24K純金箔排毒煥膚": {"peak_price": 720, "non_peak_price": 648, "is_combo": 0},
    "等離子抗菌消炎細胞機": {"peak_price": 880, "non_peak_price": 792, "is_combo": 0},
    "PULXE電脈衝緊緻水光槍": {"peak_price": 780, "non_peak_price": 702, "is_combo": 0},
    "韓國幹細胞嬰兒針": {"peak_price": 980, "non_peak_price": 882, "is_combo": 0},
    "Einxel Plus膠原修復針": {"peak_price": 1880, "non_peak_price": 1680, "is_combo": 0},
    "組合療程：Exuviance果酸煥膚 + 等離子抗菌消炎細胞機": {"peak_price": 1071, "non_peak_price": 963, "is_combo": 1},
    "組合療程：韓國幹細胞嬰兒針 + 等離子抗菌消炎細胞機": {"peak_price": 1395, "non_peak_price": 1255, "is_combo": 1},
    "組合療程：Einxel Plus膠原修復針 6次包套": {"peak_price": 9926, "non_peak_price": 9926, "is_combo": 1},
    "熱能氣化 - 單次脫墨/疣": {"peak_price": 0, "non_peak_price": 0, "is_combo": 0},
    "熱能氣化 - 10粒": {"peak_price": 1600, "non_peak_price": 1600, "is_combo": 0},
    "熱能氣化 - 20粒": {"peak_price": 1800, "non_peak_price": 1800, "is_combo": 0},
    "熱能氣化 - 40粒": {"peak_price": 2800, "non_peak_price": 2800, "is_combo": 0},
    "熱能氣化 - 任脫 3800": {"peak_price": 3800, "non_peak_price": 3800, "is_combo": 0},
    "熱能氣化 - 任脫 4800": {"peak_price": 4800, "non_peak_price": 4800, "is_combo": 0}
}