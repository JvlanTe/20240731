# ソースコード
import openpyxl
import os
from datetime import datetime

wb = openpyxl.Workbook()
current_date = datetime.now()
formatted_date = current_date.strftime("%Y年%m月")
output_folder = f"請求書_{formatted_date}"
# フォルダを作成
os.makedirs(output_folder, exist_ok=True)
