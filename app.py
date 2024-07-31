# 現在時刻を取得するモジュール
from datetime import datetime

current_time = datetime.now()

print(current_time)

#フォーマットした時間
formatted_time = current_time.strftime("%Y年%m月%d日")

print(formatted_time)
