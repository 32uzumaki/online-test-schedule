import pandas as pd
import numpy as np
from datetime import datetime

# ファイルパス
equipment_schedule_path = '装置搬入スケジュール改訂版.csv'
available_hours_path = '/労働可能時間.csv'
monthly_capacity_path = '月ごとのテスト可能台数.csv'

# データ読み込み
equipment_schedule = pd.read_csv(equipment_schedule_path)
available_hours = pd.read_csv(available_hours_path)
monthly_capacity = pd.read_csv(monthly_capacity_path)


# 新しい条件に基づいてデータをフィルタリング
filtered_schedule = equipment_schedule[(equipment_schedule['オンライン'] == '◯') & (equipment_schedule['区分'] == '初号機')]

# 以降のスケジュール生成ロジックは、filtered_scheduleを使用して実行

# 優先順位の数値化
def extract_priority(note):
    if pd.isna(note):
        return np.nan
    elif "特別優先" in note:
        return int(note.replace("特別優先", "").replace("位", ""))
    return np.nan

# 特別優先順位を抽出
filtered_schedule['特別優先順位'] = filtered_schedule['備考'].apply(extract_priority)

# 工程の優先順位をマッピング
process_priority = {'SubBE': 3, 'EPI': 2, 'WP表': 1}
filtered_schedule['工程優先順位'] = filtered_schedule['工程'].map(process_priority)

# 特別優先順位と工程優先順位を統合して最終的な優先順位を設定
filtered_schedule['最終優先順位'] = filtered_schedule['特別優先順位'].fillna(filtered_schedule['工程優先順位'])
filtered_schedule.sort_values(by=['最終優先順位', '搬入日'], ascending=[True, True], inplace=True)

# 担当者の利用可能時間のデータフレームを担当者名をキーにした辞書に変換
available_hours_dict = available_hours.set_index('担当者').to_dict('index')

# 月ごとのテスト可能台数を辞書に変換
monthly_capacity_dict = monthly_capacity.set_index('月')['テスト可能台数'].to_dict()

# スケジュールの初期化
schedule = pd.DataFrame(index=pd.date_range('2024-11-01', '2025-10-31', freq='M').strftime('%Y-%m'), columns=process_priority.keys())

# テスト割り当ての実装
for index, row in filtered_schedule.iterrows():
    device_name = row['機種名']
    process = row['工程']
    incharge = row['仕様決め担当者']
    start_date = pd.to_datetime(row['搬入日'])
    start_month = start_date.strftime('%Y-%m')
    test_hours_needed = 40  # テストに必要な時間
    
    for month in pd.date_range(start_month, '2025-10-31', freq='M').strftime('%Y-%m'):
        if available_hours_dict[incharge][month] >= test_hours_needed and monthly_capacity_dict[month] > 0:
            entry = f'({incharge}) {device_name}'
            if pd.isna(schedule.at[month, process]):
                schedule.at[month, process] = entry
            else:
                schedule.at[month, process] += '\n' + entry
            available_hours_dict[incharge][month] -= test_hours_needed
            monthly_capacity_dict[month] -= 1
            break

schedule.fillna('', inplace=True)
# スケジュールをCSVファイルに保存
transposed_schedule = schedule.T # 行と列を入れ替え
transposed_schedule.to_csv('test_schedule2.csv')
print("スケジュールがCSVファイルに保存されました。")