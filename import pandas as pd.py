import pandas as pd
import numpy as np
from datetime import datetime

# ファイルパス
equipment_schedule_path = r"装置アドレス、オンラインテスト管理表.xlsx"
available_hours_path = r'労働可能時間.xlsx'
monthly_capacity_path = r'月ごとのテスト可能台数.csv'

# データ読み込み
equipment_schedule = pd.read_excel(equipment_schedule_path, sheet_name='管理表', header=0)
available_hours = pd.read_excel(available_hours_path)
monthly_capacity = pd.read_csv(monthly_capacity_path)

new_equipment_schedule = equipment_schedule[['エリア', '図面装置No', '設備', '号機', 'オンライン対応', 'オンライン備考', 'リリース予定日', '装置型式毎の初回テスト対象', 'オンラインテスト担当者']]

# 有効なエリアのみをフィルタリング
valid_areas = ['SubBE', 'EPI', 'WP表', 'WP裏']
new_equipment_schedule = new_equipment_schedule[new_equipment_schedule['エリア'].isin(valid_areas)]

# 新しい条件に基づいてデータをフィルタリング
filtered_schedule = new_equipment_schedule[
    (new_equipment_schedule['オンライン対応'] == '〇') &
    (new_equipment_schedule['号機'] == 1) &
    (new_equipment_schedule['装置型式毎の初回テスト対象'] == '〇')
]

# 優先順位の数値化
def extract_priority(note):
    if pd.isna(note):
        return np.nan
    elif "先行オンライン優先度：" in note:
        return int(note.replace("先行オンライン優先度：", "").replace("位", ""))
    return np.nan

# 特別優先順位を抽出
filtered_schedule['先行オンライン優先度：'] = filtered_schedule['オンライン備考'].apply(extract_priority)

# 工程の優先順位をマッピング
process_priority = {'SubBE': 4, 'EPI': 3, 'WP表': 2, 'WP裏': 1}
filtered_schedule['エリア優先順位'] = filtered_schedule['エリア'].map(process_priority)

# 特別優先順位と工程優先順位を統合して最終的な優先順位を設定
filtered_schedule['最終優先順位'] = filtered_schedule['先行オンライン優先度：'].fillna(filtered_schedule['エリア優先順位'])
filtered_schedule.sort_values(by=['最終優先順位', 'リリース予定日'], ascending=[True, True], inplace=True)

# 担当者の利用可能時間のデータフレームを担当者名をキーにした辞書に変換
available_hours_dict = available_hours.set_index('担当者').to_dict('index')

# 月ごとのテスト可能台数を辞書に変換
monthly_capacity_dict = monthly_capacity.set_index('月')['テスト可能台数'].to_dict()

# スケジュールの初期化
schedule = pd.DataFrame(index=pd.date_range('2024-10-01', '2025-10-31', freq='ME').strftime('%Y-%m'), columns=process_priority.keys())

# 日付検証関数
def validate_date(value, default_date='2024-11-01'):
    try:
        return pd.to_datetime(value)
    except:
        return pd.to_datetime(default_date)

# 'リリース予定日'の補正
filtered_schedule['リリース予定日'] = filtered_schedule['リリース予定日'].apply(validate_date)

# テスト割り当ての実装
for index, row in filtered_schedule.iterrows():
    device_name = row['設備']
    process = row['エリア']
    incharge = row['オンラインテスト担当者']
    drawing_no = row['図面装置No']
    prosess = row['エリア']
    start_date = pd.to_datetime(row['リリース予定日'])
    start_month = start_date.strftime('%Y-%m')
    test_hours_needed = 40  # テストに必要な時間

    for month in pd.date_range(start_month, '2025-10-31', freq='ME').strftime('%Y-%m'):
        if available_hours_dict[incharge][month] >= test_hours_needed and monthly_capacity_dict[month] > 0:
            entry = f'({incharge}) {prosess} No.{drawing_no} {device_name}'
            if pd.isna(schedule.at[month, process]):
                schedule.at[month, process] = entry
            else:
                schedule.at[month, process] += '\n' + entry
            available_hours_dict[incharge][month] -= test_hours_needed
            monthly_capacity_dict[month] -= 1
            break

schedule.fillna('', inplace=True)
# スケジュールをCSVファイルに保存
transposed_schedule = schedule.T  # 行と列を入れ替え

transposed_schedule.to_csv('test_schedule5.csv')
print("スケジュールがCSVファイルに保存されました。")