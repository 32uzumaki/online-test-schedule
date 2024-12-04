import pandas as pd
import numpy as np
from datetime import datetime

# ファイルパス
equipment_schedule_path = r"装置アドレス、オンラインテスト管理表.csv"
available_hours_path = r'労働可能時間.csv'
monthly_capacity_path = r'月ごとのテスト可能台数.csv'

# データ読み込み
equipment_schedule = pd.read_csv(equipment_schedule_path, sheet_name='管理表', header=0)
available_hours = pd.read_csv(available_hours_path)
monthly_capacity = pd.read_csv(monthly_capacity_path)

# 必要な列を選択
new_equipment_schedule = equipment_schedule[['エリア', '図面装置No', '設備', '号機', 'オンライン対応', 'オンライン備考', 'リリース予定日', '装置型式毎の初回テスト対象', 'オンラインテスト担当者', '受け入れテスト実施日']]

# 有効なエリアのみをフィルタリング
valid_areas = ['SubBE', 'EPI', 'WP表', 'WP裏', 'EDS']
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
process_priority = {'SubBE': 5, 'EPI': 4, 'WP表': 3, 'WP裏': 2, 'EDS': 1}
filtered_schedule['エリア優先順位'] = filtered_schedule['エリア'].map(process_priority)

# 特別優先順位と工程優先順位を統合して最終的な優先順位を設定
filtered_schedule['最終優先順位'] = filtered_schedule['先行オンライン優先度：'].fillna(filtered_schedule['エリア優先順位'])
filtered_schedule.sort_values(by=['最終優先順位', 'リリース予定日'], ascending=[True, True], inplace=True)

# 担当者の利用可能時間のデータフレームを担当者名をキーにした辞書に変換
available_hours_dict = available_hours.set_index('担当者').to_dict('index')

# 月ごとのテスト可能台数を辞書に変換
monthly_capacity_dict = monthly_capacity.set_index('月')['テスト可能台数'].to_dict()

# スケジュールの初期化
schedule = pd.DataFrame(index=pd.date_range('2024-10-01', '2025-10-31', freq='M').strftime('%Y-%m'), columns=process_priority.keys())

# 日付検証関数
def validate_date(value, default_date='2024-11-01'):
    try:
        return pd.to_datetime(value)
    except:
        return pd.NaT

# 'リリース予定日'と'受け入れテスト実施日'の補正
filtered_schedule['リリース予定日'] = filtered_schedule['リリース予定日'].apply(validate_date)
filtered_schedule['受け入れテスト実施日'] = filtered_schedule['受け入れテスト実施日'].apply(validate_date)

# '受け入れテスト実施日'が存在する設備を最優先で処理
fixed_schedule_entries = filtered_schedule[filtered_schedule['受け入れテスト実施日'].notna()]
remaining_schedule_entries = filtered_schedule[filtered_schedule['受け入れテスト実施日'].isna()]

# '受け入れテスト実施日'がある設備をスケジュールに反映
for index, row in fixed_schedule_entries.iterrows():
    device_name = row['設備']
    process = row['エリア']
    incharge = row['オンラインテスト担当者']
    drawing_no = row['図面装置No']
    prosess = row['エリア']
    start_date = pd.to_datetime(row['受け入れテスト実施日'])
    start_month = start_date.strftime('%Y-%m')
    test_hours_needed = 40  # テストに必要な時間

    # 担当者と月の利用可能時間を確認
    if incharge not in available_hours_dict:
        print(f"担当者 {incharge} の利用可能時間のデータがありません。")
        continue

    if start_month not in available_hours_dict[incharge]:
        print(f"担当者 {incharge} の利用可能時間に月 {start_month} がありません。")
        continue

    if available_hours_dict[incharge][start_month] >= test_hours_needed and monthly_capacity_dict.get(start_month, 0) > 0:
        entry = f'({incharge}) {prosess} No.{drawing_no} {device_name}'
        if pd.isna(schedule.at[start_month, process]):
            schedule.at[start_month, process] = entry
        else:
            schedule.at[start_month, process] += '\n' + entry
        available_hours_dict[incharge][start_month] -= test_hours_needed
        monthly_capacity_dict[start_month] -= 1
    else:
        print(f"担当者 {incharge} の月 {start_month} の利用可能時間または月のテスト可能台数が不足しています。")
        # 必要に応じてエラーを出すか、処理を続けるか判断

# 残りの設備を通常通りスケジュールに反映
for index, row in remaining_schedule_entries.iterrows():
    device_name = row['設備']
    process = row['エリア']
    incharge = row['オンラインテスト担当者']
    drawing_no = row['図面装置No']
    prosess = row['エリア']
    start_date = pd.to_datetime(row['リリース予定日'])
    start_month = start_date.strftime('%Y-%m')
    test_hours_needed = 40  # テストに必要な時間

    for month in pd.date_range(start_month, '2025-10-31', freq='M').strftime('%Y-%m'):
        if available_hours_dict[incharge][month] >= test_hours_needed and monthly_capacity_dict.get(month, 0) > 0:
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