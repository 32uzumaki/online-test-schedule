import pandas as pd
import numpy as np
from datetime import datetime

# ファイルパス
equipment_schedule_path = r"装置アドレス、オンラインテスト管理表.csv"

# データ読み込み
equipment_schedule = pd.read_csv(equipment_schedule_path)

# 必要な列を選択
new_equipment_schedule = equipment_schedule[['エリア', '図面装置No', '設備', '号機', 'オンライン対応', 'オンライン備考', 'リリース予定日', '初号機テスト実施時期', '装置型式毎の初回テスト対象', 'オンラインテスト担当者', '受け入れテスト実施日']]

# 有効なエリアのみをフィルタリング
valid_areas = ['SubBE', 'EPI', 'WP表', 'WP裏', 'EDS']
new_equipment_schedule = new_equipment_schedule[new_equipment_schedule['エリア'].isin(valid_areas)]

# 新しい条件に基づいてデータをフィルタリング
filtered_schedule = new_equipment_schedule[
    (new_equipment_schedule['オンライン対応'] == '〇') &
    (new_equipment_schedule['装置型式毎の初回テスト対象'] == '増設機')
]

# 日付検証関数
def validate_date(value):
    try:
        return pd.to_datetime(value)
    except:
        return pd.NaT

# 'リリース予定日'と'初号機テスト実施時期'の補正
filtered_schedule['リリース予定日'] = filtered_schedule['リリース予定日'].apply(validate_date)
filtered_schedule['初号機テスト実施時期'] = filtered_schedule['初号機テスト実施時期'].apply(validate_date)

# '初号機テスト実施時期'がNaTの場合、'リリース予定日'の月を代入
filtered_schedule['初号機テスト実施時期'] = filtered_schedule.apply(
    lambda row: row['リリース予定日'] if pd.isna(row['初号機テスト実施時期']) else row['初号機テスト実施時期'],
    axis=1
)

# 'リリース予定日'が'初号機テスト実施時期'より後になるようデータをフィルタリング
filtered_schedule = filtered_schedule[
    filtered_schedule['リリース予定日'] > filtered_schedule['初号機テスト実施時期']
]

# スケジュールの初期化
schedule = pd.DataFrame(columns=['年月', 'エリア', '設備', '図面装置No', 'オンラインテスト担当者'])

# 労働可能時間と月ごとのテスト可能台数を辞書に変換
available_hours_dict = available_hours.set_index('担当者').to_dict('index')
monthly_capacity_dict = monthly_capacity.set_index('年月')['テスト可能台数'].to_dict()

# 各エントリに対してスケジュールを作成
for index, row in filtered_schedule.iterrows():
    device_name = row['設備']
    process = row['エリア']
    incharge = row['オンラインテスト担当者']
    drawing_no = row['図面装置No']
    start_date = row['初号機テスト実施時期']
    start_month = start_date.strftime('%Y-%m')
    test_hours_needed = 40  # テストに必要な時間

    # 担当者の労働可能時間を取得
    if incharge in available_hours_dict and start_month in available_hours_dict[incharge]:
        if available_hours_dict[incharge][start_month] >= test_hours_needed and monthly_capacity_dict.get(start_month, 0) > 0:
            # スケジュールに追加
            schedule = schedule.append({
                '年月': start_month,
                'エリア': process,
                '設備': device_name,
                '図面装置No': drawing_no,
                'オンラインテスト担当者': incharge
            }, ignore_index=True)
            # 労働可能時間とテスト可能台数を減算
            available_hours_dict[incharge][start_month] -= test_hours_needed
            monthly_capacity_dict[start_month] -= 1

# スケジュールを年月とエリアでソート
schedule = schedule.sort_values(by=['年月', 'エリア'])

# スケジュールをCSVファイルに保存
schedule.to_csv('test_schedule_updated.csv', index=False)
print("スケジュールがCSVファイルに保存されました。")