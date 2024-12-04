import pandas as pd
from datetime import datetime

# ファイルパス
equipment_schedule_path = "/Users/komatsutomoaki/Desktop/online-test/online-test-schedule/装置搬入スケジュール.csv"

# データ読み込み
equipment_schedule = pd.read_csv(equipment_schedule_path)

# 必要な列を選択
new_equipment_schedule = equipment_schedule[['工程', '機種名', 'リリース予定日', '初号機テスト実施時期']]

# 有効なエリアのみをフィルタリング
valid_areas = ['SubBE', 'EPI', 'WP表', 'WP裏', 'EDS']
filtered_schedule = new_equipment_schedule[new_equipment_schedule['工程'].isin(valid_areas)]

# 日付検証関数
def validate_date(value):
    try:
        return pd.to_datetime(value)
    except:
        return pd.NaT

# 日付列の整備
filtered_schedule['リリース予定日'] = filtered_schedule['リリース予定日'].apply(validate_date)
filtered_schedule['初号機テスト実施時期'] = filtered_schedule['初号機テスト実施時期'].apply(validate_date)

# '初号機テスト実施時期'がNaTの場合、'リリース予定日'の月を代入
filtered_schedule['初号機テスト実施時期'] = filtered_schedule.apply(
    lambda row: row['リリース予定日'] if pd.isna(row['初号機テスト実施時期']) else row['初号機テスト実施時期'],
    axis=1
)

# 'リリース予定日' > '初号機テスト実施時期'の場合、'リリース予定日'の月にテストを実施
filtered_schedule['初号機テスト実施時期'] = filtered_schedule.apply(
    lambda row: row['リリース予定日'] if row['リリース予定日'] > row['初号機テスト実施時期'] else row['初号機テスト実施時期'],
    axis=1
)

# 範囲外の日付を除外（2024-10-01から2026-03-31の間）
start_range = pd.Timestamp('2024-10-01')
end_range = pd.Timestamp('2026-03-31')
filtered_schedule = filtered_schedule[
    (filtered_schedule['初号機テスト実施時期'] >= start_range) &
    (filtered_schedule['初号機テスト実施時期'] <= end_range)
]

# 月列と工程列を基準にデータをピボット
filtered_schedule['年月'] = filtered_schedule['初号機テスト実施時期'].dt.strftime('%Y-%m')

# ピボットテーブルの作成
pivot_table = filtered_schedule.pivot_table(
    index='工程',  # 行に工程名
    columns='年月',  # 列に年月
    values='機種名',  # セルに機種名
    aggfunc=lambda x: '\n'.join(x)  # 同じ月・工程で機種名を改行して表示
)

# 欠損値を空文字で埋める
pivot_table = pivot_table.fillna('')

# 列を2024年10月から2026年3月の順に並べ替え
date_columns = pd.date_range('2024-10', '2026-03', freq='M').strftime('%Y-%m')
pivot_table = pivot_table.reindex(columns=date_columns, fill_value='')

# CSVファイルに保存
pivot_table.to_csv('test_schedule_pivot.csv', encoding='utf-8-sig')
print("スケジュールがCSVファイルに保存されました。")