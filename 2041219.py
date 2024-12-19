import pandas as pd

# ファイルパス
equipment_schedule_path = "excel/装置搬入スケジュール2.csv"

# データ読み込み
equipment_schedule = pd.read_csv(equipment_schedule_path)

# 必要な列を選択
new_equipment_schedule = equipment_schedule[['工程', '機種名', 'リリース予定日', '増設機テスト実施時期']]

# 有効なエリアのみをフィルタリング
valid_areas = ['SubBE', 'EPI', 'WP表', 'WP裏']
filtered_schedule = new_equipment_schedule[new_equipment_schedule['工程'].isin(valid_areas)]

# 日付検証関数
def validate_date(value):
    try:
        return pd.to_datetime(value)
    except:
        return pd.NaT

# '搬入日未定'データを抽出
undecided_schedule = filtered_schedule[filtered_schedule['リリース予定日'] == '搬入日未定']

# リリース予定日と初号機テスト実施時期を検証
filtered_schedule['リリース予定日'] = filtered_schedule['リリース予定日'].apply(validate_date)

# 範囲外データの抽出
start_range = pd.Timestamp('2024-10-01')
end_range = pd.Timestamp('2026-03-31')

out_of_range_schedule = filtered_schedule[
    (filtered_schedule['リリース予定日'] < start_range) |
    (filtered_schedule['リリース予定日'] > end_range)
]

# 範囲内のデータのみ保持
filtered_schedule = filtered_schedule[
    (filtered_schedule['リリース予定日'] >= start_range) &
    (filtered_schedule['リリース予定日'] <= end_range)
]

# 月列と工程列を基準にデータをピボット
filtered_schedule['年月'] = filtered_schedule['リリース予定日'].dt.strftime('%Y-%m')

# ピボットテーブルの作成
pivot_table = filtered_schedule.pivot_table(
    index='工程',  # 行に工程名
    columns='年月',  # 列に年月
    values='機種名',  # セルに機種名
    aggfunc=lambda x: '\n'.join(x)  # 同じ月・工程で機種名を改行して表示
)

# 欠損値を空文字で埋める
pivot_table = pivot_table.fillna('')

# 範囲外データ列を追加
pivot_table['範囲外'] = out_of_range_schedule.groupby('工程')['機種名'].apply(lambda x: '\n'.join(x)).reindex(pivot_table.index, fill_value='')

# 搬入日未定列を追加
pivot_table['搬入日未定'] = undecided_schedule.groupby('工程')['機種名'].apply(lambda x: '\n'.join(x)).reindex(pivot_table.index, fill_value='')

# 列を2024年10月から2026年3月の順に並べ替え
date_columns = pd.date_range('2024-10', '2026-03', freq='ME').strftime('%Y-%m').tolist()
pivot_table = pivot_table.reindex(columns=date_columns + ['搬入日未定', '範囲外'], fill_value='')

# 工程のソート順を指定
custom_order = ['SubBE', 'EPI', 'WP表', 'WP裏']
pivot_table = pivot_table.reindex(custom_order)

# エクセルファイルに保存
# エクセルファイルに保存
output_path = '/U/test_schedule_with_out_of_range_sorted.xlsx'
pivot_table.to_excel(output_path, engine='openpyxl')
print(f"スケジュールがエクセルファイルに保存されました: {output_path}")
