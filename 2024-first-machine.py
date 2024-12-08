import pandas as pd
from datetime import datetime, timedelta

# ファイルパス
equipment_schedule_path = "/2.csv"

# データ読み込み
equipment_schedule = pd.read_csv(equipment_schedule_path)

# 必要な列を選択、リリース実績もカウントできるようにしてたい。
new_equipment_schedule = equipment_schedule[['工程', '機種名', 'リリース予定日', '開発テスト完了予定日','受入テスト実施日']]

#リリース予定日、受入テスト実施日(装置アドレス)、初講義テスト実施実機(初号機のテストスケジュールより算出)
#確定分だけでいいのなら、初号機テスト実施実機いらない。
#DXC開発スケジュールを基に算出する必要があるのなら、初号機のテストスケジュールを作成してやる必要性がある。

# 有効なエリアのみをフィルタリング
valid_areas = ['SubBE', 'EPI', 'WP表', 'WP裏', 'EDS']
filtered_schedule = new_equipment_schedule[new_equipment_schedule['工程'].isin(valid_areas)]

"""
# 新しい条件に基づいてデータをフィルタリング
filtered_schedule = new_equipment_schedule[
    (new_equipment_schedule['オンライン対応'] == '〇') &
    (new_equipment_schedule['装置型式毎の初回テスト対象'] == '〇')
]
"""

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
filtered_schedule['開発テスト完了予定日'] = filtered_schedule['開発テスト完了予定日'].apply(validate_date)
filtered_schedule['受入テスト実施日'] = filtered_schedule['受入テスト実施日'].apply(validate_date)

# 新しい列を追加: 調整後のテスト実施時期
def determine_test_date(row):
    if pd.notna(row['受入テスト実施日']):
        # 受入テスト実施日が設定されている場合、その月を確定
        return row['受入テスト実施日']+ pd.DateOffset(months=1)
    elif pd.isna(row['リリース予定日']) or pd.isna(row['開発テスト完了予定日']):
        # リリース予定日または初号機テスト実施時期が欠損している場合はNaT
        return pd.NaT
    elif row['リリース予定日'] < row['開発テスト完了予定日']:
        # リリース予定日が初号機テスト実施時期より早い場合、翌月を設定
        return row['開発テスト完了予定日'] + pd.DateOffset(months=1)
    elif row['リリース予定日'] > row['開発テスト完了予定日']:
        # 初号機テスト実施時期がリリース予定日より早い場合、リリース予定日を設定
        return row['リリース予定日']
    else:
        # 同じ月の場合、翌月を設定
        return row['開発テスト完了予定日'] + pd.DateOffset(months=1)

filtered_schedule['調整後テスト実施時期'] = filtered_schedule.apply(determine_test_date, axis=1)

# 範囲外データの抽出
start_range = pd.Timestamp('2024-10-01')
end_range = pd.Timestamp('2026-03-31')

out_of_range_schedule = filtered_schedule[
    (filtered_schedule['調整後テスト実施時期'] < start_range) |
    (filtered_schedule['調整後テスト実施時期'] > end_range)
]

# 範囲内のデータのみ保持
filtered_schedule = filtered_schedule[
    (filtered_schedule['調整後テスト実施時期'] >= start_range) &
    (filtered_schedule['調整後テスト実施時期'] <= end_range)
]

# 月列と工程列を基準にデータをピボット
filtered_schedule['年月'] = filtered_schedule['調整後テスト実施時期'].dt.strftime('%Y-%m')

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
date_columns = pd.date_range('2024-10', '2026-03', freq='M').strftime('%Y-%m').tolist()
pivot_table = pivot_table.reindex(columns=date_columns + ['搬入日未定', '範囲外'], fill_value='')

# CSVファイルに保存
pivot_table.to_csv('range3.csv', encoding='utf-8-sig')
print("スケジュールがCSVファイルに保存されました。")