import pandas as pd

# CSVファイルを読み込む
file_path = "/Users/komatsutomoaki/Desktop/Sample_Data_Horizontal__Monthly_Integers__2024-12_to_2025-12_のコピー.csv"  # ファイルパスを指定
data = pd.read_csv(file_path)

# データの整形
data = data.set_index('Unnamed: 0').T.reset_index()
data.columns = ['Month', 'Testable']
data['Testable'] = data['Testable'].astype(int)

# テスト上限を固定値として設定
test_limit = 3

# 計算用の列を初期化
data['Tested'] = 0
data['CarryOver'] = 0

carry_over = 0  # 初期繰越

# 月ごとの計算
for i in range(len(data)):
    available_units = data.loc[i, 'Testable'] + carry_over
    tested_units = min(available_units, test_limit)
    carry_over = available_units - tested_units

    data.loc[i, 'Tested'] = tested_units
    data.loc[i, 'CarryOver'] = carry_over

# テストが終わるまで月を追加
while carry_over > 0:
    next_month = f"Month {len(data) + 1}"
    data = pd.concat(
        [
            data,
            pd.DataFrame({"Month": [next_month], "Testable": [0], "Tested": [0], "CarryOver": [0]}),
        ],
        ignore_index=True,
    )

    available_units = carry_over
    tested_units = min(available_units, test_limit)
    carry_over = available_units - tested_units

    data.loc[len(data) - 1, 'Tested'] = tested_units
    data.loc[len(data) - 1, 'CarryOver'] = carry_over

# データを月を横軸に転置
transposed_data = data.set_index('Month').T

# 結果を保存
output_path = '/Users/komatsutomoaki/Desktop/online-test/online-test-schedule/Calculated_Test_Results1.csv'
transposed_data.to_csv(output_path)

print(f"計算結果が保存されました（横軸に月を設定）: {output_path}")
