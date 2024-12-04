import pandas as pd

# CSVファイルを読み込む
file_path = "24-12_to_2025-12_.csv"  # ファイルパスを指定
data = pd.read_csv(file_path)

# データの整形
data = data.set_index('Unnamed: 0').T.reset_index()
data.columns = ['Month', 'Testable', 'TestLimit']
data['Testable'] = data['Testable'].astype(int)
data['TestLimit'] = data['TestLimit'].astype(int)

# 計算用の列を初期化
data['Tested'] = 0
data['CarryOver'] = 0

carry_over = 0  # 初期繰越

# 月ごとの計算
for i in range(len(data)):
    available_units = data.loc[i, 'Testable'] + carry_over
    tested_units = min(available_units, data.loc[i, 'TestLimit'])
    carry_over = available_units - tested_units

    data.loc[i, 'Tested'] = tested_units
    data.loc[i, 'CarryOver'] = carry_over

# テストが終わるまで月を追加
while carry_over > 0:
    next_month = f"Month {len(data) + 1}"
    data = pd.concat(
        [
            data,
            pd.DataFrame({"Month": [next_month], "Testable": [0], "TestLimit": [3], "Tested": [0], "CarryOver": [0]}),
        ],
        ignore_index=True,
    )

    available_units = carry_over
    tested_units = min(available_units, data.loc[len(data) - 1, 'TestLimit'])
    carry_over = available_units - tested_units

    data.loc[len(data) - 1, 'Tested'] = tested_units
    data.loc[len(data) - 1, 'CarryOver'] = carry_over
# データを月を横軸に転置
transposed_data = data.set_index('Month').T
# 結果を保存
output_path = 'Calculated_Test_Results.csv'
transposed_data.to_csv(output_path, index=False)

print(f"計算結果が保存されました: {output_path}")
