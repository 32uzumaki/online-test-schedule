import pandas as pd

# =============================================================================
# 1) Excelファイルから累計データを読み込む
# =============================================================================
excel_file = "TestData.xlsx"  # エクセルファイルのパス
sheet_name = "TestSheet"      # シート名（実際の名前を指定）

# header=0: 1行目を列名にする
df_cumulative = pd.read_excel(
    excel_file,
    sheet_name=sheet_name,
    header=0
)

# もしA列の列名が "Date" ではなく、「Unnamed: 0」などの列名になっている場合はリネーム
if df_cumulative.columns[0] != "Date":
    df_cumulative.rename(columns={df_cumulative.columns[0]: "Date"}, inplace=True)

# Excelから読み取った列が "SubBE", "EPI", "WP表", "WP裏" という名前になっていると仮定
priority_order = ["SubBE", "EPI", "WP表", "WP裏"]

# 1か月あたり最大テスト台数
MAX_PER_MONTH = 3

# =============================================================================
# 2) "Date"列を日付型に変換し、データを整形
# =============================================================================
df_cumulative["Date"] = pd.to_datetime(df_cumulative["Date"], errors="coerce")
# 日付がNaTの行は削除し、昇順に並べ替え
df_cumulative = df_cumulative.dropna(subset=["Date"]).sort_values("Date").reset_index(drop=True)

if df_cumulative.empty:
    print("Excelに有効な日付データがありません。処理を終了します。")
else:
    # =============================================================================
    # 3) 「累計」→「月ごとの増分」に変換
    # =============================================================================
    df_diff = df_cumulative.copy()
    for i in range(len(df_diff)):
        if i == 0:
            # 初月は前月が無いので、そのまま累計値を「増分」とみなす
            pass
        else:
            # 前行との差分をとる
            for col in priority_order:
                df_diff.loc[i, col] = df_cumulative.loc[i, col] - df_cumulative.loc[i - 1, col]

    # =============================================================================
    # 4) テストシミュレーション (月次)
    # =============================================================================
    backlog = {p: 0 for p in priority_order}  # バックログ(未テスト台数)管理
    results = []

    # 差分データ行を走査するためのインデックス
    idx = 0
    max_idx = len(df_diff)

    # シミュレーション開始: 最初の行の日付(Excelで最も古い日付)
    current_date = df_diff.loc[0, "Date"]

    while True:
        # 今月の月初(date_key)で処理する (normalizeで時刻を00:00に統一)
        date_key = current_date.normalize()

        # (a) 差分データがあれば、当月に追加
        if idx < max_idx:
            row_date = df_diff.loc[idx, "Date"].normalize()
            # 当月が差分データの行と同じ日付なら、追加台数をバックログに加算
            if row_date == date_key:
                for p in priority_order:
                    backlog[p] += df_diff.loc[idx, p]
                idx += 1  # 次の行へ

        # (b) 優先度順に最大3台テスト
        remaining = MAX_PER_MONTH
        tested = {p: 0 for p in priority_order}
        for p in priority_order:
            if backlog[p] > 0 and remaining > 0:
                can_test = min(backlog[p], remaining)
                tested[p] = can_test
                backlog[p] -= can_test
                remaining -= can_test

        # (c) 今月末時点バックログを記録
        row_result = {
            "Month": date_key.strftime("%Y-%m-%d"),
        }
        # 今月テストした台数
        for p in priority_order:
            row_result[f"Tested_{p}"] = tested[p]
        # 月末バックログ
        for p in priority_order:
            row_result[f"Backlog_{p}"] = backlog[p]

        results.append(row_result)

        # (d) バックログが全て0 & 差分データの行をすべて取り込み済みなら終了
        if all(backlog[p] == 0 for p in priority_order) and (idx >= max_idx):
            break

        # (e) 翌月1日に進む
        y = current_date.year
        m = current_date.month
        if m == 12:
            y += 1
            m = 1
        else:
            m += 1
        current_date = pd.Timestamp(y, m, 1)

    # =============================================================================
    # 5) 結果表示
    # =============================================================================
    df_results = pd.DataFrame(results)
    print("▼ テスト進捗状況(各月ごとのテスト実施数・月末の残数)")
    display(df_results)

    final_month = df_results["Month"].iloc[-1]
    print(f"\nすべての装置のテストが完了したのは {final_month} 時点です。")