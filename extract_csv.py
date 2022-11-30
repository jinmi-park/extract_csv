import sys
import os
import pandas as pd

csv = '.csv'
excel = '.xlsx'
args = sys.argv

folder_path = os.path.dirname(os.path.abspath(__file__))
target_file = args[1]
target_path = folder_path + '\\' + target_file
# CSV出力対象シートのDataFrame
df = pd.read_excel(target_path + excel, sheet_name=args[2])

# 日付の型変換
df['日付'] = pd.to_datetime(df['日付'], format="%Y/%m/%d")

# 数値の型変換
for col_name in ['明細No.', '借方勘定科目ｺｰﾄﾞ', '借方補助科目ｺｰﾄﾞ', '借方本体金額', '貸方勘定科目ｺｰﾄﾞ', '貸方本体金額']:
    df[col_name] = df[col_name].astype('Int64')

output_path = folder_path + '\\' + args[2] + csv
df.to_csv(output_path, index=None, encoding='cp932')

# Galileoptの仕様上一度開いて閉じる
csv_file = open(output_path, mode="w")
csv_file.close()

