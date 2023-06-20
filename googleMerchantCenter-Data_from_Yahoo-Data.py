#coding: UTF-8

import subprocess
import pandas as pd
import PySimpleGUI as sg
import pathlib
sg.theme('DarkTeal')

# インプットcsvを指定、親ディレクトリ取得
indata= sg.popup_get_file('元にするYahooデータを指定してください',title='インプットcsv選択', file_types=(('CSV', '*.csv'),),)
dir= pathlib.Path(indata).parent

# pandasでcsvを開く
df= pd.read_csv(indata, encoding='shift-jis')

# 欠損値を埋める※あれば結合結果も欠損値になってしまう
df['headline']= df['headline'].fillna(' ')
df['explanation']= df['explanation'].fillna(' ')

# 列名を変更
df.rename(columns={ \
'code': 'id', \
'name': 'タイトル', \
'headline': '概要1', \
'explanation': '概要2', \
'price': '価格', \
'jan': 'gtin', \
# 'product-code': '製品番号', \
'postage-set': '送料ラベル', \
'lead-time-instock': 'お届け日数ラベル' \
}, inplace=True)

# 対象列を指定し置換
df['概要2']= df['概要2'].replace(r"\n",r"<br>",regex=True)

# 列同士を結合※dtypeを揃えた上で結合
df['概要']= df['概要1'].astype(str) + '<br>' + df['概要2'].astype(str)

# 商品コードをコピー加工し新たな列を作成
df['商品リンク'] =r'https://www.ec-life.co.jp/fs/eclife/' + df['id'] + r'?utm_source=google&utm_medium=cpc&utm_campaign=freeListings'
df['商品画像リンク'] = r'https://image.rakuten.co.jp/eco-life-r/cabinet/tmb2/' + df['id'] + r'.jpg'
df['製品番号']= df['id'].str.upper()

# 列名を指定して並び替え
df= df.reindex(columns=[ \
'id', \
'タイトル', \
'概要', \
'商品リンク', \
'商品画像リンク', \
'状態', \
'価格', \
'在庫状況', \
'ブランド', \
'gtin', \
'製品番号', \
'google 商品カテゴリ', \
'送料ラベル', \
'お届け日数ラベル', \
])

# 列ごとに欠損値を異なる値で置換
df.fillna({ \
'状態': '新品', \
'在庫状況': '在庫あり', \
}, inplace=True)

# csv書き出し
df.to_csv(str(dir) + r'\gmc_out.csv', encoding='shift-jis', index=False)

# フォルダを開く
subprocess.Popen(['explorer',dir],shell=True)
