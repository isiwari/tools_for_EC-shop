#coding: UTF-8
# 楽天、Yahooはそのまま店舗にアップロードできるcsv生成
# 自社店、標準データ用に検索用キーワードを表示

import PySimpleGUI as sg
import io
import itertools
import pandas as pd
import pathlib
import re
import csv
import subprocess

sg.theme('Dark Blue 3')
fin_flg= 0
ctrCul, dir, code_txt= '', '', ''
indata= 'コントロールカラム,商品管理番号（商品URL）\n'
indata_y= 'code\n'
txtColor= ['#ffffff', '#ffffcc', '#ccffcc', '#ccffff', '#ccccff', '#ffccff', '#ffcccc']

for i in itertools.cycle(txtColor):
    if fin_flg== 1:
        break
    
    # インプットウィンドウ。「event, values=...」はpg固有
    event, values= sg.Window('楽天、Yahooの商品削除用データを作成', [ \
    [sg.T('商品コード\n（標準）', text_color= i,), sg.In('ABC-1', key='-CODE-')], \
    [sg.Submit(' OK / 続けて入力 '), sg.Cancel() ], \
    [sg.Checkbox('入力を終了しCSVを作成する', key='-SAVE-')], \
    ], keep_on_top=True, auto_size_text=True, element_padding=(10,10), use_default_focus=True ).read(close=True)
    
    # ウィンドウ自体が閉じられた/キャンセルボタンが押されたら以降の処理を全てキャンセル
    if event in (None, 'Cancel'):
        sg.popup('処理を中止します',title='ウィンドウクローズ / キャンセルボタン', keep_on_top=True)
        exit()
    
    # 1行分を書き込む
    indata += ',,' + values['-CODE-'] + '\n'
    indata_y += values['-CODE-'] + '\n'
    code_txt += values['-CODE-'] + ','
    
    # 入力を続ける間この分岐は通らない
    if values['-SAVE-']:
        fin_flg= 1
        dir= sg.popup_get_folder(' csv出力先 ', title='フォルダ選択', keep_on_top=True)
        if event in (None, 'Cancel'):
            sg.popup('処理を中止します',title='ウィンドウクローズ / キャンセルボタン\n（キャンセルボタン受付）', keep_on_top=True)
            exit()
        else:
            sg.Window('csv書き出し', [ \
            [sg.T('商品削除用に\n楽天用：item.csv\nYahoo用：add_data.csv\n\nを出力します')], \
            [sg.T('自社店 / 標準用コードをコピー'), sg.In(code_txt)], \
            [sg.Submit(' OK ')], \
            ], keep_on_top=True, auto_size_text=True, element_padding=(10,10)).read(close=True)
        
    

df= pd.read_csv(io.StringIO(indata), encoding='shift-jis')

# 小文字に変換、必須フラグ設定※「.astype(str)」は数字のみの商品コード対策のため
df['コントロールカラム']= df['コントロールカラム'].fillna('d')
df['商品管理番号（商品URL）']= df['商品管理番号（商品URL）'].astype(str).str.lower()

subprocess.Popen(['explorer', pathlib.Path(dir)],shell=True)

df.to_csv(dir + r'\item.csv', encoding='shift-jis', index=False)

with open(dir + r'\add_data.csv', 'w') as y:
    w= csv.writer(y, lineterminator='\n')
    w.writerow([indata_y])
