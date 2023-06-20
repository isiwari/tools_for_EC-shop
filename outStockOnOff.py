#coding: UTF-8
# 欠品/解除用csv生成
# 自社店→楽天の順で作成※Yahooはnewフラグデータのみ出力

import PySimpleGUI as sg
import io
import itertools
import pandas as pd
import pathlib
import re
import subprocess

sg.theme('Dark Blue 3')
fin_flg= 0
ctrCul, dir, code_txt= '', '', ''
realtimeLocation= (100, 200)
# デフォルトは # realtimeLocation= (None, None)
indata= 'コントロールカラム,商品URLコード,選択肢タイプ,セレクト／ラジオボタン用項目名,セレクト／ラジオボタン用選択肢\n'
indata_y= 'code,option-name-1,option-value-1,unselectable-1\n'
indata_y_off= 'code,options\n'
txtColor= ['#ffffff', '#ffffcc', '#ccffcc', '#ccffff', '#ccccff', '#ffccff', '#ffcccc']

for i in itertools.cycle(txtColor):
    if fin_flg== 1:
        break
    
    # インプットウィンドウのレイアウト
    layout=[
    [sg.Radio('new', "g1", default=True, key='-NEW-'), sg.Radio('delete', "g1", key='-DELETE-')], \
    [sg.T('商品コード\n（標準）', text_color= i,), sg.In('XXXXX', key='-CODE-')], \
    [sg.T('選択項目名'), sg.In('配送について', key='-optionName-')], \
    [sg.T('項目選択肢'), sg.In('00月以降のお届けに同意、上中下', key='-optionValue-')], \
    [sg.Submit(' OK / 続けて入力 '), sg.Cancel() ], \
    [sg.Checkbox('入力を終了しCSVを作成する', key='-SAVE-')], \
    ]
    window= sg.Window('自社店、楽天、Yahooの項目選択肢データを作成', \
    keep_on_top=True, auto_size_text=True, element_padding=(10,10), use_default_focus=True, location=realtimeLocation ).Layout(layout)
    
    event, values= window.read() # 「event, values=...」はpg固有
    
    # ウィンドウ自体が閉じられた/キャンセルボタンが押されたら以降の処理を全てキャンセル
    if event in (None, 'Cancel'):
        sg.popup('処理を中止します',title='ウィンドウクローズ操作 / キャンセルボタン を受け付けました', keep_on_top=True, location=realtimeLocation)
        exit()
    
    # 現在のウィンドウ座標を取得してから、ウィンドウを閉じる。.read(close=True)だとエラー排出。
    realtimeLocation= window.CurrentLocation()
    window.close()
    
    # コントロールフラグを設定
    if values['-NEW-']:
        ctrCul= 'n'
        indata_y += values['-CODE-'] + ',' + values['-optionName-'] + ',選択してください,1\n' \
        + values['-CODE-'] + ',' + values['-optionName-'] + ',' + values['-optionValue-'] + ',0\n'
    elif values['-DELETE-']:
        ctrCul= 'd'
        indata_y_off += values['-CODE-'] + ',\n'
    
    # 1行分を書き込む
    indata += ctrCul + ',' + values['-CODE-'] + ',s,' + values['-optionName-'] + ',' + values['-optionValue-'] + '\n'
    code_txt += values['-CODE-'] + ','
    
    # 入力を続ける間この分岐は通らない
    if values['-SAVE-']:
        fin_flg= 1
        dir= sg.popup_get_folder(' csv出力先となるフォルダを選択してください ', title='出力先フォルダ選択', keep_on_top=True, location=realtimeLocation)
        if dir==None :
            sg.popup('csv出力処理を中止します',title='csv出力先が指定されませんでした', keep_on_top=True, location=realtimeLocation)
            exit()
        else:
            sg.Window('csv書き出し', [ \
            [sg.T('自社店用：selectFS.csv\n楽天用：select.csv\nYahoo用：option_add_0.csv\n\nを出力します')], \
            [sg.T('コードをコピー'), sg.In(code_txt)], \
            [sg.Submit(' OK ')], \
            ], keep_on_top=True, auto_size_text=True, element_padding=(10,10), location=realtimeLocation).read(close=True)
        
    

df= pd.read_csv(io.StringIO(indata), encoding='shift-jis')

df= df.reindex(columns=[ \
'コントロールカラム', '商品URLコード', '選択肢タイプ', 'セレクト／ラジオボタン用項目名', 'セレクト／ラジオボタン用選択肢', '項目選択肢前改行', '項目名位置', '項目選択肢表示', 'テキスト幅', 'テキスト幅（モバイル）', '最終更新日時' \
])

df_r= df.rename(columns={ \
'コントロールカラム': '項目選択肢用コントロールカラム', \
'商品URLコード': '商品管理番号（商品URL）', \
'選択肢タイプ': '選択肢タイプ', \
'セレクト／ラジオボタン用項目名': '項目選択肢項目名', \
'セレクト／ラジオボタン用選択肢': '項目選択肢', \
'項目選択肢前改行': '項目選択肢別在庫用横軸選択肢' \
})

df_r= df_r.reindex(columns=[ \
'項目選択肢用コントロールカラム', '商品管理番号（商品URL）', '選択肢タイプ', '項目選択肢項目名', '項目選択肢', '項目選択肢別在庫用横軸選択肢', '項目選択肢別在庫用横軸選択肢子番号', '項目選択肢別在庫用縦軸選択肢', '項目選択肢別在庫用縦軸選択肢子番号', '項目選択肢別在庫用取り寄せ可能表示', '項目選択肢別在庫用在庫数', '在庫戻しフラグ', '在庫切れ時の注文受付', '在庫あり時納期管理番号', '在庫切れ時納期管理番号', 'タグID', '画像URL', '項目選択肢選択必須' \
])

# 小文字に変換、必須フラグ設定※「.astype(str)」は数字のみの商品コード対策のため
df_r['商品管理番号（商品URL）']= df_r['商品管理番号（商品URL）'].astype(str).str.lower()
df_r['項目選択肢選択必須']= df_r['項目選択肢選択必須'].fillna('1')

df_y= pd.read_csv(io.StringIO(indata_y), encoding='shift-jis')
df_y= df_y.reindex(columns=[ \
'code', 'sub-code', 'option-name-1', 'option-value-1', 'spec-id-1', 'spec-value-id-1', 'option-charge-1', 'unselectable-1', 'option-name-2', 'option-value-2', 'spec-id-2', 'spec-value-id-2', 'lead-time-instock', 'lead-time-outstock', 'sub-code-img1', 'main-flag' \
])
df_y_off= pd.read_csv(io.StringIO(indata_y_off), encoding='shift-jis')

subprocess.Popen(['explorer', pathlib.Path(dir)],shell=True)

df.to_csv(dir + r'\selectFS.csv', encoding='shift-jis', index=False)
df_r.to_csv(dir + r'\select.csv', encoding='shift-jis', index=False)
df_y.to_csv(dir + r'\option_add_0.csv', encoding='shift-jis', index=False)
df_y_off.to_csv(dir + r'\add_data_0.csv', encoding='shift-jis', index=False)
