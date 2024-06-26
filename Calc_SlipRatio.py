# ---/実験データ（CSVファイル）の自動処理/---
# 作成日：2022.09.14
# 変更日：2022.11.08 Calc＿SlipRatio
# 
# ---/参考URL/---
# PySimpleGui : https://fuji-pocketbook.net/pysimplegui/
# Pyautogui : https://tech-blog.rakus.co.jp/entry/20220613/python
# pyinstaller(仮想環境) : https://teshi-learn.com/2020-12/windows-pyinstaller-python-exe-anaconda/
# 
# ---/コンパイル方法/----
# 仮想環境の確認 -> conda info -e (この中のenv_exeがコンパイル用仮想環境)
# 仮想環境のアクティベート -> activate env_exe
# pyinstallerの実行 -> pyinstaller *****.py --onefile --noconsole
#   （--onefile にすると一つのexeファイルとして生成される）
#   （--noconsole を追加するとexeファイル起動後に黒い画面が生成されない）
# 
# ---/MCデータの抜き出し/---
# 元データからモーションキャプチャの移動点だけを抜き出す
# ※マーカー1点のデータのみ
#
# ---/MCデータのオフセット/---
# X軸のデータを抜き出し，スリップ率の計算をする
# エクセルファイルにすべてまとめる
# 
# 
# ---/DataLoggerデータのオフセット/---
# 
# ---/FSデータのオフセット/---
# ※今の所実装予定無し


# ---/ライブラリのインポート/---
import glob
import pandas as pd
import PySimpleGUI as sg
import math
import numpy as np
import csv

sg.theme('Dark2')       # PySimpleguiのデザイン設定

# ---/GUI画面の出力/---
def main():
    # ---/MCデータの抜き出しのフレーム/---
    FRAME_MC_DATA_EXTRACTION = sg.Frame('MCデータの抜き出し',[
        [sg.Button('参照先'),sg.Text('←MCデータの保存されているフォルダを選択',font=('Yu Gothic UI Semibold',11),key = '-REFERENCE-')],
        [sg.Text('X軸の列座標:'),sg.Input(key = '-X_COODINATE-',size=(6,1))],
        [sg.Button('保存先'),sg.Text('←抜き出したデータの保存先フォルダを選択',font=('Yu Gothic UI Semibold',11),key = '-SAVE_PATH-')],
        [sg.Button('MCデータ抜き出し'),sg.Text('',key = '-DISPLAY1-')]
    ])

    # ---/MCデータのオフセット/---
    FRAME_MC_DATA_OFFSET = sg.Frame('MCデータのオフセット',[
        [sg.Button('開く'),sg.Text('←MCデータ抜き出しで保存したフォルダを選択',font=('Yu Gothic UI Semibold',11),key = '-MC_OFFSET_PATH-')],
        [sg.Text('X軸座標'),sg.Text('最小値'),sg.Input(key = '-MC_OFFSET_MIN-',size=(9,1)),sg.Text('最大値'),sg.Input(key = '-MC_OFFSET_MAX-',size=(9,1)),
        sg.Text('車輪径[mm]'),sg.Input(key = '-DIAMETER-',size=(9,1)),sg.Text('回転数[rpm]'),sg.Input(key = '-ROTATION-',size=(9,1))],
        [sg.Button('保存先の選択'),sg.Text('←オフセットデータの保存先を選択',font=('Yu Gothic UI Semibold',11),key = '-SAVE_MC_OFFSET_PATH-')],
        [sg.Text('出力ファイル名:'),sg.Input(key = '-OUT_CSV_NAME-',size=(30,1)),sg.Text('例）Offset_Output(拡張子不要)')],
        [sg.Button('MCデータオフセット'),sg.Text('',key = '-DISPLAY2-')]
    ])

    layout = [
        [FRAME_MC_DATA_EXTRACTION],
        [FRAME_MC_DATA_OFFSET]
    ]

    window = sg.Window('モーションキャプチャの１点データ自動処理.exe',layout,resizable=True)

    while True:
        event,values = window.read()
        if event == sg.WIN_CLOSED:
            break

        if event == '参照先':
            REFERENCE = sg.popup_get_folder('MCデータの保存されているフォルダを選択',font=('Yu Gothic UI Semibold',12),location=(700,250))
            window['-REFERENCE-'].update(REFERENCE)
        
        if event =='保存先':
            SAVE_PATH = sg.popup_get_folder('保存先を選択',font=('Yu Gothic UI Semibold',12),location=(700,250))
            window['-SAVE_PATH-'].update(SAVE_PATH)
            
        if event == 'MCデータ抜き出し':
            if SAVE_PATH == "" or SAVE_PATH == None:
                sg.popup('保存先フォルダを選択してください')  # OK ボタンを表示
                window['-SAVE_PATH-'].update('←抜き出したデータの保存先フォルダを選択')
            else:
                MC_DATA_EXTRACTION(REFERENCE, SAVE_PATH, values['-X_COODINATE-'])
                window['-DISPLAY1-'].update('出力完了')
        
        if event == '開く':
            MC_OFFSET_PATH = sg.popup_get_folder('MCデータの保存されているフォルダを選択',font=('Yu Gothic UI Semibold',12),location=(700,250))
            window['-MC_OFFSET_PATH-'].update(MC_OFFSET_PATH)

        if event =='保存先の選択':
            SAVE_MC_OFFSET_PATH = sg.popup_get_folder('オフセットデータの保存先を選択',font=('Yu Gothic UI Semibold',12),location=(700,250))
            window['-SAVE_MC_OFFSET_PATH-'].update(SAVE_MC_OFFSET_PATH)

        if event == 'MCデータオフセット':
            MC_DATA_OFFSET(MC_OFFSET_PATH,SAVE_MC_OFFSET_PATH,values['-MC_OFFSET_MIN-'],values['-MC_OFFSET_MAX-'],values['-DIAMETER-'],values['-ROTATION-'],values['-OUT_CSV_NAME-'])
            window['-DISPLAY2-'].update('出力完了')
    window.close()


# ---/MCデータの抜き出し/---
def MC_DATA_EXTRACTION(MC_PATH,MC_SAVE_PATH,X_COODINATE):
    X = int(X_COODINATE) - 1                                             # X軸の列座標
    Y = X + 1                                                            # CSV上の表示列座標とPythonの数値の始まりが違う（0を1カウントするかどうか）
    Z = Y + 1
    strings = str(MC_PATH)
    MC_CSV_PATH = strings + "\*csv"
    csv_file = glob.glob(MC_CSV_PATH)                           # パス内にあるCSVファイル名を取得

    for file_name_i in csv_file:
        df = pd.read_csv(file_name_i, header = 5)               # CSVファイルのヘッダー位置，これ以降の行データを取得する
        df_1 =pd.DataFrame(
            data = {'Times(Second)': pd.Series(df.iloc[:,1]),
                    'X': pd.Series(df.iloc[:,X]),
                    'Y': pd.Series(df.iloc[:,Y]),
                    'Z': pd.Series(df.iloc[:,Z]),
            }
        )

        rename = MC_SAVE_PATH + file_name_i.strip(strings).strip(".csv") + "_MC" + ".csv"             # 保存先のフルパスを指定
        df_1.to_csv(rename)


# ---/MCデータのオフセット/---
def MC_DATA_OFFSET(MC_OFFSET_PATH,SAVE_OFFSET_PATH,MC_OFFSET_MIN,MC_OFFSET_MAX,DIAMETER,ROTATION,OUT_CSV_NAME):
    strings = str(MC_OFFSET_PATH)                                       # フォルダパス取得
    MC_CSV_OFFSET_PATH = strings + "\*csv"                              # フォルダ内のCSVファイル名取得
    csv_file = glob.glob(MC_CSV_OFFSET_PATH)                            # フォルダ内にあるすべてのファイル名をリスト化
    
    INPUT_DATA = np.array([['入力データ'],['最小値:' + str(MC_OFFSET_MIN) + '[m]'],['最大値' + str(MC_OFFSET_MAX) + '[m]'],['車輪径' + str(DIAMETER) + '[mm]'],['回転数' + str(ROTATION) + '[rpm]']])
    BLANK = np.array([['*'],['*'],['*'],['*'],['*']])
    INDEX = np.array([['File name'],['Slip ratio[・]'],['X_MAX[m]'],['X_MIN[m]'],['Time[sec]']])
    
    #書き込み用CSVファイル作成
    f = open(SAVE_OFFSET_PATH + '\\' + str(OUT_CSV_NAME) + ".csv", 'a', newline='')
    writer = csv.writer(f)
    writer.writerows(np.array(INPUT_DATA.reshape(1,-1)))
    writer.writerows(np.array(BLANK.reshape(1,-1)))
    writer.writerows(np.array(INDEX.reshape(1,-1)))
    
    for file_name_i in csv_file:
        df = pd.read_csv(file_name_i)                                   # CSVファイルの読み込み

        NAME = str(file_name_i).replace(strings,'').strip('\\')              # ファイル名から不要な部分を取り消し，実験データをわかりやすく表示する名前に変更

        X_ARRAY = np.array(df.X)
        X_ZERO_OFFSET = X_ARRAY - df.X.min()                            # 新規データフレームの作成，元データの”Ｘ”列から最小値をそれぞれ引き，ゼロオフセットする
        X_ARRAY_REMOVE_min = X_ZERO_OFFSET[X_ZERO_OFFSET > float(MC_OFFSET_MIN)]                              # オフセットしたい最小値以下の値を差し引いてデータフレームを再構成
        X_ARRAY_REMOVE_MinMax = X_ARRAY_REMOVE_min[X_ARRAY_REMOVE_min < float(MC_OFFSET_MAX)]                 # オフセットしたい最大値以上の値を差し引いてデータフレームを再構成
        
        X_ARRAY_OFFSET_MIN = float(X_ARRAY_REMOVE_MinMax.min())                        # オフセットの最大値
        X_ARRAY_OFFSET_MAX = float(X_ARRAY_REMOVE_MinMax.max())                        # オフセットの最大値
        
        # 車輪が止まった後に戻った座標データ分を削除する処理を行うパート（以下4行）
        X_preOFFSET = np.array(X_ARRAY_REMOVE_MinMax.reshape(-1,1))                     # 行の配列を列の配列に変更する   
        L1 = np.where(X_preOFFSET == X_ARRAY_OFFSET_MAX)[0][0] + 1                      # 最大値の配列が格納されている場所の隣の場所を探す（最大値以降のデータを消したいから）
        L2 = len(X_preOFFSET)                                                           # 最終行を取得（L1～最終行までがいらないデータとなる）
        X_ARRAY_OFFSET = np.delete(X_preOFFSET,slice(L1,L2),axis=0)                     # いらないデータを削除
        
        TIME = int(len(X_ARRAY_OFFSET)) * 0.01                                          # オフセットした値のデータの個数を数えて＊0.01(100Hz)すると時間になる
        REAL_SPEED = (X_ARRAY_OFFSET_MAX - X_ARRAY_OFFSET_MIN) * 1000 / TIME            # 速度の実測値，データの単位がｍだから㎜に変換
        IDEAL_SPEED = float(DIAMETER) * math.pi * float(ROTATION) / 60                  # 理想速度の計算[mm/s]
        SLIP_RATIO = (IDEAL_SPEED - REAL_SPEED) / IDEAL_SPEED                           # スリップ率の計算（V-v）/V
        
        # SET_ARRAYは一番重要な実験データの入る部分．その下にX_ARRAY_OFFSET_Xを追加したものが最終的なすべての列データ（ALL_ARRAY）になる
        SET_ARRAY = np.array([[NAME],[SLIP_RATIO],[X_ARRAY_OFFSET_MAX],[X_ARRAY_OFFSET_MIN],[TIME],['*']])
        ALL_ARRAY = np.concatenate([SET_ARRAY,X_ARRAY_OFFSET])
        a = np.array(ALL_ARRAY.reshape(1,-1))
        writer.writerows(a)
    
    f.close()


# ---/メイン関数/---
if __name__ == "__main__":
    main()