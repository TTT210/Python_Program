# Arduino Data Logger
# 作成日：2022.09.06
# 変更日：2022.09.21
# 変更日：2022.11.07 V5 データロガーの保存名を”実験名” ＋ ”_DL”となるように変更
# 変更日：2022.11.08 V6 データの保存先を選択するポップアップの追加
# ====================================================================================================
# 参考URL
# PySimpleGui : https://fuji-pocketbook.net/pysimplegui/
# Pyautogui : https://tech-blog.rakus.co.jp/entry/20220613/python
# pyinstaller(仮想環境) : https://teshi-learn.com/2020-12/windows-pyinstaller-python-exe-anaconda/
# 
# ---Serial,CSV---
# https://opuktr.hatenablog.com/entry/2018/08/20/000801
# https://qiita.com/ryota765/items/0cfc2ea2d598de11b174
# ---コンパイル方法----
# 仮想環境の確認 -> conda info -e (この中のenv_exeがコンパイル用仮想環境)
# 仮想環境のアクティベート -> activate env_exe
# pyinstallerの実行 -> pyinstaller *****.py --onefile --noconsole
#   （--onefile にすると一つのexeファイルとして生成される）
#   （--noconsole を追加するとexeファイル起動後に黒い画面が生成されない）
# ---




import PySimpleGUI as sg
import csv
import serial
import time

sg.theme('Default')       # PySimpleguiのデザイン設定

def main():
    layout = [
            [sg.Text('ポート名　'),sg.Input(key = '-PORTNAME-', size=(10,1))],
            [sg.Text('ファイル名'),sg.Input(key = '-FILENAME-', size=(35,1)), sg.Button('保存先')],
            [sg.Button('START'),sg.Button('STOP')],
            [sg.Text('',key = '-DISPLAY-')],
            [sg.Text('単輪試験機駆動用モータの消費電力をCSV出力',font=('UD デジタル 教科書体 N-B',14))]]

    window = sg.Window('DataLogger_2Sensor_V6.exe', layout, size=(430,160))

    while True:
        event,values = window.read()
        if event == sg.WIN_CLOSED:
            break

        if event == '保存先':
            SAVE_PATH = sg.popup_get_folder('データの保存先を選択',font=('Yu Gothic UI Semibold',14))

        
        if event == 'START':
            window['-DISPLAY-'].update('START')
            #---Arduinoが接続されているCOMポートの指定---
            COM = serial.Serial(str(values['-PORTNAME-']),115200)           # Arduino に合わせた設定にする
            # ---Logger start---
            filename = str(SAVE_PATH) + '\\' +str(values['-FILENAME-']) + '_DL.csv'
            # ---CSVへの書き込み---
            with open(filename,'a', newline="") as f:
                writer = csv.writer(f)
                writer.writerow(["Time[sec]","VOL_1[V]","CUR_2[A]","POWER_1[W]"])      # １行目の見出し
                # ---情報の書き込み---
                delta_T = 0.0               # 微小時間変数
                TIME_COUNT = 0.0            # 実際に計測した時間
                # ---↓ループ前にシリアル通信の開始をしておかないと計測時間の開始にラグが出る（約1.5秒くらいラグが発生する）↓---
                VOL_1 = float(COM.readline().decode('utf-8').rstrip('\n'))  # シリアル通信でArduinoの電圧値を取得
                CUR_2 = float(COM.readline().decode('utf-8').rstrip('\n'))  # シリアル通信でArduinoの電流値を取得
                
                while(1):
                    TIME_START = time.time()                # データ取得時間計測開始
                    VOL_1 = float(COM.readline().decode('utf-8').rstrip('\n'))  # シリアル通信でArduinoの電圧値を取得
                    CUR_2 = float(COM.readline().decode('utf-8').rstrip('\n'))  # シリアル通信でArduinoの電流値を取得
                    # ---電力計算---
                    POWER_1 = VOL_1 * CUR_2                                     # 電力計算
                    # ---出力配列---
                    TIME_COUNT = TIME_COUNT + delta_T                           # 実際の計測時間計算
                    data = [TIME_COUNT,VOL_1,CUR_2,POWER_1]                     # CSV書き込み用データ配列
                    # ---データの書き込み---
                    writer.writerow(data)                                       # CSVにデータ書き込み
                    TIME_END = time.time()                  # データ取得時間計測終了
                    delta_T = TIME_END - TIME_START         # 微小時間計測
                    event,values = window.read(timeout=0)           # timeout = 0 がないと常に入力待機状態になる
                    if(event == 'STOP'):
                        break
                window['-DISPLAY-'].update('STOP')
                # ---logger stop---
                # ---ポートを閉じる---
                COM.close()
                f.close()
            
    window.close()


if __name__ == "__main__":
    main()