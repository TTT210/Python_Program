# 作成日：2022.08.09 V1
# 変更日：2022.11.02 V2 MCのContinueがなくなったのでコメントアウト
# =====================================================================================================
# 作成日：2022.11.08 MCFS_DL_CAM_V1 モーションキャプチャ位置の画像認識，Position.exeの廃止
#                                   mcfs_position_cameraV2をベースとしています
#                                   生成されたEXEファイルと同じ階層に写真データを入れること．
#                                   EXEファイルはショートカットで使うこと（でないと画像のパスが反映されない）
# =====================================================================================================
# 参考URL
# PySimpleGui : https://fuji-pocketbook.net/pysimplegui/
# Pyautogui : https://tech-blog.rakus.co.jp/entry/20220613/python
# pyinstaller(仮想環境) : https://teshi-learn.com/2020-12/windows-pyinstaller-python-exe-anaconda/
# ---コンパイル方法----
# 仮想環境の確認 -> conda info -e (この中のenv_exeがコンパイル用仮想環境)
# 仮想環境のアクティベート -> activate env_exe
# pyinstallerの実行 -> pyinstaller *****.py --onefile --noconsole
#   （--onefile にすると一つのexeファイルとして生成される）
#   （--noconsole を追加するとexeファイル起動後に黒い画面が生成されない）
# ---
# =====================================================================================================



import PySimpleGUI as sg
import pyautogui as pg


def main():
    layout = [
            [sg.Text('ファイル名'),sg.Input(key = '-FILENAME-')],
            [sg.Button('START'),sg.Button('STOP')],
            [sg.Text('',key = '-DISPLAY-')],
            [sg.Text('単輪試験機自動データ取得ソフト',font=('UD デジタル 教科書体 N-B',12))]]

    window = sg.Window('MCFS_DL_CAM_V1.exe',layout)

    while True:
        event,values = window.read()
        if event == sg.WIN_CLOSED:
            break
        if event == 'START':
            window['-DISPLAY-'].update('START')
            MC_REC = pg.locateOnScreen("./img/MC_REC.png")    #MC RECボタンの画像認識
            REC_ON(values['-FILENAME-'],MC_REC)
        elif event == 'STOP':
            window['-DISPLAY-'].update('STOP')
            MC_REC = pg.locateOnScreen("./img/MC_REC.png")    #MC RECボタンの画像認識
            REC_OFF(MC_REC)
    window.close()

def REC_ON(REC_KEY,MC_REC_POSITION):             # 録画開始
    pg.moveTo(330,1010)          # MCファイル名エリアへ移動
    pg.click(button="left", clicks=2)   # MCファイル名ダブルクリック（これならうまくいった）
    pg.hotkey('ctrl','a')        # MCファイル名全選択
    pg.write(str(REC_KEY),interval=0.05)       # MCファイル名入力(インターバルがないと文字を入力する前に次のエリアに進んでしまう)
    pg.click(1700,630)           # FC値リセットクリック
    pg.click(1870,400)           # FSファイル名エリアクリック
    pg.hotkey('ctrl','a')        # FSファイル名全選択
    pg.write(str(REC_KEY),interval=0.05)       # FSファイル名入力(インターバルがないと文字を入力する前に次のエリアに進んでしまう)
    pg.click(1730,920)           # Dataloggerファイル名エリアクリック
    pg.hotkey('ctrl','a')        # FSファイル名全選択
    pg.write(str(REC_KEY),interval=0.05)       # データロガーファイル名入力(インターバルがないと文字を入力する前に次のエリアに進んでしまう)
    pg.click(1490,945)           # Datalogger START_ON
    pg.click(MC_REC_POSITION,duration=1)           # MC REC ON/OFF Dataloggerの通信開始までのタイムラグを1秒設ける
    pg.click(1815,430)           # FS REC ON/OFF
    pg.click(3790,530)           # WebCamera ON/OFF
    pg.moveTo(1250,750)          # STOPボタンの位置まで戻る

def REC_OFF(MC_REC_POSITION):                   # 録画停止とCSVファイルへの出力
    pg.click(1550,945)                          # Datalogger STOP（先にデータロガーを止める）
    pg.click(MC_REC_POSITION)                            # MC REC ON/OFF
    pg.click(1815,430)                          # FS REC ON/OFF
    pg.click(3790,530)                          # WebCamera ON/OFF
    pg.click(95,1010,duration=1)            # MC EDIT
    pg.click(18,42,duration=1)              # MC FILE
    pg.click(95,290,duration=1)             # MC Export Tracking data
    pg.click(880,510,duration=2)            # MC Export
    # pg.click(725,580,duration=2)          # MC Continue
    pg.click(45,1010,duration=4)            # MC Live
    pg.click(1400,720)                     # ファイル名入力欄の位置まで戻る


if __name__ == "__main__":
    main()