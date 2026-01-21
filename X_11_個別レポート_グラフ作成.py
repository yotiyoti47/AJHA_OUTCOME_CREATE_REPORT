import os
from tkinter import filedialog
import win32com.client
import X_00_CONST as CONST
import time
import threading
import tkinter as tk
import win32com.client as win32

# ログ監視フラグ
log_done = False

def monitor_log(log_path):
    last_line = ""
    while not log_done:
        try:
            with open(log_path, "r", encoding="utf-8") as f:
                lines = f.readlines()
                if lines and lines[-1] != last_line:
                    last_line = lines[-1]
                    print("    ログ:", last_line.strip())
        except Exception:
            pass
        time.sleep(1)


def create_report():
    global log_done

    print("\n個別レポートのグラフ作成を開始します。")

    try:
        print("\n個別レポート（.xlsm形式）が格納されたフォルダを指定してください。")
       
        root = tk.Tk()
        root.withdraw()  # ルートウィンドウを非表示にする
        root.attributes('-topmost', True)  # ダイアログを最前面に表示

        folder_path = filedialog.askdirectory(initialdir=CONST.OUTPUT_FOLDER)
        if not folder_path:
            print("\nエラー: フォルダを指定してください。")
            return -1

        print("\nグラフ作成を開始します。")

        xlsm_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsm')]
        cnt = 1
        for xlsm_file in xlsm_files:
            # xlsm_fileのファイル名に~$が含まれていたら飛ばす
            if "~$" in xlsm_file:
                continue

            if os.path.isdir(os.path.join(folder_path, xlsm_file)):
                continue

            print("  No.{} {} グラフ作成  開始".format(cnt, xlsm_file))

            # ログファイルパス（ファイルごとに同名で出力される想定）
            log_path = os.path.join(folder_path, "indiv_Graph_log.txt")

            # ログ監視スレッド開始
            log_done = False
            log_thread = threading.Thread(target=monitor_log, args=(log_path,))
            log_thread.start()

            # Excel起動とマクロ実行
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            wb = excel.Workbooks.Open(os.path.join(folder_path, xlsm_file))
            ws = wb.Worksheets["レポート_表紙"]

            # セルCF6の値が空であればグラフ作成を実行
            value = ws.Range("CF6").Value
            #print("ws.Range(CF6).Value " + value)
            if value == None:
                excel.Application.Run("経年レポート作成")
                ws.Range("CF6").Value = "グラフ作成完了"
            
            wb.Close(SaveChanges=True)
            excel.Quit()

            # ログ監視終了
            log_done = True
            log_thread.join()

            print("  No.{} {} グラフ作成  完了".format(cnt, xlsm_file))
            cnt += 1

        print("\nグラフ作成を完了しました。")
        return 0

    except Exception as e:
        print("エラー:", e)
        return -1


def create_report_old():
    print("\n個別レポートのグラフ作成を開始します。")

    try:
        print("\n個別レポート（.xlsm形式）が格納されたフォルダを指定してください。")
        folder_path = filedialog.askdirectory(initialdir= CONST.OUTPUT_FOLDER)
        if not folder_path:
            print("\nエラー: フォルダを指定してください。")
            return -1

        print("\nグラフ作成を開始します。")

        # フォルダ内のすべての.xlsmファイルを取得
        xlsm_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsm')]
        cnt = 1
        for xlsm_file in xlsm_files:
            # もしxlsm_fileがフォルダであれば飛ばす
            if os.path.isdir(os.path.join(folder_path, xlsm_file)):
                continue

            print("  No.{} {} グラフ作成  開始".format(cnt, xlsm_file))
            # Excelを起動
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False  # Falseにするとバックグラウンドで実行される
            wb = excel.Workbooks.Open(os.path.join(folder_path, xlsm_file))

            ws = wb.Worksheets["レポート_表紙"]

            #セルCF6の値が空であればグラフ作成を実行
            if ws["CF6"].Value == "":
                excel.Application.Run("経年レポート作成")
                ws["CF6"].Value = "グラフ作成完了"
            
            wb.Close(SaveChanges=True) 
            print("  No.{} {} グラフ作成  完了".format(cnt, xlsm_file))
            cnt += 1

        excel.Quit()

        print("\nグラフ作成を完了しました。")
        return 0
    except Exception as e:
        print(e)
        return -1


