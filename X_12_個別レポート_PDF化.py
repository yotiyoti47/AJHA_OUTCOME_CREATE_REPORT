
import os
from tkinter import filedialog
import win32com.client
import X_00_CONST as CONST
#from PyPDF2 import PdfMerger
from pypdf import PdfWriter, PdfReader
import tkinter as tk
import time

#def create_report():
#    print("\n個別レポートのPDF化を開始します。")

#    print("\n個別レポート（.xlsm形式）が格納されたフォルダを指定してください。")
#    folder_path = filedialog.askdirectory(initialdir= CONST.OUTPUT_FOLDER)
#    if not folder_path:
#        print("\nエラー: フォルダを指定してください。")
#        return -1

#    print("\nPDF化  を開始します。")

    # フォルダ内のすべての.xlsmファイルを取得
#    xlsm_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsm')]
#    xlsm_files.sort()

    # xlsx形式のファイルは全て削除
#    for file in os.listdir(folder_path):
#        if file.endswith(".xlsx"):
#            os.remove(os.path.join(folder_path, file))

#    cnt = 1
#    for xlsm_file in xlsm_files:
#        if xlsm_file.startswith("~$"):
#            continue
#        print("  No.{} {} PDF化  開始".format(cnt, xlsm_file))
        
        # ファイル名と同名のフォルダを作成
#        temp_folder_path = os.path.join(folder_path, xlsm_file.replace(".xlsm", ""))
#        os.makedirs(temp_folder_path, exist_ok=True)        
        
        # フォルダ内のファイルを全て削除
#        for file in os.listdir(temp_folder_path):
#            os.remove(os.path.join(temp_folder_path, file))

#        # Excelを起動
#        excel = win32com.client.Dispatch("Excel.Application")
#        excel.Visible = False  # Falseにするとバックグラウンドで実行される
#        excel.DisplayAlerts = False  
#        # ファイルを開く
#        wb = excel.Workbooks.Open(os.path.join(folder_path, xlsm_file))
        # xlsx形式で保存
#        wb.SaveAs(os.path.join(folder_path, xlsm_file.replace(".xlsm", ".xlsx")), FileFormat=51)

#        # シート名一覧を取得
#        sheet_names = [sheet.Name for sheet in wb.Sheets]
#        for sheet_name in sheet_names:

#            if "CI_" in sheet_name:
#                continue
#                
#            # シートをPDF化
#            pre_text = ""
#            if sheet_name == "レポート_表紙":
#               pre_text = "00_1_"
#            elif sheet_name == "note":
#                pre_text = "00_2_"

#            ws = wb.Sheets(sheet_name)
#            ws.ExportAsFixedFormat(0, os.path.join(temp_folder_path, pre_text +sheet_name + ".pdf"))
#        
#        cnt += 1
#        # Excelを終了
#        wb.Close()
#        excel.Quit()

#        # temp_folder_path内のPDFファイルを取得
#        pdf_files = [f for f in os.listdir(temp_folder_path) if f.endswith(".pdf")]
#        pdf_files.sort()

#        merger = PdfMerger()
#        for pdf_file in pdf_files:
#            full_path = os.path.join(temp_folder_path, pdf_file)
#            bookmark_name = os.path.splitext(pdf_file)[0]  # 拡張子なしのファイル名
#            merger.append(full_path, outline_item=bookmark_name)
        # 結合ファイルを保存
#        with open(os.path.join(folder_path, xlsm_file.replace(".xlsm", ".pdf")), "wb") as f_out:
#            merger.write(f_out)

#        print("  No.{} {} PDF化  終了".format(cnt, xlsm_file))

#    print("\n個別レポートのPDF化を完了しました。")
#    return 0


def create_report_pypdf():
    print("\n個別レポートのPDF化を開始します。")

    print("\n個別レポート（.xlsm形式）が格納されたフォルダを指定してください。")
    
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    
    folder_path = filedialog.askdirectory(initialdir=CONST.OUTPUT_FOLDER)
    if not folder_path:
        print("\nエラー: フォルダを指定してください。")
        return -1

    print("\nPDF化を開始します。")

    # .xlsmファイル取得・ソート
    xlsm_files = sorted(f for f in os.listdir(folder_path) if f.endswith('.xlsm'))

    # 古い.xlsxを全削除
    for file in os.listdir(folder_path):
        if file.endswith(".xlsx"):
            os.remove(os.path.join(folder_path, file))

    cnt = 1
    for xlsm_file in xlsm_files:
        if xlsm_file.startswith("~$"):
            continue

        print(f"  No.{cnt} {xlsm_file} PDF化 開始")

        temp_folder_path = os.path.join(folder_path, xlsm_file.replace(".xlsm", ""))
        os.makedirs(temp_folder_path, exist_ok=True)

        for file in os.listdir(temp_folder_path):
            os.remove(os.path.join(temp_folder_path, file))

        excel = None
        wb = None
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            # プリンターアクセス抑制
            excel.PrintCommunication = False  

            wb_path = os.path.join(folder_path, xlsm_file)
            print("  " + wb_path + "を開きます")
            wb = excel.Workbooks.Open(wb_path)
            print("wb_path " + wb_path)
            # 保存先パスを厳密に作成（拡張子の大文字小文字差異にも対応）
            xlsx_path = os.path.normpath(os.path.splitext(wb_path)[0] + ".xlsx")
            print("xlsx_path " + xlsx_path)
            # 同名のディレクトリが存在する場合は保存不可
            if os.path.isdir(xlsx_path):
                print("エラー: 同名のフォルダが存在するため、保存できません: " + xlsx_path)
                wb.Close(SaveChanges=False)
                excel.Quit()
                return -1
            # 既存の同名ファイルがあれば削除（Excel の上書き衝突回避）
            if os.path.exists(xlsx_path):
                try:
                    os.remove(xlsx_path)
                except Exception as e:
                    print("警告: 既存ファイルの削除に失敗しました: " + str(e))
            wb.SaveAs(xlsx_path, FileFormat=51)

            sheet_names = [sheet.Name for sheet in wb.Sheets]

            for sheet_name in sheet_names:
                if "CI_" in sheet_name:
                    continue

                pre_text = ""
                if sheet_name == "レポート_表紙":
                    pre_text = "00_1_"
                elif sheet_name == "note":
                    pre_text = "00_2_"

                ws = wb.Sheets(sheet_name)
                # 逆順で走査して削除の副作用を回避
                try:
                    chart_count = ws.ChartObjects().Count
                except Exception:
                    chart_count = 0
                for idx in range(chart_count, 0, -1):
                    chart_obj = ws.ChartObjects(idx)
                    try:
                        # グラフを画像としてコピー（クリップボードへ）
                        chart_obj.CopyPicture(Appearance=1, Format=2)  # Format=2: Picture

                        # 元の位置とサイズを保存
                        left = chart_obj.Left
                        top = chart_obj.Top
                        width = chart_obj.Width
                        height = chart_obj.Height

                        # クリップボードから貼り付け（画像として）
                        ws.Paste()
                        # 直後はCOM側で非同期のため、直近のShape取得をリトライ
                        img = None
                        for retry in range(10):
                            try:
                                if ws.Shapes.Count > 0:
                                    img = ws.Shapes(ws.Shapes.Count)
                                    break
                            except Exception as e:
                                print(f"    Shape取得リトライ {retry+1}/10: {e}")
                                time.sleep(0.1)
                        
                        if img is None:
                            print("警告: 画像貼り付けに失敗しました（シート: " + ws.Name + ")")
                            # 元のグラフは削除せずスキップ
                            continue
                        
                        # 位置・サイズ設定をリトライ
                        for retry in range(3):
                            try:
                                img.Left = left
                                img.Top = top
                                img.Width = width
                                img.Height = height
                                break
                            except Exception as e:
                                print(f"    位置設定リトライ {retry+1}/3: {e}")
                                time.sleep(0.05)

                        # 元のグラフを削除
                        chart_obj.Delete()

                        # 反映待ち
                        time.sleep(0.1)
                        
                    except Exception as e:
                        print(f"警告: グラフ処理中にエラー（シート: {ws.Name}, グラフ: {idx}）: {e}")
                        # このグラフはスキップして次へ
                        continue


                print("  " + ws.Name + "エクスポート開始")
                export_path = os.path.join(temp_folder_path, pre_text + sheet_name + ".pdf")
                ws.ExportAsFixedFormat(0, export_path)
                print("  " + ws.Name + "エクスポート終了")
            wb.Close(SaveChanges=True)
            excel.Quit()
        except Exception as e:
            print("エラー: Excel処理中に例外が発生しました: " + str(e))
            # 後片付け（エクセル確実終了）
            try:
                if wb is not None:
                    wb.Close(SaveChanges=False)
            except Exception:
                pass
            try:
                if excel is not None:
                    excel.Quit()
            except Exception:
                pass
            return -1

        pdf_files = sorted(f for f in os.listdir(temp_folder_path) if f.endswith(".pdf"))
        # 表紙優先→00_2_→残りは名前順
        pdf_files.sort(key=lambda s: (0 if s.startswith("00_1_レポート_表紙") else 1 if s.startswith("00_2_note") else 2, s))
        writer = PdfWriter()
        current_page = 0

        for pdf_file in pdf_files:
            full_path = os.path.join(temp_folder_path, pdf_file)
            reader = PdfReader(full_path)
            bookmark_name = os.path.splitext(pdf_file)[0]

            for page in reader.pages:
                writer.add_page(page)

            # ブックマーク追加（最初のページを指定）
            writer.add_outline_item(bookmark_name, current_page)
            current_page += len(reader.pages)

        output_pdf_path = os.path.join(folder_path, xlsm_file.replace(".xlsm", ".pdf"))
        with open(output_pdf_path, "wb") as f_out:
            writer.write(f_out)

        print(f"  No.{cnt} {xlsm_file} PDF化 終了")
        cnt += 1

    print("\n個別レポートのPDF化を完了しました。")
    return 0


#def create_report_pypdf_old():
    print("\n個別レポートのPDF化を開始します。")

    print("\n個別レポート（.xlsm形式）が格納されたフォルダを指定してください。")
    
    root = tk.Tk()
    root.withdraw()  # ルートウィンドウを非表示にする
    root.attributes('-topmost', True)  # ダイアログを最前面に表示
    
    folder_path = filedialog.askdirectory(initialdir=CONST.OUTPUT_FOLDER)
    if not folder_path:
        print("\nエラー: フォルダを指定してください。")
        return -1

    print("\nPDF化を開始します。")

    # .xlsmファイル取得・ソート
    xlsm_files = sorted(f for f in os.listdir(folder_path) if f.endswith('.xlsm'))

    # 古い.xlsxを全削除
    for file in os.listdir(folder_path):
        if file.endswith(".xlsx"):
            os.remove(os.path.join(folder_path, file))

    cnt = 1
    for xlsm_file in xlsm_files:
        if xlsm_file.startswith("~$"):
            continue

        print(f"  No.{cnt} {xlsm_file} PDF化 開始")



        temp_folder_path = os.path.join(folder_path, xlsm_file.replace(".xlsm", ""))
        os.makedirs(temp_folder_path, exist_ok=True)

        # tempフォルダ内削除
        for file in os.listdir(temp_folder_path):
            os.remove(os.path.join(temp_folder_path, file))

        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Open(os.path.join(folder_path, xlsm_file))
        wb.SaveAs(os.path.join(folder_path, xlsm_file.replace(".xlsm", ".xlsx")), FileFormat=51)

        sheet_names = [sheet.Name for sheet in wb.Sheets]

        for sheet_name in sheet_names:
            if "CI_" in sheet_name:
                continue

            pre_text = ""
            if sheet_name == "レポート_表紙":
                pre_text = "00_1_"
            elif sheet_name == "note":
                pre_text = "00_2_"

            ws = wb.Sheets(sheet_name)
            ws.ExportAsFixedFormat(0, os.path.join(temp_folder_path, pre_text + sheet_name + ".pdf"))

        wb.Close()
        excel.Quit()

        pdf_files = sorted(f for f in os.listdir(temp_folder_path) if f.endswith(".pdf"))
        merger = PdfMerger()

        for pdf_file in pdf_files:
            full_path = os.path.join(temp_folder_path, pdf_file)
            bookmark_name = os.path.splitext(pdf_file)[0]
            merger.append(full_path, outline_item=bookmark_name)

        output_pdf_path = os.path.join(folder_path, xlsm_file.replace(".xlsm", ".pdf"))
        with open(output_pdf_path, "wb") as f_out:
            merger.write(f_out)
        merger.close()

        print(f"  No.{cnt} {xlsm_file} PDF化 終了")
        cnt += 1

    print("\n個別レポートのPDF化を完了しました。")
    return 0