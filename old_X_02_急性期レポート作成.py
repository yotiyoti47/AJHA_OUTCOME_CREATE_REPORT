import X_00_CONST as CONST
import os
import sys
import shutil

TEMP_FOLDER = ""

# 急性期レポートを作成する
def create_report(target_year):
    
    TEMP_FOLDER = os.path.join(CONST.OUTPUT_FOLDER, str(target_year), CONST.INDCTR_ACUTE,"TEMP")

    #最初に「病院名レポート」を作成する    
    
    groups = ["急性期グループ", "慢性期グループ",]
    perids = ["1Q", "2Q", "3Q", "4Q","年間"]

    for group in groups:
        for period in perids:
            copyReport(target_year,group)

                


def copyReport(target_year, group, period):

    TEMP_FOLDER = os.path.join(CONST.OUTPUT_FOLDER, str(target_year), CONST.INDCTR_ACUTE,"TEMP")

    print("ニッセイが納品した急性期レポートのフォルダパスを入力してください: ")
    #source_report_folder = input(">>")

    #
    #DEBUG
    #
    print("DEBUG: source_report_folderを設定します。")    
    #2024
    #source_report_folder = r"\\Ajha10\共有フォルダ\全日病事務局\医療の質向上委員会\1_医療の質向上委員会\2024年度\10_委員会事業\4_診療アウトカム評価事業\11_NIT納品\2024年度診療アウトカム評価事業3Q\急性期指標"
    #2023
    source_report_folder = r"\\ajha10\共有フォルダ\全日病事務局\医療の質向上委員会\1_医療の質向上委員会\2023年度\10_委員会事業\02_診療アウトカム評価事業\08_NIT納品\2023年度診療アウトカム評価事業1-4Q,年間\02.急性期指標"
    dict_report_file_names = {}
    if target_year <= 2023:
        dict_report_file_names = CONST.DICT_REPORT_FILE_NAMES_ACUTE_TO_2024_1Q
        #print("TO_2024_1Q " + hosopNamePath)
    elif target_year == 2024:
        dict_report_file_names = CONST.DICT_REPORT_FILE_NAMES_ACUTE_FROM_2024_2Q

        #print("FROM_2024_2Q " + hosopNamePath)
    elif target_year > 2024:
        dict_report_file_names = CONST.DICT_REPORT_FILE_NAMES_ACUTE_FROM_2024_2Q

        #print("FROM_2024_2Q " + hosopNamePath)

    if not source_report_folder:
        print("\nエラー: フォルダパスが入力されていません。")
        return      
    if not os.path.isdir(source_report_folder):
        print(f"\nエラー: フォルダが存在しません: {source_report_folder}")
        return      
    
    excel_folder = os.path.join(source_report_folder, "帳票_エクセル")
    if not os.path.isdir(excel_folder) :
        excel_folder = os.path.join(source_report_folder, "帳票_Excel")
        if not os.path.isdir(excel_folder):
            excel_folder = os.path.join(source_report_folder, "帳票_EXCEL")
            if not os.path.isdir(excel_folder):
                excel_folder = os.path.join(source_report_folder, "帳票_excel")

    pdf_folder = os.path.join(source_report_folder, "帳票_PDF")
    if not os.path.isdir(excel_folder) or not os.path.isdir(pdf_folder):
        print("\nエラー: 帳票_Excel または 帳票_PDF フォルダが存在しません。")
        print("excel_folder={}, pdf_folder={}".format(excel_folder, pdf_folder))
        return

    #PDFフォルダ
    group_period_folder = os.path.join(pdf_folder, group, period)
    print("group_period_folder = {}".format(group_period_folder))

    if os.path.isdir(group_period_folder):
        copyCnt = 0
        for file in os.listdir(group_period_folder):
            #print(f"ファイル名: {file}")
            if(target_year== 2024 and period == "1Q") :
                dict_report_file_names = CONST.DICT_REPORT_FILE_NAMES_ACUTE_TO_2024_1Q
                    
            if file.endswith(".pdf") :
                #ニッセイが納品したレポートをTEMPフォルダにコピーする
                src_file = os.path.join(group_period_folder, file)

                file_head2 = file[:2]
                if file_head2 in dict_report_file_names:
                    new_file_name = dict_report_file_names[file_head2]
                else:
                    print(f"警告: ファイル名の先頭2文字が辞書に存在しません: {file}")
                    sys.exit(0)
                dst_file = os.path.join(TEMP_FOLDER, new_file_name)

                os.makedirs(TEMP_FOLDER, exist_ok=True)
                try:
                    print(f"ファイルをコピー中: {file} -> {new_file_name}")
                    #shutil.copy2(src_file, dst_file)
                    copyCnt += 1
                except Exception as e:
                    print(f"ファイルのコピー中にエラーが発生しました: {src_file} -> {dst_file}\n{e}")
                    sys.exit(0)
                
        print(f"{target_year} 年度 {group}  {period} レポートを[TEMP]フォルダに{copyCnt}件コピーしました。")


def copyTEIGI(target_year, group, period):
    TEMP_FOLDER = os.path.join(CONST.OUTPUT_FOLDER, str(target_year), CONST.INDCTR_ACUTE,"TEMP")

    print("ニッセイが納品した急性期指標の定義ファイルのフォルダパスを入力してください: ")
    #source_teigi_folder = input(">>")

    