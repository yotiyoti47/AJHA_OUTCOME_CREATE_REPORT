import copy
import re
from datetime import datetime
import os
import sys
import X_00_CONST as CONST
import X_01_レポート作成 as X_01
import openpyxl
import shutil

import X_IN_05_死亡率_疾患別_重症度別 as IN_05
import X_IN_06_死亡率_疾患別_年代別 as IN_06
import X_IN_07_死亡率_疾患別_性別 as IN_07
import X_IN_12_医療費_疾患別_重症度別 as IN_12
import X_IN_13_医療費_疾患別_年代別 as IN_13
import X_IN_14_医療費_疾患別_性別 as IN_14

#出力フォルダを設定する ※各レポートで個別実装
#引数：target_year：対象年度、target_group：対象グループ、target_period：対象期間
#戻り値：出力先フォルダのパス
#ERROR:-1
def set_output_folder(target_year):
    
    try:
        #出力先に当該年度のフォルダがあるか確認する
        output_folder = CONST.OUTPUT_FOLDER
        output_folder = os.path.join(output_folder, str(target_year))
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        #output_folder内に「05_個別レポート」フォルダがあるか確認する
        individual_folder = os.path.join(output_folder, "05_個別レポート")
        if not os.path.exists(individual_folder):
            os.makedirs(individual_folder)
    
        #acute_folder内に当該グループのフォルダがあるか確認する
        acute_folder = os.path.join(individual_folder, "04_死亡率_医療費")
        if not os.path.exists(acute_folder):
            os.makedirs(acute_folder)
    except:
        print("エラー: 出力先フォルダの設定に失敗しました。年度={}".format(target_year))
        return -1
      
    #report_type_folder内のファイル、フォルダを全て削除
    delete_file_count = 0
    for file in os.listdir(acute_folder):
        try:
            os.remove(os.path.join(acute_folder, file))
            #削除したファイル数をカウント
            delete_file_count += 1
        except:
            print("フォルダ：{}".format(acute_folder))
            print(f"エラー: {file} の削除に失敗しました。")
            print("エラー内容：{}".format(sys.exc_info()[1]))
            return -1

    print("削除したファイル数：{}".format(delete_file_count))
    return acute_folder


def copyFormatFile(target_year, output_folder):
    print("\nフォーマットファイルをコピーします。")

    format_file_path = CONST.INDIV_REP_FRMT_ACUTE_DEATHRATE_COST_PATH

    if not os.path.exists(format_file_path):
        print("エラー: フォーマットファイルが存在しません。")
        print("フォーマットファイルのパス:{}".format(format_file_path))
        return -1
    try:
        #format_file_pathのファイル名を「経年比較レポート_医療費、死亡率_1_急性期グループ.xlsm」に変更
        newName1 = "経年比較レポート_医療費、死亡率_1_急性期グループ.xlsm"
        shutil.copy(format_file_path, os.path.join(output_folder, newName1))

        newName2 = "経年比較レポート_医療費、死亡率_2_慢性期グループ.xlsm"
        shutil.copy(format_file_path, os.path.join(output_folder, newName2))

    except:
        print("エラー: フォーマットファイルのコピーに失敗しました。")
        print("エラー内容：{}".format(sys.exc_info()[1]))

        return -1

    print("フォーマットファイルのコピーに成功しました。")
    return 0


def create_indiv_report(target_create_date, output_folder,is_acute):
    print("\n個別レポートを作成します。{}".format(is_acute))

    #output_folder内のxlsmファイルを取得
    xlsm_file_list = [f for f in os.listdir(output_folder) if f.endswith('.xlsm')]
    xlsm_file_list.sort()

    if len(xlsm_file_list) == 0:
        #存在しない場合もあるので正常終了で返す
        print("エラー: xlsmファイルが存在しません。")
        return 0

    for i in range(2):
        xlsm_file = xlsm_file_list[i]

        if is_acute not in xlsm_file:
            continue

        wb = openpyxl.load_workbook(os.path.join(output_folder, xlsm_file), keep_vba=True)
        
        IN_05.getRepAgeData(wb, is_acute)
        IN_06.getRepAgeData(wb, is_acute)
        IN_07.getRepAgeData(wb, is_acute)
        IN_12.getRepAgeData(wb, is_acute)
        IN_13.getRepAgeData(wb, is_acute)
        IN_14.getRepAgeData(wb, is_acute)
        
        wb.save(os.path.join(output_folder, xlsm_file))
        wb.close()


def create_report():
    print("\n経年比較 死亡率・医療費レポートを作成します。")

    print("\n年度（西暦4桁、半角）を入力してください。")
    user_input = input(":>>")
    if not user_input.isdigit() or len(user_input) != 4:
        print("\nエラー: 西暦4桁の半角数字を入力してください。")
        return -1

    target_year = int(user_input)

    print("\nレポートの作成日をyyyy/m/d形式で入力してください。")
    user_input = input(":>>")
    if not re.match(r'^\d{4}/\d{1,2}/\d{1,2}$', user_input):
        print("\nエラー: 日付の形式が不正です。yyyy/m/d形式で入力してください。")
        return -1

    target_create_date = datetime.strptime(user_input, "%Y/%m/%d")
    print(f"\n対象作成日: {target_create_date.strftime('%Y/%m/%d')}\n")

    #出力先フォルダの設定
    output_folder = set_output_folder(target_year)
    if output_folder is None or output_folder == -1:
        return -1


    if copyFormatFile(target_year, output_folder) == -1:   
        return -1

    # マスタデータをセット
    X_01.setMastaData()

    create_indiv_report(target_create_date, output_folder,"急性期")
    create_indiv_report(target_create_date, output_folder,"慢性期")


    return 0