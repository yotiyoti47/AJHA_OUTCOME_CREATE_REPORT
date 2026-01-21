# 診療アウトカム評価事業におけるHP掲載レポート（病院名を病院番号に変更）および参加病院フィードバック用レポートを作成する
# 作成者: 全日病事務局 吉田　喬
# 作成日: 2025/6/24
# 更新日: 2025/6/24
#
# ①急性期指標、②アンケート調査、③慢性期指標

import os
import sys
import shutil
from datetime import datetime
import X_00_CONST as CONST
import X_02_急性期レポート作成

if __name__ == "__main__":
    print("\n**********レポート作成プログラム**********\n")
    print("対象年度を入力してください。")
    user_input = input(">>")
    if not (user_input.isdigit() and len(user_input) == 4):
        print("エラー: 対象年度は4桁の数値（半角）で入力してください。")
        sys.exit(0)
    print(f"対象年度: {user_input}")

    target_year = int(user_input)

    # OUTPUT_FOLDER直下に名称「target_year」のフォルダがあるか確認
    target_year_folder = os.path.join(CONST.OUTPUT_FOLDER, str(target_year))
    if not os.path.isdir(target_year_folder):
        os.makedirs(target_year_folder)
        print(f"{target_year_folder} を作成しました。")

    print("対象指標を選んでください。\n")
    print("１：急性期指標、２：アンケート調査、３：慢性期指標。\n")
    
    user_input = input(":>>")
    if user_input not in ["1", "2", "3"]:
        print("エラー: 1, 2, 3（半角） のいずれかを入力してください。")
        sys.exit(0)
    
    target_indicator =""
    
    if user_input == "1":
        target_indicator = CONST.INDCTR_ACUTE   
    elif user_input == "2":
        target_indicator = CONST.INDCTR_SURVEY    
    elif user_input == "3":
        target_indicator = CONST.INDCTR_CHRONIC
    print(f"対象指標: {target_indicator}\n")

    # target_year_folder直下にtarget_indicatorのフォルダがあるか確認
    target_indicator_folder = os.path.join(target_year_folder, target_indicator)
    if not os.path.isdir(target_indicator_folder):
        os.makedirs(target_indicator_folder)
        print(f"{target_indicator_folder} を作成しました。")

    # target_indicator_folder内にpdfファイルがないか確認する
    pdf_files = [f for f in os.listdir(target_indicator_folder) if f.lower().endswith('.pdf')]
    if not pdf_files:
        print(f"{target_indicator_folder} 内にPDFファイルはありません。")
    else:
        # 本日の日時でフォルダ名を作成
        now_str = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_folder = os.path.join(target_indicator_folder, now_str)
        os.makedirs(backup_folder)

        # PDFファイルを移動
        for pdf_file in pdf_files:
            src = os.path.join(target_indicator_folder, pdf_file)
            dst = os.path.join(backup_folder, pdf_file)
            shutil.move(src, dst)

        print(f"既存のPDFファイルを {backup_folder} に移動しました。")

    # target_indicator_folder内に「TEMP」フォルダがなければ作成する
    temp_folder = os.path.join(target_indicator_folder, "TEMP")
    if not os.path.isdir(temp_folder):
        os.makedirs(temp_folder)
        print(f"{temp_folder} を作成しました。")


    print("グループを選んでください。\n")
    print("１：急性期グループ指標、２：慢性期グループ\n")
    
    user_input = input(":>>")
    if user_input not in ["1", "2", ]:
        print("エラー: 1, 2,（半角） のいずれかを入力してください。")
        sys.exit(0)

    target_group = ""
    if user_input == "1":
        target_group = CONST.GROUPS[0]  # 急性期グループ   
    elif user_input == "2":
        target_group = CONST.GROUPS[1]  # 慢性期グループ    
    print(f"対象グループ: {target_group}\n")

    print("期間を選んでください。\n")
    print("１：1Q、２：2Q、３：3Q、４：4Q、５：年間\n")

    
    user_input = input(":>>")
    if user_input not in ["1", "2", "3", "4", "5"  ]:
        print("エラー: 1, 2,3,4,5（半角） のいずれかを入力してください。")
        sys.exit(0)

    target_period = ""
    if user_input == "1":
        target_period = CONST.PERIODS[0]  # 1Q  
    elif user_input == "2":
        target_period = CONST.PERIODS[1]  # 2Q        
    elif user_input == "3":
        target_period = CONST.PERIODS[2]  # 3Q        
    elif user_input == "4":
        target_period = CONST.PERIODS[3]  # 4Q        
    elif user_input == "5":
        target_period = CONST.PERIODS[4]  # 年間
    print(f"対象期間: {target_period}\n")



    # target_year_folder直下にtarget_indicatorのフォルダがあるか確認
    target_indicator_folder = os.path.join(target_year_folder, target_indicator)
    if not os.path.isdir(target_indicator_folder):
        os.makedirs(target_indicator_folder)
        print(f"{target_indicator_folder} を作成しました。")




    if target_indicator == CONST.INDCTR_ACUTE:
        # 急性期指標のレポートを作成
        X_02_急性期レポート作成.create_report(target_year)


