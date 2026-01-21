import copy
import re
from datetime import datetime
import os
import sys
import X_00_CONST as CONST
import X_01_レポート作成 as X_01
import openpyxl
import shutil

import X_IN_69_1_ひと月あたり症例数_回復リハ as IN_69_1
import X_IN_69_2_ひと月あたり症例数_地域包括ケア as IN_69_2
import X_IN_69_3_ひと月あたり症例数_療養 as IN_69_3
import X_IN_70_1_ひと月あたり延べ在院日数_回復リハ as IN_70_1
import X_IN_70_2_ひと月あたり延べ在院日数_地域包括ケア as IN_70_2
import X_IN_70_3_ひと月あたり延べ在院日数_療養 as IN_70_3
import X_IN_71_1_平均在院日数_回復リハ as IN_71_1
import X_IN_71_2_平均在院日数_地域包括ケア as IN_71_2
import X_IN_71_3_平均在院日数_療養 as IN_71_3
import X_IN_72_1_ADLスコアの改善率_回復リハ as IN_72_1
import X_IN_72_2_ADLスコアの改善率_地域包括ケア as IN_72_2
import X_IN_72_3_ADLスコアの改善率_療養 as IN_72_3
import X_IN_73_FIM得点の改善率_回復リハ as IN_73
import X_IN_74_医療区分の改善率_療養 as IN_74
import X_IN_75_1_疾患別リハ単位数_運動器_回復リハ as IN_75_1
import X_IN_75_2_疾患別リハ単位数_呼吸器_回復リハ as IN_75_2
import X_IN_75_3_疾患別リハ単位数_心大血管疾患_回復リハ as IN_75_3
import X_IN_75_4_疾患別リハ単位数_脳血管疾患_回復リハ as IN_75_4
import X_IN_75_5_疾患別リハ単位数_廃用症候群_回復リハ as IN_75_5
import X_IN_76_1_紹介率_回復リハ as IN_76_1
import X_IN_76_2_紹介率_地域包括ケア as IN_76_2
import X_IN_76_3_紹介率_療養 as IN_76_3
import X_IN_77_1_在宅復帰率_回復リハ as IN_77_1
import X_IN_77_2_在宅復帰率_地域包括ケア as IN_77_2
import X_IN_77_3_在宅復帰率_療養 as IN_77_3
import X_IN_78_医療区分別の症例構成割合_療養 as IN_78
import X_IN_79_薬剤管理指導料の算定率_療養 as IN_79
import X_IN_80_退院時薬剤情報管理指導料の算定率_療養 as IN_80
import X_IN_81_1_目標設定等支援_管理料の算定率_回復リハ as IN_81_1
import X_IN_81_2_目標設定等支援_管理料の算定率_療養 as IN_81_2
import X_IN_82_退院時リハビリテーション指導料の算定率_療養  as IN_82
import X_IN_83_退院前訪問指導料の算定率_療養 as IN_83
import X_IN_84_1_要介護度_回復リハ as IN_84_1
import X_IN_84_2_要介護度_地域包括ケア as IN_84_2
import X_IN_84_3_要介護度_療養 as IN_84_3
import X_IN_85_1_1_要介護情報_胃瘻_腸瘻_回復リハ as IN_85_1_1
import X_IN_85_1_2_要介護情報_胃瘻_腸瘻_地域包括ケア as IN_85_1_2
import X_IN_85_1_3_要介護情報_胃瘻_腸瘻_療養 as IN_85_1_3
import X_IN_85_2_1_要介護情報_経鼻胃管_回復リハ as IN_85_2_1
import X_IN_85_2_2_要介護情報_経鼻胃管_地域包括ケア as IN_85_2_2
import X_IN_85_2_3_要介護情報_経鼻胃管_療養 as IN_85_2_3
import X_IN_85_3_1_要介護情報_摂食_嚥下機能障害_回復リハ as IN_85_3_1
import X_IN_85_3_2_要介護情報_摂食_嚥下機能障害_地域包括ケア as IN_85_3_2
import X_IN_85_3_3_要介護情報_摂食_嚥下機能障害_療養 as IN_85_3_3
import X_IN_85_4_1_要介護情報_中心静脈栄養_回復リハ as IN_85_4_1
import X_IN_85_4_2_要介護情報_中心静脈栄養_地域包括ケア as IN_85_4_2
import X_IN_85_4_3_要介護情報_中心静脈栄養_療養 as IN_85_4_3
import X_IN_85_5_1_要介護情報_低栄養_回復リハ as IN_85_5_1
import X_IN_85_5_2_要介護情報_低栄養_地域包括ケア as IN_85_5_2
import X_IN_85_5_3_要介護情報_低栄養_療養 as IN_85_5_3
import X_IN_85_6_1_要介護情報_末梢静脈栄養_回復リハ as IN_85_6_1
import X_IN_85_6_2_要介護情報_末梢静脈栄養_地域包括ケア as IN_85_6_2
import X_IN_85_6_3_要介護情報_末梢静脈栄養_療養 as IN_85_6_3




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
    
        #survey_folder内に当該グループのフォルダがあるか確認する
        survey_folder = os.path.join(individual_folder, "03_慢性期指標")
        if not os.path.exists(survey_folder):
            os.makedirs(survey_folder)
    except:
        print("エラー: 出力先フォルダの設定に失敗しました。年度={}".format(target_year))
        return -1
      
    #report_type_folder内のファイル、フォルダを全て削除
    delete_file_count = 0
    for file in os.listdir(survey_folder):
        try:
            os.remove(os.path.join(survey_folder, file))
            #削除したファイル数をカウント
            delete_file_count += 1
        except:
            print("フォルダ：{}".format(survey_folder))
            print(f"エラー: {file} の削除に失敗しました。")
            print("エラー内容：{}".format(sys.exc_info()[1]))
            return -1

    print("削除したファイル数：{}".format(delete_file_count))
    return survey_folder


# レポート作成対象となる病院を取得
# 指標「0ひと月あたり症例数_回復リハ、地ケア、療養」の最新の年度、期に登録されている病院を対象とする
def get_target_hospitals(target_year,is_acute):
    print("\nレポート作成対象となる病院を取得します。")

    hospID_list = []

    _C = ""
    if is_acute == "慢性期":
        _C = "_C"

    sql_回リハ = "SELECT DISTINCT Hospital_HOSPITAL_ID FROM CI_100" + _C + " WHERE 年度 = " + str(target_year) \
                + " AND 月 = '1ヶ月平均' " + " AND 症例数 <> -2  "
    tempHospID_回リハ = X_01.excuteSQL(sql_回リハ, None)

    sql_地ケア = "SELECT DISTINCT Hospital_HOSPITAL_ID FROM CI_101" + _C + " WHERE 年度 = " + str(target_year) \
                + " AND 月 = '1ヶ月平均' " + " AND 症例数 <> -2  "
    tempHospID_地ケア = X_01.excuteSQL(sql_地ケア, None)

    sql_療養 = "SELECT DISTINCT Hospital_HOSPITAL_ID FROM CI_102" + _C + " WHERE 年度 = " + str(target_year) \
                + " AND 月 = '1ヶ月平均' " + " AND 症例数 <> -2  "
    tempHospID_療養 = X_01.excuteSQL(sql_療養, None)


    if tempHospID_回リハ is not None:
        for a in tempHospID_回リハ:
            hospID_list.append(a[0])

    if tempHospID_地ケア is not None:
        for a in tempHospID_地ケア:
            if a[0] not in hospID_list:
                hospID_list.append(a[0])

    if tempHospID_療養 is not None:
        for a in tempHospID_療養:
            if a[0] not in hospID_list:
                hospID_list.append(a[0])

    hosIDs = ",".join(str(hospID) for hospID in hospID_list)

    sql_target_hospName = "SELECT 病院名, HOSPITAL_ID FROM HOSPITAL WHERE HOSPITAL_ID IN (" + hosIDs+ " ) ORDER BY 都道府県コード, 病院名_よみがな"
        
    temp_hospName_list = X_01.excuteSQL(sql_target_hospName, None)
    if temp_hospName_list is None:
        print("エラー: 対象病院の取得に失敗しました。")
        return -1

    hospName_list = []
    for a in range(len(temp_hospName_list)):
        hospName_list.append((temp_hospName_list[a][0],temp_hospName_list[a][1]))

    #print("hospName_list")
    #print(hospName_list)

    #print("急性期 年度{},期{}の対象病院数:{}".format(maxYear_急性期[0][0], maxQuater_急性期[0][0], len(hospID_list)))
    #print("慢性期 年度{},期{}の対象病院数:{}".format(maxYear_慢性期[0][0], maxQuater_慢性期[0][0], len(hospID_list)))
    print("{} ：対象病院数:{}".format(is_acute,len(hospName_list)))
    return hospName_list


def copyFormatFile(target_year, hospName_list, output_folder,is_acute):
    print("\nフォーマットファイルをコピーします。{}".format(is_acute))

    #病院の数だけ繰り返し
    repCnt = 1

    _C = ""
    if is_acute == "慢性期":
        _C = "_C"
        temp_repCnt = 0
        #output_folder内のxlsmファイルを取得
        xlsm_file_list = [f for f in os.listdir(output_folder) if f.endswith('.xlsm')]

        for xlsm_file in xlsm_file_list:
            xlsm_file_name = os.path.splitext(xlsm_file )[0]
            #xlsm_file_name内にある数値２桁を取得
            xlsm_file_name_num = re.search(r'\d{2}', xlsm_file_name)
       
            if xlsm_file_name_num is not None:
                xlsm_file_name_num = xlsm_file_name_num.group()
            else:
                print("エラー: 数値２桁が取得できませんでした。")
                print("ファイル名:{}".format(xlsm_file_name))
                return -1

            xlsm_file_name_num = int(xlsm_file_name_num)

            if xlsm_file_name_num > temp_repCnt:
                temp_repCnt = xlsm_file_name_num

        repCnt = temp_repCnt + 1
          
    try:
        for hospName in hospName_list:
        #print(str(repCnt) + ":" + hospName + " 開始")
        
            format_file_path = CONST.INDIV_REP_FRMT_CHRONIC_PATH
            if not os.path.exists(format_file_path):
                print("エラー: フォーマットファイルが存在しません。")
                print("フォーマットファイルのパス:{}".format(format_file_path))
                return -1

            temp_repCnt = ""
            if repCnt <= 9:
                temp_repCnt = "0" + str(repCnt)
            else:
                temp_repCnt = str(repCnt)

            newName = "個別レポート_慢性期指標_" + str(temp_repCnt) + "_" + hospName[0] + ".xlsm"

            #format_file_pathをnewNameでコピー
            shutil.copy(format_file_path, os.path.join(output_folder, newName))            
            repCnt += 1

    except:
        print("エラー: フォーマットファイルのコピーに失敗しました。")
        print("エラー内容：{}".format(sys.exc_info()[1]))
        return -1

    print("フォーマットファイルのコピーに成功しました。{}".format(is_acute))
    return 0


def create_indiv_report(target_create_date, hosp_list, output_folder,is_acute):
    print("\n個別レポートを作成します。{}".format(is_acute))

    #output_folder内のxlsmファイルを取得
    xlsm_file_list = [f for f in os.listdir(output_folder) if f.endswith('.xlsm')]
    if len(xlsm_file_list) == 0:
        #存在しない場合もあるので正常終了で返す
        print("エラー: xlsmファイルが存在しません。")
        return 0

    # hospName_list
    # 0 病院名
    # 1 病院ID
    for hosp in hosp_list:
        for xlsm_file in xlsm_file_list:
            if hosp[0] in xlsm_file:
                print("{} を作成します。".format(xlsm_file))
                print("病院ID:{}".format(hosp[1]))
                print("病院名:{}".format(hosp[0]))


                # エクセル（マクロ付き）を開く
                wb = openpyxl.load_workbook(os.path.join(output_folder, xlsm_file), keep_vba=True)

                # 竹口病院は2023は慢性期、2024は急性期、慢性期の両方に提出していたので慢性期からデータを取得
                if hosp[0] == "竹口病院" and is_acute == "急性期":
                    is_acute = "慢性期"


                IN_69_1.getRepAgeData(wb, hosp[1], hosp[0], is_acute)
                IN_69_2.getRepAgeData(wb, hosp[1], hosp[0], is_acute)
                IN_69_3.getRepAgeData(wb, hosp[1], hosp[0], is_acute)
                IN_70_1.getRepAgeData(wb, hosp[1], hosp[0], is_acute)
                IN_70_2.getRepAgeData(wb, hosp[1], hosp[0], is_acute)
                IN_70_3.getRepAgeData(wb, hosp[1], hosp[0], is_acute)
                IN_71_1.getRepAgeData(wb, hosp[1], hosp[0], is_acute)
                IN_71_2.getRepAgeData(wb, hosp[1], hosp[0], is_acute)    
                IN_71_3.getRepAgeData(wb, hosp[1], hosp[0], is_acute)
                IN_72_1.getRepAgeData(wb, hosp[1], hosp[0], is_acute)
                IN_72_2.getRepAgeData(wb, hosp[1], hosp[0], is_acute)
                IN_72_3.getRepAgeData(wb, hosp[1], hosp[0], is_acute)
                IN_73.getRepAgeData(wb, hosp[1], hosp[0], is_acute)
                IN_74.getRepAgeData(wb, hosp[1], hosp[0], is_acute)    
                IN_75_1.getRepAgeData(wb, hosp[1], hosp[0], is_acute)
                IN_75_2.getRepAgeData(wb, hosp[1], hosp[0], is_acute)
                IN_75_3.getRepAgeData(wb, hosp[1], hosp[0], is_acute)
                IN_75_4.getRepAgeData(wb, hosp[1], hosp[0], is_acute)
                IN_75_5.getRepAgeData(wb, hosp[1], hosp[0], is_acute)
                IN_76_1.getRepAgeData(wb, hosp[1], hosp[0], is_acute)
                IN_76_2.getRepAgeData(wb, hosp[1], hosp[0], is_acute)
                IN_76_3.getRepAgeData(wb, hosp[1], hosp[0], is_acute)
                IN_77_1.getRepAgeData(wb, hosp[1], hosp[0], is_acute)
                IN_77_2.getRepAgeData(wb, hosp[1], hosp[0], is_acute)
                IN_77_3.getRepAgeData(wb, hosp[1], hosp[0], is_acute)
                IN_78.getRepAgeData(wb, hosp[1], hosp[0], is_acute)
                IN_79.getRepAgeData(wb, hosp[1], hosp[0], is_acute)
                IN_80.getRepAgeData(wb, hosp[1], hosp[0], is_acute)
                IN_81_1.getRepAgeData(wb, hosp[1], hosp[0], is_acute)
                IN_81_2.getRepAgeData(wb, hosp[1], hosp[0], is_acute)
                IN_82.getRepAgeData(wb, hosp[1], hosp[0], is_acute)
                IN_83.getRepAgeData(wb, hosp[1], hosp[0], is_acute)    
                IN_84_1.getRepAgeData(wb, hosp[1], hosp[0], is_acute)
                IN_84_2.getRepAgeData(wb, hosp[1], hosp[0], is_acute)
                IN_84_3.getRepAgeData(wb, hosp[1], hosp[0], is_acute)
                IN_85_1_1.getRepAgeData(wb, hosp[1], hosp[0], is_acute)
                IN_85_1_2.getRepAgeData(wb, hosp[1], hosp[0], is_acute)
                IN_85_1_3.getRepAgeData(wb, hosp[1], hosp[0], is_acute)
                IN_85_2_1.getRepAgeData(wb, hosp[1], hosp[0], is_acute)
                IN_85_2_2.getRepAgeData(wb, hosp[1], hosp[0], is_acute)
                IN_85_2_3.getRepAgeData(wb, hosp[1], hosp[0], is_acute)
                IN_85_3_1.getRepAgeData(wb, hosp[1], hosp[0], is_acute)
                IN_85_3_2.getRepAgeData(wb, hosp[1], hosp[0], is_acute)
                IN_85_3_3.getRepAgeData(wb, hosp[1], hosp[0], is_acute)
                IN_85_4_1.getRepAgeData(wb, hosp[1], hosp[0], is_acute)
                IN_85_4_2.getRepAgeData(wb, hosp[1], hosp[0], is_acute)
                IN_85_4_3.getRepAgeData(wb, hosp[1], hosp[0], is_acute)
                IN_85_5_1.getRepAgeData(wb, hosp[1], hosp[0], is_acute)
                IN_85_5_2.getRepAgeData(wb, hosp[1], hosp[0], is_acute)
                IN_85_5_3.getRepAgeData(wb, hosp[1], hosp[0], is_acute)
                IN_85_6_1.getRepAgeData(wb, hosp[1], hosp[0], is_acute)
                IN_85_6_2.getRepAgeData(wb, hosp[1], hosp[0], is_acute)
                IN_85_6_3.getRepAgeData(wb, hosp[1], hosp[0], is_acute)
                

                wb.save(os.path.join(output_folder, xlsm_file))
                wb.close()
                
                if hosp[0] == "竹口病院" and is_acute == "慢性期":
                    is_acute = "急性期"













def create_report():
    print("\n経年比較  急性期指標レポートを作成します。")

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

    #作成対象となる病院を取得
    #まれに慢性期の方に多くデータを提出している病院がいるので注意！
    # 現状、「竹口病院」のみ　2024年度末時点

    # 0 病院名
    # 1 病院ID
    hospID_list_急性期 = get_target_hospitals(target_year,"急性期")
    if hospID_list_急性期 == -1:
        return -1

    print("hospID_list_急性期")
    print(hospID_list_急性期)

    # 0 病院名
    # 1 病院ID
    hospID_list_慢性期 = get_target_hospitals(target_year,"慢性期")
    if hospID_list_慢性期 == -1:
        return -1

    #hospID_list_慢性期からhospID_list_急性期との重複を除く
    hospID_list_慢性期 = list(set(hospID_list_慢性期) - set(hospID_list_急性期))

    print("hospID_list_慢性期")
    print(hospID_list_慢性期)

    if copyFormatFile(target_year, hospID_list_急性期, output_folder,"急性期") == -1:   
        return -1
    if copyFormatFile(target_year, hospID_list_慢性期, output_folder,"慢性期") == -1:
        return -1

   # マスタデータをセット
    X_01.setMastaData()

    create_indiv_report(target_create_date, hospID_list_急性期, output_folder,"急性期")
    create_indiv_report(target_create_date, hospID_list_慢性期, output_folder,"慢性期")