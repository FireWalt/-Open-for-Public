#-----------------------------------------------------
#概要：Excelブックを翻訳するプログラム
#補足：複数シートを翻訳する
#　　　本Pythonファイルと同じフォルダに翻訳元Excel(.xlsx)を格納しておく
#　　　翻訳後のファイル名は「翻訳元ファイル名(Translated).xlsx」とする
#-----------------------------------------------------

import datetime
import openpyxl
from googletrans import Translator
import os
import time
from openpyxl.descriptors.base import String

#-----------------------------------------------------
#翻訳実行メソッド
#-----------------------------------------------------
def translate(trans_obj, file_path: String, src_lang_code: String, dest_lang_code: String):
    trans_obj.updateLog('Start translation.\n')

    translator = Translator() #翻訳準備
    workBook = openpyxl.load_workbook(file_path) #ブックを開く
    total_sheet_num = len(workBook.sheetnames) #シート数取得
    sheet_num = 0 #翻訳中のシート番号

    for ws in workBook.worksheets:
        sheet_num += 1
        trans_obj.updateLog('Translating ' + str(sheet_num) + "/" + str(total_sheet_num) + " sheet...\n") #翻訳しているページ数を出力

        for row in ws:
            for cell in row:
                #空白セルは翻訳しない 
                if cell.value is None:
                    continue
                #セルが数値の場合は翻訳しない
                if isinstance(cell.value,int):
                    continue
                #時刻は翻訳しない
                if isinstance(cell.value,datetime.time): 
                    continue
                #日付は翻訳しない
                if isinstance(cell.value,datetime.datetime):
                    continue
                cell.value = translator.translate(str(cell.value), src=src_lang_code, dest=dest_lang_code).text
                time.sleep(0.1) #サーバ負荷軽減処理。データ数が多い場合、時間を増やすことを検討

    file_name_split = os.path.splitext(os.path.basename(file_path)) #翻訳元ファイル名を拡張子以外と拡張子に分割
    dirname = os.path.dirname(file_path) #ファイルディレクトリを取得
    workBook.save(dirname + '/' + file_name_split[0] + "(Translated)" + file_name_split[1]) #翻訳先ファイル名を指定して保存
    trans_obj.updateLog('Translation was completed.\n') #翻訳終了メッセージ

