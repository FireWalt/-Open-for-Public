#-----------------------------------------------------
#概要：Textファイルを翻訳するプログラム
#補足：複数パラグラフを翻訳する
#　　　本Pythonファイルと同じフォルダに翻訳元Textを格納しておく
#　　　翻訳後のファイル名は「翻訳元ファイル名(Translated).拡張子」とする
#　　　エンコードはUTF-8とする。
#-----------------------------------------------------

import os
from time import sleep
from googletrans import Translator
from openpyxl.descriptors.base import String
 
#-----------------------------------------------------
#翻訳実行メソッド
#-----------------------------------------------------
def translate(trans_obj, file_path: String, src_lang_code: String, dest_lang_code: String):
    trans_obj.updateLog('Start translation.\n')
    translator = Translator() #翻訳準備
    #行数カウントとデータの読み込み
    total_row_num = 0
    lines = []
    with open(file_path, "r", encoding="utf-8") as in_f:
        for data in in_f:
            total_row_num +=1
            lines.append(data)

    trans_obj.updateLog('Translating (Number of all lines:' + str(total_row_num) + ')...\n') #翻訳中である旨を表示。

    file_name_split = os.path.splitext(os.path.basename(file_path)) #翻訳元ファイル名を拡張子以外と拡張子に分割
    out_f = open(file_name_split[0] + "(Translated)" + file_name_split[1], 'w', encoding="utf-8") #上書きでファイルを開く

    for i, str_in in enumerate(lines):
        #10の倍数の時、何パラグラフ目を翻訳中か表示
        if (i+1) % 10 == 0:
            trans_obj.updateLog('Translating ' + str(i + 1) + "/" + str(total_row_num) + ' line...\n') #翻訳中である旨を表示。

        if str(str_in)=="":
            continue
        else:
            str_out = translator.translate(str_in, src=src_lang_code, dest=dest_lang_code).text
            out_f.write(str_out + "\n")
            sleep(1) #サーバ負荷軽減処理。
                
    out_f.close()
    trans_obj.updateLog('Translation was completed.\n') #翻訳終了メッセージ

