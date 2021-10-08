#-----------------------------------------------------
#概要：共通部品を格納しているプログラム
#-----------------------------------------------------
import sys
import os

#-----------------------------------------------------
#各種定数
#-----------------------------------------------------
DIC_FILE_KIND = {1:'Excel (.xlsx)', 2:'Word (.docx)', 3:'PowerPoint (.pptx)', 4:'Text (.txt)'}  #ファイル種類のディクショナリ
DIC_LANG_KIND = {1:'Japanese', 2:'English', 3:'Korean', 4:'Chinese (simplified)'}  #ファイル種類のディクショナリ

GUI_WINDOW_WIDTH = 600  # GUI画面のWINDOW幅
GUI_WINDOW_HEIGHT = 500 # GUI画面のWINDOW高さ
GUI_WINDOW_SIZE = str(GUI_WINDOW_WIDTH) + 'x' + str(GUI_WINDOW_HEIGHT)    # GUI画面のWINDOWサイズ

GUI_LABEL1_X = 0    # GUI画面の第一ラベルのx座標
GUI_LABEL1_Y = 20   # GUI画面の第一ラベルのy座標
GUI_COMBO1_X = 120  # GUI画面の第一コンボボックスのx座標
GUI_COMBO1_Y = 20   # GUI画面の第一コンボボックスのy座標

GUI_LABEL2_X = 0    # GUI画面の第二ラベルのx座標
GUI_LABEL2_Y = 60   # GUI画面の第二ラベルのy座標
GUI_COMBO2_X = 120  # GUI画面の第二コンボボックスのx座標
GUI_COMBO2_Y = 60   # GUI画面の第二コンボボックスのy座標

GUI_LABEL3_X = 0    # GUI画面の第三ラベルのx座標
GUI_LABEL3_Y = 100  # GUI画面の第三ラベルのy座標
GUI_COMBO3_X = 120  # GUI画面の第三コンボボックスのx座標
GUI_COMBO3_Y = 100  # GUI画面の第三コンボボックスのy座標

GUI_BTN_TRANSLATE_WIDTH = 15    # GUI画面のTranslateボタンの幅
GUI_BTN_TRANSLATE_X = 10        # GUI画面のTranslateボタンのx座標
GUI_BTN_TRANSLATE_Y = 140       # GUI画面のTranslateボタンのy座標

GUI_BTN_TRANSLATE_X = 0     # GUI画面のTranslateボタンのx座標
GUI_BTN_TRANSLATE_Y = 160   # GUI画面のTranslateボタンのy座標

GUI_LABEL_LOG_HEIGHT = 20  # GUI画面のLOG領域の高さ
GUI_LABEL_LOG_WIDTH = 80   # GUI画面のLOG領域の幅
GUI_LABEL_LOG_X = 0 # GUI画面のLOG領域のx座標
GUI_LABEL_LOG_Y = 200   # GUI画面のLOG領域のy座標

#-----------------------------------------------------
#翻訳コード変換処理
#-----------------------------------------------------
def transLangChange(lang_num: int) -> str:
    if lang_num == 1: #日本語
        lang_code = "ja"
    elif lang_num == 2: #英語
        lang_code = "en"
    elif lang_num == 3: #韓国語
        lang_code = "ko"
    elif lang_num == 4: #簡体中国語
        lang_code = "zh-cn"
    else: #繁体中国語
        lang_code = "zh-tw"
    return lang_code

def get_dir_path(filename):
   if getattr(sys, "frozen", False):
       # The application is frozen
       datadir = os.path.dirname(sys.executable)
   else:
       # The application is not frozen
       # Change this bit to match where you store your data files:
       datadir = os.path.dirname(filename)
   return datadir
   