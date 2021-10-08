# -*- coding:utf-8 -*- bbb
#-----------------------------------------------------
#概要：GUI出力用pythonファイル
#補足：ユーザがファイル拡張子、翻訳元言語、翻訳先言語を選択できるよう、GUIを用意。
#-----------------------------------------------------
from genericpath import isdir
import tkinter.filedialog, tkinter.messagebox
import tkinter as tk
from tkinter import ttk
import translate_excel, translate_word, translate_powerpoint, translate_text, const
import tkinter.scrolledtext as tksc
import datetime
import sys

#-----------------------------------------------------
# メインメソッド
#-----------------------------------------------------
def main():
    trans = Translate()
    trans.createGUI()
    trans.root.mainloop()

#-----------------------------------------------------
# Translateクラス
#-----------------------------------------------------
class Translate():
    def __init__(self):
        self.root = tk.Tk()
        self.file_kind = tk.StringVar()
        self.src_lang = tk.StringVar()
        self.dest_lang = tk.StringVar()
        self.log_txt = 'Log Area\n\n'
        self.log_insert_point = 1

    #-----------------------------------------------------
    # Translate押下後のメソッド
    #-----------------------------------------------------
    def btn_translate_click(self):
        self.st.configure(state='normal')
        self.st.delete('1.0', 'end')        
        self.st.configure(state='disabled') 

        self.updateLog(self.log_txt)
        self.updateLog('---START---\n')
        str_time = datetime.datetime.now().strftime('%Y/%m/%d %H:%M:%S') 
        self.updateLog('Start time: ' + str_time + '\n')

        str_file_kind = self.file_kind.get()
        file_extention = str_file_kind[str_file_kind.find('(') +1:-1]

        fTyp = [(file_extention,file_extention)]
        iDir = const.get_dir_path(sys.argv[0])
        file_path = tkinter.filedialog.askopenfilename(filetypes = fTyp,initialdir = iDir)
        
        # キャンセルの場合はGUI画面に戻る
        if file_path == "":
            self.updateLog('Cancel the translation.\n')
            str_time = datetime.datetime.now().strftime('%Y/%m/%d %H:%M:%S') 
            self.updateLog('End time: ' + str_time + '\n')
            self.updateLog('---END---\n\n')
            return

        # 翻訳コードへ変換
        str_src_lang = self.src_lang.get()
        int_src_lang = int(str_src_lang[:str_src_lang.find('.')])
        src_lang_code = const.transLangChange(int_src_lang)
        str_dest_lang = self.dest_lang.get()
        int_dest_lang = int(str_dest_lang[:str_dest_lang.find('.')])
        dest_lang_code = const.transLangChange(int_dest_lang)

        #翻訳の実行
        if file_extention == '.xlsx':
            translate_excel.translate(self, file_path, src_lang_code, dest_lang_code)
        elif file_extention == '.docx':
            translate_word.translate(self, file_path, src_lang_code, dest_lang_code)
        elif file_extention == '.pptx':
            translate_powerpoint.translate(self, file_path, src_lang_code, dest_lang_code)
        elif file_extention == '.txt':
            translate_text.translate(self, file_path, src_lang_code, dest_lang_code)

        #終了ログ出力
        str_time = datetime.datetime.now().strftime('%Y/%m/%d %H:%M:%S') 
        self.updateLog('End time: ' + str_time + '\n')
        self.updateLog('---END---\n\n')


    #-----------------------------------------------------
    # GUI作成メソッド
    #-----------------------------------------------------
    def createGUI(self):
        # ウインドウのタイトルを定義する
        self.root.title('Translation Tool')

        # ここでウインドウサイズを定義する
        self.root.geometry(const.GUI_WINDOW_SIZE)

        # ラベルを使って文字を画面上に出す
        Static1 = tk.Label(text='1. File type?')
        Static1.pack()
        Static1.place(x=const.GUI_LABEL1_X, y=const.GUI_LABEL1_Y)

        # ファイル種類のドロップダウンリストを出現させる
        style = ttk.Style()
        style.theme_use("winnative")

        tmp_list = []
        tmp_list = const.DIC_FILE_KIND.items()
        combo_list = [r[1] for r in tmp_list]

        combobox = ttk.Combobox(self.root, textvariable= self.file_kind, values=combo_list, state="readonly")
        combobox.current(0)
        combobox.pack()
        combobox.place(x=const.GUI_COMBO1_X, y=const.GUI_COMBO1_Y)

        # ラベルを使って文字を画面上に出す
        Static2 = tk.Label(text='2. Translate from?')
        Static2.pack()
        Static2.place(x=const.GUI_LABEL2_X, y=const.GUI_LABEL2_Y)

        # 翻訳元言語のドロップダウンリストを出現させる
        style2 = ttk.Style()
        style2.theme_use("winnative")

        combo_list = []
        for key,val in const.DIC_LANG_KIND.items():
            combo_list.append(str(key) + '.' + val)

        combobox2 = ttk.Combobox(self.root, textvariable= self.src_lang, values=combo_list, state="readonly")
        combobox2.current(0)
        combobox2.pack()
        combobox2.place(x=const.GUI_COMBO2_X, y=const.GUI_COMBO2_Y)

        # ラベルを使って文字を画面上に出す
        Static3 = tk.Label(text='3. Translate To?')
        Static3.pack()
        Static3.place(x=const.GUI_LABEL3_X, y=const.GUI_LABEL3_Y)

        # 翻訳先言語のドロップダウンリストを出現させる
        style3 = ttk.Style()
        style3.theme_use("winnative")
        combobox3 = ttk.Combobox(self.root, textvariable= self.dest_lang, values=combo_list, state="readonly")
        combobox3.current(1)
        combobox3.pack()
        combobox3.place(x=const.GUI_COMBO3_X, y=const.GUI_COMBO3_Y)

        # Translateボタンを出現させる
        btn_translate = tk.Button(text='Translate!', width=const.GUI_BTN_TRANSLATE_WIDTH, command=self.btn_translate_click)
        btn_translate.pack()
        btn_translate.place(x=const.GUI_BTN_TRANSLATE_X, y=const.GUI_BTN_TRANSLATE_Y)

        # Logフィールドを出現させる
        self.st = tksc.ScrolledText(self.root, width=const.GUI_LABEL_LOG_WIDTH,height=const.GUI_LABEL_LOG_HEIGHT,bg='#dcdcdc')
        self.st.pack(fill='x')
        self.st.place(x = const.GUI_LABEL_LOG_X, y = const.GUI_LABEL_LOG_Y)
        self.updateLog(self.log_txt)

    #-----------------------------------------------------
    # ログ更新メソッド
    #-----------------------------------------------------
    def updateLog(self,text):
        self.st.configure(state='normal')
        self.st.insert(str(self.log_insert_point) + '.0', text)
        self.st.configure(state='disabled')
        self.log_insert_point += text.count('\n')


#-----------------------------------------------------
# メイン処理
#-----------------------------------------------------
if __name__ == "__main__":
    main()
