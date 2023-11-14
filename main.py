import tkinter as tk
import customtkinter as ctk
from customtkinter import filedialog
import os
import pdf_siori018 as pdfs

#import news_release_scrape as nrs
FONT_TYPE = "meiryo"

jd = pdfs.get_json_data() # JSONデータの全て

class MyTabView(ctk.CTkTabview):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.tablist = ["PDF-csv", "PDF-PDF", "Merge", "Option"] 

        # create tabs
        for ti in self.tablist:
            self.add(ti)

    def get_tablist(self):
        return self.tablist

class App(ctk.CTk):

    def __init__(self):
        super().__init__()

        # メンバー変数の設定
        self.fonts = (FONT_TYPE, 15)
        self.csv_filepath = None

        # フォームのセットアップをする
        self.setup_form()

    def setup_form(self):
        # ctk のフォームデザイン設定
        ctk.set_appearance_mode("dark")  # Modes: system (default), light, dark
        ctk.set_default_color_theme("blue")  # Themes: blue (default), dark-blue, green

        # フォームサイズ設定
        self.geometry("800x300")
        self.title("csv-pdf converter ver18.5 TTC")

        # 行方向のマスのレイアウトを設定する。リサイズしたときに一緒に拡大したい行をweight 1に設定。
        self.grid_rowconfigure(1, weight=1)
        # 列方向のマスのレイアウトを設定する
        self.grid_columnconfigure(0, weight=1)

        # タブの設定
        self.tab_view = MyTabView(master=self)
        self.tab_view.grid(row=0, column=0, padx=20, pady=20, sticky="ew")

        tablist = self.tab_view.get_tablist()

        # 機能1
        self.url_csv_frame = Frame01(master=self.tab_view.tab(tablist[0]), header_name="PDFからcsv, csvからPDFへの栞データ変換")
        self.url_csv_frame.grid(row=0, column=0, padx=20, pady=20, sticky="ew")

        # stickyは拡大したときに広がる方向のこと。nsew で4方角で指定する。
        # 機能2
        self.read_file_frame = Frame02(master=self.tab_view.tab(tablist[1]), header_name="PDFからPDFへの栞移行")
        self.read_file_frame.grid(row=0, column=0, padx=20, pady=20, sticky="ew")

        # 機能3
        self.excel_pdf_frame = Frame03(master=self.tab_view.tab(tablist[2]), header_name="栞付きPDFの合致")
        self.excel_pdf_frame.grid(row=0, column=0, padx=20, pady=20, sticky="ew")

        # 機能4
        self.option_frame = Frame04(master=self.tab_view.tab(tablist[3]), header_name="作業フォルダ(ディレクトリ)の変更")
        self.option_frame.grid(row=0, column=0, padx=20, pady=20, sticky="ew")

class Frame01(ctk.CTkFrame):
    def __init__(self, *args, header_name="UrlCsvFrame", **kwargs):
        super().__init__(*args, **kwargs)
        
        self.fonts = (FONT_TYPE, 15)
        self.header_name = header_name

        # フォームのセットアップをする
        self.setup_form()

    def setup_form(self):
        # 行方向のマスのレイアウトを設定する。リサイズしたときに一緒に拡大したい行をweight 1に設定。
        self.grid_rowconfigure(0, weight=1)
        # 列方向のマスのレイアウトを設定する
        self.grid_columnconfigure(0, weight=1)

        # フレームのラベルを表示
        self.label = ctk.CTkLabel(self, text=self.header_name, font=(FONT_TYPE, 11))
        self.label.grid(row=0, column=0, padx=30)

        # 選択
        self.select_box = ctk.CTkComboBox(master=self, values=['csv to PDF', 'PDF to csv', 'new csv'], command=self.on_select)
        self.select_box.set('csv to PDF')  # 初期選択肢を設定
        self.select_box.grid(row=1, column=0, pady=10, padx=10, columnspan=1) #, pady=20

        # ファイルパスを指定するテキストボックス。これだけ拡大したときに、幅が広がるように設定する。
        self.textbox1 = ctk.CTkEntry(master=self, placeholder_text="PDF ファイル選択", font=self.fonts, width=400)
        self.textbox1.grid(row=3, column=0, padx=10, columnspan=2, sticky="ew") #, pady=(0,10)
        # デフォルト値を設定
        #self.textbox1.insert(0, "data.csv")

        # ファイル選択ボタン
        self.button_select = ctk.CTkButton(master=self, 
            fg_color="transparent", border_width=2, text_color=("gray10", "#DCE4EE"),   # ボタンを白抜きにする
            command=self.button_select_callback_pdf, text="pdf選択", font=self.fonts)
        self.button_select.grid(row=3, column=2, padx=10, ) #pady=(0,10)

        # ファイルパスを指定するテキストボックス。これだけ拡大したときに、幅が広がるように設定する。
        self.textbox2 = ctk.CTkEntry(master=self, placeholder_text="csv ファイル選択", font=self.fonts, width=400)
        self.textbox2.grid(row=4, column=0, padx=10, columnspan=2, sticky="ew") #, pady=(0,10)
        #self.textbox2.insert(0, "data.csv")

        # ファイル選択ボタン
        self.button_select = ctk.CTkButton(master=self, 
            fg_color="transparent", border_width=2, text_color=("gray10", "#DCE4EE"),   # ボタンを白抜きにする
            command=self.button_select_callback_csv, text="csv選択", font=self.fonts)
        self.button_select.grid(row=4, column=2, padx=10) #, pady=(0,10)
        
        # 後付け文字
        back_label = ctk.CTkLabel(master=self, text="作成PDF後付文字")
        back_label.grid(row=5, column=0, padx=20, sticky="")

        # 〇
        self.textbox31 = ctk.CTkEntry(master=self, placeholder_text="_栞", width=200, font=self.fonts)
        self.textbox31.grid(row=5, column=1, padx=10, sticky="")
        
        # 実行ボタン
        self.button_open = ctk.CTkButton(master=self, command=self.button_open_callback, text="書込み", font=self.fonts)
        self.button_open.grid(row=5, column=2, padx=10) #, pady=(0,10)

    def button_select_callback_csv(self):
        """
        選択ボタンが押されたときのコールバック。ファイル選択ダイアログを表示する
        """
        # エクスプローラーを表示してファイルを選択する
        file_name = Frame02.file_read_csv()

        if file_name is not None:
            # ファイルパスをテキストボックスに記入
            self.textbox2.delete(0, tk.END)
            self.textbox2.insert(0, file_name)

    def button_select_callback_pdf(self):
        """
        選択ボタンが押されたときのコールバック。ファイル選択ダイアログを表示する
        """
        # エクスプローラーを表示してファイルを選択する
        file_name = Frame02.file_read_pdf()

        if file_name is not None:
            # ファイルパスをテキストボックスに記入
            self.textbox1.delete(0, tk.END)
            self.textbox1.insert(0, file_name)

    def button_open_callback(self):
        """
        書込みボタンが押されたときのコールバック。暫定機能として、ファイルの中身をprintする
        """
        # 変数メモ
        # self.select_box :: 'csv to PDF', 'PDF to csv', 'new csv'
        # self.textbox1 :: '.pdf'
        # self.textbox2 :: '.csv'
        # self.textbox31:: '_栞'
        select = self.select_box.get() # 'csv to PDF', 'PDF to csv', 'new csv' 
        pdf = self.textbox1.get() # PDF
        csv = self.textbox2.get()
        backstr = self.textbox31.get()
        if backstr == "":
            backstr = "_栞"
        # ファイル名と拡張子を取得
        file_name, file_extension = os.path.splitext(pdf)
        # 新しいファイルパスを作成
        save = file_name + backstr + file_extension
        
        if select == 'csv to PDF':
            print(f"変換方法={select}\ncsv={csv}\npdf={pdf}\n新規pdf={save}")
            pdfs.csv_to_pdf(pdf, csv, save)

        elif select == 'PDF to csv':
            print(f"変換方法={select}\npdf={pdf}\ncsv={csv}")
            pdfs.pdf_to_csv(pdf, csv)
        
        elif select == 'new csv':
            print(f"CSV作成={select}\ncsv={csv}")
            pdfs.default_csv(csv)
        
        else:
            print('SELECTED ERROR')

        """
        
        関数実装
        
        """
        print("作成完了")

    def on_select(self, event):
        selected_item = self.select_box.get() # 'csv to PDF', 'PDF to csv', 'new csv' 
        textbox1_item = self.textbox1.get() # PDF
        textbox2_item = self.textbox2.get() # CSV
        
        textbox31_item = self.textbox31.get()
        # 選択が変更されたときに実行したい処理をここに記述する
        print("選択されたアイテム:", selected_item)
        
        if selected_item == "PDF to csv":
            if textbox1_item == "入力不要":
                self.textbox1.delete(0, 'end')
            #if textbox31_item == "入力不要":
            #    self.textbox31.delete(0, 'end')
            if textbox2_item == "new.csv":
                self.textbox2.delete(0, 'end')
                self.textbox2.insert(0, "data.csv")
            
            self.textbox31.delete(0, 'end')
            self.textbox31.insert(0, '入力不要')

        elif selected_item == "csv to PDF":
            if textbox1_item == "入力不要":
                self.textbox1.delete(0, 'end')
            if textbox31_item == "入力不要":
                self.textbox31.delete(0, 'end')
            if textbox2_item == "new.csv":
                self.textbox2.delete(0, 'end')
                self.textbox2.insert(0, "data.csv")
                
        elif selected_item == "new csv":
            self.textbox2.delete(0, 'end')
            self.textbox2.insert(0, "new.csv")

            self.textbox1.delete(0, 'end')
            self.textbox1.insert(0, "入力不要")

            self.textbox31.delete(0, 'end')
            self.textbox31.insert(0, "入力不要")
            
class Frame02(ctk.CTkFrame):
    def __init__(self, *args, header_name="ReadFileFrame", **kwargs):
        super().__init__(*args, **kwargs)
        
        self.fonts = (FONT_TYPE, 15)
        self.header_name = header_name

        # フォームのセットアップをする
        self.setup_form()

    def setup_form(self):
        # 行方向のマスのレイアウトを設定する。リサイズしたときに一緒に拡大したい行をweight 1に設定。
        self.grid_rowconfigure(0, weight=1)
        # 列方向のマスのレイアウトを設定する
        self.grid_columnconfigure(0, weight=1)

        # フレームのラベルを表示
        self.label = ctk.CTkLabel(self, text=self.header_name, font=(FONT_TYPE, 11))
        self.label.grid(row=0, column=0, padx=30)

        # ファイルパスを指定するテキストボックス。これだけ拡大したときに、幅が広がるように設定する。
        self.textbox1 = ctk.CTkEntry(master=self, placeholder_text="PDF ファイル選択", font=self.fonts, width=400)
        self.textbox1.grid(row=3, column=0, padx=20, pady=(0,10), columnspan=2, sticky="ew")

        # ファイル選択ボタン
        self.button_select = ctk.CTkButton(master=self, 
            fg_color="transparent", border_width=2, text_color=("gray10", "#DCE4EE"),   # ボタンを白抜きにする
            command=self.button_select_callback_pdf, text="栞ありpdf選択", font=self.fonts)
        self.button_select.grid(row=3, column=2, padx=10, pady=(0,10))

        # ファイルパスを指定するテキストボックス。これだけ拡大したときに、幅が広がるように設定する。
        self.textbox2 = ctk.CTkEntry(master=self, placeholder_text="PDF ファイル選択", font=self.fonts)
        self.textbox2.grid(row=4, column=0, padx=20, pady=(0,10), columnspan=2, sticky="ew")
        #self.textbox2.insert(0, "data.csv")

        # ファイル選択ボタン
        self.button_select = ctk.CTkButton(master=self, 
            fg_color="transparent", border_width=2, text_color=("gray10", "#DCE4EE"),   # ボタンを白抜きにする
            command=self.button_select_callback_pdf2, text="栞なしpdf選択", font=self.fonts)
        self.button_select.grid(row=4, column=2, padx=10, pady=(0,10))
        
        # 年
        year_label = ctk.CTkLabel(master=self, text="作成PDF後付文字")
        year_label.grid(row=5, column=0, padx=20, sticky="")

        # 〇
        self.textbox31 = ctk.CTkEntry(master=self, placeholder_text="_栞", width=200, font=self.fonts)
        self.textbox31.grid(row=5, column=1, padx=10, sticky="")

        
        # 実行ボタン
        self.button_open = ctk.CTkButton(master=self, command=self.button_open_callback, text="書込み", font=self.fonts)
        self.button_open.grid(row=5, column=2, padx=10) #, pady=(0,10)

    def button_select_callback_pdf2(self):
        """
        選択ボタンが押されたときのコールバック。ファイル選択ダイアログを表示する
        """
        # エクスプローラーを表示してファイルを選択する
        file_name = Frame02.file_read_pdf()

        if file_name is not None:
            # ファイルパスをテキストボックスに記入
            self.textbox2.delete(0, tk.END)
            self.textbox2.insert(0, file_name)

    def button_select_callback_pdf(self):
        """
        選択ボタンが押されたときのコールバック。ファイル選択ダイアログを表示する
        """
        # エクスプローラーを表示してファイルを選択する
        file_name = Frame02.file_read_pdf()

        if file_name is not None:
            # ファイルパスをテキストボックスに記入
            self.textbox1.delete(0, tk.END)
            self.textbox1.insert(0, file_name)

    def button_open_callback(self):
        
        pdf = self.textbox1.get() # PDF
        pdf2 = self.textbox2.get()
        backstr = self.textbox31.get()
        
        if backstr == "":
            backstr = "_栞"

        # ファイル名と拡張子を取得
        file_name, file_extension = os.path.splitext(pdf2)
        # 新しいファイルパスを作成
        save = file_name + backstr + file_extension
        
        print(f"変換方法=PDF to PDF\n栞ありpdf={pdf}\n栞なしpdf={pdf2}\n新規pdf={save}")

        """
        
        関数実装
        
        """
        pdfs.pdf_to_pdf(pdf, pdf2, save)
        print("作成完了")
    
    @staticmethod
    def file_read_csv():
        """
        ファイル選択ダイアログを表示する
        """
        current_dir = os.path.abspath(os.path.dirname(__file__))
        file_path = tk.filedialog.askopenfilename(filetypes=[("csvファイル","*.csv")],initialdir=current_dir)

        if len(file_path) != 0:
            return file_path
        else:
            # ファイル選択がキャンセルされた場合
            return None

    @staticmethod
    def file_read_elsx():
        """
        ファイル選択ダイアログを表示する
        """
        current_dir = os.path.abspath(os.path.dirname(__file__))
        file_path = tk.filedialog.askopenfilename(filetypes=[("Excelファイル","*.xlsx")],initialdir=current_dir)

        if len(file_path) != 0:
            return file_path
        else:
            # ファイル選択がキャンセルされた場合
            return None
        
    @staticmethod
    def file_read_pdf():
        """
        ファイル選択ダイアログを表示する
        """
        current_dir = os.path.abspath(os.path.dirname(__file__))
        #file_path = tk.filedialog.askopenfilename(filetypes=[("PDFファイル","*.pdf")],initialdir=current_dir)
        filetypes = [("PDFファイル", "*.pdf"), ("PDFファイル", "*.PDF")]
        file_path = tk.filedialog.askopenfilename(filetypes=filetypes, initialdir=current_dir)

        if len(file_path) != 0:
            return file_path
        else:
            # ファイル選択がキャンセルされた場合
            return None

class Frame03(ctk.CTkFrame):
    def __init__(self, *args, header_name="ExcelPdfFrame", **kwargs):
        super().__init__(*args, **kwargs)
        
        self.fonts = (FONT_TYPE, 15)
        self.header_name = header_name

        # フォームのセットアップをする
        self.setup_form()

    def setup_form(self):
        # 行方向のマスのレイアウトを設定する。リサイズしたときに一緒に拡大したい行をweight 1に設定。
        self.grid_rowconfigure(0, weight=1)
        # 列方向のマスのレイアウトを設定する
        self.grid_columnconfigure(0, weight=1)

        # フレームのラベルを表示
        self.label = ctk.CTkLabel(self, text=self.header_name, font=(FONT_TYPE, 11))
        self.label.grid(row=0, column=0, padx=30)

        # ファイルパスを指定するテキストボックス。これだけ拡大したときに、幅が広がるように設定する。
        self.textbox2 = ctk.CTkEntry(master=self, placeholder_text="PDF フォルダ選択", font=self.fonts, width=400)
        self.textbox2.grid(row=2, column=0, padx=10, columnspan=2, sticky="ew")

        # フォルダ選択ボタン
        self.button_select = ctk.CTkButton(master=self, 
            fg_color="transparent", border_width=2, text_color=("gray10", "#DCE4EE"),   # ボタンを白抜きにする
            command=self.open_folder_dialog, text="フォルダ参照", font=self.fonts)
        self.button_select.grid(row=2, column=2, padx=10)

        # 選択
        self.select_box21 = ctk.CTkComboBox(master=self, values=['栞無', 'ファイル名栞', 'ファイル名+元栞', '元栞'], command=self.on_select21)
        #self.select_box21.set('ファイル名栞')  # 初期選択肢を設定
        self.select_box21.set(jd["合致後栞方法"])  # 初期選択肢を設定
        self.select_box21.grid(row=3, column=1, padx=10, columnspan=1)

        # 選択
        self.select_box22 = ctk.CTkComboBox(master=self, values=['親path', '子path'], command=self.on_select22)
        #self.select_box22.set('親path')  # 初期選択肢を設定
        self.select_box22.set(jd["合致後保存場所"])  # 初期選択肢を設定
        self.select_box22.grid(row=4, column=1, padx=10, columnspan=1)

        # 名前
        name_label = ctk.CTkLabel(master=self, text="ファイル名")
        name_label.grid(row=5, column=0, padx=20, sticky="")

        # 〇
        self.textbox3 = ctk.CTkEntry(master=self, placeholder_text="merge.pdf", width=300, font=self.fonts)
        self.textbox3.grid(row=5, column=1, padx=10, sticky="")
        #self.textbox3.insert(0, "merge.pdf")
        self.textbox3.insert(0, jd["合致後ファイル名"])

        # ボタン
        self.button_open = ctk.CTkButton(master=self, command=self.button_open_callback, text="PDFマージ", font=self.fonts)
        self.button_open.grid(row=5, column=2, padx=10)

    def button_open_callback(self):
        path = self.textbox2.get()
        select = self.select_box21.get() # '栞無', 'ファイル名栞', 'ファイル名+元栞', '元栞'
        parent = self.select_box22.get() # '親path', '子path'
        name = self.textbox3.get() # name
        
        if parent == '親path':
            save_dir = os.path.dirname(path)
        else:
            save_dir = path
            
        save = os.path.join(save_dir, name)

        print(f"栞マージ方法={select}, 栞保存位置={parent}\npdf_path={path}\nsave={save}")

        """
        
        関数実装
        
        """
        pdfs.pdf_siori_merger(select, path, save)
        print("マージ完了")

    def open_folder_dialog(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            print("選択されたフォルダパス:", folder_path)
        self.textbox2.delete(0, tk.END)
        self.textbox2.insert(0, folder_path)
        
    def on_select21(self, event):
        selected_item = self.select_box21.get()
        # 選択が変更されたときに実行したい処理をここに記述する
        print("選択されたアイテム:", selected_item)

    def on_select22(self, event):
        selected_item = self.select_box22.get()
        # 選択が変更されたときに実行したい処理をここに記述する
        print("選択されたアイテム:", selected_item)


class Frame04(ctk.CTkFrame):
    def __init__(self, *args, header_name="OptionFrame", **kwargs):
        super().__init__(*args, **kwargs)
        
        self.fonts = (FONT_TYPE, 15)
        self.header_name = header_name

        # フォームのセットアップをする
        self.setup_form()

    def setup_form(self):
        # 行方向のマスのレイアウトを設定する。リサイズしたときに一緒に拡大したい行をweight 1に設定。
        self.grid_rowconfigure(0, weight=1)
        # 列方向のマスのレイアウトを設定する
        self.grid_columnconfigure(0, weight=1)

        # フレームのラベルを表示
        self.label = ctk.CTkLabel(self, text=self.header_name, font=(FONT_TYPE, 11))
        self.label.grid(row=0, column=0, padx=20, sticky="ew")

        # ファイルパスを指定するテキストボックス。これだけ拡大したときに、幅が広がるように設定する。
        self.textbox2 = ctk.CTkEntry(master=self, placeholder_text="CD選択", font=self.fonts, width=400)
        self.textbox2.grid(row=1, column=0, padx=10, columnspan=2, sticky="ew")

        # フォルダ参照ボタンの作成とクリックイベントのバインド
        self.folder_button = ctk.CTkButton(master=self, text="CD変更", command=self.open_folder_dialog)
        self.folder_button.grid(row=1, column=2, padx=20)
        
        # ボタン
        #self.button_open = ctk.CTkButton(master=self, command=self.button_open_callback, text="PDF削除", font=self.fonts)
        #self.button_open.grid(row=2, column=4, padx=10, sticky="")

    def open_folder_dialog(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            print("選択されたフォルダパス:", folder_path)
        self.textbox2.delete(0, tk.END)
        self.textbox2.insert(0, folder_path)
        os.chdir(folder_path)
        print("CD変更完了")

if __name__ == "__main__":
    
    # cxfreeze -c GUICTk.py --target-dir target2 の時はコメントを外す。テスト字はコメント。
    #current_dir = os.getcwd()  # 現在のディレクトリを取得
    #parent_dir = os.path.dirname(current_dir)  # 親ディレクトリを取得
    #os.chdir(parent_dir)  # 親ディレクトリに移動
    
    app = App()
    app.mainloop()
