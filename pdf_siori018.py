import os,sys
import fitz
import csv
from PyPDF2 import PdfFileWriter, PdfFileReader
import datetime
import json

dt_now = datetime.datetime.now()
if dt_now.day < 15: #前半
    pattern = "前半"
else:
    pattern = "後半" #後半
cd0 = os.getcwd()  # 現在のディレクトリを取得

def pdf_to_csv(pdf_in, csv_out):
    doc = fitz.open(pdf_in)
    page_count = doc.page_count # 総ページ数
    print('Page count: ', page_count)
    toc_list = doc.get_toc()
    toc_list = [["Depth", "Title", "Page"]] + toc_list

    with open(csv_out, 'w', encoding="utf-8-sig", newline="") as f: #空白行をなくすためnewline追加(220329)
        writer = csv.writer(f)
        writer.writerows(toc_list) #org
    return()

def csv_to_pdf(pdf_in, csv_out, save):
    with open(csv_out, "r", encoding="utf-8-sig") as f: #csvの読込み
        reader = csv.reader(f, delimiter=',')
        l = [row for row in reader]
    
    l2 = [] #csvの空データ行削除
    for i in range(len(l)):
        if l[i] == []:
            pass
        elif "Depth" in l[i]: #要素にDepthがある場合をColumnと考え飛ばす
            pass
        else:
            #print(l[i])
            l2.append(l[i])

    output = PdfFileWriter() # open output
    input1 = PdfFileReader(open(pdf_in, 'rb')) # open input
    
    n = input1.getNumPages()
    for i in range(n):
        output.addPage(input1.getPage(i)) # insert page

    par_list = [] #階層のある場合のための階層親リスト
    kaiso_list = [] #階層番号のリスト

    for i in l2:
        kaiso = int(i[0])
        kaiso_list.append(kaiso)
        name = i[1]
        page = int(i[2])
        #print(kaiso,name,page)
        #page は -1 しないとずれる(2022.03.28)
        if kaiso == 1:
            par_list.append(output.addBookmark(name, page-1, parent=None))
        else:
            ln = len(kaiso_list)
            par_kaiso = kaiso-1 #親の階層
            r_kaiso = kaiso_list[::-1] #階層リストの逆順
            ind = r_kaiso.index(par_kaiso) #逆から親の階層検索
            p_ind = ln - ind - 1 #par_listの親階層インデクス
            par = par_list[p_ind]
            par_list.append(output.addBookmark(name, page-1, parent=par))

    base_pdf_in = os.path.splitext(os.path.basename(pdf_in))[0] #拡張子なしのファイル名
    #sname = base_pdf_in + "_栞.pdf"
    #outputStream = open('result.pdf','wb') #creating result pdf JCT
    sname = save
    outputStream = open(sname,'wb') #221105
    output.write(outputStream) #writing to result pdf JCT
    outputStream.close() #closing result JCT

#221108開発 PDF[0]の栞をPDF[1]に適用する。
def pdf_to_pdf(pdf_in0, pdf_in1, save):
    csv_in0 = os.path.splitext(os.path.basename(pdf_in0))[0] + ".csv" #拡張子なしのファイル名
    pdf_to_csv(pdf_in0, csv_in0)
    print(pdf_in0)
    csv_to_pdf(pdf_in1, csv_in0, save)
    print("完了 pdf[0]->pdf[1]")

# 230620開発 何も書いていないcsvファイルを作成する
def default_csv(csv_out):
    #csv_out = "new.csv"
    toc_list = [["Depth", "Title", "Page"]]

    with open(csv_out, 'w', encoding="utf-8-sig", newline="") as f: #空白行をなくすためnewline追加(220329)
        writer = csv.writer(f)
        writer.writerows(toc_list) #org


def combine_pdfs(pdf_list, out_pdf="combined_pdfs.pdf"): #複数のPDFを合体させる関数。
    pdf_writer = PdfFileWriter()

    for filename in pdf_list:
        pdf_reader = PdfFileReader(filename)
        for page in range(pdf_reader.getNumPages()):
            pdf_writer.addPage(pdf_reader.getPage(page))

    with open(out_pdf, 'wb') as fh:
        pdf_writer.write(fh)

def tocmaker(dirPath, sorted_files):
    # from PyPDF2 import PdfFileWriter, PdfFileReader
    toc_list = []
    page0 = 1
    for si in sorted_files: #ファイル順に実行。ファイル名とページ番号の対応表作成。
        base, ext = os.path.splitext(si)
        pdf_file = dirPath + "/" + si
        pdf_reader = PdfFileReader(pdf_file)
        pages = pdf_reader.numPages
        toc_list.append([1, base, page0]) #深さ,　目次名, ページ番号
        page0 += pages #合体させた時のページ番号（１ページ開始とする）
    return(toc_list)

def tocmaker_second_toc(dirPath, sorted_files):
    # from PyPDF2 import PdfFileWriter, PdfFileReader
    toc_list = []
    page0 = 1
    for si in sorted_files: #ファイル順に実行。ファイル名とページ番号の対応表作成。
        base, ext = os.path.splitext(si)
        pdf_file = dirPath + "/" + si
        pdf_reader = PdfFileReader(pdf_file)
        pages = pdf_reader.numPages
        
        # ファイル名の栞
        toc_list.append([1, base, page0]) #深さ,　目次名, ページ番号

        # 各ファイルに含まれる栞
        doc = fitz.open(pdf_file)
        toc_list_i = doc.get_toc()
        for tl in toc_list_i:
            toc_list.append([1 + tl[0], tl[1], page0 + tl[2] - 1])

        page0 += pages #合体させた時のページ番号（１ページ開始とする）
    return(toc_list)

def tocmaker_original_toc(dirPath, sorted_files):
    # from PyPDF2 import PdfFileWriter, PdfFileReader
    toc_list = []
    page0 = 1
    for si in sorted_files: #ファイル順に実行。ファイル名とページ番号の対応表作成。
        base, ext = os.path.splitext(si)
        pdf_file = dirPath + "/" + si
        pdf_reader = PdfFileReader(pdf_file)
        pages = pdf_reader.numPages
        
        # ファイル名の栞
        #toc_list.append([1, base, page0]) #深さ,　目次名, ページ番号

        # 各ファイルに含まれる栞
        doc = fitz.open(pdf_file)
        toc_list_i = doc.get_toc()
        for tl in toc_list_i:
            toc_list.append([0 + tl[0], tl[1], page0 + tl[2] - 1])

        page0 += pages #合体させた時のページ番号（１ページ開始とする）
    return(toc_list)

def list_to_pdf(pdf_in, l2, out_pdf="toced_pdfs.pdf"):
    output = PdfFileWriter() # open output

    fh = open(pdf_in, "rb")
    #input1 = PdfFileReader(open(pdf_in, 'rb')) # open input
    input1 = PdfFileReader(fh) # open input
    n = input1.getNumPages()
    for i in range(n):
        output.addPage(input1.getPage(i)) # insert page

    par_list = [] #階層のある場合のための階層親リスト
    kaiso_list = [] #階層番号のリスト

    for i in l2:
        kaiso = int(i[0])
        kaiso_list.append(kaiso)
        name = i[1]
        page = int(i[2])
        if kaiso == 1:
            par_list.append(output.addBookmark(name, page-1, parent=None))
        else:
            ln = len(kaiso_list)
            par_kaiso = kaiso-1 #親の階層
            r_kaiso = kaiso_list[::-1] #階層リストの逆順
            ind = r_kaiso.index(par_kaiso) #逆から親の階層検索
            p_ind = ln - ind - 1 #par_listの親階層インデクス
            par = par_list[p_ind]
            par_list.append(output.addBookmark(name, page-1, parent=par))

    outputStream = open(out_pdf,'wb') #creating result pdf JCT
    output.write(outputStream) #writing to result pdf JCT
    outputStream.close() #closing result JCT

    fh.close()

def sort_files_TTC(list1, order): # TTC用の並べ替え関数
    # chatGPTによるソート辞書を作成する
    dic = {f:i for i,f in enumerate(order)}
    order_1st = order + ["業務報告"] #イレギュラーな発表にも対応したい。
    # 名前のあるファイルとないファイルを分離
    included_names = []
    excluded_names = []
    for name in list1:
        if any(n in name for n in order_1st): # 名前の有無で判断
            #if "業務報告" in name: #業務報告の有無で判断
            included_names.append(name)
        else:
            excluded_names.append(name)

    print("Included names:", included_names)
    print("Excluded names:", excluded_names)

    # 業務報告ファイルの並び替え（同じ名前が2回出る場合対応）
    sorted_files = []
    # 名前リスト毎に検索
    for orderi in order:
        boo = "1"
        while boo == "1" and len(included_names) != 0:
            # ファイルリスト毎に検索
            for i in range(len(included_names)):
                # 名前をみつけた場合
                if orderi in included_names[i]:
                    sorted_files.append(included_names[i])
                    del included_names[i] # 該当ファイルリストから要素を削除
                    boo = "1"
                    break # 一段ループを出る
                # 最後まで探しても見つからない=もう存在しない
                elif i == len(included_names)-1:
                    boo = "0"
    print("Sorted files:", sorted_files)
    sorted_files += included_names + excluded_names # includded_nameは空=[]のはずだが、想定外の場合に対応。
    print("Sorted files All:", sorted_files)
    return sorted_files

def csvStr_to_list(data): # """\nA,A,A\n"""->["A","A","A"]
    data = data.split("\n")
    
    # csv文字列データをリストに変換
    data1 = []
    for cc in data:
        data1 += cc.split(",")
    
    # 空データ削除
    data1 = [item for item in data1 if item not in (None, "", "nan"," ")]
    return data1

def select_mode1(data_js): # jsonを元にマージする時のPDF並び順番リスト作成
    if data_js["option"] == "":
        print("TTCモード")
        if pattern == "前半":
            order = csvStr_to_list(data_js["term1"]) # "第２木曜日の全員出席時"
        else:
            order = csvStr_to_list(data_js["term2"]) # "第４火曜日の全員出席時"
    else:
        print("option")
        order = csvStr_to_list(data_js["option"])
    return order

def select_mode(): # jsonファイルが存在すること確認してからマージ
    directory = cd0 # 初期位置のディレクトリを取得
    # ファイルの存在を確認
    json_path = os.path.join(directory, "init.json")
    if os.path.exists(json_path):
        # JSONファイルからデータを読み込む
        with open(json_path, "r", encoding="utf-8") as file:
            data_js = json.load(file)
        order = select_mode1(data_js)
        print(order)
    else:
        print("stableモード") # 並び順なしのソート
        order = None
    return order

def get_json_data():
    directory = cd0 # 初期位置のディレクトリを取得
    # ファイルの存在を確認
    json_path = os.path.join(directory, "init.json")
    if os.path.exists(json_path):
        # JSONファイルからデータを読み込む
        with open(json_path, "r", encoding="utf-8") as file:
            data_js = json.load(file)
    else:
        data_js = None
    return data_js


def pdf_siori_merger(how, pdf_dir, save):
    dirPath1 = pdf_dir
    
    list1 = os.listdir(dirPath1) # フォルダ内に存在するファイルを表示
    list1 = [file for file in list1 if file.endswith((".pdf", ".PDF"))]

    order = select_mode() # 並び順の取得
    if order == None:
        None
    else:
        list1 = sort_files_TTC(list1, order) # 並び順通りの並び替え。

    print(dirPath1, list1)
    list2 = [dirPath1+"/"+l2 for l2 in list1] # 絶対path作成

    out1 = "tmp1.pdf"
    out2 = save

    if how == "栞無":
        combine_pdfs(list2, out_pdf=out2) #PDF合体関数

    elif how == "ファイル名栞":
        toc_list = tocmaker(dirPath1, list1)
        combine_pdfs(list2, out_pdf=out1) #PDF合体関数
        list_to_pdf(out1, toc_list, out_pdf=out2) #合体したPDFに目次を追加する関数。    
        
        try: # 中間ファイルの削除
            os.remove(out1) #どこかでファイルを開きっぱなしだと実行できない。
        except:
            None

    elif how == "ファイル名+元栞":
        toc_list = tocmaker_second_toc(dirPath1, list1)
        combine_pdfs(list2, out_pdf=out1) #PDF合体関数
        list_to_pdf(out1, toc_list, out_pdf=out2) #合体したPDFに目次を追加する関数。 
        
        try: # 中間ファイルの削除
            os.remove(out1) #どこかでファイルを開きっぱなしだと実行できない。
        except:
            None

    elif how == "元栞":
        toc_list = tocmaker_original_toc(dirPath1, list1)
        combine_pdfs(list2, out_pdf=out1) #PDF合体関数
        list_to_pdf(out1, toc_list, out_pdf=out2) #合体したPDFに目次を追加する関数。 
        
        try: # 中間ファイルの削除
            os.remove(out1) #どこかでファイルを開きっぱなしだと実行できない。
        except:
            None
    
    else:
        print("なにもしない")


"""
01.PDFから栞のリストを取出す 220401
02.取出したリスト通りに栞を挿入しなおす。220401
03.PDFからPDFへ01をした後02を実行。221110
"""