# pdf_csv_bookmarker

main.pyを実行するとGUI起動。

1. PDF-csv
   1.1. select(PDF to csv): 栞付きPDFから栞csvデータの抜出。
   1.2. select(csv to PDF): 栞csvデータからPDFへの栞情報の書出。
   1.3. select(new csv):    初期栞ファイルの作成

2. PDF-PDF
   栞付きPDFからPDFへの栞のコピー

3. Merge
   PDFとPDFの合体。
   ファイルの並び順は名前の順。
   特に入れ替えたい要素ある際、
   init.jsonのterm1, term2を書き替える。

4. Option
   CD変更: ファイル選択ダイアログで選択する際の開始位置を変更するために用意した。
           大量のPDFに栞を付ける際に使用する。

以上
