# Replace-Multiple-Words
Excel（xlsx形式）で作成されたリストに従ってWord（docx形式）のテキストを置換するプログラム\n
ExcelのA列を「検索する文字列」、B列を「置換後の文字列」として読み取り、Wordのテキストを置換します。

# 使い方
1. ExcelのA列に「検索する文字列」、B列に「置換後の文字列」を記入して単語リストを作成
2. プログラムを起動し、「参照」ボタンをクリックしてそれぞれのファイルを選択
3. 「実行」ボタンをクリックして、ファイル名と保存先を指定
4. 処理が完了すると「完了しました」のメッセージと置換件数が表示されるので、ファイルを確認してください。\n置換対象が無かった場合は新規のファイルは保存されません。

# 注意点
* Excelリストでは、選択されているシートのみが読み込まれます。選択されていないシートは読み込まれません。
* Excelリスト内に空白のセルがある場合は、エラーとなります。詳しくは、「Sample_List」の「Error_Sample 1」をご確認ください。
* Excelリストの単語は1行目から順に置換されます。重複や部分一致の単語が無いかご確認の上、実行してください。
* 英語は大文字と小文字を区別し、完全一致の文字列のみ置換します。
* Excelに数式が入っている場合は、値が読み取られます。

# 使用したライブラリ
* tkinter
https://docs.python.org/3/library/tkinter.html
* openpyxl
`pip install openpyxl`
https://pypi.org/project/openpyxl/
* python-docx
`pip install python-docx`
https://pypi.org/project/python-docx/
