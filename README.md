# Replace-Multiple-Words
## Overview
Excel（xlsx形式）で作成されたリストに従ってWord（docx形式）のテキストを置換するプログラムです。
ExcelのA列を「検索する文字列」、B列を「置換後の文字列」として読み取り、Wordのテキストを置換します。

## Usage
1. ExcelのA列に「検索する文字列」、B列に「置換後の文字列」を記入して単語リストを作成
2. プログラムを起動し、「参照」ボタンをクリックしてそれぞれのファイルを選択
3. 「実行」ボタンをクリックして、ファイル名と保存先を指定
4. 処理が完了すると「完了しました」のメッセージと置換件数が表示されるので、ファイルを確認してください。置換対象が無かった場合は新規のファイルは保存されません。

## Note 1: Excel
* Excelリストでは、選択されているシートのみが読み込まれます。選択されていないシートは読み込まれません。
* Excelリスト内に空白のセルが検出された場合は、エラーメッセージが出て処理を中止します（具体例：「Sample_List」内「Error_Sample 1」シート）。
* Excelリストの単語は1行目から順に置換されます。重複や部分一致の単語が無いかご確認の上、実行してください。
* 英語は大文字と小文字を区別し、完全一致の文字列のみ置換します。
* Excelに数式が入っている場合は、値が読み取られます。

## Note 2: Word
* 文書内の埋め込みExcel内の文字列は置換されません（具体例：「Sample_Target.docx」）。

## Requirements
* tkinter
https://docs.python.org/3/library/tkinter.html
* openpyxl
`pip install openpyxl`
* python-docx
`pip install python-docx`
https://pypi.org/project/python-docx/

## Open Source Licenses
* tkinter  
https://docs.python.org/3/license.html
* openpyxl  
https://foss.heptapod.net/openpyxl/openpyxl/-/blob/branch/3.0/LICENCE.rst
* python-docx  
https://github.com/python-openxml/python-docx/blob/master/LICENSE
