Attribute VB_Name = "mdl_Output"
Option Explicit

'内部定数
'Private Const INT_ROW_START As Integer = 1
'Private Const INT_COL_TEXT As Integer = 1

'内部変数
'Private int_Row As Integer

'初期化処理
Public Sub Init(ByRef str_FilePath As String)
    'シートのクリア
    'sht_Output.Range("A:A").ClearContents
    '変数の初期化
    'int_Row = INT_ROW_START
    'ファイルのオープン
    Open str_FilePath For Output As #1
End Sub

'終了処理
Public Sub Final()
    'ファイルのクローズ
    Close #1
End Sub

'テキスト書込み処理
Public Sub WriteText(str_Text As String)
    'テキストの書込み
    'sht_Output.Cells(int_Row, INT_COL_TEXT).Value = str_Text
    Print #1, str_Text
    '次の行へ移動
    'int_Row = int_Row + 1
End Sub
