Attribute VB_Name = "mdl_Output"
Option Explicit

'内部定数
Private Const INT_ROW_START As Integer = 1
Private Const INT_COL_TEXT As Integer = 1

'内部変数
Private int_Row As Integer

'初期化処理
Public Sub Init()
    'シートのクリア
    sht_Output.Range("A:A").Clear
    '変数の初期化
    int_Row = INT_ROW_START
End Sub

'テキスト書込み処理
Public Sub WriteText(str_Text As String)
    'テキストの書込み
    sht_Output.Cells(int_Row, INT_COL_TEXT) = str_Text
    '次の行へ移動
    int_Row = int_Row + 1
End Sub

'簡易テスト処理
Public Sub Test()
    Dim int_Index As Integer
    '初期化処理
    Init
    '確認用の出力
    For int_Index = 1 To 10
        'テキスト書込み処理
        WriteText "OutputText" & int_Index
    Next
End Sub
