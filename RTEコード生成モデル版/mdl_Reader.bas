Attribute VB_Name = "mdl_Reader"
Option Explicit

'内部変数
Dim obj_Header As cls_Reader

'簡易テスト処理
Public Sub Test()
    '前処理
    Set obj_Header = New cls_Reader
    '初期化処理
    obj_Header.Init sht_Header
    'テキスト読込み処理
    Do While obj_Header.ReadText
        '確認用の出力
        Debug.Print obj_Header.str_Text
    Loop
    '後処理
    Set obj_Header = Nothing
End Sub
