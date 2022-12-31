Attribute VB_Name = "mdl_Input"
Option Explicit

'内部定数
Private Const INT_ROW_START As Integer = 4
Private Const INT_COL_MODULE_NAME As Integer = 1
Private Const INT_COL_ATTRIBUTE As Integer = 2
Private Const INT_COL_DATA_TYPE As Integer = 3
Private Const INT_COL_DATA_NAME As Integer = 4
Private Const INT_COL_DESCRIPTION As Integer = 5
Private Const INT_COL_PREFIX As Integer = 6

'外部公開変数
Public str_ModuleName As String
Public str_Attribute As String
Public str_DataType As String
Public str_DataName As String
Public str_Description As String
Public str_Prefix As String

'内部変数
Private int_Row As Integer

'初期化処理
Public Sub Init()
    '変数の初期化
    str_ModuleName = ""
    str_Attribute = ""
    str_DataType = ""
    str_DataName = ""
    str_Description = ""
    str_Prefix = ""
    int_Row = INT_ROW_START
End Sub

'テキスト読込み処理
Public Function ReadText() As Boolean
    '各項目の読込み
    If sht_Input.Cells(int_Row, INT_COL_MODULE_NAME) <> "" Then
        If sht_Input.Cells(int_Row, INT_COL_MODULE_NAME) <> "↑" Then
            str_ModuleName = sht_Input.Cells(int_Row, INT_COL_MODULE_NAME)
        End If
        str_Attribute = sht_Input.Cells(int_Row, INT_COL_ATTRIBUTE)
        str_DataType = sht_Input.Cells(int_Row, INT_COL_DATA_TYPE)
        str_DataName = sht_Input.Cells(int_Row, INT_COL_DATA_NAME)
        str_Description = sht_Input.Cells(int_Row, INT_COL_DESCRIPTION)
        str_Prefix = sht_Input.Cells(int_Row, INT_COL_PREFIX)
        ReadText = True
    Else
        ReadText = False
    End If
    '次の行へ移動
    int_Row = int_Row + 1
End Function

'簡易テスト処理
Public Sub Test()
    '初期化処理
    Init
    'テキスト読込み処理
    Do While ReadText
        '確認用の出力
        Debug.Print str_ModuleName, str_Attribute, str_DataType, str_DataName, str_Description, str_Prefix
    Loop
End Sub
