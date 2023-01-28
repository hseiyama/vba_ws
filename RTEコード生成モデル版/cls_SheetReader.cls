VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_SheetReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'内部定数
Private Const INT_ROW_START As Integer = 2
Private Const INT_COL_CHECK As Integer = 1
Private Const INT_COL_LINE As Integer = 2
Private Const INT_COL_TEXT As Integer = 3
Private Const STR_HEADER As String = "H"
Private Const STR_FOOTER As String = "F"
Private Const STR_INSERT As String = "I"

'外部公開変数
Public bln_State As Boolean
Public str_Text As String
Public int_Line As Integer

'内部変数
Private m_objSheet As Worksheet
Private m_intRowHeader As Integer
Private m_intRowFooter As Integer
Private m_intRowInsert As Integer

'初期化処理
Public Sub Init(ByRef obj_Sheet As Worksheet)
    '変数の初期化
    Set m_objSheet = obj_Sheet
    m_intRowHeader = INT_ROW_START
    m_intRowFooter = INT_ROW_START
    m_intRowInsert = INT_ROW_START
    bln_State = False
    str_Text = ""
    int_Line = 0
End Sub

'テキスト読込み処理（ヘッダー部）
Public Function ReadHeader() As Boolean
    ReadHeader = ReadText(STR_HEADER, m_intRowHeader)
End Function

'テキスト読込み処理（フッター部）
Public Function ReadFooter() As Boolean
    ReadFooter = ReadText(STR_FOOTER, m_intRowFooter)
End Function

'テキスト読込み処理（挿入部）
Public Function ReadInsert() As Boolean
    ReadInsert = ReadText(STR_INSERT, m_intRowInsert)
End Function

'テキスト読込み処理
Private Function ReadText(ByRef str_Check As String, ByRef int_Row As Integer) As Boolean
    bln_State = False
    '該当テキストの検索
    Do While m_objSheet.Cells(int_Row, INT_COL_CHECK).Value <> "" _
    And Not bln_State
        '各項目の読込み
        If m_objSheet.Cells(int_Row, INT_COL_CHECK).Value = str_Check Then
            str_Text = m_objSheet.Cells(int_Row, INT_COL_TEXT).Value
            If str_Check = STR_INSERT Then
                int_Line = Int(m_objSheet.Cells(int_Row, INT_COL_LINE).Value)
            End If
            bln_State = True
        End If
        '次の行へ移動
        int_Row = int_Row + 1
    Loop
    '戻り値の設定
    ReadText = bln_State
End Function

'コンストラクタ
Private Sub Class_Initialize()
End Sub

'デストラクタ
Private Sub Class_Terminate()
    Set m_objSheet = Nothing
End Sub
