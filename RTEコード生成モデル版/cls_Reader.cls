VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_Reader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'内部定数
Private Const INT_ROW_START As Integer = 2
Private Const INT_COL_CHECK As Integer = 1
Private Const INT_COL_TEXT As Integer = 2

'外部公開変数
Public str_Text As String

'内部変数
Private m_objSheet As Worksheet
Private m_intRow As Integer

'初期化処理
Public Sub Init(ByRef obj_Sheet As Worksheet)
    '変数の初期化
    Set m_objSheet = obj_Sheet
    m_intRow = INT_ROW_START
End Sub

'テキスト読込み処理
Public Function ReadText() As Boolean
    '各項目の読込み
    If m_objSheet.Cells(m_intRow, INT_COL_CHECK) <> "" Then
        str_Text = m_objSheet.Cells(m_intRow, INT_COL_TEXT)
        ReadText = True
    Else
        ReadText = False
    End If
    '次の行へ移動
    m_intRow = m_intRow + 1
End Function

'コンストラクタ
Private Sub Class_Initialize()
End Sub

'デストラクタ
Private Sub Class_Terminate()
    Set m_objSheet = Nothing
End Sub
