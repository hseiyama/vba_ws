VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_TextReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'外部公開変数
Public bln_State As Boolean
Public str_Text As String
Public int_Line As Integer

'内部変数
Private m_intFileNo As Integer

'初期化処理
Public Sub Init(ByRef str_FilePath As String, ByVal int_FileNo As Integer)
    '変数の初期化
    m_intFileNo = int_FileNo
    bln_State = False
    str_Text = ""
    int_Line = 0
    'オープン処理
    Open str_FilePath For Input As #m_intFileNo
End Sub

'テキスト読込み処理
Public Function ReadText() As Boolean
    bln_State = False
    'テキスト終端の判定
    If Not EOF(m_intFileNo) Then
        bln_State = True
        'テキストの読込み
        Line Input #m_intFileNo, str_Text
        int_Line = int_Line + 1
    End If
    '戻り値の設定
    ReadText = bln_State
End Function

'コンストラクタ
Private Sub Class_Initialize()
End Sub

'デストラクタ
Private Sub Class_Terminate()
    'クローズ処理
    Close #m_intFileNo
End Sub
