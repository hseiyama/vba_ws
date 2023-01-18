VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_TextWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'外部公開変数
Public int_Line As Integer

'内部変数
Private m_intFileNo As Integer

'初期化処理
Public Sub Init(ByRef str_FilePath As String, ByVal int_FileNo As Integer)
    '変数の初期化
    m_intFileNo = int_FileNo
    int_Line = 1
    'オープン処理
    Open str_FilePath For Output As #m_intFileNo
End Sub

'テキスト書込み処理
Public Sub WriteText(ByRef str_Text As String)
    'テキストの書込み
    Print #m_intFileNo, str_Text
    '次の行へ移動
    int_Line = int_Line + 1
End Sub

'コンストラクタ
Private Sub Class_Initialize()
End Sub

'デストラクタ
Private Sub Class_Terminate()
    'クローズ処理
    Close #m_intFileNo
End Sub
