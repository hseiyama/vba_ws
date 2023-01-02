VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "sht_Input"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'内部定数
Private Const STR_RTE_FILE As String = "rte_model.c"

Private Sub RteInfoFile_Click()
    Dim str_FileName As String
    'RTE情報ファイル名の取得
    str_FileName = Application.GetOpenFilename("Microsoft Excelブック,*.xls?")
    If str_FileName <> "False" Then
        Me.Range(mdl_Input.STR_RNG_RTE_INFO_FILE).Value = str_FileName
    End If
End Sub

Private Sub RteCodePath_Click()
    Dim str_CodePath As String
    'RTEコード生成パスの取得
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = 0 Then
            Exit Sub
        End If
        Me.Range(mdl_Input.STR_RNG_RTE_CODE_FILE).Value = .SelectedItems(1) & "\" & STR_RTE_FILE
    End With
End Sub

Private Sub GetRteInfo_Click()
    Dim int_RetCode As Integer
    'RTE情報取得処理
    Call mdl_Main.GetRteInfomation
    '次処理の推奨
    int_RetCode = MsgBox("RTE情報取得処理が完了しました。" & vbCrLf & _
        "次の処理「コード生成処理」を継続して実行しますか？", _
        vbOKCancel + vbQuestion, "RTE情報取得処理")
    If int_RetCode = vbOK Then
        '次処理「コード生成処理」の実行
        Call GenerateCode_Click
    End If
End Sub

Private Sub GenerateCode_Click()
    'コード生成処理
    Call mdl_Main.GenerateCode
    'コード生成結果を表示
    'sht_Output.Activate
    '実行結果の表示
    MsgBox "コード生成処理が完了しました。" & vbCrLf & _
        Me.Range(mdl_Input.STR_RNG_RTE_CODE_FILE).Value, _
        vbOKOnly + vbInformation, "コード生成処理"
End Sub
