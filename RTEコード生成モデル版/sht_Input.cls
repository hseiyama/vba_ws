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
        Me.Range(mdl_Input.STR_RNG_RTE_CODE_PATH).Value = .SelectedItems(1)
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
        Me.Range(mdl_Input.STR_RNG_RTE_CODE_PATH).Value & "\" & mdl_Main.STR_RTE_FILE, _
        vbOKOnly + vbInformation, "コード生成処理"
End Sub

Private Sub EditRteFile_Click()
    Dim bln_Rcode As Boolean
    '既存RTEファイル編集処理
    bln_Rcode = EditRteExistFile()
    If bln_Rcode Then
    '実行結果の表示
    MsgBox "既存RTEファイル編集処理が完了しました。" & vbCrLf & _
        Me.Range(mdl_Input.STR_RNG_RTE_CODE_PATH).Value & "\" & mdl_RteFile.STR_FILE_TARGET & vbCrLf & _
        "追加したインクルード (" & mdl_RteFile.INT_ADD_LINE & "行目)" & vbCrLf & _
        mdl_RteFile.STR_ADD_TEXT, _
        vbOKOnly + vbInformation, "既存RTEファイル編集処理"
    End If
End Sub
