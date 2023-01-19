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
    Call mdl_Main_RteInfo.Collect
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
    Call mdl_Main_RteModel.Generate
    '実行結果の表示（GenerateCode）
    Call GenerateCode_CompleteMsg(mdl_Main_RteModel.STR_RTE_FILE)
    'コード生成処理
    Call mdl_Main_RteCompiler.Generate
    '実行結果の表示（GenerateCode）
    Call GenerateCode_CompleteMsg(mdl_Main_RteCompiler.STR_RTE_FILE)
End Sub

'実行結果の表示（GenerateCode）
Private Sub GenerateCode_CompleteMsg(ByRef str_File As String)
    Const STR_MSG_TITLE As String = "コード生成処理"
    MsgBox "コード生成処理が完了しました。" & vbCrLf & _
        Me.Range(mdl_Input.STR_RNG_RTE_CODE_PATH).Value & "\" & str_File, _
        vbOKOnly + vbInformation, STR_MSG_TITLE
End Sub

Private Sub EditRteFile_Click()
    Const STR_MSG_TITLE As String = "既存RTEファイル編集処理"
    Dim str_FilePath As String
    Dim bln_RCode As Boolean
    '処理実行の確認(RteStruct)
    str_FilePath = Me.Range(mdl_Input.STR_RNG_RTE_CODE_PATH).Value & "\" & mdl_Main_RteFile.STR_RTE_STRUCT
    '実行確認の表示（EditRteFile）
    If EditRteFile_ConfirmMsg(STR_MSG_TITLE, str_FilePath) Then
        'RTEファイル編集処理(RteStruct)
        bln_RCode = mdl_Main_RteFile.EditRteStruct()
        If bln_RCode Then
            '実行結果の表示（EditRteFile）
            Call EditRteFile_CompleteMsg(STR_MSG_TITLE, str_FilePath)
        End If
    End If
    '処理実行の確認(RteType)
    str_FilePath = Me.Range(mdl_Input.STR_RNG_RTE_CODE_PATH).Value & "\" & mdl_Main_RteFile.STR_RTE_TYPE
    '実行確認の表示（EditRteFile）
    If EditRteFile_ConfirmMsg(STR_MSG_TITLE, str_FilePath) Then
        'RTEファイル編集処理(RteType)
        bln_RCode = mdl_Main_RteFile.EditRteType()
        If bln_RCode Then
            '実行結果の表示（EditRteFile）
            Call EditRteFile_CompleteMsg(STR_MSG_TITLE, str_FilePath)
        End If
    End If
End Sub

'実行確認の表示（EditRteFile）
Private Function EditRteFile_ConfirmMsg(ByRef str_MsgTitle As String, ByRef str_FilePath As String) As Boolean
    Dim int_RetCode As Integer
    Dim bln_RCode As Boolean
    '実行確認の表示
    int_RetCode = MsgBox("既存RTEファイルを編集しますか？" & vbCrLf & str_FilePath, _
        vbOKCancel + vbQuestion, str_MsgTitle)
    If int_RetCode = vbOK Then
        bln_RCode = True
    Else
        bln_RCode = False
    End If
    '戻り値の設定
    EditRteFile_ConfirmMsg = bln_RCode
End Function

'実行結果の表示（EditRteFile）
Private Sub EditRteFile_CompleteMsg(ByRef str_MsgTitle As String, ByRef str_FilePath As String)
    '実行結果の表示
    MsgBox "既存RTEファイル編集処理が完了しました。" & vbCrLf & str_FilePath, _
        vbOKOnly + vbInformation, str_MsgTitle
End Sub
