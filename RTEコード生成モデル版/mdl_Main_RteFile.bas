Attribute VB_Name = "mdl_Main_RteFile"
Option Explicit

'外部定数
Public Const STR_RTE_STRUCT As String = "rte_struct.h"
Public Const STR_RTE_TYPE As String = "Rte_Type.h"

'内部定数
Private Const STR_TMP As String = ".tmp"

'内部変数
Private obj_SheetReader As cls_SheetReader

'RTEファイル編集処理(RteStruct)
Public Function EditRteStruct() As Boolean
    'RTEファイル編集処理
    EditRteStruct = EditRteFile(sht_RteStruct, STR_RTE_STRUCT)
End Function

'RTEファイル編集処理(RteType)
Public Function EditRteType() As Boolean
    'RTEファイル編集処理
    EditRteType = EditRteFile(sht_RteType, STR_RTE_TYPE)
End Function

'RTEファイル編集処理
Private Function EditRteFile(ByRef obj_Sheet As Worksheet, ByRef str_File As String) As Boolean
    '前処理
    Set obj_SheetReader = New cls_SheetReader
    '各モジュールの初期化処理
    Call mdl_Input.Init
    Call obj_SheetReader.Init(obj_Sheet)
    '対象ファイル編集処理
    EditRteFile = EditTargetFile(str_File)
    '後処理
    Set obj_SheetReader = Nothing
End Function

'対象ファイル編集処理
Private Function EditTargetFile(ByRef str_File As String) As Boolean
    Dim str_Command As String
    Dim bln_RCode As Boolean
    '対象ファイル確認
    bln_RCode = CheckTargetFile(str_File)
    If bln_RCode Then
        '行挿入処理
        Call InsertLine(str_File)
        'コマンド実行（ファイル削除）
        str_Command = "DEL " & mdl_Input.rng_RteCodePath.Value & "\" & str_File
        bln_RCode = ExecuteCommand(str_Command)
        If bln_RCode Then
            'コマンド実行（ファイル名変更）
            str_Command = "REN " & mdl_Input.rng_RteCodePath.Value & "\" & str_File & STR_TMP & " " & str_File
            bln_RCode = ExecuteCommand(str_Command)
        End If
    End If
    '戻り値の設定
    EditTargetFile = bln_RCode
End Function

'対象ファイル確認
Private Function CheckTargetFile(ByRef str_File As String) As Boolean
    Dim str_FilePath As String
    Dim bln_RCode As Boolean
    '変数の初期化
    bln_RCode = True
    'ファイルの存在を確認
    str_FilePath = mdl_Input.rng_RteCodePath.Value & "\" & str_File
    If Dir(str_FilePath) = "" Then
        MsgBox "編集対象のファイルが存在しません。" & vbCrLf & str_FilePath, _
            vbOKOnly + vbExclamation, "編集対象ファイル確認"
        bln_RCode = False
    End If
    '戻り値の設定
    CheckTargetFile = bln_RCode
End Function

'行挿入処理
Private Sub InsertLine(ByRef str_File As String)
    Dim obj_TextReader As cls_TextReader
    Dim obj_TextWriter As cls_TextWriter
    Dim bln_NextLoop As Boolean
    '前処理
    Set obj_TextReader = New cls_TextReader
    Set obj_TextWriter = New cls_TextWriter
    '初期化処理
    Call obj_TextReader.Init(mdl_Input.rng_RteCodePath.Value & "\" & str_File, 1)
    Call obj_TextWriter.Init(mdl_Input.rng_RteCodePath.Value & "\" & str_File & STR_TMP, 2)
    'テキスト読込み処理（挿入部）
    Call obj_SheetReader.ReadInsert
    '行の追加処理
    Do While obj_TextReader.ReadText
        '挿入する情報が尽きるまで検索
        bln_NextLoop = True
        Do While bln_NextLoop
            bln_NextLoop = False
            '挿入する情報を判定
            If obj_SheetReader.bln_State _
            And obj_SheetReader.int_Line = obj_TextWriter.int_Line Then
                If obj_SheetReader.str_Text <> obj_TextReader.str_Text Then
                    'テキスト書込み処理
                    Call obj_TextWriter.WriteText(obj_SheetReader.str_Text)
                End If
                'テキスト読込み処理（挿入部）
                Call obj_SheetReader.ReadInsert
                bln_NextLoop = True
            End If
        Loop
        'テキスト書込み処理
        Call obj_TextWriter.WriteText(obj_TextReader.str_Text)
    Loop
    '後処理
    Set obj_TextReader = Nothing
    Set obj_TextWriter = Nothing
End Sub

'コマンド実行
Private Function ExecuteCommand(ByRef str_Command As String) As Boolean
    Dim obj_WShell As Object
    Dim int_RCode As Integer
    Dim bln_RCode As Boolean
    '前処理
    Set obj_WShell = CreateObject("WScript.Shell")
    '変数の初期化
    bln_RCode = True
    'コマンドの同期実行
    int_RCode = obj_WShell.Run(Command:="%ComSpec% /c " & str_Command, WindowStyle:=0, WaitOnReturn:=True)
    If int_RCode <> 0 Then
        MsgBox "コマンド実行に失敗しました。" & vbCrLf & str_Command, _
            vbOKOnly + vbCritical, "コマンド実行"
        bln_RCode = False
    End If
    '後処理
    Set obj_WShell = Nothing
    '戻り値の設定
    ExecuteCommand = bln_RCode
End Function
