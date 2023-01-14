Attribute VB_Name = "mdl_RteFile"
Option Explicit

'外部定数
Public Const STR_FILE_TARGET As String = "rte_struct.h"
Public Const STR_ADD_TEXT As String = "#incliude <Rte_types.h>"
Public Const INT_ADD_LINE As Integer = 4

'内部定数
Private Const STR_FILE_TEMP As String = "rte_struct.temp"

'RTEファイル編集処理
Public Function EditRteFile(ByRef str_Path As String) As Boolean
    Dim str_Command As String
    Dim bln_Rcode As Boolean
    
    'インクルード追加処理
    Call AddInclude(str_Path)
    'コマンド実行（ファイル削除）
    str_Command = "DEL " & str_Path & "\" & STR_FILE_TARGET
    bln_Rcode = ExecuteCommand(str_Command)
    If bln_Rcode Then
        'コマンド実行（ファイル名変更）
        str_Command = "REN " & str_Path & "\" & STR_FILE_TEMP & " " & STR_FILE_TARGET
        bln_Rcode = ExecuteCommand(str_Command)
    End If
    '戻り値の設定
    EditRteFile = bln_Rcode
End Function

'インクルード追加処理
Private Sub AddInclude(ByRef str_Path As String)
    Dim int_LineNo As Integer
    Dim str_Context As String
    '前処理
    Open str_Path & "\" & STR_FILE_TARGET For Input As #1
    Open str_Path & "\" & STR_FILE_TEMP For Output As #2
    'インクルードの追加処理
    int_LineNo = 1
    Do Until EOF(1)
        Line Input #1, str_Context
        If int_LineNo = INT_ADD_LINE _
        And str_Context <> STR_ADD_TEXT Then
            Print #2, STR_ADD_TEXT
            int_LineNo = int_LineNo + 1
        End If
        Print #2, str_Context
        int_LineNo = int_LineNo + 1
    Loop
    '後処理
    Close #1
    Close #2
End Sub

'コマンド実行
Private Function ExecuteCommand(ByRef str_Command As String) As Boolean
    Dim obj_WShell As Object
    Dim int_RCode As Integer
    Dim bln_Rcode As Boolean
    '前処理
    Set obj_WShell = CreateObject("WScript.Shell")
    '初期化処理
    bln_Rcode = True
    'コマンドの同期実行
    int_RCode = obj_WShell.Run(Command:="%ComSpec% /c " & str_Command, WindowStyle:=0, WaitOnReturn:=True)
    If int_RCode <> 0 Then
        MsgBox "コマンドの実行に失敗しました。" & vbCrLf & str_Command, _
            vbOKOnly + vbCritical, "コマンド実行"
        bln_Rcode = False
    End If
    '後処理
    Set obj_WShell = Nothing
    '戻り値の設定
    ExecuteCommand = bln_Rcode
End Function
