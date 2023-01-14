Attribute VB_Name = "mdl_RteFile"
Option Explicit

'�O���萔
Public Const STR_FILE_TARGET As String = "rte_struct.h"
Public Const STR_ADD_TEXT As String = "#incliude <Rte_types.h>"
Public Const INT_ADD_LINE As Integer = 4

'�����萔
Private Const STR_FILE_TEMP As String = "rte_struct.temp"

'RTE�t�@�C���ҏW����
Public Function EditRteFile(ByRef str_Path As String) As Boolean
    Dim str_Command As String
    Dim bln_Rcode As Boolean
    
    '�C���N���[�h�ǉ�����
    Call AddInclude(str_Path)
    '�R�}���h���s�i�t�@�C���폜�j
    str_Command = "DEL " & str_Path & "\" & STR_FILE_TARGET
    bln_Rcode = ExecuteCommand(str_Command)
    If bln_Rcode Then
        '�R�}���h���s�i�t�@�C�����ύX�j
        str_Command = "REN " & str_Path & "\" & STR_FILE_TEMP & " " & STR_FILE_TARGET
        bln_Rcode = ExecuteCommand(str_Command)
    End If
    '�߂�l�̐ݒ�
    EditRteFile = bln_Rcode
End Function

'�C���N���[�h�ǉ�����
Private Sub AddInclude(ByRef str_Path As String)
    Dim int_LineNo As Integer
    Dim str_Context As String
    '�O����
    Open str_Path & "\" & STR_FILE_TARGET For Input As #1
    Open str_Path & "\" & STR_FILE_TEMP For Output As #2
    '�C���N���[�h�̒ǉ�����
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
    '�㏈��
    Close #1
    Close #2
End Sub

'�R�}���h���s
Private Function ExecuteCommand(ByRef str_Command As String) As Boolean
    Dim obj_WShell As Object
    Dim int_RCode As Integer
    Dim bln_Rcode As Boolean
    '�O����
    Set obj_WShell = CreateObject("WScript.Shell")
    '����������
    bln_Rcode = True
    '�R�}���h�̓������s
    int_RCode = obj_WShell.Run(Command:="%ComSpec% /c " & str_Command, WindowStyle:=0, WaitOnReturn:=True)
    If int_RCode <> 0 Then
        MsgBox "�R�}���h�̎��s�Ɏ��s���܂����B" & vbCrLf & str_Command, _
            vbOKOnly + vbCritical, "�R�}���h���s"
        bln_Rcode = False
    End If
    '�㏈��
    Set obj_WShell = Nothing
    '�߂�l�̐ݒ�
    ExecuteCommand = bln_Rcode
End Function
