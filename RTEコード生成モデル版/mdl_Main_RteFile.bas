Attribute VB_Name = "mdl_Main_RteFile"
Option Explicit

'�O���萔
Public Const STR_RTE_STRUCT As String = "rte_struct.h"
Public Const STR_RTE_TYPE As String = "Rte_Type.h"

'�����萔
Private Const STR_TMP As String = ".tmp"

'�����ϐ�
Private obj_SheetReader As cls_SheetReader

'RTE�t�@�C���ҏW����(RteStruct)
Public Function EditRteStruct() As Boolean
    'RTE�t�@�C���ҏW����
    EditRteStruct = EditRteFile(sht_RteStruct, STR_RTE_STRUCT)
End Function

'RTE�t�@�C���ҏW����(RteType)
Public Function EditRteType() As Boolean
    'RTE�t�@�C���ҏW����
    EditRteType = EditRteFile(sht_RteType, STR_RTE_TYPE)
End Function

'RTE�t�@�C���ҏW����
Private Function EditRteFile(ByRef obj_Sheet As Worksheet, ByRef str_File As String) As Boolean
    '�O����
    Set obj_SheetReader = New cls_SheetReader
    '�e���W���[���̏���������
    Call mdl_Input.Init
    Call obj_SheetReader.Init(obj_Sheet)
    '�Ώۃt�@�C���ҏW����
    EditRteFile = EditTargetFile(str_File)
    '�㏈��
    Set obj_SheetReader = Nothing
End Function

'�Ώۃt�@�C���ҏW����
Private Function EditTargetFile(ByRef str_File As String) As Boolean
    Dim str_Command As String
    Dim bln_RCode As Boolean
    '�Ώۃt�@�C���m�F
    bln_RCode = CheckTargetFile(str_File)
    If bln_RCode Then
        '�s�}������
        Call InsertLine(str_File)
        '�R�}���h���s�i�t�@�C���폜�j
        str_Command = "DEL " & mdl_Input.rng_RteCodePath.Value & "\" & str_File
        bln_RCode = ExecuteCommand(str_Command)
        If bln_RCode Then
            '�R�}���h���s�i�t�@�C�����ύX�j
            str_Command = "REN " & mdl_Input.rng_RteCodePath.Value & "\" & str_File & STR_TMP & " " & str_File
            bln_RCode = ExecuteCommand(str_Command)
        End If
    End If
    '�߂�l�̐ݒ�
    EditTargetFile = bln_RCode
End Function

'�Ώۃt�@�C���m�F
Private Function CheckTargetFile(ByRef str_File As String) As Boolean
    Dim str_FilePath As String
    Dim bln_RCode As Boolean
    '�ϐ��̏�����
    bln_RCode = True
    '�t�@�C���̑��݂��m�F
    str_FilePath = mdl_Input.rng_RteCodePath.Value & "\" & str_File
    If Dir(str_FilePath) = "" Then
        MsgBox "�ҏW�Ώۂ̃t�@�C�������݂��܂���B" & vbCrLf & str_FilePath, _
            vbOKOnly + vbExclamation, "�ҏW�Ώۃt�@�C���m�F"
        bln_RCode = False
    End If
    '�߂�l�̐ݒ�
    CheckTargetFile = bln_RCode
End Function

'�s�}������
Private Sub InsertLine(ByRef str_File As String)
    Dim obj_TextReader As cls_TextReader
    Dim obj_TextWriter As cls_TextWriter
    Dim bln_NextLoop As Boolean
    '�O����
    Set obj_TextReader = New cls_TextReader
    Set obj_TextWriter = New cls_TextWriter
    '����������
    Call obj_TextReader.Init(mdl_Input.rng_RteCodePath.Value & "\" & str_File, 1)
    Call obj_TextWriter.Init(mdl_Input.rng_RteCodePath.Value & "\" & str_File & STR_TMP, 2)
    '�e�L�X�g�Ǎ��ݏ����i�}�����j
    Call obj_SheetReader.ReadInsert
    '�s�̒ǉ�����
    Do While obj_TextReader.ReadText
        '�}�������񂪐s����܂Ō���
        bln_NextLoop = True
        Do While bln_NextLoop
            bln_NextLoop = False
            '�}��������𔻒�
            If obj_SheetReader.bln_State _
            And obj_SheetReader.int_Line = obj_TextWriter.int_Line Then
                If obj_SheetReader.str_Text <> obj_TextReader.str_Text Then
                    '�e�L�X�g�����ݏ���
                    Call obj_TextWriter.WriteText(obj_SheetReader.str_Text)
                End If
                '�e�L�X�g�Ǎ��ݏ����i�}�����j
                Call obj_SheetReader.ReadInsert
                bln_NextLoop = True
            End If
        Loop
        '�e�L�X�g�����ݏ���
        Call obj_TextWriter.WriteText(obj_TextReader.str_Text)
    Loop
    '�㏈��
    Set obj_TextReader = Nothing
    Set obj_TextWriter = Nothing
End Sub

'�R�}���h���s
Private Function ExecuteCommand(ByRef str_Command As String) As Boolean
    Dim obj_WShell As Object
    Dim int_RCode As Integer
    Dim bln_RCode As Boolean
    '�O����
    Set obj_WShell = CreateObject("WScript.Shell")
    '�ϐ��̏�����
    bln_RCode = True
    '�R�}���h�̓������s
    int_RCode = obj_WShell.Run(Command:="%ComSpec% /c " & str_Command, WindowStyle:=0, WaitOnReturn:=True)
    If int_RCode <> 0 Then
        MsgBox "�R�}���h���s�Ɏ��s���܂����B" & vbCrLf & str_Command, _
            vbOKOnly + vbCritical, "�R�}���h���s"
        bln_RCode = False
    End If
    '�㏈��
    Set obj_WShell = Nothing
    '�߂�l�̐ݒ�
    ExecuteCommand = bln_RCode
End Function
