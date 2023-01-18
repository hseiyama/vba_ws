Attribute VB_Name = "mdl_Main_RteCompiler"
Option Explicit

'�O���萔
Public Const STR_RTE_FILE As String = "Rte_Compiler_Cfg.h"

'�����ϐ�
Private obj_SheetReader As cls_SheetReader
Private obj_TextWriter As cls_TextWriter

'�R�[�h��������
Public Sub Generate()
    '����������
    Call Initialize
    '�R�[�h���������i�w�b�_�[���j
    Call CopyCodeHeader
    'DEF�쐬����
    Call MakeDefine
    '�R�[�h���������i�t�b�^�[���j
    Call CopyCodeFooter
    '�I������
    Call Finalize
End Sub

'����������
Private Sub Initialize()
    '�O����
    Set obj_SheetReader = New cls_SheetReader
    Set obj_TextWriter = New cls_TextWriter
    '�e���W���[���̏���������
    Call mdl_Input.Init
    Call obj_TextWriter.Init(mdl_Input.rng_RteCodePath.Value & "\" & STR_RTE_FILE, 1)
    Call obj_SheetReader.Init(sht_RteCompiler)
End Sub

'�I������
Private Sub Finalize()
    '�㏈��
    Set obj_SheetReader = Nothing
    Set obj_TextWriter = Nothing
End Sub

'�R�[�h���������i�w�b�_�[���j
Private Sub CopyCodeHeader()
    '�e�L�X�g�Ǎ��ݏ���
    Do While obj_SheetReader.ReadHeader
        '�e�L�X�g�����ݏ���
        Call obj_TextWriter.WriteText(obj_SheetReader.str_Text)
    Loop
End Sub

'�R�[�h���������i�t�b�^�[���j
Private Sub CopyCodeFooter()
    '�e�L�X�g�Ǎ��ݏ���
    Do While obj_SheetReader.ReadFooter
        '�e�L�X�g�����ݏ���
        Call obj_TextWriter.WriteText(obj_SheetReader.str_Text)
    Loop
End Sub

'DEF�쐬����
Private Sub MakeDefine()
    Dim str_ModuleNamePre As String
    '�ϐ��̏�����
    str_ModuleNamePre = ""
    'RTE�֐��̍쐬
    Do While mdl_Input.ReadText
        If mdl_Input.str_ModuleName <> str_ModuleNamePre Then
            '�e�L�X�g�����ݏ���
            Call obj_TextWriter.WriteText("/* " & mdl_Input.str_ModuleName & "��` */")
            Call obj_TextWriter.WriteText("#define " & mdl_Input.str_ModuleName & "_CODE")
        End If
        str_ModuleNamePre = mdl_Input.str_ModuleName
    Loop
End Sub
