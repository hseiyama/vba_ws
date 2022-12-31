Attribute VB_Name = "mdl_Main"
Option Explicit

'�����萔
Private Const STR_ATTRIB_READ As String = "SWC_I/F����"
Private Const STR_ATTRIB_WRITE As String = "SWC_I/F�o��"
Private Const STR_PREFIX_BUS As String = "bus"

'�����ϐ�
Dim obj_Header As cls_Reader
Dim obj_Footer As cls_Reader

'�R�[�h��������
Public Sub GenerateCode()
    '�O����
    Set obj_Header = New cls_Reader
    Set obj_Footer = New cls_Reader
    '����������
    Initialize
    '�R�[�h�R�s�[�����i�w�b�_�[���j
    CopyCode obj_Header
    'RTE�֐��쐬����
    MakeRteFunc
    '�R�[�h�R�s�[�����i�t�b�^�[���j
    CopyCode obj_Footer
    '�㏈��
    Set obj_Header = Nothing
    Set obj_Footer = Nothing
End Sub

'����������
Private Sub Initialize()
    '�e���W���[���̏���������
    mdl_Input.Init
    mdl_Output.Init
    obj_Header.Init sht_Header
    obj_Footer.Init sht_Footer
End Sub

'�R�[�h�R�s�[����
Private Sub CopyCode(ByRef obj_Reader As cls_Reader)
    '�e�L�X�g�Ǎ��ݏ���
    Do While obj_Reader.ReadText
        '�e�L�X�g�����ݏ���
        mdl_Output.WriteText obj_Reader.str_Text
    Loop
End Sub

'RTE�֐��쐬����
Private Sub MakeRteFunc()
    Dim str_Comment As String
    Dim str_FnucText As String
    Dim str_MacroText As String
    'RTE�֐��̍쐬
    Do While mdl_Input.ReadText
        '�e������̍쐬
        str_Comment = MakeComment
        str_FnucText = MakeFnucText
        str_MacroText = MakeMacroText
        '�e�L�X�g�����ݏ���
        mdl_Output.WriteText "/* " & str_Comment & " */"
        mdl_Output.WriteText "SdtType " & str_FnucText & " {"
        mdl_Output.WriteText "    " & str_MacroText & ";"
        mdl_Output.WriteText "    return STD_OK;"
        mdl_Output.WriteText "}"
        mdl_Output.WriteText ""
    Loop
End Sub

'�R�����g�쐬
Private Function MakeComment() As String
    MakeComment = mdl_Input.str_ModuleName & "(" & mdl_Input.str_DataName & ")"
End Function

'�֐�������쐬
Private Function MakeFnucText() As String
    Dim str_Command As String
    Dim str_Param As String
    '�e������̍쐬
    If mdl_Input.str_Attribute = STR_ATTRIB_READ Then
        str_Command = "Read"
        str_Param = "*u"
    ElseIf mdl_Input.str_Attribute = STR_ATTRIB_WRITE Then
        str_Command = "Write"
        If mdl_Input.str_Prefix = STR_PREFIX_BUS Then
            str_Param = "*u"
        Else
            str_Param = "u"
        End If
    Else
        str_Command = "Unknown"
        str_Param = "u"
    End If
    '�֐�������̍쐬
    MakeFnucText = mdl_Input.str_ModuleName & "_" & _
                str_Command & "_" & _
                mdl_Input.str_Prefix & "_g_" & mdl_Input.str_DataName & _
                mdl_Input.str_Prefix & "_g_" & mdl_Input.str_DataName & _
                "(" & mdl_Input.str_DataType & " " & str_Param & ")"
End Function

'�}�N��������쐬
Private Function MakeMacroText() As String
    Dim str_Command As String
    Dim str_Param As String
    '�e������̍쐬
    If mdl_Input.str_Attribute = STR_ATTRIB_READ Then
        str_Command = "read"
        str_Param = "u"
    ElseIf mdl_Input.str_Attribute = STR_ATTRIB_WRITE Then
        str_Command = "write"
        If mdl_Input.str_Prefix = STR_PREFIX_BUS Then
            str_Param = "*u"
        Else
            str_Param = "u"
        End If
    Else
        str_Command = "unknown"
        str_Param = "u"
    End If
    '�}�N��������̍쐬
    MakeMacroText = LCase(mdl_Input.str_ModuleName) & "_" & _
                str_Command & "_" & _
                mdl_Input.str_DataName & _
                "(" & str_Param & ")"
End Function
