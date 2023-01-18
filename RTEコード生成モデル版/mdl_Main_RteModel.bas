Attribute VB_Name = "mdl_Main_RteModel"
Option Explicit

'�O���萔
Public Const STR_RTE_FILE As String = "rte_model.c"

'�����萔
Private Const STR_PREFIX_BUS As String = "bus"

'�����ϐ�
Private obj_SheetReader As cls_SheetReader
Private obj_TextWriter As cls_TextWriter

'�R�[�h��������
Public Sub Generate()
    '����������
    Call Initialize
    '�R�[�h���������i�w�b�_�[���j
    Call CopyCodeHeader
    'RTE�֐��쐬����
    Call MakeRteFunc
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
    Call obj_SheetReader.Init(sht_RteModel)
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
        Call obj_TextWriter.WriteText("/* " & str_Comment & " */")
        Call obj_TextWriter.WriteText("SdtType " & str_FnucText & " {")
        Call obj_TextWriter.WriteText("    " & str_MacroText & ";")
        Call obj_TextWriter.WriteText("    return STD_OK;")
        Call obj_TextWriter.WriteText("}")
        Call obj_TextWriter.WriteText("")
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
    If mdl_Input.str_Attribute = mdl_Main_RteInfo.STR_ATTRIB_READ Then
        str_Command = "Read"
        str_Param = "*u"
    ElseIf mdl_Input.str_Attribute = mdl_Main_RteInfo.STR_ATTRIB_WRITE Then
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
    If mdl_Input.str_Attribute = mdl_Main_RteInfo.STR_ATTRIB_READ Then
        str_Command = "read"
        str_Param = "u"
    ElseIf mdl_Input.str_Attribute = mdl_Main_RteInfo.STR_ATTRIB_WRITE Then
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
