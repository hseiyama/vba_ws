Attribute VB_Name = "mdl_Input"
Option Explicit

'�����萔
Private Const INT_ROW_START As Integer = 4
Private Const INT_COL_MODULE_NAME As Integer = 1
Private Const INT_COL_ATTRIBUTE As Integer = 2
Private Const INT_COL_DATA_TYPE As Integer = 3
Private Const INT_COL_DATA_NAME As Integer = 4
Private Const INT_COL_DESCRIPTION As Integer = 5
Private Const INT_COL_PREFIX As Integer = 6

'�O�����J�ϐ�
Public str_ModuleName As String
Public str_Attribute As String
Public str_DataType As String
Public str_DataName As String
Public str_Description As String
Public str_Prefix As String

'�����ϐ�
Private int_Row As Integer

'����������
Public Sub Init()
    '�ϐ��̏�����
    str_ModuleName = ""
    str_Attribute = ""
    str_DataType = ""
    str_DataName = ""
    str_Description = ""
    str_Prefix = ""
    int_Row = INT_ROW_START
End Sub

'�e�L�X�g�Ǎ��ݏ���
Public Function ReadText() As Boolean
    '�e���ڂ̓Ǎ���
    If sht_Input.Cells(int_Row, INT_COL_MODULE_NAME) <> "" Then
        If sht_Input.Cells(int_Row, INT_COL_MODULE_NAME) <> "��" Then
            str_ModuleName = sht_Input.Cells(int_Row, INT_COL_MODULE_NAME)
        End If
        str_Attribute = sht_Input.Cells(int_Row, INT_COL_ATTRIBUTE)
        str_DataType = sht_Input.Cells(int_Row, INT_COL_DATA_TYPE)
        str_DataName = sht_Input.Cells(int_Row, INT_COL_DATA_NAME)
        str_Description = sht_Input.Cells(int_Row, INT_COL_DESCRIPTION)
        str_Prefix = sht_Input.Cells(int_Row, INT_COL_PREFIX)
        ReadText = True
    Else
        ReadText = False
    End If
    '���̍s�ֈړ�
    int_Row = int_Row + 1
End Function

'�ȈՃe�X�g����
Public Sub Test()
    '����������
    Init
    '�e�L�X�g�Ǎ��ݏ���
    Do While ReadText
        '�m�F�p�̏o��
        Debug.Print str_ModuleName, str_Attribute, str_DataType, str_DataName, str_Description, str_Prefix
    Loop
End Sub
