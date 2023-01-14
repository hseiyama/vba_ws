Attribute VB_Name = "mdl_Input"
Option Explicit

'�O���萔
Public Const STR_RNG_RTE_INFO_FILE As String = "B1"
Public Const STR_RNG_RTE_CODE_PATH As String = "B2"

'�����萔
Private Const INT_ROW_START As Integer = 6
Private Const INT_COL_MODULE_NAME As Integer = 1
Private Const INT_COL_ATTRIBUTE As Integer = 2
Private Const INT_COL_DATA_TYPE As Integer = 3
Private Const INT_COL_DATA_NAME As Integer = 4
Private Const INT_COL_DESCRIPTION As Integer = 5
Private Const INT_COL_PREFIX As Integer = 6
Private Const STR_RNG_RTE_INFO_LIST As String = "A" & INT_ROW_START & ":E1048576"

'�O�����J�ϐ�
Public rng_RteInfoFile As Range
Public rng_RteCodePath As Range
Public rng_RteInfoList As Range
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
    Set rng_RteInfoFile = sht_Input.Range(STR_RNG_RTE_INFO_FILE)
    Set rng_RteCodePath = sht_Input.Range(STR_RNG_RTE_CODE_PATH)
    Set rng_RteInfoList = sht_Input.Range(STR_RNG_RTE_INFO_LIST)
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
    If sht_Input.Cells(int_Row, INT_COL_MODULE_NAME).Value <> "" Then
        If sht_Input.Cells(int_Row, INT_COL_MODULE_NAME).Value <> "��" Then
            str_ModuleName = sht_Input.Cells(int_Row, INT_COL_MODULE_NAME).Value
        End If
        str_Attribute = sht_Input.Cells(int_Row, INT_COL_ATTRIBUTE).Value
        str_DataType = sht_Input.Cells(int_Row, INT_COL_DATA_TYPE).Value
        str_DataName = sht_Input.Cells(int_Row, INT_COL_DATA_NAME).Value
        str_Description = sht_Input.Cells(int_Row, INT_COL_DESCRIPTION).Value
        str_Prefix = sht_Input.Cells(int_Row, INT_COL_PREFIX).Value
        ReadText = True
    Else
        ReadText = False
    End If
    '���̍s�ֈړ�
    int_Row = int_Row + 1
End Function
