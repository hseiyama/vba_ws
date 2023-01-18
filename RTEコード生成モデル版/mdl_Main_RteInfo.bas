Attribute VB_Name = "mdl_Main_RteInfo"
Option Explicit

'�O���萔
Public Const STR_ATTRIB_READ As String = "SWC_I/F����"
Public Const STR_ATTRIB_WRITE As String = "SWC_I/F�o��"

'�����萔
Private Const INT_ROW_START As Integer = 3
Private Const INT_COL_ATTRIBUTE As Integer = 3
Private Const INT_COL_DATA_TYPE As Integer = 4
Private Const INT_COL_DATA_NAME As Integer = 5
Private Const INT_COL_DESCRIPTION As Integer = 6
Private Const STR_RNG_SHEET_CHECK As String = "E1"
Private Const STR_RNG_MODULE_NAME As String = "D2"
Private Const STR_SHEET_CHECK As String = "SWC_I/F���"

'�����ϐ�
Private int_Range_Row As Integer
Private int_Row As Integer
Private bln_First As Boolean

'RTE���擾����
Public Sub Collect()
    '����������
    Call Initialize
    'RTE���ݒ菈��
    Call SetRteInfomation
End Sub

'����������
Private Sub Initialize()
    '�e���W���[���̏���������
    Call mdl_Input.Init
    '�ϐ��̏�����
    int_Range_Row = 1
    int_Row = INT_ROW_START
    bln_First = False
End Sub

'RTE���ݒ菈��
Private Sub SetRteInfomation()
    Dim obj_Book As Workbook
    Dim obj_Sheet As Worksheet
    '�V�[�g�͈͂̃N���A
    mdl_Input.rng_RteInfoList.ClearContents
    'RTE���t�@�C���̑S��������
    Set obj_Book = Workbooks.Open(mdl_Input.rng_RteInfoFile.Value)
    For Each obj_Sheet In obj_Book.Worksheets
        '�ΏۃV�[�g�̔���
        If obj_Sheet.Range(STR_RNG_SHEET_CHECK) = STR_SHEET_CHECK Then
            int_Row = INT_ROW_START
            bln_First = False
            'RTE��񌟍�����
            Do While SearchRteInfo(obj_Sheet, mdl_Input.rng_RteInfoList)
            Loop
        End If
    Next
    obj_Book.Close
End Sub

'RTE��񌟍�����
Private Function SearchRteInfo(ByRef obj_Sheet As Worksheet, ByRef obj_Range As Range) As Boolean
    '�e���ڂ̓Ǎ���
    If obj_Sheet.Cells(int_Row, INT_COL_ATTRIBUTE).Value = STR_ATTRIB_READ _
    Or obj_Sheet.Cells(int_Row, INT_COL_ATTRIBUTE).Value = STR_ATTRIB_WRITE Then
        If bln_First Then
            obj_Range.Cells(int_Range_Row, 1).Value = "��"
        Else
            obj_Range.Cells(int_Range_Row, 1).Value = obj_Sheet.Range(STR_RNG_MODULE_NAME).Value
            bln_First = True
        End If
        obj_Range.Cells(int_Range_Row, 2).Value = obj_Sheet.Cells(int_Row, INT_COL_ATTRIBUTE).Value
        obj_Range.Cells(int_Range_Row, 3).Value = obj_Sheet.Cells(int_Row, INT_COL_DATA_TYPE).Value
        obj_Range.Cells(int_Range_Row, 4).Value = obj_Sheet.Cells(int_Row, INT_COL_DATA_NAME).Value
        obj_Range.Cells(int_Range_Row, 5).Value = obj_Sheet.Cells(int_Row, INT_COL_DESCRIPTION).Value
        int_Range_Row = int_Range_Row + 1
    End If
    '�����I���̔���
    If obj_Sheet.Cells(int_Row, INT_COL_ATTRIBUTE).Value <> "END" Then
        '���̍s�ֈړ�
        int_Row = int_Row + 1
        '�߂�l�̐ݒ�
        SearchRteInfo = True
    Else
        '�߂�l�̐ݒ�
        SearchRteInfo = False
    End If
End Function
