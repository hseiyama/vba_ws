Attribute VB_Name = "mdl_Output"
Option Explicit

'�����萔
Private Const INT_ROW_START As Integer = 1
Private Const INT_COL_TEXT As Integer = 1

'�����ϐ�
Private int_Row As Integer

'����������
Public Sub Init()
    '�V�[�g�̃N���A
    sht_Output.Range("A:A").Clear
    '�ϐ��̏�����
    int_Row = INT_ROW_START
End Sub

'�e�L�X�g�����ݏ���
Public Sub WriteText(str_Text As String)
    '�e�L�X�g�̏�����
    sht_Output.Cells(int_Row, INT_COL_TEXT) = str_Text
    '���̍s�ֈړ�
    int_Row = int_Row + 1
End Sub

'�ȈՃe�X�g����
Public Sub Test()
    Dim int_Index As Integer
    '����������
    Init
    '�m�F�p�̏o��
    For int_Index = 1 To 10
        '�e�L�X�g�����ݏ���
        WriteText "OutputText" & int_Index
    Next
End Sub
