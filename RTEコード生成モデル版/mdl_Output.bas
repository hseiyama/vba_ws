Attribute VB_Name = "mdl_Output"
Option Explicit

'�����萔
'Private Const INT_ROW_START As Integer = 1
'Private Const INT_COL_TEXT As Integer = 1

'�����ϐ�
'Private int_Row As Integer

'����������
Public Sub Init(ByRef str_FilePath As String)
    '�V�[�g�̃N���A
    'sht_Output.Range("A:A").ClearContents
    '�ϐ��̏�����
    'int_Row = INT_ROW_START
    '�t�@�C���̃I�[�v��
    Open str_FilePath For Output As #1
End Sub

'�I������
Public Sub Final()
    '�t�@�C���̃N���[�Y
    Close #1
End Sub

'�e�L�X�g�����ݏ���
Public Sub WriteText(str_Text As String)
    '�e�L�X�g�̏�����
    'sht_Output.Cells(int_Row, INT_COL_TEXT).Value = str_Text
    Print #1, str_Text
    '���̍s�ֈړ�
    'int_Row = int_Row + 1
End Sub
