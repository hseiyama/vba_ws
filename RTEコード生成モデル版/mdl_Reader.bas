Attribute VB_Name = "mdl_Reader"
Option Explicit

'�����ϐ�
Dim obj_Header As cls_Reader

'�ȈՃe�X�g����
Public Sub Test()
    '�O����
    Set obj_Header = New cls_Reader
    '����������
    obj_Header.Init sht_Header
    '�e�L�X�g�Ǎ��ݏ���
    Do While obj_Header.ReadText
        '�m�F�p�̏o��
        Debug.Print obj_Header.str_Text
    Loop
    '�㏈��
    Set obj_Header = Nothing
End Sub
