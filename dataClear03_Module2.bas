Attribute VB_Name = "Module2"
Option Explicit

Sub data_clear03()
    
    '------------------------------------------------------------------------
    '�f�[�^�N���A
    'Date : 2022-09-07
    '------------------------------------------------------------------------
    
    Dim i As Integer                '�J�E���^�ϐ�
    
    '���s�p�����[�^���ڃZ��
    i = Workbooks("�Ζ����ԏW�v.xlsm").Worksheets("��").Range("G2")
    
    If i > 10 Then
        Workbooks("�Ζ����ԏW�v.xlsm").Worksheets("��").Range("B3:B14") = ""
        Workbooks("�Ζ����ԏW�v.xlsm").Worksheets("��").Range("C3:C7") = ""
        Workbooks("�Ζ����ԏW�v.xlsm").Worksheets("��").Range("C9:C14") = ""
        Workbooks("�Ζ����ԏW�v.xlsm").Worksheets("��").Range("D3:E14") = ""
    Else
        MsgBox "�����𒆎~���܂��B"
    End If
    
    '�������b�Z�[�W�o��
    MsgBox "��������"
    
End Sub
