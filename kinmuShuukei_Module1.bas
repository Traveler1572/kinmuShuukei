Attribute VB_Name = "Module1"
Option Explicit

Sub soukadojikan_shuukei()
    
    '----------------------------------------------------------------------
    '�Ζ����ԏW�v�����i���ʏW�v�j
    'Date �@: 2022-09-04    �V�K�쐬
    'Update : 2022-09-07    soukadojikan_shuukei2()�֐��u�R�s�[��u�b�N���J���v�u�R�s�[���u�b�N�����v�����ǉ�
    '         2022-09-14    ���[�v��r���Ŕ����鏈���ǉ�
    '         2022-09-15    �{�}�N���̑S�̏������Ԍv�������ǉ�
    '----------------------------------------------------------------------
    Dim objWorkbook As Workbook
    Dim i As Integer                 '�J�E���^�p�ϐ�
    Dim j(19) As Workbook            '�����u�b�N�p�z��
    Dim startTime As Double          '�J�n����
    Dim endTime As Double            '�I������
    Dim processTime As Double        '�������Ԍv�Z
    
    '�J�n���Ԏ擾
    startTime = Timer
    
    '�R�s�[���u�b�N���J��
    Set j(0) = Workbooks.Open("C:\Users\all_o\OneDrive\�f�X�N�g�b�v\Lenovo\LocalBkup\TNS_OfficeWork\01_�c��\03_�y����z�V�L�F(�L��@�ց^����)\�Ζ��󋵕񍐏�(201901).xls")
    Set j(1) = Workbooks.Open("C:\Users\all_o\OneDrive\�f�X�N�g�b�v\Lenovo\LocalBkup\TNS_OfficeWork\01_�c��\03_�y����z�V�L�F(�L��@�ց^����)\�Ζ��󋵕񍐏�(201902).xls")
    Set j(2) = Workbooks.Open("C:\Users\all_o\OneDrive\�f�X�N�g�b�v\Lenovo\LocalBkup\TNS_OfficeWork\01_�c��\03_�y����z�V�L�F(�L��@�ց^����)\�Ζ��󋵕񍐏�(201903).xls")
    Set j(3) = Workbooks.Open("C:\Users\all_o\OneDrive\�f�X�N�g�b�v\Lenovo\LocalBkup\TNS_OfficeWork\01_�c��\04_�y����z�L�F(��ꐶ�����V�X�e���^�x�m�ʃG�t�T�X)\�Ζ��󋵕񍐏�(201904)_����.xls")
    Set j(4) = Workbooks.Open("C:\Users\all_o\OneDrive\�f�X�N�g�b�v\Lenovo\LocalBkup\TNS_OfficeWork\01_�c��\04_�y����z�L�F(��ꐶ�����V�X�e���^�x�m�ʃG�t�T�X)\�Ζ��󋵕񍐏�(201905)_����.xls")
    Set j(5) = Workbooks.Open("C:\Users\all_o\OneDrive\�f�X�N�g�b�v\Lenovo\LocalBkup\TNS_OfficeWork\01_�c��\04_�y����z�L�F(��ꐶ�����V�X�e���^�x�m�ʃG�t�T�X)\�Ζ��󋵕񍐏�(201906)_����.xls")
    Set j(6) = Workbooks.Open("C:\Users\all_o\OneDrive\�f�X�N�g�b�v\Lenovo\LocalBkup\TNS_OfficeWork\01_�c��\04_�y����z�L�F(��ꐶ�����V�X�e���^�x�m�ʃG�t�T�X)\�Ζ��󋵕񍐏�(201907)_����.xls")
    Set j(7) = Workbooks.Open("C:\Users\all_o\OneDrive\�f�X�N�g�b�v\Lenovo\LocalBkup\TNS_OfficeWork\01_�c��\04_�y����z�L�F(��ꐶ�����V�X�e���^�x�m�ʃG�t�T�X)\�Ζ��󋵕񍐏�(201908)_����.xls")
    Set j(8) = Workbooks.Open("C:\Users\all_o\OneDrive\�f�X�N�g�b�v\Lenovo\LocalBkup\TNS_OfficeWork\01_�c��\04_�y����z�L�F(��ꐶ�����V�X�e���^�x�m�ʃG�t�T�X)\�Ζ��󋵕񍐏�(201909)_����.xls")
    Set j(9) = Workbooks.Open("C:\Users\all_o\OneDrive\�f�X�N�g�b�v\Lenovo\LocalBkup\TNS_OfficeWork\01_�c��\04_�y����z�L�F(��ꐶ�����V�X�e���^�x�m�ʃG�t�T�X)\�Ζ��󋵕񍐏�(201910)_����.xlsx")
    Set j(10) = Workbooks.Open("C:\Users\all_o\OneDrive\�f�X�N�g�b�v\Lenovo\LocalBkup\TNS_OfficeWork\01_�c��\05_�y����z���{��(�݂��ُ،��^�O�b�h�t�B�[���h�J���p�j�[)\�Ζ��󋵕񍐏�(201911)_����.xlsx")
    Set j(11) = Workbooks.Open("C:\Users\all_o\OneDrive\�f�X�N�g�b�v\Lenovo\LocalBkup\TNS_OfficeWork\01_�c��\06_�y����z��蒬(���{�p���b�g�����^���^�O�b�h�t�B�[���h�J���p�j�[)\�Ζ��󋵕񍐏�(201912)_����.xlsx")
    Set j(12) = Workbooks.Open("C:\Users\all_o\OneDrive\�f�X�N�g�b�v\Lenovo\LocalBkup\TNS_OfficeWork\01_�c��\99_���ЋΖ�\�Ζ��󋵕񍐏�(202001)_����.xlsx")
    Set j(13) = Workbooks.Open("C:\Users\all_o\OneDrive\�f�X�N�g�b�v\Lenovo\LocalBkup\TNS_OfficeWork\01_�c��\07_�y����z�Ō���(�X�y�[�X�o�����[�z�[���f�B���O�X�^����)\�Ζ��󋵕񍐏�(202002)_����.xlsx")
    Set j(14) = Workbooks.Open("C:\Users\all_o\OneDrive\�f�X�N�g�b�v\Lenovo\LocalBkup\TNS_OfficeWork\01_�c��\07_�y����z�Ō���(�X�y�[�X�o�����[�z�[���f�B���O�X�^����)\�Ζ��󋵕񍐏�(202003)_����.xlsx")
    Set j(15) = Workbooks.Open("C:\Users\all_o\OneDrive\�f�X�N�g�b�v\Lenovo\LocalBkup\TNS_OfficeWork\01_�c��\07_�y����z�Ō���(�X�y�[�X�o�����[�z�[���f�B���O�X�^����)\�Ζ��󋵕񍐏�(202004)_����.xlsx")
    Set j(16) = Workbooks.Open("C:\Users\all_o\OneDrive\�f�X�N�g�b�v\Lenovo\LocalBkup\TNS_OfficeWork\01_�c��\99_���ЋΖ�\�y����m���z�Ζ��\202005.xlsx")
    Set j(17) = Workbooks.Open("C:\Users\all_o\OneDrive\�f�X�N�g�b�v\Lenovo\LocalBkup\TNS_OfficeWork\01_�c��\08_�y����z�����Z���^�[(��ꐶ�����V�X�e���^NSD)\�y����m���z�Ζ��\202007.xlsx")
    Set j(18) = Workbooks.Open("C:\Users\all_o\OneDrive\�f�X�N�g�b�v\Lenovo\LocalBkup\TNS_OfficeWork\01_�c��\08_�y����z�����Z���^�[(��ꐶ�����V�X�e���^NSD)\�y����m���z�Ζ��\202008.xlsx")

    '�R�s�[��ƂȂ�u�b�N���J��
    Set objWorkbook = Workbooks.Open("C:\Users\all_o\OneDrive\�f�X�N�g�b�v\Lenovo\LocalBkup\TNS_OfficeWork\01_�c��\�Ζ����ԏW�v.xlsm")
    
    '���V�[�g�̃Z���͈͂��R�s�[��A�ʃu�b�N�֒l�̂ݓ\��t�����܂�
    For i = 3 To 14
        Workbooks("�Ζ��󋵕񍐏�(201901).xls").Sheets("���{").Range("G40").Copy
            objWorkbook.Sheets("��").Range("B3").PasteSpecial Paste:=xlPasteValues
        Workbooks("�Ζ��󋵕񍐏�(201902).xls").Sheets("���{").Range("G40").Copy
            objWorkbook.Sheets("��").Range("B4").PasteSpecial Paste:=xlPasteValues
        Workbooks("�Ζ��󋵕񍐏�(201903).xls").Sheets("���{").Range("G40").Copy
            objWorkbook.Sheets("��").Range("B5").PasteSpecial Paste:=xlPasteValues
        Workbooks("�Ζ��󋵕񍐏�(201904)_����.xls").Sheets("���{").Range("G40").Copy
            objWorkbook.Sheets("��").Range("B6").PasteSpecial Paste:=xlPasteValues
        Workbooks("�Ζ��󋵕񍐏�(201905)_����.xls").Sheets("���{").Range("G40").Copy
            objWorkbook.Sheets("��").Range("B7").PasteSpecial Paste:=xlPasteValues
        Workbooks("�Ζ��󋵕񍐏�(201906)_����.xls").Sheets("���{").Range("G40").Copy
            objWorkbook.Sheets("��").Range("B8").PasteSpecial Paste:=xlPasteValues
        Workbooks("�Ζ��󋵕񍐏�(201907)_����.xls").Sheets("���{").Range("G40").Copy
            objWorkbook.Sheets("��").Range("B9").PasteSpecial Paste:=xlPasteValues
        Workbooks("�Ζ��󋵕񍐏�(201908)_����.xls").Sheets("���{").Range("G40").Copy
            objWorkbook.Sheets("��").Range("B10").PasteSpecial Paste:=xlPasteValues
        Workbooks("�Ζ��󋵕񍐏�(201909)_����.xls").Sheets("���{").Range("G40").Copy
            objWorkbook.Sheets("��").Range("B11").PasteSpecial Paste:=xlPasteValues
        Workbooks("�Ζ��󋵕񍐏�(201910)_����.xlsx").Sheets("���{").Range("I43").Copy
            objWorkbook.Sheets("��").Range("B12").PasteSpecial Paste:=xlPasteValues
        Workbooks("�Ζ��󋵕񍐏�(201911)_����.xlsx").Sheets("���{").Range("I43").Copy
            objWorkbook.Sheets("��").Range("B13").PasteSpecial Paste:=xlPasteValues
        Workbooks("�Ζ��󋵕񍐏�(201912)_����.xlsx").Sheets("���{").Range("I43").Copy
            objWorkbook.Sheets("��").Range("B14").PasteSpecial Paste:=xlPasteValues
        Workbooks("�Ζ��󋵕񍐏�(202001)_����.xlsx").Sheets("���{").Range("I43").Copy
            objWorkbook.Sheets("��").Range("C3").PasteSpecial Paste:=xlPasteValues
        Workbooks("�Ζ��󋵕񍐏�(202002)_����.xlsx").Sheets("���{").Range("I43").Copy
            objWorkbook.Sheets("��").Range("C4").PasteSpecial Paste:=xlPasteValues
        Workbooks("�Ζ��󋵕񍐏�(202003)_����.xlsx").Sheets("���{").Range("I43").Copy
            objWorkbook.Sheets("��").Range("C5").PasteSpecial Paste:=xlPasteValues
        Workbooks("�Ζ��󋵕񍐏�(202004)_����.xlsx").Sheets("TNS").Range("I43").Copy
            objWorkbook.Sheets("��").Range("C6").PasteSpecial Paste:=xlPasteValues
        Workbooks("�y����m���z�Ζ��\202005.xlsx").Sheets("TNS").Range("I43").Copy
            objWorkbook.Sheets("��").Range("C7").PasteSpecial Paste:=xlPasteValues
        Workbooks("�y����m���z�Ζ��\202007.xlsx").Sheets("TNS").Range("I43").Copy
            objWorkbook.Sheets("��").Range("C9").PasteSpecial Paste:=xlPasteValues
        Workbooks("�y����m���z�Ζ��\202008.xlsx").Sheets("TNS").Range("I43").Copy
            objWorkbook.Sheets("��").Range("C10").PasteSpecial Paste:=xlPasteValues
            
                '�Z��C10�ɒl�����͂��ꂽ���_�Ń��[�v�����𔲂���
                If Range("C10").Value = Range("C10").Value Then
                    Exit For
                End If
    Next i
    
    'soukadojikan_shuukei2�֐����Ăяo��
    Call soukadojikan_shuukei2
    
    '�R�s�[���u�b�N�����
    Call j(0).Close(SaveChanges:=False)
    Call j(1).Close(SaveChanges:=False)
    Call j(2).Close(SaveChanges:=False)
    Call j(3).Close(SaveChanges:=False)
    Call j(4).Close(SaveChanges:=False)
    Call j(5).Close(SaveChanges:=False)
    Call j(6).Close(SaveChanges:=False)
    Call j(7).Close(SaveChanges:=False)
    Call j(8).Close(SaveChanges:=False)
    Call j(9).Close(SaveChanges:=False)
    Call j(10).Close(SaveChanges:=False)
    Call j(11).Close(SaveChanges:=False)
    Call j(12).Close(SaveChanges:=False)
    Call j(13).Close(SaveChanges:=False)
    Call j(14).Close(SaveChanges:=False)
    Call j(15).Close(SaveChanges:=False)
    Call j(16).Close(SaveChanges:=False)
    Call j(17).Close(SaveChanges:=False)
    Call j(18).Close(SaveChanges:=False)
    
    '�͈͑I�����������܂�
    Application.CutCopyMode = False
    
    '�I�u�W�F�N�g��������܂�
    Set objWorkbook = Nothing
    
    '�I�����Ԏ擾
    endTime = Timer
    
    '�������Ԍv�Z
    processTime = endTime - startTime
    Workbooks("�Ζ����ԏW�v.xlsm").Sheets("��").Range("I14").Value = processTime
    
    '�Z��A1���A�N�e�B�u�ɂ���
    Range("A1").Activate
    
    '�������b�Z�[�W�o��
    MsgBox "��������"
    
End Sub

Function soukadojikan_shuukei2()

    Dim objWorkbook2 As Workbook
    Dim a As Integer                '�J�E���^�p�ϐ�
    Dim c As Workbook               '�R�s�[���u�b�N�p�ϐ�
    
    '�R�s�[���u�b�N���J��
    Set c = Workbooks.Open("C:\Users\all_o\OneDrive\�f�X�N�g�b�v\Lenovo\LocalBkup\TNS_OfficeWork\01_�c��\kinmuhyo_202009�ȍ~.xlsm")
    
    '�R�s�[��ƂȂ�u�b�N���J��
    Set objWorkbook2 = Workbooks.Open("C:\Users\all_o\OneDrive\�f�X�N�g�b�v\Lenovo\LocalBkup\TNS_OfficeWork\01_�c��\�Ζ����ԏW�v.xlsm")
    
    For a = 3 To 14
        Workbooks("kinmuhyo_202009�ȍ~.xlsm").Sheets("202009").Range("I43").Copy
            objWorkbook2.Sheets("��").Range("C11").PasteSpecial Paste:=xlPasteValues
        Workbooks("kinmuhyo_202009�ȍ~.xlsm").Sheets("202010").Range("I43").Copy
            objWorkbook2.Sheets("��").Range("C12").PasteSpecial Paste:=xlPasteValues
        Workbooks("kinmuhyo_202009�ȍ~.xlsm").Sheets("202011").Range("I43").Copy
            objWorkbook2.Sheets("��").Range("C13").PasteSpecial Paste:=xlPasteValues
        Workbooks("kinmuhyo_202009�ȍ~.xlsm").Sheets("202012").Range("I43").Copy
            objWorkbook2.Sheets("��").Range("C14").PasteSpecial Paste:=xlPasteValues
        Workbooks("kinmuhyo_202009�ȍ~.xlsm").Sheets("202101").Range("I43").Copy
            objWorkbook2.Sheets("��").Range("D3").PasteSpecial Paste:=xlPasteValues
        Workbooks("kinmuhyo_202009�ȍ~.xlsm").Sheets("202102").Range("I43").Copy
            objWorkbook2.Sheets("��").Range("D4").PasteSpecial Paste:=xlPasteValues
        Workbooks("kinmuhyo_202009�ȍ~.xlsm").Sheets("202103").Range("I43").Copy
            objWorkbook2.Sheets("��").Range("D5").PasteSpecial Paste:=xlPasteValues
        Workbooks("kinmuhyo_202009�ȍ~.xlsm").Sheets("202104").Range("I43").Copy
            objWorkbook2.Sheets("��").Range("D6").PasteSpecial Paste:=xlPasteValues
        Workbooks("kinmuhyo_202009�ȍ~.xlsm").Sheets("202105").Range("I43").Copy
            objWorkbook2.Sheets("��").Range("D7").PasteSpecial Paste:=xlPasteValues
        Workbooks("kinmuhyo_202009�ȍ~.xlsm").Sheets("202106").Range("I43").Copy
            objWorkbook2.Sheets("��").Range("D8").PasteSpecial Paste:=xlPasteValues
        Workbooks("kinmuhyo_202009�ȍ~.xlsm").Sheets("202107").Range("I43").Copy
            objWorkbook2.Sheets("��").Range("D9").PasteSpecial Paste:=xlPasteValues
        Workbooks("kinmuhyo_202009�ȍ~.xlsm").Sheets("202108").Range("I43").Copy
            objWorkbook2.Sheets("��").Range("D10").PasteSpecial Paste:=xlPasteValues
        Workbooks("kinmuhyo_202009�ȍ~.xlsm").Sheets("202109").Range("I43").Copy
            objWorkbook2.Sheets("��").Range("D11").PasteSpecial Paste:=xlPasteValues
        Workbooks("kinmuhyo_202009�ȍ~.xlsm").Sheets("202110").Range("I43").Copy
            objWorkbook2.Sheets("��").Range("D12").PasteSpecial Paste:=xlPasteValues
        Workbooks("kinmuhyo_202009�ȍ~.xlsm").Sheets("202111").Range("I43").Copy
            objWorkbook2.Sheets("��").Range("D13").PasteSpecial Paste:=xlPasteValues
        Workbooks("kinmuhyo_202009�ȍ~.xlsm").Sheets("202112").Range("I43").Copy
            objWorkbook2.Sheets("��").Range("D14").PasteSpecial Paste:=xlPasteValues
        Workbooks("kinmuhyo_202009�ȍ~.xlsm").Sheets("202201").Range("I43").Copy
            objWorkbook2.Sheets("��").Range("E3").PasteSpecial Paste:=xlPasteValues
        Workbooks("kinmuhyo_202009�ȍ~.xlsm").Sheets("202202").Range("I43").Copy
            objWorkbook2.Sheets("��").Range("E4").PasteSpecial Paste:=xlPasteValues
        Workbooks("kinmuhyo_202009�ȍ~.xlsm").Sheets("202203").Range("I43").Copy
            objWorkbook2.Sheets("��").Range("E5").PasteSpecial Paste:=xlPasteValues
        Workbooks("kinmuhyo_202009�ȍ~.xlsm").Sheets("202204").Range("I43").Copy
            objWorkbook2.Sheets("��").Range("E6").PasteSpecial Paste:=xlPasteValues
        Workbooks("kinmuhyo_202009�ȍ~.xlsm").Sheets("202205").Range("I43").Copy
            objWorkbook2.Sheets("��").Range("E7").PasteSpecial Paste:=xlPasteValues
        Workbooks("kinmuhyo_202009�ȍ~.xlsm").Sheets("202206").Range("I43").Copy
            objWorkbook2.Sheets("��").Range("E8").PasteSpecial Paste:=xlPasteValues
        Workbooks("kinmuhyo_202009�ȍ~.xlsm").Sheets("202207").Range("I43").Copy
            objWorkbook2.Sheets("��").Range("E9").PasteSpecial Paste:=xlPasteValues
        Workbooks("kinmuhyo_202009�ȍ~.xlsm").Sheets("202208").Range("I43").Copy
            objWorkbook2.Sheets("��").Range("E10").PasteSpecial Paste:=xlPasteValues
            
                '�Z��E10�ɒl�����͂��ꂽ���_�Ń��[�v�����𔲂���
                If Range("E10").Value = Range("E10").Value Then
                    Exit For
                End If
    Next a
    
    '�R�s�[���u�b�N�����
    Call c.Close(SaveChanges:=False)

End Function
