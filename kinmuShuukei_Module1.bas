Attribute VB_Name = "Module1"
Option Explicit

Sub soukadojikan_shuukei()
    
    '----------------------------------------------------------------------
    '勤務時間集計処理（月別集計）
    'Date 　: 2022-09-04    新規作成
    'Update : 2022-09-07    soukadojikan_shuukei2()関数「コピー先ブックを開く」「コピー元ブックを閉じる」処理追加
    '         2022-09-14    ループを途中で抜ける処理追加
    '         2022-09-15    本マクロの全体処理時間計測処理追加
    '----------------------------------------------------------------------
    Dim objWorkbook As Workbook
    Dim i As Integer                 'カウンタ用変数
    Dim j(19) As Workbook            '複数ブック用配列
    Dim startTime As Double          '開始時間
    Dim endTime As Double            '終了時間
    Dim processTime As Double        '処理時間計算
    
    '開始時間取得
    startTime = Timer
    
    'コピー元ブックを開く
    Set j(0) = Workbooks.Open("C:\Users\all_o\OneDrive\デスクトップ\Lenovo\LocalBkup\TNS_OfficeWork\01_営業\03_【現場】新豊洲(広域機関／日立)\勤務状況報告書(201901).xls")
    Set j(1) = Workbooks.Open("C:\Users\all_o\OneDrive\デスクトップ\Lenovo\LocalBkup\TNS_OfficeWork\01_営業\03_【現場】新豊洲(広域機関／日立)\勤務状況報告書(201902).xls")
    Set j(2) = Workbooks.Open("C:\Users\all_o\OneDrive\デスクトップ\Lenovo\LocalBkup\TNS_OfficeWork\01_営業\03_【現場】新豊洲(広域機関／日立)\勤務状況報告書(201903).xls")
    Set j(3) = Workbooks.Open("C:\Users\all_o\OneDrive\デスクトップ\Lenovo\LocalBkup\TNS_OfficeWork\01_営業\04_【現場】豊洲(第一生命情報システム／富士通エフサス)\勤務状況報告書(201904)_武井.xls")
    Set j(4) = Workbooks.Open("C:\Users\all_o\OneDrive\デスクトップ\Lenovo\LocalBkup\TNS_OfficeWork\01_営業\04_【現場】豊洲(第一生命情報システム／富士通エフサス)\勤務状況報告書(201905)_武井.xls")
    Set j(5) = Workbooks.Open("C:\Users\all_o\OneDrive\デスクトップ\Lenovo\LocalBkup\TNS_OfficeWork\01_営業\04_【現場】豊洲(第一生命情報システム／富士通エフサス)\勤務状況報告書(201906)_武井.xls")
    Set j(6) = Workbooks.Open("C:\Users\all_o\OneDrive\デスクトップ\Lenovo\LocalBkup\TNS_OfficeWork\01_営業\04_【現場】豊洲(第一生命情報システム／富士通エフサス)\勤務状況報告書(201907)_武井.xls")
    Set j(7) = Workbooks.Open("C:\Users\all_o\OneDrive\デスクトップ\Lenovo\LocalBkup\TNS_OfficeWork\01_営業\04_【現場】豊洲(第一生命情報システム／富士通エフサス)\勤務状況報告書(201908)_武井.xls")
    Set j(8) = Workbooks.Open("C:\Users\all_o\OneDrive\デスクトップ\Lenovo\LocalBkup\TNS_OfficeWork\01_営業\04_【現場】豊洲(第一生命情報システム／富士通エフサス)\勤務状況報告書(201909)_武井.xls")
    Set j(9) = Workbooks.Open("C:\Users\all_o\OneDrive\デスクトップ\Lenovo\LocalBkup\TNS_OfficeWork\01_営業\04_【現場】豊洲(第一生命情報システム／富士通エフサス)\勤務状況報告書(201910)_武井.xlsx")
    Set j(10) = Workbooks.Open("C:\Users\all_o\OneDrive\デスクトップ\Lenovo\LocalBkup\TNS_OfficeWork\01_営業\05_【現場】日本橋(みずほ証券／グッドフィールドカンパニー)\勤務状況報告書(201911)_武井.xlsx")
    Set j(11) = Workbooks.Open("C:\Users\all_o\OneDrive\デスクトップ\Lenovo\LocalBkup\TNS_OfficeWork\01_営業\06_【現場】大手町(日本パレットレンタル／グッドフィールドカンパニー)\勤務状況報告書(201912)_武井.xlsx")
    Set j(12) = Workbooks.Open("C:\Users\all_o\OneDrive\デスクトップ\Lenovo\LocalBkup\TNS_OfficeWork\01_営業\99_自社勤務\勤務状況報告書(202001)_武井.xlsx")
    Set j(13) = Workbooks.Open("C:\Users\all_o\OneDrive\デスクトップ\Lenovo\LocalBkup\TNS_OfficeWork\01_営業\07_【現場】芝公園(スペースバリューホールディングス／日立)\勤務状況報告書(202002)_武井.xlsx")
    Set j(14) = Workbooks.Open("C:\Users\all_o\OneDrive\デスクトップ\Lenovo\LocalBkup\TNS_OfficeWork\01_営業\07_【現場】芝公園(スペースバリューホールディングス／日立)\勤務状況報告書(202003)_武井.xlsx")
    Set j(15) = Workbooks.Open("C:\Users\all_o\OneDrive\デスクトップ\Lenovo\LocalBkup\TNS_OfficeWork\01_営業\07_【現場】芝公園(スペースバリューホールディングス／日立)\勤務状況報告書(202004)_武井.xlsx")
    Set j(16) = Workbooks.Open("C:\Users\all_o\OneDrive\デスクトップ\Lenovo\LocalBkup\TNS_OfficeWork\01_営業\99_自社勤務\【武井洋平】勤務表202005.xlsx")
    Set j(17) = Workbooks.Open("C:\Users\all_o\OneDrive\デスクトップ\Lenovo\LocalBkup\TNS_OfficeWork\01_営業\08_【現場】多摩センター(第一生命情報システム／NSD)\【武井洋平】勤務表202007.xlsx")
    Set j(18) = Workbooks.Open("C:\Users\all_o\OneDrive\デスクトップ\Lenovo\LocalBkup\TNS_OfficeWork\01_営業\08_【現場】多摩センター(第一生命情報システム／NSD)\【武井洋平】勤務表202008.xlsx")

    'コピー先となるブックを開く
    Set objWorkbook = Workbooks.Open("C:\Users\all_o\OneDrive\デスクトップ\Lenovo\LocalBkup\TNS_OfficeWork\01_営業\勤務時間集計.xlsm")
    
    '元シートのセル範囲をコピー後、別ブックへ値のみ貼り付けします
    For i = 3 To 14
        Workbooks("勤務状況報告書(201901).xls").Sheets("原本").Range("G40").Copy
            objWorkbook.Sheets("月").Range("B3").PasteSpecial Paste:=xlPasteValues
        Workbooks("勤務状況報告書(201902).xls").Sheets("原本").Range("G40").Copy
            objWorkbook.Sheets("月").Range("B4").PasteSpecial Paste:=xlPasteValues
        Workbooks("勤務状況報告書(201903).xls").Sheets("原本").Range("G40").Copy
            objWorkbook.Sheets("月").Range("B5").PasteSpecial Paste:=xlPasteValues
        Workbooks("勤務状況報告書(201904)_武井.xls").Sheets("原本").Range("G40").Copy
            objWorkbook.Sheets("月").Range("B6").PasteSpecial Paste:=xlPasteValues
        Workbooks("勤務状況報告書(201905)_武井.xls").Sheets("原本").Range("G40").Copy
            objWorkbook.Sheets("月").Range("B7").PasteSpecial Paste:=xlPasteValues
        Workbooks("勤務状況報告書(201906)_武井.xls").Sheets("原本").Range("G40").Copy
            objWorkbook.Sheets("月").Range("B8").PasteSpecial Paste:=xlPasteValues
        Workbooks("勤務状況報告書(201907)_武井.xls").Sheets("原本").Range("G40").Copy
            objWorkbook.Sheets("月").Range("B9").PasteSpecial Paste:=xlPasteValues
        Workbooks("勤務状況報告書(201908)_武井.xls").Sheets("原本").Range("G40").Copy
            objWorkbook.Sheets("月").Range("B10").PasteSpecial Paste:=xlPasteValues
        Workbooks("勤務状況報告書(201909)_武井.xls").Sheets("原本").Range("G40").Copy
            objWorkbook.Sheets("月").Range("B11").PasteSpecial Paste:=xlPasteValues
        Workbooks("勤務状況報告書(201910)_武井.xlsx").Sheets("原本").Range("I43").Copy
            objWorkbook.Sheets("月").Range("B12").PasteSpecial Paste:=xlPasteValues
        Workbooks("勤務状況報告書(201911)_武井.xlsx").Sheets("原本").Range("I43").Copy
            objWorkbook.Sheets("月").Range("B13").PasteSpecial Paste:=xlPasteValues
        Workbooks("勤務状況報告書(201912)_武井.xlsx").Sheets("原本").Range("I43").Copy
            objWorkbook.Sheets("月").Range("B14").PasteSpecial Paste:=xlPasteValues
        Workbooks("勤務状況報告書(202001)_武井.xlsx").Sheets("原本").Range("I43").Copy
            objWorkbook.Sheets("月").Range("C3").PasteSpecial Paste:=xlPasteValues
        Workbooks("勤務状況報告書(202002)_武井.xlsx").Sheets("原本").Range("I43").Copy
            objWorkbook.Sheets("月").Range("C4").PasteSpecial Paste:=xlPasteValues
        Workbooks("勤務状況報告書(202003)_武井.xlsx").Sheets("原本").Range("I43").Copy
            objWorkbook.Sheets("月").Range("C5").PasteSpecial Paste:=xlPasteValues
        Workbooks("勤務状況報告書(202004)_武井.xlsx").Sheets("TNS").Range("I43").Copy
            objWorkbook.Sheets("月").Range("C6").PasteSpecial Paste:=xlPasteValues
        Workbooks("【武井洋平】勤務表202005.xlsx").Sheets("TNS").Range("I43").Copy
            objWorkbook.Sheets("月").Range("C7").PasteSpecial Paste:=xlPasteValues
        Workbooks("【武井洋平】勤務表202007.xlsx").Sheets("TNS").Range("I43").Copy
            objWorkbook.Sheets("月").Range("C9").PasteSpecial Paste:=xlPasteValues
        Workbooks("【武井洋平】勤務表202008.xlsx").Sheets("TNS").Range("I43").Copy
            objWorkbook.Sheets("月").Range("C10").PasteSpecial Paste:=xlPasteValues
            
                'セルC10に値が入力された時点でループ処理を抜ける
                If Range("C10").Value = Range("C10").Value Then
                    Exit For
                End If
    Next i
    
    'soukadojikan_shuukei2関数を呼び出す
    Call soukadojikan_shuukei2
    
    'コピー元ブックを閉じる
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
    
    '範囲選択を解除します
    Application.CutCopyMode = False
    
    'オブジェクトを解放します
    Set objWorkbook = Nothing
    
    '終了時間取得
    endTime = Timer
    
    '処理時間計算
    processTime = endTime - startTime
    Workbooks("勤務時間集計.xlsm").Sheets("月").Range("I14").Value = processTime
    
    'セルA1をアクティブにする
    Range("A1").Activate
    
    '完了メッセージ出力
    MsgBox "処理完了"
    
End Sub

Function soukadojikan_shuukei2()

    Dim objWorkbook2 As Workbook
    Dim a As Integer                'カウンタ用変数
    Dim c As Workbook               'コピー元ブック用変数
    
    'コピー元ブックを開く
    Set c = Workbooks.Open("C:\Users\all_o\OneDrive\デスクトップ\Lenovo\LocalBkup\TNS_OfficeWork\01_営業\kinmuhyo_202009以降.xlsm")
    
    'コピー先となるブックを開く
    Set objWorkbook2 = Workbooks.Open("C:\Users\all_o\OneDrive\デスクトップ\Lenovo\LocalBkup\TNS_OfficeWork\01_営業\勤務時間集計.xlsm")
    
    For a = 3 To 14
        Workbooks("kinmuhyo_202009以降.xlsm").Sheets("202009").Range("I43").Copy
            objWorkbook2.Sheets("月").Range("C11").PasteSpecial Paste:=xlPasteValues
        Workbooks("kinmuhyo_202009以降.xlsm").Sheets("202010").Range("I43").Copy
            objWorkbook2.Sheets("月").Range("C12").PasteSpecial Paste:=xlPasteValues
        Workbooks("kinmuhyo_202009以降.xlsm").Sheets("202011").Range("I43").Copy
            objWorkbook2.Sheets("月").Range("C13").PasteSpecial Paste:=xlPasteValues
        Workbooks("kinmuhyo_202009以降.xlsm").Sheets("202012").Range("I43").Copy
            objWorkbook2.Sheets("月").Range("C14").PasteSpecial Paste:=xlPasteValues
        Workbooks("kinmuhyo_202009以降.xlsm").Sheets("202101").Range("I43").Copy
            objWorkbook2.Sheets("月").Range("D3").PasteSpecial Paste:=xlPasteValues
        Workbooks("kinmuhyo_202009以降.xlsm").Sheets("202102").Range("I43").Copy
            objWorkbook2.Sheets("月").Range("D4").PasteSpecial Paste:=xlPasteValues
        Workbooks("kinmuhyo_202009以降.xlsm").Sheets("202103").Range("I43").Copy
            objWorkbook2.Sheets("月").Range("D5").PasteSpecial Paste:=xlPasteValues
        Workbooks("kinmuhyo_202009以降.xlsm").Sheets("202104").Range("I43").Copy
            objWorkbook2.Sheets("月").Range("D6").PasteSpecial Paste:=xlPasteValues
        Workbooks("kinmuhyo_202009以降.xlsm").Sheets("202105").Range("I43").Copy
            objWorkbook2.Sheets("月").Range("D7").PasteSpecial Paste:=xlPasteValues
        Workbooks("kinmuhyo_202009以降.xlsm").Sheets("202106").Range("I43").Copy
            objWorkbook2.Sheets("月").Range("D8").PasteSpecial Paste:=xlPasteValues
        Workbooks("kinmuhyo_202009以降.xlsm").Sheets("202107").Range("I43").Copy
            objWorkbook2.Sheets("月").Range("D9").PasteSpecial Paste:=xlPasteValues
        Workbooks("kinmuhyo_202009以降.xlsm").Sheets("202108").Range("I43").Copy
            objWorkbook2.Sheets("月").Range("D10").PasteSpecial Paste:=xlPasteValues
        Workbooks("kinmuhyo_202009以降.xlsm").Sheets("202109").Range("I43").Copy
            objWorkbook2.Sheets("月").Range("D11").PasteSpecial Paste:=xlPasteValues
        Workbooks("kinmuhyo_202009以降.xlsm").Sheets("202110").Range("I43").Copy
            objWorkbook2.Sheets("月").Range("D12").PasteSpecial Paste:=xlPasteValues
        Workbooks("kinmuhyo_202009以降.xlsm").Sheets("202111").Range("I43").Copy
            objWorkbook2.Sheets("月").Range("D13").PasteSpecial Paste:=xlPasteValues
        Workbooks("kinmuhyo_202009以降.xlsm").Sheets("202112").Range("I43").Copy
            objWorkbook2.Sheets("月").Range("D14").PasteSpecial Paste:=xlPasteValues
        Workbooks("kinmuhyo_202009以降.xlsm").Sheets("202201").Range("I43").Copy
            objWorkbook2.Sheets("月").Range("E3").PasteSpecial Paste:=xlPasteValues
        Workbooks("kinmuhyo_202009以降.xlsm").Sheets("202202").Range("I43").Copy
            objWorkbook2.Sheets("月").Range("E4").PasteSpecial Paste:=xlPasteValues
        Workbooks("kinmuhyo_202009以降.xlsm").Sheets("202203").Range("I43").Copy
            objWorkbook2.Sheets("月").Range("E5").PasteSpecial Paste:=xlPasteValues
        Workbooks("kinmuhyo_202009以降.xlsm").Sheets("202204").Range("I43").Copy
            objWorkbook2.Sheets("月").Range("E6").PasteSpecial Paste:=xlPasteValues
        Workbooks("kinmuhyo_202009以降.xlsm").Sheets("202205").Range("I43").Copy
            objWorkbook2.Sheets("月").Range("E7").PasteSpecial Paste:=xlPasteValues
        Workbooks("kinmuhyo_202009以降.xlsm").Sheets("202206").Range("I43").Copy
            objWorkbook2.Sheets("月").Range("E8").PasteSpecial Paste:=xlPasteValues
        Workbooks("kinmuhyo_202009以降.xlsm").Sheets("202207").Range("I43").Copy
            objWorkbook2.Sheets("月").Range("E9").PasteSpecial Paste:=xlPasteValues
        Workbooks("kinmuhyo_202009以降.xlsm").Sheets("202208").Range("I43").Copy
            objWorkbook2.Sheets("月").Range("E10").PasteSpecial Paste:=xlPasteValues
            
                'セルE10に値が入力された時点でループ処理を抜ける
                If Range("E10").Value = Range("E10").Value Then
                    Exit For
                End If
    Next a
    
    'コピー元ブックを閉じる
    Call c.Close(SaveChanges:=False)

End Function
