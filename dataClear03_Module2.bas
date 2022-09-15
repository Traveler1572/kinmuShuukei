Attribute VB_Name = "Module2"
Option Explicit

Sub data_clear03()
    
    '------------------------------------------------------------------------
    'データクリア
    'Date : 2022-09-07
    '------------------------------------------------------------------------
    
    Dim i As Integer                'カウンタ変数
    
    '実行パラメータ搭載セル
    i = Workbooks("勤務時間集計.xlsm").Worksheets("月").Range("G2")
    
    If i > 10 Then
        Workbooks("勤務時間集計.xlsm").Worksheets("月").Range("B3:B14") = ""
        Workbooks("勤務時間集計.xlsm").Worksheets("月").Range("C3:C7") = ""
        Workbooks("勤務時間集計.xlsm").Worksheets("月").Range("C9:C14") = ""
        Workbooks("勤務時間集計.xlsm").Worksheets("月").Range("D3:E14") = ""
    Else
        MsgBox "処理を中止します。"
    End If
    
    '完了メッセージ出力
    MsgBox "処理完了"
    
End Sub
