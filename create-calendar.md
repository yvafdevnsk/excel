# Excel2502: カレンダーを作る

## 環境
- Microsoft 365 Apps Excel 2502

## ワークブックを開いたときに動的にテキストボックスとボタンを配置する
ThisWorkbook
```
private Sub Workbook_Open()
    Dim obj As OLEObject
    Dim txtFrom As OLEObject
    Dim txtTo As OLEObject
    Dim btnCreateCalendar As OLEObject

    'コントロールをすべて削除して毎回作り直す。
    For Each obj In ActiveSheet.OLEObjects
        obj.Delete
    Next obj

    'セル全体を「すべてクリア」する。
    Cells.Clear

    'セル全体を文字列書式にする。
    Cells.NumberFormat = "@"

    '期間Fromのラベルを設定する。
    Range("B2").Value = "From: "
    '期間Fromを入力するテキストボックスを配置する。
    'テキストボックスの初期値を現在年月にする。
    Set txtFrom = ActiveSheet.OLEObjects.Add( _
        ClassType:="Forms.TextBox.1", _
        Left:=Range("C2").Left, _
        Top:=Range("C2").Top, _
        Width:=80, _
        Height:=20)
    txtFrom.Name = "TxtYearMonthFrom"
    txtFrom.Object.Text = Format(Date, "yyyymm")

    '期間Toのラベルを設定する。
    Range("E2").Value = "From: "
    '期間Toを入力するテキストボックスを配置する。
    'テキストボックスの初期値を現在年月の翌月にする。
    Set txtTo = ActiveSheet.OLEObjects.Add( _
        ClassType:="Forms.TextBox.1", _
        Left:=Range("F2").Left, _
        Top:=Range("F2").Top, _
        Width:=80, _
        Height:=20)
    txtTo.Name = "TxtYearMonthTo"
    txtTo.Object.Text = Format(DateAdd("m", 1, Date), "yyyymm")

    '「カレンダーを作成」ボタンを配置する。
    Set btnCreateCalendar = ActiveSheet.OLEObjects.Add( _
        ClassType:="Forms.CommandButton.1", _
        Left:=Range("H2").Left, _
        Top:=Range("H2").Top, _
        Width:=120, _
        Height:=30)
    btnCreateCalendar.Name = "BtnCreateCalendar"
    btnCreateCalendar.Object.Caption = "カレンダーを作成"
End Sub
```

参考情報。
- [Workbook.Open event (Excel) | Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.open)
