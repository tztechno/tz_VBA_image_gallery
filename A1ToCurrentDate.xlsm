Sub A1ToCurrentDate()
    ' 今日の日付を取得
    Dim today As Date
    today = Now()
    Dim year As String
    year = Format(today, "yyyy") ' 4桁の年
    Dim month As String
    month = Format(today, "mm") ' 2桁の月
    Dim day As String
    day = Format(today, "dd") ' 2桁の日

    ' シート名を今日の日付に変更
    Dim sheet As Worksheet
    Set sheet = ActiveWorkbook.ActiveSheet
    sheet.Name = year & "-" & month & "-" & day

    ' A1セルの値を今日の日付に変更
    sheet.Range("A1").Value = year & "-" & month & "-" & day
End Sub
