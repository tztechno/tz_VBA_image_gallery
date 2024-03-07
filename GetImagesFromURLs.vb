' 動作未確認

Sub GetImagesFromURLs()
    Dim i As Integer
    Dim pic As Picture
    Dim rng As Range
    Dim shp As Shape
    
    ' ループで各セルを処理
    For i = 1 To Range("A" & Rows.Count).End(xlUp).Row
        ' URLを取得
        Set rng = Range("A" & i)
        If rng <> "" Then
            ' 画像を取得
            Set shp = ActiveSheet.Shapes.AddPicture(rng.Value, _
                msoFalse, msoTrue, rng.Left, rng.Top, rng.Width, rng.Height)
            ' 画像サイズを調整
            shp.LockAspectRatio = msoTrue
            shp.Width = rng.Offset(0, 1).Width
            shp.Height = rng.Offset(0, 1).Height
        End If
    Next i
End Sub






