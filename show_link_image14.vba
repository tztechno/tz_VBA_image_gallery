Sub FetchImageURLsAndGenerateHTML()
    Dim ws As Worksheet
    Dim htmlContent As String
    Dim lastRow As Long
    Dim imgURL As String
    Dim filteredRange As Range
    Dim cell As Range
    Dim filePath As String
    Dim successCount As Long
    
    ' エラーハンドリングの設定
    On Error GoTo ErrorHandler
    
    ' シートの初期化
    Set ws = ThisWorkbook.Sheets(1)
    
    ' HTMLの開始部分
    htmlContent = "<html>" & vbCrLf & _
                  "<head>" & vbCrLf & _
                  "    <style>" & vbCrLf & _
                  "        .image-container {" & vbCrLf & _
                  "            display: grid;" & vbCrLf & _
                  "            grid-template-columns: repeat(5, 1fr);" & vbCrLf & _
                  "            gap: 10px;" & vbCrLf & _
                  "            padding: 10px;" & vbCrLf & _
                  "        }" & vbCrLf & _
                  "        .image-wrapper {" & vbCrLf & _
                  "            display: flex;" & vbCrLf & _
                  "            justify-content: center;" & vbCrLf & _
                  "            align-items: center;" & vbCrLf & _
                  "            height: 300px;" & vbCrLf & _
                  "            border: 1px solid #ddd;" & vbCrLf & _
                  "        }" & vbCrLf & _
                  "        .image-wrapper img {" & vbCrLf & _
                  "            max-width: 100%;" & vbCrLf & _
                  "            max-height: 100%;" & vbCrLf & _
                  "            object-fit: contain;" & vbCrLf & _
                  "        }" & vbCrLf & _
                  "    </style>" & vbCrLf & _
                  "</head>" & vbCrLf & _
                  "<body>" & vbCrLf & _
                  "    <div class='image-container'>" & vbCrLf
    
    ' リストの最終行を取得
    lastRow = ws.Cells(ws.Rows.Count, 11).End(xlUp).Row
    
    ' フィルタリングされた範囲を取得
    Set filteredRange = ws.Range("K1:K" & lastRow).SpecialCells(xlCellTypeVisible)
    
    ' カウンターのリセット
    successCount = 0
    
    ' URLリストをループ
    For Each cell In filteredRange
        ' URLの取得（ハイパーリンクと通常のテキストの両方に対応）
        If cell.Hyperlinks.Count > 0 Then
            imgURL = cell.Hyperlinks(1).Address
        Else
            imgURL = cell.Value
        End If
        
        ' 空白セルを除外し、httpまたはhttpsで始まるかを確認
        If Len(Trim(imgURL)) > 0 And (LCase(Left(imgURL, 7)) = "http://" Or LCase(Left(imgURL, 8)) = "https://") Then
            ' HTMLに画像URLを直接埋め込む
            htmlContent = htmlContent & "        <div class='image-wrapper'>" & _
                          "<img src='" & imgURL & "' alt='Image'/>" & _
                          "</div>" & vbCrLf
            successCount = successCount + 1
        End If
    Next cell
    
    ' HTMLの終了部分
    htmlContent = htmlContent & "    </div>" & vbCrLf & "</body>" & vbCrLf & "</html>"
    
    ' HTMLファイルとして保存
    filePath = ThisWorkbook.Path & "\image_gallery.html"
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open filePath For Output As #fileNum
    Print #fileNum, htmlContent
    Close #fileNum
    
    ' シェル関数を使用してブラウザで開く
    Call Shell("cmd.exe /c start " & filePath, vbNormalFocus)
    
    ' 処理結果の表示
    MsgBox "HTMLファイルが生成されました: " & filePath & vbNewLine & _
           "成功したURL: " & successCount
    
    Exit Sub
    
ErrorHandler:
    Dim errorMsg As String
    errorMsg = "エラーが発生しました。" & vbNewLine & _
               "エラー番号: " & Err.Number & vbNewLine & _
               "エラー説明: " & Err.Description
    
    MsgBox errorMsg, vbCritical
    Debug.Print errorMsg
End Sub
