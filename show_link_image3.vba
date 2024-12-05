Sub FetchImagesAndGenerateHTML()
    Dim ws As Worksheet
    Dim http As Object
    Dim fso As Object
    Dim imgData As Variant
    Dim base64Str As String
    Dim htmlContent As String
    Dim lastRow As Long
    Dim imgURL As String
    Dim i As Long

    ' シートとHTTPリクエストオブジェクトの初期化
    Set ws = ThisWorkbook.Sheets(1)
    Set http = CreateObject("MSXML2.XMLHTTP")
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' HTMLの開始部分
    htmlContent = "<html>" & vbCrLf & _
                  "<head>" & vbCrLf & _
                  "    <style>" & vbCrLf & _
                  "        .image-container {" & vbCrLf & _
                  "            display: grid;" & vbCrLf & _
                  "            grid-template-columns: repeat(4, 1fr);" & vbCrLf & _
                  "            gap: 10px;" & vbCrLf & _
                  "            padding: 10px;" & vbCrLf & _
                  "        }" & vbCrLf & _
                  "        .image-container img {" & vbCrLf & _
                  "            width: 100%;" & vbCrLf & _
                  "            height: auto;" & vbCrLf & _
                  "            border: 1px solid #ccc;" & vbCrLf & _
                  "            border-radius: 5px;" & vbCrLf & _
                  "        }" & vbCrLf & _
                  "    </style>" & vbCrLf & _
                  "</head>" & vbCrLf & _
                  "<body>" & vbCrLf & _
                  "    <div class='image-container'>" & vbCrLf

    ' リストの最終行を取得
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' URLリストをループ
    For i = 1 To lastRow
        ' ハイパーリンクがあれば、そのURLを取得
        If ws.Cells(i, 1).Hyperlinks.Count > 0 Then
            imgURL = ws.Cells(i, 1).Hyperlinks(1).Address
        Else
            imgURL = ws.Cells(i, 1).Value
        End If
        
        ' 空白セルをスキップ
        If Trim(imgURL) = "" Then GoTo SkipIteration

        ' URL形式が正しいか確認 (簡易チェック)
        If Not IsValidURL(imgURL) Then GoTo SkipIteration

        ' HTTPリクエストで画像データを取得
        On Error Resume Next
        http.Open "GET", imgURL, False
        http.Send
        On Error GoTo 0
        
        If http.Status = 200 Then
            imgData = http.ResponseBody
            
            ' Base64エンコード
            base64Str = Base64Encode(imgData)
            
            ' HTMLに画像を埋め込む
            htmlContent = htmlContent & "<img src='data:image/png;base64," & base64Str & "' alt='Image'/>" & vbCrLf
        End If

SkipIteration:
    Next i

    ' HTMLの終了部分
    htmlContent = htmlContent & "    </div>" & vbCrLf & "</body>" & vbCrLf & "</html>"

    ' HTMLファイルとして保存
    Dim htmlFile As Object
    Set htmlFile = fso.CreateTextFile(ThisWorkbook.Path & "\images.html", True)
    htmlFile.Write htmlContent
    htmlFile.Close

    MsgBox "HTMLファイルが生成されました: " & ThisWorkbook.Path & "\images.html"
End Sub

Function Base64Encode(ByVal data As Variant) As String
    Dim xmlDoc As Object
    Dim node As Object

    ' XMLオブジェクトを利用してBase64エンコードを実現
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    Set node = xmlDoc.CreateElement("Base64Data")
    node.DataType = "bin.base64"
    node.NodeTypedValue = data
    Base64Encode = node.Text

    Set node = Nothing
    Set xmlDoc = Nothing
End Function

Function IsValidURL(ByVal url As String) As Boolean
    ' URLの簡易チェック: "http://" または "https://" で始まる場合に有効と判定
    If url Like "http://*" Or url Like "https://*" Then
        IsValidURL = True
    Else
        IsValidURL = False
    End If
End Function
