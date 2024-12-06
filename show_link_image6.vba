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
    Dim filteredRange As Range
    Dim cell As Range

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
                  "            grid-template-columns: repeat(5, 1fr);" & vbCrLf & _
                  "            gap: 10px;" & vbCrLf & _
                  "            padding: 10px;" & vbCrLf & _
                  "        }" & vbCrLf & _
                  "        .image-container img {" & vbCrLf & _
                  "            width: 100%;" & vbCrLf & _
                  "            height: 100%;" & vbCrLf & _
                  "            object-fit: contain;" & vbCrLf & _
                  "            background-color: #f0f0f0;" & vbCrLf & _
                  "            border: 1px solid #ccc;" & vbCrLf & _
                  "            border-radius: 5px;" & vbCrLf & _
                  "        }" & vbCrLf & _
                  "    </style>" & vbCrLf & _
                  "</head>" & vbCrLf & _
                  "<body>" & vbCrLf & _
                  "    <div class='image-container'>" & vbCrLf


    ' リストの最終行を取得
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' フィルタリングされた範囲を取得
    On Error Resume Next ' エラー処理（フィルターがない場合にエラーを無視）
    Set filteredRange = ws.Range("A1:A" & lastRow).SpecialCells(xlCellTypeVisible) ' 可視セルを取得
    On Error GoTo 0 ' エラー処理を元に戻す

    ' フィルタリングされていない場合は、全行を処理
    If filteredRange Is Nothing Then
        Set filteredRange = ws.Range("A1:A" & lastRow)
    End If

    ' URLリストをループ（フィルタリングされた範囲内のみ）
    For Each cell In filteredRange
        ' ハイパーリンクがあれば、そのURLを取得
        If cell.Hyperlinks.Count > 0 Then
            imgURL = cell.Hyperlinks(1).Address
        Else
            imgURL = cell.Value
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
    Next cell

    ' HTMLの終了部分
    htmlContent = htmlContent & "    </div>" & vbCrLf & "</body>" & vbCrLf & "</html>"

    ' HTMLファイルとして保存
    Dim htmlFile As Object
    Set htmlFile = fso.CreateTextFile(ThisWorkbook.Path & "\image.html", True)
    htmlFile.Write htmlContent
    htmlFile.Close

    MsgBox "HTMLファイルが生成されました: " & ThisWorkbook.Path & "\image.html"
End Sub


Sub OpenHTMLInBrowser(filePath As String)
    ' デフォルトのブラウザでHTMLファイルを開く
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")
    shell.Run filePath
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
