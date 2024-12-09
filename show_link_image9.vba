Sub FetchImagesAndGenerateHTML()
    Dim ws As Worksheet
    Dim objWinHttpReq As Object
    Dim fso As Object
    Dim imgData As Variant
    Dim base64Str As String
    Dim htmlContent As String
    Dim lastRow As Long
    Dim imgURL As String
    Dim filteredRange As Range
    Dim cell As Range
    Dim filePath As String
    Dim successCount As Long
    Dim failedCount As Long
    
    ' エラーハンドリングの設定
    On Error GoTo ErrorHandler
    
    ' シートとオブジェクトの初期化
    Set ws = ThisWorkbook.Sheets(1)
    
    ' オブジェクト作成の多重防御
    On Error Resume Next
    Set objWinHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
    If objWinHttpReq Is Nothing Then
        Set objWinHttpReq = CreateObject("MSXML2.ServerXMLHTTP")
    End If
    If objWinHttpReq Is Nothing Then
        Set objWinHttpReq = CreateObject("Microsoft.XMLHTTP")
    End If
    On Error GoTo ErrorHandler
    
    ' オブジェクト作成チェック
    If objWinHttpReq Is Nothing Then
        MsgBox "HTTPリクエストオブジェクトを作成できません", vbCritical
        Exit Sub
    End If
    
    ' ファイルシステムオブジェクトの作成（同様の多重防御）
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error GoTo ErrorHandler
    
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
    failedCount = 0
    
    ' URLリストをループ
    For Each cell In filteredRange
        ' URLの取得（ハイパーリンクと通常のテキストの両方に対応）
        If cell.Hyperlinks.Count > 0 Then
            imgURL = cell.Hyperlinks(1).Address
        Else
            imgURL = cell.Value
        End If
        
        ' 空白セルや文字のみのセルを除外
        If Len(Trim(imgURL)) > 0 Then
            ' 画像形式のチェック
            If IsImageURL(imgURL) Then
                ' エラーハンドリングを強化
                On Error Resume Next
                
                ' タイムアウト設定を追加
                objWinHttpReq.Open "GET", imgURL, False
                objWinHttpReq.SetTimeouts 5000, 5000, 5000, 5000 ' タイムアウト設定（ミリ秒）
                objWinHttpReq.Send
                
                ' 画像の処理
                If objWinHttpReq.Status = 200 Then
                    imgData = objWinHttpReq.ResponseBody
                    base64Str = Base64Encode(imgData)
                    
                    ' MIMEタイプの判定
                    Dim mimeType As String
                    mimeType = GetMimeTypeFromURL(imgURL)
                    
                    ' HTMLに画像を埋め込む
                    htmlContent = htmlContent & "        <div class='image-wrapper'>" & _
                                  "<img src='data:" & mimeType & ";base64," & base64Str & "' alt='Image'/>" & _
                                  "</div>" & vbCrLf
                    successCount = successCount + 1
                Else
                    ' エラーログの追加
                    Debug.Print "画像の取得に失敗: " & imgURL & " (Status: " & objWinHttpReq.Status & ")"
                    failedCount = failedCount + 1
                End If
                
                ' エラーハンドリングをリセット
                On Error GoTo ErrorHandler
            End If
        End If
    Next cell
    
    ' HTMLの終了部分
    htmlContent = htmlContent & "    </div>" & vbCrLf & "</body>" & vbCrLf & "</html>"
    
    ' HTMLファイルとして保存
    filePath = ThisWorkbook.Path & "\image_gallery.html"
    Dim htmlFile As Object
    Set htmlFile = fso.CreateTextFile(filePath, True)
    htmlFile.Write htmlContent
    htmlFile.Close
    
    ' HTMLをブラウザで表示
    OpenHTMLInBrowser filePath
    
    ' 処理結果の表示
    MsgBox "HTMLファイルが生成されました: " & filePath & vbNewLine & _
           "成功した画像: " & successCount & vbNewLine & _
           "失敗した画像: " & failedCount
    
    Exit Sub
    
ErrorHandler:
    Dim errorMsg As String
    errorMsg = "エラーが発生しました。" & vbNewLine & _
               "エラー番号: " & Err.Number & vbNewLine & _
               "エラー説明: " & Err.Description & vbNewLine & _
               "URL: " & imgURL
    
    MsgBox errorMsg, vbCritical
    
    ' デバッグ情報の出力
    Debug.Print errorMsg
End Sub

' 以下の関数は前回と同じ
Function IsImageURL(ByVal url As String) As Boolean
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    
    ' 画像形式の拡張子をチェック（大文字小文字区別なし）
    regEx.Pattern = "\.(jpg|jpeg|png|gif|bmp|webp)(\?.*)?$"
    regEx.IgnoreCase = True
    
    IsImageURL = regEx.Test(url)
End Function

Function GetMimeTypeFromURL(ByVal url As String) As String
    Dim ext As String
    ext = LCase(Right(url, 4))
    
    Select Case ext
        Case ".jpg", "jpeg"
            GetMimeTypeFromURL = "image/jpeg"
        Case ".png"
            GetMimeTypeFromURL = "image/png"
        Case ".gif"
            GetMimeTypeFromURL = "image/gif"
        Case ".bmp"
            GetMimeTypeFromURL = "image/bmp"
        Case "webp"
            GetMimeTypeFromURL = "image/webp"
        Case Else
            GetMimeTypeFromURL = "image/png"  ' デフォルト
    End Select
End Function

Function Base64Encode(ByVal data As Variant) As String
    Dim xmlDoc As Object
    Dim node As Object
    
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    Set node = xmlDoc.CreateElement("Base64Data")
    node.DataType = "bin.base64"
    node.NodeTypedValue = data
    Base64Encode = node.Text
    
    Set node = Nothing
    Set xmlDoc = Nothing
End Function

Sub OpenHTMLInBrowser(filePath As String)
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")
    shell.Run """" & filePath & """"
End Sub
