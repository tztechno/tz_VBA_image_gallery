Sub FetchImagesAndGenerateHTML()
    Dim ws As Worksheet
    Dim xmlHttp As Object
    Dim imgData() As Byte
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
    
    ' Microsoft XML HTTP要求の作成
    Set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    
    ' HTMLの開始部分 (前回と同様)
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
                xmlHttp.Open "GET", imgURL, False
                xmlHttp.SetTimeouts 5000, 5000, 5000, 5000 ' タイムアウト設定（ミリ秒）
                xmlHttp.Send
                
                ' 画像の処理
                If xmlHttp.Status = 200 Then
                    imgData = xmlHttp.ResponseBody
                    base64Str = ConvertToBase64(imgData)
                    
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
                    Debug.Print "画像の取得に失敗: " & imgURL & " (Status: " & xmlHttp.Status & ")"
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
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open filePath For Output As #fileNum
    Print #fileNum, htmlContent
    Close #fileNum
    
    ' シェル関数を使用してブラウザで開く
    Call Shell("cmd.exe /c start " & filePath, vbNormalFocus)
    
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

' 画像URL判定関数
Function IsImageURL(ByVal url As String) As Boolean
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    
    ' 画像形式の拡張子をチェック（大文字小文字区別なし）
    regEx.Pattern = "\.(jpg|jpeg|png|gif|bmp|webp)(\?.*)?$"
    regEx.IgnoreCase = True
    
    IsImageURL = regEx.Test(url)
End Function

' MIMEタイプ取得関数
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

' Base64変換関数
Function ConvertToBase64(ByRef data() As Byte) As String
    ' Base64エンコーディングをVBA純粋実装に変更
    Const BASE64_CHARS As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
    Dim i As Long
    Dim j As Long
    Dim SrcLen As Long
    Dim ret As String
    Dim c1 As Long, c2 As Long, c3 As Long
    
    SrcLen = UBound(data) + 1
    
    For i = 0 To SrcLen - 1 Step 3
        ' 3バイト読み取り
        c1 = data(i)
        If i + 1 < SrcLen Then
            c2 = data(i + 1)
        Else
            c2 = 0
        End If
        If i + 2 < SrcLen Then
            c3 = data(i + 2)
        Else
            c3 = 0
        End If
        
        ' Base64エンコーディング
        ret = ret & Mid(BASE64_CHARS, ((c1 And &HFC) \ 4) + 1, 1)
        ret = ret & Mid(BASE64_CHARS, (((c1 And &H3) * 16) + ((c2 And &HF0) \ 16)) + 1, 1)
        
        If i + 1 < SrcLen Then
            ret = ret & Mid(BASE64_CHARS, (((c2 And &HF) * 4) + ((c3 And &HC0) \ 64)) + 1, 1)
        Else
            ret = ret & "="
        End If
        
        If i + 2 < SrcLen Then
            ret = ret & Mid(BASE64_CHARS, (c3 And &H3F) + 1, 1)
        Else
            ret = ret & "="
        End If
    Next i
    
    ConvertToBase64 = ret
End Function
