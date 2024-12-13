Option Explicit

' WinHttpリクエストを使用するためのAPIデクレアレーション
Private Declare PtrSafe Function WinHttpOpen Lib "winhttp.dll" ( _
    ByVal pszAgentName As LongPtr, _
    ByVal dwAccessType As Long, _
    ByVal pszProxyName As LongPtr, _
    ByVal pszProxyBypass As LongPtr, _
    ByVal dwFlags As Long _
) As LongPtr

Private Declare PtrSafe Function WinHttpConnect Lib "winhttp.dll" ( _
    ByVal hSession As LongPtr, _
    ByVal pswzServerName As LongPtr, _
    ByVal nServerPort As Long, _
    ByVal dwReserved As Long _
) As LongPtr

Private Declare PtrSafe Function WinHttpOpenRequest Lib "winhttp.dll" ( _
    ByVal hConnect As LongPtr, _
    ByVal pwszVerb As LongPtr, _
    ByVal pwszObjectName As LongPtr, _
    ByVal pwszVersion As LongPtr, _
    ByVal pwszReferrer As LongPtr, _
    ByVal pwszAcceptTypes As LongPtr, _
    ByVal dwFlags As Long _
) As LongPtr

Private Declare PtrSafe Function WinHttpSendRequest Lib "winhttp.dll" ( _
    ByVal hRequest As LongPtr, _
    ByVal pwszHeaders As LongPtr, _
    ByVal dwHeadersLength As Long, _
    ByVal lpOptional As LongPtr, _
    ByVal dwOptionalLength As Long, _
    ByVal dwTotalLength As Long, _
    ByVal dwContext As LongPtr _
) As Long

Private Declare PtrSafe Function WinHttpReceiveResponse Lib "winhttp.dll" ( _
    ByVal hRequest As LongPtr, _
    ByVal lpReserved As LongPtr _
) As Long

Private Declare PtrSafe Function WinHttpQueryDataAvailable Lib "winhttp.dll" ( _
    ByVal hRequest As LongPtr, _
    ByRef lpdwNumberOfBytesAvailable As Long _
) As Long

Private Declare PtrSafe Function WinHttpReadData Lib "winhttp.dll" ( _
    ByVal hRequest As LongPtr, _
    ByRef lpBuffer As Any, _
    ByVal dwNumberOfBytesRead As Long, _
    ByRef lpdwNumberOfBytesRead As Long _
) As Long

Private Declare PtrSafe Function WinHttpCloseHandle Lib "winhttp.dll" ( _
    ByVal hInternet As LongPtr _
) As Long

Private Const WINHTTP_ACCESS_TYPE_DEFAULT_PROXY = 0
Private Const WINHTTP_FLAG_SECURE = &H800000
Private Const INTERNET_DEFAULT_HTTPS_PORT = 443

Sub FetchImagesAndGenerateHTML()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim filteredRange As Range
    Dim cell As Range
    Dim imgURL As String
    Dim htmlContent As String
    Dim successCount As Long
    Dim failedCount As Long
    Dim filePath As String
    
    ' エラーハンドリングの設定
    On Error GoTo ErrorHandler
    
    ' シートとオブジェクトの初期化
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
                Dim base64Data As String
                base64Data = FetchBase64Image(imgURL)
                
                If base64Data <> "" Then
                    ' HTMLに画像を埋め込む
                    htmlContent = htmlContent & "        <div class='image-wrapper'>" & _
                                  "<img src='data:" & GetMimeTypeFromURL(imgURL) & ";base64," & base64Data & "' alt='Image'/>" & _
                                  "</div>" & vbCrLf
                    successCount = successCount + 1
                Else
                    ' エラーログの追加
                    Debug.Print "画像の取得に失敗: " & imgURL
                    failedCount = failedCount + 1
                End If
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
               "エラー説明: " & Err.Description
    
    MsgBox errorMsg, vbCritical
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
Function Base64Encode(ByRef data() As Byte) As String
    Const BASE64_CHARS As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
    Dim i As Long
    Dim ret As String
    Dim c1 As Long, c2 As Long, c3 As Long
    Dim SrcLen As Long
    
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
    
    Base64Encode = ret
End Function

' 画像フェッチとBase64変換関数
Function FetchBase64Image(ByVal url As String) As String
    Dim hSession As LongPtr
    Dim hConnect As LongPtr
    Dim hRequest As LongPtr
    Dim bytesAvailable As Long
    Dim bytesRead As Long
    Dim buffer() As Byte
    Dim fullData() As Byte
    Dim totalBytes As Long
    Dim parsedURL As Object
    
    ' URLをパース
    Set parsedURL = ParseURL(url)
    
    ' WinHttpセッションを開く
    hSession = WinHttpOpen(StrPtr("VBA Image Fetcher"), _
                           WINHTTP_ACCESS_TYPE_DEFAULT_PROXY, _
                           0, 0, 0)
    
    If hSession = 0 Then Exit Function
    
    ' サーバーに接続
    hConnect = WinHttpConnect(hSession, _
                              StrPtr(parsedURL("host")), _
                              INTERNET_DEFAULT_HTTPS_PORT, 0)
    
    If hConnect = 0 Then
        WinHttpCloseHandle hSession
        Exit Function
    End If
    
    ' リクエストを開く
    hRequest = WinHttpOpenRequest(hConnect, _
                                  StrPtr("GET"), _
                                  StrPtr(parsedURL("path")), _
                                  StrPtr("HTTP/1.1"), _
                                  0, 0, _
                                  WINHTTP_FLAG_SECURE)
    
    If hRequest = 0 Then
        WinHttpCloseHandle hConnect
        WinHttpCloseHandle hSession
        Exit Function
    End If
    
    ' リクエストを送信
    If WinHttpSendRequest(hRequest, 0, 0, 0, 0, 0, 0) = 0 Then
        WinHttpCloseHandle hRequest
        WinHttpCloseHandle hConnect
        WinHttpCloseHandle hSession
        Exit Function
    End If
    
    ' レスポンスを受信
    If WinHttpReceiveResponse(hRequest, 0) = 0 Then
        WinHttpCloseHandle hRequest
        WinHttpCloseHandle hConnect
        WinHttpCloseHandle hSession
        Exit Function
    End If
    
    ' データを読み取り
    Do
        If WinHttpQueryDataAvailable(hRequest, bytesAvailable) = 0 Then Exit Do
        
        If bytesAvailable > 0 Then
            ReDim buffer(bytesAvailable - 1)
            If WinHttpReadData(hRequest, buffer(0), bytesAvailable, bytesRead) = 0 Then Exit Do
            
            ReDim Preserve fullData(totalBytes + bytesRead - 1)
            CopyMemory ByVal VarPtr(fullData(totalBytes)), ByVal VarPtr(buffer(0)), bytesRead
            totalBytes = totalBytes + bytesRead
        End If
    Loop While bytesAvailable > 0
    
    ' ハンドルを閉じる
    WinHttpCloseHandle hRequest
    WinHttpCloseHandle hConnect
    WinHttpCloseHandle hSession
    
    ' データが取得できない場合
    If totalBytes = 0 Then Exit Function
    
    ' Base64エンコード
    FetchBase64Image = Base64Encode(fullData)
End Function

' URLパース関数
Function ParseURL(ByVal url As String) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    
    ' URLからホストとパスを分離
    Dim parts() As String
    parts = Split(Replace(url, "https://", ""), "/", 2)
    
    result("host") = parts(0)
    
    If UBound(parts) > 0 Then
        result("path") = "/" & parts(1)
    Else
        result("path") = "/"
    End If
    
    ParseURL = result
End Function

' メモリコピー用のAPIデクレアレーション
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    Destination As Any, _
    Source As Any, _
    ByVal Length As LongPtr _
)
