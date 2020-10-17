Option Explicit

'--------------------------------------------------------------------------------
' HTTP通信用クラス。
'--------------------------------------------------------------------------------

' HTTP通信用オブジェクト
Private httpObj As Object

'--------------------------------------------------------------------------------
' コンストラクタ
'--------------------------------------------------------------------------------
Public Sub Class_Initialize()
    Set httpObj = CreateObject("MSXML2.ServerXMLHTTP")    ' TLS1.2に対応
End Sub
​
'--------------------------------------------------------------------------------
' デストラクタ
'--------------------------------------------------------------------------------
Public Sub Class_Terminate()
    Set httpObj = Nothing
End Sub
​
'--------------------------------------------------------------------------------
' 引数のURLをGETメソッドで取得する。
'
' url：URL文字列。
' return：取得したページ。
'--------------------------------------------------------------------------------
Public Function GetPage(url As String) As String
    httpObj.Open "GET", url
    httpObj.send

    ' readyState=4で読み込みが完了
    Do While httpObj.readyState < 4
        DoEvents
    Loop

    Dim statusCode As Integer
    statusCode = httpObj.Status
    
    ' HTTPのステータスコードが200(OK)以外であれば、ステータスコードなどを返す。
    If (statusCode = 200) Then
        'GetPage = httpObj.responseText ' レスポンスの文字コードがShift_JIS(MS932)の時はこちらを使う。
        GetPage = StrConv(httpObj.responseBody, vbUnicode)
    Else
        GetPage = "HTTP StatusCode:" & statusCode & ", HTTP StatusText:" & httpObj.statusText
    End If

End Function

'--------------------------------------------------------------------------------
' 引数のURLにPostメソッドで送信する。
'
' url：URL文字列。
' urlParams：URLパラメーター。
' return：レスポンスの文字列。
'--------------------------------------------------------------------------------
Public Function PostContents(url As String, urlParams As String) As String
    httpObj.Open "POST", url, False
    httpObj.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    httpObj.send (urlParams)
    
    ' readyState=4で読み込みが完了
    Do While httpObj.readyState < 4
        DoEvents
    Loop

    Dim statusCode As Integer
    statusCode = httpObj.Status

    ' HTTPのステータスコードが200(OK)以外であれば、ステータスコードなどを返す。
    If (statusCode = 200) Then
        'PostContents = httpObj.responseText ' レスポンスの文字コードがShift_JIS(MS932)の時はこちらを使う。
        PostContents = StrConv(httpObj.responsebody, vbUnicode)
    Else
        PostContents = "HTTP StatusCode:" & statusCode & ", HTTP StatusText:" & httpObj.statusText
    End If
End Function
