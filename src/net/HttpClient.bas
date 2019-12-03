Option Explicit

​'--------------------------------------------------------------------------------
' HTTP通信用クラス。
'--------------------------------------------------------------------------------
​
' HTTP通信用オブジェクト
Private httpObj As Object
​
'--------------------------------------------------------------------------------
' コンストラクタ
'--------------------------------------------------------------------------------
Public Sub Class_Initialize()
    'Set httpObj = CreateObject("MSXML2.XMLHTTP")           ' TLS1.2に非対応
    Set httpObj = CreateObject("MSXML2.ServerXMLHTTP")
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
    httpObj.Send

    ' readyState=4で読み込みが完了
    Do While httpObj.readyState < 4
        DoEvents
    Loop

    GetPage = httpObj.responseText
End Function