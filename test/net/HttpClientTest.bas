Option Explicit
​
Public Sub TestHttpClient()
    Dim httpObj As HttpClient
    Set httpObj = New HttpClient

    Dim response As String
    response = httpObj.GetPage("https://www8.cao.go.jp/chosei/shukujitsu/syukujitsu.csv")
    Debug.Print response
End Sub