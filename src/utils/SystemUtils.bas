Option Explicit

'--------------------------------------------------------------------------------
' ホスト名を取得する。
'
' return：ホスト名。
'--------------------------------------------------------------------------------
Public Function GetHostName() As String
    Dim netObj As Object
    Set netObj = CreateObject("WScript.Network")
    GetHostName = netObj.ComputerName
End Function

'--------------------------------------------------------------------------------
' ログインユーザー名を取得する。
'
' return：ログインユーザー名。
'--------------------------------------------------------------------------------
Public Function GetUserName() As String
    Dim netObj As Object
    Set netObj = CreateObject("WScript.Network")
    GetUserName = netObj.UserName
End Function