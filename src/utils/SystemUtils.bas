Option Explicit

'--------------------------------------------------------------------------------
' Excelのバージョンを取得する。
' Office365の場合は、購入時のバージョンとなる。
' https://ja.wikipedia.org/wiki/Microsoft_Excel#%E6%AD%B4%E5%8F%B2
' https://answers.microsoft.com/ja-jp/msoffice/forum/all/office365%E3%81%AEapplicationversion%E3%81%AB/3c406a7e-831e-4bda-bdf0-564f5bfa88f0
'
' return：Excelのバージョン。
'--------------------------------------------------------------------------------
Public Function GetExcelVersion() As String
    Dim version As String
    version = Application.version
    
    Select Case version
        Case "16.0"
            ' Excel2019もVersionが16.0なので、Excel2016として判定されてしまう。
            GetExcelVersion = "Excel 2016"
        Case "15.0"
            GetExcelVersion = "Excel 2013"
        Case "14.0"
            GetExcelVersion = "Excel 2010"
        Case "12.0"
            GetExcelVersion = "Excel 2007"
        Case "11.0"
            GetExcelVersion = "Excel 2003"
        Case "10.0"
            GetExcelVersion = "Excel 2002"
        Case "9.0"
            GetExcelVersion = "Excel 2000"
    End Select
End Function

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