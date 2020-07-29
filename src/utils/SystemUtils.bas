Option Explicit

'--------------------------------------------------------------------------------
' CPUの論理コア数（スレッド数）を取得する。
'
' return：CPUの論理コア数（スレッド数）。
'--------------------------------------------------------------------------------
Public Function GetCpuCoreCount()
    GetCpuCoreCount = Environ("NUMBER_OF_PROCESSORS")
End Function

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
        Else
            GetExcelVersion = "Unknown Version"
    End Select
End Function

'--------------------------------------------------------------------------------
' ホスト名を取得する。
'
' return：ホスト名。
'--------------------------------------------------------------------------------
Public Function GetHostName() As String
    GetHostName = GetHostName = Environ("COMPUTERNAME")
End Function

'--------------------------------------------------------------------------------
' 環境変数のPATH（コマンド検索パス）をセミコロンで区切って、フォルダパスの配列として返す。
'
' return：環境変数のPATHを構成しているフォルダパスの配列。
'--------------------------------------------------------------------------------
Public Function GetPathArray() As String()
    GetPathArray = Split(Environ("PATH"), ";")
End Function

'--------------------------------------------------------------------------------
' 環境変数のPATH（コマンド検索パス）そのものを文字列として返す。
'
' return：環境変数のPATH。
'--------------------------------------------------------------------------------
Public Function GetPathString() As String
    GetPathString = Environ("PATH")
End Function

'--------------------------------------------------------------------------------
' ログインユーザー名を取得する。
'
' return：ログインユーザー名。
'--------------------------------------------------------------------------------
Public Function GetUserName() As String
    GetUserName = GetUserName = Environ("USERNAME")
End Function