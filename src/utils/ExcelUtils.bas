Option Explicit

'--------------------------------------------------------------------------------
' Excelアプリケーションに関する処理をまとめた標準モジュール
'--------------------------------------------------------------------------------

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
        Case Else
            GetExcelVersion = "Unknown Version"
    End Select
End Function

'--------------------------------------------------------------------------------
' 参照設定しているライブラリ名の一覧を返す。
' ※［オプション］から［セキュリティセンター］に入り、
' 「VBAプロジェクトオブジェクトモデルへのアクセスを信頼する」にチェックを入れる。
' http://officetanaka.net/excel/vba/tips/tips100.htm
'
' return：参照設定しているライブラリ名の配列。
'--------------------------------------------------------------------------------
Public Function GetReferences() As String()

    Dim references() As String
    Dim count As Long
    Dim ref As Variant
    
    For Each ref In ActiveWorkbook.VBProject.references
        ReDim Preserve references(count)
        references(count) = ref.Name & "," & ref.Description
        count = count + 1
    Next ref
    
    GetReferences = references
    
End Function

'--------------------------------------------------------------------------------
' Excelが64bitであるかを判定する。
' https://www.ozgrid.com/forum/index.php?thread/137842-determine-32-vs-64-bit-test-compatibility/
'
' return：Excelが64bit版であればTrue、32bit版であればFalseを返す。
'--------------------------------------------------------------------------------
Public Function Is64BitExcel() As Boolean
    #If Win64 Then
        Is64BitExcel = True
    #Else
        Is64BitExcel = False
    #End If
End Function

'--------------------------------------------------------------------------------
' 参照不可のライブラリがないことを検証する。
' ※［オプション］から［セキュリティセンター］に入り、
' 「VBAプロジェクトオブジェクトモデルへのアクセスを信頼する」にチェックを入れる。
' http://officetanaka.net/excel/vba/tips/tips100.htm
'
' return：参照不可のライブラリが存在しなければTrue、
'         参照不可のライブラリが存在すればFalseを返す。
'--------------------------------------------------------------------------------
Public Function ValidateReferences() As Boolean

    Dim ref As Variant
    Dim result As Boolean
    result = True
    
    For Each ref In ActiveWorkbook.VBProject.references
        If ref.IsBroken Then
            result = False
            Exit For
        End If
    Next ref
    
    ValidateReferences = result
End Function