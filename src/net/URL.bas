Option Explicit

'--------------------------------------------------------------------------------
' 引数の文字列をURLエンコードする。
' ※Excel2013以降で対応。
'
' url：URLエンコードする文字列。
' return：URLエンコードした文字列。
'--------------------------------------------------------------------------------
Public Function Encode(url As String) As String
    Encode = WorksheetFunction.EncodeURL(url)
End Function

