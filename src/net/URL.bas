Option Explicit

'--------------------------------------------------------------------------------
' 引数の文字列（URLの一部分）をURLエンコードする。
' ※Excel2013以降で対応。
'
' targetString：URLエンコードする文字列。
' return：URLエンコードした文字列。
'--------------------------------------------------------------------------------
Public Function Encode(targetString As String) As String
    Encode = WorksheetFunction.EncodeURL(targetString)
End Function

