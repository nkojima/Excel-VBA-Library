Option Explicit

'--------------------------------------------------------------------------------
' 文字列関連のユーティリティー処理をまとめた標準モジュール
'--------------------------------------------------------------------------------

'--------------------------------------------------------------------------------
' 探索対象の文字列中に、探索する部分文字列が存在するかを調べる。探索する部分文字列が空文字の場合もtrueを返す。
'
' targetString：探索対象の文字列
' substring：探索する部分文字列。
' return：探索対象の文字列中に、探索する文字列が存在すればtrueを返す。
'--------------------------------------------------------------------------------
Public Function Contains(targetString As String, substring As String) As Boolean
    Contains = (InStr(targetString, substring) > 0)
End Function

'--------------------------------------------------------------------------------
' 探索対象の文字列が、指定された接頭辞で始まるかを判定する。
'
' targetString：探索対象の文字列
' prefix：探索する接頭辞。
' return：探索対象の文字列が、指定された接頭辞で始まればtrueを返す。接頭辞が空文字の場合もtrueを返す。
'--------------------------------------------------------------------------------
Public Function StartsWith(targetString As String, prefix As String) As Boolean
    StartsWith = (InStr(targetString, prefix) = 1)
End Function

'--------------------------------------------------------------------------------
' 探索対象の文字列が、指定された接尾辞で終わるかを判定する。
'
' targetString：探索対象の文字列
' suffix：探索する接尾辞。
' return：探索対象の文字列が、指定された接尾辞で終わればtrueを返す。接尾辞が空文字の場合もtrueを返す。
'--------------------------------------------------------------------------------
Public Function EndsWith(targetString As String, suffix As String) As Boolean
    If (Len(targetString) >= Len(suffix)) Then
        ' 文字数が【探索対象の文字列】≧【接尾辞】の時、探索対象の文字列の末尾が接尾辞と一致すればtrueを返す。
        Dim subStr As String
        subStr = Right(targetString, Len(suffix))
        EndsWith = (subStr = suffix)
    Else
        ' 文字数が【探索対象の文字列】＜【接尾辞】の時、falseを返す。
        EndsWith = False
    End If
End Function

'--------------------------------------------------------------------------------
' 大文字/小文字、全角/半角の区別なく、2つの文字列を辞書的に比較する。
'
' targetString：比較対象の文字列。
' anotherString：もう一方の比較対象の文字列。
' return：比較対象の文字列が等しい場合はtrueを返す。
' see: https://www.moug.net/tech/exvba/0140044.html
'--------------------------------------------------------------------------------
Public Function Compare(targetString As String, anotherString As String) As Boolean
    ' 2つの文字列を「大文字の半角文字」に直して比較する。
    Compare = (StrConv(UCase(targetString), vbNarrow) = StrConv(UCase(anotherString), vbNarrow))
End Function