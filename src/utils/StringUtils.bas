Option Explicit

'--------------------------------------------------------------------------------
' 文字列関連のユーティリティー処理をまとめた標準モジュール
'--------------------------------------------------------------------------------

'--------------------------------------------------------------------------------
' 大文字/小文字、全角/半角の区別なく、2つの文字列を辞書的に比較する。
'
' targetString：比較対象の文字列。
' anotherString：もう一方の比較対象の文字列。
' return：比較対象の文字列が等しい場合はtrueを返す。
'--------------------------------------------------------------------------------
Public Function Compare(targetString As String, anotherString As String) As Boolean
    ' 2つの文字列を「大文字の半角文字」に直して比較する。
    Compare = (StrConv(UCase(targetString), vbNarrow) = StrConv(UCase(anotherString), vbNarrow))
End Function

'--------------------------------------------------------------------------------
' 探索対象の文字列中に、探索する部分文字列が存在するかを調べる。探索する部分文字列が空文字の場合もtrueを返す。
'
' targetString：探索対象の文字列
' substring：探索する部分文字列。
' return：探索対象の文字列中に、部分文字列が存在すればtrueを返す。
'--------------------------------------------------------------------------------
Public Function Contains(targetString As String, substring As String) As Boolean
    Contains = (InStr(targetString, substring) > 0)
End Function

'--------------------------------------------------------------------------------
' 探索対象の文字列中に、1ないし複数個の部分文字列が存在するかを調べる。
'
' targetString：対象の文字列。
' searchTexts：探索する部分文字列。1ないし複数個を指定できる。
' return：探索対象の文字列中に、部分文字列が存在すればtrueを返す。
'--------------------------------------------------------------------------------
Function ContainsAny(targetString As String, ParamArray searchTexts() As Variant) As Boolean
    Dim txt As Variant
    Dim result As Boolean
    For Each txt In searchTexts
        If (InStr(targetString, CStr(txt)) > 0) Then
            result = True
            Exit For
        End If
    Next txt

    ContainsAny = result
End Function

'--------------------------------------------------------------------------------
' 引数の文字列中に、探索文字列が何回出現したかを返す。
'
' targetString：対象の文字列。
' subStr：探索する部分文字列。
' return：探索文字列が出現した回数を返す。
'--------------------------------------------------------------------------------
Public Function CountMatches(targetString As String, subStr As String) As String
    MsgBox Replace(targetString, subStr, "")
    CountMatches = (Len(targetString) - Len(Replace(targetString, subStr, ""))) / Len(subStr)
End Function

'--------------------------------------------------------------------------------
' 引数の文字列から空白を除去する。
'
' targetString：比較対象の文字列。
' return：文字列の前後、および文字列中から以下の文字を除去した文字列を返す。
'         全角空白、半角空白、タブ、改行コード（CR,LF,CRLF）
'--------------------------------------------------------------------------------
Public Function DeleteWhitespace(targetString As String) As String
    ' 正規表現オブジェクトの作成
    Dim reg As Object
    Set reg = CreateObject("VBScript.RegExp")

    ' 正規表現のオプションの指定
    With reg
        .Pattern = "　| |\t|\r|\n"
        .IgnoreCase = False         '大文字と小文字を区別する
        .Global = True              '文字列全体を検索する
    End With

    DeleteWhitespace = reg.Replace(targetString, "")
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
' 引数の文字列を改行コード（行終了記号）で区切った配列として返す。
'
' targetString：対象の文字列。
' return：改行コードの単位で区切られた配列。
'--------------------------------------------------------------------------------
Public Function Lines(targetString) As String()
    ' 正規表現オブジェクトの作成
    Dim reg As Object
    Set reg = CreateObject("VBScript.RegExp")

    '正規表現の指定（CRLF/CR/LF）
    With reg
        .Pattern = "\r\n|\r|\n"
        .IgnoreCase = False         '大文字と小文字を区別する
        .Global = True              '文字列全体を検索する
    End With

    ' CRLF/CR/LFをCRLFに統一する
    Dim buff As String
    buff = reg.Replace(targetString, vbCrLf)

    Lines = Split(buff, vbCrLf)
End Function

'--------------------------------------------------------------------------------
' 引数の文字列を指定した回数繰り返す。
'
' targetString：対象の文字列。
' repeatCount：繰り返しの回数。1以上の値を指定する。
' return：引数の文字列を指定した回数繰り返した結果の文字列を返す。
'--------------------------------------------------------------------------------
Public Function Repeat(targetString As String, repeatCount As Integer) As String
    Dim buff As String
    Dim i As Integer

    For i = 1 To repeatCount
        buff = buff + targetString
    Next i

    Repeat = buff
End Function

'--------------------------------------------------------------------------------
' 引数の文字列を循環シフト（circular shift）させる。
'
' StringUtils.Rotate("abcdefg", 0) => "abcdefg"
' StringUtils.Rotate("abcdefg", 2) => "fgabcde"
' StringUtils.Rotate("abcdefg", -2) => "cdefgab"
'
' targetString：循環シフトさせる文字列。
' shift：シフトさせる文字数。正の値なら右循環シフト、負の値なら左循環シフトとなる。
'--------------------------------------------------------------------------------
Public Function Rotate(targetString As String, shift As Long) As String
    
    If (shift > 0) Then
        ' 正の値なら右循環シフト。
        Rotate = Right(targetString, shift) & Left(targetString, Len(targetString) - shift)
    
    ElseIf (shift < 0) Then
        ' 負の値なら左循環シフト。
        Rotate = Right(targetString, Len(targetString) - Abs(shift)) & Left(targetString, Abs(shift))
    Else
        ' shift=0の時は、引数の文字列をそのまま返す。
        Rotate = targetString
    End If
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
' 引数の文字列を文字配列に変換して返す。
'
' targetString：対象の文字列。
' return：文字配列。
'--------------------------------------------------------------------------------
Public Function ToCharArray(targetString As String) As String()
​
    Dim aryLength As Long
    Dim charAry() As String

    aryLength = Len(targetString)
    ReDim charAry(aryLength - 1)

    Dim i As Long
    For i = 1 To aryLength
        charAry(i - 1) = Mid(targetString, i, 1)
    Next i

    ToCharArray = charAry
End Function

'--------------------------------------------------------------------------------
' 引数の文字列を指定した文字数に切り詰める。
'
' targetString：対象の文字列。
' maxLength：文字列の最大文字数。この文字数を超えた部分は削除される。
' return：引数の文字列を指定した文字数に切り詰めた文字列を返す。
'--------------------------------------------------------------------------------
Public Function Truncate(targetString As String, maxLength As Integer) As String
    Truncate = Left(targetString, maxLength)
End Function