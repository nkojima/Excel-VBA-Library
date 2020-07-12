Option Explicit

'--------------------------------------------------------------------------------
' 数値関連のユーティリティー処理をまとめた標準モジュール
'--------------------------------------------------------------------------------

'--------------------------------------------------------------------------------
' 引数の中から最小値を求める。
' ※ワークシート関数の「Min」に相当する処理が、関数やメソッドとして存在しないため。
'
' numbers：可変長引数で表される整数値のリスト。
' return：配列中の最小値を返す。
'--------------------------------------------------------------------------------
Public Function Min(ParamArray numbers() As Variant) As Long

    Dim minValue As Long
    minValue = (2 ^ 31) - 1     ' Long型の最大値を初期値として与える

    Dim num As Variant
    For Each num In numbers
        If (minValue > CLng(num)) Then
            minValue = num
        End If
    Next num
    
    Min = minValue
End Function

'--------------------------------------------------------------------------------
' 引数の中から最大値を求める。
' ※ワークシート関数の「Max」に相当する処理が、関数やメソッドとして存在しないため。
'
' numbers：可変長引数で表される整数値のリスト。
' return：配列中の最大値を返す。
'--------------------------------------------------------------------------------
Public Function Max(ParamArray numbers() As Variant) As Long

    Dim maxValue As Long
    maxValue = -(2 ^ 31)      ' Long型の最小値を初期値として与える

    Dim num As Variant
    For Each num In numbers
        If (maxValue < CLng(num)) Then
            maxValue = num
        End If
    Next num
    
    Max = maxValue
End Function