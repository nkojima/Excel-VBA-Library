Option Explicit

'------------------------------------------------------------------------------
' 指定したシートの最終列の列番号を取得する。
' xlToLeftだと「見えている列の最終列」になってしまうため、事前に列の非表示を解除する。
'
' sheetName: シート名。
' return: 最終列の列番号。
'------------------------------------------------------------------------------
Public Function GetLastColumn(sheetName As String) As Long
    GetLastColumn = Worksheets(sheetName).Cells(1, Columns.Count).End(xlToLeft).Column    ' 1行目の最終列を取得する。
End Function

'------------------------------------------------------------------------------
' 指定したシートの最終行の行番号を取得する。
' xlUpだと「見えている行の最終行」になってしまうため、事前に行の非表示を解除する。
'
' sheetName: シート名。
' return: 最終行の行番号。
'------------------------------------------------------------------------------
Public Function GetLastRow(sheetName As String) As Long
    GetLastRow = Worksheets(sheetName).Cells(Rows.Count, 1).End(xlUp).Row  ' 1列目の最終行を取得する。
End Function

'--------------------------------------------------------------------------------
' Excelの列番号をR1C1形式からA1形式に変換する。
'
' r1c1ColIdx：R1C1形式の列番号。
' return：Excelのバージョン。
'--------------------------------------------------------------------------------
Public Function ToA1(r1c1ColIdx As Long) As String
    Dim remainder As Integer, quotinent As Long
    remainder = (r1c1ColIdx - 1) Mod 26
    quotinent = Int((r1c1ColIdx - 1) / 26)  ' 除算の結果が小数になるケースがあるため、Int関数で切り捨てる
    
    If (quotinent > 0) Then
        ToA1 = ToA1(quotinent) + ToColAlphabet(remainder)
    Else
        ToA1 = ToColAlphabet(remainder)
    End If
End Function

'--------------------------------------------------------------------------------
' 1～26の整数を、A～Zのアルファベットに変換する。
'
' idx：1～26の整数。
' return：A～Zのアルファベット。
'--------------------------------------------------------------------------------
Private Function ToColAlphabet(idx As Integer) As String
    ToColAlphabet = Chr(idx + 65) ' Chr(65)は'A'となる。
End Function

'--------------------------------------------------------------------------------
' Excelの列番号をA1形式からR1C1形式に変換する。
'
' a1ColIdx：A1形式の列番号。
' return：Excelのバージョン。
'--------------------------------------------------------------------------------
Public Function ToR1C1(a1ColIdx As String) As Long
    Dim substr As String
    substr = a1ColIdx
    
    Dim i As Integer
    
    For i = 1 To Len(a1ColIdx)
        If (i = 1) Then
            ToR1C1 = ToR1C1 + ToColNumber(Right(substr, 1))
        Else
            ToR1C1 = ToR1C1 + (ToColNumber(Right(substr, 1)) * (i - 1) * 26)
        End If
        substr = Left(substr, Len(substr) - 1)
    Next i
End Function

'--------------------------------------------------------------------------------
' A～Zのアルファベットを、1～26の整数に変換する。
'
' idx：A～Zのアルファベット。
' return：1～26の整数。
'--------------------------------------------------------------------------------
Private Function ToColNumber(idx As String) As Integer
    ToColNumber = Asc(idx) - 65 + 1 ' Asc("A")は65となる。
End Function
