Option Explicit

'------------------------------------------------------------------------------
' 指定したシートの最終列の列番号を取得する。
' xlToLeftだと「見えている列の最終列」になってしまうため、事前に列の非表示を解除する。
'
' sheetName: シート名。
' return: 最終列の列番号。
'------------------------------------------------------------------------------

Function GetLastColumn(sheetName As String) As Long

    Dim lastCol As Long
    lastCol = Worksheets(sheetName).Cells(1, Columns.count).End(xlToLeft).Column

    GetLastColumn = lastCol

End Function

'------------------------------------------------------------------------------
' 指定したシートの最終行の行番号を取得する。
' xlUpだと「見えている行の最終行」になってしまうため、事前に行の非表示を解除する。
'
' sheetName: シート名。
' return: 最終行の行番号。
'------------------------------------------------------------------------------
Function GetLastRow(sheetName As String) As Long

    Dim lastRow As Long
    lastRow = Worksheets(sheetName).Cells(Rows.count, 1).End(xlUp).Row

    GetLastRow = lastRow

End Function