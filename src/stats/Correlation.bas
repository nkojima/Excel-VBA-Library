Option Explicit

'------------------------------------------------------------------------------
' 相関係数を求める処理をまとめた標準モジュール
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' Pearsonの積率相関係数を求める。
'
' x: 相関係数を求める値（系列1）。
' y: 相関係数を求める値（系列2）。
' return: 相関係数r（-1<=r<=1）。エラー時には-100を返す。
'------------------------------------------------------------------------------
Public Function CalcPearson(x() As Double, y() As Double) As Double
    
    ' 2つの系列（配列）の長さが違う場合は、相関係数を求められないのでエラーとする。
    Dim xLength As Long, yLength As Long
    xLength = UBound(x) - LBound(x) + 1
    yLength = UBound(y) - LBound(y) + 1
    
    On Error GoTo CORR_ERROR
    
    If (xLength <> yLength) Then
        Err.Raise 65535, "CalcPearson", "引数で指定した配列の長さが異なっているため、相関係数を求められません。"
    End If
    
    CalcPearson = WorksheetFunction.Pearson(x, y)
    Exit Function
    
CORR_ERROR:
    ' エラーが起きた場合、エラーメッセージを表示した後に-100を返す。
    Debug.Print "ErrorNumber:" & Err.Number & vbCrLf & _
            "Source:" & Err.Source & vbCrLf & _
           "Description:" & Err.Description
    CalcPearson = -100
End Function
