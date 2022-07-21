Option Explicit

'------------------------------------------------------------------------------
' 類似度指数を求める処理をまとめた標準モジュール
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' Jaccardの類似度指数を求める。
'
' m: 2地点に共通して出現する種数。
' a: 地点Aのみに出現する種数。
' b: 地点Bのみに出現する種数。
' return: 2地点の類似度指数（完全に一致：1、全く異なる：0）。エラー時は-100を返す。
'------------------------------------------------------------------------------
Public Function CalcJaccard(m As Integer, a As Integer, b As Integer) As Double
    
    ' 種数が負の値0である時はエラーとする。
    On Error GoTo JACCARD_ERROR
    
    If (m < 0) Then
        Err.Raise 65535, "CalcPearson", "引数で指定した「2地点に共通して出現する種数m」が負の値のため、類似度指数を求められません。"
    ElseIf (a < 0) Then
        Err.Raise 65535, "CalcPearson", "引数で指定した「地点Aのみに出現する種数a」が負の値のため、類似度指数を求められません。"
    ElseIf (b < 0) Then
        Err.Raise 65535, "CalcPearson", "引数で指定した「地点Bのみに出現する種数b」が負の値のため、類似度指数を求められません。"
    End If
        
    CalcJaccard = m / (m + a + b)
    Exit Function
    
JACCARD_ERROR:
    ' エラーが起きた場合、エラーメッセージを表示した後に-100を返す。
    Debug.Print "ErrorNumber:" & Err.Number & vbCrLf & _
            "Source:" & Err.Source & vbCrLf & _
           "Description:" & Err.Description
    CalcJaccard = -100
End Function

'------------------------------------------------------------------------------
' Sorensenの類似度指数を求める。
'
' m: 2地点に共通して出現する種数。
' a: 地点Aのみに出現する種数。
' b: 地点Bのみに出現する種数。
' return: 2地点の類似度指数（完全に一致：1、全く異なる：0）。エラー時は-100を返す。
'------------------------------------------------------------------------------
Public Function CalcSorensen(m As Integer, a As Integer, b As Integer) As Double
    
    ' 種数が負の値0である時はエラーとする。
    On Error GoTo SORENSEN_ERROR
    
    If (m < 0) Then
        Err.Raise 65535, "CalcPearson", "引数で指定した「2地点に共通して出現する種数m」が負の値のため、類似度指数を求められません。"
    ElseIf (a < 0) Then
        Err.Raise 65535, "CalcPearson", "引数で指定した「地点Aのみに出現する種数a」が負の値のため、類似度指数を求められません。"
    ElseIf (b < 0) Then
        Err.Raise 65535, "CalcPearson", "引数で指定した「地点Bのみに出現する種数b」が負の値のため、類似度指数を求められません。"
    End If
        
    CalcSorensen = (2 * m) / (2 * m + a + b)
    Exit Function
    
SORENSEN_ERROR:
    ' エラーが起きた場合、エラーメッセージを表示した後に-100を返す。
    Debug.Print "ErrorNumber:" & Err.Number & vbCrLf & _
            "Source:" & Err.Source & vbCrLf & _
           "Description:" & Err.Description
    CalcSorensen = -100
End Function
