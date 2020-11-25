Option Explicit

'--------------------------------------------------------------------------------
' 配列に関する処理をまとめた標準モジュール
'--------------------------------------------------------------------------------

'--------------------------------------------------------------------------------
' 配列の要素数を求める。
'
' ary：対象となる配列。
' return：配列の要素数。引数として初期化されていない配列を指定した時は-1、配列以外を指定した時は-100を返す。
'--------------------------------------------------------------------------------
Function GetLength(ary As Variant) As Integer
    If (IsArray(ary)) Then
        If (IsInitialized(ary)) Then
            GetLength = UBound(ary) - LBound(ary) + 1
        Else
            GetLength = -1
        End If
    Else
        GetLength = -100
    End If

End Function

'--------------------------------------------------------------------------------
' 配列が初期化されているかをチェックする。
'
' ary：対象となる配列。
' return：配列が初期化済みならTrue、そうでなければFalseを返す。
'--------------------------------------------------------------------------------
Function IsInitialized(ary As Variant) As Boolean
    On Error GoTo NOT_INITIALIZED_ERROR
    Dim length As Long: length = UBound(ary)    ' 動的配列が初期化されていなければ、ここでエラーが発生する。
    IsInitialized = True
    Exit Function

' 配列が初期化されていない場合はここに飛ばされる。
NOT_INITIALIZED_ERROR:
    IsInitialized = False
End Function
