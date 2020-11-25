Option Explicit

'--------------------------------------------------------------------------------
' ArrayUtilsのテストをまとめた標準モジュール
'--------------------------------------------------------------------------------

Sub Test_CalcArrayLength()
    ' 要素数が11の配列（※添え字が0～10）
    Dim ary1(10) As Integer
    Debug.Print "ary1の長さ：" & ArrayUtils.GetLength(ary1)

    ' 要素数が10の配列（※添え字が1～10）
    Dim ary2(1 To 10) As String
    Debug.Print "ary2の長さ：" & ArrayUtils.GetLength(ary2)

    ' 初期化済みの動的配列（※添え字が0～5）
    Dim ary3() As String
    ReDim ary3(5)
    Debug.Print "ary3の長さ：" & ArrayUtils.GetLength(ary3)

    ' 引数が初期化されていない動的配列→-1が返される
    Dim ary4() As String
    Debug.Print "ary4の長さ：" & ArrayUtils.GetLength(ary4)

    ' 引数が配列以外→-100が返される。
    Debug.Print "配列以外：" & ArrayUtils.GetLength("abc")
End Sub
