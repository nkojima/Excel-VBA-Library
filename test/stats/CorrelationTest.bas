Option Explicit

' 正常系のテスト
Public Sub Test_CalcPearson()
    Dim x(3) As Double, y(3) As Double
    
    x(0) = 1
    x(1) = 2
    x(2) = 4
    x(3) = 8
    y(0) = 7
    y(1) = 5
    y(2) = 3
    y(3) = 1
    
    Debug.Print Correlation.CalcPearson(x, y)
End Sub

' 意図的にエラーを発生させるテスト
Public Sub Test_CalcPearson_Error()
    Dim x(3) As Double, y(4) As Double
    
    x(0) = 1
    x(1) = 2
    x(2) = 4
    x(3) = 8
    y(0) = 7
    y(1) = 5
    y(2) = 3
    y(3) = 1
    y(4) = 0
    
    Debug.Print Correlation.CalcPearson(x, y)
End Sub
