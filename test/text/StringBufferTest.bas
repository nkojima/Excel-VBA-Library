Sub StringBuffer_Test()
    ' 結合させる文字列
    Dim sampleTxt As String
    sampleTxt = "ABCDE"
    
    Dim i As Long
    Dim buff As StringBuffer
    Set buff = New StringBuffer
    
    ' 文字列を繰り返し連結する。
    For i = 1 To 5
        Call buff.Append(sampleTxt)
    Next i
    
    Debug.Print "正順：" & buff.ToString()
    
    ' バッファの文字数を取得する。
    Debug.Print "バッファの文字数：" & buff.Length
    
    ' バッファを逆順にする。
    Call buff.Reverse
    Debug.Print "逆順：" & buff.ToString()
    
    ' 部分文字列を取り出す
    Debug.Print "逆順の3文字目から2文字分切り出し：" & buff.Substring(3, 2)
End Sub
