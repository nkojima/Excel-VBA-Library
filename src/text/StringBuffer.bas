Option Explicit

'--------------------------------------------------------------------------------
' 文字列バッファのクラス。
'--------------------------------------------------------------------------------

' 動的配列のバッファ
Private buffer() As String

' バッファの長さ
Private buffLength As Long

'--------------------------------------------------------------------------------
' コンストラクタ
'--------------------------------------------------------------------------------
Public Sub Class_Initialize()
    ReDim buffer(0)
End Sub

'--------------------------------------------------------------------------------
' 引数の文字列をバッファに追加する。
'
' str: 追加する文字列。
'--------------------------------------------------------------------------------
Sub Append(str As String)
    ' バッファの長さを調整する。
    Dim strLength As Long
    strLength = Len(str)
    
    If (strLength > 0) Then
        ReDim Preserve buffer(buffLength + strLength - 1)
    End If
    
    ' バッファに1文字ずつ追加する。
    Dim i As Long
    For i = 0 To (strLength - 1)
        buffer(buffLength + i) = Mid(str, (i + 1), 1)
    Next i
    
    buffLength = buffLength + strLength
End Sub

'--------------------------------------------------------------------------------
' バッファの長さ（文字数）を返す。
'--------------------------------------------------------------------------------
Function Length() As Long
    Length = buffLength
End Function

'--------------------------------------------------------------------------------
' バッファの文字を逆順にする。
'--------------------------------------------------------------------------------
Sub Reverse()
    ' 逆順にしたデータを格納する一時的なバッファを作る。
    Dim tempBuffer() As String
    ReDim tempBuffer(buffLength)
    
    ' 一時的なバッファに逆順にした文字を入れていく。
    Dim i As Long
    For i = LBound(buffer) To UBound(buffer)
        tempBuffer(i) = buffer(UBound(buffer) - i)
    Next i
    
    ' 一時的なバッファ（逆順）とバッファ（正順）を入れ替える。
    buffer = tempBuffer
End Sub

'--------------------------------------------------------------------------------
' 開始位置と終了位置を指定して、部分文字列を返す。
'
' startIdx: 開始位置（1～n）。
' endIdx: 終了位置（1～n）。
'--------------------------------------------------------------------------------
Function Substring(startIdx As Long, endIdx As Long) As String
    ' バッファを文字列にした後、必要な部分をMid関数で切り取る。
    Dim buffStr As String
    buffStr = ToString()
    
    Substring = Mid(buffStr, startIdx, endIdx)
End Function

'--------------------------------------------------------------------------------
' バッファを文字列として返す。
'--------------------------------------------------------------------------------
Function ToString() As String
    ToString = Join(buffer, "")
End Function