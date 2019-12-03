Option Explicit

'--------------------------------------------------------------------------------
' ハッシュ値を取得する処理の標準モジュール
'--------------------------------------------------------------------------------

'--------------------------------------------------------------------------------
' 引数で指定したファイルのハッシュ値（ダイジェスト値）を取得する。
'
' filePath：ファイルパス。
' hashAlgorithm：ハッシュアルゴリズムの種類。MD5,SHA1,SHA256などを指定可能。
' return：ファイルのハッシュ値（ダイジェスト値）
'--------------------------------------------------------------------------------
Public Function HashFile(filePath As String, hashAlgorithm As String) As String

    Dim wsh As Object, wExec As Object, command As String, output As String

    Set wsh = CreateObject("WScript.Shell")
    command = "certutil -hashfile " & filePath & " " & hashAlgorithm & " | findstr /V CertUtil | findstr /V " & hashAlgorithm
    Set wExec = wsh.Exec("%ComSpec% /c " & command)

    Do While wExec.Status = 0
        DoEvents
    Loop

    output = wExec.stdOut.ReadAll

    Set wExec = Nothing
    Set wsh = Nothing

    HashFile = output

End Function