Option Explicit

'--------------------------------------------------------------------------------
' ファイル、フォルダ関連のユーティリティー処理をまとめた標準モジュール
'--------------------------------------------------------------------------------

'--------------------------------------------------------------------------------
' 引数のファイルパスが存在するかを確認する。
'
' path：対象となるファイルパス。
' return：ファイルが存在すればTrue、存在しなければFalseを返す。
'--------------------------------------------------------------------------------
Public Function ExistsFile(path As String) As Boolean
    If (Dir(path) <> "") Then
        ExistsFile = True
    Else
        ExistsFile = False
    End If
End Function

'--------------------------------------------------------------------------------
' 引数のフォルダパスが存在するかを確認する。
'
' path：対象となるフォルダパス。
' return：フォルダが存在すればTrue、存在しなければFalseを返す。
'--------------------------------------------------------------------------------
Public Function ExistsFolder(path As String) As Boolean
    If (Dir(path, vbDirectory) <> "" And Dir(path) = "") Then
        ExistsFolder = True
    Else
        ExistsFolder = False
    End If
End Function