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

'--------------------------------------------------------------------------------
' 指定したパスが有効なファイルパスであるかを判定する。
'
' path：対象となるパス。
' return：パスがファイルパスであればTrue、そうでなければFalseが返される。
'--------------------------------------------------------------------------------
Function IsFile(path As String) As Boolean
    If (Dir(path) <> "") Then
        IsFile = True
    Else
        IsFile = False
    End If
End Function

'--------------------------------------------------------------------------------
' 指定したパスが有効なフォルダパスであるかを判定する。
'
' path：対象となるパス。
' return：パスがフォルダパスであればTrue、そうでなければFalseが返される。
'--------------------------------------------------------------------------------
Function IsFolder(path As String) As Boolean
    If (Exists(path)) Then
        ' GetAttr関数でファイル属性を調べる。
        If (GetAttr(path) = vbDirectory) Then
            IsFolder = True
        Else
            IsFolder = False
        End If
    Else
        IsFolder = False
    End If
End Function

'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------
' ファイルパスからファイル名を取得する。
'
' path：対象となるファイルパス。
' return：ファイル名。ファイルが存在しない時は空文字が返される。
'--------------------------------------------------------------------------------
Public Function GetBaseName(path As String) As String
    GetBaseName = Dir(path)
End Function

'--------------------------------------------------------------------------------
' ファイルパスから拡張子名を取得する。
'
' path：対象となるファイルパス。
' return：拡張子名。ファイルが存在しない時、拡張子が存在しない時は空文字が返される。
'--------------------------------------------------------------------------------
Public Function GetExtensionName(path As String) As String
    Dim fileName As String
    fileName = Dir(path)
    
    If (fileName <> "") Then
        Dim periodIdx As Long
        periodIdx = InStrRev(fileName, ".")
        GetExtensionName = Mid(fileName, periodIdx + 1)
    Else
        ' ファイルが存在しない時は空文字を返す。
        GetExtensionName = ""
    End If
End Function