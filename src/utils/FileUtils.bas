Option Explicit

'--------------------------------------------------------------------------------
' ファイル、フォルダ関連のユーティリティー処理をまとめた標準モジュール
'--------------------------------------------------------------------------------

'--------------------------------------------------------------------------------
' 指定したファイルパスが存在するかを判定する。
'
' path：対象となるパス。
' return：ファイルが存在すればTrue、存在しなければFalseが返される。
'--------------------------------------------------------------------------------
Function ExistsFile(path As String) As Boolean
    ' パスの末尾が「\」であってもフォルダパスとして認識されるように、末尾の「\」を除去する。
    path = RemoveLastChar(path)

    If (Dir(path) <> "") Then
        ExistsFile = True
    Else
        ExistsFile = False
    End If
End Function

'--------------------------------------------------------------------------------
' 指定したパスが有効なフォルダパスであるかを判定する。
'
' path：対象となるパス。
' return：フォルダが存在すればTrue、存在しなければFalseが返される。
'--------------------------------------------------------------------------------
Function ExistsFolder(path As String) As Boolean
    ' パスの末尾が「\」であってもフォルダパスとして認識されるように、末尾の「\」を除去する。
    path = RemoveLastChar(path)

    If (Dir(path, vbDirectory) <> "" And Dir(path) = "") Then
        ExistsFolder = True
    Else
        ExistsFolder = False
    End If
End Function

'--------------------------------------------------------------------------------
' ファイルパスから拡張子名を取得する。
'
' path：対象となるファイルパス。
' return：拡張子名。ファイルが存在しない時、拡張子が存在しない時は空文字が返される。
'--------------------------------------------------------------------------------
Public Function GetExtensionName(path As String) As String
    ' パスの末尾が「\」であってもフォルダパスとして認識されるように、末尾の「\」を除去する。
    path = RemoveLastChar(path)

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

'--------------------------------------------------------------------------------
' ファイルパスからファイル名を取得する。
'
' path：対象となるファイルパス。
' return：ファイル名。ファイルが存在しない時は空文字が返される。
'--------------------------------------------------------------------------------
Public Function GetFileName(path As String) As String
    GetFileName = Dir(path)
End Function

'--------------------------------------------------------------------------------
' パスの末尾が「\」であれば除去する。
'
' path：対象となるファイルパス。
' return：パスの末尾の「\」を除去した文字列。
'--------------------------------------------------------------------------------
Private Function RemoveLastChar(path As String) As String
    Dim lastChar As String
    lastChar = Right(path, 1)
    If (lastChar = "\") Then
        path = Left(path, Len(path) - 1)
    End If
    
    RemoveLastChar = path
End Function