Option Explicit

'--------------------------------------------------------------------------------
' OSやハードウェアに関する処理をまとめた標準モジュール
'--------------------------------------------------------------------------------

'--------------------------------------------------------------------------------
' CPUの論理コア数（スレッド数）を取得する。
'
' return：CPUの論理コア数（スレッド数）。
'--------------------------------------------------------------------------------
Public Function GetCpuCoreCount()
    GetCpuCoreCount = Environ("NUMBER_OF_PROCESSORS")
End Function

'--------------------------------------------------------------------------------
' ホスト名を取得する。
'
' return：ホスト名。
'--------------------------------------------------------------------------------
Public Function GetHostName() As String
    GetHostName = GetHostName = Environ("COMPUTERNAME")
End Function

'--------------------------------------------------------------------------------
' 環境変数のPATH（コマンド検索パス）をセミコロンで区切って、フォルダパスの配列として返す。
'
' return：環境変数のPATHを構成しているフォルダパスの配列。
'--------------------------------------------------------------------------------
Public Function GetPathArray() As String()
    GetPathArray = Split(Environ("PATH"), ";")
End Function

'--------------------------------------------------------------------------------
' 環境変数のPATH（コマンド検索パス）そのものを文字列として返す。
'
' return：環境変数のPATH。
'--------------------------------------------------------------------------------
Public Function GetPathString() As String
    GetPathString = Environ("PATH")
End Function

'--------------------------------------------------------------------------------
' ログインユーザー名を取得する。
'
' return：ログインユーザー名。
'--------------------------------------------------------------------------------
Public Function GetUserName() As String
    GetUserName = GetUserName = Environ("USERNAME")
End Function