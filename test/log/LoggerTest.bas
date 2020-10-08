Option Explicit

'--------------------------------------------------------------------------------
' Loggerのテストをまとめた標準モジュール
'--------------------------------------------------------------------------------

' ログ出力先となるシート名@ThisWorkbook
Private Const LOG_SHEET_NAME As String = "ログ"

'--------------------------------------------------------------------------------
' ログのシートを初期化する。
'--------------------------------------------------------------------------------
Sub Test_Initialize()
    ' ログのシートを消す。
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(LOG_SHEET_NAME).Delete
    Application.DisplayAlerts = True

    Call Logger.Initialize
    
    ' 初期化が完了していれば、ログのシートが存在するはず。
    Dim existsSheet As Boolean
    Dim sheet As Variant
    
    For Each sheet In ThisWorkbook.Worksheets
        If (sheet.Name = LOG_SHEET_NAME) Then
            existsSheet = True
            Exit For
        End If
    Next
    
    ' シートが存在しない時（＝初期化が完了していない時）はテストを停止させる。
    Debug.Assert existsSheet
End Sub