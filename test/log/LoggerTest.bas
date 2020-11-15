Option Explicit

'--------------------------------------------------------------------------------
' Loggerのテストをまとめた標準モジュール
'--------------------------------------------------------------------------------

'--------------------------------------------------------------------------------
' ログのシートを初期化する。
'--------------------------------------------------------------------------------
Sub Test_Initialize()
    ' ログのシートを消す。
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(Logger.LOG_SHEET_NAME).Delete
    Application.DisplayAlerts = True

    Call Logger.Initialize
    
    ' 初期化が完了していれば、ログのシートが存在するはず。
    Dim existsSheet As Boolean
    Dim sheet As Variant
    
    For Each sheet In ThisWorkbook.Worksheets
        If (sheet.Name = Logger.LOG_SHEET_NAME) Then
            existsSheet = True
            Exit For
        End If
    Next
    
    ' シートが存在しない時（＝初期化が完了していない時）はテストを停止させる。
    Debug.Assert existsSheet
End Sub

'--------------------------------------------------------------------------------
' ログのシートをクリアして､見出しを再設定する｡
'--------------------------------------------------------------------------------
Public Sub Test_Clear()
    Call Logger.Clear
    
    ' A1セルが「日時」でない時（＝見出しがセットされていない時）はテストを停止させる。
    Debug.Assert (ThisWorkbook.Worksheets(Logger.LOG_SHEET_NAME).Range("A1").Value = "日時")
End Sub

'--------------------------------------------------------------------------------
' ログ出力先のシート名を返す。
'--------------------------------------------------------------------------------
Public Sub Test_GetLogSheetName()
    Debug.Print "ログ出力先のシート名：" & Logger.LOG_SHEET_NAME
End Sub