Option Explicit

'------------------------------------------------------------------------------
' ログ出力処理をまとめた標準モジュール
'------------------------------------------------------------------------------

' ログ出力先となるシート名@ThisWorkbook
Private Const LOG_SHEET_NAME As String = "ログ"

'--------------------------------------------------------------------------------
' ログのシートを初期化する。
'--------------------------------------------------------------------------------
Public Sub Initialize()
    ' ログのシートが存在しなければ作成する。
    Dim existsSheet As Boolean
    Dim sheet As Variant
    
    For Each sheet In ThisWorkbook.Worksheets
        If (sheet.Name = LOG_SHEET_NAME) Then
            existsSheet = True
            Exit For
        End If
    Next
    
    If Not (existsSheet) Then
        ThisWorkbook.Worksheets.Add
        ActiveSheet.Name = LOG_SHEET_NAME
    End If
    
    ' ログのシートをクリアして、見出しを再設定する。
    With ThisWorkbook.Worksheets(LOG_SHEET_NAME)
        .Cells.ClearContents
    End With
    Call SetColumnName
End Sub

'--------------------------------------------------------------------------------
' ログのシートの見出しをセットする。
'--------------------------------------------------------------------------------
Public Sub SetColumnName()
    With ThisWorkbook.Worksheets(LOG_SHEET_NAME)
        .Range("A1").Value = "日時"
        .Range("B1").Value = "ログレベル"
        .Range("C1").Value = "内容"
    End With
End Sub

'--------------------------------------------------------------------------------
' 「情報」レベルのログを出力する。
'--------------------------------------------------------------------------------
Public Sub Info(message As String)
    Call Logging(message, "INFO")
End Sub

'--------------------------------------------------------------------------------
' 「警告」レベルのログを出力する。
'--------------------------------------------------------------------------------
Public Sub Warn(message As String)
    Call Logging(message, "WARNING")
End Sub

'--------------------------------------------------------------------------------
' 「エラー」レベルのログを出力する。
'--------------------------------------------------------------------------------
Public Sub Error(message As String)
    Call Logging(message, "ERROR")
End Sub

'------------------------------------------------------------------------------
' ログの出力
'
' message: ログの内容
' logLevel: ログのレベル（INFO/WARNING/ERROR）
'------------------------------------------------------------------------------
Private Sub Logging(message As String, logLevel As String)
    ' 「A列の最終行」の次の行にログを出力する。
    With ThisWorkbook.Worksheets(LOG_SHEET_NAME)
        Dim lastRow As Integer
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).row
        
        .Cells(lastRow + 1, 1).Value = Now
        .Cells(lastRow + 1, 2).Value = logLevel
        .Cells(lastRow + 1, 3).Value = message
    End With
End Sub
