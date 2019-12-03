Option Explicit

'--------------------------------------------------------------------------------
' 日付関連のユーティリティー処理をまとめた標準モジュール
'--------------------------------------------------------------------------------

'--------------------------------------------------------------------------------
' 引数の日付の会計年度を返す。
'
' d：日付。
' startMonth：新年度の開始月。新年度が4月から始まる場合は4を指定する。
' return：YYYY形式の西暦で表された年度。
'--------------------------------------------------------------------------------
Public Function CalcFiscalYear(d As Date, startMonth) As Integer
    CalcFiscalYear = Year(DateAdd("m", 1 - startMonth, d))
End Function

'--------------------------------------------------------------------------------
' 引数の日付の月における「月末日」を返す。
'
' d：日付。
' return：引数で指定した日付における月末日。
'--------------------------------------------------------------------------------
Public Function CalcLastDayOfMonth(d As Date)
  CalcLastDayOfMonth = DateSerial(Year(d), Month(d) + 1, 0)
End Function

'--------------------------------------------------------------------------------
' 引数の日付が属する四半期（1Q,2Q,3Q,4Q）を返す。
'
' d：日付。
' startMonth：新年度の開始月。新年度が4月から始まる場合は4を指定する。
' return：四半期を表す文字列（1Q,2Q,3Q,4Q）。
'--------------------------------------------------------------------------------
Public Function CalcQuarter(d As Date, startMonth As Integer) As String
    Dim monthDiff As Integer, currentMonth As Integer
    currentMonth = Month(d)

    If (currentMonth < startMonth) Then
        monthDiff = 12 - Abs(currentMonth - startMonth)
    Else
        monthDiff = Abs(currentMonth - startMonth)
    End If
​
    Dim quarter As String
    quarter = CStr(Application.RoundDown(monthDiff / 3, 0) + 1)
    CalcQuarter = quarter & "Q"
End Function

'--------------------------------------------------------------------------------
' 指定した年月において、指定した曜日の日付を返す。
' FindWantDayOfWeek(2019, 11, vbMonday)とした場合、2019/11/4, 2019/11/11, 2019/11/18, 2019/11/25が返される。
'
' year：西暦年。
' month ：月。
' wantDayOfWeek：曜日。VbDayOfWeek列挙型で表される。
' return：指定した年月における、指定した曜日のリスト。
'--------------------------------------------------------------------------------
Public Function FindWantDayOfWeek(year As Integer, month As Integer, wantDayOfWeek As VbDayOfWeek) As Collection

    Dim firstDay As Date, lastDay As Date
    firstDay = DateSerial(year, month, 1)       ' 指定年月の初日
    lastDay = CalcLastDayOfMonth(firstDay)      ' 指定年月の末日

    Dim i As Integer
    Dim dateBuff As Date, dateList As Collection
    Set dateList = New Collection
    dateBuff = firstDay

    For i = 1 To Day(lastDay)
        If (Weekday(dateBuff) = wantDayOfWeek) Then
            dateList.Add dateBuff
        End If
        dateBuff = DateAdd("d", 1, dateBuff)
    Next i

    Set FindWantDayOfWeek = dateList
End Function

'--------------------------------------------------------------------------------
' 引数の年がうるう年であるかを判定する。
'
' year：判定対象の西暦年。
' return：うるう年であればtrue、そうでなければfalseを返す。
'--------------------------------------------------------------------------------
Public Function IsLeapYear(year As Integer) As Boolean
    IsLeapYear = ((year Mod 4 = 0) And (year Mod 100 <> 0)) Or (year Mod 400 = 0)
End Function