Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setSplashLoading("Building calendar view...")

TempVars.Add "selMonth", Month(Date)
TempVars.Add "selYear", Year(Date)

initializeDateButtons

Exit Sub
Err_Handler:: Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number): End Sub

Function initializeDateButtons()
On Error GoTo Err_Handler

Me.prevMonth.SetFocus
Me.today.Visible = True

Dim MonthDayOne As Date, MonthDayLast As Date, MonthLength As Integer, DayOfWeek As Integer
Dim i As Integer, Y As Integer, X As Integer, btn As CommandButton
Dim OutOfCurrentMonth As Boolean, OutOfDate As Date, intMaxWeek As Integer, curDate As Date
Dim countT As Long

Me.txtDay = Date
Me.txtLongDay = Format(Date, "dddd, mmmm dd, yyyy")

Dim db As Database
Set db = CurrentDb()
Dim rsHolidays As Recordset
Set rsHolidays = db.OpenRecordset("tblHolidays")

Me.lblMonth.Caption = MonthName(TempVars!selMonth)
Me.lblYear.Caption = TempVars!selYear

MonthDayOne = DateSerial(TempVars!selYear, TempVars!selMonth, 1)
DayOfWeek = DatePart("w", MonthDayOne, vbUseSystemDayOfWeek)
MonthLength = DatePart("d", DateAdd("d", -1, DateAdd("m", 1, MonthDayOne)))

i = 2 - DayOfWeek
For Y = 0 To 5 'for each week
    For X = 0 To 6 'each each weekday
        Set btn = Me.Controls("d" & Y & X)
        btn.Visible = True
        Me.Controls("bd" & Y & X).Visible = True
        curDate = DateSerial(TempVars!selYear, TempVars!selMonth, i)
        
        If X = 6 Then 'for the last day of each week, find the week number
            Me.Controls("lblW" & Y) = DatePart("ww", curDate)
            Me.Controls("lblW" & Y).Visible = True
        End If
        
        rsHolidays.FindFirst "holidayDate = #" & curDate & "#"
        
        countT = 0
        Me.Controls("bd" & Y & X).Visible = False
        
        If curDate < Date Then Me.Controls("bd" & Y & X).BackColor = rgb(230, 0, 0)
        Me.Controls("bd" & Y & X).Caption = countT
        
        'If i falls within legal days for this month, show  this button.
        If (i > 0) And (i <= MonthLength) Then
            OutOfCurrentMonth = False
            btn.Caption = i
            btn.tag = "btnContrastBorder.L2"
            
            'get row value for final day of selected month
            If i = MonthLength Then intMaxWeek = Y
            If X = 0 Or X = 6 Then btn.tag = "btnDisContrastBorder.L2"
            If Not rsHolidays.noMatch Then btn.tag = "btnXcontrastBorder.L0" ' IF HOLIDAY -> show it
        Else
            btn.tag = "btnDis.L1"
            OutOfCurrentMonth = True
            OutOfDate = DateAdd("d", i - 1, DateSerial(TempVars!selYear, TempVars!selMonth, 1))
            btn.Caption = Day(OutOfDate)
            If i > MonthLength And Y > intMaxWeek Then 'wk6
                btn.Visible = False
                Me.lblW5.Visible = False
                Me.Controls("bd" & Y & X).Visible = False
            End If
            If Not rsHolidays.noMatch Then btn.tag = "btnXdis.L0" ' IF HOLIDAY -> show it
        End If
        
        btn.FontWeight = 400
        btn.BorderStyle = 1
        If i = Day(Date) And Me.lblYear.Caption = TempVars!selYear And Month(Date) = TempVars!selMonth Then 'current date
            Me.today.Visible = False
            btn.tag = "btnContrastBorder.L4"
            btn.BorderStyle = 3
            btn.FontWeight = 700
        End If
        ' Advance to next day.
        i = i + 1
    Next
Next

rsHolidays.CLOSE
Set rsHolidays = Nothing
Set db = Nothing

Call setTheme(Me)

Exit Function
Err_Handler:
    Call handleError(Me.name, "drawDateButtons", Err.DESCRIPTION, Err.number)
End Function

Function drawDateButtons()
On Error GoTo Err_Handler

Me.prevMonth.SetFocus
Me.today.Visible = True

Dim MonthDayOne As Date, MonthDayLast As Date, MonthLength As Integer, DayOfWeek As Integer
Dim i As Integer, Y As Integer, X As Integer, btn As CommandButton
Dim OutOfCurrentMonth As Boolean, OutOfDate As Date, intMaxWeek As Integer, curDate As Date
Dim countT As Long

Me.txtDay = Date
Me.txtLongDay = Format(Date, "dddd, mmmm dd, yyyy")

Dim db As Database
Set db = CurrentDb()
Dim rsHolidays As Recordset, rsTasks As Recordset
Set rsHolidays = db.OpenRecordset("tblHolidays")

Dim sqlSel As String, sqlWhere As String, sqlStatement As String

sqlSel = "COUNT(ID) AS cTasks, due"
sqlWhere = "person = '" & Environ("username") & "' GROUP BY due"

Select Case userData("Level")
    Case "Supervisor", "Manager"
        sqlStatement = "SELECT " & sqlSel & " FROM sqryCalendarItems_Approvals_SupervisorsUp WHERE " & sqlWhere & _
            " UNION ALL SELECT " & sqlSel & " FROM sqryCalendarItems_WOS WHERE " & sqlWhere & _
            " UNION ALL SELECT " & sqlSel & " FROM sqryCalendarItems_Issues WHERE " & sqlWhere & ";"
    Case Else
        sqlStatement = "SELECT " & sqlSel & " FROM sqryCalendarItems_Approvals WHERE " & sqlWhere & _
            " UNION ALL SELECT " & sqlSel & " FROM sqryCalendarItems_Steps WHERE " & sqlWhere & _
            " UNION ALL SELECT " & sqlSel & " FROM sqryCalendarItems_WOS WHERE " & sqlWhere & _
            " UNION ALL SELECT " & sqlSel & " FROM sqryCalendarItems_Issues WHERE " & sqlWhere & ";"
End Select
            
Set rsTasks = db.OpenRecordset(sqlStatement)

Me.lblMonth.Caption = MonthName(TempVars!selMonth)
Me.lblYear.Caption = TempVars!selYear

MonthDayOne = DateSerial(TempVars!selYear, TempVars!selMonth, 1)
DayOfWeek = DatePart("w", MonthDayOne, vbUseSystemDayOfWeek)
MonthLength = DatePart("d", DateAdd("d", -1, DateAdd("m", 1, MonthDayOne)))

i = 2 - DayOfWeek
For Y = 0 To 5 'for each week
    For X = 0 To 6 'each each weekday
        Set btn = Me.Controls("d" & Y & X)
        btn.Visible = True
        Me.Controls("bd" & Y & X).Visible = True
        curDate = DateSerial(TempVars!selYear, TempVars!selMonth, i)
        
        If X = 6 Then 'for the last day of each week, find the week number
            Me.Controls("lblW" & Y) = DatePart("ww", curDate)
            Me.Controls("lblW" & Y).Visible = True
        End If
        
        rsHolidays.FindFirst "holidayDate = #" & curDate & "#"
        rsTasks.FindFirst "due = #" & curDate & "#"
        
        If Not rsTasks.noMatch Then
            countT = rsTasks!cTasks
        Else
            countT = 0
        End If
        
        Select Case countT
            Case Is > 9
                Me.Controls("bd" & Y & X).Caption = "9+"
                Me.Controls("bd" & Y & X).BackColor = rgb(230, 0, 0)
            Case 0
                Me.Controls("bd" & Y & X).Visible = False
            Case Else
                Me.Controls("bd" & Y & X).Caption = countT
                Me.Controls("bd" & Y & X).BackColor = rgb(60, 170, 60)
        End Select
        
        If curDate < Date Then Me.Controls("bd" & Y & X).BackColor = rgb(230, 0, 0)
        Me.Controls("bd" & Y & X).Caption = countT
        
        'If i falls within legal days for this month, show  this button.
        If (i > 0) And (i <= MonthLength) Then
            OutOfCurrentMonth = False
            btn.Caption = i
            btn.tag = "btnContrastBorder.L2"
            
            'get row value for final day of selected month
            If i = MonthLength Then intMaxWeek = Y
            If X = 0 Or X = 6 Then btn.tag = "btnDisContrastBorder.L2"
            If Not rsHolidays.noMatch Then btn.tag = "btnXcontrastBorder.L0" ' IF HOLIDAY -> show it
        Else
            btn.tag = "btnDis.L1"
            OutOfCurrentMonth = True
            OutOfDate = DateAdd("d", i - 1, DateSerial(TempVars!selYear, TempVars!selMonth, 1))
            btn.Caption = Day(OutOfDate)
            If i > MonthLength And Y > intMaxWeek Then 'wk6
                btn.Visible = False
                Me.lblW5.Visible = False
                Me.Controls("bd" & Y & X).Visible = False
            End If
            If Not rsHolidays.noMatch Then btn.tag = "btnXdis.L0" ' IF HOLIDAY -> show it
        End If
        
        btn.FontWeight = 400
        btn.BorderStyle = 1
        If i = Day(Date) And Me.lblYear.Caption = TempVars!selYear And Month(Date) = TempVars!selMonth Then 'current date
            Me.today.Visible = False
            btn.tag = "btnContrastBorder.L4"
            btn.BorderStyle = 3
            btn.FontWeight = 700
        End If
        ' Advance to next day.
        i = i + 1
    Next
Next

rsTasks.CLOSE
rsHolidays.CLOSE

Set rsTasks = Nothing
Set rsHolidays = Nothing
Set db = Nothing

Call setTheme(Me)

Exit Function
Err_Handler:
    Call handleError(Me.name, "drawDateButtons", Err.DESCRIPTION, Err.number)
End Function

Function openCalendarItem()
On Error GoTo Err_Handler

Dim yx As String, datePicked As Date
yx = Right(Me.ActiveControl.name, 2)

If Me.Controls("bd" & yx).Visible = False Then Exit Function 'if the bubble isn't visible, just ignore the click

'find the default day value
datePicked = DateSerial(Me.lblYear.Caption, TempVars!selMonth, Me.Controls("d" & yx).Caption)

'check if the clicked date is NOT in the current month
If Me.Controls("d" & yx).tag = "btnXdis.L0" Or Me.Controls("d" & yx).tag = "btnDis.L1" Then
    If CLng(Me.Controls("d" & yx).Caption) > 15 Then 'if the datevalue is in the last half of the month, it should be in the month prior
        datePicked = DateSerial(Me.lblYear.Caption, TempVars!selMonth - 1, Me.Controls("d" & yx).Caption)
    Else 'otherwise it in in the future month
        datePicked = DateSerial(Me.lblYear.Caption, TempVars!selMonth + 1, Me.Controls("d" & yx).Caption)
    End If
End If

Form_DASHBOARD.sfrmCalendarItems.Form.RecordSource = "qryCalendarItems"
Form_DASHBOARD.sfrmCalendarItems.Form.filter = "person = '" & Environ("username") & "' AND due = #" & datePicked & "#"
Form_DASHBOARD.sfrmCalendarItems.Form.FilterOn = True
Form_DASHBOARD.sfrmCalendarItems.Visible = True
Form_DASHBOARD.sfrmCalendarItems.SetFocus
Form_DASHBOARD.sfrmCalendarView.Visible = False

Exit Function
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Function

Private Sub Form_Timer()
On Error Resume Next
drawDateButtons

End Sub

Private Sub loadCalendar_Click()
On Error GoTo Err_Handler

Call logClick(Me.ActiveControl.name, Me.name)

Me.TimerInterval = 3600000 'every hour
drawDateButtons
Me.loadCalendar.Visible = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub nextMonth_Click()
On Error GoTo Err_Handler

If TempVars!selMonth = 12 Then
    TempVars.Add "selYear", TempVars!selYear + 1
    TempVars.Add "selMonth", 1
Else
    TempVars.Add "selMonth", TempVars!selMonth + 1
End If

drawDateButtons

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub prevMonth_Click()
On Error GoTo Err_Handler

If TempVars!selMonth = 1 Then
    TempVars.Add "selYear", TempVars!selYear - 1
    TempVars.Add "selMonth", 12
Else
    TempVars.Add "selMonth", TempVars!selMonth - 1
End If

drawDateButtons

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub today_Click()
On Error GoTo Err_Handler

TempVars.Add "selMonth", Month(Date)
TempVars.Add "selYear", Year(Date)

drawDateButtons

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
