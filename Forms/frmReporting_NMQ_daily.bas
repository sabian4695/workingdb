Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Function PPAP_applyFilter(filterCond As String, filterName As String)

If Nz(Me.fltUnit) <> "" Then
    filterCond = Split(filterCond, " AND unitId")(0) 'remove unit filter if present
    filterCond = filterCond & " AND unitId = " & Me.fltUnit
    
    filterName = Split(filterName, " [")(0) 'remove unit from caption if present
    filterName = filterName & " [" & Me.fltUnit.column(1) & "]"
End If

Me.sfrmReporting_NMQ_daily_PPAP.Form.filter = filterCond
Me.sfrmReporting_NMQ_daily_PPAP.Form.FilterOn = True

Dim db As Database
Set db = CurrentDb()

Dim qdf As QueryDef
Set qdf = db.QueryDefs("sfrmReporting_NMQ_daily_Chart1_sub")

qdf.sql = Split(qdf.sql, "WHERE")(0) & " WHERE " & filterCond

db.QueryDefs.refresh

Set qdf = Nothing
Set db = Nothing

Me.Chart1.Requery

Me.lblTitle.Caption = filterName

End Function

Function trials_applyFilter(filterCond As String, filterName As String)

If Nz(Me.trials_fltUnit) <> "" Then
    filterCond = Split(filterCond, " AND unitId")(0) 'remove unit filter if present
    filterCond = filterCond & " AND unitId = " & Me.trials_fltUnit
    
    filterName = Split(filterName, " [")(0) 'remove unit from caption if present
    filterName = filterName & " [" & Me.trials_fltUnit.column(1) & "]"
End If

Me.sfrmReporting_NMQ_daily_trials.Form.filter = filterCond
Me.sfrmReporting_NMQ_daily_trials.Form.FilterOn = True

'Dim db As Database
'Set db = CurrentDb()
'
'Dim qdf As QueryDef
'Set qdf = db.QueryDefs("sfrmReporting_NMQ_daily_Chart1_sub")
'
'qdf.sql = Split(qdf.sql, "WHERE")(0) & " WHERE " & filterCond
'
'db.QueryDefs.refresh
'
'Set qdf = Nothing
'Set db = Nothing

'Me.Chart1.Requery

Me.trials_lblTitle.Caption = filterName

End Function

Private Sub fltUnit_AfterUpdate()
On Error GoTo Err_Handler

If Me.lblTitle.Caption = "No Report Selected" Or Me.lblTitle.Caption = "Please select a report first" Then 'no report selected
    Me.lblTitle.Caption = "Please select a report first"
    Me.ActiveControl = Null
    Exit Sub
End If

Call PPAP_applyFilter(Me.sfrmReporting_NMQ_daily_PPAP.Form.filter, Me.lblTitle.Caption)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub PPAP_subNotApproved_Click()
On Error GoTo Err_Handler

Dim filt As String
filt = "PPAPsubmit IS NOT NULL AND PPAPapproval IS NULL"

If Nz(Me.fltUnit) <> "" Then
    filt = filt & " AND unitId = " & Me.fltUnit
End If

Call PPAP_applyFilter(filt, "Submitted but Not Approved")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub PPAP_upcLate_Click()
On Error GoTo Err_Handler

Dim filt As String
filt = "PPAPsubmit IS NULL AND PPAPdue < Date() + 30"

If Nz(Me.fltUnit) <> "" Then
    filt = filt & " AND unitId = " & Me.fltUnit
End If

Call PPAP_applyFilter(filt, "Upcoming (<3 months) and Late")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub trials_fltUnit_AfterUpdate()
On Error GoTo Err_Handler

If Me.trials_lblTitle.Caption = "No Report Selected" Or Me.trials_lblTitle.Caption = "Please select a report first" Then 'no report selected
    Me.trials_lblTitle.Caption = "Please select a report first"
    Me.ActiveControl = Null
    Exit Sub
End If

Call trials_applyFilter(Me.sfrmReporting_NMQ_daily_trials.Form.filter, Me.trials_lblTitle.Caption)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub trials_upcoming_Click()
On Error GoTo Err_Handler

Dim filt As String
filt = "trialDate < Date() + 7 AND trialStatus = 'scheduled'"

If Nz(Me.trials_fltUnit) <> "" Then
    filt = filt & " AND unitId = " & Me.trials_fltUnit
End If

Call trials_applyFilter(filt, "Upcoming (7 days) Trials")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub trials_yesterday_Click()
On Error GoTo Err_Handler

Dim filt As String
filt = "trialDate = Date() -1"

If Nz(Me.trials_fltUnit) <> "" Then
    filt = filt & " AND unitId = " & Me.trials_fltUnit
End If

Call trials_applyFilter(filt, "Yesterday's Trials")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
