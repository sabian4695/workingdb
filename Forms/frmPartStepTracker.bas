Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Function applyTheFilters()
On Error GoTo Err_Handler
Dim filt
filt = ""

If Me.fltPartNumber <> "" Then filt = "partNumber = '" & Me.fltPartNumber & "'"

If Me.fltUser <> "" Then
    If filt <> "" Then filt = filt & " AND "
    filt = filt & "person = '" & Me.fltUser & "'"
End If

If Me.fltModel <> "" Then
    If filt <> "" Then filt = filt & " AND "
    filt = filt & "modelCode = '" & Me.fltModel & "'"
End If

filtNow:
If filt <> "" Then filt = filt & " AND "
Me.filter = filt & "due < #" & Date + 14 & "#"
Me.FilterOn = True
Me.OrderBy = "due"
Me.OrderByOn = True

Exit Function
Err_Handler:
    Call handleError(Me.name, "applyTheFilters", Err.DESCRIPTION, Err.number)
End Function

Private Sub export_Click()
On Error GoTo Err_Handler

Dim FileName As String, sqlString As String, filt As String
FileName = "H:\Step_Tracker_" & nowString & ".xlsx"
filt = ""
If Me.Form.filter <> "" And Me.Form.FilterOn Then filt = " WHERE " & Me.Form.filter
sqlString = "SELECT * FROM qryStepApprovalTracker" & filt

Call exportSQL(sqlString, FileName)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub fltModel_AfterUpdate()
On Error GoTo Err_Handler
applyTheFilters
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub fltPartNumber_AfterUpdate()
On Error GoTo Err_Handler
applyTheFilters
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub fltUser_AfterUpdate()
On Error GoTo Err_Handler
applyTheFilters
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Select Case userData("Level")
    Case "Supervisor", "Manager"
        Me.RecordSource = "SELECT * FROM sqryStepApprovalTracker_Approvals_SupervisorsUp"
    Case Else
        Me.RecordSource = "SELECT * FROM sqryStepApprovalTracker_Approvals UNION SELECT * FROM sqryStepApprovalTracker_Steps;"
End Select


Me.OrderBy = "due"
Me.OrderByOn = True

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub partTrackingHelp_Click()
On Error GoTo Err_Handler

Call openPath(mainFolder(Me.ActiveControl.name))

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub openDetails_Click()
On Error GoTo Err_Handler

openPartProject (Me.partNumber)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Public Sub refresh_Click()
On Error GoTo Err_Handler
Me.Requery
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
