Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cbClose_Click()
On Error GoTo Err_Handler

If Me.cbClose = True Then
    Me.Close_Date = Date
    Me.status = "DONE"
Else
    Me.Close_Date = Null
    Me.status = ""
End If
Me.Requery
Call forms("frmDRSdashboard").setProgressBar
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub deleteTime_Click()
On Error GoTo Err_Handler

Dim db As Database
Set db = CurrentDb()
db.Execute "DELETE from tblTaskTracker WHERE Task_ID = " & Me.TASK_ID
Set db = Nothing
Me.Requery

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub newTask_Click()
On Error GoTo Err_Handler

Dim db As Database
Set db = CurrentDb()
db.Execute "INSERT INTO tblTaskTracker(Task,Control_Number) VALUES (''," & forms.frmDRSdashboard.Control_Number & ");"
Set db = Nothing
Me.Requery
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub showClosedToggle_Click()
On Error GoTo Err_Handler

Dim mainFilt As String
If Me.showClosedToggle.Value = True Then
        mainFilt = "[Control_Number] = " & forms.frmDRSdashboard.Control_Number & " AND [Status] = 'Done'"
    Else
        mainFilt = "[Close_Date] is null AND [Control_Number] = " & forms.frmDRSdashboard.Control_Number
End If

Me.filter = mainFilt
Me.FilterOn = True

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
