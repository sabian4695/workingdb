Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub closeTask_Click()
On Error GoTo Err_Handler

Me.closeDate = Date
Me.Requery

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub deleteTask_Click()
On Error GoTo Err_Handler

If IsNull(Me.recordId) Then Exit Sub

If MsgBox("Are you sure you want to delete this?", vbYesNo, "Please confirm") = vbNo Then Exit Sub
Call registerWdbUpdates("tblTasks", Me.recordId, Nz(Me.workType.column(1), ""), "", "Deleted", Nz(Me.workItem, ""))
dbExecute "DELETE from tblTasks WHERE recordId = " & Me.recordId
Me.Requery

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setSplashLoading("Building task tracker...")

Call setTheme(Me)

Me.filter = "workUser = '" & Environ("username") & "' AND closeDate is null"
Me.FilterOn = True

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblDate_Click()
    Me.workDate.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
End Sub

Private Sub lblHours_Click()
    Me.workHours.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
End Sub

Private Sub lblNotes_Click()
    Me.workNotes.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
End Sub

Private Sub lblPN_Click()
    Me.workItem.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
End Sub

Private Sub lblType_Click()
    Me.workType.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
End Sub

Private Sub newTask_Click()
On Error GoTo Err_Handler

Dim db As Database
Set db = CurrentDb()

db.Execute "INSERT INTO tblTasks(workUser,workDate) VALUES ('" & Environ("username") & "','" & Date & "');"
TempVars.Add "newTaskId", db.OpenRecordset("SELECT @@identity")(0).Value
Call registerWdbUpdates("tblTasks", TempVars!newTaskId, "Task", "", "Created")
Me.Requery

Set db = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub showClosedToggle_Click()
On Error GoTo Err_Handler

Dim mainFilt As String
If Me.showClosedToggle.Value = True Then
        mainFilt = "not null"
    Else
        mainFilt = "null"
End If

Me.filter = "workUser = '" & Environ("username") & "' AND closeDate is " & mainFilt
Me.FilterOn = True

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Function trackUpdate()
On Error GoTo Err_Handler

Dim oldVal, newVal
oldVal = Me.ActiveControl.OldValue
newVal = Me.ActiveControl

If Me.ActiveControl.name = "workType" Then
    oldVal = ""
    newVal = Me.workType.column(1)
End If

Call registerWdbUpdates("tblTasks", Me.recordId, Me.ActiveControl.name, oldVal, newVal, Nz(Me.workItem, ""))

Exit Function
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Function

Private Sub tasksHistory_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmHistory"
Form_frmHistory.RecordSource = "tblWdbUpdateTracking"
Form_frmHistory.dataTag0.ControlSource = "dataTag0"
Form_frmHistory.previousData.ControlSource = "previousData"
Form_frmHistory.newData.ControlSource = "newData"
Form_frmHistory.filter = "tableName = 'tblTasks' AND updatedBy = '" & Environ("username") & "'"
Form_frmHistory.FilterOn = True
Form_frmHistory.OrderBy = "updatedDate Desc"
Form_frmHistory.OrderByOn = True

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
