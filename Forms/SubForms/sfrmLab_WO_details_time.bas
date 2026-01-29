Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub addTime_Click()
On Error GoTo Err_Handler

Dim db As Database
Set db = CurrentDb()

db.Execute "INSERT INTO tbllab_wo_time(woid,workuser,workdate) VALUES (" & Form_frmLab_WO_details.recordId & ",'" & Environ("username") & "','" & Date & "');"
TempVars.Add "woTimeId", db.OpenRecordset("SELECT @@identity")(0).Value

Set db = Nothing

Call registerLabUpdates("tbllab_wo_time", TempVars!woTimeId, "WO Time", "", "Created", Me.woid, Me.name)
Me.Requery

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub deleteItem_Click()
On Error GoTo Err_Handler

Call registerLabUpdates("tbllab_wo_time", Me.recordId, "WO Time", "", "Deleted", Form_frmLab_WO_details.recordId, Me.name)
dbExecute "DELETE from tbllab_wo_time WHERE recordid = " & Me.recordId
Me.Requery

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub imgUser_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmUserProfile", , , "user = '" & Me.workUser & "'"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub workdate_AfterUpdate()
On Error GoTo Err_Handler

Call registerLabUpdates("tbllab_wo_time", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Form_frmLab_WO_details.recordId, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub workhours_AfterUpdate()
On Error GoTo Err_Handler

Call registerLabUpdates("tbllab_wo_time", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Form_frmLab_WO_details.recordId, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub worktype_AfterUpdate()
On Error GoTo Err_Handler

Call registerLabUpdates("tbllab_wo_time", Me.recordId, Me.ActiveControl.name, "", Me.ActiveControl.column(1), Form_frmLab_WO_details.recordId, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub workuser_AfterUpdate()
On Error GoTo Err_Handler

Call registerLabUpdates("tbllab_wo_time", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Form_frmLab_WO_details.recordId, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
