Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub addResource_Click()
On Error GoTo Err_Handler

Dim db As Database
Set db = CurrentDb()

db.Execute "INSERT INTO tbllab_tech_schedule_alterations(scheduletemplateid,username) VALUES (" & Form_frmLab_tech_schedule_details.recordId & ",'" & Environ("username") & "');"
TempVars.Add "techSchedId", db.OpenRecordset("SELECT @@identity")(0).Value

Set db = Nothing

Call registerLabUpdates("tbllab_tech_schedule_alterations", TempVars!techSchedId, "Tech Schedule Alteration", "", "Created", Form_frmLab_tech_schedule_details.recordId, Me.name)
Me.Requery

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub deleteItem_Click()
On Error GoTo Err_Handler

Call registerLabUpdates("tbllab_tech_schedule_alterations", Me.recordId, "Tech Schedule Alteration", "", "Deleted", Form_frmLab_tech_schedule_details.recordId, Me.name)
dbExecute "DELETE from tbllab_tech_schedule_alterations WHERE recordid = " & Me.recordId
Me.Requery

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub scheduledate_AfterUpdate()
On Error GoTo Err_Handler

Call registerLabUpdates("tbllab_tech_schedule_alterations", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Form_frmLab_tech_schedule_details.recordId, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub schedulehours_AfterUpdate()
On Error GoTo Err_Handler

Call registerLabUpdates("tbllab_tech_schedule_alterations", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Form_frmLab_tech_schedule_details.recordId, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub schedulereason_AfterUpdate()
On Error GoTo Err_Handler

Call registerLabUpdates("tbllab_tech_schedule_alterations", Me.recordId, Me.ActiveControl.name, "", Me.ActiveControl.column(1), Form_frmLab_tech_schedule_details.recordId, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
