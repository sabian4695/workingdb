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

db.Execute "INSERT INTO tbllab_resource_schedule(woid) VALUES (" & Form_frmLab_WO_details.recordId & ");"
TempVars.Add "resourceSchedId", db.OpenRecordset("SELECT @@identity")(0).Value

Set db = Nothing

Call registerLabUpdates("tbllab_resource_schedule", TempVars!resourceSchedId, "Resource Schedule", "", "Created", Form_frmLab_WO_details.recordId, Me.name)
Me.Requery

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub actualend_AfterUpdate()
On Error GoTo Err_Handler

Call registerLabUpdates("tbllab_resource_schedule", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.woid, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub deleteItem_Click()
On Error GoTo Err_Handler

Call registerLabUpdates("tbllab_resource_schedule", Me.recordId, "Resource Schedule", "", "Deleted", Me.woid, Me.name)
dbExecute "DELETE from tbllab_resource_schedule WHERE recordid = " & Me.recordId
Me.Requery

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub schedulehours_AfterUpdate()
On Error GoTo Err_Handler

Call registerLabUpdates("tbllab_resource_schedule", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.woid, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub schedulestart_AfterUpdate()
On Error GoTo Err_Handler

Call registerLabUpdates("tbllab_resource_schedule", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.woid, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub workid_AfterUpdate()
On Error GoTo Err_Handler

Call registerLabUpdates("tbllab_resource_schedule", Me.recordId, Me.ActiveControl.name, "", Me.ActiveControl.column(1), Me.woid, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
