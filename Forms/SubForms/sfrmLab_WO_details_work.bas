Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub addWork_Click()
On Error GoTo Err_Handler

Dim db As Database
Set db = CurrentDb()

db.Execute "INSERT INTO tbllab_wo_work(woid) VALUES (" & Form_frmLab_WO_details.recordId & ");"
TempVars.Add "workoId", db.OpenRecordset("SELECT @@identity")(0).Value

Set db = Nothing

Call registerLabUpdates("tbllab_wo_work", TempVars!workoId, "WO Work", "", "Created", Form_frmLab_WO_details.recordId, Me.name)
Me.Requery

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub deleteItem_Click()
On Error GoTo Err_Handler

Call registerLabUpdates("tbllab_wo_work", Me.recordId, "WO Work", "", "Deleted", Form_frmLab_WO_details.recordId, Me.name)
dbExecute "DELETE from tbllab_wo_work WHERE recordid = " & Me.recordId
Me.Requery

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub esthours_AfterUpdate()
On Error GoTo Err_Handler

Call registerLabUpdates("tbllab_wo_work", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Form_frmLab_WO_details.recordId, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub partnumber_AfterUpdate()
On Error GoTo Err_Handler

Call registerLabUpdates("tbllab_wo_work", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Form_frmLab_WO_details.recordId, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub worktype_AfterUpdate()
On Error GoTo Err_Handler

Call registerLabUpdates("tbllab_wo_work", Me.recordId, Me.ActiveControl.name, "", Me.ActiveControl.column(1), Form_frmLab_WO_details.recordId, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
