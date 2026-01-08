Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Dirty(Cancel As Integer)
On Error GoTo Err_Handler

Me.Parent.lastModified = Now

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub btnDeleteLabWO_Click()
On Error GoTo Err_Handler

If IsNull(Me.ID) Then Exit Sub

If MsgBox("Are you sure you want to delete this Lab WO?", vbYesNo, "Warning") = vbYes Then
    Call registerCPCUpdates("tblCPC_LabWOs", ID, Me.labWorkOrder.name, Me.labWorkOrder, "Deleted", Form_frmCPC_Dashboard.ID)
    dbExecute ("DELETE * FROM tblCPC_LabWOs WHERE [id] = " & Me.ID)
    Me.Requery
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub btnDetails_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmLabWOs", , , "WONumber = " & Replace(Me.labWorkOrder, "N", "")
Form_frmLabWOs.txtSearch = Me.labWorkOrder

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub labWorkOrder_AfterUpdate()
On Error GoTo Err_Handler
Call registerCPCUpdates("tblCPC_LabWOs", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Form_frmCPC_Dashboard.ID)
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
