Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub annealing_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartAssemblyInfo", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.lblPartNumber.Caption, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub assemblyType_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartAssemblyInfo", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.lblPartNumber.Caption, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub assemblyWeight_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartAssemblyInfo", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.lblPartNumber.Caption, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub assemblyWeight100Pc_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartAssemblyInfo", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.lblPartNumber.Caption, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub btnSave_Click()
On Error GoTo Err_Handler

If Me.Dirty Then Me.Dirty = False

If Nz(DLookup("assemblyInfoId", "tblPartInfo", "recordId = " & Me.lblPartInfoId.Caption), 0) <> Me.recordId Then
    dbExecute "UPDATE tblPartInfo SET assemblyInfoId = " & Me.recordId & " WHERE recordId = " & Me.lblPartInfoId.Caption
    Form_frmPartInformation.Requery
End If

DoCmd.CLOSE acForm, "frmPartAssemblyInfo"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Dim allowEdits As Boolean
allowEdits = (Not restrict(Environ("username"), TempVars!projectOwner)) And Form_frmPartInformation.dataFreeze = False 'only project/service can edit things in this form, and only when dataFreeze is false

Me.allowEdits = allowEdits

If Form_frmPartInformation.dataFreeze Then
    lblLock.Caption = "Data Frozen"
Else
    lblLock.Caption = "only PE's can edit"
End If

Form_sfrmPartComponents.AllowAdditions = allowEdits
Form_sfrmPartComponents.allowEdits = allowEdits
Form_sfrmPartComponents.remove.Visible = allowEdits

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub assemblyInspection_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartAssemblyInfo", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.lblPartNumber.Caption, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub assemblyMeasPack_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartAssemblyInfo", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.lblPartNumber.Caption, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub machineLine_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartAssemblyInfo", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.lblPartNumber.Caption, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub partsPerHour_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartAssemblyInfo", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.lblPartNumber.Caption, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub resource_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartAssemblyInfo", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.lblPartNumber.Caption, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
