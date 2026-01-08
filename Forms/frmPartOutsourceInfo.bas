Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub assemblyType_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartAssemblyInfo", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.lblPartNumber.Caption, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub btnSave_Click()
On Error GoTo Err_Handler

If Me.Dirty Then Me.Dirty = False

If Nz(Form_frmPartInformation.outsourceInfoId, 0) <> Me.recordId Then
    Form_frmPartInformation.outsourceInfoId = Me.recordId
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Me.allowEdits = (Not restrict(Environ("username"), TempVars!projectOwner)) And Form_frmPartInformation.dataFreeze = False 'only project/service can edit things in this form, and only when dataFreeze is false

If Form_frmPartInformation.dataFreeze Then
    lblLock.Caption = "Data Frozen"
Else
    lblLock.Caption = "only PE's can edit"
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub outsourceCost_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartOutsourceInfo", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.lblPartNumber.Caption, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub outsourceVendor_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartOutsourceInfo", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.lblPartNumber.Caption, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
