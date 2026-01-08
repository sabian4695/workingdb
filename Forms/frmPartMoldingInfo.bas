Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub annealing_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartMoldingInfo", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.lblPartNumber.Caption, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub assignedPress_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartMoldingInfo", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.lblPartNumber.Caption, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub assignTool_Click()
On Error GoTo Err_Handler

Me.existingTool.Visible = True

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub automated_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartMoldingInfo", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.lblPartNumber.Caption, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub btnSave_Click()
On Error GoTo Err_Handler

If IsNull(Me.recordId) Then
    Call snackBox("error", "Hold on", "Please alter a field to save", Me.name)
    Exit Sub
End If

If Me.Dirty Then Me.Dirty = False

If Nz(DLookup("moldInfoId", "tblPartInfo", "recordId = " & Me.lblPartInfoId.Caption), 0) <> Me.recordId Then
    dbExecute "UPDATE tblPartInfo SET moldInfoId = " & Me.recordId & " WHERE recordId = " & Me.lblPartInfoId.Caption
End If

'SCAN THROUGH STEPS AND SEE IF CUSTOM ACTION IS SET UP FOR THIS FUNCTION
Call scanSteps(Form_frmPartDashboard.partNumber, "frmPartMoldingInfo_save", Me.recordId)

DoCmd.CLOSE acForm, "frmPartMoldingInfo"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub cavitation_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartMoldingInfo", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.lblPartNumber.Caption, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub existingTool_AfterUpdate()
On Error GoTo Err_Handler

DoCmd.applyFilter , "recordId = " & Me.existingTool

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub familyTool_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartMoldingInfo", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.lblPartNumber.Caption, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Dim dataFreeze As Boolean
Dim editOK As Boolean

dataFreeze = True
If CurrentProject.AllForms("frmPartInformation").IsLoaded Then 'if loaded from Part Dashboard
    dataFreeze = Form_frmPartInformation.dataFreeze 'only check for data freeze if frmPartInfo is open
    editOK = (Not restrict(Environ("username"), TempVars!projectOwner)) And dataFreeze = False 'only project/service can edit things in this form, and only when dataFreeze is false
Else 'if not loaded from part dashboard
    editOK = False
End If

Me.allowEdits = editOK
Me.assignTool.Visible = editOK
Me.unassignTool.Visible = editOK

If dataFreeze Then
    lblLock.Caption = "Data Frozen"
Else
    lblLock.Caption = "only PE's can edit"
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub gateCutting_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartMoldingInfo", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.lblPartNumber.Caption, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub insertMold_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartMoldingInfo", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.lblPartNumber.Caption, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub inspection_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartMoldingInfo", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.lblPartNumber.Caption, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub measurePack_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartMoldingInfo", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.lblPartNumber.Caption, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub partsProduced_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartMoldingInfo", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.lblPartNumber.Caption, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub piecesPerHour_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartMoldingInfo", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.lblPartNumber.Caption, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub pressSize_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartMoldingInfo", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.lblPartNumber.Caption, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub shotsPerHour_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartMoldingInfo", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.lblPartNumber.Caption, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub toolNumber_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartMoldingInfo", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.lblPartNumber.Caption, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub toolOwner_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartMoldingInfo", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.lblPartNumber.Caption, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub toolReason_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartMoldingInfo", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.lblPartNumber.Caption, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub toolType_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartMoldingInfo", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.lblPartNumber.Caption, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub twinShot_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartMoldingInfo", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.lblPartNumber.Caption, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub unassignTool_Click()
On Error GoTo Err_Handler

DoCmd.GoToRecord , , acNewRec

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
