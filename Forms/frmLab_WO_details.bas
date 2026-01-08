Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Function validate()
On Error GoTo Err_Handler

validate = False

If IsNull(Me.recordId) Then
    validate = True
    Exit Function
End If

Dim errorArray As Collection
Set errorArray = New Collection

'check stuff
If Me.sfrmLab_WO_details_work.Form.Recordset.RecordCount = 0 Then errorArray.Add "No items found in work"
If Nz(Me.Facility) = 0 Then errorArray.Add "Facility is Blank"
If Nz(Me.Requestor) = "" Then errorArray.Add "Requestor is Blank"
If Nz(Me.WOStatus) = 0 Then errorArray.Add "WO Status is Blank"

If errorArray.count > 0 Then
    Dim errorTxtLines As String, element
    errorTxtLines = ""
    For Each element In errorArray
        errorTxtLines = errorTxtLines & vbNewLine & element
    Next element
    
    MsgBox "Please fix these items: " & vbNewLine & errorTxtLines, vbOKOnly, "ACTION REQUIRED"
    Exit Function
End If

validate = True

Exit Function
Err_Handler:
    Call handleError(Me.name, "validate", Err.DESCRIPTION, Err.number)
End Function

Private Sub dueDate_AfterUpdate()
On Error GoTo Err_Handler

Call registerLabUpdates("tbllab_wo_main", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.recordId, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub facility_AfterUpdate()
On Error GoTo Err_Handler

Call registerLabUpdates("tbllab_wo_main", Me.recordId, Me.ActiveControl.name, "", Me.ActiveControl.column(1), Me.recordId, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

If IsNull(Me.recordId) Then
    Me.createdBy = Environ("username")
    Me.createdDate = Date
    Me.WOStatus = 1
End If


Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Err_Handler

If Me.Dirty Then Me.Dirty = False
If validate = False Then
    Cancel = True
    Exit Sub
End If

Form_frmLab_WO_tracker.Requery
Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Unload", Err.DESCRIPTION, Err.number)
End Sub

Private Sub remove_Click()
On Error GoTo Err_Handler

If MsgBox("Are you sure you want to delete this WO?", vbYesNo, "Please confirm") = vbYes Then
    Call registerLabUpdates("tbllab_wo_main", Me.recordId, "WO", "", "DELETED", Me.recordId, Me.name)
    dbExecute ("DELETE FROM tbllab_wo_main WHERE [recordId] = " & Me.recordId)
    DoCmd.CLOSE acForm, "frmLab_WO_details"
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub requestor_AfterUpdate()
On Error GoTo Err_Handler

Call registerLabUpdates("tbllab_wo_main", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.recordId, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub save_Click()
On Error GoTo Err_Handler

If Me.Dirty Then Me.Dirty = False
If validate = False Then Exit Sub

Call registerLabUpdates("tbllab_wo_main", Me.recordId, "WO", "", "Saved", Me.recordId, Me.name)

DoCmd.CLOSE acForm, "frmLab_WO_details"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub sfrmLab_WO_details_resources_Enter()
On Error GoTo Err_Handler

Form_sfrmLab_WO_details_resources.workid.RowSource = "SELECT tbllab_wo_work.recordid, tbllab_wo_work.partnumber, tblDropDownsSP.lab_work_type " & _
    "FROM tbllab_wo_work LEFT JOIN tblDropDownsSP ON tbllab_wo_work.worktype = tblDropDownsSP.recordid WHERE tbllab_wo_work.woid = " & Form_frmLab_WO_details.recordId
    
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub woHistory_Click()
On Error GoTo Err_Handler
If IsNull(Me.recordId) = False Then
    DoCmd.OpenForm "frmHistory"
    Form_frmHistory.RecordSource = "qryLabUpdateTracking"
    Form_frmHistory.dataTag2.ControlSource = "formname"
    Form_frmHistory.dataTag0.ControlSource = "referenceid"
    Form_frmHistory.previousData.ControlSource = "previous"
    Form_frmHistory.newData.ControlSource = "new"
    Form_frmHistory.filter = "referenceid = " & Me.recordId
    Form_frmHistory.FilterOn = True
    Form_frmHistory.OrderBy = "updatedDate Desc"
    Form_frmHistory.OrderByOn = True
End If
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub wostatus_AfterUpdate()
On Error GoTo Err_Handler

Call registerLabUpdates("tbllab_wo_main", Me.recordId, Me.ActiveControl.name, "", Me.ActiveControl.column(1), Me.recordId, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
