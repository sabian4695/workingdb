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
If Nz(Me.Facility) = 0 Then errorArray.Add "Facility is Blank"
If Nz(Me.resourcename) = "" Then errorArray.Add "Resource Name is Blank"
If Nz(Me.resourcetype) = 0 Then errorArray.Add "Resource Type is Blank"

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

Function lab_afterupdate()
On Error GoTo Err_Handler

Call registerLabUpdates("tbllab_resources", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.recordId, Me.name)

Exit Function
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Function

Private Sub facility_AfterUpdate()
On Error GoTo Err_Handler

Call registerLabUpdates("tbllab_resources", Me.recordId, Me.ActiveControl.name, "", Me.ActiveControl.column(1), Me.recordId, Me.name)

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
End If

Me.sfrmCalendar_universal!selYear = Year(Date)
Me.sfrmCalendar_universal!selMonth = Month(Date)
Me.sfrmCalendar_universal!sqlSel = "COUNT(recordId) AS cTasks, C as cDate"
Me.sfrmCalendar_universal!sqlWhere = "id = " & Me.recordId
Me.sfrmCalendar_universal!sqlGroupBy = "C"
Me.sfrmCalendar_universal!sqlFrom = "qryLab_resource_schedule"

Me.sfrmCalendar_universal.Form.universal_drawdatebuttons

Me.sfrmCalendar_universal!sfrmCalendar_universal_items.Form.RecordSource = Me.sfrmCalendar_universal!sqlFrom
Me.sfrmCalendar_universal!sfrmCalendar_universal_items.Form.filter = "C = #" & Date & "# AND " & Me.sfrmCalendar_universal!sqlWhere
Me.sfrmCalendar_universal!sfrmCalendar_universal_items.Form.FilterOn = True

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

Form_frmLab_Resources.Requery
Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Unload", Err.DESCRIPTION, Err.number)
End Sub

Private Sub remove_Click()
On Error GoTo Err_Handler

If MsgBox("Are you sure you want to delete this resource?", vbYesNo, "Please confirm") = vbYes Then
    Call registerLabUpdates("tbllab_resources", Me.recordId, "Resource", "", "Deleted", Me.recordId, Me.name)
    dbExecute ("DELETE FROM tbllab_resources WHERE [recordId] = " & Me.recordId)
    DoCmd.CLOSE acForm, "frmLab_resource_details"
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub resourceHistory_Click()
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

Private Sub resourcename_AfterUpdate()
On Error GoTo Err_Handler

Call registerLabUpdates("tbllab_resources", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.recordId, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub resourcetype_AfterUpdate()
On Error GoTo Err_Handler

Call registerLabUpdates("tbllab_resources", Me.recordId, Me.ActiveControl.name, "", Me.ActiveControl.column(1), Me.recordId, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub save_Click()
On Error GoTo Err_Handler

If Me.Dirty Then Me.Dirty = False
If validate = False Then Exit Sub

Call registerLabUpdates("tbllab_resources", Me.recordId, "Resource", "", "Saved", Me.recordId, Me.name)
DoCmd.CLOSE acForm, "frmLab_resource_details"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
