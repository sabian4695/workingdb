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
If Nz(Me.userName) = 0 Then errorArray.Add "Username is Blank"
If IsNull(Me.sunhours) Then errorArray.Add "Sunday Hours is Blank"
If IsNull(Me.monhours) Then errorArray.Add "Monday Hours is Blank"
If IsNull(Me.tuehours) Then errorArray.Add "Tuesday Hours is Blank"
If IsNull(Me.wedhours) Then errorArray.Add "Wednesday Hours is Blank"
If IsNull(Me.thuhours) Then errorArray.Add "Thursday Hours is Blank"
If IsNull(Me.frihours) Then errorArray.Add "Friday Hours is Blank"
If IsNull(Me.sathours) Then errorArray.Add "Saturday Hours is Blank"

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

Call registerLabUpdates("tbllab_tech_schedule_template", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.recordId, Me.name)

Exit Function
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Function

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Me.sfrmCalendar_universal!selYear = Year(Date)
Me.sfrmCalendar_universal!selMonth = Month(Date)
Me.sfrmCalendar_universal!sqlSel = "SUM(schedulehours) AS cTasks, C as cDate"
Me.sfrmCalendar_universal!sqlWhere = "id = " & Me.recordId
Me.sfrmCalendar_universal!sqlGroupBy = "C"
Me.sfrmCalendar_universal!sqlFrom = "qryLab_tech_schedule_calendar"

Me.sfrmCalendar_universal.Form.universal_drawdatebuttons

Me.sfrmCalendar_universal!sfrmCalendar_universal_items.Form.RecordSource = Me.sfrmCalendar_universal!sqlFrom
Me.sfrmCalendar_universal!sfrmCalendar_universal_items.Form.filter = "C = #" & Date & "# AND " & Me.sfrmCalendar_universal!sqlWhere
Me.sfrmCalendar_universal!sfrmCalendar_universal_items.Form.FilterOn = True

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub imgUser_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmUserProfile", , , "user = '" & Me.userName & "'"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub newAlteration_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmLab_tech_schedule_alteration_create", acNormal, , , acFormAdd

Form_frmLab_tech_schedule_alteration_create.userName = Me.userName
Form_frmLab_tech_schedule_alteration_create.userName.Locked = True

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub remove_Click()
On Error GoTo Err_Handler

If MsgBox("Are you sure you want to delete this resource?", vbYesNo, "Please confirm") = vbYes Then
    Call registerLabUpdates("tbllab_tech_schedule_template", Me.recordId, "Tech Schedule", "", "Deleted", Me.recordId, Me.name)
    dbExecute ("DELETE FROM tbllab_tech_schedule_template WHERE [recordId] = " & Me.recordId)
    DoCmd.CLOSE acForm, "frmLab_tech_schedule_details"
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub save_Click()
On Error GoTo Err_Handler

If Me.Dirty Then Me.Dirty = False
If validate = False Then Exit Sub

Call registerLabUpdates("tbllab_tech_schedule_template", Me.recordId, "Tech Schedule", "", "Saved", Me.recordId, Me.name)
DoCmd.CLOSE acForm, "frmLab_tech_schedule_details"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub techSchHistory_Click()
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

Private Sub username_AfterUpdate()
On Error GoTo Err_Handler

Call registerLabUpdates("tbllab_tech_schedule_template", Me.recordId, Me.ActiveControl.name, "", Me.ActiveControl.column(1), Me.recordId, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
