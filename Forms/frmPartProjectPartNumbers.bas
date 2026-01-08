Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub allHistory_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmHistory", acNormal, , "[tableName] = 'tblPartProjectPartNumbers' AND [partNumber] = '" & Form_frmPartDashboard.partNumber & "'"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub childPartNumberType_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartProjectPartNumbers", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Form_frmPartDashboard.partNumber, Nz(Me.childPartNumberType.column(1)), Form_frmPartDashboard.recordId)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub doAction_Click()
On Error GoTo Err_Handler

If Nz(Me.recordId, "") = "" Then Exit Sub
doActionSelect (Me.childPartNumber)
DoCmd.CLOSE acForm, "frmPartProjectPartNumbers"

If Nz(TempVars!partDashAction, "") = "frmPartInformation" Then Form_frmPartInformation.SetFocus

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub doActionMaster_Click()
On Error GoTo Err_Handler

doActionSelect (Me.masterPN)
DoCmd.CLOSE acForm, "frmPartProjectPartNumbers"

If Nz(TempVars!partDashAction, "") = "frmPartInformation" Then Form_frmPartInformation.SetFocus

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Function doActionSelect(partNumber As String)
On Error GoTo Err_Handler

Select Case Nz(TempVars!partDashAction, "")
    Case "frmPartInformation"
        TempVars.Add "partNumber", partNumber
        DoCmd.OpenForm "frmPartInformation"
    Case "AIF"
        Call autoUploadAIF(partNumber)
    Case "frmDropFile"
        Form_frmDropFile.TpartNumber = partNumber
    Case "rptPartOpenIssues"
        DoCmd.OpenReport "rptPartOpenIssues", acViewPreview, , "partNumber = '" & partNumber & "'"
    Case "rptPartInformation"
        DoCmd.OpenReport "rptPartInformation", acViewPreview, , "[partNumber]='" & partNumber & "'"
    Case "frmPartIssues"
        If Form_frmPartDashboard.frmPartIssues.SourceObject = "" Then
            Form_frmPartDashboard.frmPartIssues.SourceObject = "frmPartIssues"
            Form_frmPartDashboard.frmPartIssues.LinkChildFields = ""
            Form_frmPartDashboard.frmPartIssues.LinkMasterFields = ""
        End If
        Form_frmPartDashboard.frmPartIssues.Form.filter = "partNumber = '" & partNumber & "' AND [closeDate] is null"
        Form_frmPartDashboard.frmPartIssues.Form.FilterOn = True
        Form_frmPartDashboard.frmPartIssues.Form.Controls("fltPartNumber") = partNumber
        Form_frmPartDashboard.frmPartIssues.Visible = True
End Select

Exit Function
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Function

Private Sub Form_AfterInsert()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartProjectPartNumbers", Me.recordId, "Child Part Number", Nz(Me.childPartNumber), "Created", Form_frmPartDashboard.partNumber, Nz(Me.childPartNumberType.column(1)), Form_frmPartDashboard.recordId)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Dim allowStuff As Boolean
allowStuff = Not restrict(Environ("username"), Nz(TempVars!projectOwner, ""))

Me.allowEdits = allowStuff
Me.AllowAdditions = allowStuff
Me.remove.Visible = allowStuff

Me.projectId.DefaultValue = Form_frmPartDashboard.recordId
Me.masterPN = Form_frmPartDashboard.partNumber

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Err_Handler

Dim rs1 As Recordset, errorTxt As String
Set rs1 = Me.RecordsetClone


errorTxt = ""

If rs1.RecordCount > 0 Then
    rs1.MoveFirst
    Do While Not rs1.EOF
        If Len(rs1!childPartNumber) < 5 Then errorTxt = "Part Number must be >4 digits"
        If Nz(rs1!childPartNumberType) = 0 Then errorTxt = "Please select part number type"
        If rs1!childPartNumber = Me.masterPN Then errorTxt = rs1!childPartNumber & " is your master part number. It cannot also be a related part number. Please remove to continue."
        rs1.MoveNext
    Loop
End If

If errorTxt <> "" Then
    MsgBox errorTxt, vbCritical, "Please fix"
    Cancel = True
    GoTo exitThis
End If



exitThis:
rs1.CLOSE
Set rs1 = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Unload", Err.DESCRIPTION, Err.number)
End Sub

Private Sub partNumberType_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartProjectPartNumbers", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Form_frmPartDashboard.partNumber, Nz(Me.childPartNumberType.column(1)), Form_frmPartDashboard.recordId)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub remove_Click()
On Error GoTo Err_Handler

If IsNull(Me.recordId) Then
    MsgBox "This is an empty record.", vbInformation, "Can't do that"
    Exit Sub
End If

If MsgBox("Are you sure you want to delete this?", vbYesNo, "Please confirm") = vbYes Then
    Call registerPartUpdates("tblPartProjectPartNumbers", Me.recordId, "Child Part Number", Nz(Me.childPartNumber), "Deleted", Form_frmPartDashboard.partNumber, Nz(Me.childPartNumberType.column(1)), Form_frmPartDashboard.recordId)
    Dim db As Database
    Set db = CurrentDb()
    db.Execute ("DELETE FROM tblPartProjectPartNumbers WHERE [recordId] = " & Me.recordId)
    Set db = Nothing
    Me.Requery
    Call snackBox("success", "Success!", "Part Number deleted", Me.name)
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
