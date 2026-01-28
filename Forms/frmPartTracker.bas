Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub automationGates_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmPartAutomationGates", , , "partNumber = '" & Me.partNumber & "'"
Form_frmPartAutomationGates.fltPartNumber = Me.partNumber

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Detail_Paint()
On Error Resume Next

Me.automationGates.Transparent = Nz(Me.partType, "") <> "Assembled"

End Sub

Private Sub details_Click()
On Error GoTo Err_Handler

openPartProject (Me.partNumber)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub export_Click()
On Error GoTo Err_Handler

Dim FileName As String, sqlString As String, filt As String
FileName = "H:\Part_Tracker_" & nowString & ".xlsx"
filt = ""
If Me.Form.filter <> "" And Me.Form.FilterOn Then filt = " WHERE " & Me.Form.filter
sqlString = "SELECT * FROM " & Me.RecordSource & filt

Call exportSQL(sqlString, FileName)

Exit Sub
Err_Handler:
Dim db As Database
Set db = CurrentDb()
db.QueryDefs.Delete "myExportQueryDef"
Set db = Nothing
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Function applyTheFilters()
On Error GoTo Err_Handler

If Me.showClosedToggle.Value = True Then
        If Me.RecordSource <> "qryPartTrackerClosed" Then Me.RecordSource = "qryPartTrackerClosed"
    Else
        If Me.RecordSource <> "qryPartTracker" Then Me.RecordSource = "qryPartTracker"
End If

Dim filt
filt = ""

If Me.onHoldToggle.Value Then
    filt = "partProjectStatus = 'On Hold'"
Else
    filt = "partProjectStatus <> 'On Hold'"
End If

If Me.fltUser <> "" Then filt = filt & " AND (partNumber IN (SELECT partNumber FROM tblPartTeam WHERE person = '" & Me.fltUser & "'))"
If Me.fltModel <> "" Then
    If filt <> "" Then filt = filt & " AND "
    filt = filt & " modelCode = '" & Me.fltModel & "'"
End If

Me.filter = filt
Me.FilterOn = filt <> ""

Exit Function
Err_Handler:
    Call handleError(Me.name, "applyTheFilters", Err.DESCRIPTION, Err.number)
End Function

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Me.OrderBy = "[Due] ASC"
Me.OrderByOn = True

If (restrict(Environ("username"), "Project") And restrict(Environ("username"), "Service")) Then Me.newPartProject.Visible = False
Me.nmqDash.Visible = Not restrict(Environ("username"), "New Model Quality")

Me.reports.Visible = Not restrict(Environ("username"), "New Model Quality", "Manager")

Me.chooseDept = Nz(userData("Dept"), "")

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblCustPN_Click()
On Error GoTo Err_Handler

Me.customerPN.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblDescription_Click()
On Error GoTo Err_Handler

Me.DESCRIPTION.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblGate_Click()
On Error GoTo Err_Handler

Me.currentGate.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblModelCode_Click()
On Error GoTo Err_Handler

Me.modelCode.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblNotes_Click()
On Error GoTo Err_Handler

Me.stepDescription.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblOem_Click()
On Error GoTo Err_Handler

Me.OEM.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblPN_Click()
On Error GoTo Err_Handler

Me.partNumber.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblStep_Click()
On Error GoTo Err_Handler

Me.currentStep.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblStepDue_Click()
On Error GoTo Err_Handler

Me.stepDue.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblType_Click()
On Error GoTo Err_Handler

Me.partType.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub newPartProject_Click()
On Error GoTo Err_Handler

Dim X
X = InputBox("Enter part number", "Input Part Number", Form_DASHBOARD.partNumberSearch)
If StrPtr(X) = 0 Or X = "" Then Exit Sub

Form_DASHBOARD.partNumberSearch = X
Call Form_DASHBOARD.filterbyPN_Click

openPartProject (X)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub nmqDash_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmNMQDashboard", , , "partNumber = '" & Me.partNumber & "'"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub partTrackingHelp_Click()
On Error GoTo Err_Handler

Dim dept As String
dept = userData("Dept")

Select Case dept
    Case "Service", "Project"
        Call openPath(mainFolder(Me.ActiveControl.name))
    Case "Tooling"
        Call openPath(mainFolder(Me.ActiveControl.name & "_Tooling"))
    Case Else
        Call openPath(mainFolder(Me.ActiveControl.name & "_Universal"))
End Select

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Public Sub refresh_Click()
On Error GoTo Err_Handler
Me.Requery
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub reports_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmReporting"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
