Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim filt As String
Dim userID

Private Sub updateButtons(role As String, status As String)
On Error GoTo Err_Handler
userID = DLookup("ID", "tblPermissions", "user = '" & Me.fltAssignee & "'")

If role <> "" Then
    Me.filtAssignee.tag = "btn.L2"
    Me.filtCheck.tag = "btn.L2"
    Me.filtAll.tag = "btn.L2"
End If

If status <> "" Then
    Me.fltOpen.tag = "btn.L2"
    Me.fltClosed.tag = "btn.L2"
    Me.fltPend.tag = "btn.L2"
End If

Dim reqAssFilt As String
reqAssFilt = "([Requester] = '" & LCase(Environ("username")) & "' OR [Assignee] = " & userID & ")"

Dim openFilt As String, pendFilt As String, closedFilt As String
openFilt = "([Approval_Status] = 2 AND [Completed_Date] IS NULL)"
pendFilt = "([Approval_Status] = 1 AND [Completed_Date] IS NULL)"
closedFilt = "([Approval_Status] = 2 AND [Completed_Date] IS NOT NULL)"

Dim assFilt As String, checkFilt As String, allFilt As String, allPendFilt As String
assFilt = "([Assignee] = " & userID & ")"
checkFilt = "([Checker_1] = " & userID & " OR [Checker_2] = " & userID & ")"
If TempVars!iLevel = 1 Then
    allFilt = ""
    allPendFilt = ""
Else
    allFilt = "(" & assFilt & " OR " & checkFilt & ") AND "
    allPendFilt = "(" & reqAssFilt & " OR " & checkFilt & ") AND "
End If

Dim openCount As Long, pendCount As Long, closeCount As Long

Select Case role
    Case "Assignee"
        Me.filtAssignee.tag = "btn.L4"
        openCount = DCount("[Control_Number]", "[dbo_tblDRS]", assFilt & " AND " & openFilt)
        pendCount = DCount("[Control_Number]", "[dbo_tblDRS]", reqAssFilt & " AND " & pendFilt)
        closeCount = DCount("[Control_Number]", "[dbo_tblDRS]", assFilt & " AND " & closedFilt)
    Case "Checker"
        Me.filtCheck.tag = "btn.L4"
        openCount = DCount("[Control_Number]", "[dbo_tblDRS]", checkFilt & " AND " & openFilt)
        pendCount = DCount("[Control_Number]", "[dbo_tblDRS]", checkFilt & " AND " & pendFilt)
        closeCount = DCount("[Control_Number]", "[dbo_tblDRS]", checkFilt & " AND " & closedFilt)
    Case "All"
        Me.filtAll.tag = "btn.L4"
        openCount = DCount("[Control_Number]", "[dbo_tblDRS]", allFilt & openFilt)
        pendCount = DCount("[Control_Number]", "[dbo_tblDRS]", allPendFilt & pendFilt)
        closeCount = DCount("[Control_Number]", "[dbo_tblDRS]", allFilt & closedFilt)
End Select

If openCount + pendCount + closeCount > 0 Then
    Me.fltOpen.Caption = " Open (" & openCount & ")"
    Me.fltPend.Caption = " Pending (" & pendCount & ")"
    Me.fltClosed.Caption = " Closed (" & closeCount & ")"
End If

Dim assCount As Long, checkCount As Long, allCount As Long

Select Case status
    Case "Pending"
        Me.fltPend.tag = "btn.L4"
        assCount = DCount("[Control_Number]", "[dbo_tblDRS]", reqAssFilt & " AND " & pendFilt)
        checkCount = DCount("[Control_Number]", "[dbo_tblDRS]", checkFilt & " AND " & pendFilt)
        allCount = DCount("[Control_Number]", "[dbo_tblDRS]", allPendFilt & pendFilt)
    Case "Closed"
        Me.fltClosed.tag = "btn.L4"
        assCount = DCount("[Control_Number]", "[dbo_tblDRS]", assFilt & " AND " & closedFilt)
        checkCount = DCount("[Control_Number]", "[dbo_tblDRS]", checkFilt & " AND " & closedFilt)
        allCount = DCount("[Control_Number]", "[dbo_tblDRS]", allFilt & closedFilt)
    Case "Approved"
        Me.fltOpen.tag = "btn.L4"
        assCount = DCount("[Control_Number]", "[dbo_tblDRS]", assFilt & " AND " & openFilt)
        checkCount = DCount("[Control_Number]", "[dbo_tblDRS]", checkFilt & " AND " & openFilt)
        allCount = DCount("[Control_Number]", "[dbo_tblDRS]", allFilt & openFilt)
End Select

If assCount + checkCount + allCount > 0 Then
    Me.filtAssignee.Caption = "Assignee (" & assCount & ")"
    Me.filtCheck.Caption = "Checker (" & checkCount & ")"
    Me.filtAll.Caption = "All (" & allCount & ")"
End If

Call setTheme(Me)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Function grabStatus()
    Select Case True
        Case Me.fltOpen.tag = "btn.L4"
            grabStatus = "Approved"
        Case Me.fltClosed.tag = "btn.L4"
            grabStatus = "Closed"
        Case Me.fltPend.tag = "btn.L4"
            grabStatus = "Pending"
    End Select
End Function

Function grabRole()
Select Case True
    Case Me.filtAssignee.tag = "btn.L4"
        grabRole = "Assignee"
    Case Me.filtCheck.tag = "btn.L4"
        grabRole = "Checker"
    Case Me.filtAll.tag = "btn.L4"
        grabRole = "All"
End Select
End Function

Private Sub btnEditImage_Click()
On Error GoTo Err_Handler

Form_DASHBOARD.partNumberSearch = Me.Part_Number
DoCmd.OpenForm "frmPartPicture"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub classCodes_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmClassCodes"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub designWOdashboard_Click()
On Error GoTo Err_Handler

openPath (mainFolder(Me.ActiveControl.name))

DoCmd.CLOSE acForm, "frmCheckerFunctions"
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub designWOhelp_Click()
On Error GoTo Err_Handler

Call openPath(mainFolder(Me.ActiveControl.name))
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub drsFull_Click()
On Error GoTo Err_Handler

TempVars.Add "drsNewRec", "False"
DoCmd.OpenForm "frmApproveDRS"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub filtAll_Click()
On Error GoTo Err_Handler

If TempVars!iLevel <> 1 Then
    filt = "(([Checker1] = '" & Me.fltAssignee.Value & "' OR [Checker2] = '" & Me.fltAssignee.Value & "') OR "
    filt = filt & "([Requester] = '" & LCase(Me.fltAssignee.Value) & "' OR [Assignee] = '" & Me.fltAssignee.Value & "')) AND " & "[Approval_Status] = "
Else
    filt = "[Approval_Status] = "
End If

Select Case grabStatus()
    Case "Pending"
        filt = filt & "1 AND [Completed_Date] IS NULL"
    Case "Closed"
        filt = filt & "2 AND [Completed_Date] IS NOT NULL"
    Case "Approved"
        filt = filt & "2 AND [Completed_Date] IS NULL"
End Select

Call updateButtons("All", "")

Me.filter = filt
Me.FilterOn = True
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub filtAssignee_Click()
On Error GoTo Err_Handler

Select Case grabStatus()
    Case "Pending"
        filt = "([Requester] = '" & LCase(Me.fltAssignee.Value) & "' OR [Assignee] = '" & Me.fltAssignee.Value & "') AND " & "[Approval_Status] = 1"
    Case "Closed"
        filt = "[Assignee] = '" & Me.fltAssignee.Value & "' AND " & "[Approval_Status] = 2 AND [Completed_Date] IS NOT NULL"
    Case "Approved"
        filt = "[Assignee] = '" & Me.fltAssignee.Value & "' AND " & "[Approval_Status] = 2 AND [Completed_Date] IS NULL"
End Select

Call updateButtons("Assignee", "")

Me.filter = filt
Me.FilterOn = True
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub filtCheck_Click()
On Error GoTo Err_Handler

filt = "([Checker1] = '" & Me.fltAssignee.Value & "' OR [Checker2] = '" & Me.fltAssignee.Value & "') AND [Approval_Status] = "
Select Case grabStatus()
    Case "Pending"
        filt = filt & "1 AND [Completed_Date] IS NULL"
    Case "Closed"
        filt = filt & "2 AND [Completed_Date] IS NOT NULL"
    Case "Approved"
        filt = filt & "2 AND [Completed_Date] IS NULL"
End Select

Call updateButtons("Checker", "")

Me.filter = filt
Me.FilterOn = True
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub filterToMe_Click()
On Error GoTo Err_Handler

Dim showFilters As Boolean
showFilters = True

If Me.ActiveControl Then
    showFilters = False
    filt = "([Checker1] = '" & Me.fltAssignee.Value & "' AND Check_In_Prog = 'In Check' AND Completed_Date is null) OR ([Checker2] = '" & Me.fltAssignee.Value & "' AND Check_In_Prog = 'In Approval') AND approval_Status = 2 AND Completed_Date is null"
    Me.filter = filt
    Me.FilterOn = True
Else
    Select Case grabRole()
        Case "Assignee"
            Call filtAssignee_Click
        Case "Checker"
            Call filtCheck_Click
        Case "All"
            Call filtAll_Click
    End Select
End If

Me.Label31.Visible = showFilters
Me.fltAssignee.Visible = showFilters
Me.Label25.Visible = showFilters
Me.filtAssignee.Visible = showFilters
Me.filtCheck.Visible = showFilters
Me.filtAll.Visible = showFilters
Me.Label22.Visible = showFilters
Me.fltOpen.Visible = showFilters
Me.fltPend.Visible = showFilters
Me.fltClosed.Visible = showFilters

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub fltAssignee_AfterUpdate()
On Error GoTo Err_Handler

If Nz(Me.fltAssignee, "") = "" Then
    Call snackBox("error", "oh no", "Please fill in the assignee filter", Me.name)
    Exit Sub
End If

Select Case grabRole()
    Case "Assignee"
        Call filtAssignee_Click
    Case "Checker"
        Call filtCheck_Click
    Case "All"
        Call filtAll_Click
End Select

Call updateButtons(grabRole(), grabStatus())

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub fltClosed_Click()
On Error GoTo Err_Handler

Dim filtStart As String, filtAll As String
filtStart = "[Completed_Date] IS NOT NULL AND [Approval_Status] = 2"

If TempVars!iLevel <> 1 Then
    filtAll = filtStart & " AND (([Checker1] = '" & Me.fltAssignee.Value & "' OR [Checker2] = '" & Me.fltAssignee.Value & "') OR [Assignee] = '" & Me.fltAssignee.Value & "')"
Else
    filtAll = filtStart
End If

Select Case grabRole()
    Case "Assignee"
        filt = filtStart & " AND [Assignee] = '" & Me.fltAssignee.Value & "'"
    Case "Checker"
        filt = filtStart & " AND ([Checker1] = '" & Me.fltAssignee.Value & "' OR [Checker2] = '" & Me.fltAssignee.Value & "')"
    Case "All"
        filt = filtAll
End Select

Me.Due.ControlSource = "Completed_Date"
Me.lblDue.Caption = "Completed"

Me.filter = filt
Me.FilterOn = True

Call updateButtons("", "Closed")
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub fltOpen_Click()
On Error GoTo Err_Handler

Dim filtStart As String, filtAll As String
filtStart = "[Completed_Date] IS NULL AND [Approval_Status] = 2"

If TempVars!iLevel <> 1 Then
    filtAll = filtStart & " AND (([Checker1] = '" & Me.fltAssignee.Value & "' OR [Checker2] = '" & Me.fltAssignee.Value & "') OR [Assignee] = '" & Me.fltAssignee.Value & "')"
Else
    filtAll = filtStart
End If

Select Case grabRole()
    Case "Assignee"
        filt = filtStart & " AND [Assignee] = '" & Me.fltAssignee.Value & "'"
    Case "Checker"
        filt = filtStart & " AND ([Checker1] = '" & Me.fltAssignee.Value & "' OR [Checker2] = '" & Me.fltAssignee.Value & "')"
    Case "All"
        filt = filtAll
End Select

Me.Due.ControlSource = "Due"
Me.lblDue.Caption = "Due -"

Me.filter = filt
Me.FilterOn = True

Call updateButtons("", "Approved")
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub fltPend_Click()
On Error GoTo Err_Handler

Dim filtStart As String, filtAll As String
filtStart = "[Completed_Date] IS NULL AND [Approval_Status] = 1"

If TempVars!iLevel <> 1 Then
    filtAll = filtStart & " AND (([Checker1] = '" & Me.fltAssignee.Value & "' OR [Checker2] = '" & Me.fltAssignee.Value & "') OR ([Requester] = '" & LCase(Me.fltAssignee.Value) & "' OR  [Assignee] = '" & Me.fltAssignee.Value & "'))"
Else
    filtAll = filtStart
End If

Select Case grabRole()
    Case "Assignee"
        filt = filtStart & " AND ([Requester] = '" & LCase(Me.fltAssignee.Value) & "' OR [Assignee] = '" & Me.fltAssignee.Value & "')"
    Case "Checker"
        filt = filtStart & " AND ([Checker1] = '" & Me.fltAssignee.Value & "' OR [Checker2] = '" & Me.fltAssignee.Value & "')"
    Case "All"
        filt = filtAll
End Select

Me.filter = filt
Me.FilterOn = True

Call updateButtons("", "Pending")
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Dim iLevel, userID, checkerBool As Boolean
Me.fltAssignee.Value = Environ("username")

iLevel = Nz(DLookup("[designWOpermissions]", "tblPermissions", "[user] = '" & Environ("username") & "'"), 3)
checkerBool = iLevel = 1

TempVars.Add "iLevel", iLevel

Me.drsFull.Visible = checkerBool
Me.designWOdashboard.Visible = checkerBool
Me.fltAssignee.Visible = checkerBool
Me.Label31.Visible = checkerBool

userID = DLookup("[ID]", "[tblPermissions]", "[user] = '" & Me.fltAssignee.Value & "'")
Me.Repaint
Call updateButtons("Assignee", "Approved")

Me.OrderBy = "Due"
Me.OrderByOn = True

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub pendingAllDRS_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmDRScheckerTracker"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub fullDRS_Click()
On Error GoTo Err_Handler

If Nz(DLookup("[designWOpermissions]", "tblPermissions", "[user] = '" & Environ("username") & "'"), 3) = 1 Then
    TempVars.Add "drsNewRec", "False"
    DoCmd.OpenForm "frmApproveDRS", , , "[Control_Number] = " & Me.Control_Number
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblApprover_Click()
On Error GoTo Err_Handler

Me.Checker2L.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblAssignee_Click()
On Error GoTo Err_Handler

Me.Assignee.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblChecker_Click()
On Error GoTo Err_Handler

Me.Checker1L.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblComments_Click()
On Error GoTo Err_Handler

Me.Comments.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblCtrl_Click()
On Error GoTo Err_Handler

Me.Control_Number.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblDescription_Click()
On Error GoTo Err_Handler

Me.PART_DESCRIPTION.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblDue_Click()
On Error GoTo Err_Handler

Me.Due.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblPacketStatus_Click()
On Error GoTo Err_Handler

Me.Check_In_Prog.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblPN_Click()
On Error GoTo Err_Handler

Me.Part_Number.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblType_Click()
On Error GoTo Err_Handler

Me.Request_Type.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub openDetails_Click()
On Error GoTo Err_Handler

TempVars.Add "controlNumber", Me.Control_Number.Value
If CurrentProject.AllForms("frmDRSdashboard").IsLoaded = True Then DoCmd.CLOSE acForm, "frmDRSdashboard"

On Error Resume Next
DoCmd.OpenForm "frmDRSdashboard"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Public Sub refresh_Click()
On Error GoTo Err_Handler
Me.Requery
userID = DLookup("[ID]", "[tblPermissions]", "[user] = '" & Me.fltAssignee.Value & "'")
Call updateButtons(grabRole(), grabStatus())
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub requestDRS_Click()
On Error GoTo Err_Handler

TempVars.Add "drsNewRec", "True"
DoCmd.OpenForm "frmApproveDRS"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub searchPN_Click()
On Error GoTo Err_Handler

Form_DASHBOARD.partNumberSearch = Me.Part_Number
Form_DASHBOARD.filterbyPN_Click
Form_DASHBOARD.SetFocus

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub scoreCard_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmDRSscoreCard"
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub time_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmTimeView"
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
