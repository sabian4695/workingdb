Option Compare Database
Option Explicit

Public Function daysSinceLastNudgeCPC(stepId As Long)
On Error GoTo err_handler

Dim lastNudgeDate
lastNudgeDate = Nz(DLookup("updatedDate", "tblCPC_UpdateTracking", "tableRecordId = " & stepId & " AND columnName = 'Nudge'"), 0)

If Nz(lastNudgeDate, 0) = 0 Then
    daysSinceLastNudgeCPC = "N/A"
Else
    daysSinceLastNudgeCPC = Date - lastNudgeDate
End If

Exit Function
err_handler:
    Call handleError("wdbCPCfunctions", "daysSinceLastNudgeCPC", Err.DESCRIPTION, Err.number)
End Function

Function notifyCPC(projId As Long, notiType As String, stepTitle As String, Optional sendAlways As Boolean = False, Optional stepAction As Boolean = False) As Boolean
On Error GoTo err_handler

notifyCPC = False

Dim db As Database
Set db = CurrentDb()
Dim rsPartTeam As Recordset
Set rsPartTeam = db.OpenRecordset("SELECT * from tblCPC_XFteams where projectId = " & projId)
If rsPartTeam.RecordCount = 0 Then Exit Function

Do While Not rsPartTeam.EOF
    Dim rsPermissions As Recordset, sendTo As String
    If IsNull(rsPartTeam!memberName) Then GoTo nextRec
    sendTo = rsPartTeam!memberName
    Set rsPermissions = db.OpenRecordset("SELECT user, userEmail from tblPermissions where user = '" & sendTo & "' AND Dept = 'Project' AND Level = 'Engineer'")
    If rsPermissions.RecordCount = 0 Then GoTo nextRec
    If sendTo = Environ("username") And Not sendAlways Then GoTo nextRec
    
    'actually send notification
    Dim body As String, closedBy As String
    If stepAction Then
        closedBy = "stepAction"
    Else
        closedBy = getFullName()
    End If
    
    Dim bodyTitle As String, emailTitle As String, subjectLine As String
    subjectLine = "Step " & notiType
    emailTitle = "WDB Step " & notiType
    bodyTitle = "This step has been " & notiType
    
    body = emailContentGen(subjectLine, emailTitle, bodyTitle, stepTitle & " Issue", "Who: " & closedBy, "When: " & CStr(Date), "")
    Call sendNotification(sendTo, 10, 2, stepTitle & " has been " & notiType, body, "CPC Project", projId)
    
nextRec:
    rsPartTeam.MoveNext
Loop

notifyCPC = True

rsPartTeam.Close
Set rsPartTeam = Nothing
Set db = Nothing

Exit Function
err_handler:
    Call handleError("wdbCPCfunctions", "notifyCPC", Err.DESCRIPTION, Err.number)
End Function

Function closeCPCstep(stepId As Long, frmActive As String) As Boolean
On Error GoTo err_handler

closeCPCstep = False

Dim db As Database
Set db = CurrentDb()
Dim rsStep As Recordset, projectOwner As String
Dim errorText As String, testThis
errorText = ""
Set rsStep = db.OpenRecordset("SELECT * from tblCPC_Steps WHERE ID = " & stepId)

projectOwner = "CPC"

'check if all steps before this pillar are closed

If Not IsNull(rsStep!dueDate) And DCount("ID", "tblCPC_Steps", "projectId = " & rsStep!projectId & " AND indexOrder < " & rsStep!indexOrder & " AND [status] <> 'Closed'") > 0 Then
    errorText = "This step is a pillar. All steps before this pillar must be closed before this step."
    GoTo errorOut
End If

skipPillarCheck: 'BETA -TO BE REMOVED

If restrict(Environ("username"), projectOwner) = False Then GoTo theCorrectFellow 'is the bro an owner?

'FIRST: are you the right person for the job???
If Nz(rsStep!responsible) = userData("Dept") And DCount("ID", "tblCPC_XFteams", "person = '" & Environ("username") & "' AND projectId = " & rsStep!projectId) > 0 Then GoTo theCorrectFellow 'if the bro is responsible AND CHECK IF ON CF TEAM
If restrict(Environ("username"), projectOwner, "Manager") = False Then GoTo theCorrectFellow 'is the bro an owner Manager?
If restrict(Environ("username"), Nz(rsStep!responsible), "Manager") = False Then GoTo theCorrectFellow  'is the bro a manager in the department of the "responsible" person?
Call snackBox("error", "Woops", "Only the 'Responsible' person, their manager, or a project/service Manager can close a step", frmActive)
GoTo exit_handler
theCorrectFellow:

If IsNull(rsStep!closeDate) = False Then errorText = "This is already closed - what's the point in closing again?"
If getApprovalsCompleteCPC(rsStep!ID) < getTotalApprovalsCPC(rsStep!ID) Then errorText = "I spy with my little eye: open approval(s) on this step!"

'IF DOCUMENT REQUIRED, CHECK FOR DOCUMENTS
If Nz(rsStep!documentType, 0) <> 0 Then
    'First, check if any files are added. error out if not
    Dim countAttach As Long
    countAttach = DCount("ID", "tblPartAttachmentsSP", "partStepId = " & rsStep!ID)
    If countAttach = 0 Then
        errorText = "This step required a file to be added to close it"
        GoTo errorOut
    End If
End If

If errorText <> "" Then GoTo errorOut

'---CHECK STEP ACTIONS---
If IsNull(rsStep!stepActionId) Then GoTo stepActionOK

Dim rsStepAction As Recordset
Set rsStepAction = db.OpenRecordset("SELECT * from tblPartStepActions WHERE recordID = " & rsStep!stepActionId)

If rsStepAction.RecordCount = 0 Then GoTo stepActionOK 'no step action found
If rsStepAction!whenToRun <> "closeStep" And rsStepAction!whenToRun <> "firstTimeRun" Then GoTo stepActionOK 'check if this action should be running now. Ones marked "closeStep" are checks on close, meant to run now

Dim rsMoldInfo As Recordset

Select Case rsStepAction!stepAction
    Case "emailPartInfo"
        If emailPartInfo(rsStep!partNumber, Nz(rsStep!stepDescription)) = False Then Err.Raise vbObjectError + 999, , "Email couldn't send..."
    Case "emailToolShipAuthorization"
        Dim toolNum As String, shipMethod As String, moldInfoId As Long
        moldInfoId = Nz(DLookup("moldInfoId", "tblPartInfo", "partNumber = '" & rsStep!partNumber & "'"))
        If moldInfoId = 0 Then errorText = "Need a tool associated with this part to close this step."
        If errorText <> "" Then GoTo errorOut
        
        Set rsMoldInfo = db.OpenRecordset("select * from tblPartMoldingInfo where ID = " & moldInfoId)
        
        If IsNull(rsMoldInfo!toolNumber) Then errorText = "Need a tool associated with this part to send tool ship email!"
        If IsNull(rsMoldInfo!shipMethod) Then errorText = "Need to select ship method in molding info before closing this step!"
        
        If errorText <> "" Then GoTo errorOut
        
        toolNum = rsMoldInfo!toolNumber
        shipMethod = DLookup("shipMethod", "tblDropDownsSP", "ID = " & rsMoldInfo!shipMethod)
        
        Call toolShipAuthorizationEmail(toolNum, rsStep!ID, shipMethod, rsStep!partNumber)
    Case "PVtestPlanCreated"
        If DCount("ID", "tblPartTesting", "partNumber = '" & rsStep!partNumber & "'") = 0 Then 'are there any tests added?
            errorText = "Tests need added to the testing tracker for this part!"
            GoTo errorOut
        End If
    Case "PVtestPlanCompleted"
        If DCount("ID", "tblPartTesting", "partNumber = '" & rsStep!partNumber & "'") = 0 Then 'are there any tests added?
            errorText = "Tests need added to the Testing Tracker for this part!"
            GoTo errorOut
        End If
        If DCount("ID", "tblPartTesting", "partNumber = '" & rsStep!partNumber & "' AND actualEnd is null") > 0 Then 'are there any not yet complete?
            errorText = "All tests need to be complete in the Testing Tracker"
            GoTo errorOut
        End If
    Case "emailPartApprovalNotification"
        Call emailPartApprovalNotification(rsStep!ID, rsStep!partNumber)
    Case "closeStep"
        'these steps are closed based on Oracle values being present - this is checked on the firstTimeRun module
        'we can have it check here as well! just run the exact same module
        'this means that the ONLY way to close these steps is if Oracle shows the data properly. clicking the close button here just runs the same check on Oracle

        'for these steps - check if the project is in NCM. for NCM folks, do NOT check Oracle data.
        Dim rsPI As Recordset
        Set rsPI = db.OpenRecordset("SELECT developingLocation FROM tblPartInfo WHERE partNumber = '" & rsStep!partNumber & "'")
        If rsPI!developingLocation <> "NCM" Then
            Call scanSteps(rsStep!partNumber, "firstTimeRun")
            Call snackBox("info", "FYI", "This step is automatically closed when specific data is present. Clicking 'Close' ran this check manually", frmActive)
            GoTo exit_handler: 'keep stepActionChecks FALSE so it doesn't re-close the step if it was closed in the scanSteps area.
        End If
        rsPI.Close
        Set rsPI = Nothing
    Case "emailApprovedCapitalPacket"
        'check for capital packet number
        Dim CapNum As String
        CapNum = Nz(DLookup("projectCapitalNumber", "tblPartProject", "ID = " & rsStep!partProjectId), "")
        If CapNum = "" Then
            errorText = "Please enter a Capital Packet Number"
            GoTo errorOut
        End If
        If emailApprovedCapitalPacket(rsStep!ID, rsStep!partNumber, CapNum) = False Then
            errorText = "Couldn't send email, double-check the attachments"
            GoTo errorOut
        End If
End Select

stepActionOK:

Dim currentDate
currentDate = Now()

Call registerCPCUpdates("tblCPC_Steps", rsStep!ID, "closeDate", "", currentDate, rsStep!projectId, rsStep!stepName)
Call registerCPCUpdates("tblCPC_Steps", rsStep!ID, "status", rsStep!status, "Closed", rsStep!projectId, rsStep!stepName)

rsStep.Edit
rsStep!closeDate = currentDate
rsStep!status = "Closed"
rsStep.Update

Call notifyCPC(rsStep!projectId, "Closed", rsStep!stepName)

closeCPCstep = True

exit_handler:
On Error Resume Next
rsStepAction.Close
Set rsStepAction = Nothing
rsMoldInfo.Close
Set rsMoldInfo = Nothing
rsPI.Close
Set rsPI = Nothing
rsStep.Close
Set rsStep = Nothing
Set db = Nothing

Exit Function

errorOut:
Call snackBox("error", "Darn", errorText, frmActive)

Exit Function
err_handler:
    Call handleError("wdbCPCfunctions", "closeCPCstep", Err.DESCRIPTION, Err.number)
End Function

Public Function getApprovalsCompleteCPC(stepId As Long) As Long
On Error GoTo err_handler

getApprovalsCompleteCPC = 0
Dim db As Database
Set db = CurrentDb()
Dim rs1 As Recordset
Set rs1 = db.OpenRecordset("SELECT count(approvedOn) as appCount from tblCPC_StepApprovals WHERE [stepId] = " & stepId)

getApprovalsCompleteCPC = Nz(rs1!appCount, 0)

rs1.Close
Set rs1 = Nothing
Set db = Nothing

err_handler:
End Function

Public Function getTotalApprovalsCPC(stepId As Long) As Long
On Error GoTo err_handler

getTotalApprovalsCPC = 0
Dim db As Database
Set db = CurrentDb()
Dim rs1 As Recordset
Set rs1 = db.OpenRecordset("SELECT count(ID) as appCount from tblCPC_StepApprovals WHERE [stepId] = " & stepId)

getTotalApprovalsCPC = Nz(rs1!appCount, 0)

rs1.Close
Set rs1 = Nothing
Set db = Nothing

err_handler:
End Function

Function findDeptCPC(projId As Long, dept As String, Optional returnMe As Boolean = False, Optional returnFullName As Boolean = False) As String
On Error GoTo err_handler

findDeptCPC = ""

Dim db As Database
Set db = CurrentDb()
Dim rsPermissions As Recordset, permEm
Dim primaryProjId As Long
Dim primaryProjPN As String

Set rsPermissions = db.OpenRecordset("SELECT user, firstName, lastName from tblPermissions where Dept = '" & dept & "' AND Level = 'Engineer' AND user IN " & _
                                    "(SELECT memberName as user FROM tblCPC_XFTeams WHERE projectId = " & projId & ")")

Do While Not rsPermissions.EOF
    If rsPermissions!User = Environ("username") And Not returnMe Then GoTo nextRec
    If returnFullName Then
        findDeptCPC = findDeptCPC & rsPermissions!firstName & " " & rsPermissions!lastName & ","
    Else
        findDeptCPC = findDeptCPC & rsPermissions!User & ","
    End If
nextRec:
    rsPermissions.MoveNext
Loop
If findDeptCPC <> "" Then findDeptCPC = Left(findDeptCPC, Len(findDeptCPC) - 1)

rsPermissions.Close
Set rsPermissions = Nothing
Set db = Nothing

Exit Function
err_handler:
    Call handleError("wdbCPCfunctions", "findDeptCPC", Err.DESCRIPTION, Err.number)
End Function

Public Sub registerCPCUpdates(table As String, ID As Variant, column As String, oldVal As Variant, newVal As Variant, projectId As Long, Optional tag0 As String = "", Optional tag1 As String = "")
On Error GoTo err_handler

If (VarType(oldVal) = vbDate) Then oldVal = Format(oldVal, "mm/dd/yyyy")
If (VarType(newVal) = vbDate) Then newVal = Format(newVal, "mm/dd/yyyy")

Dim db As Database
Set db = CurrentDb()
Dim rs1 As Recordset
Set rs1 = db.OpenRecordset("tblCPC_UpdateTracking")

If Len(oldVal) > 255 Then oldVal = Left(oldVal, 255)
If Len(newVal) > 255 Then newVal = Left(newVal, 255)
If Len(tag0) > 100 Then newVal = Left(tag0, 100)
If Len(tag1) > 100 Then newVal = Left(tag1, 100)
If ID = "" Then ID = Null

With rs1
    .addNew
        !tableName = table
        !tableRecordId = ID
        !updatedBy = Environ("username")
        !updatedDate = Now()
        !columnName = column
        !previousData = StrQuoteReplace(CStr(Nz(oldVal, "")))
        !newData = StrQuoteReplace(CStr(Nz(newVal, "")))
        !projectId = projectId
        !dataTag0 = StrQuoteReplace(tag0)
        !dataTag1 = StrQuoteReplace(tag1)
    .Update
    .Bookmark = .lastModified
End With

rs1.Close
Set rs1 = Nothing
Set db = Nothing

Exit Sub
err_handler:
    Call handleError("wdbCPCfunctions", "registerCPCUpdates", Err.DESCRIPTION, Err.number)
End Sub

Function getYear(projectNumber As String)
On Error GoTo err_handler

    If Len(projectNumber) = 7 Then
        getYear = Left(Year(Now), 2) & Mid(projectNumber, 2, 2)
    Else
        getYear = Left(Year(Now), 2) & Mid(projectNumber, 3, 2)
    End If
    
Exit Function
err_handler:
    Call handleError("wdbCPCfunctions", "getYear", Err.DESCRIPTION, Err.number)
End Function