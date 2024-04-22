Option Compare Database
Option Explicit

Public Function getAttachmentsCount(stepId As Long) As Long
On Error GoTo err_handler

getAttachmentsCount = 0
Dim rs1 As Recordset
Set rs1 = CurrentDb().OpenRecordset("SELECT count(ID) as countIt from tblPartAttachmentsSP WHERE [partStepId] = " & stepId)

getAttachmentsCount = Nz(rs1!countIt, 0)

rs1.Close
Set rs1 = Nothing

err_handler:
End Function

Function trialScheduleEmail(Title As String, data() As String, columns, rows) As String

Dim tblHeading As String, tblArraySection As String, strHTMLBody As String

tblHeading = "<table style=""width: 100%; margin: 0 auto; padding: .1em; text-align: center; background-color: #fafafa;"">" & _
                            "<tbody>" & _
                                "<tr><td><h2 style=""color: #414141; font-size: 28px; margin-top: 0;"">" & Title & "</h2></td></tr>" & _
                            "</tbody>" & _
                        "</table>"
                        
Dim i As Long, titleRow, dataRows, j As Long
i = 0
tblArraySection = ""

titleRow = "<tr style=""padding: .1em;"">"
For i = 0 To columns
    titleRow = titleRow & "<th>" & data(0, i) & "</th>"
Next i
titleRow = titleRow & "</tr>"

dataRows = ""
For j = 1 To rows
    dataRows = dataRows & "<tr style=""border-collapse: collapse;"">"
    For i = 0 To columns
        dataRows = dataRows & "<td style=""padding: .1em;"">" & data(j, i) & "</td>"
    Next i
    dataRows = dataRows & "</tr>"
Next j

    
tblArraySection = tblArraySection & "<table style=""width: 100%; margin: 0 auto; background: #2b2b2b; color: rgb(255,255,255);""><tbody>" & titleRow & dataRows & "</tbody></table>"

strHTMLBody = "" & _
"<!DOCTYPE html><html lang=""en"" xmlns=""http://www.w3.org/1999/xhtml"" xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:o=""urn:schemas-microsoft-com:office:office"">" & _
    "<head><meta charset=""utf-8""><title>Working DB Notification</title></head>" & _
    "<body style=""margin: 0 auto; Font-family: 'Montserrat', sans-serif; font-weight: 400; font-size: 15px; line-height: 1.8;"">" & _
        "<table style=""max-width: 600px; margin: 0 auto; text-align: center; "">" & _
            "<tbody>" & _
                "<tr><td>" & tblHeading & "</td></tr>" & _
                "<tr><td>" & tblArraySection & "</td></tr>" & _
                "<tr><td><p style=""color: rgb(192, 192, 192); text-align: center;"">This email was created by  &copy; workingDB</p></td></tr>" & _
            "</tbody>" & _
        "</table>" & _
    "</body>" & _
"</html>"

trialScheduleEmail = strHTMLBody

End Function

Public Function grabHistoryRef(dataValue As Variant, columnName As String) As String
On Error GoTo err_handler

grabHistoryRef = dataValue

If dataValue = "0" Then
    grabHistoryRef = "0 / False"
    Exit Function
ElseIf dataValue = "-1" Then
    grabHistoryRef = "True"
    Exit Function
End If

dataValue = CDbl(dataValue)

Dim rs1 As Recordset
Set rs1 = CurrentDb.OpenRecordset("SELECT " & columnName & " FROM tblDropDownsSP WHERE ID = " & dataValue)

grabHistoryRef = rs1(columnName)

err_handler:
End Function

Public Function completelyDeletePartProjectAndInfo()

'-----THIS SUB IS NOT YET USABLE

Dim db As Database, partInfoId, partNum

partNum = "26587"

Set db = CurrentDb()

'-----Part Project Data
db.Execute "delete * from tblPartProject where partNumber = '" & partNum & "'"
db.Execute "delete * from tblPartGates where partNumber = '" & partNum & "'"
db.Execute "delete * from tblPartSteps where partNumber = '" & partNum & "'"
db.Execute "delete * from tblPartTrackingApprovals where partNumber = '" & partNum & "'"
db.Execute "UPDATE tblPartAttachmentsSP SET fileStatus='deleting' where partNumber = '" & partNum & "'"

'-----Part Number based data
db.Execute "delete * from tblPartTesting where partNumber = '" & partNum & "'"
db.Execute "delete * from tblPartTeam where partNumber = '" & partNum & "'"
db.Execute "delete * from tblPartComponents where assemblyNumber = '" & partNum & "'"

'-----Part Info based data
Dim rsPartInfo As Recordset, rsPackaging As Recordset
Set rsPartInfo = db.OpenRecordset("SELECT * from tblPartInfo WHERE partNumber = '" & partNum & "'")

partInfoId = rsPartInfo!recordId
db.Execute "delete * from tblPartQuoteInfo where recordId = " & rsPartInfo!quoteInfoId
db.Execute "delete * from tblPartAssemblyInfo where recordId = " & rsPartInfo!assemblyInfoId
db.Execute "delete * from tblPartOutsourceInfo where recordId = " & rsPartInfo!outsourceInfoId

rsPartInfo.Close
Set rsPartInfo = Nothing

'-----Part Packaging and Components
Set rsPackaging = db.OpenRecordset("SELECT * from tblPartPackaging WHERE partInfoId = " & partInfoId)
Do While Not rsPackaging.EOF
    db.Execute "Delete * from tblPartPackagingComponents WHERE packagingInfoId = " & rsPackaging!recordId
    rsPackaging.MoveNext
Loop
rsPackaging.Delete
rsPackaging.Close
Set rsPackaging = Nothing

'-----Part Meetings and Attendees
Dim rsMeetings As Recordset
Set rsMeetings = db.OpenRecordset("SELECT * from tblPartMeetings where partNum = '" & partNum & "'")
Do While Not rsMeetings.EOF
    db.Execute "Delete * from tblPartMeetingAttendees WHERE meetingId = " & rsMeetings!recordId
    rsMeetings.MoveNext
Loop
rsMeetings.Close
Set rsMeetings = Nothing

'-----Part Info
db.Execute "delete * from tblPartInfo where partNumber = '" & partNum & "'"

MsgBox "All done.", vbInformation, "It is finished."

'Call registerWdbUpdates("tblPartProjects", partNum, "Part Project", partNum, "Deleted", "frmPartTrackingSettings")

End Function

Public Function getApprovalsComplete(stepId As Long, partNumber As String) As Long
On Error GoTo err_handler

getApprovalsComplete = 0
Dim rs1 As Recordset
Set rs1 = CurrentDb().OpenRecordset("SELECT count(approvedOn) as appCount from tblPartTrackingApprovals WHERE [partNumber] = '" & partNumber & "' AND [tableRecordId] = " & stepId & " AND [tableName] = 'tblPartSteps'")

getApprovalsComplete = Nz(rs1!appCount, 0)

rs1.Close
Set rs1 = Nothing

err_handler:
End Function

Public Function getTotalApprovals(stepId As Long, partNumber As String) As Long
On Error GoTo err_handler

getTotalApprovals = 0
Dim rs1 As Recordset
Set rs1 = CurrentDb().OpenRecordset("SELECT count(recordId) as appCount from tblPartTrackingApprovals WHERE [partNumber] = '" & partNumber & "' AND [tableRecordId] = " & stepId & " AND [tableName] = 'tblPartSteps'")

getTotalApprovals = Nz(rs1!appCount, 0)

rs1.Close
Set rs1 = Nothing

err_handler:
End Function

Public Function recalcStepDueDates(projId As Long, oldDueDate As Date, moveBy As Long)
On Error Resume Next

Dim rsSteps As Recordset
Set rsSteps = CurrentDb().OpenRecordset("Select dueDate from tblPartSteps Where partProjectId = " & projId & " AND dueDate > #" & oldDueDate & "#")

Do While Not rsSteps.EOF
    rsSteps.Edit
    rsSteps!dueDate = addWorkdays(rsSteps!dueDate, moveBy)
    rsSteps.Update
    rsSteps.MoveNext
Loop

rsSteps.Close
Set rsSteps = Nothing

End Function

Public Function getCurrentStepDue(projId As Long) As String
On Error Resume Next

getCurrentStepDue = ""

Dim rs1 As Recordset
Set rs1 = CurrentDb().OpenRecordset("SELECT Min(dueDate) as minDue from tblPartSteps WHERE partProjectId = " & projId & " AND status <> 'Closed'")

getCurrentStepDue = Nz(rs1!minDue, "")

rs1.Close
Set rs1 = Nothing

End Function

Public Function getDOH(partNum As String) As Long
On Error GoTo err_handler

Dim db As Database
Dim rsSupplyDemand As Recordset, rsPartInfo As Recordset, rsOnHand As Recordset
Set db = CurrentDb
Set rsSupplyDemand = db.OpenRecordset("SELECT sum([QTY_DUE]) AS openOrders FROM APPS_XXCUS_SUPPLY_DEMAND_V WHERE ITEM='" & partNum & "'")
Dim openOrders
openOrders = rsSupplyDemand!openOrders

Set rsPartInfo = db.OpenRecordset("SELECT monthlyVolume from tblPartInfo WHERE partNumber = '" & partNum & "'")

Dim monthlyVolCalc, monthlyVol
monthlyVol = rsPartInfo!monthlyVolume
If IsNull(openOrders) Or monthlyVol > openOrders Then
    monthlyVolCalc = monthlyVol / 22
Else
    monthlyVolCalc = openOrders / 22
End If


Set rsOnHand = db.OpenRecordset("SELECT sum(TRANSACTION_QUANTITY) as TransQty FROM APPS_MTL_ONHAND_QUANTITIES WHERE INVENTORY_ITEM_ID = " & idNAM(partNum, "NAM"))
getDOH = Nz(rsOnHand!TransQty, 0) / monthlyVolCalc


rsPartInfo.Close
rsSupplyDemand.Close
rsOnHand.Close
Set rsPartInfo = Nothing
Set rsSupplyDemand = Nothing
Set rsOnHand = Nothing

Exit Function
err_handler:
    Call handleError("wdbProjectE", "getDOH", Err.Description, Err.number)
End Function

Public Function openOrdersByDay(partNum As String, dayNum As Long) As Double
On Error GoTo err_handler

Dim filt As String

Select Case dayNum
    Case 1
        filt = "REQUIREMENT_DATE>=Date() AND REQUIREMENT_DATE<=Date()+1"
    Case 2
        filt = "REQUIREMENT_DATE>Date()+1 AND REQUIREMENT_DATE<=Date()+2"
    Case 3
        filt = "REQUIREMENT_DATE>Date()+2 AND REQUIREMENT_DATE<=Date()+3"
    Case 4
        filt = "REQUIREMENT_DATE>Date()+3 AND REQUIREMENT_DATE<=Date()+4"
    Case 5
        filt = "REQUIREMENT_DATE>Date()+4 AND REQUIREMENT_DATE<=Date()+5"
    Case 6
        filt = "REQUIREMENT_DATE>Date()+6 AND REQUIREMENT_DATE<=Date()+11"
    Case 0 'back orders
        filt = "REQUIREMENT_DATE<Date()"
End Select

Dim db As Database
Set db = CurrentDb
Dim rsSupplyDemand As Recordset
Set rsSupplyDemand = db.OpenRecordset("SELECT SUM([QTY_DUE]) AS QtyReq FROM APPS_XXCUS_SUPPLY_DEMAND_V WHERE ITEM = '" & partNum & "' AND " & filt, dbOpenSnapshot)
openOrdersByDay = Nz(rsSupplyDemand!QtyReq, 0)

rsSupplyDemand.Close
Set rsSupplyDemand = Nothing

Exit Function
err_handler:
    Call handleError("wdbProjectE", "openOrdersByDay", Err.Description, Err.number)
End Function

Public Function createPartProject(projId)
On Error GoTo err_handler

Dim db As DAO.Database
Set db = CurrentDb()
Dim rsProject As Recordset, rsStepTemplate As Recordset, rsApprovalsTemplate As Recordset, rsGateTemplate As Recordset
Dim strInsert As String, strInsert1 As String
Dim projTempId As Long, pNum As String, childPnum As String, runningDate As Date

Set rsProject = db.OpenRecordset("SELECT * from tblPartProject WHERE recordId = " & projId)

projTempId = rsProject!projectTemplateId
pNum = rsProject!partNumber
childPnum = Nz(rsProject!childPartNumber, "")
runningDate = rsProject!projectStartDate

db.Execute "INSERT INTO tblPartTeam(partNumber,person) VALUES ('" & pNum & "','" & Environ("username") & "')", dbFailOnError 'assign project engineer
Set rsGateTemplate = db.OpenRecordset("Select * FROM tblPartGateTemplate WHERE [projectTemplateId] = " & projTempId, dbOpenSnapshot)

'--GO THROUGH EACH GATE
Do While Not rsGateTemplate.EOF
    '--ADD THIS GATE
    db.Execute "INSERT INTO tblPartGates(projectId,partNumber,gateTitle) VALUES (" & projId & ",'" & pNum & "','" & rsGateTemplate![gateTitle] & "')", dbFailOnError
    TempVars.Add "gateId", db.OpenRecordset("SELECT @@identity")(0).Value
    
    '--ADD STEPS FOR THIS GATE
    Set rsStepTemplate = db.OpenRecordset("SELECT * from tblPartStepTemplate WHERE [gateTemplateId] = " & rsGateTemplate![recordId] & " ORDER BY indexOrder Asc", dbOpenSnapshot)
    Do While Not rsStepTemplate.EOF
        If (IsNull(rsStepTemplate![Title]) Or rsStepTemplate![Title] = "") Then GoTo nextStep
        runningDate = addWorkdays(runningDate, Nz(rsStepTemplate![duration], 1))
        strInsert = "INSERT INTO tblPartSteps" & _
            "(partNumber,partProjectId,partGateId,stepType,openedBy,status,openDate,lastUpdatedDate,lastUpdatedBy,stepActionId,documentType,responsible,indexOrder,duration,dueDate) VALUES"
        strInsert = strInsert & "('" & pNum & "'," & projId & "," & TempVars!gateId & ",'" & StrQuoteReplace(rsStepTemplate![Title]) & "','" & _
            Environ("username") & "','Not Started','" & Now() & "','" & Now() & "','" & Environ("username") & "',"
        strInsert = strInsert & Nz(rsStepTemplate![stepActionId], "NULL") & "," & Nz(rsStepTemplate![documentType], "NULL") & ",'" & _
            Nz(rsStepTemplate![responsible], "") & "'," & rsStepTemplate![indexOrder] & "," & Nz(rsStepTemplate![duration], 1) & ",'" & runningDate & "');"
        db.Execute strInsert, dbFailOnError
        
        '--ADD APPROVALS FOR THIS STEP
        If Not rsStepTemplate![approvalRequired] Then GoTo nextStep
        TempVars.Add "stepId", db.OpenRecordset("SELECT @@identity")(0).Value
        Set rsApprovalsTemplate = db.OpenRecordset("SELECT * FROM tblPartStepTemplateApprovals WHERE [stepTemplateId] = " & rsStepTemplate![recordId], dbOpenSnapshot)
        
        Do While Not rsApprovalsTemplate.EOF
            strInsert1 = "INSERT INTO tblPartTrackingApprovals(partNumber,requestedBy,requestedDate,dept,reqLevel,tableName,tableRecordId) VALUES ('" & _
                pNum & "','" & Environ("username") & "','" & Now() & "','" & _
                Nz(rsApprovalsTemplate![dept], "") & "','" & Nz(rsApprovalsTemplate![reqLevel], "") & "','tblPartSteps'," & TempVars!stepId & ");"
            CurrentDb().Execute strInsert1
            rsApprovalsTemplate.MoveNext
        Loop
nextStep:
        rsStepTemplate.MoveNext
    Loop
    db.Execute "UPDATE tblPartGates SET plannedDate = '" & runningDate & "' WHERE recordId = " & TempVars!gateId 'set the planned date as the last step due date in this gate
    rsGateTemplate.MoveNext
Loop

rsGateTemplate.Close
Set rsGateTemplate = Nothing
rsStepTemplate.Close
Set rsStepTemplate = Nothing

Exit Function
err_handler:
    Call handleError("wdbProjectE", "createPartProject", Err.Description, Err.number)
End Function

Public Function grabTitle(user) As String
On Error GoTo err_handler

If IsNull(user) Then
    grabTitle = ""
    Exit Function
End If

Dim rsPermissions As Recordset
Set rsPermissions = CurrentDb().OpenRecordset("SELECT * from tblPermissions where user = '" & user & "'")
grabTitle = rsPermissions!dept & " " & rsPermissions!Level

err_handler:
End Function

Public Function setProgressBarPROJECT()
On Error GoTo err_handler

Dim percent As Double, width As Long
width = 17820

Dim rsSteps As Recordset
Set rsSteps = CurrentDb().OpenRecordset("SELECT * from tblPartSteps WHERE partProjectId = " & Form_frmPartDashboard.recordId)

Dim totalSteps, closedSteps
rsSteps.MoveLast
totalSteps = rsSteps.RecordCount

rsSteps.filter = "status = 'Closed'"
Set rsSteps = rsSteps.OpenRecordset
If rsSteps.RecordCount = 0 Then
    percent = 0
    GoTo setBar
End If
rsSteps.MoveFirst
rsSteps.MoveLast
closedSteps = rsSteps.RecordCount
percent = closedSteps / totalSteps

setBar:
Call setBarColorPercent(percent, "progressBarPROJECT", width)

Exit Function
err_handler:
    Call handleError("wdbProjectE", "setProgressBarPROJECT", Err.Description, Err.number)
End Function

Public Function setProgressBarSTEPS(gateId As Long)
On Error GoTo err_handler

Dim percent As Double, width As Long
width = 12000

Dim rsSteps As Recordset
Set rsSteps = CurrentDb().OpenRecordset("SELECT * from tblPartSteps WHERE partGateId = " & gateId)

Dim totalSteps, closedSteps
rsSteps.MoveLast
totalSteps = rsSteps.RecordCount

rsSteps.filter = "status = 'Closed'"
Set rsSteps = rsSteps.OpenRecordset
If rsSteps.RecordCount = 0 Then
    percent = 0
    GoTo setBar
End If
rsSteps.MoveFirst
rsSteps.MoveLast
closedSteps = rsSteps.RecordCount
percent = closedSteps / totalSteps

setBar:
Call setBarColorPercent(percent, "progressBarSTEPS", width)

Exit Function
err_handler:
    Call handleError("wdbProjectE", "setProgressBarSTEPS", Err.Description, Err.number)
End Function

Function setBarColorPercent(percent As Double, controlName As String, barWidth As Long)
On Error GoTo err_handler

If percent < 0.1 Then
    Form_frmPartDashboard.Controls(controlName).width = 1
Else
    Form_frmPartDashboard.Controls(controlName).width = percent * barWidth
End If

Dim pColor
Select Case True
    Case percent < 0.25
        pColor = RGB(150, 100, 100)
    Case percent >= 0.25 And percent < 0.5
        pColor = RGB(185, 130, 100)
    Case percent >= 0.5 And percent < 0.75
        pColor = RGB(140, 150, 100)
    Case percent >= 0.75
        pColor = RGB(100, 150, 100)
End Select
Form_frmPartDashboard.Controls(controlName).BackColor = pColor
Form_frmPartDashboard.Controls(controlName & "_").BorderColor = pColor
Form_frmPartDashboard.Controls(controlName).BorderColor = pColor

Exit Function
err_handler:
    Call handleError("wdbProjectE", "setBarColorPercent", Err.Description, Err.number)
End Function

Function notifyPE(partNum As String, notiType As String, stepTitle As String) As Boolean
On Error GoTo err_handler

notifyPE = False

Dim rsPartTeam As Recordset
Set rsPartTeam = CurrentDb().OpenRecordset("SELECT * from tblPartTeam where partNumber = '" & partNum & "'")
If rsPartTeam.RecordCount = 0 Then Exit Function

Do While Not rsPartTeam.EOF
    Dim rsPermissions As Recordset, SendTo As String
    If IsNull(rsPartTeam!person) Then GoTo nextRec
    SendTo = rsPartTeam!person
    Set rsPermissions = CurrentDb().OpenRecordset("SELECT user, userEmail from tblPermissions where user = '" & SendTo & "' AND Dept = 'Project' AND Level = 'Engineer'")
    If rsPermissions.RecordCount = 0 Then GoTo nextRec
    If SendTo = Environ("username") Then GoTo nextRec
    
    'actually send notification
    Dim body As String
    body = emailContentGen(partNum & " Step " & notiType, "WDB Step " & notiType, "This step has been " & notiType, stepTitle, "Part Number: " & partNum, "Closed by: " & getFullName(), "Closed On: " & CStr(Date))
    Call sendNotification(SendTo, 10, 2, stepTitle & " for " & partNum & " has been " & notiType, body, "Part Project", CLng(partNum))
    
nextRec:
    rsPartTeam.MoveNext
Loop

notifyPE = True

Exit Function
err_handler:
    Call handleError("wdbProjectE", "notifyPE", Err.Description, Err.number)
End Function

Function scanSteps(partNum As String, routineName As String, Optional identifier As Variant = "notFound") As Boolean
On Error GoTo err_handler

scanSteps = False
'this scans to see if there is a step action that needs to be deleted per its own requirements

Dim rsSteps As Recordset, rsStepActions As Recordset
Set rsSteps = CurrentDb().OpenRecordset("SELECT recordId, stepActionId, stepType FROM tblPartSteps WHERE partNumber = '" & partNum & "' AND stepActionId <> 0")

If rsSteps.RecordCount = 0 Then Exit Function 'no steps have actions attached!

Do While Not rsSteps.EOF
    Set rsStepActions = CurrentDb().OpenRecordset("SELECT * FROM tblPartStepActions WHERE recordId = " & rsSteps!stepActionId)
    If Nz(rsStepActions!whenToRun, "") <> routineName Then GoTo nextOne 'check if this is the right time to run this actions step
    
    Dim matches, rsLookItUp As Recordset, matchingCol As String, meetsCriteria As Boolean
    matchingCol = "partNumber"
    If identifier = "notFound" Then identifier = "'" & partNum & "'"
    If routineName = "frmPartMoldingInfo_save" Then matchingCol = "recordId"
    Set rsLookItUp = CurrentDb().OpenRecordset("SELECT " & rsStepActions!compareColumn & " FROM " & rsStepActions!compareTable & " WHERE " & matchingCol & " = " & identifier)
    
    meetsCriteria = False
    
    If InStr(rsStepActions!compareData, ",") > 0 Then 'check for multiple values - always seen as an OR command, not AND
        'make an array of the values and check if any match
        Dim checkIf() As String, ITEM
        checkIf = Split(rsStepActions!compareData, ",")
        For Each ITEM In checkIf
            matches = CStr(Nz(rsLookItUp(rsStepActions!compareColumn), "")) = ITEM
            If matches Then meetsCriteria = True
        Next ITEM
    Else
        matches = CStr(Nz(rsLookItUp(rsStepActions!compareColumn), "")) = rsStepActions!compareData
        If matches Then meetsCriteria = True
    End If
    
    If meetsCriteria <> rsStepActions!compareAction Then GoTo nextOne
    
    If rsStepActions!stepAction = "deleteStep" Then
        Call registerPartUpdates("tblPartSteps", rsSteps!recordId, "Deleted - stepAction", rsSteps!stepType, "", partNum, rsSteps!stepType)
        CurrentDb().Execute "DELETE FROM tblPartSteps WHERE recordId = " & rsSteps!recordId
        If CurrentProject.AllForms("frmPartDashboard").IsLoaded Then Form_sfrmPartDashboard.Requery
    End If

nextOne:
    rsSteps.MoveNext
Loop

scanSteps = True

Exit Function
err_handler:
    Call handleError("wdbProjectE", "scanSteps", Err.Description, Err.number)
End Function

Function iHaveOpenApproval(stepId As Long)
On Error GoTo err_handler

iHaveOpenApproval = False

Dim rsPermissions As Recordset, rsApprovals As Recordset
Set rsPermissions = CurrentDb().OpenRecordset("SELECT * from tblPermissions where user = '" & Environ("username") & "'")
Set rsApprovals = CurrentDb().OpenRecordset("SELECT * from tblPartTrackingApprovals WHERE approvedOn is null AND tableName = 'tblPartSteps' AND tableRecordId = " & stepId & " AND ((dept = '" & rsPermissions!dept & "' AND reqLevel = '" & rsPermissions!Level & "') OR approver = '" & Environ("username") & "')")

If rsApprovals.RecordCount > 0 Then iHaveOpenApproval = True

rsPermissions.Close
Set rsPermissions = Nothing
rsApprovals.Close
Set rsApprovals = Nothing

Exit Function
err_handler:
    Call handleError("wdbProjectE", "iHaveOpenApproval", Err.Description, Err.number)
End Function

Function iAmApprover(approvalId As Long) As Boolean
On Error GoTo err_handler

iAmApprover = False

Dim rsPermissions As Recordset, rsApprovals As Recordset
Set rsPermissions = CurrentDb().OpenRecordset("SELECT * from tblPermissions where user = '" & Environ("username") & "'")
Set rsApprovals = CurrentDb().OpenRecordset("SELECT * from tblPartTrackingApprovals WHERE approvedOn is null AND recordId = " & approvalId & " AND ((dept = '" & rsPermissions!dept & "' AND reqLevel = '" & rsPermissions!Level & "') OR approver = '" & Environ("username") & "')")

If rsApprovals.RecordCount > 0 Then iAmApprover = True

rsPermissions.Close
Set rsPermissions = Nothing
rsApprovals.Close
Set rsApprovals = Nothing

Exit Function
err_handler:
    Call handleError("wdbProjectE", "iAmApprover", Err.Description, Err.number)
End Function

Function issueCount(partNum As String) As Long
On Error GoTo err_handler

issueCount = DCount("recordId", "tblPartIssues", "partNumber = '" & partNum & "' AND [closeDate] is null")

Exit Function
err_handler:
    Call handleError("wdbProjectE", "issueCount", Err.Description, Err.number)
End Function

Function emailPartInfo(partNum As String, noteTxt As String) As Boolean
On Error GoTo err_handler
emailPartInfo = False

Dim SendItems As New clsOutlookCreateItem               ' outlook class
    Dim strTo As String                                     ' email recipient
    Dim strSubject As String                                ' email subject
    
    Set SendItems = New clsOutlookCreateItem

    Dim rs2 As Recordset
    Set rs2 = CurrentDb.OpenRecordset("SELECT * FROM tblPartTeam WHERE partNumber = '" & partNum & "'", dbOpenSnapshot)
    strTo = ""

    Do While Not rs2.EOF
        If rs2!person <> Environ("username") Then strTo = strTo & getEmail(rs2!person) & "; "
        rs2.MoveNext
    Loop
    
    strSubject = partNum & " Sales Kickoff Meeting"
    
    Dim z As String, tempFold As String
    tempFold = "\\data\mdbdata\WorkingDB\_docs\Temp\" & Environ("username") & "\"
    If FolderExists(tempFold) = False Then MkDir (tempFold)
    z = tempFold & Format(Date, "YYMMDD") & "_" & partNum & "_Part_Information.pdf"
    DoCmd.OpenReport "rptPartInformation", acViewPreview, , "[partNumber]='" & partNum & "'", acHidden
    DoCmd.OutputTo acOutputReport, "rptPartInformation", acFormatPDF, z, False
    DoCmd.Close acReport, "rptPartInformation"
    
    SendItems.CreateMailItem SendTo:=strTo, _
                             subject:=strSubject, _
                             Attachments:=z
    Set SendItems = Nothing
    
Dim FSO
Set FSO = CreateObject("Scripting.FileSystemObject")
Call FSO.deleteFile(z)
    
emailPartInfo = True

Exit Function
err_handler:
    Call handleError("wdbProjectE", "emailPartInfo", Err.Description, Err.number)
End Function

Public Function registerPartUpdates(table As String, ID As Variant, column As String, oldVal As Variant, newVal As Variant, partNumber As String, Optional tag1 As String = "", Optional tag2 As Variant = "")
On Error GoTo err_handler

Dim sqlColumns As String, sqlValues As String

If (VarType(oldVal) = vbDate) Then oldVal = Format(oldVal, "mm/dd/yyyy")
If (VarType(newVal) = vbDate) Then newVal = Format(newVal, "mm/dd/yyyy")

Dim rs1 As Recordset
Set rs1 = CurrentDb().OpenRecordset("tblPartUpdateTracking")

With rs1
    .addNew
        !tableName = table
        !tableRecordId = ID
        !updatedBy = Environ("username")
        !updatedDate = Now()
        !columnName = column
        !previousData = StrQuoteReplace(CStr(Nz(oldVal, "")))
        !newData = StrQuoteReplace(CStr(Nz(newVal, "")))
        !partNumber = partNumber
        !dataTag1 = StrQuoteReplace(tag1)
        !dataTag2 = StrQuoteReplace(tag2)
    .Update
    .Bookmark = .lastModified
End With

rs1.Close
Set rs1 = Nothing

Exit Function
err_handler:
    Call handleError("wdbProjectE", "registerPartUpdates", Err.Description, Err.number)
End Function

Function toolShipAuthorizationEmail(toolNumber As String, stepId As Long, shipMethod As String, partNumber As String) As Boolean
On Error GoTo err_handler

toolShipAuthorizationEmail = False

Dim rsApprovals As Recordset
Set rsApprovals = CurrentDb().OpenRecordset("Select * from tblPartTrackingApprovals WHERE tableName = 'tblPartSteps' AND tableRecordId = " & stepId)

Dim approvalsBool
approvalsBool = True
If rsApprovals.RecordCount = 0 Then
    approvalsBool = False
    GoTo noApprovals
End If

Dim arr() As Variant, i As Long
i = 0
rsApprovals.MoveLast
ReDim Preserve arr(rsApprovals.RecordCount)
rsApprovals.MoveFirst

Do While Not rsApprovals.EOF
    arr(i) = rsApprovals!approver & " - " & rsApprovals!approvedOn
    i = i + 1
    rsApprovals.MoveNext
Loop

noApprovals:
Dim toolEmail As String, subjectLine As String
subjectLine = "Tool Ship Authorization"
If approvalsBool Then
    toolEmail = generateEmailWarray("Tool Ship Authorization", toolNumber & " has been approved to ship", "Ship Method: " & shipMethod, "Approvals: ", arr)
Else
    toolEmail = generateHTML("Tool Ship Authorization", toolNumber & " has been approved to ship", "Ship Method: " & shipMethod, "Approvals: none", "", "", False)
End If

Dim rs2 As Recordset, strTo As String
Set rs2 = CurrentDb.OpenRecordset("SELECT * FROM tblPartTeam WHERE partNumber = '" & partNumber & "'", dbOpenSnapshot)
strTo = ""

Do While Not rs2.EOF
    If rs2!person <> Environ("username") Then strTo = strTo & getEmail(rs2!person) & "; "
    rs2.MoveNext
Loop

Dim SendItems As New clsOutlookCreateItem
Set SendItems = New clsOutlookCreateItem

SendItems.CreateMailItem SendTo:=strTo, _
                             subject:=subjectLine, _
                             htmlBody:=toolEmail
    Set SendItems = Nothing

toolShipAuthorizationEmail = True

Exit Function
err_handler:
    Call handleError("wdbProjectE", "toolShipAuthorizationEmail", Err.Description, Err.number)
End Function

Function emailPartApprovalNotification(stepId As Long, partNumber As String) As Boolean
On Error GoTo err_handler

emailPartApprovalNotification = False

Dim emailBody As String, subjectLine As String
subjectLine = "Part Approval Notification"
emailBody = generateHTML(subjectLine, partNumber & " has received customer approval", "Part Approved", "No extra details...", "", "", False)

Dim rs2 As Recordset, strTo As String
Set rs2 = CurrentDb.OpenRecordset("SELECT * FROM tblPartTeam WHERE partNumber = '" & partNumber & "'", dbOpenSnapshot)
strTo = ""

Do While Not rs2.EOF
    If rs2!person <> Environ("username") Then strTo = strTo & getEmail(rs2!person) & "; "
    rs2.MoveNext
Loop

Dim SendItems As New clsOutlookCreateItem
Set SendItems = New clsOutlookCreateItem

SendItems.CreateMailItem SendTo:=strTo, _
                             subject:=subjectLine, _
                             htmlBody:=emailBody
    Set SendItems = Nothing

emailPartApprovalNotification = True

Exit Function
err_handler:
    Call handleError("wdbProjectE", "emailPartApprovalNotification", Err.Description, Err.number)
End Function

Function generateEmailWarray(Title As String, subTitle As String, primaryMessage As String, detailTitle As String, arr() As Variant) As String
On Error GoTo err_handler

Dim tblHeading As String, tblFooter As String, strHTMLBody As String, extraFooter As String, detailTable As String

Dim ITEM, i
i = 0
detailTable = ""
For Each ITEM In arr()
    If i = UBound(arr) Then
        detailTable = detailTable & "<tr style=""border-collapse: collapse;""><td style=""padding: .1em 2em 1em 2em;"">" & ITEM & "</td></tr>"
    Else
        detailTable = detailTable & "<tr style=""border-collapse: collapse;""><td style=""padding: .1em 2em;"">" & ITEM & "</td></tr>"
    End If
    i = i + 1
Next ITEM

tblHeading = "<table style=""width: 100%; margin: 0 auto; padding: 2em 3em; text-align: center; background-color: #fafafa;"">" & _
                            "<tbody>" & _
                                "<tr><td><h2 style=""color: #414141; font-size: 28px; margin-top: 0;"">" & Title & "</h2></td></tr>" & _
                                "<tr><td><p style=""color: rgb(73, 73, 73);"">" & subTitle & "</p></td></tr>" & _
                                 "<tr><td><table style=""padding: 1em; text-align: center;"">" & _
                                     "<tr><td style=""padding: 1em 1.5em; background: #FF6B00; "">" & primaryMessage & "</td></tr>" & _
                                "</table></td></tr>" & _
                            "</tbody>" & _
                        "</table>"
                        
tblFooter = "<table style=""width: 100%; margin: 0 auto; background: #2b2b2b; color: rgba(255,255,255,.5);"">" & _
                        "<tbody>" & _
                            "<tr style=""border-collapse: collapse;""><td style=""padding: 1em; color: #c9c9c9;"">" & detailTitle & "</td></tr>" & _
                            detailTable & _
                        "</tbody>" & _
                    "</table>"
                    
extraFooter = "<tr><td><p style=""color: rgb(192, 192, 192); text-align: center;"">This email address is not monitored, please do not reply to this email</p></td></tr>"

strHTMLBody = "" & _
"<!DOCTYPE html><html lang=""en"" xmlns=""http://www.w3.org/1999/xhtml"" xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:o=""urn:schemas-microsoft-com:office:office"">" & _
    "<head><meta charset=""utf-8""><title>Working DB Notification</title></head>" & _
    "<body style=""margin: 0 auto; Font-family: 'Montserrat', sans-serif; font-weight: 400; font-size: 15px; line-height: 1.8;"">" & _
        "<table style=""max-width: 600px; margin: 0 auto; text-align: center; "">" & _
            "<tbody>" & _
                "<tr><td>" & tblHeading & "</td></tr>" & _
                "<tr><td>" & tblFooter & "</td></tr>" & _
                extraFooter & _
                "<tr><td><p style=""color: rgb(192, 192, 192); text-align: center;"">This email was created by  &copy; workingDB</p></td></tr>" & _
            "</tbody>" & _
        "</table>" & _
    "</body>" & _
"</html>"

generateEmailWarray = strHTMLBody

Exit Function
err_handler:
    Call handleError("wdbProjectE", "generateEmailWarray", Err.Description, Err.number)
End Function