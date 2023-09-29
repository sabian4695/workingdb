Option Compare Database
Option Explicit

Public Function setProgressBarPROJECT()
Dim percent As Double, width As Long
width = 18774

Dim rsSteps As Recordset
Set rsSteps = CurrentDb().OpenRecordset("SELECT * from tblPartSteps WHERE partProjectId = " & Form_frmPartDashboard.recordID)

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

End Function

Public Function setProgressBarSTEPS(gateId As Long)
Dim percent As Double, width As Long
width = 12906

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

End Function

Function setBarColorPercent(percent As Double, controlName As String, barWidth As Long)

If percent < 0.1 Then
    Form_frmPartDashboard.Controls(controlName).width = 1
Else
    Form_frmPartDashboard.Controls(controlName).width = percent * barWidth
End If

Dim pColor
Select Case True
    Case percent < 0.25
        pColor = RGB(210, 110, 90)
    Case percent >= 0.25 And percent < 0.5
        pColor = RGB(225, 170, 70)
    Case percent >= 0.5 And percent < 0.75
        pColor = RGB(200, 210, 100)
    Case percent >= 0.75
        pColor = RGB(125, 215, 100)
End Select
Form_frmPartDashboard.Controls(controlName).BackColor = pColor
Form_frmPartDashboard.Controls(controlName & "_").BorderColor = pColor
Form_frmPartDashboard.Controls(controlName).BorderColor = pColor

End Function

Function notifyPE(partNum As String, notiType As String, stepTitle As String) As Boolean
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
    body = generateHTML("WDB Step Closed", "This step has been " & notiType, stepTitle, "Part Number: " & partNum, "Closed by: " & getFullName(), "Closed On: " & CStr(Date), False, True)
    Call sendNotification(SendTo, 10, 2, stepTitle & " for " & partNum & " has been " & notiType, "Part Project", CLng(partNum))
    Call sendSMTP(rsPermissions!userEmail, "WDB Step " & notiType, body)
    
nextRec:
    rsPartTeam.MoveNext
Loop

notifyPE = True
End Function

Function scanSteps(partNum As String, routineName As String) As Boolean
scanSteps = False

'this scans to see if there is a step action that needs to be deleted per its own requirements

Dim rsSteps As Recordset, rsStepActions As Recordset
Set rsSteps = CurrentDb().OpenRecordset("SELECT recordId, stepActionId, stepType FROM tblPartSteps WHERE partNumber = '" & partNum & "' AND stepActionId <> 0")

If rsSteps.RecordCount = 0 Then Exit Function 'no steps have actions attached!

Do While Not rsSteps.EOF
    Set rsStepActions = CurrentDb().OpenRecordset("SELECT * FROM tblPartStepActions WHERE recordId = " & rsSteps!stepActionId)
    If rsStepActions!whenToRun = routineName Then 'this is the one!
        Dim matches, rsLookItUp As Recordset
        Set rsLookItUp = CurrentDb().OpenRecordset("SELECT " & rsStepActions!compareColumn & " FROM " & rsStepActions!compareTable & " WHERE partNumber = '" & partNum & "'")
        matches = CStr(rsLookItUp(rsStepActions!compareColumn)) = rsStepActions!compareData
        If matches <> rsStepActions!compareAction Then GoTo nextOne 'if it matches what it's supposed to be, then keep going
        
        If rsStepActions!stepAction = "deleteStep" Then
            Call registerPartUpdates("tblPartSteps", rsSteps!recordID, "Deleted - stepAction", rsSteps!stepType, "", partNum, rsSteps!stepType)
            CurrentDb().Execute "DELETE FROM tblPartSteps WHERE recordId = " & rsSteps!recordID
            If CurrentProject.AllForms("frmPartDashboard").IsLoaded Then Form_sfrmPartDashboard.Requery
        End If
    End If

nextOne:
    rsSteps.MoveNext
Loop

scanSteps = True
End Function

Function iHaveOpenApproval(stepId As Long)
iHaveOpenApproval = False

Dim rsPermissions As Recordset, rsApprovals As Recordset
Set rsPermissions = CurrentDb().OpenRecordset("SELECT * from tblPermissions where user = '" & Environ("username") & "'")
Set rsApprovals = CurrentDb().OpenRecordset("SELECT * from tblPartTrackingApprovals WHERE approvedOn is null AND tableName = 'tblPartSteps' AND tableRecordId = " & stepId & " AND ((dept = '" & rsPermissions!Dept & "' AND reqLevel = '" & rsPermissions!Level & "') OR approver = '" & Environ("username") & "')")

If rsApprovals.RecordCount > 0 Then iHaveOpenApproval = True

End Function

Function iAmApprover(approvalId As Long) As Boolean
iAmApprover = False

Dim rsPermissions As Recordset, rsApprovals As Recordset
Set rsPermissions = CurrentDb().OpenRecordset("SELECT * from tblPermissions where user = '" & Environ("username") & "'")
Set rsApprovals = CurrentDb().OpenRecordset("SELECT * from tblPartTrackingApprovals WHERE approvedOn is null AND recordId = " & approvalId & " AND ((dept = '" & rsPermissions!Dept & "' AND reqLevel = '" & rsPermissions!Level & "') OR approver = '" & Environ("username") & "')")

If rsApprovals.RecordCount > 0 Then iAmApprover = True

End Function

Function issueCount(partNum As String) As Long

issueCount = DCount("recordId", "tblPartIssues", "partNumber = '" & partNum & "' AND [closeDate] is null")

End Function

Function partProjectFolder(partNum As String) As String

Dim thousZeros, hundZeros, mainPath, fullFilePath

thousZeros = Left(partNum, 2) & "000\"
hundZeros = Left(partNum, 3) & "00\"
mainPath = mainFolder("partTracking")
fullFilePath = mainPath & thousZeros & hundZeros & partNum & "\"

If FolderExists(fullFilePath) Then
    partProjectFolder = fullFilePath
Else
'check each level!!
    If FolderExists(mainPath & thousZeros) = False Then MkDir (mainPath & thousZeros)
    If FolderExists(mainPath & thousZeros & hundZeros) = False Then MkDir (mainPath & thousZeros & hundZeros)
    MkDir (fullFilePath)
    partProjectFolder = fullFilePath
End If

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
err_handler:
End Function

Public Sub registerPartUpdates(table As String, ID As Variant, column As String, oldVal As Variant, newVal As Variant, partNumber As String, Optional tag1 As String = "", Optional tag2 As Variant = "")
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

End Sub

Function toolShipAuthorizationEmail(toolNumber As String, stepId As Long, shipMethod As String, partNumber As String) As Boolean
toolShipAuthorizationEmail = False

Dim rsApprovals As Recordset
Set rsApprovals = CurrentDb().OpenRecordset("Select * from tblPartTrackingApprovals WHERE tableName = 'tblPartSteps' AND tableRecordId = " & stepId)

Dim arr() As Variant, i As Long
i = 0
rsApprovals.MoveLast
ReDim Preserve arr(rsApprovals.RecordCount)
rsApprovals.MoveFirst

Do While Not rsApprovals.EOF
    arr(i) = rsApprovals!Approver & " - " & rsApprovals!approvedOn
    i = i + 1
    rsApprovals.MoveNext
Loop

Dim toolEmail As String, subjectLine As String
toolEmail = generateEmailWarray("Tool Ship Authorization", toolNumber & " has been approved to ship", "Ship Method: " & shipMethod, "Approvals: ", arr)
subjectLine = "Tool Ship Authorization"

Dim rs2 As Recordset, strTo As String
Set rs2 = CurrentDb.OpenRecordset("SELECT * FROM tblPartTeam WHERE partNumber = '" & partNumber & "'", dbOpenSnapshot)
strTo = ""

Do While Not rs2.EOF
    If rs2!person <> Environ("username") Then strTo = strTo & "{""email"":""" & getEmail(rs2!person) & """},"
    rs2.MoveNext
Loop

strTo = Left(strTo, Len(strTo) - 1)

If sendSMTP("", subjectLine, toolEmail, "", strTo) = False Then Exit Function

toolShipAuthorizationEmail = True

End Function

Function generateEmailWarray(Title As String, subTitle As String, primaryMessage As String, detailTitle As String, arr() As Variant) As String

Dim tblHeading As String, tblFooter As String, strHTMLBody As String, extraFooter As String, detailTable As String

Dim item, i
i = 0
detailTable = ""
For Each item In arr()
    If i = UBound(arr) Then
        detailTable = detailTable & "<tr style=""border-collapse: collapse;""><td style=""padding: .1em 2em 1em 2em;"">" & item & "</td></tr>"
    Else
        detailTable = detailTable & "<tr style=""border-collapse: collapse;""><td style=""padding: .1em 2em;"">" & item & "</td></tr>"
    End If
    i = i + 1
Next item

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

End Function