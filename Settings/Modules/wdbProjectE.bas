Option Compare Database
Option Explicit

Dim XL As Excel.Application, WB As Excel.Workbook, WKS As Excel.Worksheet
Dim inV As Long

Public Function grabHistoryRef(dataValue As Variant, columnName As String) As String
On Error GoTo Err_Handler

grabHistoryRef = dataValue

If dataValue = "0" Then
    grabHistoryRef = "0 / False"
    Exit Function
ElseIf dataValue = "-1" Then
    grabHistoryRef = "True"
    Exit Function
End If

dataValue = CDbl(dataValue)

Dim db As Database
Set db = CurrentDb()
Dim rs1 As Recordset
Set rs1 = db.OpenRecordset("SELECT " & columnName & " FROM tblDropDownsSP WHERE recordID = " & dataValue)

grabHistoryRef = rs1(columnName)

rs1.Close
Set rs1 = Nothing
Set db = Nothing

Err_Handler:
End Function

Public Function completelyDeletePartProjectAndInfo()
On Error GoTo Err_Handler
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
Set db = Nothing

MsgBox "All done.", vbInformation, "It is finished."

'Call registerWdbUpdates("tblPartProjects", partNum, "Part Project", partNum, "Deleted", "frmPartTrackingSettings")
Exit Function
Err_Handler:
    Call handleError("wdbProjectE", "completelyDeletePartProjectAndInfo", Err.DESCRIPTION, Err.Number)
End Function

Public Function grabTitle(User) As String
On Error GoTo Err_Handler

If IsNull(User) Then
    grabTitle = ""
    Exit Function
End If

Dim db As Database
Set db = CurrentDb()
Dim rsPermissions As Recordset
Set rsPermissions = db.OpenRecordset("SELECT * from tblPermissions where user = '" & User & "'")
grabTitle = rsPermissions!dept & " " & rsPermissions!Level

rsPermissions.Close
Set rsPermissions = Nothing
Set db = Nothing

Err_Handler:
End Function

Function notifyPE(partNum As String, notiType As String, stepTitle As String, Optional sendAlways As Boolean = False, Optional stepAction As Boolean = False, Optional notStepRelated As Boolean = False) As Boolean
On Error GoTo Err_Handler

notifyPE = False

Dim db As Database
Set db = CurrentDb()
Dim rsPartTeam As Recordset
Set rsPartTeam = db.OpenRecordset("SELECT * from tblPartTeam where partNumber = '" & partNum & "'")
If rsPartTeam.RecordCount = 0 Then Exit Function

Do While Not rsPartTeam.EOF
    Dim rsPermissions As Recordset, sendTo As String
    If IsNull(rsPartTeam!person) Then GoTo nextRec
    sendTo = rsPartTeam!person
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
    If notStepRelated Then
        subjectLine = partNum & " " & notiType '13251 Issue Created"
        emailTitle = "Issue Added" 'Internal Tooling Issue Added
        bodyTitle = stepTitle & " Issue Added"
    Else
        subjectLine = partNum & " Step " & notiType
        emailTitle = "WDB Step " & notiType
        bodyTitle = "This step has been " & notiType
    End If
    
    body = emailContentGen(subjectLine, emailTitle, bodyTitle, stepTitle & " Issue", "Part Number: " & partNum, "Who: " & closedBy, "When: " & CStr(Date))
    Call sendNotification(sendTo, 10, 2, stepTitle & " for " & partNum & " has been " & notiType, body, "Part Project", CLng(partNum))
    
nextRec:
    rsPartTeam.MoveNext
Loop

notifyPE = True

rsPartTeam.Close
Set rsPartTeam = Nothing
Set db = Nothing

Exit Function
Err_Handler:
    Call handleError("wdbProjectE", "notifyPE", Err.DESCRIPTION, Err.Number)
End Function

Function findDept(partNumber As String, dept As String, Optional returnMe As Boolean = False) As String
On Error GoTo Err_Handler

Dim db As Database
Set db = CurrentDb()
Dim rsPermissions As Recordset, permEm
Set rsPermissions = db.OpenRecordset("SELECT user, userEmail from tblPermissions where Dept = '" & dept & "' AND Level = 'Engineer' AND user IN " & _
                                    "(SELECT person FROM tblPartTeam WHERE partNumber = '" & partNumber & "')")
If rsPermissions.RecordCount = 0 Then Exit Function

Do While Not rsPermissions.EOF
    If rsPermissions!User = Environ("username") And Not returnMe Then GoTo nextRec
    findDept = findDept & rsPermissions!User & ","
nextRec:
    rsPermissions.MoveNext
Loop
If findDept <> "" Then findDept = Left(findDept, Len(findDept) - 1)

rsPermissions.Close
Set rsPermissions = Nothing
Set db = Nothing

Exit Function
Err_Handler:
    Call handleError("wdbProjectE", "findDept", Err.DESCRIPTION, Err.Number)
End Function

Public Function registerPartUpdates(table As String, ID As Variant, column As String, _
    oldVal As Variant, newVal As Variant, partNumber As String, _
    Optional tag1 As String = "", Optional tag2 As Variant = "", Optional optionExtra As String = "")
On Error GoTo Err_Handler

Dim sqlColumns As String, sqlValues As String

If (VarType(oldVal) = vbDate) Then oldVal = Format(oldVal, "mm/dd/yyyy")
If (VarType(newVal) = vbDate) Then newVal = Format(newVal, "mm/dd/yyyy")

Dim db As Database
Set db = CurrentDb()
Dim rs1 As Recordset
Set rs1 = db.OpenRecordset("tblPartUpdateTracking")

Dim updatedBy As String
updatedBy = Environ("username")
If optionExtra <> "" Then updatedBy = optionExtra

If Len(oldVal) > 255 Then oldVal = Left(oldVal, 255)
If Len(newVal) > 255 Then newVal = Left(newVal, 255)
If Len(tag1) > 100 Then newVal = Left(tag1, 100)
If Len(tag2) > 100 Then newVal = Left(tag2, 100)

With rs1
    .addNew
        !tableName = table
        !tableRecordId = ID
        !updatedBy = updatedBy
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
Set db = Nothing

Exit Function
Err_Handler:
    Call handleError("wdbProjectE", "registerPartUpdates", Err.DESCRIPTION, Err.Number)
End Function

Function generateEmailWarray(Title As String, subTitle As String, primaryMessage As String, detailTitle As String, arr() As Variant) As String
On Error GoTo Err_Handler

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
Err_Handler:
    Call handleError("wdbProjectE", "generateEmailWarray", Err.DESCRIPTION, Err.Number)
End Function