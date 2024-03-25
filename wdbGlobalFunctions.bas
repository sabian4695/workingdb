Option Compare Database
Option Explicit

Public bClone As Boolean

Public Function registerWdbUpdates(table As String, ID As Variant, column As String, oldVal As Variant, newVal As Variant, Optional tag0 As String = "", Optional tag1 As Variant = "")
On Error GoTo err_handler

Dim sqlColumns As String, sqlValues As String

If (VarType(oldVal) = vbDate) Then oldVal = Format(oldVal, "mm/dd/yyyy")
If (VarType(newVal) = vbDate) Then newVal = Format(newVal, "mm/dd/yyyy")

Dim rs1 As Recordset
Set rs1 = CurrentDb().OpenRecordset("tblWdbUpdateTracking")

With rs1
    .addNew
        !tableName = table
        !tableRecordId = ID
        !updatedBy = Environ("username")
        !updatedDate = Now()
        !columnName = column
        !previousData = StrQuoteReplace(CStr(Nz(oldVal, "")))
        !newData = StrQuoteReplace(CStr(Nz(newVal, "")))
        !dataTag0 = StrQuoteReplace(tag0)
        !dataTag1 = StrQuoteReplace(tag1)
    .Update
    .Bookmark = .lastModified
End With

rs1.Close
Set rs1 = Nothing

Exit Function
err_handler:
    Call handleError("wdbGlocalFunctions", "registerWdbUpdates", Err.description, Err.number)
End Function

Public Function addWorkdays(dateInput As Date, daysToAdd As Long) As Date
On Error GoTo err_handler

Dim i As Long, testDate As Date, daysLeft As Long, rsHolidays As Recordset, intDirection
testDate = dateInput
daysLeft = Abs(daysToAdd)
intDirection = 1
If daysToAdd < 0 Then intDirection = -1

Do While daysLeft > 0
    testDate = testDate + intDirection
    If Weekday(testDate) = 7 Or Weekday(testDate) = 1 Then ' IF WEEKEND -> skip
        testDate = testDate + intDirection
        GoTo skipDate
    End If
    Set rsHolidays = CurrentDb().OpenRecordset("SELECT * from tblHolidays WHERE holidayDate = #" & testDate & "#")
    If rsHolidays.RecordCount > 0 Then GoTo skipDate ' IF HOLIDAY -> skip to next day
     daysLeft = daysLeft - 1
skipDate:
Loop

addWorkdays = testDate

Exit Function
err_handler:
    Call handleError("wdbGlobalFunctions", "addWorkdays", Err.description, Err.number)
End Function

Public Function countWorkdays(oldDate As Date, newDate As Date) As Long
On Error GoTo err_handler

Dim total, sunday, saturday, weekdays, holidays

total = DateDiff("d", [oldDate], [newDate], vbSunday)
sunday = DateDiff("ww", [oldDate], [newDate], 1)
saturday = DateDiff("ww", [oldDate], [newDate], 7)
holidays = DCount("recordId", "tblHolidays", "holidayDate > #" & oldDate - 1 & "# AND holidayDate < #" & newDate & "#")
countWorkdays = total - sunday - saturday - holidays

Exit Function
err_handler:
    Call handleError("wdbGlobalFunctions", "countWorkdays", Err.description, Err.number)
End Function

Function getFullName() As String

Dim rs1 As Recordset
Set rs1 = CurrentDb().OpenRecordset("SELECT firstName, lastName FROM tblPermissions WHERE User = '" & Environ("username") & "'", dbOpenSnapshot)
getFullName = rs1!firstName & " " & rs1!lastName
rs1.Close: Set rs1 = Nothing

End Function

Function notificationsCount()

Dim unRead
unRead = DCount("ID", "tblNotificationsSP", "recipientUser = '" & Environ("username") & "' AND readDate is null")
If unRead > 9 Then
    Form_DASHBOARD.Form.notifications.Caption = "9+"
Else
    Form_DASHBOARD.Form.notifications.Caption = CStr(unRead)
End If
If unRead = 0 Then
    Form_DASHBOARD.Form.notifications.BackColor = RGB(60, 170, 60)
Else
    Form_DASHBOARD.Form.notifications.BackColor = RGB(230, 0, 0)
End If

End Function

Function loadECOtype(changeNotice As String) As String

Dim rs1 As Recordset
Set rs1 = CurrentDb().OpenRecordset("SELECT [CHANGE_ORDER_TYPE_ID] from ENG_ENG_ENGINEERING_CHANGES where [CHANGE_NOTICE] = '" & changeNotice & "'", dbOpenSnapshot)

loadECOtype = DLookup("[ECO_Type]", "[tblOracleDropDowns]", "[ECO_Type_ID]=" & rs1!CHANGE_ORDER_TYPE_ID)

rs1.Close
Set rs1 = Nothing
End Function

Function randomNumber(low As Long, high As Long) As Long
Randomize
randomNumber = Int((high - low + 1) * Rnd() + low)
End Function

Function getAPI(url, header1, header2)
Dim reader As New XMLHTTP60
    reader.open "GET", url, False
    reader.setRequestHeader header1, header2
    reader.send
        Do Until reader.ReadyState = 4
            DoEvents
        Loop
If reader.status = 200 Then
    getAPI = reader.responseText
Else
    MsgBox reader.status
End If
End Function

Function generateHTML(Title As String, subTitle As String, primaryMessage As String, detail1 As String, detail2 As String, detail3 As String, hasLink As Boolean) As String

Dim tblHeading As String, tblFooter As String, strHTMLBody As String

If hasLink Then
    primaryMessage = "<a href = '" & primaryMessage & "'>Check Folder</a>"
End If

tblHeading = "<table style=""width: 100%; margin: 0 auto; padding: 2em 3em; text-align: center; background-color: #fafafa;"">" & _
                            "<tbody>" & _
                                "<tr><td><h2 style=""color: #414141; font-size: 28px; margin-top: 0;"">" & Title & "</h2></td></tr>" & _
                                "<tr><td><p style=""color: rgb(73, 73, 73);"">" & subTitle & "</p></td></tr>" & _
                                 "<tr><td><table style=""padding: 1em; text-align: center;"">" & _
                                     "<tr><td style=""padding: 1em 1.5em; background: #FF6B00; "">" & primaryMessage & "</td></tr>" & _
                                "</table></td></tr>" & _
                            "</tbody>" & _
                        "</table>"
                        
tblFooter = "<table style=""width: 100%; margin: 0 auto; padding: 3em; background: #2b2b2b; color: rgba(255,255,255,.5);"">" & _
                        "<tbody>" & _
                            "<tr style=""border-collapse: collapse;""><td style=""padding: 1em; color: #c9c9c9;"">Details</td></tr>" & _
                            "<tr style=""border-collapse: collapse;""><td style=""padding: .1em 2em;"">" & detail1 & "</td></tr>" & _
                            "<tr style=""border-collapse: collapse;""><td style=""padding: .1em 2em;"">" & detail2 & "</td></tr>" & _
                            "<tr style=""border-collapse: collapse;""><td style=""padding: .1em 2em 1em 2em;"">" & detail3 & "</td></tr>" & _
                        "</tbody>" & _
                    "</table>"

strHTMLBody = "" & _
"<!DOCTYPE html><html lang=""en"" xmlns=""http://www.w3.org/1999/xhtml"" xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:o=""urn:schemas-microsoft-com:office:office"">" & _
    "<head><meta charset=""utf-8""><title>Working DB Notification</title></head>" & _
    "<body style=""margin: 0 auto; Font-family: 'Montserrat', sans-serif; font-weight: 400; font-size: 15px; line-height: 1.8;"">" & _
        "<table style=""max-width: 600px; margin: 0 auto; text-align: center; "">" & _
            "<tbody>" & _
                "<tr><td>" & tblHeading & "</td></tr>" & _
                "<tr><td>" & tblFooter & "</td></tr>" & _
                "<tr><td><p style=""color: rgb(192, 192, 192); text-align: center;"">This email was created by  &copy; workingDB</p></td></tr>" & _
            "</tbody>" & _
        "</table>" & _
    "</body>" & _
"</html>"

generateHTML = strHTMLBody

End Function

Function dailySummary(Title As String, subTitle As String, lates() As String, todays() As String, nexts() As String) As String

Dim tblHeading As String, tblStepOverview As String, strHTMLBody As String

tblHeading = "<table style=""width: 100%; margin: 0 auto; padding: 2em 2em 1em 2em; text-align: center; background-color: #fafafa;"">" & _
                            "<tbody>" & _
                                "<tr><td><h2 style=""color: #414141; font-size: 28px; margin-top: 0;"">" & Title & "</h2></td></tr>" & _
                                "<tr><td><p style=""color: rgb(73, 73, 73);"">Here is what you have happening...</p></td></tr>" & _
                            "</tbody>" & _
                        "</table>"
                        
Dim i As Long, lateTable As String, lateCount As Long, todayCount As Long, nextCount As Long, todayTable As String, nextTable As String, varStr As String
i = 0
lateCount = UBound(lates)
todayCount = UBound(todays)
nextCount = UBound(nexts)
tblStepOverview = ""

varStr = ""
If lates(1) <> "" Then
    For i = 1 To UBound(lates)
        lateTable = lateTable & "<tr style=""border-collapse: collapse;"">" & _
                                                "<td style=""padding: .1em 2em;"">" & Split(lates(i), ",")(0) & "</td>" & _
                                                "<td style=""padding: .1em 2em;"">" & Split(lates(i), ",")(1) & "</td>" & _
                                                "<td style=""padding: .1em 2em;  color: rgb(255,195,195);"">" & Split(lates(i), ",")(2) & "</td></tr>"
    Next i
    If lateCount > 1 Then varStr = "s"
    tblStepOverview = tblStepOverview & "<table style=""width: 100%; margin: 0 auto; background: #2b2b2b; color: rgb(255,255,255);""><tr><th style=""padding: 1em; font-size: 20px; color: rgb(255,150,150); display: table-header-group;"" colspan=""3"">You have " & _
                                                                lateCount & " item" & varStr & " overdue</th></tr><tbody>" & _
                                                            "<tr style=""padding: .1em 2em;""><th style=""text-align: left"">Part Number</th><th style=""text-align: left"">Item</th><th style=""text-align: left"">Due</th></tr>" & lateTable & "</tbody></table>"
End If

varStr = ""
If todays(1) <> "" Then
    For i = 1 To UBound(todays)
        todayTable = todayTable & "<tr style=""border-collapse: collapse;"">" & _
                                                "<td style=""padding: .1em 2em;"">" & Split(todays(i), ",")(0) & "</td>" & _
                                                "<td style=""padding: .1em 2em;"">" & Split(todays(i), ",")(1) & "</td>" & _
                                                "<td style=""padding: .1em 2em;"">" & Split(todays(i), ",")(2) & "</td></tr>"
    Next i
    If todayCount > 1 Then varStr = "s"
    tblStepOverview = tblStepOverview & "<table style=""width: 100%; margin: 0 auto; background: #2b2b2b; color: rgb(255,255,255);""><tr><th style=""padding: 1em; font-size: 20px; color: rgb(235,200,200); display: table-header-group;"" colspan=""3"">You have " & _
                                                                todayCount & " item" & varStr & " due today</th></tr><tbody>" & _
                                                            "<tr style=""padding: .1em 2em;""><th style=""text-align: left"">Part Number</th><th style=""text-align: left"">Item</th><th style=""text-align: left"">Due</th></tr>" & todayTable & "</tbody></table>"
End If

varStr = ""
If nexts(1) <> "" Then
    For i = 1 To UBound(nexts)
        nextTable = nextTable & "<tr style=""border-collapse: collapse;"">" & _
                                                "<td style=""padding: .1em 2em;"">" & Split(nexts(i), ",")(0) & "</td>" & _
                                                "<td style=""padding: .1em 2em;"">" & Split(nexts(i), ",")(1) & "</td>" & _
                                                "<td style=""padding: .1em 2em;"">" & Split(nexts(i), ",")(2) & "</td></tr>"
    Next i
    If nextCount > 1 Then varStr = "s"
    tblStepOverview = tblStepOverview & "<table style=""width: 100%; margin: 0 auto; background: #2b2b2b; color: rgb(255,255,255);""><tr><th style=""padding: 1em; font-size: 20px; color: rgb(235,235,235); display: table-header-group;"" colspan=""3"">You have " & _
                                                                nextCount & " item" & varStr & " due soon</th></tr><tbody>" & _
                                                            "<tr style=""padding: .1em 2em;""><th style=""text-align: left"">Part Number</th><th style=""text-align: left"">Item</th><th style=""text-align: left"">Due</th></tr>" & nextTable & "</tbody></table>"
End If

strHTMLBody = "" & _
"<!DOCTYPE html><html lang=""en"" xmlns=""http://www.w3.org/1999/xhtml"" xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:o=""urn:schemas-microsoft-com:office:office"">" & _
    "<head><meta charset=""utf-8""><title>Working DB Notification</title></head>" & _
    "<body style=""margin: 0 auto; Font-family: 'Montserrat', sans-serif; font-weight: 400; font-size: 15px; line-height: 1.8;"">" & _
        "<table style=""max-width: 600px; margin: 0 auto; text-align: center; "">" & _
            "<tbody>" & _
                "<tr><td>" & tblHeading & "</td></tr>" & _
                "<tr><td>" & tblStepOverview & "</td></tr>" & _
                "<tr><td><p style=""color: rgb(192, 192, 192); text-align: center;"">If you wish to no longer receive these emails,  go into your account menu in the workingDB to disable daily summary notifications</p></td></tr>" & _
                "<tr><td><p style=""color: rgb(192, 192, 192); text-align: center;"">This email was created by  &copy; workingDB</p></td></tr>" & _
            "</tbody>" & _
        "</table>" & _
    "</body>" & _
"</html>"

dailySummary = strHTMLBody

End Function

Public Sub registerCPCUpdates(table As String, ID As Variant, column As String, oldVal As Variant, newVal As Variant, Optional tag0 As String, Optional tag1 As String)
Dim sqlColumns As String, sqlValues As String

If (VarType(oldVal) = vbDate) Then
    oldVal = Format(oldVal, "mm/dd/yyyy")
End If

If (VarType(newVal) = vbDate) Then
    newVal = Format(newVal, "mm/dd/yyyy")
End If

If (IsNull(oldVal)) Then
    oldVal = ""
End If

If (IsNull(newVal)) Then
    newVal = ""
End If

sqlColumns = "(tableName,tableRecordId,updatedBy,updatedDate,columnName,previousData,newData,dataTag0"
                    
sqlValues = " values ('" & table & "', '" & ID & "', '" & Environ("username") & "', '" & Now() & "', '" & column & "', '" & StrQuoteReplace(CStr(oldVal)) & "', '" & StrQuoteReplace(CStr(newVal)) & "','" & tag0 & "'"

If (IsNull(tag1)) Then
    sqlColumns = sqlColumns & ")"
    sqlValues = sqlValues & ");"
Else
    sqlColumns = sqlColumns & ",dataTag1)"
    sqlValues = sqlValues & ",'" & tag1 & "');"
End If


CurrentDb().Execute "INSERT INTO tblCPC_UpdateTracking" & sqlColumns & sqlValues

End Sub

Function emailContentGen(subject As String, Title As String, subTitle As String, primaryMessage As String, detail1 As String, detail2 As String, detail3 As String) As String
emailContentGen = subject & "," & Title & "," & subTitle & "," & primaryMessage & "," & detail1 & "," & detail2 & "," & detail3
End Function

Function sendNotification(SendTo As String, notType As Integer, notPriority As Integer, desc As String, emailContent As String, Optional appName As String = "", Optional appId As Long) As Boolean
sendNotification = True

On Error GoTo err_handler

'has this person been notified about this thing today already?
Dim rsNotifications As Recordset
Set rsNotifications = CurrentDb().OpenRecordset("SELECT * from tblNotificationsSP WHERE recipientUser = '" & SendTo & "' AND notificationDescription = '" & StrQuoteReplace(desc) & "' AND sentDate > #" & Date - 1 & "#")
If rsNotifications.RecordCount > 0 Then
    If rsNotifications!notificationType = 1 Then
        Dim msgTxt As String
        If rsNotifications!senderUser = Environ("username") Then
            msgTxt = "Yo, you already did that today, let's wait 'til tomorrow to do it again."
        Else
            msgTxt = SendTo & " has already been nudged about this today by " & rsNotifications!senderUser & ". Let's wait until tomorrow to nudge them again."
        End If
        MsgBox msgTxt, vbInformation, "Hold on a minute..."
        sendNotification = False
        Exit Function
    End If
End If

Dim strValues
strValues = "'" & SendTo & "','" & getEmail(SendTo) & "','" & Environ("username") & "','" & getEmail(Environ("username")) & "','" & Now() & "'," & notType & "," & notPriority & ",'" & StrQuoteReplace(desc) & "','" & appName & "'," & appId & ",'" & StrQuoteReplace(emailContent) & "'"

CurrentDb().Execute "INSERT INTO tblNotificationsSP(recipientUser,recipientEmail,senderUser,senderEmail,sentDate,notificationType,notificationPriority,notificationDescription,appName,appId,emailContent) VALUES(" & strValues & ");"

Exit Function
err_handler:
sendNotification = False
MsgBox Err.description, vbCritical, "Notification Error"
End Function

Function privilege(pref) As Boolean
    privilege = DLookup("[" & pref & "]", "[tblPermissions]", "[User] = '" & Environ("username") & "'")
End Function

Function userData(data) As String
    userData = DLookup("[" & data & "]", "[tblPermissions]", "[User] = '" & Environ("username") & "'")
End Function

Function restrict(userName As String, dept As String, Optional Level As String) As Boolean
Dim d As Boolean, l As Boolean
d = False
l = False

    If (DLookup("[Dept]", "[tblPermissions]", "[User] = '" & userName & "'") = dept) Then
        d = True
    End If
    
    If (IsNull(Level) Or Level = "") Then
        restrict = Not (d)
    Else
        If (DLookup("[Level]", "[tblPermissions]", "[User] = '" & userName & "'") = Level) Then
            l = True
        End If
        restrict = Not (d And l)
    End If

End Function

Public Sub checkForFirstTimeRun()

Dim db As Database
Set db = CurrentDb()
Dim rsAnalytics As Recordset, ranThisWeek As Boolean

Set rsAnalytics = db.OpenRecordset("SELECT max(dateUsed) as anaDate from tblAnalytics WHERE module = 'firstTimeRun'")
If Format(rsAnalytics!anaDate, "mm/dd/yyyy") = Format(Date, "mm/dd/yyyy") Then Exit Sub 'if max date is today, then this has already ran.

'Call grabSummaryInfo 'disabled while in Beta
Call checkProgramEvents

db.Execute "INSERT INTO tblAnalytics (module,form,userName,dateUsed) VALUES ('firstTimeRun','Form_frmSplash','" & Environ("username") & "','" & Now() & "')"

End Sub

Function checkNewPartNumbers()

On Error Resume Next

Dim pnLogMax, spListMax
'grab highest part number of each and compare
pnLogMax = DMax("Part_Number", "dbo_tblParts")
spListMax = DMax("newPartNumber", "tblPartNumbers")

'if OK, exit
If pnLogMax = spListMax Then Exit Function

'if NG, add to current and check after each one
Dim db As Database
Set db = CurrentDb()
Dim rsSP As Recordset, rsLog As Recordset

Set rsSP = db.OpenRecordset("tblPartNumbers")
Set rsLog = db.OpenRecordset("dbo_tblParts")

Do While pnLogMax > spListMax
    rsSP.addNew
    
    spListMax = spListMax + 1
    
    rsSP!newPartNumber = spListMax
    
    Set rsLog = db.OpenRecordset("SELECT * from dbo_tblParts WHERE Part_Number = " & spListMax)
    If rsLog.RecordCount = 0 Then GoTo nextOne
    
    rsSP!creator = Nz(rsLog!Issuer, "workingdb")
    rsSP!PartDescription = Nz(rsLog!Part_Description, "empty")
    rsSP!customerId = Nz(rsLog!customer, 0)
    rsSP!customerPartNumber = rsLog!Customer_Part_Number
    rsSP!materialType = Nz(DLookup("Material_Type", "dbo_tblMaterialTypes", "Material_Type_ID = " & Nz(rsLog!Material_Type, 0)))
    rsSP!Color = Nz(DLookup("Color_Name", "dbo_tblColors", "Color_ID = " & Nz(rsLog!Color, 0)), "")
    rsSP!NJPpartNumber = rsLog!Nifco_Japan_Part_Number
    rsSP!Notes = DLookup("Notes", "dbo_tblNotes", "Part_Number = " & spListMax)
    rsSP!partNumberType = 1
    
    rsSP.Update
nextOne:
Loop


Do While spListMax > pnLogMax
    pnLogMax = pnLogMax + 1

    Set rsLog = db.OpenRecordset("dbo_tblParts")
    
    rsLog.addNew
    rsLog!Part_Number = pnLogMax
    rsLog.Update
    rsLog.MoveNext

nextOne1:
Loop

End Function

Function grabSummaryInfo(Optional specificUser As String = "") As Boolean
grabSummaryInfo = False

Dim db As Database
Set db = CurrentDb()
Dim rsPeople As Recordset, rsPartNumbers As Recordset, rsOpenSteps As Recordset, rsOpenWOs As Recordset, rsNoti As Recordset, rsAnalytics As Recordset
Dim lateSteps() As String, todaySteps() As String, nextSteps() As String
Dim li As Long, ti As Long, ni As Long
Dim strQry, ranThisWeek As Boolean

Set rsAnalytics = db.OpenRecordset("SELECT max(dateUsed) as anaDate from tblAnalytics WHERE module = 'firstTimeRun'")
ranThisWeek = Format(rsAnalytics!anaDate, "ww", vbMonday, vbFirstFourDays) = Format(Date, "ww", vbMonday, vbFirstFourDays)

strQry = ""
If specificUser <> "" Then strQry = " AND user = '" & specificUser & "'"

Set rsPeople = db.OpenRecordset("SELECT * from tblPermissions WHERE Inactive = False" & strQry)
    li = 1
    ti = 1
    ni = 1
    ReDim Preserve lateSteps(li)
    ReDim Preserve todaySteps(ti)
    ReDim Preserve nextSteps(ni)
    lateSteps(li) = ""
    todaySteps(ti) = ""
    nextSteps(ni) = ""

Do While Not rsPeople.EOF 'go through every active person
    If rsPeople!notifications = 1 And specificUser = "" Then GoTo nextPerson 'this person wants no notifications
    If rsPeople!notifications = 2 And ranThisWeek And specificUser = "" Then GoTo nextPerson 'this person only wants weekly notifications
    
    If rsPeople!dept = "Design" Then
        Set rsOpenWOs = db.OpenRecordset("SELECT * from qryWOsforNotifications WHERE assignee = '" & rsPeople!User & "'")
    
        Do While Not rsOpenWOs.EOF
            Select Case rsOpenWOs!Due
                    Case Date 'due today
                        ReDim Preserve todaySteps(ti)
                        todaySteps(ti) = rsOpenWOs!Part_Number & ",WO: " & rsOpenWOs!Request_Type & ",Today"
                        ti = ti + 1
                    Case Is < Date 'over due
                        ReDim Preserve lateSteps(li)
                        lateSteps(li) = rsOpenWOs!Part_Number & ",WO: " & rsOpenWOs!Request_Type & "," & Format(rsOpenWOs!Due, "mm/dd/yyyy")
                        li = li + 1
                    Case Is <= (Date + 7) 'due in next week
                        ReDim Preserve nextSteps(ni)
                        nextSteps(ni) = rsOpenWOs!Part_Number & ",WO: " & rsOpenWOs!Request_Type & "," & Format(rsOpenWOs!Due, "mm/dd/yyyy")
                        ni = ni + 1
                End Select
            rsOpenWOs.MoveNext
        Loop
        rsOpenWOs.Close
        Set rsOpenWOs = Nothing
    End If

    Set rsPartNumbers = db.OpenRecordset("SELECT * from tblPartTeam WHERE person = '" & rsPeople!User & "'") 'find all of their projects, go through every part project they are on
    Do While Not rsPartNumbers.EOF
        Set rsOpenSteps = db.OpenRecordset("SELECT * from tblPartSteps WHERE partNumber = '" & rsPartNumbers!partNumber & "' AND responsible = '" & rsPeople!dept & "' AND status <> 'Closed'")
        
        Do While Not rsOpenSteps.EOF
            Select Case rsOpenSteps!dueDate
                Case Date 'due today
                    ReDim Preserve todaySteps(ti)
                    todaySteps(ti) = rsOpenSteps!partNumber & "," & rsOpenSteps!stepType & ",Today"
                    ti = ti + 1
                Case Is < Date 'over due
                    ReDim Preserve lateSteps(li)
                    lateSteps(li) = rsOpenSteps!partNumber & "," & rsOpenSteps!stepType & "," & Format(rsOpenSteps!dueDate, "mm/dd/yyyy")
                    li = li + 1
                Case Is <= (Date + 7) 'due in next week
                    ReDim Preserve nextSteps(ni)
                    nextSteps(ni) = rsOpenSteps!partNumber & "," & rsOpenSteps!stepType & "," & Format(rsOpenSteps!dueDate, "mm/dd/yyyy")
                    ni = ni + 1
            End Select
        
            rsOpenSteps.MoveNext
        Loop
        rsOpenSteps.Close
        Set rsOpenSteps = Nothing
        rsPartNumbers.MoveNext
    Loop
    rsPartNumbers.Close
    Set rsPartNumbers = Nothing
    
    If ti + li + ni > 1 Then
        Set rsNoti = db.OpenRecordset("tblNotificationsSP")
        With rsNoti
            .addNew
            !recipientUser = rsPeople!User
            !recipientEmail = rsPeople!userEmail
            !senderUser = Environ("username")
            !senderEmail = getEmail(Environ("username"))
            !sentDate = Now()
            !readDate = Now()
            !notificationType = 9
            !notificationPriority = 2
            !notificationDescription = "Summary Email"
            !emailContent = StrQuoteReplace(dailySummary("Hi " & rsPeople!firstName, "Here is your daily summary!", lateSteps(), todaySteps(), nextSteps()))
            .Update
        End With
        rsNoti.Close
        Set rsNoti = Nothing
    End If
    
nextPerson:
    rsPeople.MoveNext
Loop

grabSummaryInfo = True

End Function

Function checkProgramEvents() As Boolean

Dim db As Database
Set db = CurrentDb()

Dim rsProgram As Recordset, rsEvents As Recordset, rsWO As Recordset, rsComments As Recordset, rsPeople As Recordset, rsNoti As Recordset
Dim controlNum As Long, comments As String, dueDate, body As String, strValues

dueDate = addWorkdays(Date, 5)

Set rsEvents = db.OpenRecordset("SELECT * from tblProgramEvents WHERE designWOcreated = False AND eventDate BETWEEN #" & Date & "# AND #" & Date + 50 & "#")
Set rsPeople = db.OpenRecordset("SELECT * from tblPermissions WHERE designWOid = 1 AND InActive = FALSE")

Do While Not rsEvents.EOF
    Set rsProgram = db.OpenRecordset("SELECT * from tblPrograms WHERE ID = " & rsEvents!programId)
    
    Set rsWO = db.OpenRecordset("dbo_tblDRS")
    rsWO.addNew
        With rsWO
            !Issue_Date = Date
            !Approval_Status = 1
            !Requester = "automated"
            !DR_Level = 1
            !Request_Type = 23
            !Design_Level = 4 'ETA
            !Due_Date = dueDate
            !Part_Number = "D8157"
            !Part_Description = "Program Review"
            !Model_Code = rsProgram!modelCode
        End With
    rsWO.Update
    
    controlNum = db.OpenRecordset("SELECT @@identity")(0).Value
    comments = "'Hold program review for " & rsProgram!modelCode & " " & rsEvents!eventTitle & "'"
    
    db.Execute "INSERT INTO dbo_tblComments(Control_Number, Comments) VALUES(" & controlNum & "," & comments & ")"
    
    body = emailContentGen("Program Review WO", "WO Notice", "WO Auto-Created for " & rsProgram!modelCode & " Program Review", "Event: " & rsEvents!eventTitle, "WO#" & controlNum, "Due: " & dueDate, "Sent On: " & CStr(Now()))
    
    rsEvents.Edit
    rsEvents!designWOcreated = True
    rsEvents.Update
    
    Set rsNoti = db.OpenRecordset("tblNotificationsSP")
    rsPeople.MoveFirst
    Do While Not rsPeople.EOF
        With rsNoti
            .addNew
            !recipientUser = rsPeople!User
            !recipientEmail = rsPeople!userEmail
            !senderUser = Environ("username")
            !senderEmail = getEmail(Environ("username"))
            !sentDate = Now()
            !notificationType = 10
            !notificationPriority = 2
            !notificationDescription = "WO Auto-Created for " & rsProgram!modelCode & " Program Review"
            !appName = "Design WO"
            !appId = controlNum
            !emailContent = body
            .Update
        End With
        rsPeople.MoveNext
    Loop
    
    rsNoti.Close
    Set rsNoti = Nothing
        
    rsProgram.Close
    Set rsProgram = Nothing
    
    rsEvents.MoveNext
Loop

rsEvents.Close
Set rsEvents = Nothing

rsPeople.Close
Set rsPeople = Nothing

End Function

Function getEmail(userName As String) As String

On Error GoTo tryOracle
Dim rsPermissions As Recordset
Set rsPermissions = CurrentDb().OpenRecordset("SELECT * from tblPermissions WHERE user = '" & userName & "'")
getEmail = rsPermissions!userEmail
rsPermissions.Close

Exit Function
tryOracle:
Dim rsEmployee As Recordset
Set rsEmployee = CurrentDb().OpenRecordset("SELECT FIRST_NAME, LAST_NAME, EMAIL_ADDRESS FROM APPS_XXCUS_USER_EMPLOYEES_V WHERE USER_NAME = '" & StrConv(userName, vbUpperCase) & "'")
getEmail = Nz(rsEmployee!EMAIL_ADDRESS, "")

End Function

Function splitString(a, b, c) As String
    On Error GoTo errorCatch
    splitString = Split(a, b)(c)
    Exit Function
errorCatch:
    splitString = ""
End Function

Function getYear(projectNumber As String)
    If Len(projectNumber) = 7 Then
        getYear = Left(Year(Now), 2) & Mid(projectNumber, 2, 2)
    Else
        getYear = Left(Year(Now), 2) & Mid(projectNumber, 3, 2)
    End If
End Function

Function labelCycle(checkLabel As String, nameLabel As String, Optional controlSourceVal As String = "") As String()
    Dim returnVal(0 To 1) As String, sortLabel As String
    If controlSourceVal = "" Then
        sortLabel = nameLabel
    Else
        sortLabel = controlSourceVal
    End If
    Select Case True
        Case InStr(checkLabel, "-") > 0
            returnVal(0) = nameLabel & " >"
            returnVal(1) = sortLabel & " DESC"
        Case InStr(checkLabel, ">") > 0
            returnVal(0) = nameLabel & " <"
            returnVal(1) = sortLabel & " ASC"
        Case Else
            returnVal(0) = nameLabel & " -"
            returnVal(1) = sortLabel & " ASC"
    End Select
    labelCycle = returnVal
End Function

Function idNAM(inputVal As Variant, typeVal As Variant) As Variant
Dim rs1 As Recordset
idNAM = ""

If inputVal = "" Then Exit Function

If typeVal = "ID" Then
    Set rs1 = CurrentDb.OpenRecordset("SELECT SEGMENT1 FROM APPS_MTL_SYSTEM_ITEMS WHERE INVENTORY_ITEM_ID = " & inputVal, dbOpenSnapshot)
    If rs1.RecordCount = 0 Then GoTo exitFunction
    idNAM = rs1("SEGMENT1")
End If

If typeVal = "NAM" Then
    Set rs1 = CurrentDb.OpenRecordset("SELECT INVENTORY_ITEM_ID FROM APPS_MTL_SYSTEM_ITEMS WHERE SEGMENT1 = '" & inputVal & "'", dbOpenSnapshot)
    If rs1.RecordCount = 0 Then GoTo exitFunction
    idNAM = rs1("INVENTORY_ITEM_ID")
End If

exitFunction:
rs1.Close
Set rs1 = Nothing
End Function

Function getDescriptionFromId(inventId As Long) As String
Dim rs1 As Recordset

getDescriptionFromId = ""
If IsNull(inventId) Then Exit Function

Set rs1 = CurrentDb.OpenRecordset("SELECT DESCRIPTION FROM APPS_MTL_SYSTEM_ITEMS WHERE INVENTORY_ITEM_ID = " & inventId, dbOpenSnapshot)
If rs1.RecordCount = 0 Then GoTo exitFunction
getDescriptionFromId = rs1("DESCRIPTION")

exitFunction:
rs1.Close
Set rs1 = Nothing
End Function

Public Function StrQuoteReplace(strValue)
  StrQuoteReplace = Replace(strValue, "'", "''")
End Function

Public Function wdbEmail(ByVal strTo As String, ByVal strCC As String, ByVal strSubject As String, body As String) As Boolean
wdbEmail = True
Dim SendItems As New clsOutlookCreateItem
Set SendItems = New clsOutlookCreateItem

If IsNull(strCC) Then strCC = ""

SendItems.CreateMailItem SendTo:=strTo, _
                             CC:=strCC, _
                             subject:=strSubject, _
                             HTMLBody:=body
    Set SendItems = Nothing
    
Exit Function
err_handler:
wdbEmail = False
MsgBox "something went wrong, sorry! Please let Jacob Brown know.", vbCritical, "Uh oh."
End Function