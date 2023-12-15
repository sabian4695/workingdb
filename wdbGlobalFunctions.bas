Option Compare Database
Option Explicit

Public bClone As Boolean

Public Function addWeekdays(dateInput As Date, daysToAdd As Long) As Date
On Error GoTo err_handler

Dim i As Long, testDate As Date, daysLeft As Long, rsHolidays As Recordset
testDate = dateInput
daysLeft = daysToAdd

Do While daysLeft > 0
    testDate = testDate + 1
    If Weekday(testDate) = 7 Then ' IF SATURDAY -> skip to monday
        testDate = testDate + 1
        GoTo skipDate
    End If
    Set rsHolidays = CurrentDb().OpenRecordset("SELECT * from tblHolidays WHERE holidayDate = #" & testDate & "#")
    If rsHolidays.RecordCount > 0 Then GoTo skipDate ' IF HOLIDAY -> skip to next day
     daysLeft = daysLeft - 1
skipDate:
Loop

addWeekdays = testDate

Exit Function
err_handler:
    Call handleError("wdbGlobalFunctions", "addWeekdays", Err.description, Err.number)
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

Dim tblHeading As String, tblFooter As String, strHTMLBody As String, extraFooter As String

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
                extraFooter & _
                "<tr><td><p style=""color: rgb(192, 192, 192); text-align: center;"">This email was created by  &copy; workingDB</p></td></tr>" & _
            "</tbody>" & _
        "</table>" & _
    "</body>" & _
"</html>"

generateHTML = strHTMLBody

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
    If rsNotifications!notificationType = 1 Or rsNotifications!notificationType = 9 Then
        Dim msgTxt As String
        If rsNotifications!senderUser = Environ("username") Then
            msgTxt = "Yo, you already did that today, let's wait 'til tomorrow to do it again."
        Else
            msgTxt = SendTo & " has already been nudged about this today, by " & rsNotifications!senderUser & ". Let's wait until tomorrow to nudge them again."
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

Function restrict(userName As String, Dept As String, Optional Level As String) As Boolean
Dim d As Boolean, l As Boolean
d = False
l = False

    If (DLookup("[Dept]", "[tblPermissions]", "[User] = '" & userName & "'") = Dept) Then
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
Dim rs1 As Recordset
Set db = CurrentDb()
Set rs1 = db.OpenRecordset("SELECT * from tblAnalytics WHERE dateUsed > #" & Date - 1 & "# AND module = 'firstTimeRun'")

If rs1.RecordCount <> 0 Then Exit Sub

'DO STUFF

'CurrentDb().Execute ("Delete * from cubeCoreAnalytics")
'CurrentDb().Execute ("INSERT INTO cubeCoreAnalytics (module,form,userName,dateUsed) SELECT module,form,userName,dateUsed FROM tblAnalytics WHERE dateUsed>#" & Date - 14 & "#")

GoTo SKIPALL
Dim rsPartSteps As Recordset, rsOverdueMsgs As Recordset, rsPermissions As Recordset, msg As String, rsPartTeam As Recordset
Dim body As String, stepTitle As String, partNum As String

Set rsPartSteps = db.OpenRecordset("SELECT * from tblPartSteps WHERE responsible is not null AND closeDate is null AND dueDate is not null")
Set rsOverdueMsgs = db.OpenRecordset("SELECT recordId, partTrackingOverdueMessages from tblWdbExtras WHERE partTrackingOverdueMessages is not null")

Do While Not rsPartSteps.EOF
    Select Case rsPartSteps!dueDate
        Case Date
            Dim count As Long, whichVal As Integer
            rsOverdueMsgs.MoveLast
            count = rsOverdueMsgs.RecordCount
            whichVal = randomNumber(1, count)
            rsOverdueMsgs.MoveFirst
            rsOverdueMsgs.FindFirst "recordId = " & whichVal
            msg = rsOverdueMsgs!partTrackingOverdueMessages
        Case Is > Date
            msg = "Yo, this step is due today, please complete!"
        Case Date + 7
            msg = "This is your 1 week away warning, this step is due soon"
        Case Else
            GoTo nextRec
    End Select
    
    Set rsPartTeam = db.OpenRecordset("select * from tblPartTeam where partNumber = '" & partNum & "'")
    
    Do While Not rsPartTeam.EOF
        Set rsPermissions = db.OpenRecordset("select * from tblPermissions where user = '" & rsPartTeam!person & "'")
        If rsPermissions!Dept = rsPartSteps!responsible Then
            partNum = rsPartSteps!partNumber
            stepTitle = rsPartSteps!stepType
            body = emailContentGen("Just a reminder...", "WDB Reminder", msg, stepTitle, "Part Number: " & partNum, "This is an automated message. Jokes included are for the purpose of making this reminder fun and light hearted.", "Sent On: " & CStr(Date))
            Call sendNotification(rsPartSteps!responsible, 9, 2, "Please complete " & stepTitle & " for " & partNum, body, "Part Project", CInt(partNum))
        End If
nextOne:
        rsPartTeam.MoveNext
    Loop
    
nextRec:
    rsPartSteps.MoveNext
Loop

SKIPALL:
db.Execute "INSERT INTO tblAnalytics (module,form,userName,dateUsed) VALUES ('firstTimeRun','Form_frmSplash','" & Environ("username") & "','" & Now() & "')"

End Sub

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