Option Compare Database
Option Explicit

Public Function runAll()
On Error Resume Next

Call grabSummaryInfo
Call grabIssueInfo

Application.Quit

End Function

Function grabSummaryInfo(Optional specificUser As String = "") As Boolean
On Error Resume Next

grabSummaryInfo = False

Dim db As Database
Set db = CurrentDb()
Dim rsPeople As Recordset, rsOpenSteps As Recordset, rsOpenWOs As Recordset, rsNoti As Recordset, rsAnalytics As Recordset
Dim lateSteps() As String, todaySteps() As String, nextSteps() As String
Dim li As Long, ti As Long, ni As Long
Dim strQry, ranThisWeek As Boolean
Dim recordsetName As String

Set rsAnalytics = db.OpenRecordset("SELECT max(dateUsed) as anaDate from tblAnalytics WHERE module = 'firstTimeRun'")
ranThisWeek = Format(rsAnalytics!anaDate, "ww", vbMonday, vbFirstFourDays) = Format(Date, "ww", vbMonday, vbFirstFourDays)

rsAnalytics.Close: Set rsAnalytics = Nothing

CurrentDb.Execute "INSERT INTO tblAnalytics (module,form,userName,dateUsed) VALUES ('summaryEmail','Form_frmSplash','" & Environ("username") & "','" & Now() & "')"

strQry = ""
If specificUser <> "" Then strQry = " AND user = '" & specificUser & "'"

Set rsPeople = db.OpenRecordset("SELECT * from tblPermissions WHERE Inactive = False" & strQry)
    li = 0
    ti = 0
    ni = 0
    ReDim Preserve lateSteps(li)
    ReDim Preserve todaySteps(ti)
    ReDim Preserve nextSteps(ni)

Do While Not rsPeople.EOF 'go through every active person
    If rsPeople!notifications = 1 And specificUser = "" Then GoTo nextPerson 'this person wants no notifications
    If rsPeople!notifications = 2 And ranThisWeek And specificUser = "" Then GoTo nextPerson 'this person only wants weekly notifications
    
    li = 0
    ti = 0
    ni = 0
    Erase lateSteps, todaySteps, nextSteps
    ReDim lateSteps(li)
    ReDim todaySteps(ti)
    ReDim nextSteps(ni)

    If rsPeople!Level = "Engineer" Then
        recordsetName = "SELECT * FROM qryStepApprovalTracker"
    Else
        recordsetName = "SELECT * FROM sqryStepApprovalTracker_Approvals_SupervisorsUp"
    End If

    Set rsOpenSteps = db.OpenRecordset(recordsetName & _
                                " WHERE person = '" & rsPeople!User & "' AND due <= Date()+7")
    
    Do While (Not rsOpenSteps.EOF And Not (ti > 15 And li > 15 And ni > 15))
        Select Case rsOpenSteps!due
            Case Date 'due today
                If ti > 15 Then
                    ti = ti + 1
                    GoTo nextStep
                End If
                ReDim Preserve todaySteps(ti)
                todaySteps(ti) = rsOpenSteps!partNumber & "," & rsOpenSteps!Action & ",Today"
                ti = ti + 1
            Case Is < Date 'over due
                If li > 15 Then
                    li = li + 1
                    GoTo nextStep
                End If
                ReDim Preserve lateSteps(li)
                lateSteps(li) = rsOpenSteps!partNumber & "," & rsOpenSteps!Action & "," & Format(rsOpenSteps!due, "mm/dd/yyyy")
                li = li + 1
            Case Is <= (Date + 7) 'due in next week
                If ni > 15 Then
                    ni = ni + 1
                    GoTo nextStep
                End If
                ReDim Preserve nextSteps(ni)
                nextSteps(ni) = rsOpenSteps!partNumber & "," & rsOpenSteps!Action & "," & Format(rsOpenSteps!due, "mm/dd/yyyy")
                ni = ni + 1
        End Select
nextStep:
        rsOpenSteps.MoveNext
    Loop
    rsOpenSteps.Close
    Set rsOpenSteps = Nothing
    
    If ti + li + ni > 0 Then
        Set rsNoti = db.OpenRecordset("tblNotificationsSP")
        With rsNoti
            .AddNew
            !recipientUser = rsPeople!User
            !recipientEmail = rsPeople!userEmail
            !senderUser = "workingDB"
            !senderEmail = "workingDB@us.nifco.com"
            !sentDate = Now()
            !readDate = Now()
            !notificationType = 9
            !notificationPriority = 2
            !notificationDescription = "Summary Email"
            !emailContent = StrQuoteReplace(dailySummary("Hi " & rsPeople!firstName, "Here is what you have going on...", lateSteps(), todaySteps(), nextSteps(), li, ti, ni))
            .Update
        End With
        rsNoti.Close
        Set rsNoti = Nothing
        Debug.Print rsPeople!User
    End If
    
nextPerson:
    rsPeople.MoveNext
Loop

grabSummaryInfo = True

On Error Resume Next
rsPeople.Close: Set rsPeople = Nothing
rsOpenSteps.Close: Set rsOpenSteps = Nothing
rsOpenWOs.Close: Set rsOpenWOs = Nothing
rsNoti.Close: Set rsNoti = Nothing

End Function

Function grabIssueInfo(Optional specificUser As String = "") As Boolean
On Error Resume Next

grabIssueInfo = False

Dim db As Database
Set db = CurrentDb()
Dim rsPeople As Recordset, rsNoti As Recordset, rsOpenIssues As Recordset
Dim lateSteps() As String, todaySteps() As String, nextSteps() As String
Dim li As Long, ti As Long, ni As Long
Dim strQry

strQry = ""
If specificUser <> "" Then strQry = " AND user = '" & specificUser & "'"

Set rsPeople = db.OpenRecordset("SELECT * from tblPermissions WHERE Inactive = False" & strQry)

li = 0
ti = 0
ni = 0
ReDim Preserve lateSteps(li)
ReDim Preserve todaySteps(ti)
ReDim Preserve nextSteps(ni)

Do While Not rsPeople.EOF 'go through every active person
    li = 0
    ti = 0
    ni = 0
    Erase lateSteps, todaySteps, nextSteps
    ReDim lateSteps(li)
    ReDim todaySteps(ti)
    ReDim nextSteps(ni)
    
    Set rsOpenIssues = db.OpenRecordset("SELECT * FROM qryOpenIssues_summaryEmail WHERE inCharge = '" & rsPeople!User & "' AND closeDate is null AND dueDate <= Date()+7")
    
    Do While Not rsOpenIssues.EOF
        Select Case rsOpenIssues!dueDate
            Case Date 'due today
                ReDim Preserve todaySteps(ti)
                todaySteps(ti) = rsOpenIssues!partNumber & ",Open Issue: " & rsOpenIssues!issueType & "-" & rsOpenIssues!issueSource & ",Today"
                ti = ti + 1
            Case Is < Date 'over due
                ReDim Preserve lateSteps(li)
                lateSteps(li) = rsOpenIssues!partNumber & ",Open Issue: " & rsOpenIssues!issueType & "-" & rsOpenIssues!issueSource & "," & Format(rsOpenIssues!dueDate, "mm/dd/yyyy")
                li = li + 1
            Case Is <= (Date + 7) 'due in next week
                ReDim Preserve nextSteps(ni)
                nextSteps(ni) = rsOpenIssues!partNumber & ",Open Issue: " & rsOpenIssues!issueType & "-" & rsOpenIssues!issueSource & "," & Format(rsOpenIssues!dueDate, "mm/dd/yyyy")
                ni = ni + 1
        End Select
        rsOpenIssues.MoveNext
    Loop
    
    If ti + li + ni > 0 Then
        Set rsNoti = db.OpenRecordset("tblNotificationsSP")
        With rsNoti
            .AddNew
            !recipientUser = rsPeople!User
            !recipientEmail = rsPeople!userEmail
            !senderUser = "workingDB"
            !senderEmail = "workingDB@us.nifco.com"
            !sentDate = Now()
            !readDate = Now()
            !notificationType = 9
            !notificationPriority = 2
            !notificationDescription = "Current Open Issues"
            !emailContent = StrQuoteReplace(dailySummary("Hi " & rsPeople!firstName, "Here are the issues assigned to you...", lateSteps(), todaySteps(), nextSteps(), li, ti, ni, False))
            .Update
        End With
        rsNoti.Close
        Set rsNoti = Nothing
        Debug.Print rsPeople!User
    End If
    
nextPerson:
    rsPeople.MoveNext
Loop

grabIssueInfo = True

On Error Resume Next
rsPeople.Close: Set rsPeople = Nothing
rsOpenIssues.Close: Set rsOpenIssues = Nothing
rsNoti.Close: Set rsNoti = Nothing

End Function

Function dailySummary(Title As String, subTitle As String, lates() As String, todays() As String, nexts() As String, lateCount As Long, todayCount As Long, nextCount As Long, Optional disclaimer As Boolean = True) As String

Dim tblHeading As String, tblStepOverview As String, strHTMLBody As String

tblHeading = "<table style=""width: 100%; margin: 0 auto; padding: 2em 2em 1em 2em; text-align: center; background-color: #fafafa;"">" & _
                            "<tbody>" & _
                                "<tr><td><h2 style=""color: #414141; font-size: 28px; margin-top: 0;"">" & Title & "</h2></td></tr>" & _
                                "<tr><td><p style=""color: rgb(73, 73, 73);"">Here is what you have happening...</p></td></tr>" & _
                            "</tbody>" & _
                        "</table>"
                        
Dim i As Long, lateTable As String, todayTable As String, nextTable As String, varStr As String, varStr1 As String, seeMore As String
seeMore = "<tr style=""border-collapse: collapse;""><td style=""padding: .1em 2em; font-style: italic;"" colspan=""3"">see the rest in the workingdb...</td></tr>"
i = 0
tblStepOverview = ""

varStr = ""
varStr1 = ""
If lates(0) <> "" Then
    For i = 0 To UBound(lates)
        lateTable = lateTable & "<tr style=""border-collapse: collapse;"">" & _
                                                "<td style=""padding: .1em 2em;"">" & Split(lates(i), ",")(0) & "</td>" & _
                                                "<td style=""padding: .1em 2em;"">" & Split(lates(i), ",")(1) & "</td>" & _
                                                "<td style=""padding: .1em 2em;  color: rgb(255,195,195);"">" & Split(lates(i), ",")(2) & "</td></tr>"
    Next i
    If lateCount > 1 Then varStr = "s"
    If lateCount > 15 Then varStr1 = seeMore
    tblStepOverview = tblStepOverview & "<table style=""width: 100%; margin: 0 auto; background: #2b2b2b; color: rgb(255,255,255);""><tr><th style=""padding: 1em; font-size: 20px; color: rgb(255,150,150); display: table-header-group;"" colspan=""3"">You have " & _
                                                                lateCount & " item" & varStr & " overdue</th></tr><tbody>" & _
                                                            "<tr style=""padding: .1em 2em;""><th style=""text-align: left"">Part Number</th><th style=""text-align: left"">Item</th><th style=""text-align: left"">Due</th></tr>" & lateTable & varStr1 & "</tbody></table>"
End If

varStr = ""
varStr1 = ""
If todays(0) <> "" Then
    For i = 0 To UBound(todays)
        todayTable = todayTable & "<tr style=""border-collapse: collapse;"">" & _
                                                "<td style=""padding: .1em 2em;"">" & Split(todays(i), ",")(0) & "</td>" & _
                                                "<td style=""padding: .1em 2em;"">" & Split(todays(i), ",")(1) & "</td>" & _
                                                "<td style=""padding: .1em 2em;"">" & Split(todays(i), ",")(2) & "</td></tr>"
    Next i
    If todayCount > 1 Then varStr = "s"
    If todayCount > 15 Then varStr1 = seeMore
    tblStepOverview = tblStepOverview & "<table style=""width: 100%; margin: 0 auto; background: #2b2b2b; color: rgb(255,255,255);""><tr><th style=""padding: 1em; font-size: 20px; color: rgb(235,200,200); display: table-header-group;"" colspan=""3"">You have " & _
                                                                todayCount & " item" & varStr & " due today</th></tr><tbody>" & _
                                                            "<tr style=""padding: .1em 2em;""><th style=""text-align: left"">Part Number</th><th style=""text-align: left"">Item</th><th style=""text-align: left"">Due</th></tr>" & todayTable & varStr1 & "</tbody></table>"
End If

varStr = ""
varStr1 = ""
If nexts(0) <> "" Then
    For i = 0 To UBound(nexts)
        nextTable = nextTable & "<tr style=""border-collapse: collapse;"">" & _
                                                "<td style=""padding: .1em 2em;"">" & Split(nexts(i), ",")(0) & "</td>" & _
                                                "<td style=""padding: .1em 2em;"">" & Split(nexts(i), ",")(1) & "</td>" & _
                                                "<td style=""padding: .1em 2em;"">" & Split(nexts(i), ",")(2) & "</td></tr>"
    Next i
    If nextCount > 1 Then varStr = "s"
    If nextCount > 15 Then varStr1 = seeMore
    tblStepOverview = tblStepOverview & "<table style=""width: 100%; margin: 0 auto; background: #2b2b2b; color: rgb(255,255,255);""><tr><th style=""padding: 1em; font-size: 20px; color: rgb(235,235,235); display: table-header-group;"" colspan=""3"">You have " & _
                                                                nextCount & " item" & varStr & " due soon</th></tr><tbody>" & _
                                                            "<tr style=""padding: .1em 2em;""><th style=""text-align: left"">Part Number</th><th style=""text-align: left"">Item</th><th style=""text-align: left"">Due</th></tr>" & nextTable & varStr1 & "</tbody></table>"
End If

Dim disclaimerTxt As String
disclaimerTxt = ""
If disclaimer Then disclaimerTxt = "<tr><td><p style=""color: rgb(192, 192, 192); text-align: center;"">If you wish to no longer receive these emails,  go into your account menu in the workingDB to disable daily summary notifications</p></td></tr>"

strHTMLBody = "" & _
"<!DOCTYPE html><html lang=""en"" xmlns=""http://www.w3.org/1999/xhtml"" xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:o=""urn:schemas-microsoft-com:office:office"">" & _
    "<head><meta charset=""utf-8""><title>Working DB Notification</title></head>" & _
    "<body style=""margin: 0 auto; Font-family: 'Montserrat', sans-serif; font-weight: 400; font-size: 15px; line-height: 1.8;"">" & _
        "<table style=""max-width: 600px; margin: 0 auto; text-align: center; "">" & _
            "<tbody>" & _
                "<tr><td>" & tblHeading & "</td></tr>" & _
                "<tr><td>" & tblStepOverview & "</td></tr>" & _
                disclaimerTxt & _
                "<tr><td><p style=""color: rgb(192, 192, 192); text-align: center;"">This email was created by  &copy; workingDB</p></td></tr>" & _
            "</tbody>" & _
        "</table>" & _
    "</body>" & _
"</html>"

dailySummary = strHTMLBody

End Function