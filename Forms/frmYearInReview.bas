Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Dim dbG As Database
Set dbG = OpenDatabase("C:\workingdb\WorkingDB_ghost.accde")
Dim rs As Recordset

Set rs = dbG.OpenRecordset("SELECT count(recordId) as idCount FROM tblWdbSessions WHERE user = '" & Environ("username") & "'")

Me.dbSessions = rs!idCount

rs.CLOSE
Set rs = Nothing
dbG.CLOSE
Set dbG = Nothing

DoEvents

'clicks
Dim db As Database
Set db = CurrentDb()
Set rs = db.OpenRecordset("SELECT * FROM tblAnalytics WHERE userName = '" & Environ("username") & "' AND year(dateUsed) = " & Year(Date))
Dim rsCount As Recordset

If rs.EOF Then
   Me.foldersOpened = 0
Else
   rs.MoveLast
   Me.foldersOpened = rs.RecordCount
End If

'parts searched
rs.filter = "module = 'filterbyPN'"
Set rsCount = rs.OpenRecordset

If rsCount.EOF Then
   Me.partsSearched = 0
Else
   rsCount.MoveLast
   Me.partsSearched = rsCount.RecordCount
End If

rs.CLOSE
Set rs = Nothing
rsCount.CLOSE
Set rsCount = Nothing

'busiest day
Set rs = db.OpenRecordset("SELECT DateValue(dateUsed) AS DAYUSED, COUNT(recordId) AS idCount FROM tblAnalytics WHERE userName = '" & Environ("username") & "' AND year(dateUsed) = " & Year(Date) & " GROUP BY DateValue(dateUsed) ORDER BY COUNT(recordId) DESC")
Me.busiestDay = rs!DAYUSED & vbNewLine & rs!idCount & " clicks"
rs.CLOSE
Set rs = Nothing

'top part numbers
Set rs = db.OpenRecordset("SELECT TOP 8 dataTag0, COUNT(recordId) AS idCount FROM tblAnalytics WHERE dataTag0 is not null AND dataTag0 <> '' AND userName = '" & Environ("username") & "' AND year(dateUsed) = " & Year(Date) & " GROUP BY dataTag0 ORDER BY COUNT(recordId) DESC")

Dim i As Long
i = 0
Do While Not rs.EOF
    i = i + 1
    If i > 8 Then Exit Do
    Me.Controls("partNumber" & i) = rs!dataTag0
    Me.Controls("imgPN" & i).Picture = "\\data\mdbdata\WorkingDB\_docs\Part_Pictures\" & rs!dataTag0 & ".png"
    rs.MoveNext
Loop

rs.CLOSE
Set rs = Nothing

Set rs = db.OpenRecordset("SELECT * FROM tblPartUpdateTracking WHERE year(updatedDate) = " & Year(Date) & " AND updatedBy = '" & Environ("username") & "'")

'steps closed
rs.filter = "newData = 'Closed' AND tableName = 'tblPartSteps'"
Set rsCount = rs.OpenRecordset

If rsCount.EOF Then
   Me.stepsClosed = 0
Else
   rsCount.MoveLast
   Me.stepsClosed = rsCount.RecordCount
End If

'steps approved
rs.filter = "newData <> 'Deleted' AND newData <> 'Created' AND tableName = 'tblPartTrackingApprovals'"
Set rsCount = rs.OpenRecordset

If rsCount.EOF Then
   Me.stepsApproved = 0
Else
   rsCount.MoveLast
   Me.stepsApproved = rsCount.RecordCount
End If

'files uploaded
rs.filter = "tableName = 'tblPartAttachmentsSP' AND newData = 'Uploaded'"
Set rsCount = rs.OpenRecordset

If rsCount.EOF Then
   Me.filesUploaded = 0
Else
   rsCount.MoveLast
   Me.filesUploaded = rsCount.RecordCount
End If

Set rs = db.OpenRecordset("SELECT * FROM tblPartUpdateTracking WHERE year(updatedDate) = " & Year(Date))
'nudges sent
rs.filter = "columnName = 'Nudge' AND previousData = 'From: " & Environ("username") & "'"
Set rsCount = rs.OpenRecordset

If rsCount.EOF Then
   Me.nudgesSent = 0
Else
   rsCount.MoveLast
   Me.nudgesSent = rsCount.RecordCount
End If

'nudges received
rs.filter = "columnName = 'Nudge' AND newData = 'To: " & Environ("username") & "'"
Set rsCount = rs.OpenRecordset

If rsCount.EOF Then
   Me.nudgesReceived = 0
Else
   rsCount.MoveLast
   Me.nudgesReceived = rsCount.RecordCount
End If

Set rs = db.OpenRecordset("SELECT DateValue(updatedDate) AS DAYUSED, COUNT(recordId) AS idCount FROM tblPartUpdateTracking " & _
    "WHERE newData = 'Closed' AND tableName = 'tblPartSteps' AND updatedBy = '" & Environ("username") & "' AND year(updatedDate) = " & Year(Date) & _
    " GROUP BY DateValue(updatedDate) ORDER BY COUNT(recordId) DESC")
Me.dayStepsClosed = rs!DAYUSED & vbNewLine & "(" & rs!idCount & " steps)"

'Most nudges sent to
Set rs = db.OpenRecordset("SELECT TOP 1 newData, COUNT(recordID) AS idCount FROM tblPartUpdateTracking WHERE " & _
    "year(updatedDate) = " & Year(Date) & " AND columnName = 'Nudge' AND previousData = 'From: " & Environ("username") & "' GROUP BY newData ORDER BY COUNT(recordId) DESC")
Me.mostNudgesSent = getFullName(Replace(rs!newData, "To: ", ""))

'Most nudges received from
Set rs = db.OpenRecordset("SELECT TOP 1 previousData, COUNT(recordID) AS idCount FROM tblPartUpdateTracking WHERE " & _
    "year(updatedDate) = " & Year(Date) & " AND columnName = 'Nudge' AND newData = 'To: " & Environ("username") & "' GROUP BY previousData ORDER BY COUNT(recordId) DESC")
Me.mostNudgesReceived = getFullName(Replace(rs!previousData, "From: ", ""))

'Top 5 Team Members
Set rs = db.OpenRecordset("SELECT TOP 10 person, count(recordId) FROM tblPartTeam WHERE " & _
    "person <> '" & Environ("username") & "' AND partNumber IN (SELECT partNumber FROM tblPartTeam WHERE person = '" & Environ("username") & "') " & _
    "GROUP BY person ORDER BY count(recordId) DESC")

i = 0
Do While Not rs.EOF
    i = i + 1
    If i > 10 Then Exit Do
    Me.Controls("TM" & i) = getFullName(rs!person)
    Me.Controls("imgTM" & i).Picture = "\\data\mdbdata\WorkingDB\Pictures\Avatars\" & rs!person & ".png"
    rs.MoveNext
Loop

rs.CLOSE
Set rs = Nothing
rsCount.CLOSE
Set rsCount = Nothing
Set db = Nothing
    
Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub
