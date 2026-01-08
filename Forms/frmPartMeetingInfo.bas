Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub dateOfMeeting_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartMeetings", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNum, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

showChecklist

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub meetingHistory_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmHistory", acNormal, , "[tableName] = 'tblPartMeetings' AND [tableRecordId] = " & Me.recordId

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub meetingNotes_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartMeetings", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNum, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Function showChecklist()
On Error GoTo Err_Handler

Dim checkListB As Boolean
checkListB = DCount("recordId", "tblPartMeetingTemplates", "meetingType = " & Me.meetingType) > 0
Me.tabCtl.Pages("checkItems").Visible = checkListB 'only show checklist if there are items

Exit Function
Err_Handler:
    Call handleError(Me.name, "showChecklist", Err.DESCRIPTION, Err.number)
End Function

Private Sub meetingType_AfterUpdate()
On Error GoTo Err_Handler

Dim typeName As String
typeName = Me.meetingType.column(1)

If DCount("recordId", "tblPartMeetingTemplates", "meetingType = " & Me.meetingType) = 0 Then GoTo skipChecklist 'if there is a checklist template for this meeting type, run the check / add items if necessary

If DCount("recordId", "tblPartMeetingInfo", "meetingId = " & Me.recordId) <> 0 Then 'check if current checklist exists, and stop if they don't want to erase
    If MsgBox("Are you sure? This adds the " & typeName & " Check Items and erases current check items", vbYesNo, "Please confirm") <> vbYes Then
        Me.Undo
        Exit Sub
    End If
End If

Call createMeetingCheckItems(Me.meetingType, Me.recordId)
If Me.Dirty Then Me.Dirty = False
Call registerPartUpdates("tblPartMeetings", Me.recordId, "Check Items", "", "Added", Me.partNum, Me.name, Me.meetingType.column(1))
Me.sfrmPartMeetingChecklist.Requery

skipChecklist:
showChecklist
Call registerPartUpdates("tblPartMeetings", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNum, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub remove_Click()
On Error GoTo Err_Handler

Dim db As Database
Set db = CurrentDb()

If MsgBox("Are you sure you want to delete this Meeting?", vbYesNo, "Please confirm") = vbYes Then
    Call registerPartUpdates("tblPartMeetings", Me.recordId, "Meeting", Nz(Me.meetingType.column(1)), "Deleted", Nz(Me.TpartNumber), Me.name)
    db.Execute ("DELETE FROM tblPartMeetingAttendees WHERE [meetingId] = " & Me.recordId)
    db.Execute ("DELETE FROM tblPartMeetings WHERE [recordId] = " & Me.recordId)
    DoCmd.CLOSE
    If CurrentProject.AllForms("frmPartMeetings").IsLoaded Then Form_frmPartMeetings.Requery
End If

Set db = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub sendNotes_Click()
On Error GoTo Err_Handler

Dim SendItems As New clsOutlookCreateItem               ' outlook class
Dim strTo As String                                     ' email recipient
Dim strSubject As String, strNotes As String

Set SendItems = New clsOutlookCreateItem

Dim db As Database
Set db = CurrentDb()
Dim rs2 As Recordset, rs3 As Recordset
Set rs2 = db.OpenRecordset("SELECT * FROM tblPartMeetingAttendees WHERE meetingId = " & Me.recordId, dbOpenSnapshot)
strTo = ""

Dim pplStr As String
pplStr = "Attendees: "
Do While Not rs2.EOF
    Set rs3 = db.OpenRecordset("SELECT * from tblPermissions WHERE user = '" & rs2!attendeeUsername & "'")
    pplStr = pplStr & "<br/>" & rs3!firstName & " " & rs3!lastName
    
    If rs2!attendeeUsername = Environ("username") Then GoTo nextOne
    
    strTo = strTo & rs3!userEmail & "; "
nextOne:
    rs2.MoveNext
Loop

Dim rsMeetingItems As Recordset, meetingItems As String
Set rsMeetingItems = db.OpenRecordset("SELECT * FROM tblPartMeetingInfo WHERE checkItem is not null AND meetingId = " & Me.recordId)

meetingItems = ""

If rsMeetingItems.RecordCount > 0 Then

    meetingItems = "<br/><table><tbody>"
    meetingItems = meetingItems & "<tr><th>Review Item</th><th>Response</th><th>Comments</th></tr>"
    
    Do While Not rsMeetingItems.EOF
        meetingItems = meetingItems & "<tr>" & _
            "<td>" & rsMeetingItems!checkItem & "</td>"
            
        Select Case rsMeetingItems!checkResponse
            Case 1 'OK
                meetingItems = meetingItems & "<td style=""color: rgb(73,155, 73); font-weight: 400"">OK</td>"
            Case 2 'NG
                meetingItems = meetingItems & "<td style=""color: rgb(80, 20, 20); font-weight: 400; background: #EE9999;"">NG</td>"
            Case 3 'N/A
                meetingItems = meetingItems & "<td style=""color: rgb(73, 73, 73); font-weight: 400"">N/A</td>"
        End Select
        
        meetingItems = meetingItems & "<td>" & rsMeetingItems!checkComments & "</td>" & _
            "</tr>"
            
        rsMeetingItems.MoveNext
    Loop
    meetingItems = meetingItems & "</tbody></table>"
End If

strSubject = Me.partNum & " " & Me.meetingType.column(1) & " on " & Me.dateOfMeeting
strNotes = Me.meetingType.column(1) & " Meeting Notes: <br/>" & Replace(Nz(Me.meetingNotes, ""), vbNewLine, "<br/>")
strNotes = strNotes & "<br/>" & meetingItems
strNotes = strNotes & "<br/><br/>" & pplStr

SendItems.CreateMailItem sendTo:=strTo, _
                         subject:=strSubject, _
                         htmlBody:=strNotes
Set SendItems = Nothing

Set db = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
