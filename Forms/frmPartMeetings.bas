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

Private Sub Form_Current()
On Error GoTo Err_Handler

Form_sfrmPartMeetingAttendees.lblPerson.tag = Me.partNum

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Me.OrderBy = "dateOfMeeting Desc"
Me.OrderByOn = True

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblDescription_Click()
On Error GoTo Err_Handler

Me.meetingType.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblMeetingDate_Click()
On Error GoTo Err_Handler

Me.dateOfMeeting.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblNotes_Click()
On Error GoTo Err_Handler

Me.meetingNotes.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblPN_Click()
On Error GoTo Err_Handler

Me.partNum.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
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
Call registerPartUpdates("tblPartMeetings", Me.recordId, "Check Items", "", "Added", Me.partNum, Me.name, Me.meetingType.column(1))

skipChecklist:
If Me.Dirty Then Me.Dirty = False
Call registerPartUpdates("tblPartMeetings", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNum, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub newPartMeeting_Click()
On Error GoTo Err_Handler

Dim db As Database
Set db = CurrentDb()

db.Execute "INSERT INTO tblPartMeetings(partNum,partProjectId,partStepId,dateOfMeeting) VALUES " & _
                                    "('" & Me.TpartNumber & "'," & Me.TprojectId & "," & Nz(Me.TstepId, "Null") & ",Date());"
TempVars.Add "meetingId", db.OpenRecordset("SELECT @@identity")(0).Value
Call registerPartUpdates("tblPartMeetings", TempVars!meetingId, "Meeting Creation", "", "Created", Me.TpartNumber, Me.name)
Me.Requery

Set db = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub openDetails_Click()
On Error GoTo Err_Handler

Select Case Nz(Me.meetingType, 0)
    Case 1 'CFKO
        DoCmd.OpenForm "frmCrossFunctionalKO", , , "meetId = " & Me.recordId
        Form_frmCrossFunctionalKO.allowEdits = False
        Form_frmCrossFunctionalKO.copyPI.Enabled = False
        Exit Sub
    Case Else
        Dim typeName As String
        typeName = Me.meetingType.column(1)
        
        If DCount("recordId", "tblPartMeetingTemplates", "meetingType = " & Me.meetingType) = 0 Then GoTo skipChecklist 'if there is a checklist template for this meeting type, run the check / add items if necessary
        If DCount("recordId", "tblPartMeetingInfo", "meetingId = " & Me.recordId) > 0 Then GoTo skipChecklist 'check if current checklist exists, and stop if they don't want to erase
        
        Call createMeetingCheckItems(Me.meetingType, Me.recordId)
        If Me.Dirty Then Me.Dirty = False
        Call registerPartUpdates("tblPartMeetings", Me.recordId, "Check Items", "", "Added", Me.partNum, Me.name, Me.meetingType.column(1))
End Select

skipChecklist:
DoCmd.OpenForm "frmPartMeetingInfo", , , "recordId = " & Me.recordId

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Public Sub refresh_Click()
On Error GoTo Err_Handler
Me.Requery
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
    Me.Requery
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
    pplStr = pplStr & vbNewLine & rs3!firstName & " " & rs3!lastName
    
    If rs2!attendeeUsername = Environ("username") Then GoTo nextOne
    
    strTo = strTo & rs3!userEmail & "; "
nextOne:
    rs2.MoveNext
Loop

strSubject = Me.partNum & " " & Me.meetingType.column(1) & " on " & Me.dateOfMeeting
strNotes = Me.meetingType.column(1) & " Meeting Notes: " & vbNewLine & Nz(Me.meetingNotes, "")
strNotes = strNotes & vbNewLine & vbNewLine & pplStr

SendItems.CreateMailItem sendTo:=strTo, _
                         subject:=strSubject, _
                         body:=strNotes
Set SendItems = Nothing

Set db = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
