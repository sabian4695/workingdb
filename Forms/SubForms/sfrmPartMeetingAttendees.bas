Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub addCFteam_Click()
On Error GoTo Err_Handler

If MsgBox("This adds all members from the CF team to this meeting - are you sure?", vbYesNo, "Are you sure?") = vbYes Then
    Dim partNum As String
    partNum = Me.lblPerson.tag
    
    Dim team() As String, ITEM
    team = Split(grabPartTeam(partNum, False, False, True), ",")
    
    For Each ITEM In team
        DoCmd.GoToRecord , , acNewRec
        Me.attendeeUsername = ITEM
    Next ITEM
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub remove_Click()
On Error GoTo Err_Handler

Me.Dirty = False

If IsNull(Me.recordId) Then
    MsgBox "This is an empty record, don't worry about deleting this.", vbInformation, "Can't do that"
    Exit Sub
End If

If MsgBox("Are you sure you want to delete this?", vbYesNo, "Please confirm") <> vbYes Then Exit Sub

dbExecute ("DELETE FROM tblPartMeetingAttendees WHERE [recordId] = " & Me.recordId)
Me.Requery
MsgBox "Attendee Deleted", vbOKOnly, "Deleted"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
