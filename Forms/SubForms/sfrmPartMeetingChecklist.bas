Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnResponse_Click()
On Error GoTo Err_Handler

Dim topVal As Long, oldVal As Long, oldValName As String
topVal = DMax("recordid", "tblDropDownsSP", "meetingItemResponse is not null")

oldVal = Me.checkResponse
oldValName = Nz(Me.checkResponse.column(1), "")

Select Case Me.checkResponse
    Case topVal
        Me.checkResponse = 0
    Case Else
        Me.checkResponse = Me.checkResponse + 1
End Select
    
Me.checkResponse.Requery

Call registerPartUpdates("tblPartMeetingInfo", Me.meetingId, "checkResponse", oldValName, Me.checkResponse.column(1), Form_frmPartMeetingInfo.partNum, Me.name, Me.checkItem)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub checkComments_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartMeetingInfo", Me.meetingId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Form_frmPartMeetingInfo.partNum, Me.name, Me.checkItem)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub history_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmHistory", acNormal, , "[tableName] = 'tblPartMeetingInfo' AND [tableRecordId] = " & Me.meetingId

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
