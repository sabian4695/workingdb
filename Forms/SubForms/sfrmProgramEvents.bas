Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub correlatedGate_AfterUpdate()
On Error GoTo Err_Handler
Call registerPartUpdates("tblProgramEvents", Me.programId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "")
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub dataSubmitted_AfterUpdate()
On Error GoTo Err_Handler
Call registerPartUpdates("tblProgramEvents", Me.programId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "")
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub dataSubmittedDate_AfterUpdate()
On Error GoTo Err_Handler
Call registerPartUpdates("tblProgramEvents", Me.programId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "")
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub deletebtn_Click()
On Error GoTo Err_Handler

If IsNull(Me.ID) Then
    MsgBox "This is an empty record.", vbInformation, "Can't do that"
    Exit Sub
End If

If MsgBox("Are you sure you want to delete this?", vbYesNo, "Please confirm") = vbYes Then
    Call registerPartUpdates("tblProgramEvents", Me.programId, "DELETE", Me.eventTitle, "DELETED", "")
    dbExecute "DELETE FROM tblProgramEvents WHERE ID = " & Me.ID
    Me.Requery
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub eventTitle_AfterUpdate()
On Error GoTo Err_Handler
Call registerPartUpdates("tblProgramEvents", Me.programId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "")
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub eventDate_AfterUpdate()
On Error GoTo Err_Handler
Call registerPartUpdates("tblProgramEvents", Me.programId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "")
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub eventType_AfterUpdate()
On Error GoTo Err_Handler
Call registerPartUpdates("tblProgramEvents", Me.programId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "")
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
