Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Function frmAfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartContacts", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber)

Exit Function
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Function

Private Sub contactHistory_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmHistory", acNormal, , "[tableName] = 'tblPartContacts' AND [partNumber] = '" & Form_frmNMQDashboard.partNumber & "'"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub deletebtn_Click()
On Error GoTo Err_Handler

If IsNull(Me.recordId) Then Exit Sub

If MsgBox("Are you sure you want to delete this record?" & vbNewLine & "You cannot undo this action.", vbYesNo, "Warning") = vbYes Then
    Call registerPartUpdates("tblPartContacts", Me.recordId, Me.ActiveControl.name, "", "Deleted", Me.partNumber)
    dbExecute "DELETE FROM [tblPartContacts] WHERE [recordId] = " & Me.recordId
    DoCmd.Requery
End If
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_AfterInsert()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartContacts", Me.recordId, Me.ActiveControl.name, "", "Contact Added", Me.partNumber)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
