Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

DoCmd.GoToRecord , , acNewRec
Me.noteText.SetFocus

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Current()
On Error GoTo Err_Handler

Dim name As String

name = getFullName()

If Nz(name) = "" Then
    Exit Sub
End If

Me.createdBy.DefaultValue = "'" & name & "'"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Dirty(Cancel As Integer)
On Error GoTo Err_Handler

Me.Parent.lastModified = Now

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub btnDeleteNote_Click()
On Error GoTo Err_Handler

Dim ID As Long

If IsNull(Me.ID) Then Exit Sub

If MsgBox("Are you sure you want to delete this note?", vbYesNo, "Warning") = vbYes Then
    ID = Me.ID
    Call registerCPCUpdates("tblCPC_Notes", ID, Me.noteText.name, Me.noteText, "Deleted", Form_frmCPC_Dashboard.ID)
    DoCmd.GoToRecord , , acNewRec
    dbExecute ("DELETE * FROM tblCPC_Notes WHERE [id] = " & ID)
    Me.Requery
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub noteText_AfterUpdate()
On Error GoTo Err_Handler
Call registerCPCUpdates("tblCPC_Notes", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Form_frmCPC_Dashboard.ID)
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
