Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub clear_Click()
On Error GoTo Err_Handler
Me.srchBox = ""
Me.FilterOn = False
Me.srchBox.SetFocus
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Me.srchBox = Environ("username")

Me.filter = "[USER_NAME] = '" & UCase(Me.srchBox) & "'"
Me.FilterOn = True

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub partNumSrchBtn_Click()
On Error GoTo Err_Handler

Me.filter = "[USER_NAME] = '" & UCase(Me.srchBox) & "'"
Me.FilterOn = True
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
