Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub clear_Click()
On Error GoTo Err_Handler
Me.srchBox = ""
Me.Form.FilterOn = False
Me.srchBox.SetFocus
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Dim partNum
partNum = Form_DASHBOARD.partNumberSearch
Me.srchBox = partNum
Me.filter = "[Tool] = '" & Me.srchBox & "'"
Me.FilterOn = True
Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub partNumSrchBtn_Click()
On Error GoTo Err_Handler

Me.filter = "[Tool] = '" & Me.srchBox & "'"
Me.FilterOn = True
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
