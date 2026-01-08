Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnSearch_Click()
On Error GoTo Err_Handler

Dim partNum
partNum = Me.srchBox
If partNum = "" Then Exit Sub
DoCmd.applyFilter , "[Item] = '" & partNum & "'"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub clear_Click()
On Error GoTo Err_Handler
Me.srchBox.SetFocus
Me.srchBox = ""
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
If partNum = "" Then Exit Sub
DoCmd.applyFilter , "[Item] = '" & partNum & "'"

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub
