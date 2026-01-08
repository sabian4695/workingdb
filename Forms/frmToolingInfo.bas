Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub clear_Click()
On Error GoTo Err_Handler
Me.NAMsrchBox.SetFocus
Me.NAMsrchBox = ""
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Me.NAMsrchBox = Form_DASHBOARD.partNumberSearch
Call NAMsrch_Click

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub NAMsrch_Click()
On Error GoTo Err_Handler

Dim toolNum
toolNum = Me.NAMsrchBox
If toolNum = "" Then Exit Sub

If Right(UCase(toolNum), 1) <> "T" Then toolNum = toolNum & "T"

DoCmd.applyFilter , "[Tool Number] = '" & UCase(toolNum) & "'"

If Me.RecordsetClone.RecordCount = 0 Then
        Me.FilterOn = False
        MsgBox "There are no records" & vbCrLf & "that match this filter.", vbInformation + vbOKOnly, "No records returned"
        Exit Sub
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub wdbMoldingInfo_Click()
On Error GoTo Err_Handler

Dim toolNum
toolNum = Me.NAMsrchBox
If toolNum = "" Then Exit Sub

If Right(toolNum, 1) <> "T" Then toolNum = toolNum & "T"

DoCmd.OpenForm "frmPartMoldingInfo", , , "toolNumber = '" & toolNum & "'"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
