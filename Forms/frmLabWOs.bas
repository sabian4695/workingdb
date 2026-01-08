Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Current()
On Error GoTo Err_Handler

Me.pgDimensional.Visible = Me.Dms.Value > 0
Me.pgForce.Visible = Me.Frc.Value > 0
Me.pgEnvironmental.Visible = Me.Env.Value > 0

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub btnSearch_Click()
On Error GoTo Err_Handler

Dim strResponse As String

Me.txtSearch.SetFocus
Me.filter = "WONumber = " & Replace(Me.txtSearch, "N", "")
Me.FilterOn = True

If Me.RecordsetClone.RecordCount = 0 Then
    DoCmd.applyFilter , "WONumber = 1002"
    strResponse = "There are no records" & vbCrLf & "that match this filter."
    MsgBox strResponse, vbInformation + vbOKOnly, "No records returned"
    Exit Sub
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub clear_Click()
On Error GoTo Err_Handler

Me.txtSearch = ""
Me.txtSearch.SetFocus

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub btnSearchWO_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmLabWO_Search"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Me.txtSearch.SetFocus

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub
