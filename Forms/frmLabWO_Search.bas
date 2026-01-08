Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnOpenWO_Click()
On Error GoTo Err_Handler

If CurrentProject.AllForms("frmLabWOs").IsLoaded = False Then
    DoCmd.OpenForm "frmLabWOs"
End If

Form_frmLabWOs.txtSearch = "N" & Me.WONumber

Form_frmLabWOs.Form.filter = "WONumber = " & Replace(Form_frmLabWOs.txtSearch, "N", "")
Form_frmLabWOs.Form.FilterOn = True

DoCmd.CLOSE acForm, "frmLabWO_Search"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub btnSearchPN_Click()
On Error GoTo Err_Handler

Me.filter = "[PartNumber] Like '*" & Me.txtSearch & "*'"
Me.FilterOn = True

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub btnSearchRequester_Click()
On Error GoTo Err_Handler

Me.filter = "[Requestor] Like '*" & Me.txtSearch & "*'"
Me.FilterOn = True

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub clear_Click()
On Error GoTo Err_Handler
Me.txtSearch.SetFocus
Me.txtSearch = ""
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Me.txtSearch = Form_DASHBOARD.partNumberSearch
Call btnSearchPN_Click

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
