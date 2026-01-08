Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnCategories_Click()
On Error GoTo Err_Handler

Form_DASHBOARD.partNumberSearch = Me.NAM
If CurrentProject.AllForms("frmItemCategories").IsLoaded = False Then
    DoCmd.CLOSE acForm, "frmItemCategories"
End If

Dim filterVal, pNum
filterVal = Form_DASHBOARD.partNumberSearch
If filterVal = "" Or IsNull(filterVal) Then filterVal = "29123"
filterVal = "[INVENTORY_ITEM_ID] = " & idNAM(filterVal, "NAM")

DoCmd.OpenForm "frmItemCategories", , , filterVal
Form_frmItemCategories.NAMsrchBox = Nz(Form_DASHBOARD.partNumberSearch)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Custsrch_Click()
On Error GoTo Err_Handler

Me.Form.filter = "[CUSTOMER_ITEM_NUMBER] LIKE '" & Me.CustsrchBox & "'"
Me.Form.FilterOn = True

If Me.RecordsetClone.RecordCount = 0 Then
        MsgBox "No Results.", vbInformation + vbOKOnly, "No records returned"
        Exit Sub
End If
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub CustsrchBox_GotFocus()
On Error GoTo Err_Handler

Me.NAMsrch.default = False
Me.Custsrch.default = True

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub NAMsrch_Click()
On Error GoTo Err_Handler

Me.Form.filter = "[NAM] = '" & Me.NAMsrchBox & "'"
Me.Form.FilterOn = True

If Me.RecordsetClone.RecordCount = 0 Then
    MsgBox "No Results.", vbInformation + vbOKOnly, "No records returned"
    Exit Sub
End If
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub NAMsrchBox_GotFocus()
On Error GoTo Err_Handler

Me.Custsrch.default = False
Me.NAMsrch.default = True

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
