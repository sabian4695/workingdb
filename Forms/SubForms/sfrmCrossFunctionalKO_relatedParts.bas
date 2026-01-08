Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub relatedPN_Click()
On Error GoTo Err_Handler

Form_frmCrossFunctionalKO.sfrmCrossFunctionalKO_relatedPartsClass.Form.filter = "partNumber = '" & Me.relatedPN & "'"
Form_frmCrossFunctionalKO.sfrmCrossFunctionalKO_relatedPartsClass.Form.FilterOn = True

Form_frmCrossFunctionalKO.sfrmCrossFunctionalKO_relatedPartsClass.Visible = True
Form_frmCrossFunctionalKO.lblClass.Visible = True

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub type_Click()
On Error GoTo Err_Handler

Form_frmCrossFunctionalKO.sfrmCrossFunctionalKO_relatedPartsClass.Form.filter = "partNumber = '" & Me.relatedPN & "'"
Form_frmCrossFunctionalKO.sfrmCrossFunctionalKO_relatedPartsClass.Form.FilterOn = True

Form_frmCrossFunctionalKO.sfrmCrossFunctionalKO_relatedPartsClass.Visible = True
Form_frmCrossFunctionalKO.lblClass.Visible = True

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
