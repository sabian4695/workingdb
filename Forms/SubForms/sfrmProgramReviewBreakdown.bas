Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnEditImage_Click()
On Error GoTo Err_Handler

Form_DASHBOARD.partNumberSearch = Me.partNumber
DoCmd.OpenForm "frmPartPicture"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblDescription_Click()
    Me.DESCRIPTION.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
End Sub

Private Sub lblPE_Click()
    Me.PE.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
End Sub

Private Sub lblPN_Click()
    Me.partNumber.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
End Sub

Private Sub lblType_Click()
    Me.partType.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
End Sub

Private Sub lblUnit_Click()
    Me.MP_Unit.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
End Sub

Private Sub openDetails_Click()
On Error GoTo Err_Handler

openPartProject (Me.partNumber)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
