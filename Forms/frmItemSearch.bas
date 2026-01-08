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

Me.ShortcutMenu = True

Me.Form.FilterOn = False
Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub ITEM_TYPE_Click()
    Me.ITEM_TYPE.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
End Sub

Private Sub lblDescription_Click()
    Me.DESCRIPTION.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
End Sub

Private Sub lblNAM_Click()
    Me.NAM.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
End Sub

Private Sub lblStatus_Click()
    Me.INVENTORY_ITEM_STATUS_CODE.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
End Sub

Private Sub srch_Click()
On Error GoTo Err_Handler

Me.Form.filter = "[DESCRIPTION] LIKE '*" & UCase(Me.srchBox) & "*'"
Me.Form.FilterOn = True
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
