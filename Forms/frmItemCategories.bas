Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnClass_Click()
On Error GoTo Err_Handler

Call openPath(mainFolder("catalog"))

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblCatType_Click()
    Me.category_type.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
End Sub

Private Sub lblSeg1_Click()
    Me.SEGMENT1.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
End Sub

Private Sub lblSeg2_Click()
    Me.SEGMENT2.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
End Sub

Private Sub lblSeg3_Click()
    Me.SEGMENT3.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
End Sub

Private Sub lblSeg4_Click()
    Me.SEGMENT4.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
End Sub

Private Sub NAMsrch_Click()
On Error GoTo Err_Handler
Dim partNum
partNum = Me.NAMsrchBox
If partNum <> "" Then partNum = idNAM(partNum, "NAM")
If partNum = "" Then Exit Sub

DoCmd.applyFilter , "[INVENTORY_ITEM_ID] = " & partNum
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub clear_Click()
On Error GoTo Err_Handler
Me.NAMsrchBox.SetFocus
Me.NAMsrchBox = ""
Me.FilterOn = False
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
