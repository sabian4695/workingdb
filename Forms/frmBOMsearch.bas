Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub clear_Click()
On Error GoTo Err_Handler

Me.NAMsrchBox = ""
Me.NAMsrchBox.SetFocus

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Function getID() As Long
getID = 0

Dim partNum
partNum = Me.NAMsrchBox
If IsNull(partNum) Or partNum = "" Then
    MsgBox "please type something in to search..", vbInformation, "Can't find it - sorry"
    Exit Function
End If
Dim idVal
idVal = idNAM(partNum, "NAM")
If idVal = "" Then
    MsgBox "Part number not found in System Items Table. Sometimes this happens if the part isn't active yet.", vbInformation, "Sorry about that."
    Exit Function
End If

getID = idVal

End Function

Private Sub ComptSrch_Click()
On Error GoTo Err_Handler

Dim checkIt
checkIt = getID
If checkIt = 0 Then Exit Sub
Me.Form.filter = "[COMPONENT_ITEM_ID] = " & checkIt
Me.Form.FilterOn = True

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub export_Click()
On Error GoTo Err_Handler

Dim FileName As String, sqlString As String
FileName = "H:\BOMsearch_" & nowString & ".xlsx"
sqlString = "Select Org, Assy, Compt from qryBOM where " & Me.Form.filter

Call exportSQL(sqlString, FileName)
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Assysrch_Click()
On Error GoTo Err_Handler
Dim checkIt
checkIt = getID
If checkIt = 0 Then Exit Sub
Me.Form.filter = "[ASSEMBLY_ITEM_ID] = " & checkIt
Me.Form.FilterOn = True
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub btnOHQ_Click()
On Error GoTo Err_Handler

Form_DASHBOARD.partNumberSearch = Me.Compt

If CurrentProject.AllForms("frmOnHandQty").IsLoaded = True Then DoCmd.CLOSE acForm, "frmOnHandQty"

DoCmd.OpenForm "frmOnHandQty"
DoCmd.CLOSE acForm, "frmBOMsearch"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblAssy_Click()
On Error GoTo Err_Handler

Me.Assy.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblAssyDesc_Click()
On Error GoTo Err_Handler

Me.assyDescription.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblAssyStatus_Click()
On Error GoTo Err_Handler

Me.assyStatus.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblCompDesc_Click()
On Error GoTo Err_Handler

Me.compDescription.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblCompStatus_Click()
On Error GoTo Err_Handler

Me.compStatus.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblCompt_Click()
On Error GoTo Err_Handler

Me.Compt.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblImp_Click()
On Error GoTo Err_Handler

Me.IMPLEMENTATION_DATE.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblInverse_Click()
On Error GoTo Err_Handler

Me.Inverse_Qty.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblOrg_Click()
On Error GoTo Err_Handler

Me.Org.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblQty_Click()
On Error GoTo Err_Handler

Me.Qty.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
