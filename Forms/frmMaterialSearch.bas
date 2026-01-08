Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub export_Click()
On Error GoTo Err_Handler

Dim FileName As String
FileName = "H:\materialSearch" & nowString & ".xlsx"
DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, "qryMaterialSearch", FileName, True
MsgBox "Export Complete. File path: " & FileName, vbOKOnly, "Notice"
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblColor_Click()
    Me.Color.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
End Sub

Private Sub lblGrade_Click()
    Me.Grade.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
End Sub

Private Sub lblIndicator_Click()
    Me.Indicator.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
End Sub

Private Sub lblMaterial_Click()
    Me.Material.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
End Sub

Private Sub lblNAM_Click()
    Me.NAM.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
End Sub

Private Sub lblOrg_Click()
    Me.Org.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
End Sub

Private Sub lblSegment5_Click()
    Me.Segment5.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
End Sub

Private Sub lblStatus_Click()
    Me.INVENTORY_ITEM_STATUS_CODE.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
End Sub

Private Sub lblType_Click()
    Me.type.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
End Sub

Private Sub ohq_Click()
On Error GoTo Err_Handler
Form_DASHBOARD.partNumberSearch = Me.NAM

If CurrentProject.AllForms("frmOnHandQty").IsLoaded = True Then
    DoCmd.CLOSE acForm, "frmOnHandQty"
    DoCmd.OpenForm "frmOnHandQty"
End If

DoCmd.OpenForm "frmOnHandQty"

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub
