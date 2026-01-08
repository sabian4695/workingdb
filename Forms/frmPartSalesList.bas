Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub export_Click()
On Error GoTo Err_Handler

Dim FileName As String
FileName = "H:\Part_Sales_List_" & nowString & ".xlsx"
DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, "qrySales", FileName, True
MsgBox "Export Complete. File path: " & FileName, vbOKOnly, "Notice"
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Function clickLabel(fieldName As String)
On Error Resume Next

Me(fieldName).SetFocus
DoCmd.RunCommand acCmdFilterMenu
End Function

Public Sub refresh_Click()
On Error GoTo Err_Handler
Me.Requery
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
