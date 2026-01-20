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

Private Sub export_Click()
On Error GoTo Err_Handler

Dim FileName As String, sqlString As String, filt As String
FileName = "H:\Item_Categories_" & nowString & ".xlsx"
filt = ""
If Me.Form.filter <> "" And Me.Form.FilterOn Then filt = " WHERE " & Me.Form.filter
sqlString = Replace(Me.RecordSource, ";", "") & filt
sqlString = Replace(sqlString, "_frmItemCategories", "frmItemCategories")

Call exportSQL(sqlString, FileName)

Exit Sub
Err_Handler:
Dim db As Database
Set db = CurrentDb()
db.QueryDefs.Delete "myExportQueryDef"
Set db = Nothing
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

DoCmd.applyFilter , "[PN] = '" & partNum & "'"
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
