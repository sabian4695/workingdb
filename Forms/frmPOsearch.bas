Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnSearch_Click()
On Error GoTo Err_Handler

Dim CapNum
CapNum = Me.srchBox
If CapNum = "" Then Exit Sub
DoCmd.applyFilter , "[capNum] = '" & CapNum & "'"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub clear_Click()
On Error GoTo Err_Handler
Me.srchBox.SetFocus
Me.srchBox = ""
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub export_Click()
On Error GoTo Err_Handler

Dim FileName As String, sqlString As String, filt As String
FileName = "H:\POs_" & nowString & ".xlsx"
filt = " WHERE " & Replace(Me.Form.filter, "[capNum]", "PO_PO_REQUISITION_HEADERS_ALL.ATTRIBUTE3")
If Me.FilterOn = False Then filt = ""
sqlString = Left(Me.RecordSource, Len(Me.RecordSource) - 1) & filt

Call exportSQL(sqlString, FileName)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
