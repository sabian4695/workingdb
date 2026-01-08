Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnSearch_Click()
On Error GoTo Err_Handler

Dim partNum
partNum = Me.srchBox
If partNum <> "" Then partNum = idNAM(partNum, "NAM")
If partNum = "" Then Exit Sub

DoCmd.applyFilter , "[INVENTORY_ITEM_ID] = " & partNum

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

Dim FileName As String, sqlString As String
FileName = "H:\PartForecast_" & nowString & ".xlsx"
sqlString = Split(Me.RecordSource, "ORDER BY")(0)
sqlString = Left(sqlString, Len(sqlString) - 1) & " AND " & Me.filter

Call exportSQL(sqlString, FileName)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Dim partNum
partNum = Form_DASHBOARD.partNumberSearch
Me.srchBox = partNum
If partNum <> "" Then partNum = idNAM(partNum, "NAM")
If partNum = "" Then Exit Sub

DoCmd.applyFilter , "[INVENTORY_ITEM_ID] = " & partNum
Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub
