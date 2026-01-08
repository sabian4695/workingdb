Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnDetails_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmPackagingDetails", , , "partNumber = '" & Me.partNumber & "'"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub btnRefresh_Click()
On Error GoTo Err_Handler

Me.Requery

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub export_Click()
On Error GoTo Err_Handler

Dim FileName As String, sqlString As String, filt As String
FileName = "H:\Packaging_Tracker_" & nowString & ".xlsx"
filt = " WHERE " & Me.Form.filter
If Me.FilterOn = False Then filt = ""

sqlString = "SELECT * FROM qryPackagingTrackerExport " & filt

Call exportSQL(sqlString, FileName)

Exit Sub
Err_Handler:
Dim db As Database
Set db = CurrentDb()
db.QueryDefs.Delete "myExportQueryDef"
Set db = Nothing
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Function clickFilter(controlName As String)
On Error GoTo Err_Handler

Me.Controls(controlName).SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Function
Err_Handler:
    Call handleError(Me.name, controlName, Err.DESCRIPTION, Err.number)
End Function

Function applyTheFilters()
On Error GoTo Err_Handler

Dim filt
filt = "partNumber IN (SELECT PN FROM qryPackagingPartNumbers)"

If Me.fltUser <> "" Then filt = filt & " AND (partNumber IN (SELECT partNumber FROM tblPartTeam WHERE person = '" & Me.fltUser & "'))"
If Me.fltModel <> "" Then filt = filt & " AND modelCode = '" & Me.fltModel & "'"

'Select Case Me.byPackStatus
'    Case "Closed"
'        'pack test status
'        filt = filt & " AND (packTestStatusCalc = 'N/A' or packTestStatusCalc = 'Closed' or packTestStatusCalc = 'Not Found')"
'        'NPIF status
'        'filt = filt & " AND NPIFstatusCalc = 'Uploaded'"
'        'customer approval status
'        'filt = filt & " AND (customerPackApproval = 'Not Found' or customerPackApproval = 'Closed')"
'        'fit trial status
'        'filt = filt & " AND (fitTrialStatus = 'N/A' or fitTrialStatus = 'Complete')"
'    Case "Open"
'        'pack test status
'        filt = filt & " AND NOT ((packTestStatusCalc = 'N/A' or packTestStatusCalc = 'Closed' or packTestStatusCalc = 'Not Found')"
'        'NPIF status
'        filt = filt & " AND (NPIFstatusCalc = 'Not Found' or NPIFstatusCalc = 'Uploaded')"
'        'customer approval status
'        filt = filt & " AND (customerPackApproval = 'Not Found' or customerPackApproval = 'Closed')"
'        'fit trial status
'        filt = filt & " AND (fitTrialStatus = 'N/A' or fitTrialStatus = 'Complete'))"
'End Select

If Nz(Me.byProjectStatus, 0) <> 0 Then
    filt = filt & " AND (partNumber IN (SELECT partNumber FROM tblPartProject WHERE projectStatus = " & Me.byProjectStatus & "))"
End If

Me.filter = filt
Me.FilterOn = filt <> ""

Exit Function
Err_Handler:
    Call handleError(Me.name, "applyTheFilters", Err.DESCRIPTION, Err.number)
End Function

Private Sub NPIF_Click()
On Error GoTo Err_Handler

Dim mainPath As String, adFilter As String
mainPath = mainFolder(Me.ActiveControl.name)

adFilter = "?FilterField1=Part%5Fx0020%5FNumber&FilterValue1=" & Me.partNumber

openPath (mainPath & adFilter)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub openDetails_Click()
On Error GoTo Err_Handler

openPartProject (Me.partNumber)
DoCmd.CLOSE acForm, "frmPartPackagingDetails"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
