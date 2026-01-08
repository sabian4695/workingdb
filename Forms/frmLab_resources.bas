Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub details_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmLab_resource_details", acNormal, , "recordId = " & Me.recordId

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub export_Click()
On Error GoTo Err_Handler

Dim FileName As String, sqlString As String, filt As String
FileName = "H:\Lab_resources_" & nowString & ".xlsx"
filt = ""
If Me.Form.filter <> "" And Me.Form.FilterOn Then filt = " WHERE " & Me.Form.filter
sqlString = Me.RecordSource & filt

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

Private Sub labHelp_Click()
On Error GoTo Err_Handler

Call openPath(mainFolder(Me.ActiveControl.name))

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblName_Click()
On Error GoTo Err_Handler

Me.resourcename.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblType_Click()
On Error GoTo Err_Handler

Me.lab_resource_type.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub newResource_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmLab_resource_details", acNormal, , , acFormAdd

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Public Sub refresh_Click()
On Error GoTo Err_Handler
Me.Requery
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
