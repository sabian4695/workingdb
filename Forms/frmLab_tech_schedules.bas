Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub details_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmLab_tech_schedule_details", acNormal, , "recordId = " & Me.recordId

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub export_Click()
On Error GoTo Err_Handler

Dim FileName As String, sqlString As String, filt As String
FileName = "H:\Lab_tech_schedules_" & nowString & ".xlsx"
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

Private Sub lblFri_Click()
On Error GoTo Err_Handler

Me.frihours.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblMon_Click()
On Error GoTo Err_Handler

Me.monhours.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblName_Click()
On Error GoTo Err_Handler

Me.userName.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblSat_Click()
On Error GoTo Err_Handler

Me.sathours.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblSun_Click()
On Error GoTo Err_Handler

Me.sunhours.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblThu_Click()
On Error GoTo Err_Handler

Me.thuhours.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblTue_Click()
On Error GoTo Err_Handler

Me.tuehours.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblWed_Click()
On Error GoTo Err_Handler

Me.wedhours.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub newAlteration_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmLab_tech_schedule_alteration_create", acNormal, , , acFormAdd

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub newTech_Click()
On Error GoTo Err_Handler

Dim db As Database
Set db = CurrentDb()

db.Execute "INSERT INTO tbllab_tech_schedule_template(username) VALUES('')"
TempVars.Add "techSchedule", db.OpenRecordset("SELECT @@identity")(0).Value

Set db = Nothing

Call registerLabUpdates("tbllab_tech_schedule_template", TempVars!techSchedule, "Tech Schedule", "", "Created", "", Me.name)

Me.Requery

DoCmd.OpenForm "frmLab_tech_schedule_details", acNormal, , "recordId = " & TempVars!techSchedule

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
