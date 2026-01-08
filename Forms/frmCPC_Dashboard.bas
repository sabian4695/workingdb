Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub emailCF_Click()
On Error GoTo Err_Handler

Dim SendItems As New clsOutlookCreateItem
Dim db As DAO.Database
Dim rs As DAO.Recordset
Dim strTo As String
Dim strSubject As String

Me.Requery

Set SendItems = New clsOutlookCreateItem
Set db = CurrentDb
Set rs = db.OpenRecordset("SELECT memberEmail FROM tblCPC_XFTeams WHERE projectId = " & Form_frmCPC_Dashboard.ID, dbOpenSnapshot)

If Not (rs.BOF And rs.EOF) Then
    rs.MoveFirst
    Do While Not rs.EOF
        strTo = strTo & rs("memberEmail") & "; "
        rs.MoveNext
    Loop
End If

If strTo = "" Then
    MsgBox "Please add members to cross functional team.", vbOKOnly + vbExclamation, "Error"
    Exit Sub
End If

strSubject = Me.projectNumber

SendItems.CreateMailItem sendTo:=strTo, subject:=strSubject

Set SendItems = Nothing
rs.CLOSE
Set rs = Nothing
Set db = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Dirty(Cancel As Integer)
On Error GoTo Err_Handler

Me.lastModified = Now

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub btnEmail_Click()
On Error GoTo Err_Handler

Dim SendItems As New clsOutlookCreateItem
Dim strTo As String
Dim strSubject As String

If IsNull(Me.contactEmail) Then
    MsgBox "Please enter a contact email.", vbOKOnly + vbExclamation, "Error"
    Exit Sub
End If

Set SendItems = New clsOutlookCreateItem

strTo = Me.contactEmail & ";"
strSubject = Me.projectNumber

SendItems.CreateMailItem sendTo:=strTo, subject:=strSubject

Set SendItems = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Dim ctrl As Control

For Each ctrl In Me.Controls
    If ctrl.ControlType = acComboBox Then
        If InStr(ctrl.name, "status") > 0 Then
            Call Format_Task(ctrl)
        End If
    End If
Next ctrl

Me.sfrmCPC_Dashboard.Form.filter = "status <> 'Closed'"
Me.sfrmCPC_Dashboard.Form.FilterOn = True

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Err_Handler
If CurrentProject.AllForms("frmCPC_WorkTracker").IsLoaded Then Form_frmCPC_WorkTracker.Requery

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Unload", Err.DESCRIPTION, Err.number)
End Sub

Private Sub projectHistory_Click()
On Error GoTo Err_Handler
DoCmd.OpenForm "frmHistory"
Form_frmHistory.RecordSource = "tblCPC_UpdateTracking"
Form_frmHistory.dataTag0.ControlSource = "dataTag0"
Form_frmHistory.filter = "dataTag0 = '" & Me.projectNumber & "'"
Form_frmHistory.FilterOn = True
Form_frmHistory.OrderBy = "updatedDate Desc"
Form_frmHistory.OrderByOn = True
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub btnClose_Click()
On Error GoTo Err_Handler

If MsgBox("Are you sure you want to close this project?", vbYesNo + vbQuestion, "Close Confirmation") = vbNo Then Exit Sub

Me.status = "Closed"
Me.dateClosed = Date
Call registerCPCUpdates("tblCPC_Projects", ID, Me.status.name, Me.status.OldValue, Me.status, Me.ID)
DoCmd.CLOSE

If CurrentProject.AllForms("frmCPC_WorkTracker").IsLoaded Then Form_frmCPC_WorkTracker.Requery

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub btnDelete_Click()
On Error GoTo Err_Handler

If MsgBox("Are you sure you want to delete this project?", vbYesNo + vbExclamation, "Delete Confirmation") = vbNo Then Exit Sub

Me.status = "Deleted"
Me.dateClosed = Date
Call registerCPCUpdates("tblCPC_Projects", ID, Me.status.name, Me.status.OldValue, Me.status, Me.ID)
DoCmd.CLOSE

If CurrentProject.AllForms("frmCPC_WorkTracker").IsLoaded Then Form_frmCPC_WorkTracker.Requery

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Format_Task(cbo As Control)
On Error GoTo Err_Handler

Dim colorVal

If cbo.Value = "N/A" Then colorVal = rgb(89, 89, 89)
If cbo.Value = "Incomplete" Then colorVal = rgb(135, 0, 0)
If cbo.Value = "In Progress" Then colorVal = rgb(135, 135, 0)
If cbo.Value = "Complete" Then colorVal = rgb(0, 135, 0)

cbo.BorderColor = colorVal

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Function CPC_trackUpdate()
On Error GoTo Err_Handler
Call registerCPCUpdates("tblCPC_Projects", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.ID)
Exit Function
Err_Handler:
    Call handleError(Me.name, "CPC_trackUpdate", Err.DESCRIPTION, Err.number)
End Function

Function CPC_trackUpdate_Task()
On Error GoTo Err_Handler
Call CPC_trackUpdate
Call Format_Task(Me.ActiveControl)
Exit Function
Err_Handler:
    Call handleError(Me.name, "CPC_trackUpdate_Task", Err.DESCRIPTION, Err.number)
End Function
