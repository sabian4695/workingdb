Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Me.filter = "recipientUser = '" & Environ("username") & "' AND readDate is null"
Me.FilterOn = True

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Function updateFilter()

Dim mainFilt As String, var1 As String, var2 As String
var1 = ""
var2 = "recipientUser"

If Me.showReadToggle Then var1 = "not "
If Me.showSentToggle Then var2 = "senderUser"
Me.markAsRead.Enabled = Not (Me.showSentToggle Or Me.showReadToggle)

mainFilt = "readDate is " & var1 & "null AND " & var2 & " = '" & Environ("username") & "'"

Me.filter = mainFilt
Me.FilterOn = True

End Function

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Err_Handler
Call notificationsCount

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Unload", Err.DESCRIPTION, Err.number)
End Sub

Private Sub imgSender_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmUserProfile", , , "user = '" & Me.senderUser & "'"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub markAllAsRead_Click()
On Error GoTo Err_Handler

dbExecute "UPDATE tblNotificationsSP SET readDate = '" & Now() & "' WHERE recipientUser = '" & Environ("username") & "' AND readDate is null"
Me.Requery

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub markAsRead_Click()
On Error GoTo Err_Handler

If Me.recipientUser = Environ("username") Then
    If IsNull(Me.readDate) Then Me.readDate = Now()
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub notificationOpen_Click()
On Error GoTo Err_Handler

If Me.appName = "" Then Exit Sub

Select Case Me.appName
    Case "Design WO"
        If CurrentProject.AllForms("frmDRSdashboard").IsLoaded = True Then
             DoCmd.CLOSE acForm, "frmDRSdashboard"
                  On Error Resume Next
            TempVars.Add "controlNumber", Me.appId.Value
            DoCmd.OpenForm "frmDRSdashboard"
        Else
            TempVars.Add "controlNumber", Me.appId.Value
            DoCmd.OpenForm "frmDRSdashboard"
        End If
    Case "Part Project"
        openPartProject (Me.appId.Value)
    Case "Issue"
        DoCmd.OpenForm "frmPartIssues", , , "recordId = " & Me.appId.Value
    Case "Trial"
        DoCmd.OpenForm "frmPartTrialDetails", , , "recordId = " & Me.appId.Value
End Select

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub showReadToggle_Click()
On Error GoTo Err_Handler

Call updateFilter

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub showSentToggle_Click()
On Error GoTo Err_Handler

Call updateFilter

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
