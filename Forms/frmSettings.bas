Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim dbLoc, fso

Function checkIfAdminDev() As Boolean

checkIfAdminDev = False

Dim errorTxt As String: errorTxt = ""
If (privilege("admin") = False) Then errorTxt = "You need admin privilege to do this"
If (privilege("developer") = False) Then errorTxt = "You need developer privilege to do this"

If errorTxt <> "" Then
    MsgBox errorTxt, vbCritical, "Access Denied"
    checkIfAdminDev = False
    Exit Function
End If

checkIfAdminDev = True

End Function

Private Sub disShift_Click()
On Error GoTo Err_Handler
If Not checkIfAdminDev Then Exit Sub
ap_DisableShift
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub enableShift_Click()
On Error GoTo Err_Handler
If Not checkIfAdminDev Then Exit Sub
ap_EnableShift
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub hideNav_Click()
On Error GoTo Err_Handler
Call DoCmd.NavigateTo("acNavigationCategoryObjectType")
Call DoCmd.RunCommand(acCmdWindowHide)
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub hideRibbon_Click()
On Error GoTo Err_Handler

DoCmd.ShowToolbar "Ribbon", acToolbarNo

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub openSettings_Click()
On Error GoTo Err_Handler

openPath ("\\data\mdbdata\WorkingDB\Batch\Working DB SETTINGS.lnk")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub showRibbon_Click()
On Error GoTo Err_Handler

DoCmd.ShowToolbar "Ribbon", acToolbarYes

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub showNav_Click()
On Error GoTo Err_Handler

Call DoCmd.SelectObject(acTable, , True)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Dim dev As Boolean
dev = privilege("Developer") 'dev dashboard

Me.enableShift.Visible = dev
Me.disShift.Visible = dev
Me.showNav.Visible = dev
Me.showRibbon.Visible = dev
Me.hideNav.Visible = dev
Me.hideRibbon.Visible = dev

If privilege("Edit") Then
    Me.openSettings.Enabled = True
    Me.openSettings.Caption = " Open Settings App"
Else
    Me.openSettings.Enabled = False
    Me.openSettings.Caption = " Open Settings App - NEED EDIT PRIVILEGE"
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub plmSettings_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmPLMsettings"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
