Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub backBtn_Click()
On Error GoTo Err_Handler

Form_DASHBOARD.sfrmCalendarView.Visible = True
Form_DASHBOARD.sfrmCalendarView.SetFocus
Form_DASHBOARD.sfrmCalendarItems.Visible = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setSplashLoading("Building calendar items...")

Call setTheme(Me)

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub openItem_Click()
On Error GoTo Err_Handler

Select Case Me.type
    Case "Design WO"
        If CurrentProject.AllForms("frmDRSdashboard").IsLoaded = True Then DoCmd.CLOSE acForm, "frmDRSdashboard"
        TempVars.Add "controlNumber", Me.ID.Value
        DoCmd.OpenForm "frmDRSdashboard"
    Case "Step Approval", "Step"
        openPartProject (Me.ID)
    Case "Part Issue"
        DoCmd.OpenForm "frmPartIssues", , , "recordId = " & Me.ID
End Select

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
