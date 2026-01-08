Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub deletebtn_Click()
On Error GoTo Err_Handler

If IsNull(Me.ID) Then Exit Sub

If MsgBox("Are you sure you want to delete this record?" & vbNewLine & "You cannot undo this action.", vbYesNo, "Warning") = vbYes Then
    Call registerDRSUpdates("tblTimeTrackChild", Me.ID, Me.ActiveControl.name, "", "time deleted", Me.Control_Number)
    dbExecute "DELETE FROM [dbo_tblTimeTrackChild] WHERE [ID] = " & Me.ID
    DoCmd.Requery
End If
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_AfterInsert()
On Error GoTo Err_Handler

Call registerDRSUpdates("tblTimeTrackChild", Me.ID, Me.ActiveControl.name, "", "Time Added", Me.Control_Number)

If Form_frmDRSdashboard.Check_In_Prog = "Not Started" Then
    Form_frmDRSdashboard.Check_In_Prog = "In Progress"
    Call registerDRSUpdates("tblDRStrackerExtras", Me.Control_Number, "Check_In_Prog", "Not Started", "In Progress")
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)
Me.Associate_ID.DefaultValue = DLookup("[ID]", "[tblPermissions]", "[user] = '" & Environ("username") & "'")

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub TimeTrack_Work_Date_AfterUpdate()
On Error GoTo Err_Handler
Call registerDRSUpdates("tblTimeTrackChild", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.Control_Number)
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub TimeTrack_Work_Hours_AfterUpdate()
On Error GoTo Err_Handler
Call registerDRSUpdates("tblTimeTrackChild", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.Control_Number)
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub TimeTrack_Work_Type_AfterUpdate()
On Error GoTo Err_Handler
Call registerDRSUpdates("tblTimeTrackChild", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.Control_Number)
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
