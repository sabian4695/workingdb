Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub noteHistory_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmHistory"
Form_frmHistory.RecordSource = "qryWdbUpdateTracking"
Form_frmHistory.dataTag0.ControlSource = "dataTag1"
Form_frmHistory.dataTag2.ControlSource = "dataTag0"
Form_frmHistory.dataTag1.ControlSource = "tableName"
Form_frmHistory.filter = "[tableName] = 'tblReporting_notes' AND tableRecordId = " & Me.noteId
Form_frmHistory.FilterOn = True
Form_frmHistory.OrderBy = "updatedDate Desc"
Form_frmHistory.OrderByOn = True


Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub reportNote_AfterUpdate()
On Error GoTo Err_Handler

registerWdbUpdates "tblReporting_notes", Me.noteId, "reportNote", Me.reportNote.OldValue, Me.reportNote.Value, Me.name, Me.partNumber

Me.updatedBy = Environ("username")
Me.updatedDate = Now()
Me.dataTag0 = "nmq_morning_trials"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
