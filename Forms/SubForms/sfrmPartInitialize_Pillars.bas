Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Sub updateAllPillars()
On Error Resume Next

If Me.Dirty Then Me.Dirty = False

Dim db As Database
Set db = CurrentDb()

Dim rs As Recordset
Set rs = Me.RecordsetClone

rs.MoveFirst
Do While Not rs.EOF
    rs.Edit
    rs!pillarDue = addWorkdays(Form_frmPartInitialize.opDate, rs!pillarLength)
    rs.Update
    rs.MoveNext
Loop

rs.CLOSE
Set rs = Nothing
Set db = Nothing
Me.Requery

End Sub

Private Sub pillarDue_AfterUpdate()
On Error Resume Next

If Me.Dirty Then Me.Dirty = False

If Me.pillarLength = 0 Then 'this is the hinge pillar - OPT0 typically
    'if you change the hinge pillar, you need to recalculate stuff.
    'first, change the OPT0 date in the master form, then run through all numbers and recalculate
    
    Form_frmPartInitialize.opDate = Me.pillarDue
    
    Dim db As Database
    Set db = CurrentDb()
    
    Dim rs As Recordset
    Set rs = Me.RecordsetClone
    
    rs.MoveFirst
    Do While Not rs.EOF
        rs.Edit
        rs!pillarLength = countWorkdays(Form_frmPartInitialize.opDate, rs!pillarDue)
        rs.Update
        rs.MoveNext
    Loop
    
    rs.CLOSE
    Set rs = Nothing
    Set db = Nothing
    Me.Requery
Else
    Me.pillarLength = countWorkdays(Form_frmPartInitialize.opDate, Me.pillarDue)
End If

End Sub

Private Sub pillarLength_AfterUpdate()
On Error Resume Next
Me.pillarDue = addWorkdays(Form_frmPartInitialize.opDate, Me.pillarLength)

End Sub

Private Sub remove_Click()
On Error GoTo Err_Handler

Dim db As Database
Set db = CurrentDb()

If Me.NewRecord Then GoTo exit_handler
If MsgBox("Are you sure?", vbYesNo, "Just checking.") <> vbYes Then GoTo exit_handler

db.Execute "DELETE * FROM tblSessionVariables WHERE ID = " & Me.ID
Me.Requery

updateAllPillars

exit_handler:
Set db = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
