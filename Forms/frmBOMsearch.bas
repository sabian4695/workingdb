Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub clear_Click()
On Error GoTo Err_Handler

Me.NAMsrchBox = ""
Me.NAMsrchBox.SetFocus

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub


Private Sub ComptSrch_Click()
On Error GoTo Err_Handler

Dim db As Database
Set db = CurrentDb()

Dim qdf As QueryDef

Set qdf = db.QueryDefs("sqryBOM")

If Nz(Me.NAMsrchBox, "") <> "" Then
    qdf.sql = Split(qdf.sql, "WHERE")(0) & " WHERE (sysItems1.SEGMENT1 = '" & Me.NAMsrchBox & "' AND bomInv.DISABLE_DATE Is Null);"
Else
    qdf.sql = Split(qdf.sql, "WHERE")(0) & " WHERE (sysItems.SEGMENT1 <> '' AND bomInv.DISABLE_DATE Is Null);"
End If
    
db.QueryDefs.refresh

Set qdf = Nothing
Set db = Nothing

Me.Requery

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub export_Click()
On Error GoTo Err_Handler

Dim db As Database
Set db = CurrentDb()

Dim qdf As QueryDef

Set qdf = db.QueryDefs("frmBOM")

Dim FileName As String, sqlString As String
FileName = "H:\BOMsearch_" & nowString & ".xlsx"
sqlString = qdf.sql

Call exportSQL(sqlString, FileName)

Set qdf = Nothing
Set db = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Assysrch_Click()
On Error GoTo Err_Handler

Dim db As Database
Set db = CurrentDb()

Dim qdf As QueryDef

Set qdf = db.QueryDefs("sqryBOM")

If Nz(Me.NAMsrchBox, "") <> "" Then
    qdf.sql = Split(qdf.sql, "WHERE")(0) & " WHERE (sysItems.SEGMENT1 = '" & Me.NAMsrchBox & "' AND bomInv.DISABLE_DATE Is Null);"
Else
    qdf.sql = Split(qdf.sql, "WHERE")(0) & " WHERE (sysItems.SEGMENT1 <> '' AND bomInv.DISABLE_DATE Is Null);"
End If
    
db.QueryDefs.refresh

Set qdf = Nothing
Set db = Nothing

Me.Requery

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub btnOHQ_Click()
On Error GoTo Err_Handler

Form_DASHBOARD.partNumberSearch = Me.Compt

If CurrentProject.AllForms("frmOnHandQty").IsLoaded = True Then DoCmd.CLOSE acForm, "frmOnHandQty"

DoCmd.OpenForm "frmOnHandQty"
DoCmd.CLOSE acForm, "frmBOMsearch"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Current()
On Error GoTo Err_Handler

Dim db As Database
Set db = CurrentDb()

Dim qdf As QueryDef

Set qdf = db.QueryDefs("sfrm_sqryBOM")
qdf.sql = Split(qdf.sql, "WHERE")(0) & " WHERE (sysItems.SEGMENT1 = '" & Me.Compt & "' AND bomInv.DISABLE_DATE Is Null);"
    
db.QueryDefs.refresh

Set qdf = Nothing
Set db = Nothing

Me.sfrmBOMsearch.Requery

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblAssy_Click()
On Error GoTo Err_Handler

Me.Assy.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblAssyDesc_Click()
On Error GoTo Err_Handler

Me.assyDescription.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblAssyStatus_Click()
On Error GoTo Err_Handler

Me.assyStatus.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblCompDesc_Click()
On Error GoTo Err_Handler

Me.compDescription.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblCompStatus_Click()
On Error GoTo Err_Handler

Me.compStatus.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblCompt_Click()
On Error GoTo Err_Handler

Me.Compt.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblImp_Click()
On Error GoTo Err_Handler

Me.IMPLEMENTATION_DATE.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblInverse_Click()
On Error GoTo Err_Handler

Me.Inverse_Qty.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblOrg_Click()
On Error GoTo Err_Handler

Me.Org.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblQty_Click()
On Error GoTo Err_Handler

Me.Qty.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
