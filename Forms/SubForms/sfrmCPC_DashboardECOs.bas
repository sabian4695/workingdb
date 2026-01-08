Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Dirty(Cancel As Integer)
On Error GoTo Err_Handler

Me.Parent.lastModified = Now

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub btnDeleteECO_Click()
On Error GoTo Err_Handler

If IsNull(Me.ID) Then Exit Sub

If MsgBox("Are you sure you want to delete this ECO?", vbYesNo, "Warning") = vbYes Then
    Call registerCPCUpdates("tblCPC_ECOs", ID, Me.ecoNumber.name, Me.ecoNumber, "Deleted", Form_frmCPC_Dashboard.ID)
    dbExecute ("DELETE * FROM tblCPC_ECOs WHERE [id] = " & Me.ID)
    Me.Requery
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub btnDetails_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmECOs", , , "[Change_Notice] = '" & Me.ecoNumber & "'"
Form_frmECOs.ECOsrch = Me.ecoNumber

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub ECONumber_AfterUpdate()
On Error GoTo Err_Handler

Call registerCPCUpdates("tblCPC_ECOs", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Form_frmCPC_Dashboard.ID)

Dim message
Dim db As DAO.Database
Dim rs1 As DAO.Recordset
Dim rs2 As DAO.Recordset
Dim rowsAdded As Long

message = MsgBox("Do you want to load the revised items from this ECO into Part Tracking?", vbYesNo + vbQuestion, "Load Revised Items")

If message <> vbYes Then
    Exit Sub
End If

Set db = CurrentDb

Set rs1 = db.OpenRecordset("SELECT ENG_ENG_REVISED_ITEMS.CHANGE_NOTICE, APPS_MTL_SYSTEM_ITEMS.SEGMENT1, ENG_ENG_REVISED_ITEMS.NEW_ITEM_REVISION " & _
"FROM APPS_MTL_SYSTEM_ITEMS INNER JOIN ENG_ENG_REVISED_ITEMS ON APPS_MTL_SYSTEM_ITEMS.INVENTORY_ITEM_ID = ENG_ENG_REVISED_ITEMS.REVISED_ITEM_ID " & _
"GROUP BY ENG_ENG_REVISED_ITEMS.CHANGE_NOTICE, APPS_MTL_SYSTEM_ITEMS.SEGMENT1, ENG_ENG_REVISED_ITEMS.NEW_ITEM_REVISION " & _
"HAVING ENG_ENG_REVISED_ITEMS.CHANGE_NOTICE='" & Me.ecoNumber & "';", dbOpenSnapshot)

Set rs2 = db.OpenRecordset("SELECT partNumber FROM tblCPC_Parts WHERE projectId = " & Form_frmCPC_Dashboard.ID, dbOpenSnapshot)

rs1.MoveFirst
Do While Not rs1.EOF
    rs2.FindFirst ("partNumber = '" & rs1("SEGMENT1") & "'")
    If rs2.noMatch Then
        db.Execute ("INSERT INTO tblCPC_Parts (projectNumber, partNumber, newRev) " & _
                        "SELECT '" & Me.projectNumber & "', '" & rs1("SEGMENT1") & "', '" & rs1("NEW_ITEM_REVISION") & "'")
    End If
    rs1.MoveNext
Loop

rowsAdded = db.RecordsAffected

If rowsAdded <> 0 Then
    Form_sfrmCPC_DashboardParts.Requery
    Form_sfrmCPC_DashboardParts.Insert_Many_Description
End If

rs1.CLOSE
rs2.CLOSE
Set rs1 = Nothing
Set rs2 = Nothing
Set db = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
