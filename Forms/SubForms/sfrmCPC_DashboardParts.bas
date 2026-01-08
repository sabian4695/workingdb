Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim db As DAO.Database
Dim rs1 As DAO.Recordset
Dim rs2 As DAO.Recordset

Private Sub custPartNumber_AfterUpdate()
On Error GoTo Err_Handler
Call registerCPCUpdates("tblCPC_Parts", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Form_frmCPC_Dashboard.ID)
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub custTrackingNumber_AfterUpdate()
On Error GoTo Err_Handler
Call registerCPCUpdates("tblCPC_Parts", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Form_frmCPC_Dashboard.ID)
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub description_AfterUpdate()
On Error GoTo Err_Handler
Call registerCPCUpdates("tblCPC_Parts", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Form_frmCPC_Dashboard.ID)
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub firstShipDate_AfterUpdate()
On Error GoTo Err_Handler
Call registerCPCUpdates("tblCPC_Parts", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Form_frmCPC_Dashboard.ID)
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Insert_Many_Description

DoCmd.GoToRecord , , acNewRec
Me.partNumber.SetFocus

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Dirty(Cancel As Integer)
On Error GoTo Err_Handler

Me.Parent.lastModified = Now

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub newRev_AfterUpdate()
On Error GoTo Err_Handler

Call registerCPCUpdates("tblCPC_Parts", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Form_frmCPC_Dashboard.ID)
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub partnumber_AfterUpdate()
On Error GoTo Err_Handler

Insert_Single_Description
Update_Row_Source

Call registerCPCUpdates("tblCPC_Parts", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Form_frmCPC_Dashboard.ID)
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub custPartNumber_Enter()
On Error GoTo Err_Handler
Update_Row_Source
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub docHisSearch_Click()
On Error GoTo Err_Handler

Call openDocumentHistoryFolder(Me.partNumber)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub btnDeletePart_Click()
On Error GoTo Err_Handler

Dim ID As Long

If IsNull(Me.ID) Then Exit Sub

If MsgBox("Are you sure you want to delete this part?", vbYesNo, "Warning") = vbYes Then
    ID = Me.ID
    Call registerCPCUpdates("tblCPC_Parts", ID, Me.partNumber.name, Me.partNumber, "Deleted", Form_frmCPC_Dashboard.ID)
    DoCmd.GoToRecord , , acNewRec
    dbExecute ("DELETE * FROM tblCPC_Parts WHERE [id] = " & ID)
    Me.Requery
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Sub Insert_Many_Description()
On Error GoTo Err_Handler

Dim QUERY As String

Set db = CurrentDb
Set rs1 = db.OpenRecordset("SELECT partNumber, description, unit FROM tblCPC_Parts WHERE projectId = " & Form_frmCPC_Dashboard.ID, dbOpenSnapshot)

QUERY = "SELECT APPS_MTL_SYSTEM_ITEMS.DESCRIPTION, APPS_MTL_CATEGORIES_VL.SEGMENT1 " & _
    "FROM (APPS_MTL_SYSTEM_ITEMS INNER JOIN INV_MTL_ITEM_CATEGORIES ON APPS_MTL_SYSTEM_ITEMS.INVENTORY_ITEM_ID = INV_MTL_ITEM_CATEGORIES.INVENTORY_ITEM_ID) INNER JOIN APPS_MTL_CATEGORIES_VL ON INV_MTL_ITEM_CATEGORIES.CATEGORY_ID = APPS_MTL_CATEGORIES_VL.CATEGORY_ID " & _
    "GROUP BY APPS_MTL_SYSTEM_ITEMS.DESCRIPTION, APPS_MTL_CATEGORIES_VL.SEGMENT1, APPS_MTL_SYSTEM_ITEMS.SEGMENT1 " & _
    "HAVING APPS_MTL_CATEGORIES_VL.SEGMENT1 Like 'U*' AND APPS_MTL_SYSTEM_ITEMS.SEGMENT1='"

If rs1.RecordCount = 0 Then
    Exit Sub
End If

If Not (rs1.BOF And rs1.EOF) Then
    rs1.MoveFirst
    Do While Not rs1.EOF
        If IsNull(rs1("description")) Or IsNull(rs1("unit")) Then
            On Error Resume Next
            Set rs2 = db.OpenRecordset(QUERY & rs1("partNumber") & "';", dbOpenSnapshot)
            If rs2.RecordCount = 0 Then
                rs1.MoveNext
            End If
            rs1.Edit
            rs1("description") = rs2("DESCRIPTION")
            rs1("unit") = rs2("SEGMENT1")
            rs1.Update
        End If
        rs1.MoveNext
    Loop
    
    If db.RecordsAffected > 0 Then
        Me.Requery
    End If
End If

Set db = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Insert_Single_Description()
On Error GoTo Err_Handler

Dim QUERY As String

QUERY = "SELECT APPS_MTL_SYSTEM_ITEMS.DESCRIPTION, APPS_MTL_CATEGORIES_VL.SEGMENT1 " & _
    "FROM (APPS_MTL_SYSTEM_ITEMS INNER JOIN INV_MTL_ITEM_CATEGORIES ON APPS_MTL_SYSTEM_ITEMS.INVENTORY_ITEM_ID = INV_MTL_ITEM_CATEGORIES.INVENTORY_ITEM_ID) INNER JOIN APPS_MTL_CATEGORIES_VL ON INV_MTL_ITEM_CATEGORIES.CATEGORY_ID = APPS_MTL_CATEGORIES_VL.CATEGORY_ID " & _
    "GROUP BY APPS_MTL_SYSTEM_ITEMS.DESCRIPTION, APPS_MTL_CATEGORIES_VL.SEGMENT1, APPS_MTL_SYSTEM_ITEMS.SEGMENT1 " & _
    "HAVING APPS_MTL_CATEGORIES_VL.SEGMENT1 Like 'U*' AND APPS_MTL_SYSTEM_ITEMS.SEGMENT1='"

If IsNull(Me.partNumber) Then
    Exit Sub
End If

Set db = CurrentDb
Set rs1 = db.OpenRecordset(QUERY & Me.partNumber & "';", dbOpenSnapshot)

If rs1.RecordCount = 0 Then
    Exit Sub
End If

Me.DESCRIPTION = rs1("DESCRIPTION")
Me.unit = rs1("SEGMENT1")

rs1.CLOSE
Set rs1 = Nothing
Set db = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Sub Update_Row_Source()
On Error GoTo Err_Handler

If IsNull(Me.partNumber) Then
    Me.custPartNumber.RowSource = ""
    Exit Sub
End If

Me.custPartNumber.RowSource = "SELECT INV_MTL_CUSTOMER_ITEMS.CUSTOMER_ITEM_NUMBER " & _
    "FROM (INV_MTL_CUSTOMER_ITEM_XREFS INNER JOIN APPS_MTL_SYSTEM_ITEMS ON INV_MTL_CUSTOMER_ITEM_XREFS.INVENTORY_ITEM_ID = APPS_MTL_SYSTEM_ITEMS.INVENTORY_ITEM_ID) INNER JOIN INV_MTL_CUSTOMER_ITEMS ON INV_MTL_CUSTOMER_ITEM_XREFS.CUSTOMER_ITEM_ID = INV_MTL_CUSTOMER_ITEMS.CUSTOMER_ITEM_ID " & _
    "WHERE APPS_MTL_SYSTEM_ITEMS.SEGMENT1='" & Me.partNumber & "' " & _
    "GROUP BY INV_MTL_CUSTOMER_ITEMS.CUSTOMER_ITEM_NUMBER " & _
    "ORDER BY INV_MTL_CUSTOMER_ITEMS.CUSTOMER_ITEM_NUMBER;"

If Me.custPartNumber.ListCount = 1 Then
    Me.custPartNumber = Me.custPartNumber.column(0, 0)
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub runoutDate_AfterUpdate()
On Error GoTo Err_Handler
Call registerCPCUpdates("tblCPC_Parts", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Form_frmCPC_Dashboard.ID)
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub shipped_AfterUpdate()
On Error GoTo Err_Handler
Call registerCPCUpdates("tblCPC_Parts", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Form_frmCPC_Dashboard.ID)
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub unit_AfterUpdate()
On Error GoTo Err_Handler
Call registerCPCUpdates("tblCPC_Parts", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Form_frmCPC_Dashboard.ID)
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
