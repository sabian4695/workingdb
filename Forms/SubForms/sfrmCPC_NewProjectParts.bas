Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim rs As DAO.Recordset

Private Sub btnDeletePart_Click()

On Error GoTo Err_Handler

Dim ID As Long

If IsNull(Me.ID) Then
    Exit Sub
End If

If MsgBox("Are you sure you want to delete this part?", vbYesNo, "Warning") = vbYes Then
    ID = Me.ID
    DoCmd.GoToRecord , , acNewRec
    dbExecute ("DELETE * FROM tblCPC_Parts WHERE [id] = " & ID)
    Me.Requery
End If

exit_handler:
    Exit Sub

Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
    Resume exit_handler

End Sub

Private Sub txtPartNumber_AfterUpdate()

On Error GoTo Err_Handler

Dim message
Dim ID As Long

Insert_Single_Description
Update_Row_Source

If Me.txtUnit = "U12" Then
    message = MsgBox("This part is in U12" & vbCrLf & "Are you sure you want to add this part?", vbYesNo + vbExclamation, "Warning")
    If message <> vbNo Then
        Exit Sub
    End If
    
    ID = Me.ID
    DoCmd.GoToRecord , , acNewRec
    dbExecute ("DELETE * FROM tblCPC_Parts WHERE id = " & ID)
    Me.Requery
End If
    
exit_handler:
    Exit Sub

Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
    Resume exit_handler

End Sub

Private Sub cmbCustPartNumber_Enter()

On Error GoTo Err_Handler

If IsNull(Me.txtPartNumber) Then
    Me.cmbCustPartNumber.RowSource = ""
    Exit Sub
End If

Me.cmbCustPartNumber.RowSource = "SELECT INV_MTL_CUSTOMER_ITEMS.CUSTOMER_ITEM_NUMBER " & _
    "FROM (INV_MTL_CUSTOMER_ITEM_XREFS INNER JOIN APPS_MTL_SYSTEM_ITEMS ON INV_MTL_CUSTOMER_ITEM_XREFS.INVENTORY_ITEM_ID = APPS_MTL_SYSTEM_ITEMS.INVENTORY_ITEM_ID) INNER JOIN INV_MTL_CUSTOMER_ITEMS ON INV_MTL_CUSTOMER_ITEM_XREFS.CUSTOMER_ITEM_ID = INV_MTL_CUSTOMER_ITEMS.CUSTOMER_ITEM_ID " & _
    "WHERE APPS_MTL_SYSTEM_ITEMS.SEGMENT1='" & Me.partNumber & "' " & _
    "GROUP BY INV_MTL_CUSTOMER_ITEMS.CUSTOMER_ITEM_NUMBER " & _
    "ORDER BY INV_MTL_CUSTOMER_ITEMS.CUSTOMER_ITEM_NUMBER;"

If Me.cmbCustPartNumber.ListCount = 1 Then
    Me.cmbCustPartNumber = Me.cmbCustPartNumber.column(0, 0)
End If

exit_handler:
    Exit Sub

Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
    Resume exit_handler

End Sub

Private Sub Insert_Single_Description()

On Error GoTo Err_Handler

Dim QUERY As String

QUERY = "SELECT APPS_MTL_SYSTEM_ITEMS.DESCRIPTION, APPS_MTL_CATEGORIES_VL.SEGMENT1 " & _
    "FROM (APPS_MTL_SYSTEM_ITEMS INNER JOIN INV_MTL_ITEM_CATEGORIES ON APPS_MTL_SYSTEM_ITEMS.INVENTORY_ITEM_ID = INV_MTL_ITEM_CATEGORIES.INVENTORY_ITEM_ID) INNER JOIN APPS_MTL_CATEGORIES_VL ON INV_MTL_ITEM_CATEGORIES.CATEGORY_ID = APPS_MTL_CATEGORIES_VL.CATEGORY_ID " & _
    "GROUP BY APPS_MTL_SYSTEM_ITEMS.DESCRIPTION, APPS_MTL_CATEGORIES_VL.SEGMENT1, APPS_MTL_SYSTEM_ITEMS.SEGMENT1 " & _
    "HAVING APPS_MTL_CATEGORIES_VL.SEGMENT1 Like 'U*' AND APPS_MTL_SYSTEM_ITEMS.SEGMENT1='"

If IsNull(Me.txtPartNumber) Then
    Exit Sub
End If

Dim db As DAO.Database
Set db = CurrentDb
Set rs = db.OpenRecordset(QUERY & Me.txtPartNumber & "';", dbOpenSnapshot)

If rs.RecordCount = 0 Then
    Exit Sub
End If

' Me.txtDescription = rs("DESCRIPTION")
Me.txtUnit = rs("SEGMENT1")

rs.CLOSE
Set rs = Nothing
Set db = Nothing

exit_handler:
    Exit Sub

Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
    Resume exit_handler

End Sub

Sub Update_Row_Source()

On Error GoTo Err_Handler

If IsNull(Me.txtPartNumber) Then
    Me.cmbCustPartNumber.RowSource = ""
    Exit Sub
End If

Me.cmbCustPartNumber.RowSource = "SELECT INV_MTL_CUSTOMER_ITEMS.CUSTOMER_ITEM_NUMBER " & _
    "FROM (INV_MTL_CUSTOMER_ITEM_XREFS INNER JOIN APPS_MTL_SYSTEM_ITEMS ON INV_MTL_CUSTOMER_ITEM_XREFS.INVENTORY_ITEM_ID = APPS_MTL_SYSTEM_ITEMS.INVENTORY_ITEM_ID) INNER JOIN INV_MTL_CUSTOMER_ITEMS ON INV_MTL_CUSTOMER_ITEM_XREFS.CUSTOMER_ITEM_ID = INV_MTL_CUSTOMER_ITEMS.CUSTOMER_ITEM_ID " & _
    "WHERE APPS_MTL_SYSTEM_ITEMS.SEGMENT1='" & Me.partNumber & "' " & _
    "GROUP BY INV_MTL_CUSTOMER_ITEMS.CUSTOMER_ITEM_NUMBER " & _
    "ORDER BY INV_MTL_CUSTOMER_ITEMS.CUSTOMER_ITEM_NUMBER;"

If Me.cmbCustPartNumber.ListCount = 1 Then
    Me.cmbCustPartNumber = Me.cmbCustPartNumber.column(0, 0)
End If

exit_handler:
    Exit Sub

Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
    Resume exit_handler

End Sub
