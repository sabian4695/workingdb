Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub addNew_Click()
On Error GoTo Err_Handler

Dim packType As Long
If Me.Recordset.RecordCount > 0 Then
    packType = 2
Else
    packType = 1
End If

Dim db As Database
Set db = CurrentDb()
db.Execute "INSERT INTO tblPartPackagingInfo(partInfoId, packType) VALUES (" & Me.partInfoId & "," & packType & ");"
TempVars.Add "packId", db.OpenRecordset("SELECT @@identity")(0).Value
Set db = Nothing

Call registerPartUpdates("tblPartPackagingInfo", TempVars!packId, DLookup("packagingType", "tblDropDownsSP", "recordid = " & packType), "", "Created", Me.Parent.partNumber, Me.name)
Me.Requery

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub boxesPerSkid_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartPackagingInfo", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.Parent.partNumber, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Dim allowEdits As Boolean

allowEdits = (Not restrict(Environ("username"), "Packaging") Or Not restrict(Environ("username"), "Project")) And Me.Parent.dataFreeze = False 'only project/service can edit things in this form, and only when dataFreeze is false

Me.addNew.Visible = allowEdits
Me.allowEdits = allowEdits
Me.remove.Visible = allowEdits
Form_sfrmPartPackagingComponents.Form.allowEdits = allowEdits
Form_sfrmPartPackagingComponents.Form.AllowAdditions = allowEdits
Form_sfrmPartPackagingComponents.remove.Visible = allowEdits

If Me.Parent.dataFreeze Then
    lblLock.Caption = "Data Frozen"
Else
    lblLock.Caption = "only Packaging/Project can edit"
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub packType_AfterUpdate()
On Error GoTo Err_Handler

If Me.Dirty Then Me.Dirty = False
If DCount("recordId", "tblPartPackagingInfo", "partInfoId = " & Me.partInfoId & " AND packType = 1") > 1 Then
    Me.packType = 2
    Call snackBox("error", "No.", "You can only have one primary pack for a part", Me.Parent.name)
    Exit Sub
End If

Call registerPartUpdates("tblPartPackagingInfo", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.Parent.partNumber, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub primaryCustomer_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartPackagingInfo", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.Parent.partNumber, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub remove_Click()
On Error GoTo Err_Handler

If IsNull(Me.recordId) Then Exit Sub

If (Not restrict(Environ("username"), "Packaging") Or Not restrict(Environ("username"), "Project")) = False And Me.Parent.dataFreeze = False Then
    MsgBox "Only Packaging Engineers can do this", vbCritical, "Denied"
    Exit Sub
End If

If MsgBox("Are you sure you want to delete this?", vbYesNo, "Please confirm") <> vbYes Then Exit Sub

Call registerPartUpdates("tblPartPackagingInfo", Me.recordId, "Part Packaging Info", Me.packType.column(1), "Deleted", Me.Parent.partNumber, Me.name)

Dim db As Database
Set db = CurrentDb()
db.Execute ("DELETE FROM tblPartPackagingComponents WHERE [packagingInfoId] = " & Me.recordId)
db.Execute ("DELETE FROM tblPartPackagingInfo WHERE [recordId] = " & Me.recordId)

Set db = Nothing

Me.Requery
If Me.Recordset.RecordCount = 0 Then Me.sfrmPartPackagingComponents.Visible = False
Call snackBox("success", "Nice", "Packaging successfully deleted.", Me.Parent.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
