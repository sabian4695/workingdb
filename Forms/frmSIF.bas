Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub clear_Click()
On Error GoTo Err_Handler
Me.srchBox.SetFocus
Me.srchBox = ""
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Dim i As Long
For i = 1 To 51
    Me.Controls("lbl" & i).Visible = False
Next i

Me.srchBox.SetFocus
Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Sub srch_Click()
On Error GoTo Err_Handler

Dim i As Long
For i = 1 To 48
    Me.Controls("lbl" & i).Visible = False
Next i

Dim db As DAO.Database
Dim rsSIF As Recordset
Dim srchTxt As String
srchTxt = Me.srchBox

Set db = CurrentDb()
Set rsSIF = db.OpenRecordset("SELECT * FROM APPS_Q_SIF_NEW_MOLDED_PART_V where SIFNUM = '" & srchTxt & "'")

If rsSIF.RecordCount = 0 Then Set rsSIF = db.OpenRecordset("SELECT * FROM APPS_Q_SIF_NEW_ASSEMBLED_PART_V where SIFNUM = '" & srchTxt & "'")
If rsSIF.RecordCount = 0 Then Set rsSIF = db.OpenRecordset("SELECT * FROM APPS_Q_SIF_NEW_PURCHASING_PART_V where SIFNUM = '" & srchTxt & "'")

If rsSIF.RecordCount = 0 Then
    MsgBox "no records found!", vbInformation, "Woopsy"
    Exit Sub
End If

Dim counter As Long
counter = 1
Dim fld As DAO.Field
For Each fld In rsSIF.Fields
    Select Case fld.name
        Case "ROW_ID", "PLAN_ID", "ORGANIZATION_ID", "COLLECTION_ID", "OCCURRENCE", "LAST_UPDATED_BY_ID", "CREATED_BY_ID", "LAST_UPDATE_LOGIN"
        Case "NOTES"
            Me.txtNotes.Visible = True
            Me.txtNotes = fld.Value
        Case Else
            Me.Controls("lbl" & counter).Caption = Replace(fld.name, "_", " ")
            Me.Controls("lbl" & counter).Visible = True
            Me.Controls("val" & counter) = fld.Value
                counter = counter + 1
    End Select
Next

On Error Resume Next
Set fld = Nothing
rsSIF.CLOSE
Set rsSIF = Nothing
Set db = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
