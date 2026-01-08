Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub businessCode_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartInfo", Me("tblPartInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub copyClass_Click()
On Error GoTo Err_Handler

Dim db As Database
Set db = CurrentDb()

Dim X As String

X = InputBox("Please enter a part number to copy from", "Enter Part Number", Form_frmCrossFunctionalKO.partNumber)

If StrPtr(X) = 0 Then GoTo exit_handler
If X = "" Then GoTo exit_handler

Dim rsPI As Recordset
Set rsPI = db.OpenRecordset("SELECT * FROM tblPartInfo WHERE partNumber = '" & X & "'")

If rsPI.RecordCount = 0 Then
    MsgBox "No class info found"
    GoTo exit_handler
End If

Call registerPartUpdates("tblPartInfo", Me("tblPartInfo.recordId"), Me.partClassCode.name, Me.partClassCode, rsPI!partClassCode, Me.partNumber, Me.name)
Call registerPartUpdates("tblPartInfo", Me("tblPartInfo.recordId"), Me.subClassCode.name, Me.subClassCode, rsPI!subClassCode, Me.partNumber, Me.name)
Call registerPartUpdates("tblPartInfo", Me("tblPartInfo.recordId"), Me.businessCode.name, Me.businessCode, rsPI!businessCode, Me.partNumber, Me.name)
Call registerPartUpdates("tblPartInfo", Me("tblPartInfo.recordId"), Me.focusAreaCode.name, Me.focusAreaCode, rsPI!focusAreaCode, Me.partNumber, Me.name)

Me.partClassCode = rsPI!partClassCode
Me.subClassCode = rsPI!subClassCode
Me.businessCode = rsPI!businessCode
Me.focusAreaCode = rsPI!focusAreaCode

If Me.Dirty Then Me.Dirty = False

rsPI.CLOSE
Set rsPI = Nothing

exit_handler:
Set db = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub focusAreaCode_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartInfo", Me("tblPartInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Current()
On Error GoTo Err_Handler

Me.subClassCode.RowSource = "SELECT recordId, subClassCode, subClassCodeName, subClassCodeCat From tblPartClassification WHERE subClassCode Is Not Null AND subClassCodeCat = '" & Me.partClassCode.column(3) & "'"

Select Case Me.partClassCode.column(3)
    Case "FBU"
        Me.businessCode = 4
    Case "ADAS"
        Me.businessCode = 9
        Me.focusAreaCode = 5
    Case "FCS"
        Me.businessCode = 1
    Case "PF"
        Me.businessCode = 3
    Case "MCD"
        Me.businessCode = 2
    Case "LSC"
        Me.businessCode = 5
End Select

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Current", Err.DESCRIPTION, Err.number)
End Sub

Private Sub partClassCode_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartInfo", Me("tblPartInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)

Me.subClassCode.RowSource = "SELECT recordId, subClassCode, subClassCodeName, subClassCodeCat From tblPartClassification WHERE subClassCode Is Not Null AND subClassCodeCat = '" & Me.partClassCode.column(3) & "'"

Select Case Me.partClassCode.column(3)
    Case "FBU"
        Me.businessCode = 4
    Case "ADAS"
        Me.businessCode = 9
        Me.focusAreaCode = 5
    Case "FCS"
        Me.businessCode = 1
    Case "PF"
        Me.businessCode = 3
    Case "MCD"
        Me.businessCode = 2
    Case "LSC"
        Me.businessCode = 5
End Select

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub subClassCode_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartInfo", Me("tblPartInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
