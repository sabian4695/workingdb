Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub componentNumber_AfterUpdate()
On Error GoTo Err_Handler

'make sure this part can be added
Dim errorTxt As String, partNum As String
errorTxt = ""
partNum = Nz(Me.componentNumber)

If partNum = "" Then Exit Sub

Form_DASHBOARD.partNumberSearch = partNum
Form_DASHBOARD.filterbyPN_Click

If DCount("recordId", "tblPartProject", "partNumber = '" & partNum & "'") > 0 Then errorTxt = "Project for this part already exists"
If DCount("recordId", "tblPartProjectPartNumbers", "childPartNumber = '" & partNum & "'") > 0 Then errorTxt = "This part is linked to another project"
If Form_DASHBOARD.lblErrors.Visible = True And Form_DASHBOARD.lblErrors.Caption = "Part not found in Oracle" Then errorTxt = "This part doesn't show up in Oracle"

If errorTxt <> "" Then
    MsgBox errorTxt, vbCritical, "Sorry"
    Me.componentNumber = ""
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub remove_Click()
On Error GoTo Err_Handler

Dim db As Database
Set db = CurrentDb()

If Me.NewRecord Then Exit Sub
db.Execute "DELETE * FROM tblSessionVariables WHERE ID = " & Me.ID
Me.Requery

Set db = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
