Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Function validate()
On Error GoTo Err_Handler

validate = False

Dim errorArray As Collection
Set errorArray = New Collection

'check stuff
If Nz(Me.userName) = "" Then errorArray.Add "Username is Blank"
If Nz(Me.schedulereason) = 0 Then errorArray.Add "Reason is Blank"
If Nz(Me.startdate) = "" Then errorArray.Add "Start Date is Blank"
If Nz(Me.endDate) = "" Then errorArray.Add "End Date is Blank"
If Nz(Me.hoursPerDay) = 0 Then errorArray.Add "Hours per Day is Blank"

If errorArray.count > 0 Then
    Dim errorTxtLines As String, element
    errorTxtLines = ""
    For Each element In errorArray
        errorTxtLines = errorTxtLines & vbNewLine & element
    Next element
    
    MsgBox "Please fix these items: " & vbNewLine & errorTxtLines, vbOKOnly, "ACTION REQUIRED"
    Exit Function
End If

validate = True

Exit Function
Err_Handler:
    Call handleError(Me.name, "validate", Err.DESCRIPTION, Err.number)
End Function

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub save_Click()
On Error GoTo Err_Handler

If Me.Dirty Then Me.Dirty = False
If validate = False Then Exit Sub

'do loop through the days and add the records
Dim intDate As Date
intDate = Me.startdate

Dim refId As Long
refId = DLookup("recordId", "tbllab_tech_schedule_template", "username = '" & Me.userName & "'")

Dim db As Database
Set db = CurrentDb()

Dim rs As Recordset
Set rs = db.OpenRecordset("tbllab_tech_schedule_alterations")

Do While intDate <= Me.endDate
    
    With rs
        .addNew
        !userName = Me.userName
        !scheduledate = intDate
        !schedulehours = Me.hoursPerDay
        !schedulereason = Me.schedulereason
        !scheduletemplateid = refId
        .Update
    End With
    
    TempVars.Add "scheduleId", db.OpenRecordset("SELECT @@identity")(0).Value
    Call registerLabUpdates("tbllab_tech_schedule_alterations", TempVars!scheduleId, "Tech Schedule", "", "Saved", CStr(refId), Me.name)
    
    intDate = DateAdd("d", 1, intDate)
Loop

rs.CLOSE
Set rs = Nothing
Set db = Nothing

DoCmd.CLOSE acForm, "frmLab_tech_schedule_alteration_create"
If CurrentProject.AllForms("frmLab_tech_schedule_details").IsLoaded Then Form_frmLab_tech_schedule_details.Requery

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub startDate_AfterUpdate()
On Error GoTo Err_Handler

If IsNull(Me.endDate) Then
    Me.endDate = Me.startdate
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
