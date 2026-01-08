Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnSave_Click()
On Error GoTo Err_Handler

Dim db As DAO.Database
Set db = CurrentDb()

If Me.Dirty Then Me.Dirty = False
Dim errorTxt As String
errorTxt = ""

If (Me.projectTemplateId = "" Or IsNull(Me.projectTemplateId)) Then errorTxt = "Please select a template"
If IsNull(Me.projectStartDate) Then errorTxt = "Please enter a project start date"
If IsNull(Me.opDate) Then errorTxt = "Please fill out foundational pillar date first"

If errorTxt <> "" Then
    MsgBox errorTxt, vbCritical, "Nope"
    Exit Sub
End If

'Dim Assy As Boolean
'If Me.projectTemplateId <> 8 Then
'    Assy = False
'    GoTo skipAssyCheck
'End If

'PILLAR IMPLEMENTATION - remove ASSY creation area. Not sure how to implement pillars for this yet
'Dim rs As Recordset
'Set rs = db.OpenRecordset("SELECT * from tblSessionVariables WHERE assemblyNumber = '" & Me.partNumber & "' AND componentNumber is not null")
'Assy = True
'If rs.RecordCount = 0 Then
'    Assy = False
'    GoTo skipAssyCheck
'End If

If errorTxt <> "" Then
    MsgBox errorTxt, vbCritical, "Nope"
    Exit Sub
End If

'If Assy Then 'assembly
'    rs.MoveFirst
'    Do While Not rs.EOF
'        If Nz(rs!componentProjId) <> 0 Then 'create a project if template is selected
'            db.Execute "INSERT INTO tblPartProject(partNumber,projectTemplateId,projectStartDate,projectStatus) VALUES('" & rs!componentNumber & "'," & rs!componentProjId & ",'" & Me.projectStartDate & "',1)"
'            TempVars.Add "projId", db.OpenRecordset("SELECT @@identity")(0).Value
'            Call createPartProject(TempVars!projId)
'        End If
'
'        'check if component is already on this BOM
'        If DCount("recordId", "tblPartComponents", "assemblyNumber = '" & Me.partNumber & "' AND componentNumber = '" & rs!componentNumber & "'") = 0 Then
'            db.Execute "INSERT INTO tblPartComponents(assemblyNumber,componentNumber) VALUES('" & Me.partNumber & "','" & rs!componentNumber & "')"
'        End If
'        rs.MoveNext
'    Loop
'End If

'rs.Close
'Set rs = Nothing


'skipAssyCheck:
db.Execute "INSERT INTO tblPartProject(partNumber,projectTemplateId,projectStartDate,projectStatus) VALUES('" & Me.partNumber & "'," & Me.projectTemplateId.column(0) & ",'" & Me.projectStartDate & "',1)"
TempVars.Add "projId", db.OpenRecordset("SELECT @@identity")(0).Value

Set db = Nothing
Call createPartProject(TempVars!projId)

DoCmd.CLOSE acForm, "frmPartInitialize"
DoCmd.OpenForm "frmPartDashboard"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Me.sfrmPartInitialize.Visible = False
Me.partNumber = TempVars!partNumber

Dim tempType As Long

Select Case userData("Dept")
    Case "Project" 'new model
        tempType = 1
    Case "Service" 'service
        tempType = 2
End Select

If userData("Developer") Then
    Me.projectTemplateId.RowSource = "SELECT tblPartProjectTemplate.recordId, tblPartProjectTemplate.projectTitle From tblPartProjectTemplate WHERE tblPartProjectTemplate.recordId<>10 AND tblPartProjectTemplate.projectTitle Is Not Null"
Else
    Me.projectTemplateId.RowSource = "SELECT tblPartProjectTemplate.recordId, tblPartProjectTemplate.projectTitle From tblPartProjectTemplate WHERE tblPartProjectTemplate.recordId<>10 AND tblPartProjectTemplate.projectTitle Is Not Null AND templateType = " & tempType
End If

Me.sfrmPartInitialize_Pillars.Visible = False

dbExecute "DELETE * FROM tblSessionVariables WHERE pillarTitle is not null"
Me.sfrmPartInitialize_Pillars.Form.Requery

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub projectTemplateId_AfterUpdate()
On Error GoTo Err_Handler

'Me.sfrmPartInitialize.Visible = Me.projectTemplateId = 8 'new assembly

If IsNull(Me.opDate) Then
    MsgBox "Please fill out foundational pillar date first", vbInformation, "Woops"
    Me.projectTemplateId = Null
    Exit Sub
End If
    
Me.sfrmPartInitialize_Pillars.Visible = True


Dim db As Database
Set db = CurrentDb()

db.Execute "DELETE * FROM tblSessionVariables WHERE pillarTitle is not null"

Dim rsTemplate As Recordset, rsSess As Recordset
Set rsTemplate = db.OpenRecordset("SELECT * FROM tblPartStepTemplate WHERE gateTemplateId IN (SELECT recordId FROM tblPartGateTemplate WHERE projectTemplateId = " & Me.projectTemplateId & ") AND pillarStep = TRUE")
Set rsSess = db.OpenRecordset("tblSessionVariables")

If rsSess.RecordCount > 0 Then
    rsSess.MoveFirst
    Do While Not rsSess.EOF
        rsSess.Delete
        rsSess.MoveNext
    Loop
    rsSess.MoveFirst
End If

Do While Not rsTemplate.EOF
    
    rsSess.addNew
    rsSess!pillarTitle = rsTemplate!Title
    rsSess!pillarLength = rsTemplate!durationDays
    rsSess!pillarDue = addWorkdays(Me.opDate, rsTemplate![durationDays])
    rsSess!pillarStepId = rsTemplate!recordId
    rsSess.Update
    
    rsTemplate.MoveNext
Loop

rsSess.CLOSE
Set rsSess = Nothing
rsTemplate.CLOSE
Set rsTemplate = Nothing
Set db = Nothing

Me.sfrmPartInitialize_Pillars.Form.Requery

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
