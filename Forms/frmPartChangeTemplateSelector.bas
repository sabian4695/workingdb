Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub selectTemplate_Click()
On Error GoTo Err_Handler

If MsgBox("This is will all the steps from this template to the CURRENT selected gate. Are you sure?", vbYesNo, "Please confirm!") <> vbYes Then Exit Sub

Dim templateId As Long, gateId As Long, projId As Long, pNum As String
templateId = Me.recordId
gateId = Me.partGateId
projId = Me.partProjectId
pNum = Me.partNumber

'for due date calc, find all CURRENT steps past selected step due date, and adjust by full change template amount, THEN calc the due dates in the template as they are made
Dim templateDuration As Long, startingDate As Date
templateDuration = DSum("duration", "tblPartStepTemplate", "gateTemplateId = " & templateId)
'startingDate = Form_sfrmPartDashboard.dueDate
Call recalcStepDueDates(projId, startingDate, templateDuration)

Dim db As DAO.Database
Set db = CurrentDb()

Dim rsStepTemplate As Recordset
Dim rsApprovalsTemplate As Recordset
Dim rsFiltApprovals As Recordset
Dim strInsert As String, strInsert1 As String, runningDate As Date

Set rsStepTemplate = db.OpenRecordset("SELECT * from tblPartStepTemplate WHERE gateTemplateId = " & templateId & " ORDER BY indexOrder Asc", dbOpenSnapshot)
Set rsApprovalsTemplate = db.OpenRecordset("tblPartStepTemplateApprovals", dbOpenSnapshot)
'runningDate = startingDate

'how much to offset all items
Dim templateIndexCount As Long, indexVal As Long
templateIndexCount = rsStepTemplate.RecordCount
indexVal = Form_sfrmPartDashboard.indexOrder

db.Execute "UPDATE tblPartSteps SET indexOrder = indexOrder + " & templateIndexCount & " WHERE partGateId = " & Form_sfrmPartDashboard.partGateId & " AND indexOrder > " & indexVal

Do While Not rsStepTemplate.EOF
    If (IsNull(rsStepTemplate![Title]) Or rsStepTemplate![Title] = "") Then GoTo nextStep
    indexVal = indexVal + 1
    'runningDate = addWorkdays(runningDate, Nz(rsStepTemplate![duration], 1))
    'runningDate = Null
    strInsert = "INSERT INTO tblPartSteps" & _
        "(partNumber,partProjectId,partGateId,stepType,openedBy,status,openDate,lastUpdatedDate,lastUpdatedBy,stepActionId,documentType,responsible,indexOrder,duration) VALUES"
    strInsert = strInsert & "('" & pNum & "'," & projId & "," & gateId & ",'" & StrQuoteReplace(rsStepTemplate![Title]) & "','" & _
        Environ("username") & "','Not Started','" & Now() & "','" & Now() & "','" & Environ("username") & "',"
    strInsert = strInsert & Nz(rsStepTemplate![stepActionId], "NULL") & "," & Nz(rsStepTemplate![documentType], "NULL") & ",'" & _
        Nz(rsStepTemplate![responsible], "") & "'," & indexVal & "," & Nz(rsStepTemplate![duration], 1) & ");"
    db.Execute strInsert, dbFailOnError
    
    '--ADD APPROVALS FOR THIS STEP
    If Not rsStepTemplate![approvalRequired] Then GoTo nextStep
    TempVars.Add "stepId", db.OpenRecordset("SELECT @@identity")(0).Value
    Set rsApprovalsTemplate = db.OpenRecordset("SELECT * FROM tblPartStepTemplateApprovals WHERE [stepTemplateId] = " & rsStepTemplate![recordId], dbOpenSnapshot)
    
    Do While Not rsApprovalsTemplate.EOF
        strInsert1 = "INSERT INTO tblPartTrackingApprovals(partNumber,requestedBy,requestedDate,dept,reqLevel,tableName,tableRecordId) VALUES ('" & _
            pNum & "','" & Environ("username") & "','" & Now() & "','" & _
            Nz(rsApprovalsTemplate![dept], "") & "','" & Nz(rsApprovalsTemplate![reqLevel], "") & "','tblPartSteps'," & TempVars!stepId & ");"
        db.Execute strInsert1
        rsApprovalsTemplate.MoveNext
    Loop
nextStep:
    rsStepTemplate.MoveNext
Loop

On Error Resume Next
rsStepTemplate.CLOSE
rsApprovalsTemplate.CLOSE
rsFiltApprovals.CLOSE
Set rsStepTemplate = Nothing
Set rsApprovalsTemplate = Nothing
Set rsFiltApprovals = Nothing

Set db = Nothing

On Error GoTo Err_Handler
DoCmd.CLOSE acForm, "frmPartChangeTemplateSelector"
Form_frmPartDashboard.partDash_refresh_Click

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
