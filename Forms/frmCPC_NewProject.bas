Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmbProjectType_AfterUpdate()
On Error GoTo Err_Handler

Dim projectType As String
Dim projectYear As String
Dim projectCount As String
Dim projectNumber As String

projectType = Nz(Me.cmbProjectType, "")
projectYear = Right(Year(Date), 2)

Select Case projectType
    Case "General"
        projectCount = Format(DCount("[projectNumber]", "tblCPC_Projects", "[projectNumber] LIKE 'P" & projectYear & "*'") + 1, "000")
        projectNumber = "P" & projectYear & "-" & projectCount
    Case "Material Change"
        projectCount = Format(DCount("[projectNumber]", "tblCPC_Projects", "[projectNumber] LIKE 'MC" & projectYear & "*'") + 1, "000")
        projectNumber = "MC" & projectYear & "-" & projectCount
    Case "Study"
        projectCount = Format(DCount("[projectNumber]", "tblCPC_Projects", "[projectNumber] LIKE 'S" & projectYear & "*'") + 1, "000")
        projectNumber = "S" & projectYear & "-" & projectCount
    Case Else
        projectNumber = ""
End Select

Me.txtProjectNumber = projectNumber
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub cmbChangeType_Enter()
On Error GoTo Err_Handler

If IsNull(Me.cmbProjectType) Then
    MsgBox "Please select a project type.", vbOKOnly, "Warning"
    Me.cmbProjectType.SetFocus
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub cmbLocation_Enter()
On Error GoTo Err_Handler

If IsNull(Me.cmbProjectType) Then
    MsgBox "Please select a project type.", vbOKOnly, "Warning"
    Me.cmbProjectType.SetFocus
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub cmbOwner_Enter()
On Error GoTo Err_Handler

If IsNull(Me.cmbProjectType) Then
    MsgBox "Please select a project type.", vbOKOnly, "Warning"
    Me.cmbProjectType.SetFocus
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Err_Handler

If validate <> "" Then
    If MsgBox("Are you sure?" & vbNewLine & "Your current record will be deleted.", vbYesNo, "Please confirm") <> vbYes Then
        Cancel = True
        Exit Sub
    End If

    DoCmd.SetWarnings False
    If Nz(Me.ID) <> "" Then DoCmd.RunCommand acCmdDeleteRecord
    DoCmd.SetWarnings True
End If


Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Unload", Err.DESCRIPTION, Err.number)
End Sub

Private Sub sfrmECOs_Enter()
On Error GoTo Err_Handler

If IsNull(Me.cmbProjectType) Then
    MsgBox "Please select a project type.", vbOKOnly, "Warning"
    Me.cmbProjectType.SetFocus
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub sfrmPartNumbers_Enter()
On Error GoTo Err_Handler

If IsNull(Me.cmbProjectType) Then
    MsgBox "Please select a project type.", vbOKOnly, "Warning"
    Me.cmbProjectType.SetFocus
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub txtDescription_Enter()
On Error GoTo Err_Handler

If IsNull(Me.cmbProjectType) Then
    MsgBox "Please select a project type.", vbOKOnly, "Warning"
    Me.cmbProjectType.SetFocus
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub btnECO_Click()
On Error GoTo Err_Handler

Dim db As DAO.Database
Dim rsECO As DAO.Recordset
Dim sqlMain As String
Dim sqlWhere As String
Dim rowsAdded As Long

Set db = CurrentDb
Set rsECO = db.OpenRecordset("SELECT ecoNumber FROM tblCPC_ECOs WHERE projectId = " & Me.ID, dbOpenSnapshot)

db.Execute "DELETE * FROM tblCPC_Parts WHERE projectId = " & Me.ID

sqlMain = "INSERT INTO tblCPC_Parts (projectId, partNumber, newRev, unit) " & _
    "SELECT " & Me.ID & ", APPS_MTL_SYSTEM_ITEMS.SEGMENT1 AS partNumber, ENG_ENG_REVISED_ITEMS.NEW_ITEM_REVISION AS newRev, APPS_MTL_CATEGORIES_VL.SEGMENT1 AS unit " & _
    "FROM ENG_ENG_REVISED_ITEMS INNER JOIN ((APPS_MTL_CATEGORIES_VL INNER JOIN INV_MTL_ITEM_CATEGORIES ON APPS_MTL_CATEGORIES_VL.CATEGORY_ID = INV_MTL_ITEM_CATEGORIES.CATEGORY_ID) " & _
    "INNER JOIN APPS_MTL_SYSTEM_ITEMS ON INV_MTL_ITEM_CATEGORIES.INVENTORY_ITEM_ID = APPS_MTL_SYSTEM_ITEMS.INVENTORY_ITEM_ID) ON ENG_ENG_REVISED_ITEMS.REVISED_ITEM_ID = " & _
    "APPS_MTL_SYSTEM_ITEMS.INVENTORY_ITEM_ID " & _
    "GROUP BY APPS_MTL_SYSTEM_ITEMS.SEGMENT1, ENG_ENG_REVISED_ITEMS.NEW_ITEM_REVISION, APPS_MTL_CATEGORIES_VL.SEGMENT1, ENG_ENG_REVISED_ITEMS.CHANGE_NOTICE "

If Not (rsECO.BOF And rsECO.EOF) Then
    rsECO.MoveFirst
    Do While Not rsECO.EOF
        If rsECO("ecoNumber") <> "" Then
            sqlWhere = "HAVING APPS_MTL_CATEGORIES_VL.SEGMENT1 Like 'U*' AND ENG_ENG_REVISED_ITEMS.CHANGE_NOTICE='" & rsECO("ecoNumber") & "';"
            db.Execute sqlMain & sqlWhere, dbFailOnError
        End If
        rsECO.MoveNext
    Loop
Else
    MsgBox "No ECO entered.", vbOKOnly, "Warning"
    Exit Sub
End If

rowsAdded = db.RecordsAffected

If rowsAdded = 0 Then
    MsgBox "This ECO does not have any revised items.", vbOKOnly, "Warning"
Else
    Me.sfrmPartNumbers.Requery
End If

rsECO.CLOSE
Set rsECO = Nothing
Set db = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Function validate() As String
validate = ""

Select Case True
    Case Nz(Me.cmbProjectType) = ""
        validate = "Please select a project type."
    Case Nz(Me.cmbChangeType) = ""
        validate = "Please select a change type."
    Case Nz(Me.cmbLocation) = ""
        validate = "Please select a location."
    Case Nz(Me.kickoffDate) = ""
        validate = "Please select a Kickoff Date."
    Case Nz(Me.userName) = ""
        validate = "Please select a project owner."
    Case Form_sfrmCPC_NewProjectParts.RecordsetClone.RecordCount = 0
        validate = "Please enter a part number."
    Case Nz(Me.txtDescription) = ""
        validate = "Please enter a description."
End Select

End Function

Private Sub btnSave_Click()
On Error GoTo Err_Handler

Dim val
val = validate
If val <> "" Then
    MsgBox val, vbInformation, "Warning..."
    Exit Sub
End If

If Me.Dirty Then Me.Dirty = False

Dim opT0 As Date

Dim db As DAO.Database
Set db = CurrentDb()
Dim rsProject As Recordset, rsStepTemplate As Recordset, rsApprovalsTemplate As Recordset
Dim strInsert As String, strInsert1 As String
Dim projTempId As Long, pillarDue As Date

Set rsProject = db.OpenRecordset("SELECT * from tblCPC_Projects WHERE ID = " & Me.ID)

'find the pillar step that is ZERO, then calculate everything from there
'this is called the foundational pillar

projTempId = 1

Set rsStepTemplate = db.OpenRecordset("SELECT * from tblCPC_Steps_Template WHERE [projectTemplateId] = " & projTempId & " ORDER BY indexOrder Asc", dbOpenSnapshot)
rsStepTemplate.FindFirst "pillarStep = TRUE AND duration = 0"

opT0 = Me.kickoffDate

If DCount("ID", "tblCPC_XFteams", "projectId = " & Me.ID & " AND memberName = '" & Environ("username") & "'") = 0 Then db.Execute "INSERT INTO tblCPC_XFteams(projectId,memberName) VALUES (" & Me.ID & ",'" & Environ("username") & "')", dbFailOnError 'assign project engineer

'--ADD STEPS
rsStepTemplate.MoveFirst
Do While Not rsStepTemplate.EOF
    If (IsNull(rsStepTemplate![stepName]) Or rsStepTemplate![stepName] = "") Then GoTo nextStep
        'if NOT pillar, don't add date!
        If rsStepTemplate!pillarStep Then
            pillarDue = addWorkdays(opT0, rsStepTemplate!duration)
            strInsert = "INSERT INTO tblCPC_Steps" & _
                "(projectId,stepName,openedBy,status,openedDate,lastUpdatedDate,lastUpdatedBy,stepActionId,documentType,responsible,indexOrder,duration,dueDate) VALUES"
            strInsert = strInsert & "(" & Me.ID & ",'" & StrQuoteReplace(rsStepTemplate![stepName]) & "','" & _
                Environ("username") & "','Not Started','" & Now() & "','" & Now() & "','" & Environ("username") & "',"
            strInsert = strInsert & Nz(rsStepTemplate![stepActionId], "NULL") & "," & Nz(rsStepTemplate![documentType], "NULL") & ",'" & _
                Nz(rsStepTemplate![responsible], "") & "'," & rsStepTemplate![indexOrder] & "," & Nz(rsStepTemplate![duration], 1) & ",'" & pillarDue & "');"
        Else
             strInsert = "INSERT INTO tblCPC_Steps" & _
                "(projectId,stepName,openedBy,status,openedDate,lastUpdatedDate,lastUpdatedBy,stepActionId,documentType,responsible,indexOrder,duration) VALUES"
            strInsert = strInsert & "(" & Me.ID & ",'" & StrQuoteReplace(rsStepTemplate![stepName]) & "','" & _
                Environ("username") & "','Not Started','" & Now() & "','" & Now() & "','" & Environ("username") & "',"
            strInsert = strInsert & Nz(rsStepTemplate![stepActionId], "NULL") & "," & Nz(rsStepTemplate![documentType], "NULL") & ",'" & _
                Nz(rsStepTemplate![responsible], "") & "'," & rsStepTemplate![indexOrder] & "," & Nz(rsStepTemplate![duration], 1) & ");"
        End If
        
    db.Execute strInsert, dbFailOnError
    
    '--ADD APPROVALS FOR THIS STEP
    TempVars.Add "stepId", db.OpenRecordset("SELECT @@identity")(0).Value
    Set rsApprovalsTemplate = db.OpenRecordset("SELECT * FROM tblPartStepTemplateApprovals WHERE [stepTemplateId] = " & rsStepTemplate![ID], dbOpenSnapshot)
    
    Do While Not rsApprovalsTemplate.EOF
        strInsert1 = "INSERT INTO tblCPC_Approvals_Template(requestedBy,requestedOn,dept,reqLevel,stepId) VALUES ('" & _
            Environ("username") & "','" & Now() & "','" & _
            Nz(rsApprovalsTemplate![dept], "") & "','" & Nz(rsApprovalsTemplate![reqLevel], "") & "'," & TempVars!stepId & ");"
        db.Execute strInsert1
        rsApprovalsTemplate.MoveNext
    Loop
nextStep:
    rsStepTemplate.MoveNext
Loop

DoEvents

rsStepTemplate.CLOSE
Set rsStepTemplate = Nothing
Set db = Nothing

Call registerCPCUpdates("tblCPC_Projects", Me.ID, "Project_Creation", "", "Project_Creation", Me.ID)
DoCmd.CLOSE

If CurrentProject.AllForms("frmCPC_WorkTracker").IsLoaded Then Form_frmCPC_WorkTracker.Requery

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
