Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnAssigneeJudgement_Click()
On Error GoTo Err_Handler

Me.assigneeJudgement = Not Me.assigneeJudgement

Call registerDRSUpdates("tblDesignChecksheet", Me.controlNumber, "assigneeJudgement", Me.assigneeJudgement.OldValue, Me.assigneeJudgement, Me.reviewItem, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub btnCheckerJudgement_Click()
On Error GoTo Err_Handler

Me.checkerJudgement = Not Me.checkerJudgement

Call registerDRSUpdates("tblDesignChecksheet", Me.controlNumber, "checkerJudgement", Me.checkerJudgement.OldValue, Me.checkerJudgement, Me.reviewItem, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub cmbPartType_Enter()
On Error GoTo Err_Handler

Me.cmbPartType.tag = Me.cmbPartType.Value

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub cmbPartType_AfterUpdate()
On Error GoTo Err_Handler

Dim controlNum
controlNum = Form_frmDRSdashboard.Control_Number

If DCount("recordId", "tblDesignChecksheet", "controlNumber = " & controlNum) = 0 Then
    createChecksheet
    Me.Requery
    
    Call registerDRSUpdates("tblDesignChecksheet", controlNum, Me.ActiveControl.name, Me.cmbPartType.tag, Me.cmbPartType, Me.reviewItem, Me.name)
    Exit Sub
End If

If MsgBox("This will reset the current checksheet, inlcuding all comments." & vbCrLf & "Do you want to continue?", vbYesNo, "Reset Checksheet") = vbYes Then
    Dim db As Database
    Set db = CurrentDb()
    Dim rsChecksheet As Recordset
    Set rsChecksheet = db.OpenRecordset("SELECT * FROM tblDesignChecksheet WHERE controlNumber = " & controlNum)

    Do While Not rsChecksheet.EOF
        rsChecksheet.Delete
        rsChecksheet.MoveNext
    Loop
    
    rsChecksheet.CLOSE
    Set rsChecksheet = Nothing
    Set db = Nothing
    
    Call registerDRSUpdates("tblDesignChecksheet", controlNum, "Drawing Checksheet", "", "Deleted", "Deleted Checksheet", "frmDesignChecksheet")
    
    createChecksheet
    Me.Requery
    
    Call registerDRSUpdates("tblDesignChecksheet", controlNum, Me.ActiveControl.name, Me.cmbPartType.tag, Me.cmbPartType, Me.reviewItem, Me.name)
Else
    Me.cmbPartType = Me.cmbPartType.tag
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Comments_AfterUpdate()
On Error GoTo Err_Handler

Call registerDRSUpdates("tblDesignChecksheet", Me.controlNumber, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.reviewItem, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub exportChecksheet_Click()
On Error GoTo Err_Handler

Dim controlNum, chkFold As String
controlNum = Me.controlNumber
chkFold = Nz(DLookup("Check_Folder", "tblDRStrackerExtras", "Control_Number = " & controlNum))

If Nz(chkFold) = "" Then
    Call snackBox("error", "No Check Folder", "You need a check folder set up to do this.", Me.name)
    Exit Sub
End If

Dim Y As String, z As String, tempFold As String, fso
Set fso = CreateObject("Scripting.FileSystemObject")
Y = addLastSlash(chkFold) & "2. " & controlNum & " Design Checksheet.pdf"

tempFold = getTempFold
If FolderExists(tempFold) = False Then MkDir (tempFold)

z = tempFold & controlNum & "TEMPdesignChecksheet.pdf"
DoCmd.OpenReport "rptDesignChecksheet", acViewPreview, , "[controlNumber]=" & controlNum, acHidden
DoCmd.OutputTo acOutputReport, "rptDesignChecksheet", acFormatPDF, z, False
DoCmd.CLOSE acReport, "rptDesignChecksheet"

Call fso.CopyFile(z, Y)
Call fso.deleteFile(z)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub history_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmHistory"
Form_frmHistory.RecordSource = "qryDRSupdateTracking"
Form_frmHistory.dataTag0.ControlSource = "dataTag0"
Form_frmHistory.previousData.ControlSource = "previous"
Form_frmHistory.newData.ControlSource = "new"
Form_frmHistory.filter = "tableRecordId = " & Me.controlNumber & " AND dataTag1 = 'frmDesignChecksheet'"
Form_frmHistory.FilterOn = True
Form_frmHistory.OrderBy = "updatedDate Desc"
Form_frmHistory.OrderByOn = True

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub createChecksheet()
On Error GoTo Err_Handler

Dim controlNum, drawingType, designResponsible, partType
controlNum = Form_frmDRSdashboard.Control_Number

If Me.txtDrawingType = "Internal Drawing" Then
    drawingType = 1
Else
    drawingType = 2 'Customer Drawing
End If

Select Case Me.txtDesignResponsible
    Case "NAM Responsible"
        designResponsible = 2
    Case "NJP Responsible"
        designResponsible = 1
    Case Else 'Customer only, or customer with our support
        designResponsible = 3
End Select

Select Case Me.cmbPartType
    Case "Assembled Part"
        partType = 2
    Case "Molded Part"
        partType = 1
    Case "Purchased Part"
        partType = 3
    Case Else
        MsgBox "Please enter a valid part type.", vbOKOnly, "Invalid Part Type"
        Exit Sub
End Select

If DCount("recordId", "tblDesignChecksheet", "controlNumber = " & controlNum) = 0 Then
    Dim db As Database
    Set db = CurrentDb()
    Dim rs1 As Recordset, rsChecksheet As Recordset
    Set rs1 = db.OpenRecordset("SELECT * from tblDesignChecksheetDefaults WHERE drawingType LIKE '*" & drawingType & "*' AND designResponsible LIKE '*" & designResponsible & "*' AND partType LIKE '*" & partType & "*' ORDER BY indexOrder Asc")
    Set rsChecksheet = db.OpenRecordset("tblDesignChecksheet")

    Do While Not rs1.EOF
        rsChecksheet.addNew

        rsChecksheet!controlNumber = controlNum
        rsChecksheet!reviewItem = rs1!recordId

        rsChecksheet.Update
        rs1.MoveNext
    Loop

    rs1.CLOSE
    Set rs1 = Nothing
    Set db = Nothing
    
    Call registerDRSUpdates("tblDesignChecksheet", controlNum, "Drawing Checksheet", "", "Created", "Created Checksheet", "frmDesignChecksheet")
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
