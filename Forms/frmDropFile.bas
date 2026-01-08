Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnSave_Click()
On Error GoTo Err_Handler

Dim errorText As String
If Nz(Me.documentType, 0) = 0 Then errorText = "Need to select document type"
If Nz(Me.dragDrop) = "" Then errorText = "Please select a document to upload..."

If errorText <> "" Then
    MsgBox errorText, vbCritical, "Hold up"
    Exit Sub
End If

Dim fso As Object
Set fso = CreateObject("Scripting.FileSystemObject")

Dim fileExt As String, currentLoc As String, fullPath As String, attchFullFileName As String, tempFold As String, newFile As String
currentLoc = TempVars!dragDropLocation.Value

'transfer file to temp location
tempFold = getTempFold
fileExt = fso.GetExtensionName(currentLoc)
newFile = tempFold & "tempUpload" & nowString & "." & fileExt

If FolderExists(tempFold) = False Then MkDir (tempFold)
Call fso.CopyFile(currentLoc, newFile)

currentLoc = newFile

Dim testType As String
testType = "-" & Nz(Me.TtestType, "")
If testType = "-" Then testType = ""

Select Case True
    Case Me.documentType = 30 'OTHER document type
        Me.attachName = Me.customName & "-" & DMax("ID", "tblPartAttachmentsSP") + 1
    Case Me.TtestId <> "" 'Test/Trial attachment
        Me.attachName = DLookup("fileName", "tblPartAttachmentStandards", "recordId = " & Me.documentType) & "-" & DMax("ID", "tblPartAttachmentsSP") + 1
    Case Else 'step attachment
        Me.attachName = DLookup("fileName", "tblPartAttachmentStandards", "recordId = " & Me.documentType) & testType & "-" & DMax("ID", "tblPartAttachmentsSP") + 1
End Select

attchFullFileName = Replace(Me.attachName, " ", "_") & "." & fileExt
        
Dim testId, stepId, gateNum As Long
testId = Me.TtestId
If testId = "" Then testId = Null
stepId = Me.TstepId
If stepId = "" Then
    stepId = Null
    If Nz(Me.TprojectId, 0) <> 0 Then gateNum = CLng(Right(Left(DLookup("gateTitle", "tblPartGates", "projectId = " & Nz(Me.TprojectId, 0) & " AND actualDate is null"), 2), 1))
Else
    gateNum = CLng(Right(Left(DLookup("gateTitle", "tblPartGates", "recordId = " & DLookup("partGateId", "tblPartSteps", "recordId = " & Me.TstepId)), 2), 1))
End If

Dim db As DAO.Database
Dim rsPartAtt As DAO.Recordset
Dim rsPartAttChild As DAO.Recordset2
Set db = CurrentDb
Set rsPartAtt = db.OpenRecordset("tblPartAttachmentsSP", dbOpenDynaset)

rsPartAtt.addNew
rsPartAtt!fileStatus = "Created"

rsPartAtt.Update
rsPartAtt.MoveLast

rsPartAtt.Edit
Set rsPartAttChild = rsPartAtt.Fields("Attachments").Value

rsPartAttChild.addNew
Dim fld As DAO.Field2
Set fld = rsPartAttChild.Fields("FileData")
fld.LoadFromFile (currentLoc)
rsPartAttChild.Update

rsPartAtt!partNumber = Me.TpartNumber
rsPartAtt!testId = testId
rsPartAtt!partStepId = stepId
rsPartAtt!partProjectId = Nz(Me.TprojectId, 0)
rsPartAtt!documentType = CLng(Me.documentType)
rsPartAtt!uploadedBy = Environ("username")
rsPartAtt!uploadedDate = Now()
rsPartAtt!attachName = Me.attachName
rsPartAtt!attachFullFileName = attchFullFileName
rsPartAtt!fileStatus = "Uploading"
rsPartAtt!gateNumber = gateNum
rsPartAtt!documentTypeName = Me.documentType.column(1)
rsPartAtt!businessArea = DLookup("businessArea", "tblPartAttachmentStandards", "recordId = " & Me.documentType)
rsPartAtt.Update

Select Case Me.documentType
    Case 22 'test attachment
        Call registerPartUpdates("tblPartAttachmentsSP", Me.TtestId, "Test Attachment", Me.attachName, "Uploaded", Me.TpartNumber, Form_frmPartAttachments.itemName.Caption, Me.TprojectId)
    Case 32 'trial attachment
        Call registerPartUpdates("tblPartAttachmentsSP", Me.TtestId, "Trial Attachment", Me.attachName, "Uploaded", Me.TpartNumber, Form_frmPartAttachments.itemName.Caption, Me.TprojectId)
    Case Else 'step attachment
        Call registerPartUpdates("tblPartAttachmentsSP", Me.TstepId, "Step Attachment", Me.attachName, "Uploaded", Me.TpartNumber, Form_frmPartAttachments.itemName.Caption, Me.TprojectId)
End Select

On Error Resume Next
Set fld = Nothing
rsPartAttChild.CLOSE: Set rsPartAttChild = Nothing
rsPartAtt.CLOSE: Set rsPartAtt = Nothing
Set db = Nothing

DoCmd.CLOSE acForm, "frmDropFile"
Form_frmPartAttachments.Requery

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub dragDrop_AfterUpdate()
On Error GoTo Err_Handler

Me.refresh
Dim strText
Dim dragDropVal() As String, i As Integer, ITEM
dragDropVal = Split(Me.dragDrop.Value, "#")

i = 0
For Each ITEM In dragDropVal
    If i = 0 Then GoTo nextItem
    Select Case i
        Case 0
            GoTo nextItem
        Case 1
            strText = ITEM
        Case Else
            If ITEM <> "" Then strText = strText & "#" & ITEM
    End Select
nextItem:
    i = i + 1
Next ITEM

TempVars.Add "dragDropLocation", strText
Me.dragDropView = strText
Me.getFocus.SetFocus
If strText = "" Then
    MsgBox "Didn't get that - please try again", vbInformation, "Oh no!"
    Exit Sub
End If
MsgBox "Got it! Make sure details below are correct, then click Save + Close to Upload File", vbInformation, "Nice"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub filePicker_Click()
On Error GoTo Err_Handler

Dim strFile

With Application.FileDialog(msoFileDialogOpen)
    .Title = "Choose a File"
    .AllowMultiSelect = False
    .Show
    
    On Error Resume Next
    strFile = .SelectedItems(1)
End With

On Error GoTo Err_Handler

If IsNull(strFile) Then Exit Sub

Me.dragDrop = strFile
TempVars.Add "dragDropLocation", strFile
Me.dragDropView = strFile

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Me.dragDrop = ""
Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub
