Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Public nppfAddressFull As Variant
Public toolNum
Public partSelected
Public toolNumFirstTwo
Public toolNumFirstThree
Public nppfDirectory As Variant

'Public nppfAddressFull

Function afterUpdate_tblPartInfo()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartInfo", Me("tblPartInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)

Exit Function
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Function

Function afterUpdate_tblPartAppearanceInfo()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartAppearanceInfo", Me("tblPartAppearanceInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)

Exit Function
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Function

Private Sub APQPFormButton_Click()
On Error GoTo Err_Handler

Dim excelApp, targetWorkbook
Set excelApp = CreateObject("Excel.application", "")

excelApp.Visible = True
Set targetWorkbook = excelApp.Workbooks.open("\\shelbyville\data\Quality\1 - APQP\00 - APQP Templates\1. Auto Templates\1 Internal\" & GetExcelFileName("\\shelbyville\data\Quality\1 - APQP\00 - APQP Templates\1. Auto Templates\1 Internal", "1a"), ReadOnly:=True, UpdateLinks:=False)
targetWorkbook.Worksheets("General Info").Range("Tool_No") = Me.toolNumber

'If Me.assembly = True Then
'        targetWorkbook.Worksheets("General Info").Range("Assy") = "Yes"
'Else
'        targetWorkbook.Worksheets("General Info").Range("Assy") = "No"
'End If

If Me.familyTool = True Then
        targetWorkbook.Worksheets("General Info").Range("Fam") = "Yes"
Else
        targetWorkbook.Worksheets("General Info").Range("Fam") = "No"
End If

If Left(Me.[tblUnits.Org], 3) = "SLB" Then
    targetWorkbook.Worksheets("General Info").Range("outsourced") = "No"
Else
    targetWorkbook.Worksheets("General Info").Range("outsourced") = "Yes"
End If

If Me.familyTool = True Then
    targetWorkbook.Worksheets("General Info").Range("Cav") = (Me.cavitation.Value) & "+" & (Me.cavitation.Value)
    targetWorkbook.Worksheets("General Info").Range("NAM_NO_LH") = Left(Me.partNumber, 4) & Left(Right(Me.partNumber, 2), 1)
    targetWorkbook.Worksheets("General Info").Range("Desc") = Left(Me.DESCRIPTION, Len(Me.DESCRIPTION) - 3)
    targetWorkbook.Worksheets("General Info").Range("Desc_LH") = Left(Me.DESCRIPTION, Len(Me.DESCRIPTION) - 5) & "LH"
Else
    targetWorkbook.Worksheets("General Info").Range("Cav") = Me.cavitation
    targetWorkbook.Worksheets("General Info").Range("Desc") = Me.DESCRIPTION
End If

targetWorkbook.Worksheets("General Info").Range("NAM_NO") = Left(Me.partNumber, 5)

targetWorkbook.Worksheets("General Info").Range("Customer") = Me.customerId
targetWorkbook.Worksheets("General Info").Range("Cust_No") = Me.customerPN
targetWorkbook.Worksheets("General Info").Range("Model_Code") = Me.modelCode
targetWorkbook.Worksheets("General Info").Range("Mat") = Me.materialNumber
targetWorkbook.Worksheets("General Info").Range("Mat_Type") = Me.materialSymbol
targetWorkbook.Worksheets("General Info").Range("Colorant") = Me.materialNumber1
targetWorkbook.Worksheets("General Info").Range("Color") = Me.colorCode

'If Me.DeltaRFlammability = True Then
'    targetWorkbook.Worksheets("General Info").Range("deltaRFlammability") = "Yes"
'Else
'    targetWorkbook.Worksheets("General Info").Range("deltaRFlammability") = "No"
'End If
'
'If Me.DeltaRPart_Marking = True Then
'    targetWorkbook.Worksheets("General Info").Range("deltaRPartMarking") = "Yes"
'Else
'    targetWorkbook.Worksheets("General Info").Range("deltaRPartMarking") = "No"
'End If
'
'If Me.DeltaRDimension = True Then
'    targetWorkbook.Worksheets("General Info").Range("deltaRDimension") = "Yes"
'Else
'    targetWorkbook.Worksheets("General Info").Range("deltaRDimension") = "No"
'End If

'targetWorkbook.Worksheets("General Info").Range("qe") = Me.Quality_Engineer
'targetWorkbook.Worksheets("General Info").Range("te") = Me.Tooling_Engineer
'targetWorkbook.Worksheets("General Info").Range("pe") = Me.Project_Engineer
'targetWorkbook.Worksheets("General Info").Range("de") = Me.DevDesignEng


targetWorkbook.Worksheets("General Info").Range("H42") = Date
targetWorkbook.Worksheets("General Info").Range("G42") = UCase(Right(Environ("username"), 1) & Left(Environ("username"), 1))

Set targetWorkbook = Nothing
Set excelApp = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub autoTolerancerButton_Click()
On Error GoTo Err_Handler

Dim excelApp, targetWorkbook
Set excelApp = CreateObject("Excel.application", "")
excelApp.Visible = True
Set targetWorkbook = excelApp.Workbooks.open("\\shelbyville\data\Quality\1 - APQP\00 - APQP Templates\1. Auto Templates\1 Internal\" & GetExcelFileName("\\shelbyville\data\Quality\1 - APQP\00 - APQP Templates\1. Auto Templates\1 Internal", "6a"), ReadOnly:=True)

Set targetWorkbook = Nothing
Set excelApp = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub



Private Sub cfConceptButton_Click()
On Error GoTo Err_Handler

Dim partNum, thousZeros, FolName, mainPath, strPath

partNum = Me.partNumber
thousZeros = Left(partNum, 2) & "000\"
mainPath = mainFolder("slbProject")
strPath = mainPath & thousZeros

If Len(partNum) = 5 Then
    FolName = Dir(strPath & partNum & "*", vbDirectory)
    Call openPath(strPath & FolName & "\QUALITY DOCUMENTS\" & "CHECK FIXTURE INFO")
Else
    Call openPath(mainPath)
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub CFPacketButton_Click()
On Error GoTo Err_Handler

Dim excelApp, targetWorkbook
Set excelApp = CreateObject("Excel.application", "")
excelApp.Visible = True
Set targetWorkbook = excelApp.Workbooks.open("\\shelbyville\data\Quality\1 - APQP\00 - APQP Templates\1. Auto Templates\2 External\" & GetExcelFileName("\\shelbyville\data\Quality\1 - APQP\00 - APQP Templates\1. Auto Templates\2 External", "2a"), ReadOnly:=True)
targetWorkbook.Worksheets("INFO").Range("toolnum") = Me.partNumber
targetWorkbook.Worksheets("INFO").Range("partnum") = Me.customerPN
targetWorkbook.Worksheets("INFO").Range("partname") = Me.DESCRIPTION
'targetWorkbook.Worksheets("INFO").Range("nifcoqe") = Me.Quality_Engineer
targetWorkbook.Worksheets("INFO").Range("toolnum") = Me.partNumber
targetWorkbook.Worksheets("INFO").Range("model") = Me.modelCode
targetWorkbook.Worksheets("INFO").Range("toolnum") = Me.partNumber
targetWorkbook.Worksheets("INFO").Range("sigdate") = Date
targetWorkbook.Worksheets("INFO").Range("eci") = Me.customerRevLevel

Set targetWorkbook = Nothing
Set excelApp = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub cfQuoteButton_Click()
On Error Resume Next

If IsNull(Me.partNumber) Then Exit Sub
toolNum = Me.partNumber.Value
Dim programNum As Variant

programNum = DLookup("[Vehicle Program]", "tblToolsMain", "[Tool Number]= '" & toolNum & "'")

If programNum = "200D & 220D" Then
    Application.FollowHyperlink "\\shelbyville.nifcoam.local\data\Quality\1 - APQP\200D\200D_CF_QUOTE"

ElseIf programNum = "200D" Then
   Application.FollowHyperlink "\\shelbyville.nifcoam.local\data\Quality\1 - APQP\200D\200D_CF_QUOTE"

ElseIf programNum = "220D" Then
   Application.FollowHyperlink "\\shelbyville.nifcoam.local\data\Quality\1 - APQP\200D\200D_CF_QUOTE"

Else
   Application.FollowHyperlink "\\shelbyville.nifcoam.local\data\Quality\1 - APQP\" & programNum & "\" & programNum & "_CF_QUOTE"

End If

End Sub

Private Sub CFRequired_AfterUpdate()
On Error GoTo Err_Handler

If Me.checkFixture = True Then
    'Me.cfReqOrdered.Visible = True
    'Me.CFconcept.Visible = True
    'Me.CFApproval.Visible = True
    'Me.noCFLabel.Visible = False
Else
    'Me.cfReqOrdered.Visible = False
    'Me.CFconcept.Visible = False
    'Me.CFApproval.Visible = False
    'Me.noCFLabel.Visible = True
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub cfScheduleButton_Click()
On Error Resume Next

If IsNull(Me.partNumber) Then Exit Sub
toolNum = Me.partNumber.Value
Dim programNum As Variant
Dim programNumWithExtra

programNum = DLookup("[Vehicle Program]", "tblToolsMain", "[Tool Number]= '" & toolNum & "'")

    Dim objFS As Variant
    Dim objfolder As Variant
    Dim objFile As Variant
    Dim intLengthOfPartialName As Integer
    Dim strfilenamefull As String
    
    Set objFS = CreateObject("Scripting.FileSystemObject")
    
    If programNum = "200D & 220D" Or programNum = "200D" Or programNum = "220D" Then
        Set objfolder = objFS.GetFolder("\\shelbyville.nifcoam.local\data\Quality\1 - APQP\200D")
        programNum = "200D"
        programNumWithExtra = programNum & "_CF_Schedule"
    Else
        Set objfolder = objFS.GetFolder("\\shelbyville.nifcoam.local\data\Quality\1 - APQP\" & programNum)
        programNumWithExtra = programNum & "_CF_Schedule"
    End If
    
    'work out how long the partial file name is
    intLengthOfPartialName = Len(programNumWithExtra)
    
    For Each objFile In objfolder.Files
    
        'Test to see if the first 4 digits of the folder name matches the input
        If Left(objFile.name, intLengthOfPartialName) = programNumWithExtra Then
            strfilenamefull = objFile.name 'get the full file name
        
    Exit For
    Else
    End If
    
    Next objFile
    
If programNum = "200D & 220D" Or programNum = "200D" Or programNum = "220D" Then
    Application.FollowHyperlink "\\shelbyville.nifcoam.local\data\Quality\1 - APQP\200D\" & strfilenamefull 'add name of schedule file .xlsx

'ElseIf programNum = "200D" Then
'   Application.FollowHyperlink "\\shelbyville.nifcoam.local\data\Quality\1 - APQP\200D"
'
'ElseIf programNum = "220D" Then
'   Application.FollowHyperlink "\\shelbyville.nifcoam.local\data\Quality\1 - APQP\200D"

Else
   Application.FollowHyperlink "\\shelbyville.nifcoam.local\data\Quality\1 - APQP\" & programNum & "\" & strfilenamefull
End If

End Sub

Private Sub colorantCertButton_Click()
On Error GoTo Err_Handler

Application.FollowHyperlink "\\shelbyville\data\Quality\Receiving\Incoming Logs\Raw Material & Certs\color\Color\Cert\" & Nz(Me.ColorantNumber)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub


Private Sub custdwgButton_Click()
On Error GoTo Err_Handler

Form_DASHBOARD.partNumberSearch = Me.partNumber
Call Form_DASHBOARD.openCust_Click

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub customerSDSButton_Click()
On Error GoTo Err_Handler
Dim partNum, thousZeros, FolName, mainPath, strPath

partNum = Me.partNumber
thousZeros = Left(partNum, 2) & "000\"
mainPath = mainFolder("slbProject")
strPath = mainPath & thousZeros

If Len(partNum) = 5 Then
    FolName = Dir(strPath & partNum & "*", vbDirectory)
    Call openPath(strPath & FolName & "\QUALITY DOCUMENTS\" & "SDS FOLDER\" & "Customer Drawing SDS")
Else
    Call openPath(mainPath)
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub dcDataButton_Click()
On Error GoTo Err_Handler

Application.FollowHyperlink "\\design.nifcoam.local\data\Documents\Data\INTERNAL\Send\" & Nz(Me.partNumber)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub docHistoryButton_Click()
On Error GoTo Err_Handler

Call openDocumentHistoryFolder(Me.partNumber)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub engineeringStandardsButton_Click()
On Error GoTo Err_Handler

Application.FollowHyperlink "\\design\data\Documents\Engineering_Standards\Customers\Toyota"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub



Private Sub gageCalibrationButton_Click()
On Error Resume Next

Dim excelApp, targetWorkbook
Set excelApp = CreateObject("Excel.application", "")
excelApp.Visible = True
Set targetWorkbook = excelApp.Workbooks.open("\\shelbyville\data\Quality\1 - APQP\00 - APQP Templates\1. Auto Templates\1 Internal\" & GetExcelFileName("\\shelbyville\data\Quality\1 - APQP\00 - APQP Templates\1. Auto Templates\1 Internal", "13a"), ReadOnly:=True)

targetWorkbook.Worksheets("Sheet1").Range("todaysDate") = Date
'targetWorkbook.Worksheets("Sheet1").Range("qe") = Me.Quality_Engineer
targetWorkbook.Worksheets("Sheet1").Range("namNo") = Me.toolNum
targetWorkbook.Worksheets("Sheet1").Range("custNum") = Me.customerPN
targetWorkbook.Worksheets("Sheet1").Range("customer") = Me.customerId
targetWorkbook.Worksheets("Sheet1").Range("program") = Me.programId

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)

End Sub

Private Sub Form_Current() 'runs when a new record is opened/filtered to or the form is opened again
On Error Resume Next


Form_sfrmNMQDashboard.setFilter

'Pops the current part number into the testing tracker filter box
forms!frmNMQDashboard!sfrmPartTestingTracker.Form!fltPartNumber = Me.partNumberBox

If summonMike = True Then
    mascotImage.Visible = True
End If

If summonMike = False Then
    mascotImage.Visible = False
End If


'Pulling up the part image
If Not Me.partNumber = "" Then
    partPhotoImage.Visible = True
    partPhotoImage.Picture = "\\data\mdbdata\WorkingDB\_docs\Part_Pictures\" & Me.partNumber.Value & ".png"
Else
    partPhotoImage.Visible = False
End If


'---Approval / conditional approval label on part photo----
If IsNull(Me.PPAPapproval) Then
    PPAPApprovedLabel.Caption = "Not Approved"
    PPAPApprovedLabel.ForeColor = rgb(200, 75, 50)
Else
    PPAPApprovedLabel.Caption = "Approved!"
    PPAPApprovedLabel.ForeColor = rgb(0, 200, 50)
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Dim db As Database
Set db = CurrentDb()
Dim rsPermissions As Recordset
Dim rsOrgNames As Recordset

'Mike-related
mascotImage.Visible = False
wisdomText.Visible = False
summonMike.Caption = "Summon Mike"

Set rsPermissions = db.OpenRecordset("SELECT * from tblPermissions where user = '" & Environ("username") & "'")

Me.orgLabel.Caption = DLookup("permissionsLocation", "tblDropDownsSP", "recordid = " & rsPermissions!Org)
helloBox.Caption = getFullName(firstOnly:=True) 'get first name

rsPermissions.CLOSE

Set rsPermissions = Nothing
Set db = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub internalSDSButton_Click()
On Error GoTo Err_Handler

Form_DASHBOARD.partNumberSearch = Me.partNumber
Form_DASHBOARD.slbSDS_Click

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub IPDbutton_Click()
On Error GoTo Err_Handler

Form_DASHBOARD.partNumberSearch = Me.partNumber
Call Form_DASHBOARD.openInternalDwg_Click

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Function GetExcelFileName(directory, searchTerm)
On Error GoTo Err_Handler

Dim objFS As Variant
Dim objfolder As Variant
Dim objFile As Variant
Dim intLengthOfPartialName As Integer
Set objFS = CreateObject("Scripting.FileSystemObject")
Set objfolder = objFS.GetFolder(directory)
Dim strfilenamefull As String

'work out how long the partial file name is
intLengthOfPartialName = Len(searchTerm)

For Each objFile In objfolder.Files
    'Test to see if the first 5 digits of the pdf name matches the input
    If Left(objFile.name, intLengthOfPartialName) = searchTerm Then
        strfilenamefull = objFile.name 'get the full file name
Exit For
Else
End If

Next objFile

GetExcelFileName = strfilenamefull

Exit Function
Err_Handler:
    Call handleError(Me.name, "GetExcelFileName", Err.DESCRIPTION, Err.number)
End Function

Private Sub ISButton_Click()
On Error GoTo Err_Handler
Dim partNum, thousZeros, FolName, mainPath, strPath

partNum = Me.partNumber
thousZeros = Left(partNum, 2) & "000\"
mainPath = mainFolder("slbProject")
strPath = mainPath & thousZeros

If Len(partNum) = 5 Then
    FolName = Dir(strPath & partNum & "*", vbDirectory)
    Call openPath(strPath & FolName & "\QUALITY DOCUMENTS\" & "INSPECTION STANDARD FOLDER")
Else
    Call openPath(mainPath)
End If
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub konotesButton_Click()
On Error GoTo Err_Handler
Dim partNum, thousZeros, FolName, mainPath, strPath

partNum = Me.partNumber
thousZeros = Left(partNum, 2) & "000\"
mainPath = mainFolder("slbProject")
strPath = mainPath & thousZeros

If Len(partNum) = 5 Then
    FolName = Dir(strPath & partNum & "*", vbDirectory)
    Call openPath(strPath & FolName & "\PROJECT DOCUMENTS\Kick Off meeting notes")
Else
    Call openPath(mainPath)
End If
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
Application.FollowHyperlink nppfAddressFull
End Sub

Private Sub Labdatabutton_Click()
On Error Resume Next

If IsNull(partSelected) Then Exit Sub

Dim objFS As Variant
Dim objfolder As Variant
Dim objFile As Variant
Dim intLengthOfPartialName As Integer
Dim strfilenamefull As String
Dim strNumberPlusT As String
Dim intLengthOfPartialNamePlusT As Integer
Dim directory


toolNumFirstTwo = Left(partSelected, 2) & "000"
toolNumFirstThree = Left(partSelected, 3) & "00"
directory = "\\shelbyville\data\Quality\20 - Lab Data\" & toolNumFirstTwo & "\" & toolNumFirstThree

Set objFS = CreateObject("Scripting.FileSystemObject")
Set objfolder = objFS.GetFolder(directory)

'work out how long the partial file name is
intLengthOfPartialName = 5 'Len(partSelected)

For Each objfolder In objfolder.SubFolders
    'Test to see if the first 5 digits of the folder name matches the input
    If Left(objfolder.name, intLengthOfPartialName) = partNumberBox.Value Then
        Application.FollowHyperlink "\\shelbyville\data\Quality\20 - Lab Data\" & toolNumFirstTwo & "\" & toolNumFirstThree & "\" & objfolder.name 'get the full file name
Exit For
Else
End If

Next objfolder

End Sub

Private Sub masterPartListButton_Click()
On Error GoTo Err_Handler

Application.FollowHyperlink "\\shelbyville\data\Quality\1 - APQP\Master Parts List"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub mascotImage_DblClick(Cancel As Integer)
On Error Resume Next
Dim randomNumber As Integer
Dim wisdom As String

randomNumber = Int((10 - 1 + 1) * Rnd + 1)

Select Case randomNumber
    Case 1
     wisdom = "The strength of a car lies not in its metal, but in the resilience of its plastic parts."
     Case 2
        wisdom = "Just as the Stoics found wisdom in simplicity, so too does the efficiency of a vehicle depend on the quality of its plastic components."
    Case 3
        wisdom = "A well-crafted plastic part can withstand the test of time, much like a well-trained mind can endure life's challenges."
    Case 4
        wisdom = "In the world of automobiles, the smallest plastic part can make the biggest difference, just as small actions can lead to great wisdom."
    Case 5
        wisdom = "The innovation in plastic car parts is a reminder that progress often comes from unexpected places, much like wisdom from life's challenges."
    Case 6
        wisdom = "In the assembly of a car, every plastic part has its purpose, just as every experience contributes to our growth."
    Case 7
        wisdom = "Begin at once to live, and count each separate day as a separate life."
    Case 8
        wisdom = "Man is not worried by real problems so much as by his imagined anxieties about real problems."
    Case 9
        wisdom = "It is not death that a man should fear, but he should fear never beginning to live."
    Case 10
        wisdom = "Not what we have but what we enjoy constitutes our abundance."
    Case Else
        wisdom = "Mike grows very tired, please try again tomorrow."
End Select

wisdomText.Caption = "Mike says: " & vbNewLine & wisdom
End Sub

Private Sub materialCertButton_Click()
On Error GoTo Err_Handler

Application.FollowHyperlink "https://nifcoam.sharepoint.com/sites/logistics/inbound/Material%20Certifications/Forms/AllItems.aspx?view=7&q=" & Nz(Me.materialNumber)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub materialTestInfoButton_Click()
On Error GoTo Err_Handler

'DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12, "matTestQ", "\\shelbyville\data\Quality\1 - APQP\Material Test Info\materialTestingInfo.xlsx"
'DoCmd.OutputTo acOutputQuery, "matTestQ", acFormatXLSX, , True

'Dim dbs As Database

'Set dbs = CurrentDb


'Set rsQuery = dbs.OpenRecordset("matTestQ")

'Set excelApp = CreateObject("Excel.application", "")
'excelApp.Visible = True
'Set targetWorkbook = excelApp.workbooks.Open("\\shelbyville\data\Quality\1 - APQP\Material Test Info\materialTestingInfo.xlsx")
'targetWorkbook.Worksheets("data").Range("A2").CopyFromRecordset rsQuery

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub matingPanelRequest_Click()
On Error GoTo Err_Handler

Dim excelApp, targetWorkbook
Set excelApp = CreateObject("Excel.application", "")
excelApp.Visible = True
Set targetWorkbook = excelApp.Workbooks.open("\\shelbyville\data\Quality\1 - APQP\00 - APQP Templates\1. Auto Templates\1 Internal\" & GetExcelFileName("\\shelbyville\data\Quality\1 - APQP\00 - APQP Templates\1. Auto Templates\1 Internal", "8a"), ReadOnly:=True)
targetWorkbook.Worksheets("General Info").Range("PartNumber") = Me.partNumber
targetWorkbook.Worksheets("General Info").Range("PartName") = Me.DESCRIPTION

Set targetWorkbook = Nothing
Set excelApp = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub matTestButton_Click()
On Error GoTo Err_Handler

Application.FollowHyperlink "\\shelbyville\data\Quality\24 - Material Test Data\" & Nz(Me.materialNumber)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub moldbookButton_Click()
On Error GoTo Err_Handler

Dim partNum, thousZeros, FolName, mainPath, strPath

partNum = Me.partNumber
thousZeros = Left(partNum, 2) & "000\"
mainPath = mainFolder("slbProject")
strPath = mainPath & thousZeros

If Len(partNum) = 5 Then
    FolName = Dir(strPath & partNum & "*", vbDirectory)
    Call openPath(strPath & FolName & "\TRIAL DOCUMENTS\PASS DOCUMENTS")
Else
    Call openPath(mainPath)
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
 
Private Sub njpPPAPButton_Click()
On Error GoTo Err_Handler

Dim excelApp, targetWorkbook
Set excelApp = CreateObject("Excel.application", "")
excelApp.Visible = True
Set targetWorkbook = excelApp.Workbooks.open("\\shelbyville\data\Quality\1 - APQP\00 - APQP Templates\1. Auto Templates\1 Internal\" & GetExcelFileName("\\shelbyville\data\Quality\1 - APQP\00 - APQP Templates\1. Auto Templates\1 Internal", "7a"), ReadOnly:=True)

Set targetWorkbook = Nothing
Set excelApp = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub nmqHistory_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmHistory", acNormal, , "[partNumber] = '" & Me.partNumber & "' AND dataTag1 = 'frmNMQDashboard'"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub NPPFButton_Click()
On Error GoTo Err_Handler
Dim partNum, thousZeros, FolName, mainPath, strPath

partNum = Me.partNumber
thousZeros = Left(partNum, 2) & "000\"
mainPath = mainFolder("slbProject")
strPath = mainPath & thousZeros

If Len(partNum) = 5 Then
    FolName = Dir(strPath & partNum & "*", vbDirectory)
    Call openPath(strPath & FolName)
Else
    Call openPath(mainPath)
End If
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Function GetFullFileName()
On Error GoTo Err_Handler

Dim objFS As Variant
Dim objfolder As Variant
Dim objFile As Variant
Dim intLengthOfPartialName As Integer
Dim strfilenamefull As String
Dim strNumberPlusT As String
Dim intLengthOfPartialNamePlusT As Integer

toolNum = Me.partNumber.Value
toolNumFirstTwo = Left(toolNum, 2) & "000"
nppfDirectory = "\\shelbyville\data\New Project Part Folders\" & toolNumFirstTwo
strNumberPlusT = toolNum & "T"


Set objFS = CreateObject("Scripting.FileSystemObject")
Set objfolder = objFS.GetFolder(nppfDirectory)

'work out how long the partial file name is
intLengthOfPartialName = Len(toolNum)
intLengthOfPartialNamePlusT = Len(strNumberPlusT)


For Each objfolder In objfolder.SubFolders

    'Test to see if the first 5 digits of the folder name matches the input
    If Left(objfolder.name, intLengthOfPartialName) = toolNum Then

        strfilenamefull = objfolder.name 'get the full file name

    'Test to see if the 5 or 6 digits of the folder name matches the input
    ElseIf Right(objfolder.name, intLengthOfPartialNamePlusT) = strNumberPlusT Then

        strfilenamefull = objfolder.name 'pulls the full file name

Exit For

Else

End If

Next objfolder

'Return the full file name as the function's value
GetFullFileName = strfilenamefull

Exit Function
Err_Handler:
    Call handleError(Me.name, "GetFullFileName", Err.DESCRIPTION, Err.number)
End Function

Function GetFolderName(directory As String)
On Error GoTo Err_Handler

Dim objFS As Variant
Dim objfolder As Variant
Dim objFile As Variant
Dim intLengthOfPartialName As Integer
Dim strfilenamefull As String

toolNum = Me.partNumber.Value

Set objFS = CreateObject("Scripting.FileSystemObject")
Set objfolder = objFS.GetFolder(directory)

'work out how long the partial file name is
intLengthOfPartialName = Len(toolNum)

For Each objfolder In objfolder.SubFolders
    'Test to see if the first 5 digits of the folder name matches the input
    If Left(objfolder.name, intLengthOfPartialName) = toolNum Then
        GetFolderName = objfolder.name 'get the full file name
Exit For

Else

End If

Next objfolder

'Return the full file name as the function's value
Exit Function
Err_Handler:
    Call handleError(Me.name, "GetFolderName", Err.DESCRIPTION, Err.number)
End Function

Private Sub ohq_Click()
On Error GoTo Err_Handler

Dim filterVal, pNum
filterVal = idNAM(Nz(Me.partNumber), "NAM")
If filterVal = "" Then filterVal = "44388"
filterVal = "[INVENTORY_ITEM_ID] = " & filterVal

DoCmd.OpenForm "frmOnHandQty", , , filterVal
Form_frmOnHandQty.NAMsrchBox = Nz(Me.partNumber)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub openDash_Click()
On Error GoTo Err_Handler

openPartProject (Me.partNumber)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub openMasterSetup_Click()
On Error GoTo Err_Handler

Form_DASHBOARD.partNumberSearch = Me.partNumber
Form_DASHBOARD.cnlMasterSetups.SetFocus
Form_DASHBOARD.openMasterSetups

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub openTestingTracker_Click()
On Error GoTo Err_Handler

If (CurrentProject.AllForms("frmPartTestingTracker").IsLoaded = True) Then DoCmd.CLOSE acForm, "frmPartTestingTracker"
DoCmd.OpenForm "frmPartTestingTracker", , , "partNumber = '" & Me.partNumber & "'"

Form_frmPartTestingTracker.fltPartNumber = Me.partNumber

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub PAPacketButton_Click()
On Error Resume Next
Dim excelApp, targetWorkbook
Set excelApp = CreateObject("Excel.application", "")

excelApp.Visible = True
Set targetWorkbook = excelApp.Workbooks.open("\\shelbyville\data\Quality\1 - APQP\00 - APQP Templates\1. Auto Templates\2 External\" & GetExcelFileName("\\shelbyville\data\Quality\1 - APQP\00 - APQP Templates\1. Auto Templates\2 External", "1a"), ReadOnly:=True)
targetWorkbook.Worksheets("Info").Range("D1") = Me.partNumber
targetWorkbook.Worksheets("Info").Range("D2") = Me.toolNumber
targetWorkbook.Worksheets("Info").Range("D3") = Me.partNumber
'targetWorkbook.Worksheets("Info").Range("D4") = Me.Tooling_Engineer
'targetWorkbook.Worksheets("Info").Range("D5") = Me.Project_Engineer
'targetWorkbook.Worksheets("Info").Range("D6") = Me.Quality_Engineer
'targetWorkbook.Worksheets("Info").Range("D7") = Me.DevDesignEng
'targetWorkbook.Worksheets("Info").Range("D8") = Me.sales
targetWorkbook.Worksheets("Info").Range("D9") = Me.materialSymbol
targetWorkbook.Worksheets("Info").Range("D10") = Me.customerPN
targetWorkbook.Worksheets("Info").Range("D11") = Me.DESCRIPTION
targetWorkbook.Worksheets("Info").Range("D12") = Me.cavitation
targetWorkbook.Worksheets("Info").Range("D13") = Me.programId
targetWorkbook.Worksheets("Info").Range("D14") = Me.customerId
targetWorkbook.Worksheets("Info").Range("D15") = Me.colorCode
targetWorkbook.Worksheets("Info").Range("D16") = Me.GrainRequired
targetWorkbook.Worksheets("Info").Range("D17") = Me.materialSpec
targetWorkbook.Worksheets("Info").Range("D18") = Me.familyTool
'targetWorkbook.Worksheets("Info").Range("D19") = Me.Production_Location
'targetWorkbook.Worksheets("Info").Range("D20") = Me.assembly
targetWorkbook.Worksheets("Info").Range("D21") = Form_sfrmNMQ_PartContacts.contactName
targetWorkbook.Worksheets("Info").Range("D22") = Me.checkFixture
targetWorkbook.Worksheets("Info").Range("D23") = Me.modelCode
targetWorkbook.Worksheets("Info").Range("D24") = Me.customerRevLevel
'targetWorkbook.Worksheets("Info").Range("D25") = Me.DeltaRFlammability
'targetWorkbook.Worksheets("Info").Range("D26") = Me.DeltaRPart_Marking
'targetWorkbook.Worksheets("Info").Range("D27") = Me.DeltaRVOC
'targetWorkbook.Worksheets("Info").Range("D28") = Me.DeltaRDimension
targetWorkbook.Worksheets("Info").Range("D29") = Me.materialNumber
targetWorkbook.Worksheets("Info").Range("D30") = Me.generalTolerance
targetWorkbook.Worksheets("Info").Range("D31") = Me.NominalWeight
targetWorkbook.Worksheets("Info").Range("D32") = Me.massTolerance
targetWorkbook.Worksheets("Info").Range("D33") = Me.GrainRequired
targetWorkbook.Worksheets("Info").Range("D34") = Me.appearanceType
targetWorkbook.Worksheets("Info").Range("D35") = Me.GrainDepth
targetWorkbook.Worksheets("Info").Range("D36") = Me.testPanel
targetWorkbook.Worksheets("Info").Range("D37") = Me.annealing <> 0
targetWorkbook.Worksheets("Info").Range("D38") = Me.annealingDetails
targetWorkbook.Worksheets("Info").Range("D39") = Me.ColorantNumber
targetWorkbook.Worksheets("Info").Range("D40") = Me.colorCode
targetWorkbook.Worksheets("Info").Range("D41") = Me.GlossValue
targetWorkbook.Worksheets("Info").Range("D42") = Me.IMDSnumber
targetWorkbook.Worksheets("Info").Range("headersdate") = Date

Set targetWorkbook = Nothing
Set excelApp = Nothing

End Sub

Private Sub partDataButton_Click()
On Error GoTo Err_Handler
Dim partNum, thousZeros, FolName, mainPath, strPath

partNum = Me.partNumber
thousZeros = Left(partNum, 2) & "000\"
mainPath = mainFolder("slbProject")
strPath = mainPath & thousZeros

If Len(partNum) = 5 Then
    FolName = Dir(strPath & partNum & "*", vbDirectory)
    Call openPath(strPath & FolName & "\TOOLING DOCUMENTS\" & "Part Data")
Else
    Call openPath(mainPath)
End If
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub PPAPAppButton_Click()
On Error GoTo Err_Handler
Dim partNum, thousZeros, FolName, mainPath, strPath

partNum = Me.partNumber
thousZeros = Left(partNum, 2) & "000\"
mainPath = mainFolder("slbProject")
strPath = mainPath & thousZeros

If Len(partNum) = 5 Then
    FolName = Dir(strPath & partNum & "*", vbDirectory)
    Call openPath(strPath & FolName & "\QUALITY DOCUMENTS\" & "PA_PPAP\APPROVAL")
Else
    Call openPath(mainPath)
End If
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub PPAPbutton_Click()
On Error GoTo Err_Handler
Dim partNum, thousZeros, FolName, mainPath, strPath

partNum = Me.partNumber
thousZeros = Left(partNum, 2) & "000\"
mainPath = mainFolder("slbProject")
strPath = mainPath & thousZeros

If Len(partNum) = 5 Then
    FolName = Dir(strPath & partNum & "*", vbDirectory)
    Call openPath(strPath & FolName & "\QUALITY DOCUMENTS\" & "PA_PPAP")
Else
    Call openPath(mainPath)
End If
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub PQSIPIRBUTTON_Click()
On Error GoTo Err_Handler

toolNum = CStr(Me.partNumber)
'targetName = GetExcelFileName("\\shelbyville.nifcoam.local\data\Quality\3 - Internal Documents\PQS - IPIR new Template", toolNum)

'Method 1: Search each file name
'If Left(targetName, 5) = Left(toolNum, 5) Then
'    answer = MsgBox("APQP found, press (Yes) to open Excel file or (No) to open file path.", vbYesNoCancel)
'
'        If answer = vbYes Then
'            Application.FollowHyperlink "\\shelbyville\data\Quality\3 - Internal Documents\PQS - IPIR new Template\" & targetName
'        ElseIf answer = vbNo Then
'            Application.FollowHyperlink "\\shelbyville\data\Quality\3 - Internal Documents\PQS - IPIR new Template\"
'        Else
'            Exit Sub
'        End If
'Else
'    MsgBox ("No APQP found in folder. The naming convention may be incorrect or it may not exist!" & vbNewLine & "Press OK to open folder.")
'    Application.FollowHyperlink "\\shelbyville\data\Quality\3 - Internal Documents\PQS - IPIR new Template\"
'End If

'Method 2: just hyperlink to that bad boy and provide alternative options if it is not named correctly/doesn't exist
On Error GoTo ErrHandler1

Dim excelApp, targetWorkbook, lastNum, FamilyToolNumbers

If Me.familyTool = False Then
    Set excelApp = CreateObject("Excel.application", "")
    excelApp.Visible = True
    Set targetWorkbook = excelApp.Workbooks.open("\\shelbyville\data\Quality\3 - Internal Documents\PQS - IPIR new Template\" & Me.partNumber & "-APQP.xlsm", UpdateLinks:=False)
    'Application.FollowHyperlink "\\shelbyville\data\Quality\3 - Internal Documents\PQS - IPIR new Template\" & Me.partNumber & "-APQP.xlsm"
Else
    On Error GoTo ErrHandler2
    lastNum = Right(Me.partNumber, 1) + 1
    FamilyToolNumbers = toolNum & "(" & lastNum & ")"
    
    Set excelApp = CreateObject("Excel.application", "")
    excelApp.Visible = True
    Set targetWorkbook = excelApp.Workbooks.open("\\shelbyville\data\Quality\3 - Internal Documents\PQS - IPIR new Template\" & FamilyToolNumbers & "-APQP.xlsm", UpdateLinks:=False)
    'Application.FollowHyperlink "\\shelbyville\data\Quality\3 - Internal Documents\PQS - IPIR new Template\" & FamilyToolNumbers & "-APQP.xlsm"
End If

Set targetWorkbook = Nothing
Set excelApp = Nothing

Exit Sub
ErrHandler1:
MsgBox ("No APQP found in folder. The naming convention may be incorrect or it may not exist!" & vbNewLine & "Correct naming convention:  XXXXX-APQP or XXXXX(X)-APQP (for family tools) ")
Application.FollowHyperlink "\\shelbyville\data\Quality\3 - Internal Documents\PQS - IPIR new Template\"

Exit Sub
ErrHandler2:
FamilyToolNumbers = toolNum & "(" & Right(Me.partNumber, 2) + 1 & ")"
Application.FollowHyperlink "\\shelbyville\data\Quality\3 - Internal Documents\PQS - IPIR new Template\" & FamilyToolNumbers & "-APQP.xlsm"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub projectButton_Click()
On Error GoTo Err_Handler
Dim partNum, thousZeros, FolName, mainPath, strPath

partNum = Me.partNumber
thousZeros = Left(partNum, 2) & "000\"
mainPath = mainFolder("slbProject")
strPath = mainPath & thousZeros

If Len(partNum) = 5 Then
    FolName = Dir(strPath & partNum & "*", vbDirectory)
    Call openPath(strPath & FolName & "\PROJECT DOCUMENTS\")
Else
    Call openPath(mainPath)
End If
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub projectrequestButton_Click()
On Error GoTo Err_Handler
Dim partNum, thousZeros, FolName, mainPath, strPath

partNum = Me.partNumber
thousZeros = Left(partNum, 2) & "000\"
mainPath = mainFolder("slbProject")
strPath = mainPath & thousZeros

If Len(partNum) = 5 Then
    FolName = Dir(strPath & partNum & "*", vbDirectory)
    Call openPath(strPath & FolName & "\PROJECT DOCUMENTS\Project Request")
Else
    Call openPath(mainPath)
End If
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub qualityFolderButton_Click()
On Error GoTo Err_Handler
Dim partNum, thousZeros, FolName, mainPath, strPath

partNum = Me.partNumber
thousZeros = Left(partNum, 2) & "000\"
mainPath = mainFolder("slbProject")
strPath = mainPath & thousZeros

If Len(partNum) = 5 Then
    FolName = Dir(strPath & partNum & "*", vbDirectory)
    Call openPath(strPath & FolName & "\QUALITY DOCUMENTS\")
Else
    Call openPath(mainPath)
End If
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
'Application.FollowHyperlink nppfAddressFull & "\QUALITY DOCUMENTS\"
End Sub

Private Sub searchCerts_Click()
On Error GoTo Err_Handler

Dim ip
ip = InputBox("Enter material number", "Input here", Nz(Me.materialNumber))
If ip = "" Then Exit Sub
Application.FollowHyperlink "\\shelbyville\data\Quality\Receiving\Incoming Logs\Raw Material & Certs\resin\Cert\" & ip

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub searchColorantButton_Click()
On Error GoTo Err_Handler

Dim ip
ip = InputBox("Enter colorant number", "Input here", Nz(Me.materialNumber1))
If ip = "" Then Exit Sub

Application.FollowHyperlink "\\shelbyville\data\Quality\Receiving\Incoming Logs\Raw Material & Certs\color\Color\Cert\" & ip

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub searchMatTestDataButton_Click()
On Error GoTo Err_Handler

Dim ip
ip = InputBox("Enter material number", "Input here", Nz(Me.materialNumber))
Application.FollowHyperlink "\\shelbyville\data\Quality\24 - Material Test Data\" & ip

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub searchTools_Click()
On Error GoTo Err_Handler

partSelected = InputBox("Enter Nifco Part Number", "Part Number Entry")
If partSelected = vbNullString Or StrPtr(partSelected) = 0 Then Exit Sub


If Len(partSelected) < 5 Then
    MsgBox ("Part number invalid. Try adding more numbers ;)")
    Exit Sub
End If

If Len(partSelected) > 6 Then
    MsgBox ("Part number invalid. Try removing some numbers lol")
    Exit Sub
End If

DoCmd.applyFilter , "[partNumber]=" & "'" & [partSelected] & "'"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub slbSDS_Click()
On Error GoTo Err_Handler
Dim partNum, thousZeros, FolName, mainPath, strPath

partNum = Me.partNumber
thousZeros = Left(partNum, 2) & "000\"
mainPath = mainFolder("slbProject")
strPath = mainPath & thousZeros

If Len(partNum) = 5 Then
    FolName = Dir(strPath & partNum & "*", vbDirectory)
    Call openPath(strPath & FolName & "\QUALITY DOCUMENTS\" & "SDS FOLDER")
Else
    Call openPath(mainPath)
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub staffFolderButton_Click()
On Error GoTo Err_Handler

Dim nameOfUser
nameOfUser = Environ("username")
nameOfUser = Left(nameOfUser, Len(nameOfUser) - 1)

Dim objFS As Variant
Dim objfolder As Variant
Dim objFile As Variant
Dim intLengthOfPartialName As Integer
Dim strfilenamefull As String

Set objFS = CreateObject("Scripting.FileSystemObject")
Set objfolder = objFS.GetFolder("\\shelbyville\data\Quality\8 - Staff Folders\")

'work out how long the partial file name is
intLengthOfPartialName = Len(nameOfUser)

For Each objfolder In objfolder.SubFolders

    'Test to see if the first 5 digits of the folder name matches the input
    If Right(objfolder.name, intLengthOfPartialName) = nameOfUser Then
        strfilenamefull = objfolder.name 'get the full file name
    
Exit For
Else
End If

Next objfolder

Application.FollowHyperlink "\\shelbyville\data\Quality\8 - Staff Folders\" & strfilenamefull

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Test_Panels_Required_AfterUpdate()
On Error GoTo Err_Handler

If Me.testPanel = True Then
    'Me.testPanelsOrdered.Visible = True
    'Me.testpanelslabel.Visible = False
Else
    'Me.testPanelsOrdered.Visible = False
    'Me.testpanelslabel.Visible = True
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub testingFolderButton_Click()
On Error GoTo Err_Handler

Application.FollowHyperlink nppfAddressFull & "\QUALITY DOCUMENTS\TESTING FOLDER"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub summonMike_Click()
If summonMike = True Then
    mascotImage.Visible = True
    wisdomText.Visible = True
    summonMike.Caption = "Dismiss Mike"
    wisdomText.Caption = "..."
End If

If summonMike = False Then
    mascotImage.Visible = False
    wisdomText.Visible = False
    summonMike.Caption = "Summon Mike"
End If
End Sub

Private Sub tirFolderButton_Click()
On Error GoTo Err_Handler
Dim partNum, thousZeros, FolName, mainPath, strPath

partNum = Me.partNumber
thousZeros = Left(partNum, 2) & "000\"
mainPath = mainFolder("slbProject")
strPath = mainPath & thousZeros

If Len(partNum) = 5 Then
    FolName = Dir(strPath & partNum & "*", vbDirectory)
    Call openPath(strPath & FolName & "\TOOLING DOCUMENTS\TIR Folder")
Else
    Call openPath(mainPath)
End If
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub crossRefButton_Click()
On Error GoTo Err_Handler

Form_DASHBOARD.partNumberSearch = Me.partNumber
Form_DASHBOARD.xref_Click

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub TIRFormButton_Click()
On Error GoTo Err_Handler

Dim excelApp, targetWorkbook
Set excelApp = CreateObject("Excel.application", "")
excelApp.Visible = True
Set targetWorkbook = excelApp.Workbooks.open("\\shelbyville\data\Quality\1 - APQP\00 - APQP Templates\1. Auto Templates\1 Internal\" & GetExcelFileName("\\shelbyville\data\Quality\1 - APQP\00 - APQP Templates\1. Auto Templates\1 Internal", "2a"), ReadOnly:=True)
targetWorkbook.Worksheets("Cover Page").Range("model") = Me.modelCode
targetWorkbook.Worksheets("Cover Page").Range("toolNum") = Me.toolNumber
targetWorkbook.Worksheets("Cover Page").Range("dateToday") = Date
targetWorkbook.Worksheets("Cover Page").Range("custnum") = Me.customerPN
targetWorkbook.Worksheets("Cover Page").Range("namNo") = Me.partNumber
targetWorkbook.Worksheets("Cover Page").Range("partName") = Me.DESCRIPTION
targetWorkbook.Worksheets("Cover Page").Range("tirNum") = Me.partNumber & "-" & Format(Date, Format:="yymmdd")
'targetWorkbook.Worksheets("Cover Page").Range("qe") = Me.Quality_Engineer 'need to access recordset
'targetWorkbook.Worksheets("Cover Page").Range("pe") = Me.Project_Engineer
'targetWorkbook.Worksheets("Cover Page").Range("de") = Me.DevDesignEng
'targetWorkbook.Worksheets("Cover Page").Range("te") = Me.Tooling_Engineer

Set targetWorkbook = Nothing
Set excelApp = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub toolingFolderButton_Click()
On Error GoTo Err_Handler
Dim partNum, thousZeros, FolName, mainPath, strPath

partNum = Me.partNumber
thousZeros = Left(partNum, 2) & "000\"
mainPath = mainFolder("slbProject")
strPath = mainPath & thousZeros

If Len(partNum) = 5 Then
    FolName = Dir(strPath & partNum & "*", vbDirectory)
    Call openPath(strPath & FolName & "\TOOLING DOCUMENTS")
Else
    Call openPath(mainPath)
End If
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub toyotaISButton_Click()
On Error GoTo Err_Handler

Dim excelApp, targetWorkbook
Set excelApp = CreateObject("Excel.application", "")
excelApp.Visible = True
Set targetWorkbook = excelApp.Workbooks.open("\\shelbyville\data\Quality\1 - APQP\00 - APQP Templates\1. Auto Templates\2 External\" & GetExcelFileName("\\shelbyville\data\Quality\1 - APQP\00 - APQP Templates\1. Auto Templates\2 External", "4a"), ReadOnly:=True)
'Info Page Info
targetWorkbook.Worksheets("Info").Range("sigdate") = Date
targetWorkbook.Worksheets("Info").Range("nifco_number") = Me.partNumber
targetWorkbook.Worksheets("Info").Range("PART_NUMBER") = Me.customerPN
targetWorkbook.Worksheets("Info").Range("description") = Me.DESCRIPTION
targetWorkbook.Worksheets("Info").Range("nifco_number") = Me.partNumber
targetWorkbook.Worksheets("Info").Range("model") = Me.modelCode
targetWorkbook.Worksheets("Info").Range("program") = Me.programId
targetWorkbook.Worksheets("Info").Range("nifco_number") = Me.partNumber
targetWorkbook.Worksheets("Info").Range("NAMC") = Me.customerId
targetWorkbook.Worksheets("Info").Range("eci") = Me.customerRevLevel
targetWorkbook.Worksheets("Info").Range("custqe") = Me.customerId
targetWorkbook.Worksheets("Info").Range("nifco_number") = Me.partNumber
'targetWorkbook.Worksheets("Info").Range("qe") = Me.Quality_Engineer
targetWorkbook.Worksheets("Info").Range("date") = Date
targetWorkbook.Worksheets("Info").Range("cfRequired") = Me.checkFixture
'targetWorkbook.Worksheets("Info").Range("C16") = Me.DeltaRFlammability
'targetWorkbook.Worksheets("Info").Range("C17") = Me.DeltaRPart_Marking
'targetWorkbook.Worksheets("Info").Range("C18") = Me.DeltaRVOC
'targetWorkbook.Worksheets("Info").Range("C19") = Me.DeltaRDimension
targetWorkbook.Worksheets("Info").Range("material") = Me.materialNumber
targetWorkbook.Worksheets("Info").Range("gentol") = Me.generalTolerance
targetWorkbook.Worksheets("Info").Range("weight") = Me.NominalWeight
targetWorkbook.Worksheets("Info").Range("masstolselect") = Me.massTolerance
'targetWorkbook.Worksheets("Info").Range("C16") = Me.DeltaRFlammability
targetWorkbook.Worksheets("Info").Range("mat_name") = Me.materialSymbol
targetWorkbook.Worksheets("Info").Range("TSM") = Me.materialSpec
'BCD Page Info
targetWorkbook.Worksheets("BCD").Range("grainreq") = Me.GrainRequired
'targetWorkbook.Worksheets("BCD").Range("AS9") = Me.GrainType
'targetWorkbook.Worksheets("BCD").Range("AS10") = Me.GrainDepth
'targetWorkbook.Worksheets("BCD").Range("AS10") = Me.GrainDepth
'EFG Page Info
'targetWorkbook.Worksheets("EFG").Range("AR18") = Me.ColorAge
'targetWorkbook.Worksheets("EFG").Range("AR19") = Me.VOC
'targetWorkbook.Worksheets("EFG").Range("AR20") = Me.GlassFog
'targetWorkbook.Worksheets("EFG").Range("AR21") = Me.SmellTest

Set targetWorkbook = Nothing
Set excelApp = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub toyotaSDSButton_Click()
On Error GoTo Err_Handler

Dim excelApp, targetWorkbook
Set excelApp = CreateObject("Excel.application", "")
excelApp.Visible = True
Set targetWorkbook = excelApp.Workbooks.open("\\shelbyville\data\Quality\1 - APQP\00 - APQP Templates\1. Auto Templates\2 External\" & GetExcelFileName("\\shelbyville\data\Quality\1 - APQP\00 - APQP Templates\1. Auto Templates\2 External", "3a"), ReadOnly:=True)
targetWorkbook.Worksheets("Info").Range("RHPN") = Me.customerPN
targetWorkbook.Worksheets("Info").Range("LHPN") = Me.customerPN
targetWorkbook.Worksheets("Info").Range("custnum") = Me.customerPN
targetWorkbook.Worksheets("Info").Range("RHNAME") = Me.DESCRIPTION
targetWorkbook.Worksheets("Info").Range("LHNAME") = Me.DESCRIPTION
targetWorkbook.Worksheets("Info").Range("partName") = Me.DESCRIPTION
targetWorkbook.Worksheets("Info").Range("model") = Me.modelCode
targetWorkbook.Worksheets("Info").Range("eci") = Me.customerRevLevel
targetWorkbook.Worksheets("Info").Range("custQE") = Me.customerId
'targetWorkbook.Worksheets("Info").Range("nifcoqe") = Me.Quality_Engineer
targetWorkbook.Worksheets("Info").Range("RHNAME") = Me.DESCRIPTION
targetWorkbook.Worksheets("Info").Range("date") = Date
targetWorkbook.Worksheets("Info").Range("sigdate") = Date

Set targetWorkbook = Nothing
Set excelApp = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub gageInfoSheet_Click()
On Error GoTo Err_Handler

Dim excelApp, targetWorkbook
Set excelApp = CreateObject("Excel.application", "")
excelApp.Visible = True
Set targetWorkbook = excelApp.Workbooks.open("\\shelbyville\data\Quality\1 - APQP\00 - APQP Templates\1. Auto Templates\1 Internal\" & GetExcelFileName("\\shelbyville\data\Quality\1 - APQP\00 - APQP Templates\1. Auto Templates\1 Internal", "10a"), ReadOnly:=True)
targetWorkbook.Worksheets("Gauge Info Sheet").Range("QE") = getFullName()
targetWorkbook.Worksheets("Gauge Info Sheet").Range("subDate") = Date
targetWorkbook.Worksheets("Gauge Info Sheet").Range("nifcoNum") = Me.partNumber

Set targetWorkbook = Nothing
Set excelApp = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub japanCFRequest_Click()
On Error GoTo Err_Handler

Dim excelApp, targetWorkbook
Set excelApp = CreateObject("Excel.application", "")
excelApp.Visible = True
Set targetWorkbook = excelApp.Workbooks.open("\\shelbyville\data\Quality\1 - APQP\00 - APQP Templates\1. Auto Templates\1 Internal\" & GetExcelFileName("\\shelbyville\data\Quality\1 - APQP\00 - APQP Templates\1. Auto Templates\1 Internal", "11a"), ReadOnly:=True)

targetWorkbook.Worksheets("1").Range("todaysDate") = Date
targetWorkbook.Worksheets("1").Range("qe") = getFullName()
targetWorkbook.Worksheets("1").Range("namNo") = Me.partNumber
targetWorkbook.Worksheets("1").Range("custNum") = Me.customerPN
targetWorkbook.Worksheets("1").Range("partName") = Me.DESCRIPTION
targetWorkbook.Worksheets("1").Range("projectCode") = Me.programId.Value
targetWorkbook.Worksheets("1").Range("eciNum") = Me.customerRevLevel

Set targetWorkbook = Nothing
Set excelApp = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub pcrPacket_Click()
On Error GoTo Err_Handler

Dim excelApp, targetWorkbook
Set excelApp = CreateObject("Excel.application", "")
excelApp.Visible = True
Set targetWorkbook = excelApp.Workbooks.open("\\shelbyville\data\Quality\1 - APQP\00 - APQP Templates\1. Auto Templates\2 External\" & GetExcelFileName("\\shelbyville\data\Quality\1 - APQP\00 - APQP Templates\1. Auto Templates\2 External", "5a"), ReadOnly:=True)
'targetWorkbook.Worksheets("Info").Range("RHPN") = Me.customerPN

Set targetWorkbook = Nothing
Set excelApp = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
Private Sub tprButton_Click()
On Error GoTo Err_Handler
Dim partNum, thousZeros, FolName, mainPath, strPath

partNum = Me.partNumber
thousZeros = Left(partNum, 2) & "000\"
mainPath = mainFolder("slbProject")
strPath = mainPath & thousZeros

If Len(partNum) = 5 Then
    FolName = Dir(strPath & partNum & "*", vbDirectory)
    Call openPath(strPath & FolName & "\PROJECT DOCUMENTS\Schedules_progress reports")
Else
    Call openPath(mainPath)
End If
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub trialReviewsButton_Click()
On Error GoTo Err_Handler

toolNum = Me.partNumber.Value
toolNumFirstTwo = Left(toolNum, 2) & "000"
Application.FollowHyperlink "\\shelbyville\data\Manufacturing\Master Set-ups\" & toolNumFirstTwo & "\" & toolNum & "\" & "Trial Reports"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub WOFormButton_Click()
On Error Resume Next
'Dim dbs As Database
'Set dbs = CurrentDb
'Set rsQuery = dbs.OpenRecordset("WorkOrderQ")

Dim excelApp, targetWorkbook, answer, answer1, answer2
Set excelApp = CreateObject("Excel.application", "")

answer = MsgBox("Internal Work Order format?", vbYesNoCancel)
        If answer = vbYes Then
            Set targetWorkbook = excelApp.Workbooks.open("\\shelbyville\data\Quality\1 - APQP\00 - APQP Templates\1. Auto Templates\1 Internal\" & GetExcelFileName("\\shelbyville\data\Quality\1 - APQP\00 - APQP Templates\1. Auto Templates\1 Internal", "4 AUTO"), ReadOnly:=True)
        ElseIf answer = vbNo Then
            answer1 = MsgBox("Toyota format?", vbYesNoCancel)
            If answer1 = vbYes Then
                Set targetWorkbook = excelApp.Workbooks.open("\\shelbyville\data\Quality\1 - APQP\00 - APQP Templates\1. Auto Templates\1 Internal\" & GetExcelFileName("\\shelbyville\data\Quality\1 - APQP\00 - APQP Templates\1. Auto Templates\1 Internal", "3 AUTO"), ReadOnly:=True)
            End If
            If answer1 = vbNo Then
                answer2 = MsgBox("TBA format?", vbYesNoCancel)
                If answer2 = vbYes Then
                Set targetWorkbook = excelApp.Workbooks.open("\\shelbyville\data\Quality\1 - APQP\00 - APQP Templates\1. Auto Templates\1 Internal\" & GetExcelFileName("\\shelbyville\data\Quality\1 - APQP\00 - APQP Templates\1. Auto Templates\1 Internal", "5 AUTO"), ReadOnly:=True)
                End If
            End If
        End If
excelApp.Visible = True

'Set targetWorkbook = excelApp.workbooks.Open("\\shelbyville\data\Quality\1 - APQP\00 - APQP Templates\1. Auto Templates\1 Internal\" & GetExcelFileName("\\shelbyville\data\Quality\1 - APQP\00 - APQP Templates\1. Auto Templates\1 Internal", "3 AUTO"))
'targetWorkbook.Worksheets("data").Range("A1").CopyFromRecordset rsQuery
targetWorkbook.Worksheets("Lab Work Order").Range("J13") = Me.toolNumber
targetWorkbook.Worksheets("Lab Work Order").Range("J11") = Me.customerId.Value
targetWorkbook.Worksheets("Lab Work Order").Range("J15") = Me.customerPN
targetWorkbook.Worksheets("Lab Work Order").Range("J17") = Me.partNumber
targetWorkbook.Worksheets("Lab Work Order").Range("J19") = Me.DESCRIPTION
targetWorkbook.Worksheets("Lab Work Order").Range("J21") = getFullName()
targetWorkbook.Worksheets("Lab Work Order").Range("AB21") = Date

If Me.annealing <> 0 Then
    targetWorkbook.Worksheets("Lab Work Order").Range("K23") = "X"
    targetWorkbook.Worksheets("Lab Work Order").Range("AB23") = Me.annealingDetails
Else
    targetWorkbook.Worksheets("Lab Work Order").Range("P23") = "X"
End If

'targetWorkbook.Worksheets("Lab Work Order").Range("AB15") = Me.Production_Location
targetWorkbook.Worksheets("Lab Work Order").Range("J25") = Me.customerRevLevel
targetWorkbook.Worksheets("Lab Work Order").Range("J27") = Me.customerRevLevel
'targetWorkbook.Worksheets("Lab Work Order").Range("G29") = Me.Project_Engineer
'targetWorkbook.Worksheets("Lab Work Order").Range("Q29") = Me.Tooling_Engineer
'targetWorkbook.Worksheets("Lab Work Order").Range("AB29") = Me.DevDesignEng
targetWorkbook.Worksheets("Lab Work Order").Range("V25") = Me.programId
targetWorkbook.Worksheets("Lab Work Order").Range("AB9") = Date
targetWorkbook.Worksheets("Lab Work Order").Range("AB17") = Me.materialNumber
targetWorkbook.Worksheets("Lab Work Order").Range("AB19") = Me.cavitation
targetWorkbook.Worksheets("Lab Work Order").Range("C38") = "X"

If Me.checkFixture = True Then
    targetWorkbook.Worksheets("Lab Work Order").Range("M45") = "X"
End If

'If Me.forceTestingRequired = True Then
'    targetWorkbook.Worksheets("Lab Work Order").Range("C52") = "X"
'    targetWorkbook.Worksheets("Lab Work Order").Range("G59") = "3"
'    targetWorkbook.Worksheets("Lab Work Order").Range("M53") = "X"
'    targetWorkbook.Worksheets("Lab Work Order").Range("AA59") = "STD"
'End If

If Me.testPanel = True Then
    targetWorkbook.Worksheets("Lab Work Order").Range("T55") = "X"
End If

Set targetWorkbook = Nothing
Set excelApp = Nothing

End Sub

Function openMainPath()
On Error GoTo Err_Handler

Call openPath(mainFolder(Me.ActiveControl.name))

Exit Function
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Function


'Sub listFiles(mySourcePath)
''-Runs through each folder in a given source path and compares folder names to the values in nppfFoldersT to see if any are not listed.
''-If not listed, this procedure will list the missing file.
''-This assumes for now that we will only be worried about listing the new files for the 30000 NPPF.
''-Currently, it will not work correctly on previous part folders (28000, 27000), but can be adjusted to do so.
'
'Dim Counter As Integer
'Set myObject = CreateObject("Scripting.FileSystemObject")
'Set mySource = myObject.GetFolder(mySourcePath)
'
''on error Resume Next
'
'Dim db As Database
'Dim rs As Recordset
'Dim mysql As String
'
'Set db = CurrentDb
'mysql = "nppfFoldersT"
'Set rs = db.OpenRecordset(mysql, dbOpenDynaset)
'
'newCounter = 0
'
'For Each myFile In mySource.Subfolders
'
'    rs.MoveLast
'    safetyCounter = 0
'    Do Until safetyCounter = 500
'        safetyCounter = safetyCounter + 1
'
'        If rs![nppfName] = myFile.name Then
'            Exit Do
'        End If
'        rs.MovePrevious
'    Loop
'
'    If rs![nppfName] = myFile.name Then
'        skippedCounter = skippedCounter + 1
'        'MsgBox ("Skipped the file: " & myFile.Name & vbNewLine & "Reason - Table entry already exists: " & rs![nppfName])
'    End If
'
'    If Not rs![nppfName] = myFile.name Then
'        rs.addNew
'        rs![nppfPath] = myFile.Path
'        rs![nppfName] = myFile.name
'        rs.Update
'        newCounter = newCounter + 1
'        'MsgBox ("Added the file: " & myFile.Name)
'    End If
'
'Next
'
'' Clean up
'rs.Close
'db.Close
'Set rs = Nothing
'Set db = Nothing
'
'Set myObject = Nothing
'Set mySource = Nothing
'
'MsgBox "Listed " & newCounter & " new files and skipped " & skippedCounter & " files that already existed."
'End Sub

Sub getPartTeam()
On Error GoTo Err_Handler

'--Jacob -- FYI there is a global function for this : grabPartTeam

Err_Handler:
End Sub
