Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Function setApp(appName As String, shortName As String, Optional filterString As String = "")
On Error GoTo Err_Handler

Me.partNumberSearch.SetFocus

If TempVars!smallScreen = "False" Then
    Me.appContainer.SourceObject = appName
    Me.appContainer.Visible = True
    
    Call resetAppButtons
    
    Me.Controls("ln" & shortName).Visible = True
    Me.Controls("cs" & shortName).Visible = True
    
    Me.Controls("app" & shortName).tag = "btn.L1"
    Me.Controls("app" & shortName).FontWeight = 700
    Me.Controls("app" & shortName).Picture = Replace(Me.Controls("app" & shortName).Picture, "_inactive", "")
    
    Call setTheme(Me)
    
    If filterString <> "" Then
        Me.appContainer.Form.filter = filterString
        Me.appContainer.Form.FilterOn = True
    End If
    
    Me.appContainer.SetFocus
Else
    DoCmd.OpenForm appName
    
    Call resetAppButtons
    
    Me.Controls("ln" & shortName).Visible = True
    Me.Controls("cs" & shortName).Visible = True
    
    Me.Controls("app" & shortName).tag = "btn.L1"
    Me.Controls("app" & shortName).FontWeight = 700
    Me.Controls("app" & shortName).Picture = Replace(Me.Controls("app" & shortName).Picture, "_inactive", "")
    
    Call setTheme(Me)
    
    If filterString <> "" Then
        forms(appName).Form.filter = filterString
        forms(appName).Form.FilterOn = True
    End If
    
End If

Exit Function
Err_Handler:
    Call handleError(Me.name, "setApp", Err.DESCRIPTION, Err.number)
End Function

Private Sub appAutomation_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)
Call setApp("frmPartAutomationGates", "Automation", "(partNumber IN (SELECT partNumber FROM tblPartTeam WHERE person = '" & Environ("username") & "'))")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub appCapacity_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)

If userData("Dept") = "Sales" Or userData("Developer") Or userData("Strategic Planning") Then
    Call setApp("frmLab_WO_tracker", "Lab", "lab_wo_status <> 'Closed'")
Else
    Call snackBox("error", "You don't have access.", "You must a beta tester to open this app", Me.name)
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub appCPC_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)

Call setApp("frmCPC_WorkTracker", "CPC")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub appDesignWOs_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)
If userData("Dept") = "Design" Or userData("Developer") Then
    Call setApp("frmDRSworkTracker", "DesignWOs", "[Assignee] = '" & Environ("username") & "' AND [Approval_Status] = 2 AND [Completed_Date] IS NULL")
Else
    Call snackBox("error", "You don't have access.", "You must be in Design to open this App", Me.name)
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub appLab_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)

If userData("Beta") Then
    Call setApp("frmLab_WO_tracker", "Lab", "lab_wo_status <> 'Closed'")
Else
    Call snackBox("error", "You don't have access.", "You must a beta tester to open this app", Me.name)
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub appNewParts_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)
If userData("Dept") = "Sales" Or userData("Developer") Then
    Call setApp("frmNewPartNumbers", "NewParts", "createdDate > Date() - 60")
Else
    Call snackBox("error", "You don't have access.", "You must be in Sales to open this App", Me.name)
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub appOpenIssues_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)

Call setApp("frmPartIssues", "OpenIssues", "inCharge = '" & Environ("username") & "' AND [closeDate] is null")

Form_frmPartIssues.fltInCharge = Environ("username")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub appOpenSteps_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)

Call setApp("frmPartStepTracker", "OpenSteps", "person = '" & Environ("username") & "'")

Form_frmPartStepTracker.fltUser = Environ("username")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub appPackaging_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)

Call setApp("frmPackagingTracker", "Packaging", "partNumber IN (SELECT PN FROM qryPackagingPartNumbers) AND (partNumber IN (SELECT partNumber FROM tblPartTeam WHERE person = '" & Environ("username") & "'))")
Form_frmPackagingTracker.fltUser = Environ("username")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub appPartTracking_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)

Call setApp("frmPartTracker", "PartTracking", "partProjectStatus <> 'On Hold' AND (partNumber IN (SELECT partNumber FROM tblPartTeam WHERE person = '" & Environ("username") & "'))")

Form_frmPartTracker.fltUser = Environ("username")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub appPrograms_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)

Call setApp("frmPrograms", "Programs")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub appTesting_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)

Call setApp("frmPartTestingTracker", "Testing", "partNumber IN (SELECT partNumber FROM tblPartTeam WHERE person = '" & Environ("username") & "')")

Form_frmPartTestingTracker.fltUser = Environ("username")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub appTrials_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)

Call setApp("frmPartTrialTracker", "Trials", "trialStatus <> 3 AND [trialStatus] <> 4 AND (partNumber IN (SELECT partNumber FROM tblPartTeam WHERE person = '" & Environ("username") & "') OR creator = '" & Environ("username") & "')")

Form_frmPartTrialTracker.fltUser = Environ("username")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub bomSearch_Click()
On Error GoTo Err_Handler

Dim filterVal, pNum
filterVal = idNAM(Nz(Me.partNumberSearch), "NAM")
If Nz(filterVal) = "" Then filterVal = "3147035"
filterVal = "[ASSEMBLY_ITEM_ID] = " & filterVal

DoCmd.OpenForm "frmBOMsearch", , , filterVal
Form_frmBOMsearch.NAMsrchBox = Nz(Me.partNumberSearch)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub btn3DexSheet_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)
DoCmd.OpenForm "frmPLM"
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub btnEditPicture_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmPartPicture"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub btnLearnMore_Click(): On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)
Call openPath(mainFolder(Me.ActiveControl.name))
Exit Sub
Err_Handler:: Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number): End Sub

Private Sub btnReports_Click()
On Error GoTo Err_Handler
DoCmd.OpenForm "frmReports"
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub catiaMacros_Click()
On Error GoTo Err_Handler
DoCmd.OpenForm "frmCatiaMacros"
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub checkXFteam_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name, Me.partNumberSearch)
DoCmd.OpenForm "frmXFteam"
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Function resetAppButtons()
On Error GoTo Err_Handler

Dim AppList As Collection
Set AppList = New Collection

AppList.Add "PartTracking", "PartTracking"
AppList.Add "OpenSteps", "OpenSteps"
AppList.Add "OpenIssues", "OpenIssues"
AppList.Add "Testing", "Testing"
AppList.Add "Trials", "Trials"
AppList.Add "Programs", "Programs"
AppList.Add "DesignWOs", "DesignWOs"
AppList.Add "CPC", "CPC"
AppList.Add "NewParts", "NewParts"
AppList.Add "Automation", "Automation"
AppList.Add "Packaging", "Packaging"
AppList.Add "Lab", "Lab"

Dim element
For Each element In AppList
    Me.Controls("ln" & element).Visible = False
    Me.Controls("cs" & element).Visible = False
    Me.Controls("app" & element).FontWeight = 400
    Me.Controls("app" & element).tag = "btn.L0"
    If InStr(Me.Controls("app" & element).Picture, "_inactive") = 0 Then
        Me.Controls("app" & element).Picture = Replace(Me.Controls("app" & element).Picture, ".ico", "_inactive.ico")
    End If
Next element

Exit Function
Err_Handler:
    Call handleError(Me.name, "closeApp", Err.DESCRIPTION, Err.number)
End Function

Function closeApp()
On Error GoTo Err_Handler

Me.partNumberSearch.SetFocus
Me.appContainer.Visible = False

Call resetAppButtons
Call setTheme(Me)

Exit Function
Err_Handler:
    Call handleError(Me.name, "closeApp", Err.DESCRIPTION, Err.number)
End Function

Private Sub CLOSE_Click()
DoCmd.CLOSE acForm, "DASHBOARD"
End Sub

Private Sub cnlToolData_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name, Me.partNumberSearch)
Dim partNum, thousZeros, hundZeros, FolName, mainfolderpath, strFilePath, prtpath

partNum = Me.partNumberSearch

If Len(partNum) = 5 Then
    If Left(partNum, 2) = "00" Then GoTo Singles
    If Left(partNum, 1) = "0" Then
        thousZeros = Right(Left(partNum, 2) & "000\", 5)
        hundZeros = Right(Left(partNum, 3) & "00\", 5)
        GoTo ifSet
    End If
    thousZeros = Left(partNum, 2) & "000\"
    hundZeros = Left(partNum, 3) & "00\"
    GoTo ifSet
ElseIf Len(partNum) = 4 Then
    thousZeros = Left(partNum, 1) & "000\"
    hundZeros = Left(partNum, 2) & "00\"
ElseIf Len(partNum) = 3 Then
Singles:
    thousZeros = "0-999\"
    hundZeros = ""
End If

ifSet:
mainfolderpath = mainFolder(Me.ActiveControl.name)
prtpath = mainfolderpath & thousZeros & hundZeros
FolName = Dir(prtpath & partNum & "*", vbDirectory)
strFilePath = prtpath & FolName

If Len(Me.partNumberSearch) > 0 Or Me.partNumberSearch Like "*P" Then
    If Len(FolName) = 0 Then
        If MsgBox("This folder does not exist. Do you want to go to the main folder?", vbYesNo, "Error") = vbYes Then Call openPath(mainfolderpath)
        Exit Sub
    Else
        Call openPath(strFilePath)
    End If
Else
    Call openPath(mainfolderpath)
End If
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub docHisSearch_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name, Me.partNumberSearch)
Call openDocumentHistoryFolder(Me.partNumberSearch)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub fakeHistory_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmSearchHistory"

Exit Sub
Err_Handler:
Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub feedback_Click()
On Error GoTo Err_Handler

Call logClick(Me.ActiveControl.name, Me.name)
Call openPath(mainFolder(Me.ActiveControl.name))

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Function dNumberSearch(partNum As String)
On Error GoTo Err_Handler

Dim FolName2, FolName1 As String
FolName2 = Dir(mainFolder("DocHisD") & partNum & "*", vbDirectory)
If Len(FolName2) > 1 Then Me.Controls("docHisSearch").BorderStyle = 1
FolName1 = Dir(mainFolder("ModelV5D") & partNum & "*", vbDirectory)
If Len(FolName1) > 1 Then Me.Controls("modelV5Search").BorderStyle = 1

Exit Function
Err_Handler:
    Call handleError(Me.name, "dNumberSearch", Err.DESCRIPTION, Err.number)
End Function

Function showOracleTags(boolShow As Boolean)
On Error GoTo Err_Handler

Me.NAM.Visible = boolShow
Me.MaxOfMaxOfNEW_ITEM_REVISION.Visible = boolShow
Me.INVENTORY_ITEM_STATUS_CODE.Visible = boolShow
Me.ITEM_TYPE.Visible = boolShow
Me.DESCRIPTION.Visible = boolShow
Me.lblOracle0.Visible = boolShow
Me.lblOracle1.Visible = boolShow
Me.lblOracle2.Visible = boolShow
Me.lblOracle3.Visible = boolShow
Me.lblOracle4.Visible = boolShow

Exit Function
Err_Handler:
    Call handleError(Me.name, "showOracleTags", Err.DESCRIPTION, Err.number)
End Function

Function setDrawingCaption(Title As String, Button As String)
On Error GoTo Err_Handler

If Button = "internal" Then
    Me.openInternalDwg.Caption = Title
    Me.openInternalDwg.Visible = True
Else
    Me.openCust.Caption = Title
    Me.openCust.Visible = True
End If

Me.Label157.Visible = True

Exit Function
Err_Handler:
    Call handleError(Me.name, "setDrawingCaption", Err.DESCRIPTION, Err.number)
End Function

Function setErrorText(Title As String, Optional Show As Boolean = True)
On Error GoTo Err_Handler

If Show = False Then
    Me.lblErrors.Visible = False
    Exit Function
End If

Me.lblErrors.Caption = Title
Me.lblErrors.Visible = True

Exit Function
Err_Handler:
    Call handleError(Me.name, "setErrorText", Err.DESCRIPTION, Err.number)
End Function

Function grabQuoteNum(partNum As String) As String
On Error GoTo Err_Handler

Dim db As Database
Set db = CurrentDb()
Dim rsQuoteNum As Recordset
Set rsQuoteNum = db.OpenRecordset("SELECT tblPartQuoteInfo.quoteNumber as QuoteNum FROM tblPartInfo INNER JOIN tblPartQuoteInfo ON tblPartInfo.quoteInfoId = tblPartQuoteInfo.recordId WHERE tblPartInfo.partNumber='" & partNum & "'")

If rsQuoteNum.RecordCount > 0 Then
    grabQuoteNum = Format(rsQuoteNum!quoteNum, "00000")
Else
    grabQuoteNum = ""
End If

rsQuoteNum.CLOSE
Set rsQuoteNum = Nothing
Set db = Nothing

Exit Function
Err_Handler:
    Call handleError(Me.name, "grabQuoteNum", Err.DESCRIPTION, Err.number)
End Function

Public Sub filterbyPN_Click()
On Error GoTo Err_Handler

Dim errorTracker As String
errorTracker = ""

'First, wipe everything!
DoCmd.Hourglass True
Me.Painting = False
DoCmd.Echo False
Call message_Click
Me.hidecmd.Visible = True
Me.partNumberSearch.SetFocus
Me.viewLinks.Visible = False
Me.shrtCutBadge.Visible = False
Me.lblErrors.Visible = False
Me.openInternalDwg.Visible = False
Me.openCust.Visible = False
Me.Label157.Visible = False
Call showOracleTags(False)
Me.docHisSearch.Width = 2.125 * 1440
Me.docHisSearch.BorderStyle = 0
Me.openProgram.Visible = False

Dim db As Database
Set db = CurrentDb()

errorTracker = "db"

'check if you typed something in
Dim partNum As String
If Nz(Me.partNumberSearch) = "" Then
    Call setErrorText("Please type something in")
    GoTo exitFunc
End If

errorTracker = "partNum"
partNum = Me.partNumberSearch
errorTracker = "partNum0"
Call logClick("filterbyPN", Me.name, partNum)
errorTracker = "partNum1"
If DCount("ID", "tblSessionVariables", "searchHistory = '" & StrQuoteReplace(partNum) & "'") <> 0 Then db.Execute "DELETE FROM tblSessionVariables WHERE searchHistory= '" & partNum & "'"
errorTracker = "partNum2"

On Error Resume Next
db.Execute "Insert into tblSessionVariables (searchHistory) values ('" & partNum & "');" 'UPDATEABLE QUERY ERROR, not sure why
On Error GoTo Err_Handler

TempVars.Add "partNumber", partNum
errorTracker = "tempVar"

If partNum Like "D*" Then
    Call setErrorText("Cannot Search Oracle for D#")
    Call dNumberSearch(CStr(partNum))
    GoTo exitFunc
ElseIf (Len(partNum) < 4) Then
    If partNum Like "B*" Then GoTo searchBnum
    Call setErrorText("Must be >4 digits for Oracle. No Result")
    GoTo skipOracle
End If

If TempVars!oracleConnectionFailed = "True" Then
    Call setErrorText("No connection to Oracle")
    GoTo skipOracle
End If

errorTracker = "notSpecial"

searchBnum:
Dim rs1 As DAO.Recordset
On Error GoTo checkForOracle
Set rs1 = db.OpenRecordset("SELECT INVENTORY_ITEM_ID, SEGMENT1, INVENTORY_ITEM_STATUS_CODE, ITEM_TYPE, DESCRIPTION FROM APPS_MTL_SYSTEM_ITEMS WHERE SEGMENT1 = '" & partNum & "'", dbOpenSnapshot)
errorTracker = "oracle1"
If rs1.RecordCount = 0 Then 'not in main Oracle table, now look through SIFs
    If DCount("[ROW_ID]", "APPS_Q_SIF_NEW_ASSEMBLED_PART_V", "[NIFCO_PART_NUMBER] = '" & partNum & "'") > 0 Then 'is it in assy table?
        errorTracker = "oracle2"
        Dim rsAssy As Recordset
        Set rsAssy = db.OpenRecordset("SELECT SIFNUM, PART_DESCRIPTION FROM APPS_Q_SIF_NEW_ASSEMBLED_PART_V WHERE NIFCO_PART_NUMBER = '" & partNum & "'", dbOpenSnapshot)
        rsAssy.MoveLast
        Me.NAM = partNum
        Me.MaxOfMaxOfNEW_ITEM_REVISION = "00"
        Me.INVENTORY_ITEM_STATUS_CODE = rsAssy![sifNum]
        Me.ITEM_TYPE = "SIF ASSY"
        Me.DESCRIPTION = rsAssy!PART_DESCRIPTION
        Call showOracleTags(True)
    ElseIf DCount("[ROW_ID]", "APPS_Q_SIF_NEW_MOLDED_PART_V ", "[NIFCO_PART_NUMBER] = '" & partNum & "'") > 0 Then 'is it in molded table?
        errorTracker = "oracle3"
        Dim rsMold As Recordset
        Set rsMold = db.OpenRecordset("SELECT SIFNUM, PART_DESCRIPTION FROM APPS_Q_SIF_NEW_MOLDED_PART_V WHERE NIFCO_PART_NUMBER = '" & partNum & "'", dbOpenSnapshot)
        rsMold.MoveLast
        Me.NAM = partNum
        Me.MaxOfMaxOfNEW_ITEM_REVISION = "00"
        Me.INVENTORY_ITEM_STATUS_CODE = rsMold![sifNum]
        Me.ITEM_TYPE = "SIF MLD"
        Me.DESCRIPTION = rsMold!PART_DESCRIPTION
        Call showOracleTags(True)
    Else
        errorTracker = "oracle4"
        Call setErrorText("Part not found in Oracle")
        Call showOracleTags(False)
    End If
    GoTo skipOracle
End If
Dim rs2 As DAO.Recordset
Set rs2 = db.OpenRecordset("SELECT MAX(NEW_ITEM_REVISION) AS REV FROM ENG_ENG_REVISED_ITEMS WHERE REVISED_ITEM_ID = " & rs1("INVENTORY_ITEM_ID") & " AND IMPLEMENTATION_DATE IS NOT NULL", dbOpenSnapshot)
errorTracker = "oracle5"
If rs2.RecordCount = 0 Then GoTo skipOracle
If rs2("REV") = "" Or IsNull(rs2("REV")) Then
    Me.MaxOfMaxOfNEW_ITEM_REVISION = "00"
Else
    Me.MaxOfMaxOfNEW_ITEM_REVISION = rs2("REV")
End If
errorTracker = "oracle6"
Me.DESCRIPTION = rs1("DESCRIPTION")
Me.INVENTORY_ITEM_STATUS_CODE = rs1("INVENTORY_ITEM_STATUS_CODE")
Me.ITEM_TYPE = rs1("ITEM_TYPE")
Me.NAM = rs1("SEGMENT1")

rs1.CLOSE
rs2.CLOSE
Set rs1 = Nothing
Set rs2 = Nothing

errorTracker = "afterOracle"

Call showOracleTags(True)

skipOracle:
On Error GoTo Err_Handler
Dim thousZeros As String, hundZeros As String, i As Integer, mainPath As String, partNumLeft As String, partNumRight As String, NCMpart As Boolean, alistDim As Boolean
Dim direct As String
Dim alist() As String
Dim btnPaths As Collection, designDept As Boolean, bordCol As Long, bordStyle As Long
Set btnPaths = New Collection

thousZeros = Left(partNum, 2) & "000\"
hundZeros = Left(partNum, 3) & "00\"
partNumLeft = Left(partNum, 4)
partNumRight = Right(partNum, 1)
i = 0

'---check if NCM part number. Folders are different for these---
If partNum Like "[A-Z][A-Z]##[A-Z]##[A-Z]" Or partNum Like "[A-Z][A-Z]##[A-Z]##" Or partNum Like "##[A-Z]##" Then
    NCMpart = True
    If Not partNum Like "##[A-Z]##" Then partNum = Mid(partNum, 3, 5)
    hundZeros = Left(partNum, 3) & "00\"
    mainPath = mainFolder("ncmDrawingMaster") & hundZeros & partNum & "\Documents\"
    
    btnPaths.Add ITEM:=hundZeros & partNum & "\Documents", Key:="docHisSearch"
    btnPaths.Add ITEM:=hundZeros & partNum & "\CATIA", Key:="modelV5Search"
Else
    NCMpart = False
    mainPath = mainFolder("docHisSearch") & thousZeros & hundZeros & partNum & "\"
    
    btnPaths.Add ITEM:=thousZeros & hundZeros & partNum & "*", Key:="docHisSearch"
    btnPaths.Add ITEM:=thousZeros & hundZeros & partNum & "*", Key:="modelV5Search"
End If


direct = Dir(mainPath)
If (Right(mainPath, 10) = "DOCUMENTS\") And direct = "" Then direct = Dir(Left(mainPath, Len(mainPath) - 10)) 'NCM shortcuts are placed in the main NCM folder, not within DOCUMENTS

errorTracker = "folder"
alistDim = False

'check for files and shortcuts in main Doc His folder
Do While direct <> vbNullString
    If direct <> "." And direct <> ".." Then
        If direct Like "*.lnk" Then
            Me.viewLinks.Visible = True
            Me.shrtCutBadge.Visible = True
            Me.docHisSearch.Width = 1.75 * 1440
            GoTo nosave
        End If
        alistDim = True
        ReDim Preserve alist(i)
        alist(i) = direct
        i = i + 1
    End If
nosave:
    direct = Dir
Loop

errorTracker = "shortcuts"

'check to see if shortcut folder has anything
Dim shortcutFold, directShorts
shortcutFold = mainPath & "shortcuts\"
directShorts = Dir(shortcutFold)
Do While directShorts <> vbNullString
    If directShorts <> "." And directShorts <> ".." Then
        If directShorts Like "*.lnk" Then
            Me.viewLinks.Visible = True
            Me.shrtCutBadge.Visible = True
            Me.docHisSearch.Width = 1.75 * 1440
            GoTo nosave1
        End If
    End If
nosave1:
    directShorts = Dir
Loop

Dim fileList() As String
Dim ITEM
i = 0
ReDim Preserve fileList(1)
fileList(0) = "hi"

errorTracker = "files"

If alistDim Then
    For Each ITEM In alist
        If ITEM <> Empty Then
            Select Case True
                Case ITEM Like "*IPD*"
                    Call setDrawingCaption(" &IPD", "internal")
                Case ITEM Like "*IAD*"
                    Call setDrawingCaption(" &IAD", "internal")
                Case ITEM Like "*PPD*"
                    Call setDrawingCaption(" &PPD", "internal")
                Case ITEM Like "*ICD*"
                    Call setDrawingCaption(" I&CD", "customer")
                Case ITEM Like "*ACD*"
                    Call setDrawingCaption(" A&CD", "customer")
                Case Else
                    ReDim Preserve fileList(i)
                    Call setDrawingCaption(" &Cust", "customer")
                    fileList(i) = ITEM
                    i = i + 1
            End Select
        End If
    Next ITEM
End If

bordCol = rgb(80, 150, 80)
bordStyle = 2

On Error Resume Next
btnPaths.Add ITEM:=partNum & "*", Key:="sendSearch"
btnPaths.Add ITEM:=hundZeros & partNumLeft & "*" & partNumRight & "*", Key:="submissionsSearch"
btnPaths.Add ITEM:=thousZeros & hundZeros & partNum & "*", Key:="labdocsCNL"
btnPaths.Add ITEM:=thousZeros & partNum & "*", Key:="slbProject"
btnPaths.Add ITEM:=thousZeros & hundZeros & partNum & "*", Key:="labdocsSLB"
btnPaths.Add ITEM:=thousZeros & partNumLeft & "*" & partNumRight & "*", Key:="lvgProject"
btnPaths.Add ITEM:=Left(partNum, 2) & "000 - " & Left(partNum, 2) + 1 & "000\" & partNumLeft & "*" & partNumRight & "*", Key:="labDocsLVG"
btnPaths.Add ITEM:=thousZeros & hundZeros & partNum & "*", Key:="cnlMasterSetups"
btnPaths.Add ITEM:=thousZeros & partNum & "*", Key:="slbMasterSetups"

Dim rsLinks As Recordset, fol, altPath As String
Set rsLinks = db.OpenRecordset("SELECT link, btnName FROM tblLinks WHERE searchBorder = TRUE", dbOpenSnapshot)

errorTracker = "links2"

Do While Not rsLinks.EOF
    altPath = rsLinks!Link
    If rsLinks!btnName = "docHisSearch" And NCMpart Then altPath = mainFolder("ncmDrawingMaster")
    If rsLinks!btnName = "modelV5Search" And NCMpart Then altPath = mainFolder("ncmDrawingModelV5")

    fol = altPath & btnPaths(rsLinks!btnName)
    
    If FolderExists(fol) Then
        Me.Controls(rsLinks!btnName).BorderStyle = bordStyle
        Me.Controls(rsLinks!btnName).BorderColor = bordCol
    Else
        Me.Controls(rsLinks!btnName).BorderStyle = 0
    End If
nextLink:
    rsLinks.MoveNext
Loop

rsLinks.CLOSE
Set rsLinks = Nothing

On Error GoTo Err_Handler

errorTracker = "project"

Dim rsProj As Recordset, rsRelProj As Recordset
Set rsProj = db.OpenRecordset("SELECT recordId from tblPartProject WHERE partNumber = '" & partNum & "'")
If rsProj.RecordCount > 0 Then
    Me.openDash.BorderStyle = bordStyle
    Me.openDash.BorderColor = bordCol
Else
    Set rsRelProj = db.OpenRecordset("SELECT recordId from tblPartProjectPartNumbers WHERE childPartNumber = '" & partNum & "'")
    If rsRelProj.RecordCount > 0 Then
        Me.openDash.BorderStyle = bordStyle
        Me.openDash.BorderColor = bordCol
    Else
        Me.openDash.BorderStyle = 0
    End If
    rsRelProj.CLOSE
    Set rsRelProj = Nothing
End If

rsProj.CLOSE
Set rsProj = Nothing

Dim rsPI As Recordset
Set rsPI = db.OpenRecordset("SELECT * FROM tblPartInfo WHERE partNumber = '" & partNum & "' AND programId is not null")

If rsPI.RecordCount > 0 Then
    If Nz(rsPI!programId, 0) = 0 Then GoTo noProgram
    Me.openProgram.BorderStyle = bordStyle
    Me.openProgram.BorderColor = bordCol
    Me.openProgram.Caption = " " & DLookup("modelCode", "tblPrograms", "ID = " & rsPI!programId)
    Me.openProgram.Visible = True
Else
    Me.openProgram.BorderStyle = 0
    Me.openProgram.Visible = False
End If
noProgram:

rsPI.CLOSE
Set rsPI = Nothing

errorTracker = "quote"

Dim quoteNumber As String
quoteNumber = grabQuoteNum(partNum)

If quoteNumber <> "" Then
    Me.quoteSearchBtn.Caption = " &Quote Info: #" & quoteNumber
    Me.quoteSearchBtn.BorderStyle = bordStyle
    Me.quoteSearchBtn.BorderColor = bordCol
Else
    Me.quoteSearchBtn.Caption = " &Quote Info"
    Me.quoteSearchBtn.BorderStyle = 0
End If

exitFunc:
Me.partNumberSearch.SetFocus
DoCmd.Echo True
Me.Painting = True
DoCmd.Hourglass False
Set db = Nothing
errorTracker = "madeIt"
Exit Sub

checkForOracle:
errorTracker = "oracleError"
DoCmd.Echo True
Me.Painting = True
DoCmd.Hourglass False
If Err.number = 3151 Then
    Call setErrorText("No connection to Oracle")
    TempVars.Add "oracleConnectionFailed", "True"
    Me.NAM = partNum
    GoTo skipOracle
Else
    GoTo Err_Handler
End If

Exit Sub
Err_Handler:
DoCmd.Echo True
Me.Painting = True
DoCmd.Hourglass False
Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number, errorTracker)
End Sub

Function dayTime() As String
On Error GoTo Err_Handler

Select Case time()
    Case Is < TimeSerial(12, 0, 0)
        dayTime = "Morning"
    Case Is > TimeSerial(18, 0, 0)
        dayTime = "Evening"
    Case Else
        dayTime = "Afternoon"
End Select

Exit Function
Err_Handler:
    Call handleError(Me.name, "dayTime", Err.DESCRIPTION, Err.number)
End Function

Private Sub Form_Timer()
On Error Resume Next

Me.TimerInterval = 900000 'every 15 minutes
Call notificationsCount

End Sub

Private Sub hidecmd_Click()
On Error GoTo Err_Handler

DoCmd.Hourglass True
Application.Echo False
Me.Painting = True

Me.viewLinks.Visible = False
Me.shrtCutBadge.Visible = False
Me.docHisSearch.Width = 2.125 * 1440
Call showOracleTags(False)
Me.lblErrors.Visible = False
Me.openInternalDwg.Visible = False
Me.openCust.Visible = False
Me.Label157.Visible = False
Me.partNumberSearch.SetFocus
Me.partNumberSearch = ""
Me.partPicture.Visible = False
Me.quoteSearchBtn.Caption = " &Quote Info"
Me.openProgram.Visible = False

Me.Controls("openDash").BorderStyle = 0
Me.Controls("quoteSearchbtn").BorderStyle = 0

Dim db As Database
Set db = CurrentDb()
Dim rsLinks As Recordset
Set rsLinks = db.OpenRecordset("SELECT link, btnName FROM tblLinks WHERE searchBorder = TRUE", dbOpenSnapshot)

Do While Not rsLinks.EOF
    Me.Controls(rsLinks!btnName).BorderStyle = 0
    rsLinks.MoveNext
Loop

rsLinks.CLOSE
Set rsLinks = Nothing
Set db = Nothing

Me.partNumberSearch.SetFocus

Application.Echo True
Me.Painting = True
DoCmd.Hourglass False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub labdocsCNL_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name, Me.partNumberSearch)
Dim partNum, mainPath

mainPath = mainFolder(Me.ActiveControl.name)
partNum = Me.partNumberSearch
    If Len(partNum) = 5 Then
        Call openPath(mainPath & Left(partNum, 2) & "000\" & Left(partNum, 3) & "00\" & partNum)
    Else
        Call openPath(mainPath)
    End If
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub labDocsLVG_Click()
On Error GoTo msg
Call logClick(Me.ActiveControl.name, Me.name, Me.partNumberSearch)
Dim partNum, mainPath, level1, mainfolderpath, FolName, strFilePath, partNumLeft, partNumRight

mainPath = mainFolder(Me.ActiveControl.name)
partNum = Me.partNumberSearch
partNumLeft = Left(partNum, 4)
partNumRight = Right(partNum, 1)
level1 = Dir(mainPath & Left(partNum, 2) & "*", vbDirectory)
mainfolderpath = mainPath & level1 & "\"
FolName = Dir(mainfolderpath & partNumLeft & "*" & partNumRight & "*", vbDirectory)

If Len(FolName) = 0 Then
    Call openPath(mainfolderpath)
Else
    strFilePath = mainfolderpath & FolName
    Call openPath(strFilePath)
End If

Exit Sub

msg:
MsgBox "You may need to request access to LVG Labdocs to access this", vbCritical, "Sorry"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub labdocsSLB_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name, Me.partNumberSearch)
Dim partNum, mainPath, prtpath, FolName

partNum = Me.partNumberSearch
mainPath = mainFolder(Me.ActiveControl.name)
    If Len(partNum) = 5 Then
        prtpath = mainPath & Left(partNum, 2) & "000\" & Left(partNum, 3) & "00\"
        FolName = Dir(prtpath & partNum & "*", vbDirectory)
        Call openPath(prtpath & FolName)
    Else
        Call openPath(mainPath)
    End If
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub btnSettings_Click(): On Error GoTo Err_Handler
DoCmd.OpenForm "frmSettings"
Exit Sub
Err_Handler:: Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number): End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Dim dev As Boolean

dev = CurrentProject.Path <> "C:\workingdb"

Me.CLOSE.Visible = dev

Dim designPrograms As Boolean
designPrograms = Nz(TempVars!dept) = "Design"

Me.btn3DexSheet.Visible = designPrograms Or dev
Me.modelV5search.Visible = designPrograms Or dev
Me.catiaMacros.Visible = designPrograms Or dev

Dim Org
Org = Nz(TempVars!Org, 4)
If Org = 0 Then Org = 4
If Org < 4 Then Me.tabByOrgByPart.Value = Org - 1

Call loadUserBtns

Me.message.Caption = Nz(TempVars!Joke, "")
Me.partNumberSearch.SetFocus

Form_DASHBOARD.sfrmCalendarItems.Visible = False

Me.Caption = Environ("username")
Me.userPic.Picture = "\\data\mdbdata\WorkingDB\Pictures\Avatars\" & Environ("username") & ".png"

DoCmd.Maximize

closeApp

If TempVars!smallScreen = True Then
    Call smallScreenMode(True)
End If

Call setTheme(Me)

Exit Sub
Err_Handler:
    Call handleError("DASHBOARD", "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Function smallScreenMode(bool As Boolean)
On Error GoTo Err_Handler

'go block by block, by TAG of the items and move them all

Dim ctlVar As Control

Dim neg As Long
If bool Then
    neg = 1
Else
    neg = -1
End If

For Each ctlVar In Me.Controls
    If InStr(ctlVar.tag, "g") Then
        Select Case True
            Case InStr(ctlVar.tag, "g1") 'feedback btn/Learn more
                ctlVar.Top = ctlVar.Top - 1860 * neg '13920-12060
            Case InStr(ctlVar.tag, "g2") 'calendar stuff
                ctlVar.Left = ctlVar.Left - 16000 * neg
                ctlVar.Top = ctlVar.Top - 1740 * neg '9900-8160
            Case InStr(ctlVar.tag, "g3") 'packlist -->
                ctlVar.Left = ctlVar.Left - 10260 * neg '16800-6540
                ctlVar.Height = ctlVar.Height - 60 * neg '419-359
            Case InStr(ctlVar.tag, "g4") 'help -->
                ctlVar.Left = ctlVar.Left - 16200 * neg '27240 - 11040
            Case InStr(ctlVar.tag, "g5") 'favorites block
                ctlVar.Top = ctlVar.Top + 3720 * neg '840-4560
                ctlVar.Left = ctlVar.Left - 10800 * neg '12360-1560
            Case InStr(ctlVar.tag, "g6") 'task tracker
                ctlVar.Top = ctlVar.Top + 7400 * neg
                ctlVar.Left = ctlVar.Left - 21600 * neg
                If bool Then
                    ctlVar.Height = 4200 '8760
                Else
                    ctlVar.Height = 8760
                End If
            Case InStr(ctlVar.tag, "g7") 'app container
                ctlVar.Visible = Not bool
                If bool Then
                    ctlVar.Height = 3660
                    ctlVar.Width = 5340
                End If
        End Select
    End If
Next ctlVar

Me.appContainer.SourceObject = ""
Me.toggleTaskTracker.Visible = Not bool

If bool Then
    Me.Detail.Height = 13020
    Me.Width = 12440
Else
    Me.appContainer.Visible = False
    Me.appContainer.Height = 10140
    Me.appContainer.Width = 21300
    Me.Command590.Width = 21419
    Me.Command590.Height = 10259
    Me.Detail.Height = 14835
    Me.Width = 28440
End If

Exit Function
Err_Handler:
    Call handleError("DASHBOARD", "smallScreenMode", Err.DESCRIPTION, Err.number)
End Function

Private Sub help_Click(): On Error GoTo Err_Handler
Call openPath(mainFolder(Me.ActiveControl.name))
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lvgData_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name, Me.partNumberSearch)
Dim partNum, prtpath

partNum = Me.partNumberSearch
prtpath = mainFolder("lvgProject") & Left(partNum, 2) & "000\"

Dim FolName As String
    ChDir "S:\"
    FolName = Dir(prtpath & partNum & "*", vbDirectory)
        
Call openPath(prtpath & FolName & "\QUALITY DOCUMENTS\IR_KFD_ISIR")
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lvgProject_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name, Me.partNumberSearch)
Dim partNum, thousZeros, partNumLeft, partNumRight, mainPath, FolName
partNum = Me.partNumberSearch
mainPath = mainFolder(Me.ActiveControl.name)

If Len(partNum) = 5 Then
    thousZeros = Left(partNum, 2) & "000\"
    partNumLeft = Left(partNum, 4)
    partNumRight = Right(partNum, 1)
    ChDir "S:\"
    FolName = Dir(mainPath & thousZeros & partNumLeft & "*" & partNumRight & "*", vbDirectory)
    Call openPath(mainPath & thousZeros & FolName)
Else
    Call openPath(mainPath)
End If
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub message_Click()
On Error GoTo Err_Handler
Dim dayTimeVal, dayName, currentMessage

Me.partPicture.Visible = False
If Len(Me.partNumberSearch) > 4 Then
    Dim partPic, partPicDir
    partPicDir = "\\data\mdbdata\WorkingDB\_docs\Part_Pictures\"
    partPic = Dir((partPicDir & Me.partNumberSearch & "*"))
    If Len(partPic) > 0 Then
        Me.partPicture.Picture = partPicDir & partPic
        Me.partPicture.Visible = True
    End If
End If

Dim db As Database
Set db = CurrentDb()
Dim rs1 As Recordset
Set rs1 = db.OpenRecordset("SELECT max(ID), DatabaseVersion FROM tblReleaseNotes WHERE databaseName = 'Front_End' group by databaseversion")
rs1.MoveLast

If grabVersion() <> rs1!DatabaseVersion Then
    Me.message.Caption = "New WorkingDB version available" & vbNewLine & vbNewLine & "-Please reopen-"
    Me.message.BorderColor = rgb(180, 40, 40)
    Me.message.BackColor = rgb(60, 40, 40)
    Me.message.BorderWidth = 2
    Me.message.BorderStyle = 1
    Me.message.ForeColor = vbWhite
    Me.message.FontBold = True
    Me.message.fontSize = 14
    Me.partPicture.Visible = False
    Me.btnEditPicture.Visible = False
    Exit Sub
End If

currentMessage = DLookup("[Message]", "tblDBinfoBE", "[ID] = 1")
If Nz(currentMessage) <> "" Then
    Me.message.Caption = currentMessage
    Me.partPicture.Visible = False
End If

rs1.CLOSE
Set rs1 = Nothing
Set db = Nothing
Exit Sub
Err_Handler:
Me.message.Caption = "Good " & dayName & " " & dayTimeVal & "."
End Sub

Private Sub modelV5search_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name, Me.partNumberSearch)
Call openModelV5Folder(Me.partNumberSearch)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub notifications_Click(): On Error GoTo Err_Handler

DoCmd.OpenForm "frmNotifications"

Exit Sub
Err_Handler:: Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number): End Sub

Public Sub openCust_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name, Me.partNumberSearch)

Dim thousZeros As String, hundZeros As String, mainPath As String
Dim partNum, dochispath, IPDName, IADName, PPDName
partNum = Me.partNumberSearch

Dim fileList As Collection
Set fileList = New Collection

Dim direct As String
thousZeros = Left(partNum, 2) & "000\"
hundZeros = Left(partNum, 3) & "00\"
mainPath = mainFolder("docHisSearch") & thousZeros & hundZeros & partNum & "\"
direct = Dir(mainPath)

Do While direct <> vbNullString
    If direct <> "." And direct <> ".." Then
        Select Case True
            Case direct = "." Or direct = ".." 'filter out the dir generated dots
                GoTo nosave
            Case direct Like "*.lnk", direct Like "*IPD*", direct Like "*IAD*", direct Like "*PPD*" 'filter anything internal and links
                GoTo nosave
            Case Else 'ANYTHING other than internal drawings, folders, and links will be treated as a customer drawing
                fileList.Add direct
        End Select
    End If
nosave:
    direct = Dir
Loop

If fileList.count > 1 Then 'if for some reason there is more than one customer drawing, open the folder instead of a file
    openPath (mainPath)
Else
    openPath (addLastSlash(mainPath) & fileList(1))
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub openDash_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name, Me.partNumberSearch)
If (Nz(Me.partNumberSearch) = "") Then
    MsgBox "First, enter a part number", vbOKOnly, "Can't Open"
    Exit Sub
End If

openPartProject (Me.partNumberSearch)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Public Sub openInternalDwg_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name, Me.partNumberSearch)

Dim thousZeros As String, hundZeros As String, mainPath As String
Dim partNum, dochispath, IPDName, IADName, PPDName
partNum = Me.partNumberSearch

Dim fileList As Collection
Set fileList = New Collection

Dim direct As String
thousZeros = Left(partNum, 2) & "000\"
hundZeros = Left(partNum, 3) & "00\"
mainPath = mainFolder("docHisSearch") & thousZeros & hundZeros & partNum & "\"
direct = Dir(mainPath)

Do While direct <> vbNullString
    If direct <> "." And direct <> ".." Then
        If direct Like "*.lnk" Then GoTo nosave
        Select Case True
            Case direct Like "*.lnk" 'filter out shortcuts
                GoTo nosave
            Case direct Like "*IPD*", direct Like "*IAD*", direct Like "*PPD*" 'add all internal names drawings
                fileList.Add direct
        End Select
    End If
nosave:
    direct = Dir
Loop

Select Case fileList.count
    Case 1 'only one file found - open it
        openPath (addLastSlash(mainPath) & fileList(1))
    Case 0 'no file found, send error
        showOracleTags (False)
        setErrorText ("No drawing found")
    Case Else 'multiple found, open the folder
        openPath (mainPath)
End Select
    
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub openProgram_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)

DoCmd.OpenForm "frmProgramReview"

Dim Program
Program = Me.openProgram.Caption

Form_frmProgramReview.txtFilterInput.Value = Right(Program, Len(Program) - 1)
Form_frmProgramReview.filterByProgram_Click

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub packList_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)

DoCmd.OpenForm "frmPackingList"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub quoteSearchBtn_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name, Me.partNumberSearch)
Dim partNum As String, mainPath, FolName, strFilePath, prtFilePath, quoteNumber As String

partNum = Me.partNumberSearch.Value
mainPath = mainFolder(Me.ActiveControl.name)

If Me.quoteSearchBtn.BorderStyle = 2 Then 'quote number was found earlier
    quoteNumber = grabQuoteNum(partNum)
Else
    quoteNumber = partNum 'the person most likely just search directly for a quote number
End If

If Len(quoteNumber) = 5 Then
    prtFilePath = mainPath & zeros(quoteNumber, 2)
    FolName = Dir(prtFilePath & quoteNumber, vbDirectory)
    strFilePath = prtFilePath & FolName
    If Len(FolName) = 0 Then
        If MsgBox("Do you want to go to the main folder?", vbYesNo, "Folder Does not Exist") = vbYes Then Call openPath(mainPath)
        Exit Sub
    Else
        Call openPath(strFilePath)
    End If
Else
    Call openPath(mainPath)
End If
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub revisedItems_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name, Me.partNumberSearch)
DoCmd.OpenForm "frmECOpartHistory"
Exit Sub
Err_Handler:: Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number): End Sub

Private Sub sendSearch_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name, Me.partNumberSearch)
Dim partNum, FolName, mainPath

partNum = Me.partNumberSearch
mainPath = mainFolder(Me.ActiveControl.name)

Call checkMkDir(mainPath, partNum, "*")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub slbProject_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name, Me.partNumberSearch)
Dim partNum, thousZeros, FolName, mainPath, strPath

partNum = Me.partNumberSearch
thousZeros = Left(partNum, 2) & "000\"
mainPath = mainFolder(Me.ActiveControl.name)
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

Public Sub slbSDS_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name, Me.partNumberSearch)
Dim partNum, FolName, firstPath

partNum = Me.partNumberSearch

If Len(partNum) < 1 Then
    MsgBox "Must enter part number to go to SLB SDS folder", vbOKOnly, "Error"
    Exit Sub
End If

firstPath = mainFolder("slbProject") & Left(partNum, 2) & "000\"

FolName = Dir(firstPath & partNum & "*", vbDirectory)
Call openPath(firstPath & FolName & "\QUALITY DOCUMENTS\SDS FOLDER")
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub submissionsSearch_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name, Me.partNumberSearch)
Dim partNum, mainPath, FolName, firstPath, partNumLeft, partNumRight

partNum = Me.partNumberSearch
mainPath = mainFolder(Me.ActiveControl.name)

If Len(partNum) = 5 Then
    firstPath = mainPath & Left(partNum, 3) & "00\"
    partNumLeft = Left(partNum, 4)
    partNumRight = Right(partNum, 1)
    FolName = Dir(firstPath & partNumLeft & "*" & partNumRight & "*", vbDirectory)
    Call openPath(firstPath & FolName)
Else
    Call openPath(mainPath)
End If
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub toggleTaskTracker_Click()
On Error GoTo Err_Handler

Me.sfrmTasks.Visible = Not Me.sfrmTasks.Visible

If Me.sfrmTasks.Visible Then
    Me.toggleTaskTracker.Picture = "\\data\mdbdata\WorkingDB\Pictures\Core\upArrow.ico"
Else
    Me.toggleTaskTracker.Picture = "\\data\mdbdata\WorkingDB\Pictures\Core\downArrow.ico"
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub userCmd_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmUserView"

Exit Sub
Err_Handler:
        Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub viewLinks_Click()
On Error GoTo Err_Handler
If Len(Me.partNumberSearch) = 5 Then DoCmd.OpenForm "frmPartShortcuts"
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Public Sub xref_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name, Me.partNumberSearch)
DoCmd.OpenForm "frmCustomerXref", , , "[NAM] = '" & Nz(Form_DASHBOARD.partNumberSearch, "26589") & "'"

Form_frmCustomerXref.NAMsrchBox = Form_DASHBOARD.partNumberSearch
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Function openMasterSetups()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name, Me.partNumberSearch)

Dim mainPath As String, adFilter As String, partNum As String
partNum = Nz(Me.partNumberSearch, "")
mainPath = mainFolder(Me.ActiveControl.name)

adFilter = ""
If Len(partNum) > 3 Then adFilter = "?FilterField1=Part%5Fx0020%5FNumber&FilterValue1=" & partNum

openPath (mainPath & adFilter)

Exit Function
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Function

Private Sub lvgMasterSetups_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)
Call openPath(mainFolder(Me.ActiveControl.name))
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Public Sub loadUserBtns()
On Error GoTo Err_Handler

Dim db As Database
Set db = CurrentDb()
Dim rsUserButton As Recordset
Set rsUserButton = db.OpenRecordset("SELECT * FROM tblUserButtons WHERE User = '" & Environ("username") & "'", dbOpenSnapshot)

Dim i, start
For i = 1 To 10
    rsUserButton.FindFirst "ButtonNum = '" & i & "'"
    If i < 10 Then
        start = "  &" & i & " "
    Else
        start = " 1&0 "
    End If
    If rsUserButton.noMatch Then
        Me.Controls("userdef" & i).Caption = "  not set"
        Me.Controls("userdef" & i).ForeColor = rgb(150, 150, 150)
        Me.Controls("userdef" & i).FontBold = False
    Else
        Me.Controls("userdef" & i).Caption = start & rsUserButton!Caption
        Me.Controls("userdef" & i).ForeColor = rgb(255, 255, 255)
        Me.Controls("userdef" & i).FontBold = True
    End If
Next i

rsUserButton.CLOSE
Set rsUserButton = Nothing
Set db = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, "loadUserBtns", Err.DESCRIPTION, Err.number)
End Sub

Private Sub setUserDef_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)
DoCmd.OpenForm "frmUserButtonChange"
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub userDefinedButtons(number As String)
On Error GoTo Err_Handler

Dim btnLink As String
btnLink = Nz(DLookup("Link", "tblUserButtons", "[ButtonNum] = '" & number & "' And [User] = '" & Environ("username") & "'"), "not found")

If btnLink = "not found" Then
    MsgBox "To set up your user button, click ''Set Favorites' below", vbInformation, "No Link Set Up Yet"
    Exit Sub
Else
    Call openPath(btnLink)
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, "userDefinedButtons", Err.DESCRIPTION, Err.number)
End Sub

Private Sub userdef1_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)
Call userDefinedButtons("1")
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub userdef10_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)
Call userDefinedButtons("10")
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub userdef2_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)
Call userDefinedButtons("2")
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub userdef3_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)
Call userDefinedButtons("3")
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub userdef4_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)
Call userDefinedButtons("4")
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub userdef5_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)
Call userDefinedButtons("5")
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub userdef6_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)
Call userDefinedButtons("6")
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub userdef7_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)
Call userDefinedButtons("7")
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub userdef8_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)
Call userDefinedButtons("8")
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub userdef9_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)
Call userDefinedButtons("9")
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub yearInReview_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)

DoCmd.OpenForm "frmYearInReview"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
