Option Compare Database
Option Explicit

Dim XL As Excel.Application, WB As Excel.Workbook, WKS As Excel.Worksheet
Dim inV As Long

Function openPartProject(partNum As String) As Boolean
On Error GoTo err_handler

Form_DASHBOARD.partNumberSearch = partNum
TempVars.Add "partNumber", partNum

If DCount("recordId", "tblPartProject", "partNumber = '" & partNum & "'") > 0 Then GoTo openIt 'if there is a project for this, open it
If DCount("recordId", "tblPartProjectPartNumbers", "childPartNumber = '" & partNum & "'") > 0 Then GoTo openIt 'if there is a related project for this, open it

If Form_DASHBOARD.NAM <> partNum Then
    MsgBox "Please search this part before opening the dash", vbInformation, "Sorry."
    Exit Function
End If

If userData("org") = 5 Then GoTo openIt 'bypass Oracle restrictions for NCM users

If Form_DASHBOARD.lblErrors.Visible = True And Form_DASHBOARD.lblErrors.Caption = "Part not found in Oracle" Then
    MsgBox "This part number must show up in Oracle to open the dash", vbInformation, "Sorry."
    Exit Function
End If

openIt:
If (CurrentProject.AllForms("frmPartDashboard").IsLoaded = True) Then DoCmd.Close acForm, "frmPartDashboard"
DoCmd.OpenForm "frmPartDashboard"

Exit Function
err_handler:
    Call handleError("wdbProjectE", "openPartProject", Err.DESCRIPTION, Err.number)
End Function

Public Function checkAIFfields(partNum As String) As Boolean
On Error GoTo err_handler
checkAIFfields = False

'---Setup Variables---
Dim db As Database
Set db = CurrentDb()
Dim rsPI As Recordset, rsPack As Recordset, rsPackC As Recordset, rsComp As Recordset, rsAI As Recordset, rsU As Recordset
Dim rsPE As Recordset, rsPMI As Recordset

Dim errorArray As Collection
Set errorArray = New Collection

If findDept(partNum, "Project", True) = "" Then errorArray.Add "Project Engineer"

'---Grab General Data---
Set rsPI = db.OpenRecordset("SELECT * from tblPartInfo WHERE partNumber = '" & partNum & "'")
Set rsPack = db.OpenRecordset("SELECT * from tblPartPackagingInfo WHERE partInfoId = " & rsPI!recordId & " AND packType = 1")
Set rsU = db.OpenRecordset("SELECT * from tblUnits WHERE recordId = " & rsPI!unitId)

'check part info stuff - always reqruied
If Nz(rsPI!dataStatus) = "" Then errorArray.Add "Data Status"
If Nz(rsPI!customerId) = "" Then errorArray.Add "Customer"
If Nz(rsPI!developingLocation) = "" Then errorArray.Add "Developing Org"

If rsPI!dataStatus = 2 Then
    If Nz(rsPI!unitId) = "" Then errorArray.Add "MP Unit"
    If Nz(rsPI!developingUnit) = "" Then errorArray.Add "Dev Unit"
End If

If Nz(rsPI!partType) = "" Then errorArray.Add "Part Type"
If Nz(rsPI!finishLocator) = "" Then errorArray.Add "Finish Locator"
If Nz(rsPI!finishSubInv) = "" Then errorArray.Add "Finish Sub-Inventory"
If Nz(rsPI!quoteInfoId) = "" Then errorArray.Add "Quote Information"
If Nz(DLookup("quotedCost", "tblPartQuoteInfo", "recordId = " & rsPI!quoteInfoId)) = "" Then errorArray.Add "Quoted Cost"
If Nz(rsPI!sellingPrice) = "" Then errorArray.Add "Selling Price" 'required always if FG

If rsPI!partType = 1 Or rsPI!partType = 4 Then 'molded / new color
    If Nz(rsPI!moldInfoId) = "" Then
        errorArray.Add "Molding Info" 'always required
        GoTo skipMold
    End If
    
    Set rsPMI = db.OpenRecordset("SELECT * from tblPartMoldingInfo WHERE recordId = " & rsPI!moldInfoId)

    'always required
    If Nz(rsPMI!inspection) = "" Then errorArray.Add "Tool Inspection Level"
    If Nz(rsPMI!measurePack) = "" Then errorArray.Add "Tool Measure Pack Level"
    If Nz(rsPMI!annealing) = "" Then errorArray.Add "Tool Annealing Level"
    If Nz(rsPMI!automated) = "" Then errorArray.Add "Tool Automation Type"
    If Nz(rsPMI!toolType) = "" Then errorArray.Add "Tool Level"
    If Nz(rsPMI!gateCutting) = "" Then errorArray.Add "Tool Gate Level"
    If Nz(rsPI!materialNumber) = "" Then errorArray.Add "Material Number"
    If Nz(rsPI!pieceWeight) = "" Then errorArray.Add "Piece Weight"
    If Nz(rsPI!materialNumber1) <> "" Then 'if there is a second material, must enter wieght for that material
        If Nz(rsPI!matNum1PieceWeight) = "" Then errorArray.Add "Second Material Piece Weight"
    End If
    If Nz(rsPMI!toolNumber) = "" Then errorArray.Add "Tool Number"
    If Nz(rsPMI!pressSize) = "" Then errorArray.Add "Press Tonnage"
    If Nz(rsPMI!piecesPerHour) = "" Then errorArray.Add "Pieces Per Hour"
    
    If rsPI!dataStatus = 2 Then 'required for transfer
        If Nz(rsPI!itemWeight100Pc) = "" And rsPI!unitId = 1 Then errorArray.Add "100 Piece Weight" 'U01 only
        If Nz(rsPMI!assignedPress) = "" Then errorArray.Add "Assigned Press"
    End If
    
    rsPMI.Close
    Set rsPMI = Nothing
End If
skipMold:

If rsPI!partType = 2 Or rsPI!partType = 5 Then
    If Nz(rsPI!assemblyInfoId) = "" Then
        errorArray.Add "Assembly Info"
        GoTo skipAssy
    End If
    
    Set rsAI = db.OpenRecordset("SELECT * from tblPartAssemblyInfo WHERE recordId = " & rsPI!assemblyInfoId)

    'always required
    If Nz(rsAI!assemblyType) = "" Then errorArray.Add "Assembly Type"
    If Nz(rsAI!assemblyAnnealing) = "" Then errorArray.Add "Assembly Annealing Level"
    If Nz(rsAI!assemblyInspection) = "" Then errorArray.Add "Assembly Inspection Level)"
    If Nz(rsAI!assemblyMeasPack) = "" Then errorArray.Add "Assembly Measure Pack Level"
    If Nz(rsAI!partsPerHour) = "" Then errorArray.Add "Assembly Parts Per Hour"
    
    If rsPI!dataStatus = 2 Then 'required for transfer
        If Nz(rsAI!assemblyWeight100Pc) = "" Then errorArray.Add "100 Piece Weight"
        If Nz(rsAI!resource) = "" Then errorArray.Add "Assembly Resource"
        If Nz(rsAI!machineLine) = "" Then errorArray.Add "Assembly Machine Line"
    End If

    rsAI.Close
    Set rsAI = Nothing
    
    Set rsComp = db.OpenRecordset("SELECT * from tblPartComponents WHERE assemblyNumber = '" & partNum & "'")
    If rsComp.RecordCount = 0 Then
        errorArray.Add "Component Information"
        GoTo skipAssy
    End If
    
    Do While Not rsComp.EOF
        'always required
        If Nz(rsComp!componentNumber) = "" Then errorArray.Add "Blank Component Number" 'always required
        If Nz(rsComp!quantity) = "" Then errorArray.Add "Blank Component Quantity" 'always required
        
        If rsPI!dataStatus = 2 Then 'required for transfer
            If Nz(rsComp!finishLocator) = "" Then errorArray.Add "Component Finish Locator"
            If Nz(rsComp!finishSubInv) = "" Then errorArray.Add "Component Sub-Inventory"
        End If
        
        rsComp.MoveNext
    Loop
    rsComp.Close
    Set rsComp = Nothing
End If
skipAssy:

If rsPack.RecordCount = 0 Then
    If rsPI!dataStatus = 2 Then errorArray.Add "Packaging Information" 'required for transfer
Else
    Do While Not rsPack.EOF
        If Nz(rsPack!packType) = "" Then errorArray.Add "Packaging Type" 'required for transfer
        If rsU!Org = "CUU" Then
            If Nz(rsPack!boxesPerSkid) = "" Then errorArray.Add "Boxes Per Skid" 'if CUU org, then need to check this for transfer for MEX FREIGHT cost calc
        End If

        Set rsPackC = db.OpenRecordset("SELECT * from tblPartPackagingComponents WHERE packagingInfoId = " & rsPack!recordId)
        If rsPackC.RecordCount = 0 And rsPI!dataStatus = 2 Then errorArray.Add "Packaging Components" 'required for transfer
        
        Do While Not rsPackC.EOF
            If rsPI!dataStatus = 2 Then 'required for transfer
                If Nz(rsPackC!componentType) = "" Then errorArray.Add "Packaging Component Type"
                If Nz(rsPackC!componentPN) = "" Then errorArray.Add "Packaging Component Part Number"
                If Nz(rsPackC!componentQuantity) = "" Then errorArray.Add "Packing Component Quantity"
            End If
            rsPackC.MoveNext
        Loop
        rsPack.MoveNext
        rsPackC.Close: Set rsPackC = Nothing
    Loop
    
    rsPack.Close: Set rsPack = Nothing
End If

If rsPI!unitId = 3 And rsPI!dataStatus = 2 Then 'if U06 - these are required for transfer
    If Nz(rsPI!outsourceInfoId) = "" Then
        errorArray.Add "Outsource Info"
    Else
        If Nz(DLookup("outsourceCost", "tblPartOutsourceInfo", "recordId = " & rsPI!outsourceInfoId)) = "" Then errorArray.Add "Outsource Cost"
    End If
End If

If errorArray.count > 0 Then GoTo sendMsg

checkAIFfields = True
GoTo exitFunction

sendMsg:
Dim errorTxtLines As String, element
errorTxtLines = ""
For Each element In errorArray
    errorTxtLines = errorTxtLines & vbNewLine & element
Next element

MsgBox "Please fix these items:" & vbNewLine & errorTxtLines, vbOKOnly, "Fix this to export"

exitFunction:
On Error Resume Next
rsPI.Close: Set rsPI = Nothing
rsPack.Close: Set rsPack = Nothing
rsPackC.Close: Set rsPackC = Nothing
rsComp.Close: Set rsComp = Nothing
rsAI.Close: Set rsAI = Nothing
rsPMI.Close: Set rsPMI = Nothing
rsU.Close: Set rsU = Nothing
Exit Function

err_handler:
    Call handleError("wdbProjectE", "checkAIFfields", Err.DESCRIPTION, Err.number)
    GoTo exitFunction
End Function

Public Function exportAIF(partNum As String) As String
On Error GoTo err_handler
exportAIF = False

'---Setup Variables---
Dim db As Database
Set db = CurrentDb()
Dim rsPI As Recordset, rsPack As Recordset, rsPackC As Recordset, rsComp As Recordset, rsAI As Recordset
Dim rsPE As Recordset, rsOI As Recordset, rsU As Recordset, rsPMI As Recordset, rsDevU As Recordset
Dim outsourceCost As String
Dim mexFr As String, cartQty, mat0 As Double, mat1 As Double, resourceCSV() As String, ITEM, resID As Long, orgID As Long

'---Grab General Data---
Set rsPI = db.OpenRecordset("SELECT * from tblPartInfo WHERE partNumber = '" & partNum & "'")
Set rsPack = db.OpenRecordset("SELECT * from tblPartPackagingInfo WHERE partInfoId = " & rsPI!recordId)
Set rsU = db.OpenRecordset("SELECT * from tblUnits WHERE recordId = " & rsPI!unitId)
Set rsDevU = db.OpenRecordset("SELECT * from tblUnits WHERE recordId = " & Nz(rsPI!developingUnit, 0))
Set rsPE = CurrentDb().OpenRecordset("SELECT * from tblPermissions where Dept = 'Project' AND Level = 'Engineer' AND user IN " & _
                                    "(SELECT person FROM tblPartTeam WHERE partNumber = '" & partNum & "')")

mexFr = "0"
If rsU!Org = "CUU" And rsPI!dataStatus = 2 Then
    cartQty = Nz(DLookup("componentQuantity", "tblPartPackagingComponents", "packagingInfoId = " & rsPack!recordId & " AND componentType = 1"))
    mexFr = (cartQty * rsPack!boxesPerSkid)
    If mexFr <> 0 Then mexFr = 83.7 / (cartQty * rsPack!boxesPerSkid)
End If

outsourceCost = "0"
If Nz(rsPI!outsourceInfoId) <> "" Then
    Set rsOI = db.OpenRecordset("SELECT * from tblPartOutsourceInfo WHERE recordId = " & rsPI!outsourceInfoId)
    outsourceCost = Nz(rsOI!outsourceCost)
    rsOI.Close: Set rsOI = Nothing
End If
                                    
'---Setup Excel Form---
Set XL = New Excel.Application
Set WB = XL.Workbooks.Add
XL.Visible = False
WB.Activate
Set WKS = WB.ActiveSheet
WKS.name = "MAIN"
WKS.Range("A:E").HorizontalAlignment = xlCenter
WKS.Range("A:E").VerticalAlignment = xlCenter
inV = 1

'---Import General Data---
WKS.Range("A1:E1").Font.Italic = True
aifInsert "ACCOUNTING INFORMATION FORM", "", , "Exported: ", Date
aifInsert "PRIMARY INFORMATION", "", , , , True
aifInsert "Part Number", partNum, firstColBold:=True
aifInsert "Data Status", DLookup("partDataStatus", "tblDropDownsSP", "ID = " & rsPI!dataStatus), firstColBold:=True
aifInsert "Planner", rsPE!firstName & " " & rsPE!lastName, firstColBold:=True
aifInsert "Mark Code", Nz(rsPI!partMarkCode), firstColBold:=True
aifInsert "Customer", DLookup("CUSTOMER_NAME", "APPS_XXCUS_CUSTOMERS", "CUSTOMER_ID = " & rsPI!customerId), firstColBold:=True

If rsPI!dataStatus = 2 Then
    aifInsert "MP Unit", rsU!unitName, firstColBold:=True
    aifInsert "In-House Unit", rsDevU!unitName, firstColBold:=True
Else
    aifInsert "Unit", "U12", firstColBold:=True
End If

aifInsert "Mexico Rates", Nz(rsU!Org) = "CUU", firstColBold:=True
aifInsert "Org", Nz(rsU!Org, rsPI!developingLocation), firstColBold:=True  'is this supposed to be UNIT based, or the developing ORG?
aifInsert "Part Type", DLookup("partType", "tblDropDownsSP", "ID = " & rsPI!partType), firstColBold:=True
aifInsert "Routing Finish", Nz(DLookup("finishLocator", "tblDropDownsSP", "ID = " & rsPI!finishLocator)), firstColBold:=True
aifInsert "Sub-Location", Nz(DLookup("finishSubInv", "tblDropDownsSP", "ID = " & rsPI!finishSubInv)), firstColBold:=True
aifInsert "Mexico Freight", mexFr, firstColBold:=True, set5Dec:=True
aifInsert "Quoted Cost", Nz(DLookup("quotedCost", "tblPartQuoteInfo", "recordId = " & rsPI!quoteInfoId), 0), firstColBold:=True, set5Dec:=True
aifInsert "Selling Price", Nz(rsPI!sellingPrice), firstColBold:=True, set5Dec:=True
aifInsert "Royalty", Nz(rsPI!sellingPrice) * 0.03, firstColBold:=True, set5Dec:=True
aifInsert "Outsource Cost", outsourceCost, firstColBold:=True, set5Dec:=True

'---Molding / Assembly Specific Information---
Dim insLev As String, mpLev As String, anneal As String, laborType As String, pph As String, weight100Pc As String, orgCalc
Select Case rsPI!partType
    Case 1, 4 'molded / new color
        aifInsert "MOLDING INFORMATION", "", , , , True
        Set rsPMI = db.OpenRecordset("SELECT * from tblPartMoldingInfo WHERE recordId = " & rsPI!moldInfoId)
        weight100Pc = Nz(rsPI!itemWeight100Pc, 0)
        insLev = Nz(rsPMI!inspection)
        mpLev = Nz(rsPMI!measurePack)
        anneal = Nz(rsPMI!annealing)
        If rsPMI!insertMold Then
            laborType = "Insert Mold"
        Else
            laborType = DLookup("pressAutomation", "tblDropDownsSP", "ID = " & rsPMI!automated)
        End If
        pph = Nz(rsPMI!piecesPerHour)
        aifInsert "Tool Number", rsPMI!toolNumber, firstColBold:=True
        aifInsert "Press Tonnage", Nz(rsPMI!pressSize), firstColBold:=True
        aifInsert "Home Press", Nz(rsPMI!assignedPress), firstColBold:=True
        aifInsert "Tooling Lvl", rsPMI!toolType, firstColBold:=True
        aifInsert "Gate Lvl", rsPMI!gateCutting, firstColBold:=True
        aifInsert "Insert Mold", rsPMI!insertMold, firstColBold:=True
        aifInsert "Family Mold", rsPMI!familyTool, firstColBold:=True
        If rsPI!glass Then
            aifInsert "Glass Cost", DLookup("pressRate", "tblDropDownsSP", "pressSize = '" & rsPMI!pressSize & "'") / rsPMI!piecesPerHour / 408 / 12 / 0.85, firstColBold:=True, set5Dec:=True
        Else
            aifInsert "Glass Cost", "0", firstColBold:=True, set5Dec:=True
        End If
        If rsPI!regrind Then
            mat0 = 0: mat1 = 0
            orgCalc = Replace(Nz(rsU!Org, rsPI!developingLocation), "CUU", "MEX")
            orgID = DLookup("ID", "tblOrgs", "Org = '" & orgCalc & "'")
            If Nz(rsPI!materialNumber) <> "" Then
                mat0 = gramsToLbs(rsPI!pieceWeight) * 0.06 * DLookup("ITEM_COST", "APPS_CST_ITEM_COST_TYPE_V", "COST_TYPE = 'Frozen' AND ITEM_NUMBER = '" & Nz(rsPI!materialNumber) & "' AND ORGANIZATION_ID = " & orgID)
            End If
            If Nz(rsPI!materialNumber1) <> "" Then
                mat1 = gramsToLbs(rsPI!matNum1PieceWeight) * 0.06 * DLookup("ITEM_COST", "APPS_CST_ITEM_COST_TYPE_V", "COST_TYPE = 'Frozen' AND ITEM_NUMBER = '" & Nz(rsPI!materialNumber1) & "' AND ORGANIZATION_ID = " & orgID)
            End If
            aifInsert "Regrind Cost", mat0 + mat1, firstColBold:=True, set5Dec:=True 'multiple piece weight
        Else
            aifInsert "Regrind Cost", "0", firstColBold:=True, set5Dec:=True
        End If
        aifInsert "Material Number 1", Nz(rsPI!materialNumber), firstColBold:=True
        aifInsert "Piece Weight (lb)", gramsToLbs(Nz(rsPI!pieceWeight)), firstColBold:=True, set5Dec:=True
        aifInsert "Material Number 2", Nz(rsPI!materialNumber1), firstColBold:=True
        aifInsert "Material 2 Piece Weight (lb)", gramsToLbs(Nz(rsPI!matNum1PieceWeight)), firstColBold:=True, set5Dec:=True
        rsPMI.Close
        Set rsPMI = Nothing
    Case 2, 5 'Assembled / subassembly
        aifInsert "ASSEMBLY INFORMATION", "", , , , True
        Set rsAI = db.OpenRecordset("SELECT * from tblPartAssemblyInfo WHERE recordId = " & rsPI!assemblyInfoId)
        weight100Pc = Nz(rsAI!assemblyWeight100Pc, 0)
        laborType = DLookup("assemblyType", "tblDropDownsSP", "ID = " & rsAI!assemblyType)
        anneal = Nz(rsAI!assemblyAnnealing, 0)
        insLev = Nz(rsAI!assemblyInspection, 0)
        mpLev = Nz(rsAI!assemblyMeasPack, 0)
        pph = Nz(rsAI!partsPerHour)
        
        resID = 1
        If InStr(rsAI!resource, ",") Then
            resourceCSV = Split(rsAI!resource, ",")
            For Each ITEM In resourceCSV
                aifInsert "Resource " & resID, CStr(ITEM), firstColBold:=True
                resID = resID + 1
            Next ITEM
        End If
        
        aifInsert "Machine Line", Nz(rsAI!machineLine), firstColBold:=True
        rsAI.Close
        Set rsAI = Nothing
    Case 3 'Purchased
End Select

aifInsert "100 Piece Weight (lb)", gramsToLbs(weight100Pc), firstColBold:=True, set5Dec:=True
aifInsert "Pieces Per Hour", pph, firstColBold:=True
aifInsert "Labor Type", laborType, firstColBold:=True
aifInsert "Inspection Lvl", insLev, firstColBold:=True
aifInsert "MsPack Lvl", mpLev, firstColBold:=True
aifInsert "Annealing Lvl", anneal, firstColBold:=True

'---Component Information---
Set rsComp = db.OpenRecordset("SELECT * from tblPartComponents WHERE assemblyNumber = '" & partNum & "'")
If rsComp.RecordCount > 0 Then
    aifInsert "COMPONENT INFORMATION", "", , , , True
    aifInsert "Part Number", "Description", "Qty", "Locator", "Sub-Inventory", , True
End If
Do While Not rsComp.EOF
    aifInsert rsComp!componentNumber, _
        findDescription(rsComp!componentNumber), _
        rsComp!quantity, _
        rsComp!finishLocator, _
        Nz(DLookup("finishSubInv", "tblDropDownsSP", "ID = " & rsComp!finishSubInv))
    rsComp.MoveNext
Loop
rsComp.Close
Set rsComp = Nothing

'---Packaging Information---
Dim packType As String
If rsPack.RecordCount > 0 Then
    aifInsert "PACKAGING INFORMATION", "", , , , True
End If
Do While Not rsPack.EOF
    packType = DLookup("packagingType", "tblDropDownsSP", "ID = " & rsPack!packType)
    Set rsPackC = db.OpenRecordset("SELECT * from tblPartPackagingComponents WHERE packagingInfoId = " & rsPack!recordId)
    If rsPackC.RecordCount > 0 Then aifInsert "Packaging Type", "Component Type", "Component Number", "Component Qty", , , True
    Do While Not rsPackC.EOF
        aifInsert packType, Nz(DLookup("packComponentType", "tblDropDownsSP", "ID = " & rsPackC!componentType)), Nz(rsPackC!componentPN), Nz(rsPackC!componentQuantity)
        rsPackC.MoveNext
    Loop
    rsPack.MoveNext
Loop

'---Formatting---
WKS.Cells.columns.AutoFit
WKS.Range("B3:B4").Font.Size = 26
WKS.Range("A1:E" & inV - 1).BorderAround Weight:=xlMedium

'---Finish Up---
Dim FileName As String
FileName = "H:\" & partNum & "_Accounting_Info_" & nowString & ".xlsx"
WB.SaveAs FileName, , , , True
MsgBox "Export Complete. File path: " & FileName, vbOKOnly, "Notice"

'---Cleanup---
XL.Visible = True
Set XL = Nothing
Set WKS = Nothing
Set XL = Nothing

On Error Resume Next
rsPI.Close: Set rsPI = Nothing
rsU.Close: Set rsU = Nothing
rsPack.Close: Set rsPack = Nothing
rsPackC.Close: Set rsPackC = Nothing

exportAIF = FileName

Exit Function
err_handler:
    Call handleError("wdbProjectE", "exportAIF", Err.DESCRIPTION, Err.number)
End Function

Function aifInsert(columnVal0 As String, columnVal1 As String, Optional columnVal2 As String = ".", Optional columnVal3 As String = ".", Optional columnVal4 As String = ".", _
                                Optional heading As Boolean = False, Optional Title As Boolean = False, Optional firstColBold As Boolean = False, Optional set5Dec = False)
On Error GoTo err_handler

WKS.Cells(inV, 1) = columnVal0
WKS.Cells(inV, 2) = columnVal1
If columnVal2 <> "." Then WKS.Cells(inV, 3) = columnVal2
If columnVal3 <> "." Then WKS.Cells(inV, 4) = columnVal3
If columnVal4 <> "." Then WKS.Cells(inV, 5) = columnVal4

WKS.Range("A" & inV & ":E" & inV).Borders(xlInsideHorizontal).Weight = xlThin
WKS.Range("A" & inV & ":E" & inV).Borders(xlInsideVertical).Weight = xlThin
WKS.Range("A" & inV & ":E" & inV).Borders(xlTop).Weight = xlThin
WKS.Range("A" & inV & ":E" & inV).Borders(xlBottom).Weight = xlThin

If heading Then
    WKS.Range("A" & inV & ":E" & inV).Interior.Color = RGB(214, 220, 228)
    WKS.Range("A" & inV & ":E" & inV).Font.Size = 14
    WKS.Range("A" & inV & ":E" & inV).Font.Bold = True
    WKS.Range("A" & inV & ":E" & inV).Merge
    WKS.Range("A" & inV & ":E" & inV).Borders(xlTop).Weight = xlMedium
End If

If Title Then
    WKS.Range("A" & inV & ":E" & inV).Font.Bold = True
    WKS.Range("A" & inV & ":E" & inV).Interior.Color = RGB(242, 242, 242)
End If
If firstColBold Then
    WKS.Range("A" & inV).Font.Bold = True
    WKS.Range("A" & inV).Interior.Color = RGB(242, 242, 242)
    WKS.Range("B" & inV & ":E" & inV).Merge
    If set5Dec Then WKS.Range("B" & inV & ":E" & inV).NumberFormat = "0.00000"
End If
inV = inV + 1

Exit Function
err_handler:
    Call handleError("wdbProjectE", "aifInsert", Err.DESCRIPTION, Err.number)
End Function

Function loadPlannerECO(partNumber As String) As String
On Error Resume Next
loadPlannerECO = ""

Dim revID
revID = idNAM(partNumber, "NAM")
If revID = "" Then Exit Function

Dim rs1 As Recordset
Set rs1 = CurrentDb().OpenRecordset("SELECT [CHANGE_NOTICE] from ENG_ENG_REVISED_ITEMS where [REVISED_ITEM_ID] = " & revID & _
    " AND [CANCELLATION_DATE] IS NULL AND [CHANGE_NOTICE] IN (SELECT [CHANGE_NOTICE] FROM ENG_ENG_ENGINEERING_CHANGES WHERE [CHANGE_ORDER_TYPE_ID] = 6502)", dbOpenSnapshot)

If rs1.RecordCount > 0 Then loadPlannerECO = rs1!CHANGE_NOTICE

rs1.Close
Set rs1 = Nothing
End Function

Function loadTransferECO(partNumber As String) As String
On Error Resume Next
loadTransferECO = ""

Dim revID
revID = idNAM(partNumber, "NAM")
If revID = "" Then Exit Function

Dim rs1 As Recordset
Set rs1 = CurrentDb().OpenRecordset("SELECT [CHANGE_NOTICE] from ENG_ENG_REVISED_ITEMS where [REVISED_ITEM_ID] = " & revID & _
    " AND [CANCELLATION_DATE] IS NULL AND [CHANGE_NOTICE] IN (SELECT [CHANGE_NOTICE] FROM ENG_ENG_ENGINEERING_CHANGES WHERE [CHANGE_ORDER_TYPE_ID] = 72)", dbOpenSnapshot)

If rs1.RecordCount > 0 Then loadTransferECO = rs1!CHANGE_NOTICE

rs1.Close
Set rs1 = Nothing
End Function

Public Function getAttachmentsCount(stepId As Long) As Long
On Error GoTo err_handler

getAttachmentsCount = 0
Dim rs1 As Recordset
Set rs1 = CurrentDb().OpenRecordset("SELECT count(ID) as countIt from tblPartAttachmentsSP WHERE [partStepId] = " & stepId)

getAttachmentsCount = Nz(rs1!countIt, 0)

rs1.Close
Set rs1 = Nothing

err_handler:
End Function

Public Function getAttachmentsCountReq(stepId As Long, docType As Long, projectId As Long) As Long
On Error GoTo err_handler

getAttachmentsCountReq = 0
If Nz(docType) = "" Then Exit Function 'no document required

Dim rsAttStd As Recordset
Set rsAttStd = CurrentDb.OpenRecordset("SELECT uniqueFile FROM tblPartAttachmentStandards WHERE recordId = " & docType)

If rsAttStd!uniqueFile Then
    Dim rsRelated As Recordset
    Set rsRelated = CurrentDb.OpenRecordset("SELECT count(recordId) as countIt FROM tblPartProjectPartNumbers WHERE projectId = " & projectId)
    getAttachmentsCountReq = rsRelated!countIt + 1 'count of all related parts on this project + 1 for master
    rsRelated.Close
    Set rsRelated = Nothing
Else
    getAttachmentsCountReq = 1 'just one file for all the parts is OK
End If

rsAttStd.Close
Set rsAttStd = Nothing

err_handler:
End Function

Function trialScheduleEmail(Title As String, data() As Variant, columns, rows) As String
On Error GoTo err_handler

Dim tblHeading As String, tblArraySection As String, strHTMLBody As String

tblHeading = "<table style=""width: 100%; margin: 0 auto; padding: .1em; text-align: center; background-color: #fafafa;"">" & _
                            "<tbody>" & _
                                "<tr><td><h2 style=""color: #414141; font-size: 28px; margin-top: 0;"">" & Title & "</h2></td></tr>" & _
                            "</tbody>" & _
                        "</table>"
                        
Dim i As Long, titleRow, dataRows, j As Long
i = 0
tblArraySection = ""

titleRow = "<tr style=""padding: .1em;"">"
For i = 0 To columns
    titleRow = titleRow & "<th>" & data(i, 0) & "</th>"
Next i
titleRow = titleRow & "</tr>"

dataRows = ""
For j = 1 To rows
    dataRows = dataRows & "<tr style=""border-collapse: collapse; font-size: 12px; text-align: center; "">"
    For i = 0 To columns
        dataRows = dataRows & "<td style=""padding: .1em; border: 1px solid; "">" & data(i, j) & "</td>"
    Next i
    dataRows = dataRows & "</tr>"
Next j

    
tblArraySection = tblArraySection & "<table style=""width: 100%; margin: 0 auto; background: #2b2b2b; color: rgb(255,255,255);""><tbody>" & titleRow & dataRows & "</tbody></table>"

strHTMLBody = "" & _
"<!DOCTYPE html><html lang=""en"" xmlns=""http://www.w3.org/1999/xhtml"" xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:o=""urn:schemas-microsoft-com:office:office"">" & _
    "<head><meta charset=""utf-8""><title>Working DB Notification</title></head>" & _
    "<body style=""margin: 0 auto; Font-family: 'Montserrat', sans-serif; font-weight: 400; font-size: 10px; line-height: 1.8;"">" & _
        "<table style=""max-width: 600px; margin: 0 auto; text-align: center;"">" & _
            "<tbody>" & _
                "<tr><td>" & tblHeading & "</td></tr>" & _
                "<tr><td>" & tblArraySection & "</td></tr>" & _
                "<tr><td><p style=""color: rgb(192, 192, 192); text-align: center;"">This email was created by  &copy; workingDB</p></td></tr>" & _
            "</tbody>" & _
        "</table>" & _
    "</body>" & _
"</html>"

trialScheduleEmail = strHTMLBody

Exit Function
err_handler:
    Call handleError("wdbProjectE", "trialScheduleEmail", Err.DESCRIPTION, Err.number)
End Function

Public Function grabHistoryRef(dataValue As Variant, columnName As String) As String
On Error GoTo err_handler

grabHistoryRef = dataValue

If dataValue = "0" Then
    grabHistoryRef = "0 / False"
    Exit Function
ElseIf dataValue = "-1" Then
    grabHistoryRef = "True"
    Exit Function
End If

dataValue = CDbl(dataValue)

Dim rs1 As Recordset
Set rs1 = CurrentDb.OpenRecordset("SELECT " & columnName & " FROM tblDropDownsSP WHERE ID = " & dataValue)

grabHistoryRef = rs1(columnName)

err_handler:
End Function

Public Function completelyDeletePartProjectAndInfo()
On Error GoTo err_handler
'-----THIS SUB IS NOT YET USABLE

Dim db As Database, partInfoId, partNum

partNum = "26587"

Set db = CurrentDb()

'-----Part Project Data
db.Execute "delete * from tblPartProject where partNumber = '" & partNum & "'"
db.Execute "delete * from tblPartGates where partNumber = '" & partNum & "'"
db.Execute "delete * from tblPartSteps where partNumber = '" & partNum & "'"
db.Execute "delete * from tblPartTrackingApprovals where partNumber = '" & partNum & "'"
db.Execute "UPDATE tblPartAttachmentsSP SET fileStatus='deleting' where partNumber = '" & partNum & "'"

'-----Part Number based data
db.Execute "delete * from tblPartTesting where partNumber = '" & partNum & "'"
db.Execute "delete * from tblPartTeam where partNumber = '" & partNum & "'"
db.Execute "delete * from tblPartComponents where assemblyNumber = '" & partNum & "'"

'-----Part Info based data
Dim rsPartInfo As Recordset, rsPackaging As Recordset
Set rsPartInfo = db.OpenRecordset("SELECT * from tblPartInfo WHERE partNumber = '" & partNum & "'")

partInfoId = rsPartInfo!recordId
db.Execute "delete * from tblPartQuoteInfo where recordId = " & rsPartInfo!quoteInfoId
db.Execute "delete * from tblPartAssemblyInfo where recordId = " & rsPartInfo!assemblyInfoId
db.Execute "delete * from tblPartOutsourceInfo where recordId = " & rsPartInfo!outsourceInfoId

rsPartInfo.Close
Set rsPartInfo = Nothing

'-----Part Packaging and Components
Set rsPackaging = db.OpenRecordset("SELECT * from tblPartPackaging WHERE partInfoId = " & partInfoId)
Do While Not rsPackaging.EOF
    db.Execute "Delete * from tblPartPackagingComponents WHERE packagingInfoId = " & rsPackaging!recordId
    rsPackaging.MoveNext
Loop
rsPackaging.Delete
rsPackaging.Close
Set rsPackaging = Nothing

'-----Part Meetings and Attendees
Dim rsMeetings As Recordset
Set rsMeetings = db.OpenRecordset("SELECT * from tblPartMeetings where partNum = '" & partNum & "'")
Do While Not rsMeetings.EOF
    db.Execute "Delete * from tblPartMeetingAttendees WHERE meetingId = " & rsMeetings!recordId
    rsMeetings.MoveNext
Loop
rsMeetings.Close
Set rsMeetings = Nothing

'-----Part Info
db.Execute "delete * from tblPartInfo where partNumber = '" & partNum & "'"

MsgBox "All done.", vbInformation, "It is finished."

'Call registerWdbUpdates("tblPartProjects", partNum, "Part Project", partNum, "Deleted", "frmPartTrackingSettings")
Exit Function
err_handler:
    Call handleError("wdbProjectE", "completelyDeletePartProjectAndInfo", Err.DESCRIPTION, Err.number)
End Function

Public Function getApprovalsComplete(stepId As Long, partNumber As String) As Long
On Error GoTo err_handler

getApprovalsComplete = 0
Dim rs1 As Recordset
Set rs1 = CurrentDb().OpenRecordset("SELECT count(approvedOn) as appCount from tblPartTrackingApprovals WHERE [partNumber] = '" & partNumber & "' AND [tableRecordId] = " & stepId & " AND [tableName] = 'tblPartSteps'")

getApprovalsComplete = Nz(rs1!appCount, 0)

rs1.Close
Set rs1 = Nothing

err_handler:
End Function

Public Function getTotalApprovals(stepId As Long, partNumber As String) As Long
On Error GoTo err_handler

getTotalApprovals = 0
Dim rs1 As Recordset
Set rs1 = CurrentDb().OpenRecordset("SELECT count(recordId) as appCount from tblPartTrackingApprovals WHERE [partNumber] = '" & partNumber & "' AND [tableRecordId] = " & stepId & " AND [tableName] = 'tblPartSteps'")

getTotalApprovals = Nz(rs1!appCount, 0)

rs1.Close
Set rs1 = Nothing

err_handler:
End Function

Public Function recalcStepDueDates(projId As Long, oldDueDate As Date, moveBy As Long)
On Error Resume Next

Dim rsSteps As Recordset
Set rsSteps = CurrentDb().OpenRecordset("Select dueDate from tblPartSteps Where partProjectId = " & projId & " AND dueDate > #" & oldDueDate & "#")

Do While Not rsSteps.EOF
    rsSteps.Edit
    rsSteps!dueDate = addWorkdays(rsSteps!dueDate, moveBy)
    rsSteps.Update
    rsSteps.MoveNext
Loop

rsSteps.Close
Set rsSteps = Nothing

End Function

Public Function getCurrentStepDue(projId As Long) As String
On Error Resume Next

getCurrentStepDue = ""

Dim rs1 As Recordset
Set rs1 = CurrentDb().OpenRecordset("SELECT Min(dueDate) as minDue from tblPartSteps WHERE partProjectId = " & projId & " AND status <> 'Closed'")

getCurrentStepDue = Nz(rs1!minDue, "")

rs1.Close
Set rs1 = Nothing

End Function

Public Function createPartProject(projId)
On Error GoTo err_handler

Dim db As DAO.Database
Set db = CurrentDb()
Dim rsProject As Recordset, rsStepTemplate As Recordset, rsApprovalsTemplate As Recordset, rsGateTemplate As Recordset
Dim strInsert As String, strInsert1 As String
Dim projTempId As Long, pNum As String, childPnum As String, runningDate As Date

Set rsProject = db.OpenRecordset("SELECT * from tblPartProject WHERE recordId = " & projId)

projTempId = rsProject!projectTemplateId
pNum = rsProject!partNumber
childPnum = Nz(rsProject!childPartNumber, "")
runningDate = rsProject!projectStartDate

If Nz(pNum) = "" Then Exit Function 'escape possible part number null projects

db.Execute "INSERT INTO tblPartTeam(partNumber,person) VALUES ('" & pNum & "','" & Environ("username") & "')", dbFailOnError 'assign project engineer
Set rsGateTemplate = db.OpenRecordset("Select * FROM tblPartGateTemplate WHERE [projectTemplateId] = " & projTempId, dbOpenSnapshot)

'--GO THROUGH EACH GATE
Do While Not rsGateTemplate.EOF
    '--ADD THIS GATE
    db.Execute "INSERT INTO tblPartGates(projectId,partNumber,gateTitle) VALUES (" & projId & ",'" & pNum & "','" & rsGateTemplate![gateTitle] & "')", dbFailOnError
    TempVars.Add "gateId", db.OpenRecordset("SELECT @@identity")(0).Value
    
    '--ADD STEPS FOR THIS GATE
    Set rsStepTemplate = db.OpenRecordset("SELECT * from tblPartStepTemplate WHERE [gateTemplateId] = " & rsGateTemplate![recordId] & " ORDER BY indexOrder Asc", dbOpenSnapshot)
    Do While Not rsStepTemplate.EOF
        If (IsNull(rsStepTemplate![Title]) Or rsStepTemplate![Title] = "") Then GoTo nextStep
        runningDate = addWorkdays(runningDate, Nz(rsStepTemplate![duration], 1))
        strInsert = "INSERT INTO tblPartSteps" & _
            "(partNumber,partProjectId,partGateId,stepType,openedBy,status,openDate,lastUpdatedDate,lastUpdatedBy,stepActionId,documentType,responsible,indexOrder,duration,dueDate) VALUES"
        strInsert = strInsert & "('" & pNum & "'," & projId & "," & TempVars!gateId & ",'" & StrQuoteReplace(rsStepTemplate![Title]) & "','" & _
            Environ("username") & "','Not Started','" & Now() & "','" & Now() & "','" & Environ("username") & "',"
        strInsert = strInsert & Nz(rsStepTemplate![stepActionId], "NULL") & "," & Nz(rsStepTemplate![documentType], "NULL") & ",'" & _
            Nz(rsStepTemplate![responsible], "") & "'," & rsStepTemplate![indexOrder] & "," & Nz(rsStepTemplate![duration], 1) & ",'" & runningDate & "');"
        db.Execute strInsert, dbFailOnError
        
        '--ADD APPROVALS FOR THIS STEP
        If Not rsStepTemplate![approvalRequired] Then GoTo nextStep
        TempVars.Add "stepId", db.OpenRecordset("SELECT @@identity")(0).Value
        Set rsApprovalsTemplate = db.OpenRecordset("SELECT * FROM tblPartStepTemplateApprovals WHERE [stepTemplateId] = " & rsStepTemplate![recordId], dbOpenSnapshot)
        
        Do While Not rsApprovalsTemplate.EOF
            strInsert1 = "INSERT INTO tblPartTrackingApprovals(partNumber,requestedBy,requestedDate,dept,reqLevel,tableName,tableRecordId) VALUES ('" & _
                pNum & "','" & Environ("username") & "','" & Now() & "','" & _
                Nz(rsApprovalsTemplate![dept], "") & "','" & Nz(rsApprovalsTemplate![reqLevel], "") & "','tblPartSteps'," & TempVars!stepId & ");"
            CurrentDb().Execute strInsert1
            rsApprovalsTemplate.MoveNext
        Loop
nextStep:
        rsStepTemplate.MoveNext
    Loop
    db.Execute "UPDATE tblPartGates SET plannedDate = '" & runningDate & "' WHERE recordId = " & TempVars!gateId 'set the planned date as the last step due date in this gate
    rsGateTemplate.MoveNext
Loop

rsGateTemplate.Close
Set rsGateTemplate = Nothing
rsStepTemplate.Close
Set rsStepTemplate = Nothing

Exit Function
err_handler:
    Call handleError("wdbProjectE", "createPartProject", Err.DESCRIPTION, Err.number)
End Function

Public Function grabTitle(User) As String
On Error GoTo err_handler

If IsNull(User) Then
    grabTitle = ""
    Exit Function
End If

Dim rsPermissions As Recordset
Set rsPermissions = CurrentDb().OpenRecordset("SELECT * from tblPermissions where user = '" & User & "'")
grabTitle = rsPermissions!dept & " " & rsPermissions!Level

err_handler:
End Function

Public Function setProgressBarPROJECT()
On Error GoTo err_handler

Dim percent As Double, width As Long
width = 17820

Dim rsSteps As Recordset
Set rsSteps = CurrentDb().OpenRecordset("SELECT * from tblPartSteps WHERE partProjectId = " & Form_frmPartDashboard.recordId)

Dim totalSteps, closedSteps
rsSteps.MoveLast
totalSteps = rsSteps.RecordCount

rsSteps.filter = "status = 'Closed'"
Set rsSteps = rsSteps.OpenRecordset
If rsSteps.RecordCount = 0 Then
    percent = 0
    GoTo setBar
End If
rsSteps.MoveFirst
rsSteps.MoveLast
closedSteps = rsSteps.RecordCount
percent = closedSteps / totalSteps

setBar:
Call setBarColorPercent(percent, "progressBarPROJECT", width)

Exit Function
err_handler:
    Call handleError("wdbProjectE", "setProgressBarPROJECT", Err.DESCRIPTION, Err.number)
End Function

Public Function setProgressBarSTEPS(gateId As Long)
On Error GoTo err_handler

Dim percent As Double, width As Long
width = 11886
percent = 0

Dim rsSteps As Recordset
Set rsSteps = CurrentDb().OpenRecordset("SELECT * from tblPartSteps WHERE partGateId = " & gateId)

Dim totalSteps, closedSteps
rsSteps.MoveLast
totalSteps = rsSteps.RecordCount
If totalSteps = 0 Then GoTo setBar

rsSteps.filter = "status = 'Closed'"
Set rsSteps = rsSteps.OpenRecordset
If rsSteps.RecordCount = 0 Then GoTo setBar
rsSteps.MoveFirst
rsSteps.MoveLast
closedSteps = rsSteps.RecordCount
percent = closedSteps / totalSteps

setBar:
Call setBarColorPercent(percent, "progressBarSTEPS", width)

Exit Function
err_handler:
    Call handleError("wdbProjectE", "setProgressBarSTEPS", Err.DESCRIPTION, Err.number)
End Function

Function setBarColorPercent(percent As Double, controlName As String, barWidth As Long)
On Error GoTo err_handler

If percent < 0.1 Then
    Form_frmPartDashboard.Controls(controlName).width = 1
Else
    Form_frmPartDashboard.Controls(controlName).width = percent * barWidth
End If

Dim pColor
Select Case True
    Case percent < 0.25
        pColor = RGB(150, 100, 100)
    Case percent >= 0.25 And percent < 0.5
        pColor = RGB(185, 130, 100)
    Case percent >= 0.5 And percent < 0.75
        pColor = RGB(140, 150, 100)
    Case percent >= 0.75
        pColor = RGB(100, 150, 100)
End Select
Form_frmPartDashboard.Controls(controlName).BackColor = pColor
Form_frmPartDashboard.Controls(controlName & "_").BorderColor = pColor
Form_frmPartDashboard.Controls(controlName).BorderColor = pColor

Exit Function
err_handler:
    Call handleError("wdbProjectE", "setBarColorPercent", Err.DESCRIPTION, Err.number)
End Function

Function notifyPE(partNum As String, notiType As String, stepTitle As String, Optional sendAlways As Boolean = False, Optional stepAction As Boolean = False) As Boolean
On Error GoTo err_handler

notifyPE = False

Dim rsPartTeam As Recordset
Set rsPartTeam = CurrentDb().OpenRecordset("SELECT * from tblPartTeam where partNumber = '" & partNum & "'")
If rsPartTeam.RecordCount = 0 Then Exit Function

Do While Not rsPartTeam.EOF
    Dim rsPermissions As Recordset, sendTo As String
    If IsNull(rsPartTeam!person) Then GoTo nextRec
    sendTo = rsPartTeam!person
    Set rsPermissions = CurrentDb().OpenRecordset("SELECT user, userEmail from tblPermissions where user = '" & sendTo & "' AND Dept = 'Project' AND Level = 'Engineer'")
    If rsPermissions.RecordCount = 0 Then GoTo nextRec
    If sendTo = Environ("username") And Not sendAlways Then GoTo nextRec
    
    'actually send notification
    Dim body As String, closedBy As String
    If stepAction Then
        closedBy = "stepAction"
    Else
        closedBy = getFullName()
    End If
    body = emailContentGen(partNum & " Step " & notiType, "WDB Step " & notiType, "This step has been " & notiType, stepTitle, "Part Number: " & partNum, "Closed by: " & closedBy, "Closed On: " & CStr(Date))
    Call sendNotification(sendTo, 10, 2, stepTitle & " for " & partNum & " has been " & notiType, body, "Part Project", CLng(partNum))
    
nextRec:
    rsPartTeam.MoveNext
Loop

notifyPE = True

Exit Function
err_handler:
    Call handleError("wdbProjectE", "notifyPE", Err.DESCRIPTION, Err.number)
End Function

Function findDept(partNumber As String, dept As String, Optional returnMe As Boolean = False) As String
On Error GoTo err_handler

Dim rsPermissions As Recordset, permEm
Set rsPermissions = CurrentDb().OpenRecordset("SELECT user, userEmail from tblPermissions where Dept = '" & dept & "' AND Level = 'Engineer' AND user IN " & _
                                    "(SELECT person FROM tblPartTeam WHERE partNumber = '" & partNumber & "')")
If rsPermissions.RecordCount = 0 Then Exit Function

Do While Not rsPermissions.EOF
    If rsPermissions!User = Environ("username") And Not returnMe Then GoTo nextRec
    findDept = findDept & rsPermissions!User & ","
nextRec:
    rsPermissions.MoveNext
Loop
If findDept <> "" Then findDept = Left(findDept, Len(findDept) - 1)

Exit Function
err_handler:
    Call handleError("wdbProjectE", "findDept", Err.DESCRIPTION, Err.number)
End Function

Function scanSteps(partNum As String, routineName As String, Optional identifier As Variant = "notFound") As Boolean
On Error GoTo err_handler

scanSteps = False
'this scans to see if there is a step that needs to be deleted or closed per its step action requirements

Dim rsSteps As Recordset, rsStepActions As Recordset, dFilt As String, eFilt As String, db As Database
Set db = CurrentDb()
'grab all steps that match this partNum and routine name, and are not closed
dFilt = "SELECT * FROM tblPartSteps WHERE stepActionId IN (SELECT recordId FROM tblPartStepActions WHERE whenToRun = '" & routineName & "') AND status <> 'Closed'"
eFilt = ""
If partNum <> "all" Then eFilt = " AND partNumber = '" & partNum & "'"
Set rsSteps = db.OpenRecordset(dFilt & eFilt)

If rsSteps.RecordCount = 0 Then Exit Function 'no steps have actions attached!

Do While Not rsSteps.EOF
    Set rsStepActions = db.OpenRecordset("SELECT * FROM tblPartStepActions WHERE recordId = " & rsSteps!stepActionId)
    If Nz(rsStepActions!whenToRun, "") <> routineName Then GoTo nextOne 'check if this is the right time to run this actions step
    
    Dim matches, rsLookItUp As Recordset, matchingCol As String, meetsCriteria As Boolean
    matchingCol = "partNumber"
    If identifier = "notFound" Then identifier = "'" & partNum & "'"
    If routineName = "frmPartMoldingInfo_save" Then matchingCol = "recordId"
    
    'Check for types of actions based on table name
    Select Case rsStepActions!compareTable
        Case "INV_MTL_EAM_ASSET_ATTR_VALUES"
            Dim rsPI As Recordset, rsPMI As Recordset
            Set rsPI = db.OpenRecordset("SELECT moldInfoId FROM tblPartInfo WHERE partNumber = '" & rsSteps!partNumber & "'")
            If rsPI.RecordCount = 0 Then GoTo nextOne
            If Nz(rsPI!moldInfoId) = "" Then GoTo nextOne
            Set rsPMI = db.OpenRecordset("SELECT toolNumber FROM tblPartMoldingInfo WHERE recordId = " & rsPI!moldInfoId)
            identifier = "'" & rsPMI!toolNumber & "'"
            matchingCol = "SERIAL_NUMBER" 'toolnumber column in this table
            rsPI.Close
            Set rsPI = Nothing
            rsPMI.Close
            Set rsPMI = Nothing
        Case "ENG_ENG_ENGINEERING_CHANGES"
            Dim rsECOrev As Recordset 'find the transfer ECO
            If Nz(rsSteps!partNumber) = "" Then GoTo nextOne
            Dim pnId As String
            pnId = idNAM(rsSteps!partNumber, "NAM")
            If pnId = "" Then GoTo nextOne
            Set rsECOrev = db.OpenRecordset("select CHANGE_NOTICE from ENG_ENG_ENGINEERING_CHANGES " & _
                "where CHANGE_NOTICE IN (select CHANGE_NOTICE from ENG_ENG_REVISED_ITEMS where REVISED_ITEM_ID = " & pnId & " ) " & _
                "AND IMPLEMENTATION_DATE is not null AND REASON_CODE = 'TRANSFER'")
            If rsECOrev.RecordCount = 0 Then GoTo nextOne
            rsECOrev.Close
            Set rsECOrev = Nothing
            GoTo performAction 'transfer ECO found!
    End Select
    
    Set rsLookItUp = db.OpenRecordset("SELECT " & rsStepActions!compareColumn & " FROM " & rsStepActions!compareTable & " WHERE " & matchingCol & " = " & identifier)
    
    meetsCriteria = False
    If rsLookItUp.RecordCount = 0 Then GoTo nextOne
    
    If InStr(rsStepActions!compareData, ",") > 0 Then 'check for multiple values - always seen as an OR command, not AND
        'make an array of the values and check if any match
        Dim checkIf() As String, ITEM
        checkIf = Split(rsStepActions!compareData, ",")
        For Each ITEM In checkIf
            matches = CStr(Nz(rsLookItUp(rsStepActions!compareColumn), "")) = ITEM
            If matches Then meetsCriteria = True
        Next ITEM
    Else
        matches = CStr(Nz(rsLookItUp(rsStepActions!compareColumn))) = Nz(rsStepActions!compareData)
        If matches Then meetsCriteria = True
    End If
    
    If meetsCriteria <> rsStepActions!compareAction Then GoTo nextOne
    
performAction:
    Select Case rsStepActions!stepAction 'everything matched - what should be done with this step??
        Case "deleteStep" 'delete the step!
            Call registerPartUpdates("tblPartSteps", rsSteps!recordId, "Deleted - stepAction", rsSteps!stepType, "", partNum, rsSteps!stepType, "stepAction")
            rsSteps.Delete
            If CurrentProject.AllForms("frmPartDashboard").IsLoaded Then Form_sfrmPartDashboard.Requery
        Case "closeStep" 'close the step!
            Dim currentDate
            currentDate = Now()
            Call registerPartUpdates("tblPartSteps", rsSteps!recordId, "closeDate", rsSteps!closeDate, currentDate, rsSteps!partNumber, rsSteps!stepType, rsSteps!partProjectId, "stepAction")
            Call registerPartUpdates("tblPartSteps", rsSteps!recordId, "status", rsSteps!status, "Closed", rsSteps!partNumber, rsSteps!stepType, rsSteps!partProjectId, "stepAction")
            rsSteps.Edit
            rsSteps!closeDate = currentDate
            rsSteps!status = "Closed"
            rsSteps.Update
            
            If (DCount("recordId", "tblPartSteps", "[closeDate] is null AND partGateId = " & rsSteps!partGateId) = 0) Then 'if it's the last step in the gate, close the gate!
                Dim rsGate As Recordset
                Set rsGate = db.OpenRecordset("SELECT * FROM tblPartGates WHERE recordId = " & rsSteps!partGateId)
                Call registerPartUpdates("tblPartGates", rsSteps!partGateId, "actualDate", rsGate!gateDate, currentDate, rsSteps!partNumber, rsGate!gateTitle, rsSteps!partProjectId, "stepAction")
                
                rsGate.Edit
                rsGate!actualDate = currentDate
                rsGate.Update
                rsGate.Close
                Set rsGate = Nothing
                
                If CurrentProject.AllForms("frmPartDashboard").IsLoaded Then Form_sfrmPartDashboardDates.Requery
            End If
            
            Call notifyPE(rsSteps!partNumber, "Closed", rsSteps!stepType, True)
            If CurrentProject.AllForms("frmPartDashboard").IsLoaded Then Form_sfrmPartDashboard.Requery
    End Select

nextOne:
    rsSteps.MoveNext
Loop

On Error Resume Next
rsPI.Close
Set rsPI = Nothing
rsPMI.Close
Set rsPMI = Nothing
rsECOrev.Close
Set rsECOrev = Nothing
rsLookItUp.Close
Set rsLookItUp = Nothing
rsStepActions.Close
Set rsStepActions = Nothing
rsSteps.Close
Set rsSteps = Nothing

scanSteps = True

Exit Function
err_handler:
    Call handleError("wdbProjectE", "scanSteps", Err.DESCRIPTION, Err.number)
End Function

Function iHaveOpenApproval(stepId As Long)
On Error GoTo err_handler

iHaveOpenApproval = False

Dim rsPermissions As Recordset, rsApprovals As Recordset
Set rsPermissions = CurrentDb().OpenRecordset("SELECT * from tblPermissions where user = '" & Environ("username") & "'")
Set rsApprovals = CurrentDb().OpenRecordset("SELECT * from tblPartTrackingApprovals WHERE approvedOn is null AND tableName = 'tblPartSteps' AND tableRecordId = " & stepId & " AND ((dept = '" & rsPermissions!dept & "' AND reqLevel = '" & rsPermissions!Level & "') OR approver = '" & Environ("username") & "')")

If rsApprovals.RecordCount > 0 Then iHaveOpenApproval = True

rsPermissions.Close
Set rsPermissions = Nothing
rsApprovals.Close
Set rsApprovals = Nothing

Exit Function
err_handler:
    Call handleError("wdbProjectE", "iHaveOpenApproval", Err.DESCRIPTION, Err.number)
End Function

Function iAmApprover(approvalId As Long) As Boolean
On Error GoTo err_handler

iAmApprover = False

Dim rsPermissions As Recordset, rsApprovals As Recordset
Set rsPermissions = CurrentDb().OpenRecordset("SELECT * from tblPermissions where user = '" & Environ("username") & "'")
Set rsApprovals = CurrentDb().OpenRecordset("SELECT * from tblPartTrackingApprovals WHERE approvedOn is null AND recordId = " & approvalId & " AND ((dept = '" & rsPermissions!dept & "' AND reqLevel = '" & rsPermissions!Level & "') OR approver = '" & Environ("username") & "')")

If rsApprovals.RecordCount > 0 Then iAmApprover = True

rsPermissions.Close
Set rsPermissions = Nothing
rsApprovals.Close
Set rsApprovals = Nothing

Exit Function
err_handler:
    Call handleError("wdbProjectE", "iAmApprover", Err.DESCRIPTION, Err.number)
End Function

Function issueCount(partNum As String) As Long
On Error GoTo err_handler

issueCount = DCount("recordId", "tblPartIssues", "partNumber = '" & partNum & "' AND [closeDate] is null")

Exit Function
err_handler:
    Call handleError("wdbProjectE", "issueCount", Err.DESCRIPTION, Err.number)
End Function

Function emailPartInfo(partNum As String, noteTxt As String) As Boolean
On Error GoTo err_handler
emailPartInfo = False

Dim SendItems As New clsOutlookCreateItem               ' outlook class
    Dim strTo As String                                     ' email recipient
    Dim strSubject As String                                ' email subject
    
    Set SendItems = New clsOutlookCreateItem

    Dim rs2 As Recordset
    Set rs2 = CurrentDb.OpenRecordset("SELECT * FROM tblPartTeam WHERE partNumber = '" & partNum & "'", dbOpenSnapshot)
    strTo = ""

    Do While Not rs2.EOF
        If rs2!person <> Environ("username") Then strTo = strTo & getEmail(rs2!person) & "; "
        rs2.MoveNext
    Loop
    
    strSubject = partNum & " Sales Kickoff Meeting"
    
    Dim z As String, tempFold As String
    tempFold = "\\data\mdbdata\WorkingDB\_docs\Temp\" & Environ("username") & "\"
    If FolderExists(tempFold) = False Then MkDir (tempFold)
    z = tempFold & Format(Date, "YYMMDD") & "_" & partNum & "_Part_Information.pdf"
    DoCmd.OpenReport "rptPartInformation", acViewPreview, , "[partNumber]='" & partNum & "'", acHidden
    DoCmd.OutputTo acOutputReport, "rptPartInformation", acFormatPDF, z, False
    DoCmd.Close acReport, "rptPartInformation"
    
    SendItems.CreateMailItem sendTo:=strTo, _
                             subject:=strSubject, _
                             Attachments:=z
    Set SendItems = Nothing
    
Dim FSO
Set FSO = CreateObject("Scripting.FileSystemObject")
Call FSO.deleteFile(z)
    
emailPartInfo = True

Exit Function
err_handler:
    Call handleError("wdbProjectE", "emailPartInfo", Err.DESCRIPTION, Err.number)
End Function

Public Function registerPartUpdates(table As String, ID As Variant, column As String, _
    oldVal As Variant, newVal As Variant, partNumber As String, _
    Optional tag1 As String = "", Optional tag2 As Variant = "", Optional optionExtra As String = "")
On Error GoTo err_handler

Dim sqlColumns As String, sqlValues As String

If (VarType(oldVal) = vbDate) Then oldVal = Format(oldVal, "mm/dd/yyyy")
If (VarType(newVal) = vbDate) Then newVal = Format(newVal, "mm/dd/yyyy")

Dim rs1 As Recordset
Set rs1 = CurrentDb().OpenRecordset("tblPartUpdateTracking")

Dim updatedBy As String
updatedBy = Environ("username")
If optionExtra <> "" Then updatedBy = optionExtra

If Len(oldVal) > 255 Then oldVal = Left(oldVal, 255)
If Len(newVal) > 255 Then newVal = Left(newVal, 255)

With rs1
    .addNew
        !tableName = table
        !tableRecordId = ID
        !updatedBy = updatedBy
        !updatedDate = Now()
        !columnName = column
        !previousData = StrQuoteReplace(CStr(Nz(oldVal, "")))
        !newData = StrQuoteReplace(CStr(Nz(newVal, "")))
        !partNumber = partNumber
        !dataTag1 = StrQuoteReplace(tag1)
        !dataTag2 = StrQuoteReplace(tag2)
    .Update
    .Bookmark = .lastModified
End With

rs1.Close
Set rs1 = Nothing

Exit Function
err_handler:
    Call handleError("wdbProjectE", "registerPartUpdates", Err.DESCRIPTION, Err.number)
End Function

Function toolShipAuthorizationEmail(toolNumber As String, stepId As Long, shipMethod As String, partNumber As String) As Boolean
On Error GoTo err_handler

toolShipAuthorizationEmail = False

Dim rsApprovals As Recordset
Set rsApprovals = CurrentDb().OpenRecordset("Select * from tblPartTrackingApprovals WHERE tableName = 'tblPartSteps' AND tableRecordId = " & stepId)

Dim approvalsBool
approvalsBool = True
If rsApprovals.RecordCount = 0 Then
    approvalsBool = False
    GoTo noApprovals
End If

Dim arr() As Variant, i As Long
i = 0
rsApprovals.MoveLast
ReDim Preserve arr(rsApprovals.RecordCount)
rsApprovals.MoveFirst

Do While Not rsApprovals.EOF
    arr(i) = rsApprovals!approver & " - " & rsApprovals!approvedOn
    i = i + 1
    rsApprovals.MoveNext
Loop

noApprovals:
Dim toolEmail As String, subjectLine As String
subjectLine = "Tool Ship Authorization"
If approvalsBool Then
    toolEmail = generateEmailWarray("Tool Ship Authorization", toolNumber & " has been approved to ship", "Ship Method: " & shipMethod, "Approvals: ", arr)
Else
    toolEmail = generateHTML("Tool Ship Authorization", toolNumber & " has been approved to ship", "Ship Method: " & shipMethod, "Approvals: none", "", "")
End If

Dim rs2 As Recordset, strTo As String
Set rs2 = CurrentDb.OpenRecordset("SELECT * FROM tblPartTeam WHERE partNumber = '" & partNumber & "'", dbOpenSnapshot)
strTo = ""

Do While Not rs2.EOF
    If rs2!person <> Environ("username") Then strTo = strTo & getEmail(rs2!person) & "; "
    rs2.MoveNext
Loop

Dim SendItems As New clsOutlookCreateItem
Set SendItems = New clsOutlookCreateItem

SendItems.CreateMailItem sendTo:=strTo, _
                             subject:=subjectLine, _
                             htmlBody:=toolEmail
    Set SendItems = Nothing

toolShipAuthorizationEmail = True

Exit Function
err_handler:
    Call handleError("wdbProjectE", "toolShipAuthorizationEmail", Err.DESCRIPTION, Err.number)
End Function

Function emailPartApprovalNotification(stepId As Long, partNumber As String) As Boolean
On Error GoTo err_handler

emailPartApprovalNotification = False

Dim emailBody As String, subjectLine As String
subjectLine = "Part Approval Notification"
emailBody = generateHTML(subjectLine, partNumber & " has received customer approval", "Part Approved", "No extra details...", "", "")

Dim rs2 As Recordset, strTo As String
Set rs2 = CurrentDb.OpenRecordset("SELECT * FROM tblPartTeam WHERE partNumber = '" & partNumber & "'", dbOpenSnapshot)
strTo = ""

Do While Not rs2.EOF
    If rs2!person <> Environ("username") Then strTo = strTo & getEmail(rs2!person) & "; "
    rs2.MoveNext
Loop

Dim SendItems As New clsOutlookCreateItem
Set SendItems = New clsOutlookCreateItem

SendItems.CreateMailItem sendTo:=strTo, _
                             subject:=subjectLine, _
                             htmlBody:=emailBody
    Set SendItems = Nothing

emailPartApprovalNotification = True

Exit Function
err_handler:
    Call handleError("wdbProjectE", "emailPartApprovalNotification", Err.DESCRIPTION, Err.number)
End Function

Function emailApprovedCapitalPacket(stepId As Long, partNumber As String, capitalPacketNum As String) As Boolean
On Error GoTo err_handler

emailApprovedCapitalPacket = False

'find attachment link
Dim attachLink As String
attachLink = DLookup("directLink", "tblPartAttachmentsSP", "partStepId = " & stepId)

Dim emailBody As String, subjectLine As String
subjectLine = partNumber & " Capital Packet Approval"
emailBody = generateHTML(subjectLine, capitalPacketNum & " Capital Packet for " & partNumber & " is now Approved", "Capital Packet", "No extra details...", "", "", attachLink)

Dim rs2 As Recordset, strTo As String
Set rs2 = CurrentDb.OpenRecordset("SELECT * FROM tblPartTeam WHERE partNumber = '" & partNumber & "'", dbOpenSnapshot)
strTo = ""

Do While Not rs2.EOF
    If rs2!person <> Environ("username") Then strTo = strTo & rs2!person & ","
    rs2.MoveNext
Loop

strTo = Left(strTo, Len(strTo) - 1)

Call sendNotification(strTo, 9, 2, partNumber & " Capital Packet Approval", emailBody, "Part Project", CLng(partNumber), True)

emailApprovedCapitalPacket = True

Exit Function
err_handler:
    Call handleError("wdbProjectE", "emailApprovedCapitalPacket", Err.DESCRIPTION, Err.number)
End Function

Function generateEmailWarray(Title As String, subTitle As String, primaryMessage As String, detailTitle As String, arr() As Variant) As String
On Error GoTo err_handler

Dim tblHeading As String, tblFooter As String, strHTMLBody As String, extraFooter As String, detailTable As String

Dim ITEM, i
i = 0
detailTable = ""
For Each ITEM In arr()
    If i = UBound(arr) Then
        detailTable = detailTable & "<tr style=""border-collapse: collapse;""><td style=""padding: .1em 2em 1em 2em;"">" & ITEM & "</td></tr>"
    Else
        detailTable = detailTable & "<tr style=""border-collapse: collapse;""><td style=""padding: .1em 2em;"">" & ITEM & "</td></tr>"
    End If
    i = i + 1
Next ITEM

tblHeading = "<table style=""width: 100%; margin: 0 auto; padding: 2em 3em; text-align: center; background-color: #fafafa;"">" & _
                            "<tbody>" & _
                                "<tr><td><h2 style=""color: #414141; font-size: 28px; margin-top: 0;"">" & Title & "</h2></td></tr>" & _
                                "<tr><td><p style=""color: rgb(73, 73, 73);"">" & subTitle & "</p></td></tr>" & _
                                 "<tr><td><table style=""padding: 1em; text-align: center;"">" & _
                                     "<tr><td style=""padding: 1em 1.5em; background: #FF6B00; "">" & primaryMessage & "</td></tr>" & _
                                "</table></td></tr>" & _
                            "</tbody>" & _
                        "</table>"
                        
tblFooter = "<table style=""width: 100%; margin: 0 auto; background: #2b2b2b; color: rgba(255,255,255,.5);"">" & _
                        "<tbody>" & _
                            "<tr style=""border-collapse: collapse;""><td style=""padding: 1em; color: #c9c9c9;"">" & detailTitle & "</td></tr>" & _
                            detailTable & _
                        "</tbody>" & _
                    "</table>"
                    
extraFooter = "<tr><td><p style=""color: rgb(192, 192, 192); text-align: center;"">This email address is not monitored, please do not reply to this email</p></td></tr>"

strHTMLBody = "" & _
"<!DOCTYPE html><html lang=""en"" xmlns=""http://www.w3.org/1999/xhtml"" xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:o=""urn:schemas-microsoft-com:office:office"">" & _
    "<head><meta charset=""utf-8""><title>Working DB Notification</title></head>" & _
    "<body style=""margin: 0 auto; Font-family: 'Montserrat', sans-serif; font-weight: 400; font-size: 15px; line-height: 1.8;"">" & _
        "<table style=""max-width: 600px; margin: 0 auto; text-align: center; "">" & _
            "<tbody>" & _
                "<tr><td>" & tblHeading & "</td></tr>" & _
                "<tr><td>" & tblFooter & "</td></tr>" & _
                extraFooter & _
                "<tr><td><p style=""color: rgb(192, 192, 192); text-align: center;"">This email was created by  &copy; workingDB</p></td></tr>" & _
            "</tbody>" & _
        "</table>" & _
    "</body>" & _
"</html>"

generateEmailWarray = strHTMLBody

Exit Function
err_handler:
    Call handleError("wdbProjectE", "generateEmailWarray", Err.DESCRIPTION, Err.number)
End Function