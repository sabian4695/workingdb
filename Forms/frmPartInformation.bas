Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub addProgram_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmAddProgram"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub assyInfo_Click()
On Error GoTo Err_Handler

If Me.Dirty Then Me.Dirty = False

If IsNull(Me.assemblyInfoId) Then
    DoCmd.OpenForm "frmPartAssemblyInfo"
    DoCmd.GoToRecord acDataForm, "frmPartAssemblyInfo", acNewRec
Else
    DoCmd.OpenForm "frmPartAssemblyInfo", , , "recordId = " & Me.assemblyInfoId
End If

Form_frmPartAssemblyInfo.lblPartNumber.Caption = Me.partNumber
Form_frmPartAssemblyInfo.lblPartInfoId.Caption = Me.recordId

Form_sfrmPartComponents.filter = "assemblyNumber = '" & Me.partNumber & "'"
Form_sfrmPartComponents.FilterOn = True

Form_sfrmPartComponents.assemblyNumber.DefaultValue = Me.partNumber

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub btnSave_Click()
On Error GoTo Err_Handler

If Me.Dirty = True Then Me.Dirty = False

'SCAN THROUGH STEPS AND SEE IF CUSTOM ACTION IS SET UP FOR THIS FUNCTION
Call scanSteps(Me.partNumber, "frmPartInformation_save")

Form_frmPartDashboard.partDash_refresh_Click

DoCmd.CLOSE acForm, "frmPartInformation"
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub copyPI_Click()
On Error GoTo Err_Handler

If (restrict(Environ("username"), TempVars!projectOwner) = True) Then
    MsgBox "Only project/service Engineers can do this", vbCritical, "Denied"
    Exit Sub
End If

If Me.Dirty Then Me.Dirty = False

Dim copyPartNum As String
copyPartNum = InputBox("Enter part number", "Input Part Number")
If StrPtr(copyPartNum) = 0 Or copyPartNum = "" Then Exit Sub 'must enter something
Call copyPartInformation(copyPartNum, Me.partNumber, Me.name)

Me.Requery
Call snackBox("success", "Success", "Part Info Copied! Please double check ALL information for accuracy", Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Function afterUpdate_tblPartInfo()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartInfo", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)

If Me.Dirty Then Me.Dirty = False

Exit Function
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Function

Private Sub developingLocation_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartInfo", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)

Me.developingUnit.RowSource = "SELECT tblUnits.recordID, tblUnits.unitNumber, tblUnits.unitName, tblUnits.Org, tblUnits.Type, tblUnits.Description FROM tblUnits WHERE tblUnits.Org = '" & Me.developingLocation & "'"

If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Me.filter = "partNumber = '" & TempVars!partNumber & "'"
Me.FilterOn = True

Dim db As Database
Set db = CurrentDb()
Dim quoteInfo As Recordset
Dim rsPackInfo As Recordset

If Me.Recordset.RecordCount = 0 Then
    'check if there is a hidden part record just in case (due to relationship requiring quote table record, sometimes they are hidden)
    Dim rsPI As Recordset
    
    Set rsPI = db.OpenRecordset("SELECT * FROM tblPartInfo WHERE partNumber = '" & TempVars!partNumber & "'")
    If rsPI.RecordCount = 1 Then
        Set quoteInfo = db.OpenRecordset("tblPartQuoteInfo", dbOpenSnapshot)
        db.Execute "INSERT INTO tblPartQuoteInfo(quoteNumber) VALUES (0)", dbFailOnError
        TempVars.Add "quoteInfoId", db.OpenRecordset("SELECT @@identity")(0).Value
        
        rsPI.Edit
        rsPI!quoteInfoId = TempVars!quoteInfoId
        rsPI.Update
        
        quoteInfo.CLOSE
        Set quoteInfo = Nothing
        Me.Requery
    End If
    
    rsPI.CLOSE
    Set rsPI = Nothing
End If


If TempVars!partNumber <> Me.partNumber Or Nz(Me.partNumber) = "" Then
    Me.partNumber = TempVars!partNumber
    
    'set defaults
    Me.developingLocation = DLookup("permissionsLocation", "tblDropDownsSP", "recordid = " & userData("Org"))
    Me.dataStatus = 1
End If

'add primary pack info record
If Not IsNull(Me.recordId) Then
    Set rsPackInfo = db.OpenRecordset("SELECT * FROM tblPartPackagingInfo WHERE partInfoId = " & Me.recordId)
    If rsPackInfo.RecordCount = 0 Then db.Execute "INSERT INTO tblPartPackagingInfo(partInfoId,packType) VALUES (" & Me.recordId & ",1)"
    rsPackInfo.CLOSE
    Set rsPackInfo = Nothing
End If

If Nz(Me.quoteInfoId) = "" Then
    Set quoteInfo = db.OpenRecordset("tblPartQuoteInfo", dbOpenSnapshot)
    db.Execute "INSERT INTO tblPartQuoteInfo(quoteNumber) VALUES (0)", dbFailOnError
    TempVars.Add "quoteInfoId", db.OpenRecordset("SELECT @@identity")(0).Value
    Me.quoteInfoId = TempVars!quoteInfoId
    quoteInfo.CLOSE
    Set quoteInfo = Nothing
    Me.Requery
End If

Call enableDisableButtons

Dim allowEdits As Boolean, meDept As String
meDept = userData("Dept")

allowEdits = (meDept = TempVars!projectOwner And Me.dataFreeze = False) 'only project/service can edit things in this form, and only when dataFreeze is false
Me.pullSIF.Visible = allowEdits
Me.copyPI.Visible = allowEdits

If Me.dataFreeze = False Then
    Dim ctlVar As Control 'set lock value for all controls - IF not frozen
    For Each ctlVar In Me.Controls
        Select Case ctlVar.ControlType
            Case acTextBox, acCheckBox, acComboBox
                ctlVar.Locked = Not CBool(InStr(ctlVar.tag, meDept))
        End Select
    Next ctlVar
Else
    Me.allowEdits = False
End If

Dim lockDesignM As Boolean
lockDesignM = restrict(Environ("username"), "Design", "Manager")

Me.partClassCode.Locked = lockDesignM
Me.subClassCode.Locked = lockDesignM
Me.businessCode.Locked = lockDesignM
Me.focusAreaCode.Locked = lockDesignM

If Me.dataFreeze Then
    lblLock.Caption = "Data Frozen"
Else
    lblLock.Caption = "Permissions: NMQ - PPAP Dates; PE - All Other Fields"
End If

Set db = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Err_Handler

If CurrentProject.AllForms("frmPartAssemblyInfo").IsLoaded Then DoCmd.CLOSE acForm, "frmPartAssemblyInfo"
If CurrentProject.AllForms("frmPartMoldingInfo").IsLoaded Then DoCmd.CLOSE acForm, "frmPartMoldingInfo"
If CurrentProject.AllForms("frmPartOutsourceInfo").IsLoaded Then DoCmd.CLOSE acForm, "frmPartOutsourceInfo"

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Unload", Err.DESCRIPTION, Err.number)
End Sub

Private Sub freezeData_Click()
On Error GoTo Err_Handler

'what's the status right now
'if frozen, you only managers can unfreeze
'if not frozen, this will send notification to managers that it has been frozen

'or should it be frozen automatically by the approval or closing process?

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub history_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmHistory", acNormal, , "([tableName] = 'tblPartInfo' AND [partNumber] = '" & Me.partNumber & "') OR ([tableName] = 'tblPartQuoteInfo' AND [partNumber] = '" & Me.partNumber & "')"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub moldInfo_Click()
On Error GoTo Err_Handler

If Me.Dirty Then Me.Dirty = False

If IsNull(Me.moldInfoId) Then
    DoCmd.OpenForm "frmPartMoldingInfo"
    DoCmd.GoToRecord acDataForm, "frmPartMoldingInfo", acNewRec
    Form_frmPartMoldingInfo.toolNumber.DefaultValue = "'" & Me.partNumber & "T'"
Else
    DoCmd.OpenForm "frmPartMoldingInfo", , , "recordId = " & Me.moldInfoId
End If

Form_frmPartMoldingInfo.lblPartNumber.Caption = Me.partNumber
Form_frmPartMoldingInfo.lblPartInfoId.Caption = Me.recordId

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub partType_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartInfo", Me.recordId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
Call enableDisableButtons

If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub enableDisableButtons()
On Error GoTo Err_Handler

If Me.partType = 1 Or Me.partType = 4 Then 'molded OR new color
    Me.moldInfo.Visible = True
    Me.tabTabs.Pages("materialInfo").Visible = True
Else
    Me.moldInfo.Visible = False
    Me.tabTabs.Pages("materialInfo").Visible = False
End If

If Me.partType = 2 Or Me.partType = 5 Then 'assembled or subassembly
    Me.assyInfo.Visible = True
Else
    Me.assyInfo.Visible = False
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub pullSIF_Click()
On Error GoTo Err_Handler

If (restrict(Environ("username"), TempVars!projectOwner) = True) Then
    MsgBox "Only project/service Engineers can do this", vbCritical, "Denied"
    Exit Sub
End If

If Me.Dirty Then Me.Dirty = False

Dim db As Database
Dim rs1 As Recordset
Set db = CurrentDb()

If DCount("[ROW_ID]", "APPS_Q_SIF_NEW_MOLDED_PART_V", "[NIFCO_PART_NUMBER] = '" & Me.partNumber & "'") > 0 Then 'IF MOLDED
    Set rs1 = db.OpenRecordset("APPS_Q_SIF_NEW_MOLDED_PART_V", dbOpenSnapshot)
    rs1.FindLast "[NIFCO_PART_NUMBER] = '" & Me.partNumber & "'"
    Me.sampleQty = rs1!SAMPLE_QUANTITY
    Call registerPartUpdates("tblPartInfo", Me.recordId, "sampleQty", Me.sampleQty.OldValue, Me.sampleQty, Me.partNumber, Me.name)
    Me.sampleDue = rs1!SAMPLE_DUE_DATE_TO_NIFCO
    Call registerPartUpdates("tblPartInfo", Me.recordId, "sampleDue", Me.sampleDue.OldValue, Me.sampleDue, Me.partNumber, Me.name)
    Me.monthlyVolume = rs1!EST_MONTLY_VOLUME
    Call registerPartUpdates("tblPartInfo", Me.recordId, "monthlyVolume", Me.monthlyVolume.OldValue, Me.monthlyVolume, Me.partNumber, Me.name)
    If rs1!NEWCUSTOMERPARTNUM = Me.partNumber Or rs1!NEWCUSTOMERPARTNUM = 0 Then
    Else
        Me.customerPN = rs1!NEWCUSTOMERPARTNUM
        Call registerPartUpdates("tblPartInfo", Me.recordId, "customerPN", Me.customerPN.OldValue, Me.customerPN, Me.partNumber, Me.name)
    End If
    Me.partType = 1
ElseIf DCount("[ROW_ID]", "APPS_Q_SIF_NEW_ASSEMBLED_PART_V", "[NIFCO_PART_NUMBER] = '" & Me.partNumber & "'") > 0 Then 'IF ASSEMBLED
    Set rs1 = db.OpenRecordset("APPS_Q_SIF_NEW_ASSEMBLED_PART_V", dbOpenSnapshot)
    rs1.FindLast "[NIFCO_PART_NUMBER] = '" & Me.partNumber & "'"
    Me.sampleQty = rs1!SAMPLE_QTY
    Call registerPartUpdates("tblPartInfo", Me.recordId, "sampleQty", Me.sampleQty.OldValue, Me.sampleQty, Me.partNumber, Me.name)
    Me.sampleDue = rs1!SAMPLEDUEDATE
    Call registerPartUpdates("tblPartInfo", Me.recordId, "sampleDue", Me.sampleDue.OldValue, Me.sampleDue, Me.partNumber, Me.name)
    Me.monthlyVolume = rs1!EST_MONTLY_VOLUME
    Call registerPartUpdates("tblPartInfo", Me.recordId, "monthlyVolume", Me.monthlyVolume.OldValue, Me.monthlyVolume, Me.partNumber, Me.name)
    If rs1!CUSTOMER_PART_NUM = Me.partNumber Or rs1!CUSTOMER_PART_NUM = 0 Then
    Else
        Me.customerPN = rs1!CUSTOMER_PART_NUM
        Call registerPartUpdates("tblPartInfo", Me.recordId, "customerPN", Me.customerPN.OldValue, Me.customerPN, Me.partNumber, Me.name)
    End If
    Me.partType = 2
ElseIf DCount("[ROW_ID]", "APPS_Q_SIF_NEW_PURCHASING_PART_V", "[NIFCO_PART_NUMBER] = '" & Me.partNumber & "'") > 0 Then 'IF ASSEMBLED
    Set rs1 = db.OpenRecordset("APPS_Q_SIF_NEW_PURCHASING_PART_V", dbOpenSnapshot)
    rs1.FindLast "[NIFCO_PART_NUMBER] = '" & Me.partNumber & "'"
    Me.sampleQty = rs1!SAMPLE_QTY
    Call registerPartUpdates("tblPartInfo", Me.recordId, "sampleQty", Me.sampleQty.OldValue, Me.sampleQty, Me.partNumber, Me.name)
    Me.sampleDue = rs1!SAMPLEDUEDATE
    Call registerPartUpdates("tblPartInfo", Me.recordId, "sampleDue", Me.sampleDue.OldValue, Me.sampleDue, Me.partNumber, Me.name)
    Me.monthlyVolume = rs1!EST_MONTLY_VOLUME
    Call registerPartUpdates("tblPartInfo", Me.recordId, "monthlyVolume", Me.monthlyVolume.OldValue, Me.monthlyVolume, Me.partNumber, Me.name)
    If rs1!CUSTOMER_PART_NUM = Me.partNumber Or rs1!CUSTOMER_PART_NUM = 0 Then
    Else
        Me.customerPN = rs1!CUSTOMER_PART_NUM
        Call registerPartUpdates("tblPartInfo", Me.recordId, "customerPN", Me.customerPN.OldValue, Me.customerPN, Me.partNumber, Me.name)
    End If
    Me.partType = 3
Else
    MsgBox "Part not found in SIF's. Sorry!", vbCritical, "Notice"
    Exit Sub
End If

Call registerPartUpdates("tblPartQuoteInfo", Me.recordId, "quoteNumber", Me.quoteNumber, rs1!ENG_QUOTE_LOG_NUM, Me.partNumber, Me.name)
Me.quoteNumber = rs1!ENG_QUOTE_LOG_NUM

Me.sellingPrice = rs1!PIECE_PRICE
Call registerPartUpdates("tblPartInfo", Me.recordId, "sellingPrice", Me.sellingPrice.OldValue, Me.sellingPrice, Me.partNumber, Me.name)
Me.customerId = rs1!Customer_ID
Call registerPartUpdates("tblPartInfo", Me.recordId, "customerId", Me.customerId.OldValue, Me.customerId, Me.partNumber, Me.name)
Me.DESCRIPTION = rs1!PART_DESCRIPTION
Call registerPartUpdates("tblPartInfo", Me.recordId, "Description", Me.DESCRIPTION.OldValue, Me.DESCRIPTION, Me.partNumber, Me.name)
Me.SOPdate = rs1!EST_SOP_DATE
Call registerPartUpdates("tblPartInfo", Me.recordId, "SOPdate", Me.SOPdate.OldValue, Me.SOPdate, Me.partNumber, Me.name)

Call registerPartUpdates("tblPartQuoteInfo", Me.recordId, "quotedCost", Me.quotedCost, rs1!INTERNAL_PART_COST, Me.partNumber, Me.name)
Me.quotedCost = rs1!INTERNAL_PART_COST

Me.designResponsibility = rs1!DESIGN_RESPONSIBILITY
Call registerPartUpdates("tblPartInfo", Me.recordId, "designResponsibility", Me.designResponsibility.OldValue, Me.designResponsibility, Me.partNumber, Me.name)
Me.sifNum = rs1!sifNum
Me.PPAPdue = rs1!SUBMISSION_DUE_DATE
Call registerPartUpdates("tblPartInfo", Me.recordId, "PPAPdue", Me.PPAPdue.OldValue, Me.PPAPdue, Me.partNumber, Me.name)
Me.estTIA = rs1!ESTIMATED_TIA_SUB_DATE
Call registerPartUpdates("tblPartInfo", Me.recordId, "estTIA", Me.estTIA.OldValue, Me.estTIA, Me.partNumber, Me.name)

If Me.Dirty Then Me.Dirty = False

Call registerPartUpdates("tblPartInfo", Me.recordId, "Pull SIF Data", "", "SIF Data Pulled from " & rs1!sifNum, Me.partNumber, Me.name)
Call snackBox("success", "Success", "SIF Info Pulled!", Me.name)
Set db = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub quotedCost_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartQuoteInfo", Me.recordId, Me.ActiveControl.name, Nz(Me.ActiveControl.OldValue), Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False
Exit Sub

tryAgain:
Call registerPartUpdates("tblPartQuoteInfo", Me.recordId, Me.ActiveControl.name, "", Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
If Err.number = 3251 Then GoTo tryAgain
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub quotedEOAT_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartQuoteInfo", Me.recordId, Me.ActiveControl.name, Not Me.ActiveControl, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub quotedFixtures_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartQuoteInfo", Me.recordId, Me.ActiveControl.name, Not Me.ActiveControl, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub quotedGages_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartQuoteInfo", Me.recordId, Me.ActiveControl.name, Not Me.ActiveControl, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub quotedSPH_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartQuoteInfo", Me.recordId, Me.ActiveControl.name, Nz(Me.ActiveControl.OldValue), Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False
Exit Sub

tryAgain:
Call registerPartUpdates("tblPartQuoteInfo", Me.recordId, Me.ActiveControl.name, "", Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
If Err.number = 3251 Then GoTo tryAgain
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub quotedTesting_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartQuoteInfo", Me.recordId, Me.ActiveControl.name, Not Me.ActiveControl, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub quoteNumber_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartQuoteInfo", Me.recordId, Me.ActiveControl.name, Nz(Me.ActiveControl.OldValue), Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False
Exit Sub

tryAgain:
Call registerPartUpdates("tblPartQuoteInfo", Me.recordId, Me.ActiveControl.name, "", Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
If Err.number = 3251 Then GoTo tryAgain
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub sendNotes_Click()
On Error GoTo Err_Handler

If emailPartInfo(Me.partNumber, "") = False Then Err.Raise vbObjectError + 999, , "Email couldn't send..."

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub tabTabs_Change()
On Error GoTo Err_Handler

If Me.tabTabs.Value = 4 Then 'outsource only
    If Me.Dirty Then Me.Dirty = False

    If IsNull(Me.outsourceInfoId) Then
        Me.frmPartOutsourceInfo.SetFocus
        DoCmd.GoToRecord , , acNewRec
    Else
        Me.frmPartOutsourceInfo.Form.filter = "recordId = " & Me.outsourceInfoId
    End If
    
    Form_frmPartOutsourceInfo.lblPartNumber.Caption = Me.partNumber
    Form_frmPartOutsourceInfo.lblPartInfoId.Caption = Me.recordId
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
