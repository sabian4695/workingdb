Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub annealing_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartMoldingInfo", Me("tblPartMoldingInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub annealingDetails_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartInfo", Me("tblPartMoldingInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub appearance_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartMoldingInfo", Me("tblPartMoldingInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub assemblyConcerns_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartAssemblyInfo", Me("tblPartAssemblyInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub btnClass_Click()
On Error GoTo Err_Handler

Call openPath(mainFolder("catalog"))

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub btnSave_Click()
On Error GoTo Err_Handler

If Me.Dirty = True Then Me.Dirty = False

'SCAN THROUGH STEPS AND SEE IF CUSTOM ACTION IS SET UP FOR THIS FUNCTION

Dim meetingNotes, dateOfMeeting
meetingNotes = Me.meetingNotes
dateOfMeeting = Me.dateOfMeeting

If (CurrentProject.AllForms("frmPartMeetings").IsLoaded = True) Then
    Form_frmPartMeetings.Requery
End If
DoCmd.CLOSE acForm, "frmCrossFunctionalKO"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub businessCode_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartInfo", Me("tblPartInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub cavitation_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartMoldingInfo", Me("tblPartMoldingInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub checkFixture_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartInfo", Me("tblPartInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub colorCode_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartInfo", Me("tblPartMoldingInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub copyPI_Click()
On Error GoTo Err_Handler
If Me.Dirty Then Me.Dirty = False

Dim copyPartNum As String
copyPartNum = InputBox("Enter part number", "Input Part Number")
If StrPtr(copyPartNum) = 0 Or copyPartNum = "" Then Exit Sub 'must enter something
Call copyPartInformation(copyPartNum, Me.partNumber, Me.name, "Design")

Me.Requery
Call snackBox("success", "Success", "Part Info Copied! Please double check ALL information for accuracy", Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub criticalDimensions_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartInfo", Me("tblPartInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Ctl3Dweight_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartInfo", Me("tblPartMoldingInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub customerRevLevel_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartInfo", Me("tblPartInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub dateOfMeeting_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartMeetings", Me.meetId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub designLessonsLearned_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartMeetingInfo", Me.meetId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub designResponsibility_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartInfo", Me("tblPartInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Me.dfmeaReq.Visible = Nz(Me.designResponsibility, "") = "Nifco America"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub endOfArmTooling_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartMoldingInfo", Me("tblPartMoldingInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub EPorPLrestrictions_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartMoldingInfo", Me("tblPartMoldingInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub familyTool_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartMoldingInfo", Me("tblPartMoldingInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub fixedDimensions_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartInfo", Me("tblPartInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub focusAreaCode_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartInfo", Me("tblPartInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Select Case Me.partType
    Case 1 'Molded
        Me.TabCtl278.Pages("tabAssembly").Visible = False 'set assembly info tab to invisible
    Case 2 'Assembled
        Me.TabCtl278.Pages("tabToolInfo").Visible = False 'set tool info tab to invisible
        If Nz(Me.assemblyInfoId) = "" Then
            MsgBox "Note: this is an assembled part and there is no assembly info yet - which means you can't edit the assembly/purchased section currently. Please ask your PE to add assembly info.", vbInformation, "Uh oh"
            Me.assemblyConcerns.Locked = True
            Me.purchasePartConcerns.Locked = True
        End If
        
        'all items with matInfo in tag (material info)
        
        Dim ctlVar As Control 'set visibility value for all controls with the tag
        For Each ctlVar In Me.Controls
            Select Case ctlVar.ControlType
                Case acTextBox, acComboBox, acCommandButton, acLabel, acRectangle
                    If InStr(ctlVar.tag, "matInfo") Then ctlVar.Visible = False
            End Select
        Next ctlVar
        
End Select

Form_sfrmPartMeetingAttendees.lblPerson.tag = Me.partNumber

If Nz(Me.partClassCode, "") <> "" Then Me.subClassCode.RowSource = "SELECT recordId, subClassCode, subClassCodeName, subClassCodeCat From tblPartClassification WHERE subClassCode Is Not Null AND subClassCodeCat = '" & Me.partClassCode.column(3) & "'"

Dim designResp

If CurrentProject.AllForms("frmDRSdashboard").IsLoaded Then
    designResp = DLookup("drs_designresponsibility", "tblDropDownsSP", "recordid = " & Form_frmDRSdashboard.DESIGN_RESPONSIBILITY)

    If Nz(Me.designResponsibility, 0) <> designResp Then
        Me.designResponsibility = designResp
        If Me.Dirty Then Me.Dirty = False
    End If
End If

Me.dfmeaReq.Visible = Nz(Me.designResponsibility, "") = "Nifco America"

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub gateType_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartMoldingInfo", Me("tblPartMoldingInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub generalTolerance_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartInfo", Me("tblPartInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub glass_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartInfo", Me("tblPartMoldingInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub insertMold_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartMoldingInfo", Me("tblPartMoldingInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub letDown_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartInfo", Me("tblPartMoldingInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lifterCount_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartMoldingInfo", Me("tblPartMoldingInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub massTolerance_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartInfo", Me("tblPartInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub materialNumber_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartInfo", Me("tblPartMoldingInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub materialNumber1_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartInfo", Me("tblPartMoldingInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub materialSpec_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartInfo", Me("tblPartMoldingInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub materialSymbol_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartInfo", Me("tblPartMoldingInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub matingItem_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartInfo", Me("tblPartInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub meetingNotes_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartMeetings", Me.meetId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub moldflow_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartMoldingInfo", Me("tblPartMoldingInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub packagingTest_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartInfo", Me("tblPartInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub partClassCode_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartInfo", Me("tblPartInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)

Me.subClassCode.RowSource = "SELECT recordId, subClassCode, subClassCodeName, subClassCodeCat From tblPartClassification WHERE subClassCode Is Not Null AND subClassCodeCat = '" & Me.partClassCode.column(3) & "'"

Select Case Me.partClassCode.column(3)
    Case "FBU"
        Me.businessCode = 4
    Case "ADAS"
        Me.businessCode = 9
        Me.focusAreaCode = 5
    Case "FCS"
        Me.businessCode = 1
    Case "PF"
        Me.businessCode = 3
    Case "MCD"
        Me.businessCode = 2
    Case "LSC"
        Me.businessCode = 5
End Select

If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub partMarkings_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartMoldingInfo", Me("tblPartMoldingInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub partsProduced_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartMoldingInfo", Me("tblPartMoldingInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub partTesting_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartInfo", Me("tblPartInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub programId_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartInfo", Me("tblPartInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub purchasePartConcerns_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartAssemblyInfo", Me("tblPartAssemblyInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub qualityLessonsLearned_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartMeetingInfo", Me.meetId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub quotedCost_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartQuoteInfo", Me("tblPartQuoteInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub quotedEOAT_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartMoldingInfo", Me("tblPartMoldingInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub quotedFixtures_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartInfo", Me("tblPartInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub quotedMaterial_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartMoldingInfo", Me("tblPartMoldingInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub quotedMaterial1_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartMoldingInfo", Me("tblPartMoldingInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub quotedRegrind_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartInfo", Me("tblPartMoldingInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub quotedWeight_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartInfo", Me("tblPartMoldingInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub quoteNumber_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartQuoteInfo", Me("tblPartQuoteInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub regrind_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartInfo", Me("tblPartMoldingInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub sealOffConcern_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartMoldingInfo", Me("tblPartMoldingInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub sellingPrice_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartQuoteInfo", Me("tblPartQuoteInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub sendNotes_Click()
On Error GoTo Err_Handler

Dim SendItems As New clsOutlookCreateItem               ' outlook class
Dim strTo As String                                     ' email recipient
Dim strSubject As String                                ' email subject

Set SendItems = New clsOutlookCreateItem

strSubject = Me.partNumber & " Cross Functional Kickoff Meeting"

Dim cfmmId As Long
cfmmId = Nz(DLookup("recordId", "tblPartMeetings", "partNum = '" & Me.partNumber & "' AND meetingType = 1"), 0)

Dim z As String, tempFold As String
tempFold = getTempFold
If FolderExists(tempFold) = False Then MkDir (tempFold)
z = tempFold & Format(Date, "YYMMDD") & "_" & Me.partNumber & "_CFMM.pdf"
DoCmd.OpenReport "rptCrossFunctionalKO", acViewPreview, , "[meetingId]=" & cfmmId, acHidden
DoCmd.OutputTo acOutputReport, "rptCrossFunctionalKO", acFormatPDF, z, False
DoCmd.CLOSE acReport, "rptCrossFunctionalKO"

SendItems.CreateMailItem sendTo:=grabPartTeam(Me.partNumber, True, False, True, onlyEngineers:=True), _
                         subject:=strSubject, _
                         Attachments:=z
Set SendItems = Nothing

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
Call fso.deleteFile(z)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub slideCount_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartMoldingInfo", Me("tblPartMoldingInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub slideLifterTravelConcern_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartMoldingInfo", Me("tblPartMoldingInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub subClassCode_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartInfo", Me("tblPartInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub testPanel_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartInfo", Me("tblPartInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub textured_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartMoldingInfo", Me("tblPartMoldingInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub toolingLessonsLearned_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartMeetingInfo", Me.meetId, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub toolNumber_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartMoldingInfo", Me("tblPartMoldingInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub toolReason_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartMoldingInfo", Me("tblPartMoldingInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub twinShot_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartMoldingInfo", Me("tblPartMoldingInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub undercuts_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartMoldingInfo", Me("tblPartMoldingInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub unitId_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartInfo", Me("tblPartInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub wallRatioConcern_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartMoldingInfo", Me("tblPartMoldingInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub warpConcerns_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartMoldingInfo", Me("tblPartMoldingInfo.recordId"), Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, Me.partNumber, Me.name)
If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
