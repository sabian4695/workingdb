Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub applyAll_Click()
On Error GoTo Err_Handler

If Me.Dirty Then Me.Dirty = False

If DCount("[ID]", "tblPLM") < 2 Then
    MsgBox "Must have at least two files loaded to use this function. The purpose is the apply settings of one file to all files", vbInformation, "Notice"
    Exit Sub
ElseIf DCount("[ID]", "tblPLM", "[Sel] = TRUE") <> 1 Then 'make sure there's only one record selected
    MsgBox "Please select one file. Then clicking this button will apply the inputs from the selected file to the rest of the available files", vbInformation, "Notice"
    Exit Sub
End If

Dim db As Database
Dim rs1 As Recordset
Set db = CurrentDb()
Set rs1 = db.OpenRecordset("tblPLM", dbOpenDynaset, dbOpenSnapshot)
Dim fld As DAO.Field
Dim j As Integer
Dim startIt As Boolean
j = 0
startIt = False

rs1.FindFirst "[Sel] = -1"
Dim masterRec() As String
For Each fld In rs1.Fields
    ReDim Preserve masterRec(j)
    masterRec(j) = Nz(rs1(fld.name).Value, "")
    j = j + 1
Next

j = 0
For Each fld In rs1.Fields
    If startIt = True Then
        Form_frmPLM.Form.Dirty = False
        
        db.Execute "UPDATE tblPLM SET [" & fld.name & "] = '" & masterRec(j) & "'"
    End If
    If fld.name = "File_Data_Type" Then
        startIt = True
    End If
    j = j + 1
Next

Set fld = Nothing
rs1.CLOSE
Set rs1 = Nothing
Set db = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub approver_Today_Click()
On Error GoTo Err_Handler

Me.Approver_Date = Format(Date, "mm/dd/yy")
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Function resetIt()
Me.Dirty = False
Me.Requery
Me.Dirty = False
Me.refresh
End Function

Private Sub btnClose_Click()
On Error GoTo Err_Handler

DoCmd.CLOSE
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

Private Sub btnSettings_Click()
On Error GoTo Err_Handler

If privilege("Edit") Then
    openPath ("\\data\mdbdata\WorkingDB\Batch\Working DB SETTINGS.lnk")
Else
    Call snackBox("error", "No Can Do.", "You have to have Edit privilege to access this.", Me.name)
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub businessCode_AfterUpdate()
On Error GoTo Err_Handler

Call updateClassCode

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub checker_Today_Click()
On Error GoTo Err_Handler

Me.Checker_Date = Format(Date, "mm/dd/yy")
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub ckSLBblend_Click()
On Error GoTo Err_Handler

If Me.ckSLBblend = True Then
    Me.Material_Conc.Visible = True
    If Me.Color_Name = Nz(DLookup("[Material_Color]", "tblPLMdropDownsMaterialNum", "[Material_Num] = '" & Me.Material_Number & "'"), "") Then
        Me.Color_Name = ""
    End If
Else
    Me.Material_Conc.Visible = False
    Me.Material_Conc = ""
    If Me.Color_Name = "" Or IsNull(Me.Color_Name) Then
        Me.Color_Name = Nz(DLookup("[Material_Color]", "tblPLMdropDownsMaterialNum", "[Material_Num] = '" & Me.Material_Number & "'"), "")
    End If
End If
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub clearSheet_Click()
On Error GoTo Err_Handler

If Me.Recordset.RecordCount > 0 Then
    Call cmdClrSheetClick
End If

dbExecute "INSERT INTO tblPLM(Sel) VALUES(False)"
Me.Requery

Exit Sub
Err_Handler:
    Call handleError(Me.name, "clearSheet", Err.DESCRIPTION, Err.number)
End Sub

Private Sub Design_Control_No_AfterUpdate()
On Error GoTo Err_Handler

If Len(Me.Design_Control_No) = 5 Then
    Me.Design_Control_No = "0" & Me.Design_Control_No
ElseIf Len(Me.Design_Control_No) <> 6 Then
    MsgBox "Design Control No. should be 6 digits.", vbInformation, "Notice"
    Me.Design_Control_No.SetFocus
End If

Exit Sub
Err_Handler:
    Application.Echo True
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Design_No_AfterUpdate()
On Error GoTo Err_Handler

If Me.Recordset.RecordCount = 0 Then Exit Sub
If Nz(Me.Full_Design_No) = "" Or Nz(Me.Design_No) = "" Then
    Me.Left_DesignNo = ""
    Exit Sub
End If
Dim X, Y
X = Len(Me.Design_No)
Y = Len(Me.Full_Design_No)
If InStr(Me.Full_Design_No, "-") Then
    Me.Left_DesignNo = Left(Me.Full_Design_No, Y - X - 3)
Else
    Me.Left_DesignNo = Left(Me.Full_Design_No, Y - X)
End If
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub updateDistribution()
On Error GoTo Err_Handler
Dim i As Integer, j As Integer
j = 0
Me.distribution = ""
For i = 1 To 3
    If Len(Me.Controls("Distribution_" & i)) > 1 Then
        j = j + 1
        If Me.distribution = "" Then
            Me.distribution = Me.Controls("Distribution_" & i)
        Else
            Me.distribution = Me.distribution & "&" & Me.Controls("Distribution_" & i)
        End If
    End If
Next i
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub updateCustomer()
On Error GoTo Err_Handler
If Me.Direct_Customer = "" Or IsNull(Me.Direct_Customer) Then
    Me.OEM = ""
    Me.Customer = ""
    Me.OEM.Visible = False
    Exit Sub
Else
    If DLookup("[OEM]", "tblPLMdropDownsCustomer", "Customer = '" & Me.Direct_Customer.Value & "'") Then
        Me.OEM = ""
        Me.OEM.Visible = False
    Else
        Me.OEM.Visible = True
    End If
End If

If Me.OEM = "" Or IsNull(Me.OEM) Then
    Me.Customer = Me.Direct_Customer
Else
    Me.Customer = Me.Direct_Customer & " (" & Me.OEM & ")"
End If
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub updateClassCode()
On Error GoTo Err_Handler

Me.Product_Division_Code = Me.partClassCode & Me.subClassCode & "-" & Me.businessCode

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Direct_Customer_AfterUpdate()
On Error GoTo Err_Handler
Call updateCustomer
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Distribution_1_AfterUpdate()
On Error GoTo Err_Handler
Call updateDistribution
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Distribution_2_AfterUpdate()
On Error GoTo Err_Handler
Call updateDistribution
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Distribution_3_AfterUpdate()
On Error GoTo Err_Handler
Call updateDistribution
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub dwgStandards_Click()
On Error GoTo Err_Handler

Call openPath(mainFolder(Me.ActiveControl.name))
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub updateCustomerUnbound()
On Error GoTo Err_Handler
If Me.Customer = "" Or IsNull(Me.Customer) Then
    Me.OEM = ""
    Me.Direct_Customer = ""
    Me.OEM.Visible = False
    Exit Sub
End If
If InStr(Me.Customer, "(") Then
    Me.OEM.Visible = True
    Me.OEM = Split(Split(Me.Customer, "(")(1), ")")(0)
    Me.Direct_Customer = Trim(Split(Me.Customer, "(")(0))
Else
    If DLookup("[OEM]", "tblPLMdropDownsCustomer", "Customer = '" & Me.Customer & "'") Then
        Me.OEM = ""
        Me.OEM.Visible = False
    Else
        Me.OEM.Visible = True
    End If
    Me.Direct_Customer = Me.Customer
End If
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub updateClassCodesUnbound()
On Error GoTo Err_Handler

Me.partClassCode = ""
Me.subClassCode = ""
Me.businessCode = ""

Select Case Len(Me.Product_Division_Code)
    Case 5 'just class code
        Me.partClassCode = Me.Product_Division_Code
    Case 7 'class + sub class
        Me.partClassCode = Left(Me.Product_Division_Code, 5)
        Me.subClassCode = Right(Me.Product_Division_Code, 2)
    Case Is > 7 'all three
        Me.partClassCode = Left(Me.Product_Division_Code, 5)
        Me.subClassCode = Right(Split(Me.Product_Division_Code, "-")(0), 2)
        Me.businessCode = Split(Me.Product_Division_Code, "-")(1)
End Select

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub updateMaterial_SymbolUnbound()
On Error GoTo Err_Handler
If Me.Material_Symbol = "" Or IsNull(Me.Material_Symbol) Then
    Me.Material_Spec = ""
    Me.Material_Type = ""
    Me.Material_Spec.Visible = False
    Exit Sub
End If
If InStr(Me.Material_Symbol, "(") Then
    Me.Material_Spec.Visible = True
    Me.Material_Spec = Split(Split(Me.Material_Symbol, "(")(1), ")")(0)
    Me.Material_Type = Trim(Split(Me.Material_Symbol, "(")(0))
Else
    Me.Material_Spec.Visible = True
    Me.Material_Type = Me.Material_Symbol
End If
    Me.MaterialGrade.RowSource = "SELECT tblPLMdropDownsMaterialGrade.Material_Grade " & _
        "FROM tblPLMdropDownsMaterialType INNER JOIN tblPLMdropDownsMaterialGrade ON (tblPLMdropDownsMaterialType.Material_Type_ID " & _
        "= tblPLMdropDownsMaterialGrade.Material_Type_ID) AND (tblPLMdropDownsMaterialType.Material_Type_ID = " & _
        "tblPLMdropDownsMaterialGrade.Material_Type_ID) WHERE (((tblPLMdropDownsMaterialType.Material_Type) = '" & Me.Material_Type & "') And " & _
        "((tblPLMdropDownsMaterialGrade.Material_Grade) Is Not Null)) ORDER BY tblPLMdropDownsMaterialGrade.Material_Grade;"
    Me.Material_Spec.RowSource = "SELECT tblPLMdropDownsMaterialSpec.Material_Spec " & _
        "FROM tblPLMdropDownsMaterialType INNER JOIN tblPLMdropDownsMaterialSpec ON (tblPLMdropDownsMaterialType.Material_Type_ID = " & _
        "tblPLMdropDownsMaterialSpec.Material_Type_ID) AND (tblPLMdropDownsMaterialType.Material_Type_ID = " & _
        "tblPLMdropDownsMaterialSpec.Material_Type_ID) WHERE (((tblPLMdropDownsMaterialType.Material_Type) = '" & Me.Material_Type & "') And " & _
        "((tblPLMdropDownsMaterialSpec.Material_Spec) Is Not Null)) ORDER BY tblPLMdropDownsMaterialSpec.Material_Spec;"
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub updateMaterialGradeUnbound()
On Error GoTo Err_Handler
If Me.Material_Grade = "" Or IsNull(Me.Material_Grade) Then
    Me.Material_Number = ""
    Me.MaterialGrade = ""
    Me.Material_Conc = ""
    Me.Material_Number.Visible = False
    Me.Material_Conc.Visible = False
    Me.ckSLBblend.Visible = False
    Me.ckSLBblend = False
    Me.Label469.Visible = False
    Exit Sub
End If
Me.ckSLBblend.Visible = True
Me.Label469.Visible = True
If InStr(Me.Material_Grade, "(") Then
    Me.Material_Number.Visible = True
    If InStr(Me.Material_Grade, "+") Then
        Me.ckSLBblend = True
        Me.Material_Conc.Visible = True
        Me.Material_Number = Trim(Split(Split(Split(Me.Material_Grade, "(")(1), ")")(0), "+")(0))
        Me.Material_Conc = Trim(Split(Split(Split(Me.Material_Grade, "(")(1), ")")(0), "+")(1))
        Me.MaterialGrade = Trim(Split(Me.Material_Grade, "(")(0))
    Else
        Me.Material_Conc.Visible = False
        Me.ckSLBblend = False
        Me.Material_Number = Split(Split(Me.Material_Grade, "(")(1), ")")(0)
        Me.MaterialGrade = Trim(Split(Me.Material_Grade, "(")(0))
        Me.Material_Conc = ""
    End If
Else
    Me.Material_Number.Visible = True
    Me.MaterialGrade = Me.Material_Grade
    Me.Material_Conc.Visible = False
    Me.Material_Conc = ""
End If
    Me.Material_Number.RowSource = "SELECT tblPLMdropDownsMaterialNum.Material_Num FROM tblPLMdropDownsMaterialGrade " & _
        "INNER JOIN tblPLMdropDownsMaterialNum ON tblPLMdropDownsMaterialGrade.Material_Grade_ID = " & _
        "tblPLMdropDownsMaterialNum.Material_Grade_ID WHERE (((tblPLMdropDownsMaterialGrade.Material_Grade)= '" & _
        Me.MaterialGrade & "') AND ((tblPLMdropDownsMaterialNum.Material_Num) Is Not Null)) ORDER BY tblPLMdropDownsMaterialNum.Material_Num;"
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Current()
On Error GoTo Err_Handler

If Me.Recordset.RecordCount = 0 Then Exit Sub

Me.txtCF = Me.ID

Call updateCustomerUnbound
Call updateMaterial_SymbolUnbound
Call updateMaterialGradeUnbound
Call updateClassCodesUnbound

Me.subClassCode.RowSource = "SELECT recordId, subClassCode, subClassCodeName, subClassCodeCat From tblPartClassification WHERE subClassCode Is Not Null AND subClassCodeCat = '" & Me.partClassCode.column(3) & "'"

Select Case Me.partClassCode.column(3)
    Case "FBU"
        Me.businessCode = "FBU"
    Case "ADAS"
        Me.businessCode = "ADS"
        Me.focusAreaCode = "ADAS"
    Case "FCS"
        Me.businessCode = "FCS"
    Case "PF"
        Me.businessCode = "PWT"
    Case "MCD"
        Me.businessCode = "MCD"
    Case "LSC"
        Me.businessCode = "LSC"
End Select

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Dim editpriv As Boolean
editpriv = privilege("Edit")
Me.btnSettings.Visible = False
If editpriv = True Then
    Me.btnSettings.Visible = True
End If

Application.Echo False
If DCount("[ID]", "tblPLM") > 0 Then
Else
    Me.AllowAdditions = True
    DoCmd.GoToRecord , , acNewRec
    Me.Dirty = True
End If
Me.AllowAdditions = False
Call resetIt
Application.Echo True
Exit Sub
Err_Handler:
    Application.Echo True
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Sub findSIFdata()
On Error GoTo Err_Handler
Dim db As Database
Dim rs1 As Recordset
Set db = CurrentDb()
    If DCount("[ROW_ID]", "APPS_Q_SIF_NEW_MOLDED_PART_V", "[NIFCO_PART_NUMBER] = '" & Left(Me.Part_No, 5) & "'") > 0 Then
    '**NEW MOLDED**
        Set rs1 = db.OpenRecordset("APPS_Q_SIF_NEW_MOLDED_PART_V", dbOpenSnapshot)
        rs1.FindFirst "[NIFCO_PART_NUMBER] = '" & Left(Me.Part_No, 5) & "'"
        'Select Case True 'material
            'Case Left(rs1!MATERIAL_AND_COLOR, 3) = "TSM" 'contains TSM spec
                'Me.Material_Spec = UCase(Split(rs1!MATERIAL_AND_COLOR, "/")(0))
                'Me.Color_Name = UCase(Split(rs1!MATERIAL_AND_COLOR, "/")(1))
            'Case Left(rs1!MATERIAL_AND_COLOR, 6) = "PC/ABS" 'PC/ABS will thow off the next rule
                'Me.Material_Type = UCase(rs1!MATERIAL_AND_COLOR)
            'Case InStr(rs1!MATERIAL_AND_COLOR, "/") 'normal with /
                'Me.Material_Type = UCase(Split(rs1!MATERIAL_AND_COLOR, "/")(0))
                'Me.Color_Name = UCase(Split(rs1!MATERIAL_AND_COLOR, "/")(1))
            'Case Else
                'Me.Material_Symbol = UCase(rs1!MATERIAL_AND_COLOR)
        'End Select
        If rs1!NEWCUSTOMERPARTNUM = Left(Me.Part_No, 5) Or rs1!NEWCUSTOMERPARTNUM = 0 Then
        Else
            Me.Customer_Part_No = rs1!NEWCUSTOMERPARTNUM
        End If
    ElseIf DCount("[ROW_ID]", "APPS_Q_SIF_NEW_ASSEMBLED_PART_V", "[NIFCO_PART_NUMBER] = '" & Left(Me.Part_No, 5) & "'") > 0 Then
    '**NEW ASSEMBLED**
        Set rs1 = db.OpenRecordset("APPS_Q_SIF_NEW_ASSEMBLED_PART_V", dbOpenSnapshot)
        rs1.FindFirst "[NIFCO_PART_NUMBER] = '" & Left(Me.Part_No, 5) & "'"
        Me.Material_Symbol = "(NOTED)"
        Me.Material_Grade = "(NOTED)"
        Me.Color_Name = "(NOTED)"
        If rs1!CUSTOMER_PART_NUM = Left(Me.Part_No, 5) Or rs1!CUSTOMER_PART_NUM = 0 Then
        Else
            Me.Customer_Part_No = rs1!NEWCUSTOMERPARTNUM
        End If
    Else
        MsgBox "Part not found in SIF's. Sorry!", vbCritical, "Notice"
        Exit Sub
    End If
'**BOTH**
Me.Customer = UCase(Split(rs1!Customer, " ")(0))
Me.Product_Name = UCase(rs1!PART_DESCRIPTION)
If Len(rs1!ENG_QUOTE_LOG_NUM) = 4 Or Len(rs1!ENG_QUOTE_LOG_NUM) = 5 Then
    Me.RFQ_No = Left(rs1!ENG_QUOTE_LOG_NUM, 4)
End If
Select Case rs1!Critical_Part
    Case "YES-CRI"
        Me.Critical_Part = "CRITICAL"
    Case "YES"
        Me.Critical_Part = "CRITICAL"
    Case "YES-IMP"
        Me.Critical_Part = "IMPORTANT"
End Select
Select Case rs1!DESIGN_RESPONSIBILITY
    Case "Nifco America"
        Me.Submitted_80 = "30*######-80"
    Case "Nifco Japan"
        Me.Submitted_80 = "30*######-80"
End Select

Set db = Nothing
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub killDots()
On Error GoTo Err_Handler
Dim db As Database
Dim rs1 As Recordset
Set db = CurrentDb()
Set rs1 = db.OpenRecordset("tblPLM", dbOpenSnapshot)
Dim fld As DAO.Field
Dim j As Integer
Dim startIt As Boolean
j = 0
startIt = False

While Not rs1.EOF
    j = 0
    For Each fld In rs1.Fields
        If startIt = True Then
            If Nz(rs1(fld.name).Value, "") = "." Or Nz(rs1(fld.name).Value, "") = "  /  /  " Then
                Me.Dirty = False
                db.Execute "UPDATE tblPLM SET [" & fld.name & "] = '' WHERE [ID] = " & rs1!ID
            End If
        End If
        If fld.name = "File_Data_Type" Then
            startIt = True
        End If
        j = j + 1
    Next
    rs1.MoveNext
Wend
Set fld = Nothing
rs1.CLOSE
Set rs1 = Nothing
Set db = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub capitalize()
On Error GoTo Err_Handler
Dim db As Database
Dim rs1 As Recordset
Set db = CurrentDb()
Set rs1 = db.OpenRecordset("tblPLM", dbOpenSnapshot)
Dim fld As DAO.Field
Dim startIt As Boolean
startIt = False

While Not rs1.EOF
    For Each fld In rs1.Fields
        If startIt = True Then
            If fld.name = "ID" Or fld.name = "Current_Status" Or fld.name = "Classification" Or fld.name = "Process" Then
                GoTo skip
            End If
            Form_frmPLM.Form.Dirty = False
            
            db.Execute "UPDATE tblPLM SET [" & fld.name & "] = '" & UCase(rs1(fld.name).Value) & "' WHERE [ID] = " & rs1!ID
        End If
        If fld.name = "File_Data_Type" Then
            startIt = True
        End If
skip:
    Next
nextOne:
startIt = False
    rs1.MoveNext
Wend
rs1.CLOSE
Set rs1 = Nothing
Set fld = Nothing
Set db = Nothing
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub initialize_Click()
On Error GoTo Err_Handler

DoCmd.Hourglass True
Application.Echo False
Me.Painting = False

If Me.Recordset.RecordCount < 1 Then
    MsgBox "You'll want to have at least one file loaded to use this feature - trust me.", vbInformation, "Notice"
    GoTo leave
End If

If MsgBox("Is this a new drawing?", vbYesNo, "Notice") = vbYes Then
    Dim fullName As String
    fullName = UCase(getFullName())
    Me.In_Charge = fullName
    Me.Drafter = fullName
    Me.Designer = fullName
    Me.Drafter_Date = Format(Date, "mm/dd/yy")
    Me.Designer_Date = Format(Date, "mm/dd/yy")
    Me.Process = "1:New"
    If MsgBox("Would you like to add data from a SIF?", vbYesNo, "Question") = vbYes Then
        If Len(Me.Part_No) < 5 Or IsNull(Len(Me.Part_No)) Then
            MsgBox "Please enter a part number first", vbInformation, "Notice"
            GoTo leave
        End If
        Call findSIFdata
        Call Form_Current
    Else
        Me.Product_Name = UCase(Nz(DLookup("[Description]", "APPS_MTL_SYSTEM_ITEMS", "[SEGMENT1] = '" & Left(Me.Part_No, 5) & "'"), "."))
    End If
Else
    Me.Process = "2:Revise"
End If
If DCount("[Control_Number]", "dbo_tblDRS", "[Part_Number] = '" & Left(Me.Part_No, 5) & "'") > 0 Then
    Dim ctrl: ctrl = DMax("[Control_Number]", "dbo_tblDRS", "[Part_Number] = '" & Left(Me.Part_No, 5) & "'")
    If MsgBox("Design WO#" & ctrl & " was found. Do you want to use data from here?", vbYesNo, "Notice") = vbYes Then
        Me.Customer_Model_No = DLookup("[Model_Code]", "dbo_tblDRS", "Control_Number = " & ctrl)
        If Len(ctrl) = 5 Then
            Me.Design_Control_No = "0" & ctrl
        Else
            Me.Design_Control_No = ctrl
        End If
    End If
End If

Me.In_Charge = UCase(getFullName())
Me.Controls("Section") = "NAM"
Me.Current_Status = "Mass production"
Me.classification = "Internal data"

leave:
Application.Echo True
Me.Painting = True
DoCmd.Hourglass False

Exit Sub
Err_Handler:
Application.Echo True
Me.Painting = True
DoCmd.Hourglass False
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub loadParameters_Click()
On Error GoTo Err_Handler

DoCmd.Hourglass True
Application.Echo False
Me.Painting = False
Me.AllowAdditions = True

DoCmd.GoToRecord , , acNewRec

dbExecute "DELETE FROM tblPLM"
Form_frmPLM.Requery
DoCmd.GoToRecord , , acNewRec

Call subLoadModelBase(False)
Call killDots
Call resetIt

Me.AllowAdditions = False
Application.Echo True
Me.Painting = True
DoCmd.Hourglass False
If Me.Recordset.RecordCount > 0 Then
    Call Design_No_AfterUpdate
    Me.Recordset.MoveFirst
End If
Exit Sub
Err_Handler:
    Application.Echo True
    Me.Painting = True
    DoCmd.Hourglass False
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub loadTitleBlocks_Click()
On Error GoTo Err_Handler

DoCmd.Hourglass True
Application.Echo False
Me.Painting = False
Me.AllowAdditions = True

dbExecute "DELETE FROM tblPLM"
Form_frmPLM.Form.Requery
DoCmd.GoToRecord , , acNewRec
Me.Dirty = True

Call subLoadModelBase(True)
Call killDots
Call resetIt
Me.AllowAdditions = False
Application.Echo True
Me.Painting = True
DoCmd.Hourglass False
Call Design_No_AfterUpdate
Me.Recordset.MoveFirst
Exit Sub
Err_Handler:
    Application.Echo True
    Me.Painting = True
    DoCmd.Hourglass False
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub updateMaterialGrade()
On Error GoTo Err_Handler

If Me.MaterialGrade = "" Or IsNull(Me.MaterialGrade) Then
    Me.Material_Number = ""
    Me.Material_Conc = ""
    Me.Material_Grade = ""
    Me.Material_Number.Visible = False
    Me.Material_Conc.Visible = False
    Me.ckSLBblend.Visible = False
    Me.ckSLBblend = False
    Me.Label469.Visible = False
    Exit Sub
Else
    Me.Material_Number.Visible = True
End If

If Me.Material_Number = "" Or IsNull(Me.Material_Number) Then
    Me.Material_Grade = Me.MaterialGrade
    Me.Material_Conc = ""
    Me.Material_Conc.Visible = False
    Me.ckSLBblend.Visible = False
    Me.ckSLBblend = False
    Me.Label469.Visible = False
Else
    If Me.Material_Conc = "" Or IsNull(Me.Material_Conc) Then
        Me.Material_Grade = Me.MaterialGrade & " (" & Me.Material_Number & ")"
    Else
        Me.ckSLBblend.Visible = True
        Me.ckSLBblend = True
        Me.Label469.Visible = True
        Me.Material_Grade = Me.MaterialGrade & " (" & Me.Material_Number & "+" & Me.Material_Conc & ")"
    End If
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Material_Conc_AfterUpdate()
On Error GoTo Err_Handler
Call updateMaterialGrade
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Material_Number_AfterUpdate()
On Error GoTo Err_Handler
Call updateMaterialGrade
If Me.Material_Number = "" Or IsNull(Me.Material_Number) Then
    Me.Material_Conc = ""
    Me.Material_Conc.Visible = False
    Me.ckSLBblend.Visible = False
    Me.ckSLBblend = False
    Me.Label469.Visible = False
Else
    Me.ckSLBblend.Visible = True
    Me.ckSLBblend = False
    Me.Label469.Visible = True
    If Me.Color_Name = "" Or IsNull(Me.Color_Name) Then
        Me.Color_Name = Nz(DLookup("[Material_Color]", "tblPLMdropDownsMaterialNum", "[Material_Num] = '" & Me.Material_Number & "'"), "")
    End If
End If
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Material_Spec_AfterUpdate()
On Error GoTo Err_Handler
Call updateMaterialSymbol
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Material_Type_AfterUpdate()
On Error GoTo Err_Handler
Call updateMaterialSymbol

If Me.Material_Type = "" Or IsNull(Me.Material_Type) Then
Else
    Me.MaterialGrade.RowSource = "SELECT tblPLMdropDownsMaterialGrade.Material_Grade " & _
        "FROM tblPLMdropDownsMaterialType INNER JOIN tblPLMdropDownsMaterialGrade ON (tblPLMdropDownsMaterialType.Material_Type_ID " & _
        "= tblPLMdropDownsMaterialGrade.Material_Type_ID) AND (tblPLMdropDownsMaterialType.Material_Type_ID = " & _
        "tblPLMdropDownsMaterialGrade.Material_Type_ID) WHERE (((tblPLMdropDownsMaterialType.Material_Type) = '" & Me.Material_Type & "') And " & _
        "((tblPLMdropDownsMaterialGrade.Material_Grade) Is Not Null)) ORDER BY tblPLMdropDownsMaterialGrade.Material_Grade;"
    Me.Material_Spec.RowSource = "SELECT tblPLMdropDownsMaterialSpec.Material_Spec " & _
        "FROM tblPLMdropDownsMaterialType INNER JOIN tblPLMdropDownsMaterialSpec ON (tblPLMdropDownsMaterialType.Material_Type_ID = " & _
        "tblPLMdropDownsMaterialSpec.Material_Type_ID) AND (tblPLMdropDownsMaterialType.Material_Type_ID = " & _
        "tblPLMdropDownsMaterialSpec.Material_Type_ID) WHERE (((tblPLMdropDownsMaterialType.Material_Type) = '" & Me.Material_Type & "') And " & _
        "((tblPLMdropDownsMaterialSpec.Material_Spec) Is Not Null)) ORDER BY tblPLMdropDownsMaterialSpec.Material_Spec;"
End If
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub updateMaterialSymbol()
On Error GoTo Err_Handler

If Me.Material_Type = "" Or IsNull(Me.Material_Type) Then
    Me.Material_Spec = ""
    Me.Material_Symbol = ""
    Me.Material_Spec.Visible = False
    Exit Sub
Else
    Me.Material_Spec.Visible = True
End If

If Me.Material_Spec = "" Or IsNull(Me.Material_Spec) Then
    Me.Material_Symbol = Me.Material_Type
Else
    Me.Material_Symbol = Me.Material_Type & " (" & Me.Material_Spec & ")"
End If
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub MaterialGrade_AfterUpdate()
On Error GoTo Err_Handler
Call updateMaterialGrade
If Me.MaterialGrade = "" Or IsNull(Me.MaterialGrade) Then
Else
    Me.Material_Number.RowSource = "SELECT tblPLMdropDownsMaterialNum.Material_Num FROM tblPLMdropDownsMaterialGrade " & _
        "INNER JOIN tblPLMdropDownsMaterialNum ON tblPLMdropDownsMaterialGrade.Material_Grade_ID = " & _
        "tblPLMdropDownsMaterialNum.Material_Grade_ID WHERE (((tblPLMdropDownsMaterialGrade.Material_Grade)= '" & _
        Me.MaterialGrade & "') AND ((tblPLMdropDownsMaterialNum.Material_Num) Is Not Null)) ORDER BY tblPLMdropDownsMaterialNum.Material_Num;"
End If
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub moveData_Click()
On Error GoTo Err_Handler

Call cmdDataMoveClick
Call resetIt
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub OEM_AfterUpdate()
On Error GoTo Err_Handler
Call updateCustomer
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub partClassCode_AfterUpdate()
On Error GoTo Err_Handler

Me.subClassCode.RowSource = "SELECT recordId, subClassCode, subClassCodeName, subClassCodeCat From tblPartClassification WHERE subClassCode Is Not Null AND subClassCodeCat = '" & Me.partClassCode.column(3) & "'"

Select Case Me.partClassCode.column(3)
    Case "FBU"
        Me.businessCode = "FBU"
    Case "ADAS"
        Me.businessCode = "ADS"
        Me.focusAreaCode = "ADAS"
    Case "FCS"
        Me.businessCode = "FCS"
    Case "PF"
        Me.businessCode = "PWT"
    Case "MCD"
        Me.businessCode = "MCD"
    Case "LSC"
        Me.businessCode = "LSC"
End Select

Call updateClassCode

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub plmManual_Click()
On Error GoTo Err_Handler

Call openPath(mainFolder(Me.ActiveControl.name))
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub pullDesignNo_Click()
On Error GoTo Err_Handler

If Me.Dirty Then Me.Dirty = False
Call cmdNumberingClick
Call resetIt
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub refresh_Click()
On Error GoTo Err_Handler

Call resetIt
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Revision_No_AfterUpdate()
On Error GoTo Err_Handler
If Len(Me.Revision_No) = 1 Then
    Me.Revision_No = "0" & Me.Revision_No
ElseIf Len(Me.Revision_No) <> 2 Then
    MsgBox "Revision No. should be 2 digits. Please double check your entry.", vbInformation, "Notice"
End If
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub selAll_Click()
On Error GoTo Err_Handler

If Me.Dirty Then Me.Dirty = False
dbExecute "UPDATE [tblPLM] SET [Sel] = TRUE;"
Me.Requery
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub setParameter_Click()
On Error GoTo Err_Handler

DoCmd.Hourglass True
Application.Echo False
Me.Painting = False

Call resetIt
Call capitalize

If cmdSetPropertyClick Then
    '---ADD IN PUSH TO tblPartInfo---
    Dim db As Database
    Set db = CurrentDb()
    Dim rsPI As Recordset
    
    Set rsPI = db.OpenRecordset("SELECT * FROM tblPartInfo WHERE partNumber = '" & Left(Me.Part_No, 5) & "'")
    If rsPI.RecordCount = 1 Then
        'description
        If Nz(Me.Product_Name, "") <> "" And Nz(rsPI!DESCRIPTION, "") = "" Then
            rsPI.Edit
            rsPI!DESCRIPTION = Me.Product_Name
            rsPI.Update
            Call registerPartUpdates("tblPartInfo", rsPI!recordId, "description", "", rsPI!DESCRIPTION, rsPI!partNumber, "frmPLM Push")
        End If
        
        'class info
        If Nz(Me.partClassCode, 0) <> 0 And Nz(rsPI!partClassCode, 0) = 0 Then
                rsPI.Edit
                rsPI!partClassCode = Me.partClassCode.column(0)
                rsPI!subClassCode = Me.subClassCode.column(0)
                rsPI!focusAreaCode = Me.focusAreaCode.column(0)
                rsPI!businessCode = Me.businessCode.column(0)
                rsPI.Update
                Call registerPartUpdates("tblPartInfo", rsPI!recordId, "partClassCode", "", rsPI!partClassCode, rsPI!partNumber, "frmPLM Push")
                Call registerPartUpdates("tblPartInfo", rsPI!recordId, "subClassCode", "", rsPI!subClassCode, rsPI!partNumber, "frmPLM Push")
                Call registerPartUpdates("tblPartInfo", rsPI!recordId, "focusAreaCode", "", rsPI!focusAreaCode, rsPI!partNumber, "frmPLM Push")
                Call registerPartUpdates("tblPartInfo", rsPI!recordId, "businessCode", "", rsPI!businessCode, rsPI!partNumber, "frmPLM Push")
        End If
        
        'material
        If Nz(Me.Material_Number, "") <> "" And Nz(rsPI!materialNumber) = 0 Then
            rsPI.Edit
            rsPI!materialNumber = Me.Material_Number
            rsPI.Update
            Call registerPartUpdates("tblPartInfo", rsPI!recordId, "materialNumber", "", rsPI!materialNumber, rsPI!partNumber, "frmPLM Push")
        End If
        
        'concentrate
        If Nz(Me.Material_Conc, "") <> "" And Nz(rsPI!materialNumber1) = 0 Then
            rsPI.Edit
            rsPI!materialNumber1 = Me.Material_Conc
            rsPI.Update
            Call registerPartUpdates("tblPartInfo", rsPI!recordId, "materialNumber1", "", rsPI!materialNumber1, rsPI!partNumber, "frmPLM Push")
        End If
        
        'regrind
        If Nz(Me.regrind, "") <> "" And Nz(rsPI!regrind) = 0 Then
            rsPI.Edit
            rsPI!regrind = Me.regrind Like "USABLE*"
            rsPI.Update
            Call registerPartUpdates("tblPartInfo", rsPI!recordId, "regrind", "", rsPI!regrind, rsPI!partNumber, "frmPLM Push")
        End If
        
        '3D weight
        If Nz(Me.Mass, "") <> "" And Nz(rsPI!pieceWeight) = 0 Then
            rsPI.Edit
            rsPI!pieceWeight = Me.Mass
            rsPI.Update
            Call registerPartUpdates("tblPartInfo", rsPI!recordId, "pieceWeight", "", rsPI!pieceWeight, rsPI!partNumber, "frmPLM Push")
        End If
    End If
    
    rsPI.CLOSE
    Set rsPI = Nothing
    Set db = Nothing
End If

Call resetIt
Call Design_No_AfterUpdate

Application.Echo True
Me.Painting = True
DoCmd.Hourglass False

Exit Sub
Err_Handler:
Application.Echo True
Me.Painting = True
DoCmd.Hourglass False
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub setTitleBlocks_Click()
On Error GoTo Err_Handler

DoCmd.Hourglass True
Application.Echo False
Me.Painting = False

Call resetIt
Call capitalize
Call cmdSetTitleBlock
Call resetIt
Call Design_No_AfterUpdate

Application.Echo True
Me.Painting = True
DoCmd.Hourglass False

Exit Sub
Err_Handler:
Application.Echo True
Me.Painting = True
DoCmd.Hourglass False
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub subClassCode_AfterUpdate()
On Error GoTo Err_Handler

Call updateClassCode

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
