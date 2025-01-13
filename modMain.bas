Option Explicit

Public Type Record
    CatiaObject As Object
    ParentIndex As Long
    IsChildInstance As Boolean
    FromDB As Boolean
    Sel As String
    ID As String
    Level As String
    Amount As String
    FilePath As String
    FileName As String
    LinkTo   As String
    LinkID   As String
    partNumber As String
    InstanceName As String
    ModelDrawingID As String
    Properties() As String
End Type

Public gcurMainProperty() As String
Public gstrDesignerName As String
Public gstrNFDesigner As String
Public glstProhibitCharacter() As String

Public Sub Init()
    ReDim glstProhibitCharacter(18)
    glstProhibitCharacter(1) = "/"
    glstProhibitCharacter(2) = """"
    glstProhibitCharacter(3) = "#"
    glstProhibitCharacter(4) = "$"
    glstProhibitCharacter(5) = "@"
    glstProhibitCharacter(6) = "%"
    glstProhibitCharacter(7) = "*"
    glstProhibitCharacter(8) = "?"
    glstProhibitCharacter(9) = "\"
    glstProhibitCharacter(10) = "<"
    glstProhibitCharacter(11) = ">"
    glstProhibitCharacter(12) = "["
    glstProhibitCharacter(13) = "]"
    glstProhibitCharacter(14) = "|"
    glstProhibitCharacter(15) = ":"
    glstProhibitCharacter(16) = "'"
    glstProhibitCharacter(17) = ";"
    glstProhibitCharacter(18) = "!"
End Sub

Private Sub ManualLock()
    Call fncInitExcel
    Call modSetting.fncRead
    Call Terminate
End Sub

Public Sub subLoadModelBase(ByVal iblnLoad2dText As Boolean)
    If fncInitExcel = False Then
        Call Terminate
        Exit Sub
    End If

    If modCatia.fncInit() = False Then
        Call Terminate
        Exit Sub
    End If
    
    Dim objExcelData As CATIAPropertyTable
    Set objExcelData = modMain.fncGetProperty()

    Dim blnIsWritten As Boolean
    blnIsWritten = fncIsSheetWritten(objExcelData)
    If blnIsWritten = True Then
        If modMessage.Show("Q001") = False Then
            Call Terminate
            Exit Sub
        End If
    End If

    Dim objCatiaData As CATIAPropertyTable
    Set objCatiaData = modCatia.fncGetProperty(iblnLoad2dText)
    If objCatiaData Is Nothing Then
        Call Terminate
        Exit Sub
    End If
    
    Dim blnIsSame As Boolean: blnIsSame = False
    Call objCatiaData.fncIsSameStructure(objExcelData, blnIsSame)
    
    Dim blnOverWrite As Boolean
    If iblnLoad2dText = False And blnIsSame = True And blnIsWritten = True Then
        blnOverWrite = False
    Else
        blnOverWrite = True
    End If
    
    Call clearSheet
    
    If fncWriteExcel(objCatiaData, blnOverWrite, objExcelData) = False Then
        Call modMessage.Show("E008")
        Call Terminate
        Exit Sub
    End If

    Call ClearDeletePropertyCheckBox
    Call fncSetSelCheckBox(False)
    MsgBox "All done!", vbInformation, "Yo"
    Call Terminate
End Sub

Public Sub cmdClrSheetClick()
    If fncInitExcel = False Then
        Call Terminate
        Exit Sub
    End If
    
    Call clearSheet
    MsgBox "All done!", vbInformation, "Yo"
    Call Terminate
End Sub

Public Sub cmdNumberingClick()
    If fncInitExcel = False Then
        Call modMessage.Show("E009")
        Call Terminate
        Exit Sub
    End If
    
    Dim objExcelData As CATIAPropertyTable
    Set objExcelData = modMain.fncGetProperty()

    Dim lngSelCnt As Long
    lngSelCnt = objExcelData.fncCountSlected()
    If lngSelCnt = 0 Then
        MsgBox "Please select a row.", vbInformation, "Hmm"
        Call Terminate
        Exit Sub
    End If
    
    Dim errMsgID As String
    errMsgID = fncCheckBeforeNumbering(objExcelData)
    If errMsgID <> "" Then
        Call modMessage.Show(errMsgID)
        Call Terminate
        Exit Sub
    End If
    
    Dim blnBlank3D As Boolean
    errMsgID = fncNumbering(objExcelData, blnBlank3D)
    If errMsgID <> "" Then
        Call modMessage.Show(errMsgID)
        Call Terminate
        Exit Sub
    End If
    
    If blnBlank3D = True Then
        Call modMessage.Show("W003")
    End If
    
    Call ClearDeletePropertyCheckBox
    Call cmdOffClick
    Call fncSetSelCheckBox(False)
    MsgBox "All done!", vbInformation, "Yo"
    
    If Not objExcelData Is Nothing Then
        Set objExcelData = Nothing
    End If
    
    Call Terminate
End Sub

Public Sub cmdSetPropertyClick()
    If fncInitExcel = False Then
        Call Terminate
        Exit Sub
    End If

    Dim blnDeleteProperty As Boolean
    If fncGetDeletePropertyCheckBox(blnDeleteProperty) = False Then
        Call modMessage.Show("E017")
        cmdOffClick
        Call Terminate
        Exit Sub
    End If

    Dim blnDrawingUpdate As Boolean
    If fncGetDrawingUpdateCheckBox(blnDrawingUpdate) = False Then
        Call modMessage.Show("E017")
        cmdOffClick
        Call Terminate
        Exit Sub
    End If

    Dim objExcelData As CATIAPropertyTable
    Set objExcelData = modMain.fncGetProperty()

    Dim lngSelCnt As Long
    lngSelCnt = objExcelData.fncCountSlected()
    If lngSelCnt = 0 Then
        MsgBox "Please select a row.", vbInformation, "Hmm"
        Call Terminate
        Exit Sub
    End If

    If blnDeleteProperty = False Then
        Dim blnCheckExcelData As Boolean
        Dim strDuplicated As String
        Dim strMsgID As String
        Dim strPropertyName As String
        strMsgID = objExcelData.fncCheckBlank(strPropertyName)
        If strMsgID = "E038" Then
            MsgBox modMessage.GetMessage(strMsgID, strPropertyName), vbCritical + vbOKOnly, "Form"
            Call Terminate
            Exit Sub
        ElseIf strMsgID <> "" Then
            Call modMessage.Show(strMsgID)
            Call Terminate
            Exit Sub
        End If
        Call objExcelData.SetDefaultDesinerSection
        If 0 < objExcelData.fncCountUnknownSection() Then
            MsgBox "File_Data_Name" & " was not generated." & vbCrLf & "Because undefined Section.", vbInformation, "Hmm"
        ElseIf 0 < objExcelData.fncCountUnknownStatus Then
            MsgBox "File_Data_Name was not generated." & vbCrLf & "Because undefined Current_Status.", vbInformation, "Hmm"
        End If
        Call objExcelData.SetDummyBlank
        If objExcelData.fncReplaceProhibitCharacter() = True Then
            If modMessage.Show("Q003") = False Then
                Call Terminate
                Exit Sub
            End If
        End If
    End If

    Call objExcelData.ClearModelID

    If modCatia.fncInit() = False Then
        Call Terminate
        Exit Sub
    End If

    Dim objCatiaData As CATIAPropertyTable
    Set objCatiaData = modCatia.fncGetProperty()
    If objCatiaData Is Nothing Then
        Call Terminate
        Exit Sub
    End If

    Dim blnIsSame As Boolean
    If objCatiaData.fncIsSameStructure(objExcelData, blnIsSame) = False Then
        Call modMessage.Show("E012")
        Call Terminate
        Exit Sub
    End If
    
    If blnIsSame = False Then
        Call modMessage.Show("E010")
        Call Terminate
        Exit Sub
    End If
    
    If blnDeleteProperty = False Then
        If fncWriteExcelForUpdate(objExcelData) = False Then
            Call modMessage.Show("E008")
            Call Terminate
        End If
        If modCatia.fncSetProperty(objCatiaData, objExcelData, blnDrawingUpdate) = False Then
            Call modMessage.Show("E018")
            Call Terminate
            Exit Sub
        End If
    Else
        If modMessage.Show("Q002") = False Then
            Call Terminate
            Exit Sub
        End If
        If modCatia.fncDeleteProperty(objCatiaData) = False Then
            Call modMessage.Show("E018")
            Call Terminate
            Exit Sub
        End If
    End If
    
    Call ClearDeletePropertyCheckBox
    Call cmdOffClick
    Call fncSetSelCheckBox(False)
    MsgBox "All done!", vbInformation, "Yo"
    Call Terminate
End Sub

Public Sub cmdSetTitleBlock()
    If fncInitExcel = False Then
        Call Terminate
        Exit Sub
    End If

    Dim objExcelData As CATIAPropertyTable
    Set objExcelData = modMain.fncGetProperty()

    Dim lngSelCnt As Long
    lngSelCnt = objExcelData.fncCountSlected()
    If lngSelCnt = 0 Then
        MsgBox "Please select a row.", vbInformation, "Hmm"
        Call Terminate
        Exit Sub
    End If

    Dim blnCheckExcelData As Boolean
    Dim strDuplicated As String

    If modCatia.fncInit() = False Then
        Call cmdOffClick
        Call fncSetSelCheckBox(False)
        Call Terminate
        Exit Sub
    End If

    Dim objCatiaData As CATIAPropertyTable
    Set objCatiaData = modCatia.fncGetProperty()
    If objCatiaData Is Nothing Then
        Call cmdOffClick
        Call fncSetSelCheckBox(False)
        Call Terminate
        Exit Sub
    End If

    Dim blnIsSame As Boolean
    If objCatiaData.fncIsSameStructure(objExcelData, blnIsSame) = False Then
        Call modMessage.Show("E012")
        Call cmdOffClick
        Call fncSetSelCheckBox(False)
        Call Terminate
        Exit Sub
    End If
    
    If blnIsSame = False Then
        Call modMessage.Show("E010")
        Call cmdOffClick
        Call fncSetSelCheckBox(False)
        Call Terminate
        Exit Sub
    End If
    
    If modCatia.fncSetProperty(objCatiaData, objExcelData, False, True) = False Then
        Call modMessage.Show("E018")
        Call cmdOffClick
        Call fncSetSelCheckBox(False)
        Call Terminate
        Exit Sub
    End If

    Call cmdOffClick
    Call fncSetSelCheckBox(False)
    MsgBox "All done!", vbInformation, "Yo"
    Call Terminate
End Sub

Public Sub cmdDataMoveClick()
    If fncInitExcel = False Then
        Call Terminate
        Exit Sub
    End If

    Dim objExcelData As CATIAPropertyTable
    Set objExcelData = modMain.fncGetProperty()

    If modCatia.fncInit() = False Then
        Call Terminate
        Exit Sub
    End If

    Dim objCatiaData As CATIAPropertyTable
    Set objCatiaData = modCatia.fncGetProperty()
    If objCatiaData Is Nothing Then
        Call Terminate
        Exit Sub
    End If

    Dim blnIsSame As Boolean
    If objCatiaData.fncIsSameStructure(objExcelData, blnIsSame) = False Then
        Call modMessage.Show("E012")
        Call Terminate
        Exit Sub
    End If
    
    If blnIsSame = False Then
        Call modMessage.Show("E010")
        Call Terminate
        Exit Sub
    End If
    
    If objExcelData.fncReplaceProhibitCharacter() = True Then
        Call modMessage.Show("E032")
        Call Terminate
        Exit Sub
    End If
    
    Dim strSaveCheck As String
    strSaveCheck = modCatia.fncCheckBeforeSave(objExcelData)
    If strSaveCheck <> "" Then
        Call modMessage.Show(strSaveCheck)
        Call Terminate
        Exit Sub
    End If
    
    If FolderExists(gstr3dexCacheDir) = False Then
        Call modMessage.Show("E037")
        Call Terminate
        Exit Sub
    End If
    
    Dim strSaveDir As String
    strSaveDir = fncCreateSaveDir(objExcelData)
    If strSaveDir = "" Then
        Call modMessage.Show("E026")
        Call Terminate
        Exit Sub
    End If
    
    If modCatia.fncSaveData(objCatiaData, strSaveDir, objExcelData) = False Then
        Call modMessage.Show("E014")
        Call Terminate
        Exit Sub
    End If

    If fncWriteExcelForDataMove(objExcelData) = False Then
        Call modMessage.Show("E008")
        Call Terminate
        Exit Sub
    End If
    
    Call ClearDeletePropertyCheckBox
    Call cmdOffClick
    Call fncSetSelCheckBox(False)
    MsgBox "All done!", vbInformation, "Yo"
    Call Terminate
End Sub

Private Sub cmdOffClick()
    If fncInitExcel = False Then
        Call Terminate
        Exit Sub
    End If
    
    Dim objExcelData As CATIAPropertyTable
    Set objExcelData = modMain.fncGetProperty()

    Call Terminate
End Sub

Private Sub cmdClrModelIDClick()
    If fncInitExcel = False Then
        Call Terminate
        Exit Sub
    End If

    Dim objExcelData As CATIAPropertyTable
    Set objExcelData = modMain.fncGetProperty()
    
    Dim lngSelCnt As Long
    lngSelCnt = objExcelData.fncCountSlected()
    If lngSelCnt = 0 Then
        MsgBox "Please select a row.", vbInformation, "Hmm"
        Call Terminate
        Exit Sub
    End If
    
    If fncClrModelID(objExcelData) = False Then
        Call modMessage.Show("E008")
        Call cmdOffClick
        Call fncSetSelCheckBox(False)
        Call Terminate
        Exit Sub
    End If
    
    Call ClearDeletePropertyCheckBox
    Call cmdOffClick
    Call fncSetSelCheckBox(False)
    Call Terminate
End Sub

Private Function fncInitExcel() As Boolean
    fncInitExcel = False
    Dim strErrID As String

    Call modMain.Init

    If modSetting.fncRead() = False Then
        Call modMessage.Show("E001")
        Exit Function
    End If

    If modSetting.fncCheck() = False Then
        Call modMessage.Show("E002")
        Exit Function
    End If

    If modMain.fncReadTitle() = False Then
        Call modMessage.Show("E001")
        Exit Function
    End If
    
    If modDefineDevelopment.fncRead() = False Then
        Call modMessage.Show("E001")
        Exit Function
    End If
    
    strErrID = modDefineDevelopment.fncCheck()
    If strErrID <> "" Then
        Call modMessage.Show(strErrID)
        Exit Function
    End If
    
    If modDefineDrawing.fncRead() = False Then
        Exit Function
    End If
    
    If modDefineDrawing.fncCheck1() = False Then
        Call modMessage.Show("E016")
        Exit Function
    End If

    fncInitExcel = True
End Function

Private Function fncReadTitle() As Boolean
    fncReadTitle = False
    
    gstrDesignerName = "Designer"
    gstrNFDesigner = Nz(Form_frmPLM.In_Charge.Value, "")
    
Dim db As Database
Dim rs1 As Recordset
Set db = CurrentDb()
Set rs1 = db.OpenRecordset("tblPLM", dbOpenSnapshot)
Dim fld As DAO.Field
Dim I As Integer
Dim intCnt As Integer
Dim startIt As Boolean
I = 0
intCnt = 0
startIt = False

For Each fld In rs1.Fields
    If startIt = True Then
        intCnt = intCnt + 1
        ReDim Preserve gcurMainProperty(intCnt)
            gcurMainProperty(intCnt) = fld.name
    End If
    If fld.name = "File_Data_Name" Then
        startIt = True
    End If
Next
Set fld = Nothing
Set db = Nothing
    
    fncReadTitle = True
End Function

Public Sub Terminate()
    ReDim gcurMainProperty(0)
    gstrNFDesigner = ""
    gstrDesignerName = ""
    
    Call modCatia.Terminate
    Call modDefineDevelopment.Terminate
    Call modDefineDrawing.Terminate
    Call modSetting.Terminate
End Sub

Private Function fncIsSheetWritten(ByRef iobjExcelData As CATIAPropertyTable) As Boolean
    fncIsSheetWritten = False
    
    If iobjExcelData Is Nothing Then Exit Function
    
    Dim lngCnt As Long
    lngCnt = iobjExcelData.fncCount()
    Dim I As Long
    For I = 1 To lngCnt
    
        Dim typRecord As Record
        Call iobjExcelData.fncItem(I, typRecord)
        
        If fncIsBlankRecord(typRecord) = False Then
            fncIsSheetWritten = True
            Exit Function
        End If
    Next I
End Function

Private Sub clearSheet()
    On Error Resume Next
    
    dbExecute "DELETE FROM tblPLM"
    Form_frmPLM.Requery
    Form_sfrmPLM.Requery
End Sub

Private Function fncWriteExcel(ByRef iobjRecords As CATIAPropertyTable, _
                               Optional ByVal iblnOverWrite As Boolean = True, _
                               Optional ByRef iobjOldData As CATIAPropertyTable = Nothing) As Boolean

    fncWriteExcel = False

    Dim lngRecCnt As Long
    lngRecCnt = iobjRecords.fncCount
    
    On Error Resume Next
    Dim lngColCnt As Long
    lngColCnt = UBound(gcurMainProperty) + 12
    
    On Error GoTo 0
    Form_frmPLM.Recordset.MoveFirst

    Dim I As Long
    For I = 1 To lngRecCnt
            If I > 1 Then
                Form_frmPLM.Recordset.addNew
            End If
        Dim typRecord As Record
        If iobjRecords.fncItem(I, typRecord) = False Then
            Exit Function
        End If
        Form_frmPLM.FileName.Value = typRecord.FileName
        Form_frmPLM.fileId.Value = typRecord.ID
        Form_frmPLM.Lv.Value = typRecord.Level
        Form_frmPLM.Amount.Value = typRecord.Amount
        Form_frmPLM.Link_ID.Value = typRecord.LinkID
        Form_frmPLM.File_Path.Value = typRecord.FilePath
        Form_frmPLM.Link_To.Value = typRecord.LinkTo
        Form_frmPLM.Instance_Name.Value = typRecord.InstanceName
        Form_frmPLM.File_Part_Number.Value = typRecord.partNumber
        Form_frmPLM.Controls("ModelID/DrawingID").Value = typRecord.ModelDrawingID
        
        On Error Resume Next
        Dim lngPropCnt As Long
        lngPropCnt = UBound(typRecord.Properties)
        On Error GoTo 0
        
Dim db As Database
Dim rs1 As Recordset
Set db = CurrentDb()
Set rs1 = db.OpenRecordset("tblPLM", dbOpenSnapshot)
Dim fld As DAO.Field
Dim j As Integer
Dim startIt As Boolean
j = 0
startIt = False

    For Each fld In rs1.Fields
        If startIt = True Then
            j = j + 1
            Dim strValue As String
            If typRecord.Properties(j) = "Unset" Or _
                typRecord.Properties(j) = "999" Then
                strValue = ""
            Else
                strValue = typRecord.Properties(j)
            End If
            Form_frmPLM.Dirty = False
            
            db.Execute "UPDATE tblPLM SET [" & fld.name & "] = '" & strValue & "' WHERE [FilePath] = '" & typRecord.FilePath & "'"
        End If
        If fld.name = "File_Data_Name" Then
            startIt = True
            Form_frmPLM.Dirty = False
        End If
    Next
    Form_frmPLM.Dirty = False
Set fld = Nothing
        If I <> lngRecCnt Then
            Dim typNextRecord As Record
            If iobjRecords.fncItem(I + 1, typNextRecord) = False Then
                Exit Function
            End If
        End If
    Next I
Set db = Nothing
fncWriteExcel = True
End Function

Private Function fncWriteExcelForUpdate(ByRef iobjRecords As CATIAPropertyTable) As Boolean
    fncWriteExcelForUpdate = False
    
    Dim rs1 As Recordset
    Dim db As Database
    Set db = CurrentDb()
    Set rs1 = db.OpenRecordset("tblPLM", dbOpenSnapshot)
    
    Form_frmPLM.Recordset.MoveFirst

    Dim lngRecCnt As Long
    lngRecCnt = iobjRecords.fncCount
    
    '/ Excel
    Dim I As Long
    For I = 1 To lngRecCnt
        Dim typRecord As Record
        If I > 1 Then
            Form_frmPLM.Recordset.MoveNext
        End If
        If iobjRecords.fncItem(I, typRecord) = False Then
            Exit Function
        End If
        
        '/ FileName
        Dim objCell
        If Form_frmPLM.FileName.Value <> typRecord.FileName Then
            Form_frmPLM.FileName.Value = typRecord.FileName
        End If
        
        '/ FilePath
        If Form_frmPLM.File_Path.Value <> typRecord.FilePath Then
            Form_frmPLM.File_Path.Value = typRecord.FilePath
        End If
        
        '/ DrawLinkTo
        If Form_frmPLM.Link_To.Value <> typRecord.LinkTo Then
            Form_frmPLM.Link_To.Value = typRecord.LinkTo
        End If
        
        '/ ModelID / DrawingID
        If Form_frmPLM.Controls("ModelID/DrawingID").Value <> typRecord.ModelDrawingID Then
            Form_frmPLM.Controls("ModelID/DrawingID").Value = typRecord.ModelDrawingID
        End If
        
        On Error Resume Next
        Dim lngPropCnt As Long
        lngPropCnt = UBound(typRecord.Properties)
        On Error GoTo 0
        
        Dim j As Long
        For j = 1 To lngPropCnt
            
            '/ Input Required
            Dim strPropName As String
            strPropName = modMain.gcurMainProperty(j)
            
            Dim strReq As String
            Dim strDataType As String
            strReq = modDefineDrawing.fncGetInputRequired(strPropName)
            strDataType = modDefineDrawing.fncGetDataType(strPropName)
            
            Dim strDummyValue As String
            If strDataType = "0" Then
                strDummyValue = "Unset"
            Else
                strDummyValue = "999"
            End If
            
            Dim strValue As String
            If strReq = "0" And typRecord.Properties(j) = strDummyValue Then
                strValue = ""
            Else
                strValue = typRecord.Properties(j)
            End If
            
            If Nz(rs1(gcurMainProperty(j)), "") <> strValue Then
                
                db.Execute "UPDATE tblPLM SET [" & gcurMainProperty(j) & "] = '" & strValue & "' WHERE [ID] = " & rs1![ID]
                Form_frmPLM.Form.Dirty = False
                Form_sfrmPLM.Dirty = False
            End If
        Next j
        rs1.MoveNext
    Next I
    
    rs1.Close
    Set rs1 = Nothing
    Set db = Nothing
    Form_frmPLM.Dirty = False
    Form_sfrmPLM.Dirty = False
    fncWriteExcelForUpdate = True
End Function

Private Function fncWriteExcelForDataMove(ByRef iobjRecords As CATIAPropertyTable) As Boolean

    fncWriteExcelForDataMove = False

    Dim lngRecCnt As Long
    lngRecCnt = iobjRecords.fncCount
    
    Dim I As Long
    For I = 1 To lngRecCnt
        Dim typRecord As Record
        If iobjRecords.fncItem(I, typRecord) = False Then Exit Function
    Next I
    
    fncWriteExcelForDataMove = True
End Function

Public Function fncGetSelCheckBox(ByRef oblnCheckBox) As Boolean
    fncGetSelCheckBox = False
    fncGetSelCheckBox = fncGetCheckBox("Main", "chkSel", oblnCheckBox)
End Function

Public Function fncSetSelCheckBox(ByVal iblnChecked As Boolean) As Boolean
    fncSetSelCheckBox = False
    fncSetSelCheckBox = fncSetCheckBox("Main", "chkSel", iblnChecked)
End Function

Public Function fncGetDrawingUpdateCheckBox(ByRef oblnCheckBox) As Boolean
    fncGetDrawingUpdateCheckBox = False
    fncGetDrawingUpdateCheckBox = fncGetCheckBox("Main", "chkDrawingUpdate", oblnCheckBox)
End Function

Public Function fncGetDeletePropertyCheckBox(ByRef oblnCheckBox) As Boolean
    fncGetDeletePropertyCheckBox = False
    fncGetDeletePropertyCheckBox = fncGetCheckBox("Main", "chkDeleteProperty", oblnCheckBox)
End Function

Public Function fncGetDuplicateDesignNoCheckBox(ByRef oblnCheckBox) As Boolean
    fncGetDuplicateDesignNoCheckBox = False
    fncGetDuplicateDesignNoCheckBox = fncGetCheckBox("Main", "chkDuplicateDesignNo", oblnCheckBox)
End Function

Private Function fncGetCheckBox(ByVal istrSheetName As String, ByVal istrCheckBoxName As String, _
                                ByRef oblnCheckBox) As Boolean
fncGetCheckBox = False
oblnCheckBox = Form_frmPLM.Controls(istrCheckBoxName).Value
fncGetCheckBox = True
End Function

Private Function fncSetCheckBox(ByVal istrSheetName As String, ByVal istrCheckBoxName As String, _
                                ByVal iblnChecked As Boolean) As Boolean

fncSetCheckBox = False
If istrCheckBoxName <> "chkSel" Then
Form_frmPLM.Controls(istrCheckBoxName).Value = iblnChecked
End If
fncSetCheckBox = True
End Function

Private Sub ClearDeletePropertyCheckBox()
dbExecute "UPDATE [tblPLM] SET [Sel] = False"
End Sub

Private Function fncGetProperty() As CATIAPropertyTable
    Set fncGetProperty = New CATIAPropertyTable
    
    Dim rs1 As Recordset
    Dim db As Database
    Set db = CurrentDb()
    Set rs1 = db.OpenRecordset("tblPLM", dbOpenSnapshot)

    Dim I As Long
    For I = 1 To DCount("[ID]", "[tblPLM]")
        On Error Resume Next
        Dim lngCnt As Long
        lngCnt = UBound(modMain.gcurMainProperty)
        On Error GoTo 0
        
        Dim strProperties() As String
        ReDim strProperties(lngCnt)
        
        Dim j As Long
        For j = 1 To lngCnt
            If fncGetIndex("File_Data_Type") = j Then
                On Error Resume Next
                Select Case ""
                    Case StrConv("CATDrawing", vbUpperCase)
                        strProperties(j) = "CATDrawing"
                    Case StrConv("CATProduct", vbUpperCase)
                        strProperties(j) = "CATProduct"
                    Case StrConv("CATPart", vbUpperCase)
                        strProperties(j) = "CATPart"
                    Case Else
                        strProperties(j) = Nz(rs1(gcurMainProperty(j)), "")
                End Select
                On Error GoTo 0
            Else
                On Error Resume Next
                strProperties(j) = Nz(rs1(gcurMainProperty(j)), "")
                On Error GoTo 0
            End If
        Next j
        
        Dim typRecord As Record
        typRecord.IsChildInstance = False
        typRecord.Sel = Nz(rs1![Sel])
        typRecord.ID = Nz(rs1![File_ID])
        typRecord.Level = Nz(rs1![Lv], "")
        typRecord.Amount = Nz(rs1![Amount])
        typRecord.LinkID = Nz(rs1![Link_ID])
        typRecord.FilePath = Nz(rs1![FilePath])
        typRecord.FileName = Nz(rs1![FileName])
        typRecord.LinkTo = Nz(rs1![LinkTo])
        typRecord.partNumber = Nz(rs1![partNumber])
        typRecord.InstanceName = Nz(rs1![InstanceName])
        typRecord.ModelDrawingID = Nz(rs1![ModelID/DrawingID])
        typRecord.Properties = strProperties
        
        If fncIsBlankRecord(typRecord) = True Then
            Exit For
        Else
            Call fncGetProperty.fncAddRecord(typRecord)
        End If
        rs1.MoveNext
    Next I
    rs1.Close
    Set rs1 = Nothing
    Set db = Nothing
End Function

Public Function fncGetIndex(ByVal istrPropertyName As String) As Long
    fncGetIndex = 0
    On Error GoTo 0

    fncGetIndex = fncGetColumn(istrPropertyName) - 11 '-column L
End Function

Private Function fncGetColumn(ByVal istrPropertyName As String) As Long
Dim db As Database, rs1 As Recordset, fld As DAO.Field
Set db = CurrentDb()
Set rs1 = db.OpenRecordset("tblPLM", dbOpenSnapshot)
Dim I As Integer
I = 0

For Each fld In rs1.Fields
    If fld.name = istrPropertyName Then
        fncGetColumn = I
        Exit Function
    End If
    I = I + 1
Next
Set fld = Nothing
Set db = Nothing
End Function

Private Function fncIsBlankRecord(ByRef itypRecord As Record) As Boolean
    fncIsBlankRecord = False
    
    If itypRecord.Level <> "" Then Exit Function
    If itypRecord.FilePath <> "" Then Exit Function
    If itypRecord.FileName <> "" Then Exit Function
    If itypRecord.LinkTo <> "" Then Exit Function
    If itypRecord.partNumber <> "" Then Exit Function
    If itypRecord.InstanceName <> "" Then Exit Function
    If itypRecord.ModelDrawingID <> "" Then Exit Function
    
    On Error Resume Next
    Dim lngCnt As Long
    lngCnt = UBound(itypRecord.Properties)
    On Error GoTo 0
    
    Dim I As Long
    For I = 1 To lngCnt
        If itypRecord.Properties(I) <> "" Then Exit Function
    Next I

    fncIsBlankRecord = True
End Function

Private Function fncClrModelID(ByRef iExcelData As CATIAPropertyTable) As Boolean
    fncClrModelID = False
    
    Dim I As Long
    For I = 1 To iExcelData.fncCount
        Dim typRecord As Record
        Call iExcelData.fncItem(I, typRecord)
        
        If Trim(typRecord.Sel) = True Then
            Dim lngCol As Long
            lngCol = fncGetColumn("ModelID/DrawingID")
        End If
    Next I
    
    fncClrModelID = True
End Function

Private Function fncCheckBeforeNumbering(ByRef iobjRecords As CATIAPropertyTable) As String
    fncCheckBeforeNumbering = ""

    Dim objSheet
    On Error GoTo 0
    
    Dim blnDuplicateDesignNo As Boolean
    Call modMain.fncGetDuplicateDesignNoCheckBox(blnDuplicateDesignNo)
    If blnDuplicateDesignNo = True Then
        Dim lngRecCnt As Long
        lngRecCnt = iobjRecords.fncCount
        Dim I As Long
        
        For I = 1 To lngRecCnt
            If fncCheckNumberingRow(iobjRecords, I) = False Then GoTo CONTINUE_FNCCHECK
    
            Dim typRecord As Record
            If iobjRecords.fncItem(I, typRecord) = False Then GoTo CONTINUE_FNCCHECK
            
            Dim lngTypeIndex As Long
            lngTypeIndex = modMain.fncGetIndex("File_Data_Type")
            
            Dim strType As String
            strType = typRecord.Properties(lngTypeIndex)
            If strType <> "CATDrawing" Then GoTo CONTINUE_FNCCHECK
    
            Dim lngLinkToIndex As Long
            lngLinkToIndex = iobjRecords.fncSearchFromFilePath(typRecord.LinkTo)
            If lngLinkToIndex = 0 Then
                fncCheckBeforeNumbering = "E027"
                Exit Function
            End If
    
            Dim typLink As Record
            If iobjRecords.fncItem(lngLinkToIndex, typLink) = False Then
                fncCheckBeforeNumbering = "E027"
                Exit Function
            End If
             
            Dim lngDesignNoIndex As Long
            lngDesignNoIndex = modMain.fncGetIndex("Design_No")
            
            Dim strDesignNo As String
            strDesignNo = typLink.Properties(lngDesignNoIndex)
            If strDesignNo <> "" Then GoTo CONTINUE_FNCCHECK
            
            If Trim(typLink.Sel) <> True Then
                fncCheckBeforeNumbering = "E031"
                Exit Function
            End If
    
CONTINUE_FNCCHECK:
        Next I
    End If
End Function

Private Function fncNumbering(ByRef iobjRecords As CATIAPropertyTable, ByRef oblnBlank3D As Boolean) As String
    fncNumbering = ""
    oblnBlank3D = False
    
    On Error GoTo 0
    
    Dim curNumberedRow() As Long
    Dim lngRecCnt As Long
    lngRecCnt = iobjRecords.fncCount
    
    Dim blnIs2DNumbered As Boolean
    blnIs2DNumbered = False
    blnIs2DNumbered = iobjRecords.fncIs2DNumbered()
    If blnIs2DNumbered = True Then
        oblnBlank3D = True
        Exit Function
    End If
    
    Dim isNumbering As Boolean
    isNumbering = False
    Dim I As Long
    For I = 1 To lngRecCnt
    
        If fncCheckNumberingRow(iobjRecords, I) = False Then GoTo continue
        
        fncNumbering = fncNumberingRow(iobjRecords, I)

        If fncNumbering <> "" Then Exit Function
        
        Dim lngNumberedCnt As Long
        On Error Resume Next
        lngNumberedCnt = UBound(curNumberedRow)
        On Error GoTo 0
        ReDim Preserve curNumberedRow(lngNumberedCnt + 1) As Long
        curNumberedRow(lngNumberedCnt + 1) = I
        
        isNumbering = True
        
continue:
    Next I
    
    Dim blnDuplicateChkBox As Boolean
    blnDuplicateChkBox = False
    Call modMain.fncGetDuplicateDesignNoCheckBox(blnDuplicateChkBox)
    
    If blnDuplicateChkBox = True Then
        On Error Resume Next
        lngRecCnt = 0
        lngRecCnt = UBound(curNumberedRow)
        On Error GoTo 0
        For I = 1 To lngRecCnt
            
            fncNumbering = fncNumberingForDrawing(iobjRecords, curNumberedRow(I))
            If fncNumbering <> "" Then Exit Function
        
        Next I
    End If
    
    If oblnBlank3D = False And isNumbering = False Then
        fncNumbering = "E020"
    End If
End Function

Private Function fncCheckNumberingRow(ByRef iobjRecords As CATIAPropertyTable, ByVal I As Long) As Boolean
    fncCheckNumberingRow = False
        
    Dim typRecord As Record
    If iobjRecords.fncItem(I, typRecord) = False Then Exit Function
    
    If Trim(typRecord.Sel) <> True Then Exit Function
    
    'DesignNo
    Dim lngDesignNoIndex As Long
    lngDesignNoIndex = modMain.fncGetIndex("Design_No")
    
    Dim strDesignNo As String
    strDesignNo = typRecord.Properties(lngDesignNoIndex)
    If Trim(strDesignNo) <> "" Then Exit Function
    
    fncCheckNumberingRow = True
End Function

Private Function fncNumberingRow(ByRef iobjRecords As CATIAPropertyTable, ByVal I As Long) As String

    fncNumberingRow = ""
    
    Dim typRecord As Record
    If iobjRecords.fncItem(I, typRecord) = False Then Exit Function

    Dim strType As String
    Dim strLinkID As String
    strType = typRecord.Properties(modMain.fncGetIndex("File_Data_Type"))
    strLinkID = typRecord.LinkID
    
    Dim con As New ADODB.Connection

    On Error GoTo Error
    Dim conStr As String
    conStr = fncGetConnectString()
    con.open (conStr)
    On Error GoTo 0
    
    Dim blnDuplicateChkBox As Boolean
    blnDuplicateChkBox = False
    Call modMain.fncGetDuplicateDesignNoCheckBox(blnDuplicateChkBox)
    
    If strType = "CATDrawing" And blnDuplicateChkBox = True Then
    Else
        fncNumberingRow = fncNumberingDesignNo(con, I)
        If fncNumberingRow <> "" Then GoTo Finally
    End If
    
    GoTo Finally

Error:
    fncNumberingRow = "E021"
Finally:
    If Not con Is Nothing Then
        If con.State = adStateOpen Then con.Close
        Set con = Nothing
    End If
End Function

Private Function fncNumberingForDrawing(ByRef iobjRecords As CATIAPropertyTable, ByVal I As Long) As String

    fncNumberingForDrawing = ""

    Dim typRecord As Record
    If iobjRecords.fncItem(I, typRecord) = False Then Exit Function

    Dim strType As String
    strType = typRecord.Properties(modMain.fncGetIndex("File_Data_Type"))
    If strType <> "CATDrawing" Then Exit Function

    Dim lngColLinkTo As Long
    lngColLinkTo = fncGetColumn("LinkTo")

    Dim lngColDesignNo As Long
    lngColDesignNo = fncGetColumn("Design_No")

    Dim objCellPartDesignNo 'As Range
    Dim objCellDrawDesignNo 'As Range
    objCellDrawDesignNo.Value = objCellPartDesignNo.Value
End Function

Private Function fncGetConnectString() As String
    fncGetConnectString = ""
    
    Dim DBServ As String, DBName As String, DBUser As String, DBPass As String
    DBServ = modSetting.gstrServerName
    DBName = modSetting.gstrDBName
    DBUser = modSetting.gstrUserName
    DBPass = modSetting.gstrPassword
    
    fncGetConnectString = "Provider=Sqloledb;" & _
                            "Data Source=" & DBServ & ";" & _
                            "Initial Catalog=" & DBName & ";" & _
                            "Connect Timeout=10;" & _
                            "user id=" & DBUser & ";" & _
                            "password=" & DBPass
End Function

Private Function fncGetOldConnectString() As String
    fncGetOldConnectString = ""
    
    Dim DBServ As String, DBName As String, DBUser As String, DBPass As String
    DBServ = modSetting.gstrOldServerName
    DBName = modSetting.gstrOldDBName
    DBUser = modSetting.gstrOldUserName
    DBPass = modSetting.gstrOldPassword
    
    fncGetOldConnectString = "Provider=Sqloledb;" & _
                            "Data Source=" & DBServ & ";" & _
                            "Initial Catalog=" & DBName & ";" & _
                            "Connect Timeout=10;" & _
                            "user id=" & DBUser & ";" & _
                            "password=" & DBPass
End Function

Private Function fncNumberingDesignNo(ByRef con As ADODB.Connection, ByVal I As Long) As String
    fncNumberingDesignNo = ""
        
    Dim lRec As ADODB.Recordset
    Dim devCode As String, tblName As String, lSql As String
    devCode = modDefineDevelopment.gstrOfficeCode
    tblName = modDefineDevelopment.gstrNumberingTable

    On Error GoTo Error
    lSql = "INSERT INTO [dbo].[" & tblName & "] ([CREATEDATE],[OSUSER])" & _
                         "VALUES (GetDate(), '" & Environ("username") & "')"
    con.Execute (lSql)

    'DesignNo
    On Error GoTo Error
    Dim designNo As String
    lSql = "SELECT SCOPE_IDENTITY()"
    Set lRec = con.Execute(lSql)
    designNo = lRec.Fields(0).Value
    On Error GoTo 0

    dbExecute "UPDATE tblPLM SET Design_No = " & designNo
    
    GoTo Finally
Error:
    fncNumberingDesignNo = "E023"
Finally:
    If Not lRec Is Nothing Then
        lRec.Close
        Set lRec = Nothing
    End If
End Function

Public Function fncGetPropertyFromDB(ByRef ilstModelID() As String, _
                                     ByRef iobjTable As CATIAPropertyTable) As Boolean

    fncGetPropertyFromDB = False
    
    Dim lngCnt As Long
    On Error Resume Next
    lngCnt = UBound(ilstModelID)
    On Error GoTo 0
    If lngCnt <= 0 Then
        fncGetPropertyFromDB = True
        GoTo Finally
    End If
    
    Dim con As New ADODB.Connection
    On Error GoTo Finally
    Dim conStr As String
    conStr = fncGetOldConnectString()
    con.open (conStr)
    On Error GoTo 0
    
    'ModelID
    On Error GoTo Finally
    Dim lSql As String
    lSql = "SELECT T3.ATTRNAME, T2.ATTRVALUE, T1.MODELID FROM " & modSetting.gstrOldDBName & ".dbo.CATIAMODEL T1 " & _
           "INNER JOIN " & modSetting.gstrOldDBName & ".dbo.CATIA_ATTR_VALUE T2 ON T1.MODELID = T2.MODELID " & _
           "INNER JOIN " & modSetting.gstrOldDBName & ".dbo.CATIA_ATTR_NAME T3 ON T3.ATTR_ID = T2.ATTR_ID " & _
           "WHERE T1.modelID = "
    
    Dim I As Long
    For I = 1 To lngCnt
        If I <> 1 Then lSql = lSql & " Or T1.modelID = "
        lSql = lSql & ilstModelID(I)
    Next I
    
    Dim objRecord As ADODB.Recordset
    Set objRecord = con.Execute(lSql)
    On Error GoTo 0
    
    If iobjTable.fncSetPropertyFromDB(ilstModelID, objRecord) = False Then GoTo Finally
    fncGetPropertyFromDB = True
    
Finally:
    If Not con Is Nothing Then
        If con.State = adStateOpen Then con.Close
        Set con = Nothing
    End If
End Function

Public Function fncCreateSaveDir(ByRef iobjCatiaData As CATIAPropertyTable) As String
    fncCreateSaveDir = ""
    
    Dim typTopRecord As Record, lngFileName As Long, strFileName As String, strSaveDir As String, strResult As String
    
    If iobjCatiaData.fncItem(1, typTopRecord) = False Then Exit Function
    lngFileName = modMain.fncGetIndex("File_Data_Name") + 1
    strFileName = typTopRecord.Properties(lngFileName)
    If strFileName = "" Then Exit Function
    strSaveDir = modSetting.gstrSendToPath & "\" & strFileName
    MkDir (strSaveDir)
    
    On Error Resume Next
    strResult = Dir(strSaveDir, vbDirectory)
    On Error GoTo 0
    
    If strResult <> "" Then fncCreateSaveDir = strSaveDir
End Function

Public Function fncGetAttrVal(ByRef iobjRecord As ADODB.Recordset, ByVal istrModelID As String, _
                              ByVal istrAttrName As String, ByRef ostrValue As String) As Boolean
    fncGetAttrVal = False
    
    On Error GoTo CATCH
    iobjRecord.MoveFirst
    Do Until iobjRecord.EOF
        On Error Resume Next
        If iobjRecord.Fields(2) = istrModelID And iobjRecord.Fields(0) = istrAttrName Then
            ostrValue = iobjRecord.Fields(1)
            fncGetAttrVal = True
            On Error GoTo 0
        End If
        iobjRecord.MoveNext
        On Error GoTo 0
    Loop
CATCH:
End Function