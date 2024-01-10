Option Explicit

Declare PtrSafe Function SHCreateDirectoryEx Lib "shell32" Alias "SHCreateDirectoryExA" ( _
                                                                  ByVal hwnd As Long, _
                                                                  ByVal pszPath As String, _
                                                                  ByVal psa As Long) As Long
Const DEBUG_MODE As Boolean = False

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
    Call modMessage.Show("I001")
    Call Terminate
End Sub

Public Sub cmdClrSheetClick()
    If fncInitExcel = False Then
        Call Terminate
        Exit Sub
    End If
    
    Call clearSheet
    Call modMessage.Show("I001")
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
        Call modMessage.Show("W002")
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
    Call modMessage.Show("I001")
    
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
        Call modMessage.Show("W002")
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
            Call modMessage.Show3(strMsgID, strPropertyName)
            Call Terminate
            Exit Sub
        ElseIf strMsgID <> "" Then
            Call modMessage.Show(strMsgID)
            Call Terminate
            Exit Sub
        End If
        Call objExcelData.SetDefaultDesinerSection
        If 0 < objExcelData.fncCountUnknownSection() Then
            Call modMessage.Show("I002")
        ElseIf 0 < objExcelData.fncCountUnknownStatus Then
            Call modMessage.Show("I003")
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
    Call modMessage.Show("I001")
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
        Call modMessage.Show("W002")
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
    Call modMessage.Show("I001")
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
    
    If modSetting.fncCheck3dexCacheDir() = False Then
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
    Call modMessage.Show("I001")
    Call Terminate
End Sub

Private Sub cmdAllSelClick()
    If fncInitExcel = False Then
        Call Terminate
        Exit Sub
    End If
    
    Dim objExcelData As CATIAPropertyTable
    Set objExcelData = modMain.fncGetProperty()

    Dim blnSelChecked As Boolean
    If fncGetSelCheckBox(blnSelChecked) = False Then
        Call modMessage.Show("E999")
        Call Terminate
        Exit Sub
    End If
    
    If blnSelChecked = True Then
        Call fncSetAllSel(objExcelData, True)
    Else
        Call fncSetAllSel(objExcelData, False)
    End If
End Sub

Private Sub cmdOnClick()
    If fncInitExcel = False Then
        Call Terminate
        Exit Sub
    End If

    Dim objExcelData As CATIAPropertyTable
    Set objExcelData = modMain.fncGetProperty()

    Call fncSetAllSel(objExcelData, True)
    Call Terminate
End Sub

Private Sub cmdOffClick()
    If fncInitExcel = False Then
        Call Terminate
        Exit Sub
    End If
    
    Dim objExcelData As CATIAPropertyTable
    Set objExcelData = modMain.fncGetProperty()

    Call fncSetAllSel(objExcelData, False)
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
        Call modMessage.Show("W002")
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

    Call modConst.Init

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
    gstrNFDesigner = Nz(forms!frmPLM.Form!In_Charge.Value, "")
    
Dim db As Database
Dim rs1 As Recordset
Set db = CurrentDb()
Set rs1 = db.OpenRecordset("tblPLM", dbOpenSnapshot)
Dim fld As DAO.Field
Dim i As Integer
Dim intCnt As Integer
Dim startIt As Boolean
i = 0
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
    
    If iobjExcelData Is Nothing Then
        Exit Function
    End If
    
    Dim lngCnt As Long
    lngCnt = iobjExcelData.fncCount()
    Dim i As Long
    For i = 1 To lngCnt
    
        Dim typRecord As Record
        Call iobjExcelData.fncItem(i, typRecord)
        
        If fncIsBlankRecord(typRecord) = False Then
            fncIsSheetWritten = True
            Exit Function
        End If
    Next i
End Function

Private Sub clearSheet()
    On Error Resume Next
    
    CurrentDb().Execute "DELETE FROM tblPLM"
    forms!frmPLM.Form.Requery
    forms!frmPLM!sfrmPLM.Requery
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
    forms!frmPLM.Form.Recordset.MoveFirst

    Dim i As Long
    For i = 1 To lngRecCnt
            If i > 1 Then
                forms!frmPLM.Form.Recordset.addNew
            End If
        Dim typRecord As Record
        If iobjRecords.fncItem(i, typRecord) = False Then
            Exit Function
        End If
        forms!frmPLM.Form!FileName.Value = typRecord.FileName
        forms!frmPLM.Form!FileID.Value = typRecord.ID
        forms!frmPLM.Form!Lv.Value = typRecord.Level
        forms!frmPLM.Form!Amount.Value = typRecord.Amount
        forms!frmPLM.Form!Link_ID.Value = typRecord.LinkID
        forms!frmPLM.Form!File_Path.Value = typRecord.FilePath
        forms!frmPLM.Form!Link_To.Value = typRecord.LinkTo
        forms!frmPLM.Form!Instance_Name.Value = typRecord.InstanceName
        forms!frmPLM.Form!File_Part_Number.Value = typRecord.partNumber
        forms!frmPLM!.Controls("ModelID/DrawingID").Value = typRecord.ModelDrawingID
        
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
            If typRecord.Properties(j) = VALUE_UNSET_STR Or _
                typRecord.Properties(j) = VALUE_UNSET_NUM Then
                strValue = ""
            Else
                strValue = typRecord.Properties(j)
            End If
            forms!frmPLM.Form.Dirty = False
            
            CurrentDb().Execute "UPDATE tblPLM SET [" & fld.name & "] = '" & strValue & "' WHERE [FilePath] = '" & typRecord.FilePath & "'"
        End If
        If fld.name = "File_Data_Name" Then
            startIt = True
            forms!frmPLM.Form.Dirty = False
        End If
    Next
    forms!frmPLM.Form.Dirty = False
Set fld = Nothing
        If i <> lngRecCnt Then
            Dim typNextRecord As Record
            If iobjRecords.fncItem(i + 1, typNextRecord) = False Then
                Exit Function
            End If
        End If
    Next i
    
    fncWriteExcel = True
End Function

Private Function fncWriteExcelForUpdate(ByRef iobjRecords As CATIAPropertyTable) As Boolean
    fncWriteExcelForUpdate = False
    
    Dim rs1 As Recordset
    Dim db As Database
    Set db = CurrentDb()
    Set rs1 = db.OpenRecordset("tblPLM", dbOpenSnapshot)
    
    forms!frmPLM.Form.Recordset.MoveFirst

    Dim lngRecCnt As Long
    lngRecCnt = iobjRecords.fncCount
    
    '/ Excel
    Dim i As Long
    For i = 1 To lngRecCnt
        Dim typRecord As Record
        If i > 1 Then
            forms!frmPLM.Form.Recordset.MoveNext
        End If
        If iobjRecords.fncItem(i, typRecord) = False Then
            Exit Function
        End If
        
        '/ FileName
        Dim objCell
        If forms!frmPLM.Form!FileName.Value <> typRecord.FileName Then
            forms!frmPLM.Form!FileName.Value = typRecord.FileName
        End If
        
        '/ FilePath
        If forms!frmPLM.Form!File_Path.Value <> typRecord.FilePath Then
            forms!frmPLM.Form!File_Path.Value = typRecord.FilePath
        End If
        
        '/ DrawLinkTo
        If forms!frmPLM.Form!Link_To.Value <> typRecord.LinkTo Then
            forms!frmPLM.Form!Link_To.Value = typRecord.LinkTo
        End If
        
        '/ ModelID / DrawingID
        If forms!frmPLM!.Controls("ModelID/DrawingID").Value <> typRecord.ModelDrawingID Then
            forms!frmPLM!.Controls("ModelID/DrawingID").Value = typRecord.ModelDrawingID
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
                strDummyValue = modConst.VALUE_UNSET_STR
            Else
                strDummyValue = modConst.VALUE_UNSET_NUM
            End If
            
            Dim strValue As String
            If strReq = "0" And typRecord.Properties(j) = strDummyValue Then
                strValue = ""
            Else
                strValue = typRecord.Properties(j)
            End If
            
            If Nz(rs1(gcurMainProperty(j)), "") <> strValue Then
                
                CurrentDb().Execute "UPDATE tblPLM SET [" & gcurMainProperty(j) & "] = '" & strValue & "' WHERE [ID] = " & rs1![ID]
                forms!frmPLM.Form.Dirty = False
                forms!frmPLM.sfrmPLM.Form.Dirty = False
            End If
        Next j
        rs1.MoveNext
    Next i
    
    rs1.Close
    Set rs1 = Nothing
    forms!frmPLM.Form.Dirty = False
    forms!frmPLM.sfrmPLM.Form.Dirty = False
    fncWriteExcelForUpdate = True
End Function

Private Function fncWriteExcelForDataMove(ByRef iobjRecords As CATIAPropertyTable) As Boolean

    fncWriteExcelForDataMove = False

    Dim lngRecCnt As Long
    lngRecCnt = iobjRecords.fncCount
    
    Dim i As Long
    For i = 1 To lngRecCnt
        
        Dim typRecord As Record
        If iobjRecords.fncItem(i, typRecord) = False Then
            Exit Function
        End If
        
        '/ ModelID/DrawingID
        Dim objCell 'As Range
        'If objCell.value <> typRecord.ModelDrawingID Then
            'objCell.value = typRecord.ModelDrawingID
        'End If
        
    Next i
    
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
oblnCheckBox = forms!frmPLM!.Controls(istrCheckBoxName).Value
fncGetCheckBox = True
End Function

Private Function fncSetCheckBox(ByVal istrSheetName As String, ByVal istrCheckBoxName As String, _
                                ByVal iblnChecked As Boolean) As Boolean

fncSetCheckBox = False
If istrCheckBoxName <> "chkSel" Then
forms!frmPLM!.Controls(istrCheckBoxName).Value = iblnChecked
End If
fncSetCheckBox = True
End Function

Private Sub ClearDeletePropertyCheckBox()

    CurrentDb().Execute "UPDATE [tblPLM] SET [Sel] = False"
End Sub

Private Function fncGetProperty() As CATIAPropertyTable
    Set fncGetProperty = New CATIAPropertyTable
    
    Dim rs1 As Recordset
    Dim db As Database
    Set db = CurrentDb()
    Set rs1 = db.OpenRecordset("tblPLM", dbOpenSnapshot)

    Dim i As Long
    For i = 1 To DCount("[ID]", "[tblPLM]")
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
                    Case StrConv(CATDRAWING, vbUpperCase)
                        strProperties(j) = CATDRAWING
                    Case StrConv(CATPRODUCT, vbUpperCase)
                        strProperties(j) = CATPRODUCT
                    Case StrConv(CATPART, vbUpperCase)
                        strProperties(j) = CATPART
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
        typRecord.Sel = Nz(rs1![Sel], "")
        typRecord.ID = Nz(rs1![File_ID], "")
        typRecord.Level = Nz(rs1![Lv], "")
        typRecord.Amount = Nz(rs1![Amount], "")
        typRecord.LinkID = Nz(rs1![Link_ID], "")
        typRecord.FilePath = Nz(rs1![FilePath], "")
        typRecord.FileName = Nz(rs1![FileName], "")
        typRecord.LinkTo = Nz(rs1![LinkTo], "")
        typRecord.partNumber = Nz(rs1![partNumber], "")
        typRecord.InstanceName = Nz(rs1![InstanceName], "")
        typRecord.ModelDrawingID = Nz(rs1![ModelID/DrawingID], "")
        typRecord.Properties = strProperties
        
        If fncIsBlankRecord(typRecord) = True Then
            Exit For
        Else
            Call fncGetProperty.fncAddRecord(typRecord)
        End If
        rs1.MoveNext
    Next i
    rs1.Close
    Set rs1 = Nothing
End Function

Public Function fncGetIndex(ByVal istrPropertyName As String) As Long
    fncGetIndex = 0
    On Error GoTo 0

    fncGetIndex = fncGetColumn(istrPropertyName) - 11 '-column L
End Function

Private Function fncGetColumn(ByVal istrPropertyName As String) As Long
Dim db As Database
Dim rs1 As Recordset
Set db = CurrentDb()
Set rs1 = db.OpenRecordset("tblPLM", dbOpenSnapshot)
Dim fld As DAO.Field
Dim i As Integer
i = 0

For Each fld In rs1.Fields
    If fld.name = istrPropertyName Then
        fncGetColumn = i
        Exit Function
    End If
    i = i + 1
Next
Set fld = Nothing
End Function

Private Function fncIsBlankRecord(ByRef itypRecord As Record) As Boolean
    fncIsBlankRecord = False
    
    If itypRecord.Level <> "" Then
        Exit Function
    ElseIf itypRecord.FilePath <> "" Then
        Exit Function
    ElseIf itypRecord.FileName <> "" Then
        Exit Function
    ElseIf itypRecord.LinkTo <> "" Then
        Exit Function
    ElseIf itypRecord.partNumber <> "" Then
        Exit Function
    ElseIf itypRecord.InstanceName <> "" Then
        Exit Function
    ElseIf itypRecord.ModelDrawingID <> "" Then
        Exit Function
    End If
    
    On Error Resume Next
    Dim lngCnt As Long
    lngCnt = UBound(itypRecord.Properties)
    On Error GoTo 0
    
    Dim i As Long
    For i = 1 To lngCnt
        If itypRecord.Properties(i) <> "" Then
            Exit Function
        End If
    Next i

    fncIsBlankRecord = True
End Function

Private Function fncSetAllSel(ByRef iExcelData As CATIAPropertyTable, _
                              ByVal istrValue As Boolean) As Boolean
fncSetAllSel = False
On Error GoTo 0

If istrValue = True Then

    CurrentDb().Execute "UPDATE [tblPLM] SET [Sel] = TRUE;"
Else

    CurrentDb().Execute "UPDATE [tblPLM] SET [Sel] = FALSE;"
End If
forms!frmPLM!sfrmPLM.Form.Dirty = False
    
fncSetAllSel = True
End Function

Private Function fncClrModelID(ByRef iExcelData As CATIAPropertyTable) As Boolean
    fncClrModelID = False
    
    Dim i As Long
    For i = 1 To iExcelData.fncCount
    
        Dim typRecord As Record
        Call iExcelData.fncItem(i, typRecord)
        
        If Trim(typRecord.Sel) = True Then
            Dim lngCol As Long
            lngCol = fncGetColumn(TITLE_MODELIDDRAWID)
            
            'objCell.value = ""
        End If
    Next i
    
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
        Dim i As Long
        For i = 1 To lngRecCnt
        
            If fncCheckNumberingRow(iobjRecords, i) = False Then
                GoTo CONTINUE_FNCCHECK
            End If
    
            Dim typRecord As Record
            If iobjRecords.fncItem(i, typRecord) = False Then
                GoTo CONTINUE_FNCCHECK
            End If
            
            Dim lngTypeIndex As Long
            lngTypeIndex = modMain.fncGetIndex(TITLE_FILEDATATYPE)
            
            Dim strType As String
            strType = typRecord.Properties(lngTypeIndex)
            If strType <> CATDRAWING Then
                GoTo CONTINUE_FNCCHECK
            End If
    
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
            lngDesignNoIndex = modMain.fncGetIndex(TITLE_DESIGNNO)
            
            Dim strDesignNo As String
            strDesignNo = typLink.Properties(lngDesignNoIndex)
            If strDesignNo <> "" Then
                GoTo CONTINUE_FNCCHECK
            End If
            
            If Trim(typLink.Sel) <> True Then
                fncCheckBeforeNumbering = "E031"
                Exit Function
            End If
    
CONTINUE_FNCCHECK:
        Next i
    End If
End Function

Private Function fncNumbering(ByRef iobjRecords As CATIAPropertyTable, ByRef oblnBlank3D As Boolean) As String

    fncNumbering = ""
    oblnBlank3D = False
    
    Dim objSheet
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
    Dim i As Long
    For i = 1 To lngRecCnt
    
        If fncCheckNumberingRow(iobjRecords, i) = False Then
            GoTo CONTINUE
        End If
        
        If DEBUG_MODE = True Then
            fncNumbering = fncNumberingRow_DEBUG(iobjRecords, objSheet, i)
        Else
            fncNumbering = fncNumberingRow(iobjRecords, objSheet, i)
        End If
        
        If fncNumbering <> "" Then
            Exit Function
        End If
        
        Dim lngNumberedCnt As Long
        On Error Resume Next
        lngNumberedCnt = UBound(curNumberedRow)
        On Error GoTo 0
        ReDim Preserve curNumberedRow(lngNumberedCnt + 1) As Long
        curNumberedRow(lngNumberedCnt + 1) = i
        
        isNumbering = True
        
CONTINUE:
    Next i
    
    Dim blnDuplicateChkBox As Boolean
    blnDuplicateChkBox = False
    Call modMain.fncGetDuplicateDesignNoCheckBox(blnDuplicateChkBox)
    
    If blnDuplicateChkBox = True Then
        On Error Resume Next
        lngRecCnt = 0
        lngRecCnt = UBound(curNumberedRow)
        On Error GoTo 0
        For i = 1 To lngRecCnt
            
            fncNumbering = fncNumberingForDrawing(iobjRecords, objSheet, curNumberedRow(i))
            If fncNumbering <> "" Then
                Exit Function
            End If
        
        Next i
    End If
    
    If oblnBlank3D = False And isNumbering = False Then
        fncNumbering = "E020"
    End If
End Function

Private Function fncCheckNumberingRow(ByRef iobjRecords As CATIAPropertyTable, ByVal i As Long) As Boolean
    fncCheckNumberingRow = False
        
    Dim typRecord As Record
    If iobjRecords.fncItem(i, typRecord) = False Then
        Exit Function
    End If
    
    If Trim(typRecord.Sel) <> True Then
        Exit Function
    End If
    
    '/ DesignNo
    Dim lngDesignNoIndex As Long
    lngDesignNoIndex = modMain.fncGetIndex(TITLE_DESIGNNO)
    
    Dim strDesignNo As String
    strDesignNo = typRecord.Properties(lngDesignNoIndex)
    If Trim(strDesignNo) <> "" Then
        Exit Function
    End If
    
    fncCheckNumberingRow = True
End Function

Private Function fncNumberingRow(ByRef iobjRecords As CATIAPropertyTable, ByRef objSheet, _
                                 ByVal i As Long) As String

    fncNumberingRow = ""
    
    Dim typRecord As Record
    If iobjRecords.fncItem(i, typRecord) = False Then
        Exit Function
    End If
    
    Dim strType As String
    Dim strLinkID As String
    strType = typRecord.Properties(modMain.fncGetIndex(TITLE_FILEDATATYPE))
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
    
    If strType = CATDRAWING And blnDuplicateChkBox = True Then
    Else
        fncNumberingRow = fncNumberingDesignNo(con, objSheet, i)
        If fncNumberingRow <> "" Then
            GoTo Finally
        End If
    End If
    
    GoTo Finally

Error:
    fncNumberingRow = "E021"
Finally:
    If Not con Is Nothing Then
        If con.State = adStateOpen Then
            con.Close
        End If
        Set con = Nothing
    End If
End Function

Private Function fncNumberingRow_DEBUG(ByRef iobjRecords As CATIAPropertyTable, ByRef objSheet, _
                                       ByVal i As Long) As String
    
    fncNumberingRow_DEBUG = ""
    
    Dim typRecord As Record
    If iobjRecords.fncItem(i, typRecord) = False Then
        Exit Function
    End If
    
    Dim strType As String
    Dim strLinkID As String
    strType = typRecord.Properties(modMain.fncGetIndex(TITLE_FILEDATATYPE))
    strLinkID = typRecord.LinkID
    
    Dim blnDuplicateChkBox As Boolean
    blnDuplicateChkBox = False
    Call modMain.fncGetDuplicateDesignNoCheckBox(blnDuplicateChkBox)
    
    If strType = CATDRAWING And blnDuplicateChkBox = True Then
    Else
        fncNumberingRow_DEBUG = fncNumberingDesignNo_DEBUG(objSheet, i)
    End If
End Function

Private Function fncNumberingForDrawing(ByRef iobjRecords As CATIAPropertyTable, ByRef objSheet, _
                                        ByVal i As Long) As String

    fncNumberingForDrawing = ""
    
    Dim typRecord As Record
    If iobjRecords.fncItem(i, typRecord) = False Then
        Exit Function
    End If
    
    Dim strType As String
    strType = typRecord.Properties(modMain.fncGetIndex(TITLE_FILEDATATYPE))
    If strType <> CATDRAWING Then
        Exit Function
    End If
    
    Dim objFixedCell
    Set objFixedCell = objSheet.Range(CELL_MAIN_FIXED)
    
    Dim objUserCell
    Set objUserCell = objSheet.Range(CELL_MAIN_USER)
    
    Dim lngColLinkTo As Long
    lngColLinkTo = fncGetColumn(TITLE_LINKTO)
    
    Dim objCellLinkTo
    Set objCellLinkTo = objSheet.Cells(objFixedCell.Row + i, lngColLinkTo)
    
    Dim strLinkTo As String
    strLinkTo = objCellLinkTo.Value
    
    Dim lngPartRow As Long
    lngPartRow = iobjRecords.fncSearchFromFilePath(strLinkTo)
    If lngPartRow = 0 Then
        fncNumberingForDrawing = "E027"
        Exit Function
    End If
    
    Dim lngColDesignNo As Long
    lngColDesignNo = fncGetColumn(TITLE_DESIGNNO)
    
    Dim objCellPartDesignNo 'As Range
    Dim objCellDrawDesignNo 'As Range
    Set objCellPartDesignNo = objSheet.Cells(objUserCell.Row + lngPartRow, lngColDesignNo)
    Set objCellDrawDesignNo = objSheet.Cells(objUserCell.Row + i, lngColDesignNo)
    objCellDrawDesignNo.Value = objCellPartDesignNo.Value
End Function

Private Function fncGetConnectString() As String
    fncGetConnectString = ""
    
    Dim DBServ As String
    DBServ = modSetting.gstrServerName
    
    Dim DBName As String
    DBName = modSetting.gstrDBName
    
    Dim DBUser As String
    DBUser = modSetting.gstrUserName
    
    Dim DBPass As String
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
    
    Dim DBServ As String
    DBServ = modSetting.gstrOldServerName
    
    Dim DBName As String
    DBName = modSetting.gstrOldDBName
    
    Dim DBUser As String
    DBUser = modSetting.gstrOldUserName
    
    Dim DBPass As String
    DBPass = modSetting.gstrOldPassword
    
    fncGetOldConnectString = "Provider=Sqloledb;" & _
                            "Data Source=" & DBServ & ";" & _
                            "Initial Catalog=" & DBName & ";" & _
                            "Connect Timeout=10;" & _
                            "user id=" & DBUser & ";" & _
                            "password=" & DBPass
End Function

Private Function fncNumberingDesignNo(ByRef con As ADODB.Connection, ByRef objSheet, ByVal i As Long) As String
    fncNumberingDesignNo = ""
        
    Dim lRec As ADODB.Recordset
    Dim devCode As String
    devCode = modDefineDevelopment.gstrOfficeCode
    Dim tblName As String
    tblName = modDefineDevelopment.gstrNumberingTable

    On Error GoTo Error

    Dim objWshNetwork As Object
    Set objWshNetwork = CreateObject("WScript.Network")
    Dim osUser As String
    osUser = objWshNetwork.userName

    Dim lSql As String
    lSql = "INSERT INTO [dbo].[" & tblName & "] ([CREATEDATE],[OSUSER])" & _
                         "VALUES (GetDate(), '" & osUser & "')"
    con.Execute (lSql)
    On Error GoTo 0

    '/ +1‚³‚ê‚½DesignNo‚ðŽæ“¾
    On Error GoTo Error
    Dim designNo As String
    lSql = "SELECT SCOPE_IDENTITY()"
    Set lRec = con.Execute(lSql)
    designNo = lRec.Fields(0).Value
    On Error GoTo 0
    
    
    CurrentDb().Execute "UPDATE tblPLM SET Design_No = " & designNo
    
    GoTo Finally
Error:
    fncNumberingDesignNo = "E023"
Finally:
    If Not lRec Is Nothing Then
        lRec.Close
        Set lRec = Nothing
    End If
End Function

Private Function fncNumberingDesignNo_DEBUG(ByRef objSheet, ByVal i As Long) As String
    fncNumberingDesignNo_DEBUG = "E023"
    
    '/ Setting/MaxDesignNo
    Dim objSettingSheet
    'Set objSettingSheet = Excel.Worksheets.item("Setting")
    
    Dim devCode As String
    devCode = modDefineDevelopment.gstrOfficeCode
    
    Dim strText As String
    strText = objSettingSheet.Cells(2, 5).Text
    
    '/ MaxModelID
    Dim lngMaxID As Long
    If IsNumeric(strText) = False Then
        Exit Function
    Else
        lngMaxID = strText
        lngMaxID = lngMaxID + 1
    End If
    
    objSettingSheet.Cells(2, 5).Value = lngMaxID
    
    Dim objUserCell 'As Range
    Set objUserCell = objSheet.Range(CELL_MAIN_USER)
    
    Dim lngCol As Long
    lngCol = fncGetColumn(TITLE_DESIGNNO)
    
    Dim objCell 'As Range
    Set objCell = objSheet.Cells(objUserCell.Row + i, lngCol)
     objCell.Value = lngMaxID
    
    fncNumberingDesignNo_DEBUG = ""
End Function

Private Function fncNumberingDesignNo_Asterisk(ByRef objSheet, ByVal i As Long) As String
    fncNumberingDesignNo_Asterisk = "E023"
    
    Dim objUserCell 'As Range
    Set objUserCell = objSheet.Range(CELL_MAIN_USER)
    
    Dim lngCol As Long
    lngCol = fncGetColumn(TITLE_DESIGNNO)
    
    Dim objCell 'As Range
    Set objCell = objSheet.Cells(objUserCell.Row + i, lngCol)
    objCell.Value = True
     
    fncNumberingDesignNo_Asterisk = ""
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
    
    '/ ‹ŒÌ”ÔDB‚ÉÚ‘±
    Dim con As New ADODB.Connection
    On Error GoTo Finally
    Dim conStr As String
    conStr = fncGetOldConnectString()
    con.open (conStr)
    On Error GoTo 0
    
    '/ ‘ÎÛModelID‚Ì‘®«‚ðŒŸõ
    On Error GoTo Finally
    Dim lSql As String
    lSql = "SELECT T3.ATTRNAME, T2.ATTRVALUE, T1.MODELID FROM " & modSetting.gstrOldDBName & ".dbo.CATIAMODEL T1 " & _
           "INNER JOIN " & modSetting.gstrOldDBName & ".dbo.CATIA_ATTR_VALUE T2 ON T1.MODELID = T2.MODELID " & _
           "INNER JOIN " & modSetting.gstrOldDBName & ".dbo.CATIA_ATTR_NAME T3 ON T3.ATTR_ID = T2.ATTR_ID " & _
           "WHERE T1.modelID = "
    
    Dim i As Long
    For i = 1 To lngCnt
        If i <> 1 Then
            lSql = lSql & " Or T1.modelID = "
        End If
        lSql = lSql & ilstModelID(i)
    Next i
    
    Dim objRecord As ADODB.Recordset
    Set objRecord = con.Execute(lSql)
    On Error GoTo 0
    
    If iobjTable.fncSetPropertyFromDB(ilstModelID, objRecord) = False Then
        GoTo Finally
    End If
    
    fncGetPropertyFromDB = True
    
Finally:
    If Not con Is Nothing Then
        If con.State = adStateOpen Then
            con.Close
        End If
        Set con = Nothing
    End If
End Function

Public Function fncCreateSaveDir(ByRef iobjCatiaData As CATIAPropertyTable) As String
    fncCreateSaveDir = ""
    
    Dim typTopRecord As Record
    If iobjCatiaData.fncItem(1, typTopRecord) = False Then
        Exit Function
    End If
    
    Dim lngFileName As Long
    lngFileName = modMain.fncGetIndex(TITLE_FILEDATANAME) + 1
    
    Dim strFileName As String
    strFileName = typTopRecord.Properties(lngFileName)
    If strFileName = "" Then
       Exit Function
    End If
    
    Dim strSaveDir As String
    strSaveDir = modSetting.gstrSendToPath & "\" & strFileName
    
    On Error Resume Next
    Call SHCreateDirectoryEx(0&, strSaveDir, 0&)
    On Error GoTo 0
    
    On Error Resume Next
    Dim strResult As String
    strResult = Dir(strSaveDir, vbDirectory)
    On Error GoTo 0
    
    If strResult <> "" Then
        fncCreateSaveDir = strSaveDir
    End If
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