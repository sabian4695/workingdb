Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Option Explicit
Private mRecords() As Record
Private gUnknownSection() As String
Private gUnknownStatus() As String
Public gBlankSection As Boolean
Public gBlankStatus As Boolean

Private Sub Class_Initialize()
    ReDim mRecords(0) As Record
    ReDim gUnknownSection(0) As String
    ReDim gUnknownStatus(0) As String
    gBlankSection = False
    gBlankStatus = False
End Sub

Private Sub Class_Terminate()
    ReDim mRecords(0) As Record
    ReDim gUnknownSection(0) As String
    ReDim gUnknownStatus(0) As String
    gBlankSection = False
    gBlankStatus = False
End Sub

Public Function fncCount() As Long
    fncCount = 0
    On Error Resume Next
    fncCount = UBound(mRecords)
    On Error GoTo 0
End Function

Public Function fncCountUnknownSection() As Long
    fncCountUnknownSection = 0
    On Error Resume Next
    fncCountUnknownSection = UBound(gUnknownSection)
    On Error GoTo 0
End Function

Public Function fncCountUnknownStatus() As Long
    fncCountUnknownStatus = 0
    On Error Resume Next
    fncCountUnknownStatus = UBound(gUnknownStatus)
    On Error GoTo 0
End Function

Public Function fncItem(ByVal ilngIndex As Long, ByRef otypRecord As Record) As Boolean
    fncItem = False

    Dim lngCnt As Long
    lngCnt = Me.fncCount()
    
    If lngCnt < ilngIndex Then
        Exit Function
    End If

    otypRecord = mRecords(ilngIndex)
    
    fncItem = True
End Function

Public Function fncAddRecord(ByRef itypRecord As Record) As Long
    fncAddRecord = -1

    Dim lngCnt As Long
    lngCnt = Me.fncCount()
    
    Dim strType As String
    strType = itypRecord.Properties(modMain.fncGetIndex(TITLE_FILEDATATYPE) - 1)
    
    If strType = "Component" Then
        fncAddRecord = 0
        Exit Function
    Else
        Dim i As Integer
        For i = 1 To lngCnt
            If mRecords(i).FilePath = itypRecord.FilePath Then
                mRecords(i).Amount = mRecords(i).Amount + 1
                itypRecord.IsChildInstance = True
                itypRecord.Amount = "-"
                itypRecord.InstanceName = ""
                itypRecord.ModelDrawingID = ""
                itypRecord.partNumber = ""
                
                On Error Resume Next
                Dim lngPropCnt As Long
                lngPropCnt = 0
                lngPropCnt = UBound(itypRecord.Properties)
                On Error GoTo 0
                
                Dim j As Long
                For j = 1 To lngPropCnt
                    itypRecord.Properties(j) = ""
                Next j
                
                Exit For
            End If
        Next i

        itypRecord.ID = lngCnt + 1
            
        ReDim Preserve mRecords(lngCnt + 1)
        mRecords(lngCnt + 1) = itypRecord
        
        fncAddRecord = lngCnt + 1
    End If
End Function

Public Function fncIsSameStructure(ByRef iobjData As CATIAPropertyTable, _
                                   ByRef oblnIsSame As Boolean) As Boolean
    
    fncIsSameStructure = False
    oblnIsSame = False
    Dim lngDataSize As Long
    Dim lngMyDataSize As Long
    lngDataSize = iobjData.fncCount()
    lngMyDataSize = Me.fncCount()
    If lngDataSize <> lngMyDataSize Then
        fncIsSameStructure = True
        oblnIsSame = False
        Exit Function
    ElseIf lngDataSize <= 0 Then
        fncIsSameStructure = False
        oblnIsSame = False
    End If
    
    Dim i As Long
    For i = 1 To lngDataSize
    
        Dim typData As Record
        Dim typMyData As Record
        If iobjData.fncItem(i, typData) = False Then
            Exit Function
        ElseIf Me.fncItem(i, typMyData) = False Then
            Exit Function
        End If
        
        If typData.FilePath <> typMyData.FilePath Then
            fncIsSameStructure = True
            oblnIsSame = False
        End If

    Next i

    fncIsSameStructure = True
    oblnIsSame = True
End Function

Public Function fncGetLastLevel() As Long
    fncGetLastLevel = 0

    Dim lngCnt As Long
    lngCnt = Me.fncCount()
    
    Dim i As Integer
    For i = 1 To lngCnt
        Dim typRecord As Record
        typRecord = mRecords(i)
        If fncGetLastLevel < typRecord.Level Then
            fncGetLastLevel = typRecord.Level
        End If
    Next i
End Function

Public Function fncCheckBlank(ByRef ostrPropertyName As String) As String
    fncCheckBlank = ""

    If modSetting.gstrInputCheck <> "1" Then
        Exit Function
    End If

    Dim lngCnt As Long
    lngCnt = Me.fncCount()
    
    Dim blnNoCheckFlag As Boolean
    blnNoCheckFlag = False
    
    Dim i As Long
    For i = 1 To lngCnt
    
        Dim typRecord As Record
        If Me.fncItem(i, typRecord) = False Then
            GoTo CONTINUE
        End If
        
        Dim lngValIndex As Long
        Dim strValue As String
    
        lngValIndex = modMain.fncGetIndex(TITLE_CLASSIFICATION) - 1
        strValue = typRecord.Properties(lngValIndex)
        
        If typRecord.Level <= 1 Then
            If strValue = VALUE_2KMOULD Or _
               strValue = VALUE_SUBPRODUCT Or _
               strValue = VALUE_REFERENCE Or _
               strValue = VALUE_LAYOUT Or _
               strValue = VALUE_CUSTOMERAPPROVEDDATA Then
                blnNoCheckFlag = True
            Else
                blnNoCheckFlag = False
            End If
        End If
        
        If blnNoCheckFlag = True Then
            GoTo CONTINUE
        End If
        
        Dim strSel As String
        strSel = typRecord.Sel
        If Trim(strSel) <> True Then
            GoTo CONTINUE
        End If
        
        lngValIndex = modMain.fncGetIndex(TITLE_DESIGNNO) - 1
        strValue = typRecord.Properties(lngValIndex)
        If Trim(strValue) = "" Then
            fncCheckBlank = "E034"
            Exit Function
        End If

        lngValIndex = modMain.fncGetIndex(TITLE_CURRENTSTATUS) - 1
        strValue = typRecord.Properties(lngValIndex)
        If Trim(strValue) = "" Then
            fncCheckBlank = "E047"
            Exit Function
        End If
        
        Dim lngPropCnt As Long
        On Error Resume Next
        lngPropCnt = UBound(modMain.gcurMainProperty)
        Dim j As Long
        For j = 1 To lngPropCnt
        
            Dim strPropName As String
            strPropName = modMain.gcurMainProperty(j)
            
            Dim strReq As String
            strReq = modDefineDrawing.fncGetInputRequired(strPropName)
            
            If strReq = "1" Then
                strValue = typRecord.Properties(j)
                If Trim(strValue) = "" Then
                    fncCheckBlank = "E038"
                    ostrPropertyName = strPropName
                    Exit Function
                End If
            End If
            
        Next j
        
CONTINUE:
    Next i
End Function

Private Function fncIsDateFormat(ByVal strVal As String) As Boolean

    fncIsDateFormat = False
    
    Dim varTemp As Variant
    varTemp = Split(Trim(strVal), " ")
    
    Dim lngCnt As Long
    On Error Resume Next
    lngCnt = UBound(varTemp)
    On Error GoTo 0
    
    Dim strDate As String
    Dim strTime As String
    
    If 2 <= lngCnt Then
        Exit Function
    ElseIf lngCnt = 1 Then
        strDate = varTemp(0)
        strTime = varTemp(1)
        
    ElseIf lngCnt = 0 Then
        strDate = varTemp(0)
    End If
    
    Dim blnDate As Boolean
    blnDate = IsDate(Format(strDate, "dd/mm/yy"))
    
    Dim blnTime As Boolean
    If strTime <> "" Then
        blnTime = (Format(strTime, "hh:mm:ss") = strTime) And IsDate(strTime)
    Else
        blnTime = True
    End If
    
    If blnDate = True And blnTime = True Then
        fncIsDateFormat = True
    Else
        fncIsDateFormat = False
    End If
End Function

Public Function fncSearchFromFilePath(ByVal istrFilePath As String) As Long
    fncSearchFromFilePath = 0

    Dim lngRecCnt As Long
    lngRecCnt = Me.fncCount
    
    Dim i As Long
    For i = 1 To lngRecCnt
        If mRecords(i).FilePath = istrFilePath Then
            fncSearchFromFilePath = i
            Exit For
        End If
    Next i
End Function

Public Sub UpdatePath(ByVal istrOldPath As String, ByVal istrNewPath As String, ByVal istrNewName As String)
    Dim lngRecCnt As Long
    lngRecCnt = Me.fncCount
    
    Dim i As Long
    For i = 1 To lngRecCnt
        If mRecords(i).FilePath = istrOldPath Then
            mRecords(i).FilePath = istrNewPath
            mRecords(i).fileName = istrNewName
        End If
        If mRecords(i).LinkTo = istrOldPath Then
            mRecords(i).LinkTo = istrNewPath
        End If
    Next i
End Sub

Public Sub UpdateModelID(ByVal ilngIndex As Long, ByVal istrValue As String)
    Dim lngRecCnt As Long
    lngRecCnt = Me.fncCount
    
    If 0 < ilngIndex And ilngIndex <= lngRecCnt Then
        mRecords(ilngIndex).ModelDrawingID = istrValue
    End If
End Sub

Public Function fncReplaceProhibitCharacter() As Boolean
    fncReplaceProhibitCharacter = False

    Dim lngIndex_FileDataName As Long
    Dim lngIndex_FullDesignNo As Long
    lngIndex_FileDataName = modMain.fncGetIndex(TITLE_FILEDATANAME)
    lngIndex_FullDesignNo = modMain.fncGetIndex(TITLE_FULLDESIGNNO)

    Dim lngRecCnt As Long
    lngRecCnt = Me.fncCount
    
    Dim i As Long
    For i = 1 To lngRecCnt
    
        Dim strSel As String
        strSel = mRecords(i).Sel
        If Trim(strSel) <> True Then
            GoTo CONTINUE
        End If
        
        Dim strBuf As String
        strBuf = mRecords(i).Properties(lngIndex_FileDataName)
        Call ReplaceString(strBuf)
        If strBuf <> mRecords(i).Properties(lngIndex_FileDataName) Then
            mRecords(i).Properties(lngIndex_FileDataName) = strBuf
            fncReplaceProhibitCharacter = True
        End If
        
        strBuf = mRecords(i).Properties(lngIndex_FullDesignNo)
        Call ReplaceString(strBuf)
        If strBuf <> mRecords(i).Properties(lngIndex_FullDesignNo) Then
            mRecords(i).Properties(lngIndex_FullDesignNo) = strBuf
            fncReplaceProhibitCharacter = True
        End If
CONTINUE:
    Next i
End Function

Public Sub ClearModelID()
    Dim lngRecCnt As Long
    lngRecCnt = Me.fncCount
    
    Dim i As Long
    For i = 1 To lngRecCnt
    
        Dim strSel As String
        strSel = mRecords(i).Sel
        If Trim(strSel) <> True Then
            GoTo CONTINUE
        End If
        
        mRecords(i).ModelDrawingID = ""

CONTINUE:
    Next i
End Sub

Private Sub ReplaceString(ByRef ostrString As String)
    On Error Resume Next
    Dim lngSize As Long
    lngSize = UBound(modConst.glstProhibitCharacter)
    On Error GoTo 0

    Dim i As Long
    For i = 1 To lngSize
        Dim strFind As String
        strFind = modConst.glstProhibitCharacter(i)
        ostrString = Replace(ostrString, strFind, modConst.REPLACE_CHAR)
    Next i
End Sub

Public Sub SetDefaultDesinerSection()
    Dim lngIndex_Designer As Long
    Dim lngIndex_Section As Long
    Dim lngIndex_Status As Long
    Dim lngIndex_DesignNo As Long
    Dim lngIndex_RevisionNo As Long
    Dim lngIndex_FileDataName As Long
    Dim lngIndex_FullDesignNo As Long
    Dim lngIndex_Classification As Long
    Dim lngIndex_FileName As Long
    Dim lngIndex_FileDataType As Long
    Dim lngIndex_MaterialGrade As Long
    lngIndex_Designer = modMain.fncGetIndex(modMain.gstrDesignerName) - 1
    lngIndex_Section = modMain.fncGetIndex(TITLE_SECTION) - 1
    lngIndex_Status = modMain.fncGetIndex(TITLE_CURRENTSTATUS) - 1
    lngIndex_DesignNo = modMain.fncGetIndex(TITLE_DESIGNNO) - 1
    lngIndex_RevisionNo = modMain.fncGetIndex(TITLE_REVISIONNO) - 1
    lngIndex_FileDataName = modMain.fncGetIndex(TITLE_FILEDATANAME) - 1
    lngIndex_FullDesignNo = modMain.fncGetIndex(TITLE_FULLDESIGNNO) - 1
    lngIndex_Classification = modMain.fncGetIndex(TITLE_CLASSIFICATION) - 1
    lngIndex_FileName = modMain.fncGetIndex(TITLE_FILENAME) - 1
    lngIndex_FileDataType = modMain.fncGetIndex(TITLE_FILEDATATYPE) - 1
    lngIndex_MaterialGrade = modMain.fncGetIndex(TITLE_MATERIALGRADE) - 1
    
    Dim blnNoCheckFlag As Boolean
    blnNoCheckFlag = False
    Dim blnSetFileName As Boolean
    blnSetFileName = False

    Dim lngRecCnt As Long
    lngRecCnt = Me.fncCount
    
    Dim i As Long
    For i = 1 To lngRecCnt
        If mRecords(i).IsChildInstance = True Then
            GoTo CONTINUE
        End If
    
        Dim strDesigner As String
        strDesigner = mRecords(i).Properties(lngIndex_Designer)
        If Trim(strDesigner) = "" Then
            mRecords(i).Properties(lngIndex_Designer) = modMain.gstrNFDesigner
            strDesigner = modMain.gstrNFDesigner
        End If
        
        Dim strSection As String
        strSection = mRecords(i).Properties(lngIndex_Section)
        If Trim(strSection) = "" Then
            mRecords(i).Properties(lngIndex_Section) = modDefineDevelopment.gstrSection
            strSection = modDefineDevelopment.gstrSection
        End If
        
        Dim strClassification As String
        strClassification = mRecords(i).Properties(lngIndex_Classification)
        
        If modSetting.gstrInputCheck <> "1" Then
            blnNoCheckFlag = True
            blnSetFileName = True
        ElseIf mRecords(i).Level <= 1 Then
            If strClassification = VALUE_2KMOULD Or _
               strClassification = VALUE_SUBPRODUCT Or _
               strClassification = VALUE_REFERENCE Or _
               strClassification = VALUE_LAYOUT Or _
               strClassification = VALUE_CUSTOMERAPPROVEDDATA Then
                blnNoCheckFlag = True
            Else
                blnNoCheckFlag = False
            End If
            
            If strClassification = VALUE_CUSTOMERAPPROVEDDATA Then
                blnSetFileName = True
            Else
                blnSetFileName = False
            End If
        End If
        
        Dim strSel As String
        strSel = mRecords(i).Sel
        If Trim(strSel) <> True Then
            GoTo CONTINUE
        End If
        
        If Trim(modSetting.gstrUnsetMaterialGrade) = "1" And strClassification = VALUE_SUBMISSIONDATA Then
            mRecords(i).Properties(lngIndex_MaterialGrade) = VALUE_UNSET_STR
        End If
        
        If Trim(modSetting.gstrAutoInput) = "0" Then
            GoTo CONTINUE
        End If
        
        Dim strDesignNo As String
        strDesignNo = Trim(mRecords(i).Properties(lngIndex_DesignNo))
        If blnNoCheckFlag = False And Trim(strDesignNo) = "" Then
            mRecords(i).Properties(lngIndex_FileDataName) = ""
            mRecords(i).Properties(lngIndex_FullDesignNo) = ""
            GoTo CONTINUE
        End If
        
        If blnNoCheckFlag = False And Trim(strSection) = "" Then
            gBlankSection = True
            mRecords(i).Properties(lngIndex_FileDataName) = ""
            mRecords(i).Properties(lngIndex_FullDesignNo) = ""
            GoTo CONTINUE
        End If
        
        Dim blnOldSection As Boolean
        Dim strOfficeCode As String
        blnOldSection = False
        strOfficeCode = ""
        
        Dim strBuff1 As String
        Dim strBuff2 As String
        On Error Resume Next
        strBuff1 = ""
        strBuff2 = ""
        strBuff1 = Left(strDesignNo, 1)
        strBuff2 = Mid(strDesignNo, 2, 1)
        On Error GoTo 0
        
        If IsNumeric(strBuff1) = False And IsNumeric(strBuff2) = False Then
            blnOldSection = True
            strOfficeCode = strBuff1 & strBuff2
            strDesignNo = Mid(strDesignNo, 3)
        ElseIf IsNumeric(strBuff1) = False Then
            blnOldSection = True
            strOfficeCode = strBuff1
            strDesignNo = Mid(strDesignNo, 2)
        Else
            strOfficeCode = modDefineDevelopment.fncGetOfficeCodeFromSection(strSection)
            If blnNoCheckFlag = False And Trim(strOfficeCode) = "" Then
                On Error Resume Next
                Dim lngSectCnt As Long
                lngSectCnt = UBound(gUnknownSection)
                ReDim Preserve gUnknownSection(lngSectCnt + 1) As String
                gUnknownSection(lngSectCnt + 1) = strSection
                On Error GoTo 0
                mRecords(i).Properties(lngIndex_FileDataName) = ""
                mRecords(i).Properties(lngIndex_FullDesignNo) = ""
                GoTo CONTINUE
            End If
            If IsNumeric(strOfficeCode) = True Then
                blnOldSection = False
            Else
                blnOldSection = True
            End If
        End If

        Dim strStatus As String
        strStatus = mRecords(i).Properties(lngIndex_Status)
        If blnNoCheckFlag = False And Trim(strStatus) = "" Then
            gBlankStatus = True
            mRecords(i).Properties(lngIndex_FileDataName) = ""
            mRecords(i).Properties(lngIndex_FullDesignNo) = ""
            GoTo CONTINUE
        End If
        
        If blnNoCheckFlag = False And (strStatus <> VALUE_MASSPRODUCT And strStatus <> VALUE_PROTOTYPESTUDY) Then
            On Error Resume Next
            Dim lngStatusCnt As Long
            lngStatusCnt = UBound(gUnknownStatus)
            ReDim Preserve gUnknownStatus(lngStatusCnt + 1) As String
            gUnknownStatus(lngStatusCnt + 1) = strStatus
            On Error GoTo 0
            mRecords(i).Properties(lngIndex_FileDataName) = ""
            mRecords(i).Properties(lngIndex_FullDesignNo) = ""
            GoTo CONTINUE
        End If
        
        Dim strStatusCode As String
        If strClassification = VALUE_CUSTOMERAPPROVEDDATA Then
            strStatusCode = "C"
'        ElseIf strClassification = "Reference" Then
'            strStatusCode = "R"
        Else
            If blnOldSection = False Then
                If strStatus = VALUE_MASSPRODUCT Then
                    strStatusCode = "M"
                ElseIf strStatus = VALUE_PROTOTYPESTUDY Then
                    strStatusCode = "T"
                End If
            Else
                If strStatus = VALUE_PROTOTYPESTUDY Then
                    strStatusCode = "T"
                Else
                    strStatusCode = ""
                End If
            End If
        End If
        Dim strRevisionNo As String
        strRevisionNo = mRecords(i).Properties(lngIndex_RevisionNo)
        If blnOldSection = False And Len(strDesignNo) < 6 Then
            Dim strAddZero As String
            strAddZero = Right(String(6, "0") & Trim(strDesignNo), 6)
            strDesignNo = strAddZero
        End If
        
        Dim strFullDesignNo As String
        
        If blnOldSection = True And strStatusCode = "C" Then
            strFullDesignNo = strStatusCode & strOfficeCode & strDesignNo
        Else
            strFullDesignNo = strOfficeCode & strStatusCode & strDesignNo
        End If

        If Trim(strRevisionNo) <> "" Then
            strFullDesignNo = strFullDesignNo & "-" & strRevisionNo
        End If

        Dim strOldFullDesignNo As String
        strOldFullDesignNo = mRecords(i).Properties(lngIndex_FullDesignNo)
        
        Dim strOldFullDesignNoSplit() As String
        strOldFullDesignNoSplit = Split(strOldFullDesignNo, "&")
        
        On Error Resume Next
        Dim lngOldFullDesignNoSize As Long
        lngOldFullDesignNoSize = 0
        lngOldFullDesignNoSize = UBound(strOldFullDesignNoSplit)
        On Error GoTo 0
        
        Dim m As Long
        If 0 < lngOldFullDesignNoSize Then
            For m = 1 To lngOldFullDesignNoSize
                strFullDesignNo = strFullDesignNo + "&" + strOldFullDesignNoSplit(m)
            Next m
        End If

        mRecords(i).Properties(lngIndex_FullDesignNo) = strFullDesignNo
        
        Dim strEndChar As String
        strEndChar = ""

        If strClassification = VALUE_2KMOULD Or _
           strClassification = VALUE_REFERENCE Or _
           strClassification = VALUE_SUBPRODUCT Then
            strEndChar = "S"
        ElseIf strClassification = VALUE_LAYOUT Then
            strEndChar = "U"
        ElseIf 2 <= Len(strRevisionNo) And IsNumeric(Right(strRevisionNo, 2)) = True Then
            Dim lngRevNo As Long
            lngRevNo = Right(strRevisionNo, 2)
            If 80 <= lngRevNo Then
                strEndChar = "U"
            Else
                strEndChar = "S"
            End If
        ElseIf strClassification = VALUE_SUBMISSIONDATA Then
            strEndChar = "U"
        Else
            strEndChar = "S"
        End If
        
        If strClassification = VALUE_SUBPRODUCT Or _
           strClassification = VALUE_REFERENCE Or _
           strClassification = VALUE_LAYOUT Then
            strEndChar = "00" & strEndChar
        ElseIf strClassification = VALUE_2KMOULD Then
            strEndChar = strEndChar
        ElseIf Len(strRevisionNo) = 2 Then
        ElseIf Len(strRevisionNo) = 4 Then
            strEndChar = Left(strRevisionNo, 2) & strEndChar
        Else
            strEndChar = strRevisionNo & strEndChar
        End If
        
        Dim strHeadChar As String
        If strClassification = VALUE_SUBPRODUCT Then
            strHeadChar = "S"
        ElseIf strClassification = VALUE_REFERENCE Then
            strHeadChar = "J"
        ElseIf strClassification = VALUE_LAYOUT Then
            strHeadChar = "L"
        ElseIf strClassification = VALUE_2KMOULD Then
            strHeadChar = "W"
        Else
            strHeadChar = ""
        End If
        
        Dim strFileDataName As String
        
        If blnOldSection = True And strStatusCode = "C" Then
            strFileDataName = strStatusCode & strOfficeCode & strDesignNo
        ElseIf blnOldSection = True And _
                  (strClassification = VALUE_SUBPRODUCT Or _
                   strClassification = VALUE_REFERENCE Or _
                   strClassification = VALUE_LAYOUT Or _
                   strClassification = VALUE_2KMOULD) Then
            strFileDataName = strHeadChar & strOfficeCode & strStatusCode & strDesignNo
        ElseIf blnOldSection = True Then
            strFileDataName = strOfficeCode & strStatusCode & strDesignNo
        Else
            strFileDataName = strHeadChar & strOfficeCode & strStatusCode & strDesignNo
        End If

        Dim strLastChar As String
        Dim intIncrement As Integer
        Dim strTempName As String
        Dim blnSameName As Boolean
        
        If Len(strEndChar) = 3 Then
            strLastChar = Right(strEndChar, 1)
            If Len(strRevisionNo) = 4 Then
                intIncrement = 0
            ElseIf strRevisionNo = "" Or _
               strClassification = VALUE_SUBPRODUCT Or _
               strClassification = VALUE_REFERENCE Or _
               strClassification = VALUE_LAYOUT Then
                intIncrement = 0
            Else
                intIncrement = 1
            End If
            
            Do
                strTempName = ""
                blnSameName = False
                If intIncrement < 10 Then
                    strTempName = strFileDataName & "-0" & intIncrement & strLastChar
                Else
                    strTempName = strFileDataName & "-" & intIncrement & strLastChar
                End If
                
                Dim j As Long
                For j = 1 To lngRecCnt
                
                    If j = i Then
                        GoTo CONTINUE2
                    ElseIf i < j And Trim(mRecords(j).Sel) = True Then
                        GoTo CONTINUE2
                    End If
                    
                    Dim strFileDataType As String
                    strFileDataType = mRecords(i).Properties(lngIndex_FileDataType)
                    If mRecords(j).Properties(lngIndex_FileDataName) & mRecords(j).Properties(lngIndex_FileDataType) = _
                       strTempName & strFileDataType Then
                        blnSameName = True
                        intIncrement = intIncrement + 1
                        Exit For
                    End If
CONTINUE2:
                Next j
                
                If blnSameName = False Then
                    strFileDataName = strTempName
                    Exit Do
                End If
                
            Loop While True
        
        ElseIf Trim(strEndChar) <> "" Then
            strFileDataName = strFileDataName & "-" & strEndChar
        End If
        
        mRecords(i).Properties(lngIndex_FileDataName) = strFileDataName

CONTINUE:
        
        If Trim(modSetting.gstrAutoInput) <> "0" And blnSetFileName = True And Trim(mRecords(i).Properties(lngIndex_DesignNo)) = "" Then
            
            Dim strLoadFileName As String
            strLoadFileName = mRecords(i).fileName
            mRecords(i).Properties(lngIndex_FileDataName) = strLoadFileName
            
        End If
        
    Next i
    
    For i = 1 To lngRecCnt
        
        If mRecords(i).IsChildInstance = True Then
            GoTo CONTINUE3
        End If
        
        If Trim(mRecords(i).Sel) <> True Then
            GoTo CONTINUE3
        End If

        Dim strName As String
        Dim strType As String
        Dim strLastChar2 As String
        Dim strClassification2 As String
        Dim strRevisionNo2 As String
        strName = mRecords(i).Properties(lngIndex_FileDataName)
        strType = mRecords(i).Properties(lngIndex_FileDataType)
        strClassification2 = mRecords(i).Properties(lngIndex_Classification)
        strRevisionNo2 = mRecords(i).Properties(lngIndex_RevisionNo)
        If Right(strName, 2) = "-S" Then
            strLastChar2 = "S"
        ElseIf Right(strName, 2) = "-U" Then
            strLastChar2 = "U"
        ElseIf strClassification2 = VALUE_2KMOULD Then
            strLastChar2 = Right(strName, 1)
        Else
            GoTo CONTINUE3
        End If
        
        If Trim(strName) = "" Then
            GoTo CONTINUE3
        End If
        
        Dim lngCnt As Long
        lngCnt = 0
        Dim k As Long
        For k = 1 To lngRecCnt
        
            If k = i Then
                GoTo CONTINUE6
            End If
        
            If strName & strType = mRecords(k).Properties(lngIndex_FileDataName) & mRecords(k).Properties(lngIndex_FileDataType) Then
                lngCnt = lngCnt + 1
            End If
CONTINUE6:
        Next k

        If lngCnt <= 0 Then
            If strClassification2 = VALUE_2KMOULD And Right(strName, 2) = "-S" Then
                mRecords(i).Properties(lngIndex_FileDataName) = Left(strName, Len(strName) - 2) & "-1" & strLastChar2

            End If
            GoTo CONTINUE3
        End If
        
        Dim strHeadFileName As String
        If strClassification2 = VALUE_2KMOULD Then
            strHeadFileName = Left(strName, Len(strName) - 3)
        Else
            strHeadFileName = Left(strName, Len(strName) - 2)
        End If
        
        Dim intIncrement2 As Integer
        intIncrement2 = 1

        For k = 1 To lngRecCnt
            If strName & strType <> mRecords(k).Properties(lngIndex_FileDataName) & mRecords(k).Properties(lngIndex_FileDataType) Then
                GoTo CONTINUE4
            End If
            
            Do
                Dim blnSameName2 As Boolean
                blnSameName2 = False
            
                Dim strNewName As String

                strNewName = strHeadFileName & "-" & intIncrement2 & strLastChar2
                
                Dim l As Long
                For l = 1 To lngRecCnt
                    If k = l Then
                        GoTo CONTINUE5
                    End If
                    
                    If mRecords(l).Properties(lngIndex_FileDataName) & mRecords(l).Properties(lngIndex_FileDataType) = strNewName & strType Then
                        blnSameName2 = True
                        Exit For
                    End If
CONTINUE5:
                Next l

                If blnSameName2 = False Then
                    Exit Do
                End If
                
                intIncrement2 = intIncrement2 + 1
                
            Loop While True
            
            mRecords(k).Properties(lngIndex_FileDataName) = strNewName
            intIncrement2 = intIncrement2 + 1
CONTINUE4:
        Next k
        
CONTINUE3:
    Next i
    
End Sub

Public Sub SetDummyBlank()

    Dim lngCnt As Long
    lngCnt = Me.fncCount()
    
    Dim i As Long
    For i = 1 To lngCnt
        
        Dim lngPropCnt As Long
        On Error Resume Next
        lngPropCnt = UBound(modMain.gcurMainProperty)
        Dim j As Long
        For j = 1 To lngPropCnt
        
            Dim strPropName As String
            strPropName = modMain.gcurMainProperty(j)
            
            Dim strReq As String
            Dim strDataType As String
            strReq = modDefineDrawing.fncGetInputRequired(strPropName)
            strDataType = modDefineDrawing.fncGetDataType(strPropName)
            
            Dim strValue As String
            strValue = mRecords(i).Properties(j)
            
            If strReq = "0" And Trim(strValue) = "" Then
                
                Dim strDummyValue As String
                If strDataType = "0" Then
                    strDummyValue = VALUE_UNSET_STR
                Else
                    strDummyValue = VALUE_UNSET_NUM
                End If
                mRecords(i).Properties(j) = strDummyValue
                    
            End If
        Next j
    
    Next i
    
End Sub

Public Function fncSetPropertyFromDB(ByRef ilstModelID() As String, _
                                     ByRef iobjRecord As ADODB.Recordset) As Boolean

    fncSetPropertyFromDB = False

    On Error Resume Next
    Dim lngReplaceCnt As Long
    lngReplaceCnt = UBound(ilstModelID)
    On Error GoTo 0

    Dim i As Integer
    For i = 1 To lngReplaceCnt
        On Error Resume Next
        Dim strModelID As String
        strModelID = ""
        strModelID = ilstModelID(i)
        On Error GoTo 0
    
        Dim typRecord As modMain.Record
        typRecord.ModelDrawingID = strModelID
        
        Dim lngPropCnt As Long
        lngPropCnt = UBound(modMain.gcurMainProperty)
        
        Dim lstProperties() As String
        ReDim lstProperties(lngPropCnt)
        
        Dim blnNotFound As Boolean
        blnNotFound = True
        
        Dim j As Long
        For j = 1 To lngPropCnt
            Dim strAttrName As String
            strAttrName = modDefineDrawing.fncGetOldDBAttrName(modMain.gcurMainProperty(j))
            If strAttrName = "" Then
                lstProperties(j) = ""
            Else
                If modMain.fncGetAttrVal(iobjRecord, strModelID, strAttrName, lstProperties(j)) = True Then
                    blnNotFound = False
                End If
            End If
            
            If modMain.gcurMainProperty(j) = TITLE_FILEDATATYPE Then
                If StrConv(lstProperties(j), vbUpperCase) = StrConv(CATDRAWING, vbUpperCase) Then
                    lstProperties(j) = CATDRAWING
                ElseIf StrConv(lstProperties(j), vbUpperCase) = StrConv(CATPRODUCT, vbUpperCase) Then
                    lstProperties(j) = CATPRODUCT
                ElseIf StrConv(lstProperties(j), vbUpperCase) = StrConv(CATPART, vbUpperCase) Then
                    lstProperties(j) = CATPART
                End If
            End If
            
            
            If modMain.gcurMainProperty(j) = TITLE_REVISIONNO Then
                Dim strFullDesignNo As String
                strFullDesignNo = lstProperties(j)
                
                Dim strSplit() As String
                strSplit = Split(strFullDesignNo, "-")
                
                On Error Resume Next
                Dim lngSize As Long
                lngSize = UBound(strSplit)
                On Error GoTo 0
                
                If 0 < lngSize Then
                    Dim strRevisionNo As String
                    strRevisionNo = strSplit(1)
                    
                    Dim k As Long
                    For k = 2 To lngSize
                        strRevisionNo = strRevisionNo & "-" & strSplit(k)
                    Next k
                    
                    lstProperties(j) = strRevisionNo
                End If
            End If
            
            If modMain.gcurMainProperty(j) = TITLE_DESIGNNO Then
                Dim strDesignNo As String
                strDesignNo = lstProperties(j)
                
                If 7 = Len(strDesignNo) And _
                   IsNumeric(Left(strDesignNo, 2)) = True Then
                    
                    lstProperties(j) = Right(strDesignNo, 5)
                    
                ElseIf 8 <= Len(strDesignNo) And _
                   IsNumeric(Left(strDesignNo, 2)) = True Then
                    
                    lstProperties(j) = Right(strDesignNo, 6)
                    
                End If
            
            End If
            
        Next j
    
        If blnNotFound = True Then
            
            On Error Resume Next
            Dim lngNotFoundCnt As Long
            lngNotFoundCnt = 0
            lngNotFoundCnt = UBound(gstrNotFoundModelID)
            On Error GoTo 0
            
            ReDim Preserve gstrNotFoundModelID(lngNotFoundCnt + 1)
            gstrNotFoundModelID(lngNotFoundCnt + 1) = strModelID
            
            GoTo CONTINUE
        End If
        
        typRecord.FromDB = True
        typRecord.Properties = lstProperties
        
        If Me.fncReplaceRecord(typRecord) = False Then
            Exit Function
        End If
        
CONTINUE:
    Next i
    fncSetPropertyFromDB = True
End Function

Public Function fncReplaceRecord(ByRef iRecord As modMain.Record) As Boolean

    fncReplaceRecord = False
    
    Dim lngRecCnt As Long
    lngRecCnt = Me.fncCount
    
    Dim i As Long
    For i = 1 To lngRecCnt
    
        If mRecords(i).ModelDrawingID <> iRecord.ModelDrawingID Then
            GoTo CONTINUE
        End If
        
        mRecords(i).FromDB = iRecord.FromDB
        
        Dim lngPropCnt As Long
        On Error Resume Next
        lngPropCnt = UBound(modMain.gcurMainProperty)
        On Error GoTo 0
        
        Dim j As Long
        For j = 1 To lngPropCnt
            On Error Resume Next
            mRecords(i).Properties(j) = iRecord.Properties(j)
            On Error GoTo 0
        Next j
        
        fncReplaceRecord = True
CONTINUE:
    Next i
    
End Function

Public Function fncSetLinkID() As Boolean

    fncSetLinkID = False

    Dim lngRecCnt As Long
    lngRecCnt = Me.fncCount
    
    Dim i As Long
    For i = 1 To lngRecCnt
        
        If mRecords(i).LinkID <> "" Then
            GoTo CONTINUE
        End If
        
        If mRecords(i).LinkTo = "" And mRecords(i).LinkID = "" Then
            mRecords(i).LinkID = "-"
            GoTo CONTINUE
        End If
        
        Dim lngIndex As Long
        lngIndex = Me.fncSearchFromFilePath(mRecords(i).LinkTo)
        
        mRecords(i).LinkID = mRecords(lngIndex).ID
        mRecords(lngIndex).LinkID = mRecords(i).ID
        
CONTINUE:
    Next i
    fncSetLinkID = True
End Function


Public Function fncCountSlected() As Long
    fncCountSlected = 0

    Dim lngRecCnt As Long
    lngRecCnt = Me.fncCount
    
    Dim i As Long
    For i = 1 To lngRecCnt
        If Trim(mRecords(i).Sel) = True Then
            fncCountSlected = fncCountSlected + 1
        End If
    Next i
End Function

Public Sub SortDrawing()
    Dim lngRecCnt As Long
    Dim lngIndex_Type As Long
    lngRecCnt = Me.fncCount
    lngIndex_Type = modMain.fncGetIndex(TITLE_FILEDATATYPE)
    
    Dim blnSorted As Boolean
    Do
    
        blnSorted = False
        
        Dim i As Long
        For i = 1 To lngRecCnt - 1
            
            If mRecords(i + 1).Properties(lngIndex_Type) <> CATDRAWING Then
                Exit For
            End If
            
            If Val(mRecords(i).LinkID) > Val(mRecords(i + 1).LinkID) Then

                blnSorted = True
                Dim buf As Record
                buf = mRecords(i)
                mRecords(i) = mRecords(i + 1)
                mRecords(i + 1) = buf
            End If
        Next i
    Loop While blnSorted = True
End Sub


Public Function fncIs2DNumbered() As Boolean
    fncIs2DNumbered = False

    Dim lngIndex_Type As Long
    Dim lngIndex_DesignNo As Long
    lngIndex_Type = modMain.fncGetIndex(TITLE_FILEDATATYPE)
    lngIndex_DesignNo = modMain.fncGetIndex(TITLE_DESIGNNO)

    Dim lngRecCnt As Long
    lngRecCnt = Me.fncCount

    Dim i As Long
    For i = 1 To lngRecCnt

        Dim strType As String
        strType = mRecords(i).Properties(lngIndex_Type)

        Dim str3DDesignNo As String
        str3DDesignNo = mRecords(i).Properties(lngIndex_DesignNo)
        
        Dim typRecord As Record
        If mRecords(i).Sel = True And mRecords(i).LinkID <> "-" And str3DDesignNo = "" _
                                    And (strType = CATPART Or strType = CATPRODUCT) Then
            
            Dim lngLinkID As Long
            lngLinkID = mRecords(i).LinkID
            
            Dim str2DDesignNo As String
            str2DDesignNo = mRecords(lngLinkID).Properties(lngIndex_DesignNo)
            
            If str2DDesignNo <> "" Then
                fncIs2DNumbered = True
                Exit Function
            End If
        End If
    Next i
End Function

Public Sub CorrectOrName()
    Dim lngRecCnt As Long
    Dim lngIndex_Section As Long
    Dim lngIndex_CurrentStatus As Long
    lngRecCnt = Me.fncCount
    lngIndex_Section = modMain.fncGetIndex(TITLE_SECTION)
    lngIndex_CurrentStatus = modMain.fncGetIndex(TITLE_CURRENTSTATUS)
    
    Dim i As Long
    For i = 1 To lngRecCnt
        Dim strSection As String
        strSection = mRecords(i).Properties(lngIndex_Section)
        If UCase(VALUE_FCS) = UCase(Trim(strSection)) Then
            mRecords(i).Properties(lngIndex_Section) = VALUE_FCS
        ElseIf 0 < InStr(UCase(strSection), UCase(VALUE_EPWR)) Then
            mRecords(i).Properties(lngIndex_Section) = VALUE_FCS
        End If
        
        Dim strCurrentStatus As String
        strCurrentStatus = mRecords(i).Properties(lngIndex_CurrentStatus)
        If UCase(VALUE_PROTOTYPE) = UCase(Trim(strCurrentStatus)) Or _
           UCase(VALUE_STUDY) = UCase(Trim(strCurrentStatus)) Then
            mRecords(i).Properties(lngIndex_CurrentStatus) = VALUE_PROTOTYPESTUDY
        End If
    Next i
End Sub