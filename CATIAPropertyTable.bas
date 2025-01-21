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
    
    If lngCnt < ilngIndex Then Exit Function
    otypRecord = mRecords(ilngIndex)
    
    fncItem = True
End Function

Public Function fncAddRecord(ByRef itypRecord As Record) As Long
    fncAddRecord = -1

    Dim lngCnt As Long
    lngCnt = Me.fncCount()
    
    Dim strType As String
    strType = itypRecord.Properties(modMain.fncGetIndex("File_Data_Type") - 1)
    
    If strType = "Component" Then
        fncAddRecord = 0
        Exit Function
    Else
        Dim I As Integer
        For I = 1 To lngCnt
            If mRecords(I).FilePath = itypRecord.FilePath Then
                mRecords(I).Amount = mRecords(I).Amount + 1
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
        Next I

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
    
    Dim I As Long
    For I = 1 To lngDataSize
        Dim typData As Record
        Dim typMyData As Record
        If iobjData.fncItem(I, typData) = False Then
            Exit Function
        ElseIf Me.fncItem(I, typMyData) = False Then
            Exit Function
        End If
        
        If typData.FilePath <> typMyData.FilePath Then
            fncIsSameStructure = True
            oblnIsSame = False
        End If
    Next I

    fncIsSameStructure = True
    oblnIsSame = True
End Function

Public Function fncGetLastLevel() As Long
    fncGetLastLevel = 0

    Dim lngCnt As Long
    lngCnt = Me.fncCount()
    
    Dim I As Integer
    For I = 1 To lngCnt
        Dim typRecord As Record
        typRecord = mRecords(I)
        If fncGetLastLevel < typRecord.Level Then fncGetLastLevel = typRecord.Level
    Next I
End Function

Public Function fncCheckBlank(ByRef ostrPropertyName As String) As String
    fncCheckBlank = ""

    If modMain.gstrInputCheck <> "1" Then
        Exit Function
    End If

    Dim lngCnt As Long
    lngCnt = Me.fncCount()
    
    Dim blnNoCheckFlag As Boolean
    blnNoCheckFlag = False
    
    Dim I As Long
    For I = 1 To lngCnt
    
        Dim typRecord As Record
        If Me.fncItem(I, typRecord) = False Then
            GoTo continue
        End If
        
        Dim lngValIndex As Long
        Dim strValue As String
    
        lngValIndex = modMain.fncGetIndex("Classification") - 1
        strValue = typRecord.Properties(lngValIndex)
        
        If typRecord.Level <= 1 Then
            If strValue = "2K mould" Or _
               strValue = "SubProduct" Or _
               strValue = "Reference" Or _
               strValue = "LayOut" Or _
               strValue = "Customer approved data" Then
                blnNoCheckFlag = True
            Else
                blnNoCheckFlag = False
            End If
        End If
        
        If blnNoCheckFlag = True Then
            GoTo continue
        End If
        
        Dim strSel As String
        strSel = typRecord.Sel
        If Trim(strSel) <> True Then
            GoTo continue
        End If
        
        lngValIndex = modMain.fncGetIndex("Design_No") - 1
        strValue = typRecord.Properties(lngValIndex)
        If Trim(strValue) = "" Then
            fncCheckBlank = "E034"
            Exit Function
        End If

        lngValIndex = modMain.fncGetIndex("Current_Status") - 1
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
            strReq = getPLMpropRequired(strPropName)
            
            If strReq = "1" Then
                strValue = typRecord.Properties(j)
                If Trim(strValue) = "" Then
                    fncCheckBlank = "E038"
                    ostrPropertyName = strPropName
                    Exit Function
                End If
            End If
            
        Next j
        
continue:
    Next I
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
    
    Dim I As Long
    For I = 1 To lngRecCnt
        If mRecords(I).FilePath = istrFilePath Then
            fncSearchFromFilePath = I
            Exit For
        End If
    Next I
End Function

Public Sub UpdatePath(ByVal istrOldPath As String, ByVal istrNewPath As String, ByVal istrNewName As String)
    Dim lngRecCnt As Long
    lngRecCnt = Me.fncCount
    
    Dim I As Long
    For I = 1 To lngRecCnt
        If mRecords(I).FilePath = istrOldPath Then
            mRecords(I).FilePath = istrNewPath
            mRecords(I).FileName = istrNewName
        End If
        If mRecords(I).LinkTo = istrOldPath Then
            mRecords(I).LinkTo = istrNewPath
        End If
    Next I
End Sub

Public Sub UpdateModelID(ByVal ilngIndex As Long, ByVal istrValue As String)
    Dim lngRecCnt As Long
    lngRecCnt = Me.fncCount
    
    If 0 < ilngIndex And ilngIndex <= lngRecCnt Then mRecords(ilngIndex).ModelDrawingID = istrValue
    
End Sub

Public Function fncReplaceProhibitCharacter() As Boolean
    fncReplaceProhibitCharacter = False

    Dim lngIndex_FileDataName As Long
    Dim lngIndex_FullDesignNo As Long
    lngIndex_FileDataName = modMain.fncGetIndex("File_Data_Name")
    lngIndex_FullDesignNo = modMain.fncGetIndex("Full_Design_No")

    Dim lngRecCnt As Long
    lngRecCnt = Me.fncCount
    
    Dim I As Long
    For I = 1 To lngRecCnt
    
        Dim strSel As String
        strSel = mRecords(I).Sel
        If Trim(strSel) <> True Then
            GoTo continue
        End If
        
        Dim strBuf As String
        strBuf = mRecords(I).Properties(lngIndex_FileDataName)
        Call ReplaceString(strBuf)
        If strBuf <> mRecords(I).Properties(lngIndex_FileDataName) Then
            mRecords(I).Properties(lngIndex_FileDataName) = strBuf
            fncReplaceProhibitCharacter = True
        End If
        
        strBuf = mRecords(I).Properties(lngIndex_FullDesignNo)
        Call ReplaceString(strBuf)
        If strBuf <> mRecords(I).Properties(lngIndex_FullDesignNo) Then
            mRecords(I).Properties(lngIndex_FullDesignNo) = strBuf
            fncReplaceProhibitCharacter = True
        End If
continue:
    Next I
End Function

Public Sub ClearModelID()
    Dim lngRecCnt As Long
    lngRecCnt = Me.fncCount
    
    Dim I As Long
    For I = 1 To lngRecCnt
        Dim strSel As String
        strSel = mRecords(I).Sel
        If Trim(strSel) <> True Then GoTo continue
        mRecords(I).ModelDrawingID = ""
continue:
    Next I
End Sub

Private Sub ReplaceString(ByRef ostrString As String)
    On Error Resume Next
    Dim lngSize As Long
    lngSize = UBound(modMain.glstProhibitCharacter)
    On Error GoTo 0

    Dim I As Long
    For I = 1 To lngSize
        Dim strFind As String
        strFind = modMain.glstProhibitCharacter(I)
        ostrString = Replace(ostrString, strFind, " ")
    Next I
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
    lngIndex_Section = modMain.fncGetIndex("Section") - 1
    lngIndex_Status = modMain.fncGetIndex("Current_Status") - 1
    lngIndex_DesignNo = modMain.fncGetIndex("Design_No") - 1
    lngIndex_RevisionNo = modMain.fncGetIndex("Revision_No") - 1
    lngIndex_FileDataName = modMain.fncGetIndex("File_Data_Name") - 1
    lngIndex_FullDesignNo = modMain.fncGetIndex("Full_Design_No") - 1
    lngIndex_Classification = modMain.fncGetIndex("Classification") - 1
    lngIndex_FileName = modMain.fncGetIndex("FileName") - 1
    lngIndex_FileDataType = modMain.fncGetIndex("File_Data_Type") - 1
    lngIndex_MaterialGrade = modMain.fncGetIndex("Material_Grade") - 1
    
    Dim blnNoCheckFlag As Boolean
    blnNoCheckFlag = False
    Dim blnSetFileName As Boolean
    blnSetFileName = False

    Dim lngRecCnt As Long
    lngRecCnt = Me.fncCount
    
    Dim I As Long
    For I = 1 To lngRecCnt
        If mRecords(I).IsChildInstance = True Then GoTo continue
    
        Dim strDesigner As String
        strDesigner = mRecords(I).Properties(lngIndex_Designer)
        If Trim(strDesigner) = "" Then
            mRecords(I).Properties(lngIndex_Designer) = modMain.gstrNFDesigner
            strDesigner = modMain.gstrNFDesigner
        End If
        
        Dim strSection As String
        strSection = mRecords(I).Properties(lngIndex_Section)
        If Trim(strSection) = "" Then
            mRecords(I).Properties(lngIndex_Section) = "NAM"
            strSection = "NAM"
        End If
        
        Dim strClassification As String
        strClassification = mRecords(I).Properties(lngIndex_Classification)
        
        If modMain.gstrInputCheck <> "1" Then
            blnNoCheckFlag = True
            blnSetFileName = True
        ElseIf mRecords(I).Level <= 1 Then
            If strClassification = "2K mould (CATPart)" Or _
               strClassification = "SubProduct" Or _
               strClassification = "Reference" Or _
               strClassification = "LayOut" Or _
               strClassification = "Customer approved data" Then
                blnNoCheckFlag = True
            Else
                blnNoCheckFlag = False
            End If
            
            If strClassification = "Customer approved data" Then
                blnSetFileName = True
            Else
                blnSetFileName = False
            End If
        End If
        
        Dim strSel As String
        strSel = mRecords(I).Sel
        If Trim(strSel) <> True Then GoTo continue
        If Trim(modMain.gstrUnsetMaterialGrade) = "1" And strClassification = "Submission data" Then mRecords(I).Properties(lngIndex_MaterialGrade) = "Unset"
        If Trim(modMain.gstrAutoInput) = "0" Then GoTo continue
        
        Dim strDesignNo As String
        strDesignNo = Trim(mRecords(I).Properties(lngIndex_DesignNo))
        If blnNoCheckFlag = False And Trim(strDesignNo) = "" Then
            mRecords(I).Properties(lngIndex_FileDataName) = ""
            mRecords(I).Properties(lngIndex_FullDesignNo) = ""
            GoTo continue
        End If
        
        If blnNoCheckFlag = False And Trim(strSection) = "" Then
            gBlankSection = True
            mRecords(I).Properties(lngIndex_FileDataName) = ""
            mRecords(I).Properties(lngIndex_FullDesignNo) = ""
            GoTo continue
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
            strOfficeCode = getSectionData("Office_Code", "Section", strSection)
            If blnNoCheckFlag = False And Trim(strOfficeCode) = "" Then
                On Error Resume Next
                Dim lngSectCnt As Long
                lngSectCnt = UBound(gUnknownSection)
                ReDim Preserve gUnknownSection(lngSectCnt + 1) As String
                gUnknownSection(lngSectCnt + 1) = strSection
                On Error GoTo 0
                mRecords(I).Properties(lngIndex_FileDataName) = ""
                mRecords(I).Properties(lngIndex_FullDesignNo) = ""
                GoTo continue
            End If
            If IsNumeric(strOfficeCode) = True Then
                blnOldSection = False
            Else
                blnOldSection = True
            End If
        End If

        Dim strStatus As String
        strStatus = mRecords(I).Properties(lngIndex_Status)
        If blnNoCheckFlag = False And Trim(strStatus) = "" Then
            gBlankStatus = True
            mRecords(I).Properties(lngIndex_FileDataName) = ""
            mRecords(I).Properties(lngIndex_FullDesignNo) = ""
            GoTo continue
        End If
        
        If blnNoCheckFlag = False And (strStatus <> "Mass production" And strStatus <> "Prototype/Study") Then
            On Error Resume Next
            Dim lngStatusCnt As Long
            lngStatusCnt = UBound(gUnknownStatus)
            ReDim Preserve gUnknownStatus(lngStatusCnt + 1) As String
            gUnknownStatus(lngStatusCnt + 1) = strStatus
            On Error GoTo 0
            mRecords(I).Properties(lngIndex_FileDataName) = ""
            mRecords(I).Properties(lngIndex_FullDesignNo) = ""
            GoTo continue
        End If
        
        Dim strStatusCode As String
        If strClassification = "Customer approved data" Then
            strStatusCode = "C"
        Else
            If blnOldSection = False Then
                If strStatus = "Mass production" Then
                    strStatusCode = "M"
                ElseIf strStatus = "Prototype/Study" Then
                    strStatusCode = "T"
                End If
            Else
                If strStatus = "Prototype/Study" Then
                    strStatusCode = "T"
                Else
                    strStatusCode = ""
                End If
            End If
        End If
        Dim strRevisionNo As String
        strRevisionNo = mRecords(I).Properties(lngIndex_RevisionNo)
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
        strOldFullDesignNo = mRecords(I).Properties(lngIndex_FullDesignNo)
        
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

        mRecords(I).Properties(lngIndex_FullDesignNo) = strFullDesignNo
        
        Dim strEndChar As String
        strEndChar = ""

        If strClassification = "2K mould (CATPart)" Or _
           strClassification = "Reference" Or _
           strClassification = "SubProduct" Then
            strEndChar = "S"
        ElseIf strClassification = "LayOut" Then
            strEndChar = "U"
        ElseIf 2 <= Len(strRevisionNo) And IsNumeric(Right(strRevisionNo, 2)) = True Then
            Dim lngRevNo As Long
            lngRevNo = Right(strRevisionNo, 2)
            If 80 <= lngRevNo Then
                strEndChar = "U"
            Else
                strEndChar = "S"
            End If
        ElseIf strClassification = "Submission data" Then
            strEndChar = "U"
        Else
            strEndChar = "S"
        End If
        
        If strClassification = "SubProduct" Or _
           strClassification = "Reference" Or _
           strClassification = "LayOut" Then
            strEndChar = "00" & strEndChar
        ElseIf strClassification = "2K mould (CATPart)" Then
            strEndChar = strEndChar
        ElseIf Len(strRevisionNo) = 2 Then
        ElseIf Len(strRevisionNo) = 4 Then
            strEndChar = Left(strRevisionNo, 2) & strEndChar
        Else
            strEndChar = strRevisionNo & strEndChar
        End If
        
        Dim strHeadChar As String
        If strClassification = "SubProduct" Then
            strHeadChar = "S"
        ElseIf strClassification = "Reference" Then
            strHeadChar = "J"
        ElseIf strClassification = "LayOut" Then
            strHeadChar = "L"
        ElseIf strClassification = "2K mould (CATPart)" Then
            strHeadChar = "W"
        Else
            strHeadChar = ""
        End If
        
        Dim strFileDataName As String
        
        If blnOldSection = True And strStatusCode = "C" Then
            strFileDataName = strStatusCode & strOfficeCode & strDesignNo
        ElseIf blnOldSection = True And _
                  (strClassification = "SubProduct" Or _
                   strClassification = "Reference" Or _
                   strClassification = "LayOut" Or _
                   strClassification = "2K mould (CATPart)") Then
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
               strClassification = "SubProduct" Or _
               strClassification = "Reference" Or _
               strClassification = "LayOut" Then
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
                
                    If j = I Then
                        GoTo CONTINUE2
                    ElseIf I < j And Trim(mRecords(j).Sel) = True Then
                        GoTo CONTINUE2
                    End If
                    
                    Dim strFileDataType As String
                    strFileDataType = mRecords(I).Properties(lngIndex_FileDataType)
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
        
        mRecords(I).Properties(lngIndex_FileDataName) = strFileDataName

continue:
        
        If Trim(modMain.gstrAutoInput) <> "0" And blnSetFileName = True And Trim(mRecords(I).Properties(lngIndex_DesignNo)) = "" Then
            
            Dim strLoadFileName As String
            strLoadFileName = mRecords(I).FileName
            mRecords(I).Properties(lngIndex_FileDataName) = strLoadFileName
            
        End If
        
    Next I
    
    For I = 1 To lngRecCnt
        
        If mRecords(I).IsChildInstance = True Then
            GoTo CONTINUE3
        End If
        
        If Trim(mRecords(I).Sel) <> True Then
            GoTo CONTINUE3
        End If

        Dim strName As String
        Dim strType As String
        Dim strLastChar2 As String
        Dim strClassification2 As String
        Dim strRevisionNo2 As String
        strName = mRecords(I).Properties(lngIndex_FileDataName)
        strType = mRecords(I).Properties(lngIndex_FileDataType)
        strClassification2 = mRecords(I).Properties(lngIndex_Classification)
        strRevisionNo2 = mRecords(I).Properties(lngIndex_RevisionNo)
        If Right(strName, 2) = "-S" Then
            strLastChar2 = "S"
        ElseIf Right(strName, 2) = "-U" Then
            strLastChar2 = "U"
        ElseIf strClassification2 = "2K mould (CATPart)" Then
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
        
            If k = I Then
                GoTo CONTINUE6
            End If
        
            If strName & strType = mRecords(k).Properties(lngIndex_FileDataName) & mRecords(k).Properties(lngIndex_FileDataType) Then
                lngCnt = lngCnt + 1
            End If
CONTINUE6:
        Next k

        If lngCnt <= 0 Then
            If strClassification2 = "2K mould (CATPart)" And Right(strName, 2) = "-S" Then
                mRecords(I).Properties(lngIndex_FileDataName) = Left(strName, Len(strName) - 2) & "-1" & strLastChar2

            End If
            GoTo CONTINUE3
        End If
        
        Dim strHeadFileName As String
        If strClassification2 = "2K mould (CATPart)" Then
            strHeadFileName = Left(strName, Len(strName) - 3)
        Else
            strHeadFileName = Left(strName, Len(strName) - 2)
        End If
        
        Dim intIncrement2 As Integer
        intIncrement2 = 1

        For k = 1 To lngRecCnt
            If strName & strType <> mRecords(k).Properties(lngIndex_FileDataName) & mRecords(k).Properties(lngIndex_FileDataType) Then GoTo CONTINUE4
            
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

                If blnSameName2 = False Then Exit Do
                
                intIncrement2 = intIncrement2 + 1
            Loop While True
            
            mRecords(k).Properties(lngIndex_FileDataName) = strNewName
            intIncrement2 = intIncrement2 + 1
CONTINUE4:
        Next k
        
CONTINUE3:
    Next I
    
End Sub

Public Sub SetDummyBlank()

    Dim lngCnt As Long
    lngCnt = Me.fncCount()
    
    Dim I As Long
    For I = 1 To lngCnt
        
        Dim lngPropCnt As Long
        On Error Resume Next
        lngPropCnt = UBound(modMain.gcurMainProperty)
        Dim j As Long
        For j = 1 To lngPropCnt
        
            Dim strPropName As String
            strPropName = modMain.gcurMainProperty(j)
            
            Dim strReq As String
            Dim strDataType As String
            strReq = getPLMpropRequired(strPropName)
            strDataType = 0
            
            Dim strValue As String
            strValue = mRecords(I).Properties(j)
            
            If strReq = "0" And Trim(strValue) = "" Then
                
                Dim strDummyValue As String
                If strDataType = "0" Then
                    strDummyValue = "Unset"
                Else
                    strDummyValue = "999"
                End If
                mRecords(I).Properties(j) = strDummyValue
                    
            End If
        Next j
    
    Next I
    
End Sub

Public Function fncSetPropertyFromDB(ByRef ilstModelID() As String, _
                                     ByRef iobjRecord As ADODB.Recordset) As Boolean

    fncSetPropertyFromDB = False

    On Error Resume Next
    Dim lngReplaceCnt As Long
    lngReplaceCnt = UBound(ilstModelID)
    On Error GoTo 0

    Dim I As Integer
    For I = 1 To lngReplaceCnt
        On Error Resume Next
        Dim strModelID As String
        strModelID = ""
        strModelID = ilstModelID(I)
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
            strAttrName = ""
            If strAttrName = "" Then
                lstProperties(j) = ""
            Else
                If modMain.fncGetAttrVal(iobjRecord, strModelID, strAttrName, lstProperties(j)) = True Then
                    blnNotFound = False
                End If
            End If
            
            If modMain.gcurMainProperty(j) = "File_Data_Type" Then
                If StrConv(lstProperties(j), vbUpperCase) = StrConv("CATDrawing", vbUpperCase) Then
                    lstProperties(j) = "CATDrawing"
                ElseIf StrConv(lstProperties(j), vbUpperCase) = StrConv("CATProduct", vbUpperCase) Then
                    lstProperties(j) = "CATProduct"
                ElseIf StrConv(lstProperties(j), vbUpperCase) = StrConv("CATPart", vbUpperCase) Then
                    lstProperties(j) = "CATPart"
                End If
            End If
            
            
            If modMain.gcurMainProperty(j) = "Revision_No" Then
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
            
            If modMain.gcurMainProperty(j) = "Design_No" Then
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
            
            GoTo continue
        End If
        
        typRecord.FromDB = True
        typRecord.Properties = lstProperties
        
        If Me.fncReplaceRecord(typRecord) = False Then Exit Function
        
continue:
    Next I
    fncSetPropertyFromDB = True
End Function

Public Function fncReplaceRecord(ByRef iRecord As modMain.Record) As Boolean

    fncReplaceRecord = False
    
    Dim lngRecCnt As Long
    lngRecCnt = Me.fncCount
    
    Dim I As Long
    For I = 1 To lngRecCnt
        If mRecords(I).ModelDrawingID <> iRecord.ModelDrawingID Then GoTo continue
        
        mRecords(I).FromDB = iRecord.FromDB
        
        Dim lngPropCnt As Long
        On Error Resume Next
        lngPropCnt = UBound(modMain.gcurMainProperty)
        On Error GoTo 0
        
        Dim j As Long
        For j = 1 To lngPropCnt
            On Error Resume Next
            mRecords(I).Properties(j) = iRecord.Properties(j)
            On Error GoTo 0
        Next j
        
        fncReplaceRecord = True
continue:
    Next I
    
End Function

Public Function fncSetLinkID() As Boolean
    fncSetLinkID = False

    Dim lngRecCnt As Long
    lngRecCnt = Me.fncCount
    
    Dim I As Long
    For I = 1 To lngRecCnt
        If mRecords(I).LinkID <> "" Then GoTo continue
        If mRecords(I).LinkTo = "" And mRecords(I).LinkID = "" Then
            mRecords(I).LinkID = "-"
            GoTo continue
        End If
        
        Dim lngIndex As Long
        lngIndex = Me.fncSearchFromFilePath(mRecords(I).LinkTo)
        
        mRecords(I).LinkID = mRecords(lngIndex).ID
        mRecords(lngIndex).LinkID = mRecords(I).ID
continue:
    Next I
    fncSetLinkID = True
End Function

Public Function fncCountSlected() As Long
    fncCountSlected = 0

    Dim lngRecCnt As Long
    lngRecCnt = Me.fncCount
    
    Dim I As Long
    For I = 1 To lngRecCnt
        If Trim(mRecords(I).Sel) = True Then
            fncCountSlected = fncCountSlected + 1
        End If
    Next I
End Function

Public Sub SortDrawing()
    Dim lngRecCnt As Long
    Dim lngIndex_Type As Long
    lngRecCnt = Me.fncCount
    lngIndex_Type = modMain.fncGetIndex("File_Data_Type")
    
    Dim blnSorted As Boolean
    Do
        blnSorted = False
        
        Dim I As Long
        For I = 1 To lngRecCnt - 1
            If mRecords(I + 1).Properties(lngIndex_Type) <> "CATDrawing" Then Exit For
            
            If val(mRecords(I).LinkID) > val(mRecords(I + 1).LinkID) Then
                blnSorted = True
                Dim buf As Record
                buf = mRecords(I)
                mRecords(I) = mRecords(I + 1)
                mRecords(I + 1) = buf
            End If
        Next I
    Loop While blnSorted = True
End Sub


Public Function fncIs2DNumbered() As Boolean
    fncIs2DNumbered = False

    Dim lngIndex_Type As Long
    Dim lngIndex_DesignNo As Long
    lngIndex_Type = modMain.fncGetIndex("File_Data_Type")
    lngIndex_DesignNo = modMain.fncGetIndex("Design_No")

    Dim lngRecCnt As Long
    lngRecCnt = Me.fncCount

    Dim I As Long
    For I = 1 To lngRecCnt

        Dim strType As String
        strType = mRecords(I).Properties(lngIndex_Type)

        Dim str3DDesignNo As String
        str3DDesignNo = mRecords(I).Properties(lngIndex_DesignNo)
        
        Dim typRecord As Record
        If mRecords(I).Sel = True And mRecords(I).LinkID <> "-" And str3DDesignNo = "" _
                                    And (strType = "CATPart" Or strType = "CATProduct") Then
            
            Dim lngLinkID As Long
            lngLinkID = mRecords(I).LinkID
            
            Dim str2DDesignNo As String
            str2DDesignNo = mRecords(lngLinkID).Properties(lngIndex_DesignNo)
            
            If str2DDesignNo <> "" Then
                fncIs2DNumbered = True
                Exit Function
            End If
        End If
    Next I
End Function

Public Sub CorrectOrName()
    Dim lngRecCnt As Long
    Dim lngIndex_Section As Long
    Dim lngIndex_CurrentStatus As Long
    lngRecCnt = Me.fncCount
    lngIndex_Section = modMain.fncGetIndex("Section")
    lngIndex_CurrentStatus = modMain.fncGetIndex("Current_Status")
    
    Dim I As Long
    For I = 1 To lngRecCnt
        Dim strSection As String
        strSection = mRecords(I).Properties(lngIndex_Section)
        If UCase("FCS") = UCase(Trim(strSection)) Then
            mRecords(I).Properties(lngIndex_Section) = "FCS"
        ElseIf 0 < InStr(UCase(strSection), UCase("PWR")) Then
            mRecords(I).Properties(lngIndex_Section) = "FCS"
        End If
        
        Dim strCurrentStatus As String
        strCurrentStatus = mRecords(I).Properties(lngIndex_CurrentStatus)
        If UCase("Prototype") = UCase(Trim(strCurrentStatus)) Or _
           UCase("Study") = UCase(Trim(strCurrentStatus)) Then
            mRecords(I).Properties(lngIndex_CurrentStatus) = "Prototype/Study"
        End If
    Next I
End Sub