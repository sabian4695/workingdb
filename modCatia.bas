Option Explicit

Private mobjCATIA As Object
Private mlstDrawDoc() As Object
Private mlstProdDoc() As Object
Public gstrNotFoundText() As String
Public gstrNoLinkDoc() As String
Public gstrNotFoundModelID() As String
Public gstrNotSavedFile() As String
Dim rsPLMprops As Recordset

Public Function fncInit()
    fncInit = False
    Set mobjCATIA = Nothing
    ReDim mlstDrawDoc(0)
    ReDim mlstProdDoc(0)
    ReDim gstrNotFoundText(0)
    ReDim gstrNoLinkDoc(0)
    ReDim gstrNotFoundModelID(0)
    
    Set rsPLMprops = CurrentDb().OpenRecordset("tblPLMproperties")

    On Error Resume Next
    Set mobjCATIA = GetObject(, "CATIA.Application")
    On Error GoTo 0

    If mobjCATIA Is Nothing Then
        Call snackBox("error", "Uh oh!", "Catia isn't running!", "frmPLM")
        Exit Function
    End If

    If mobjCATIA.Windows.count = 0 Then
        Call snackBox("error", "Uh oh!", "No Catia files found", "frmPLM")
        Exit Function
    End If
    
    Dim lngDrawCnt As Long, lngProdCnt As Long, i As Long, objDoc As Object
    For i = 1 To mobjCATIA.Windows.count
        On Error Resume Next
        Set objDoc = mobjCATIA.Windows.ITEM(i).Parent
        On Error GoTo 0
        If Not objDoc Is Nothing Then
            If TypeName(objDoc) = "DrawingDocument" Then
                lngDrawCnt = lngDrawCnt + 1
                ReDim Preserve mlstDrawDoc(lngDrawCnt)
                Set mlstDrawDoc(lngDrawCnt) = objDoc
            ElseIf TypeName(objDoc) = "ProductDocument" Or _
                   TypeName(objDoc) = "PartDocument" Then
                lngProdCnt = lngProdCnt + 1
                ReDim Preserve mlstProdDoc(lngProdCnt)
                Set mlstProdDoc(lngProdCnt) = objDoc
            End If
        End If
    Next i
    fncInit = True
End Function

Public Sub Terminate()
    Set mobjCATIA = Nothing
    ReDim mlstDrawDoc(0)
    ReDim mlstProdDoc(0)
    ReDim gstrNotFoundText(0)
    ReDim gstrNoLinkDoc(0)
    ReDim gstrNotFoundModelID(0)
    ReDim gstrNotSavedFile(0)
    On Error Resume Next
    rsPLMprops.Close
    Set rsPLMprops = Nothing
End Sub

Public Function fncGetProperty(Optional ByVal iblnLoad2dText As Boolean = True) As CATIAPropertyTable
    Set fncGetProperty = New CATIAPropertyTable
    
    Dim lngParentIndex As Long
    lngParentIndex = 0
    
    On Error Resume Next
    Dim lngCnt As Long
    lngCnt = UBound(mlstDrawDoc)
    On Error GoTo 0
    
    Dim i As Long
    For i = 1 To lngCnt
        If Not mlstDrawDoc(i) Is Nothing Then
            If fncGetPropertyFromDrawing(mlstDrawDoc(i), lngParentIndex, fncGetProperty, iblnLoad2dText) = False Then
                Call modMessage.Show("E006", mlstDrawDoc(i).name)
                Set fncGetProperty = Nothing
                Exit Function
            End If
            lngParentIndex = lngParentIndex + 1
        End If
    Next i
    
    On Error Resume Next
    lngCnt = 0
    lngCnt = UBound(mlstProdDoc)
    On Error GoTo 0
    
    For i = 1 To lngCnt
        Dim objRootProd As Object
        Set objRootProd = mlstProdDoc(i).Product
        If objRootProd Is Nothing Then
            Call modMessage.Show("E006")
            Set fncGetProperty = Nothing
        End If

        Call objRootProd.ApplyWorkMode(2)

        If fncGetPropertyFromProduct(objRootProd, 1, lngParentIndex, fncGetProperty) = False Then
            Call modMessage.Show("E006")
            Set fncGetProperty = Nothing
            Exit Function
        End If
        
    Next i

    Dim strErrID As String
    strErrID = fncGetOldNumberingProperty(fncGetProperty)
    If strErrID = "E036" Then
        Call modMessage.Show2(strErrID, gstrNotFoundModelID)
    ElseIf strErrID <> "" Then
        Call modMessage.Show(strErrID)
        Set fncGetProperty = Nothing
        Exit Function
    End If
    
    If fncGetProperty.fncSetLinkID() = False Then
        Call modMessage.Show("E006")
        Set fncGetProperty = Nothing
        Exit Function
    End If
    
    Call fncGetProperty.SortDrawing
    Call fncGetProperty.CorrectOrName
End Function

Private Function fncGetDrawingLink(ByRef iobjDrawDoc As Object) As String
    fncGetDrawingLink = ""

    If iobjDrawDoc Is Nothing Then
        Exit Function
    End If

    Dim i As Long
    For i = 1 To iobjDrawDoc.Sheets.count
        Dim objSheet As Object
        Set objSheet = iobjDrawDoc.Sheets.ITEM(i)
        Dim j As Long
        For j = 1 To objSheet.Views.count
            Dim objView As Object
            Set objView = objSheet.Views.ITEM(j)
            On Error Resume Next
            Dim objLinkDoc As Object
            Set objLinkDoc = objView.GenerativeBehavior.Document
            On Error GoTo 0
            If objLinkDoc Is Nothing Then GoTo continue
            
            If TypeName(objLinkDoc) = "Product" Then
                fncGetDrawingLink = objLinkDoc.ReferenceProduct.Parent.fullName
                Exit Function
            End If
            
            Dim objParent As Object
            Set objParent = objLinkDoc
            Do While True
                If objParent Is Nothing Then Exit Do
                If TypeName(objParent) = "ProductDocument" Or TypeName(objParent) = "PartDocument" Then Exit Do
                
                On Error Resume Next
                Dim objTemp As Object
                Set objTemp = Nothing
                Set objTemp = objParent.Parent
                On Error GoTo 0
                
                If objTemp Is Nothing Then Exit Do
                Set objParent = objTemp
            Loop
            If TypeName(objParent) = "ProductDocument" Or TypeName(objParent) = "PartDocument" Then
                fncGetDrawingLink = objParent.fullName
                Exit Function
            End If
continue:
        Next j
    Next i
End Function

Private Function fncIsInTree(ByRef iobjProduct As Object, ByVal istrFullPath As String) As Boolean
    fncIsInTree = False
    
    If iobjProduct Is Nothing Then Exit Function

    On Error Resume Next
    Dim strFullName As String
    strFullName = iobjProduct.ReferenceProduct.Parent.fullName
    On Error GoTo 0
    If strFullName = istrFullPath Then
        fncIsInTree = True
        Exit Function
    End If
    
    Dim i As Long
    For i = 1 To iobjProduct.Products.count
        Dim objChildProd As Object
        Set objChildProd = iobjProduct.Products.ITEM(i)
        If fncIsInTree(objChildProd, istrFullPath) = True Then
            fncIsInTree = True
            Exit Function
        End If
    Next i
End Function

Private Function fncGetOldNumberingProperty(ByRef iobjTable As CATIAPropertyTable) As String
    fncGetOldNumberingProperty = ""

    Dim lngCnt As Long
    lngCnt = iobjTable.fncCount
    Dim lstModelID() As String
    ReDim lstModelID(0)
    Dim i As Long
    For i = 1 To lngCnt
        Dim typRecord As modMain.Record
        If iobjTable.fncItem(i, typRecord) = False Then
            fncGetOldNumberingProperty = "E035"
            Exit Function
        End If
        
        If typRecord.IsChildInstance = True Then GoTo continue
        If Trim(typRecord.ModelDrawingID) = "" Then GoTo continue
        
        Dim lngTypeIndex As Long
        lngTypeIndex = modMain.fncGetIndex("File_Data_Type")
        Dim strType As String
        strType = typRecord.Properties(lngTypeIndex)

        If strType = "CATProduct" Or strType = "CATPart" Then
            On Error Resume Next
            Dim lngIndex As Long
            lngIndex = UBound(lstModelID)
            On Error GoTo 0
            
            ReDim Preserve lstModelID(lngIndex + 1)
            lstModelID(lngIndex + 1) = typRecord.ModelDrawingID
        End If
continue:
    Next i
    
    On Error Resume Next
    lngCnt = 0
    lngCnt = UBound(lstModelID)
    On Error GoTo 0

    If 0 < lngCnt Then
        ReDim gstrNotFoundModelID(0)
        If modMain.fncGetPropertyFromDB(lstModelID, iobjTable) = False Then
            fncGetOldNumberingProperty = "E035"
            Exit Function
        End If
        
        On Error Resume Next
        Dim lngCntNotFound As Long
        lngCntNotFound = UBound(gstrNotFoundModelID)
        On Error GoTo 0
        If lngCntNotFound > 0 Then fncGetOldNumberingProperty = "E036"
    End If
End Function

Private Function fncGetPropertyFromDrawing(ByRef iobjDrawDoc As Object, _
                                           ByVal ilngParentIndex As Long, _
                                           ByRef oobjCatiaData As CATIAPropertyTable, _
                                           ByVal iblnLoad2dText As Boolean) As Boolean
    fncGetPropertyFromDrawing = False
    If iobjDrawDoc Is Nothing Then Exit Function
    
    Dim strParamName As String, strFilePath As String, strFileName As String, strDrawingID As String, strDrawLinkTo As String, strProperties() As String
    strParamName = ""
    strFilePath = iobjDrawDoc.fullName
    strFileName = fncSplitFileName(iobjDrawDoc.name)
    strParamName = getPLMpropertyData("Drawing_Parameter_Name", "Form_Name", "ModelID/DrawingID")
    strDrawingID = fncGetDrawingParam(iobjDrawDoc, strParamName)
    strDrawLinkTo = fncGetDrawingLink(iobjDrawDoc)

    Call fncGetExtendPropertiesFromDraw(iobjDrawDoc, strProperties, iblnLoad2dText)
    
    Dim typRecord As Record
    Set typRecord.CatiaObject = iobjDrawDoc
    typRecord.ParentIndex = ilngParentIndex
    typRecord.Level = 0
    typRecord.Amount = 1
    typRecord.FilePath = strFilePath
    typRecord.FileName = strFileName
    typRecord.LinkTo = strDrawLinkTo
    typRecord.ModelDrawingID = strDrawingID
    typRecord.Properties = strProperties
    Call oobjCatiaData.fncAddRecord(typRecord)
    fncGetPropertyFromDrawing = True
End Function

Private Function fncGetDrawingText(ByRef ilstDrawText() As Object, ByVal istrTextName As String, _
                                   ByRef ostrValue As String) As Boolean
    fncGetDrawingText = False
    ostrValue = ""
    
    On Error Resume Next
    Dim lngSize As Long, objText As Object, lstTargetText() As Object
    lngSize = UBound(ilstDrawText)
    On Error GoTo 0
    
    ReDim lstTargetText(0)
    
    Dim i As Long
    For i = 1 To lngSize
        Set objText = ilstDrawText(i)
        If Not objText Is Nothing Then
            If fncIsNumberedName(objText.name, istrTextName) Then
                On Error Resume Next
                Dim lngCnt As Long: lngCnt = 0
                lngCnt = UBound(lstTargetText)
                On Error GoTo 0
                ReDim Preserve lstTargetText(lngCnt + 1)
                Set lstTargetText(lngCnt + 1) = objText
            End If
        End If
    Next i
    
    Call SortTextObject(lstTargetText)
    
    On Error Resume Next
    Dim lngTargetSize As Long
    lngTargetSize = UBound(lstTargetText)
    On Error GoTo 0
    
    Dim j As Long
    For j = 1 To lngTargetSize
        Dim objTargetText As Object: Set objTargetText = Nothing
        Set objTargetText = lstTargetText(j)
        If Not objTargetText Is Nothing Then
            If fncGetDrawingText = False Then
                fncGetDrawingText = True
            Else
                ostrValue = ostrValue + "&"
            End If
            ostrValue = ostrValue + objTargetText.Text
        End If
    Next j
End Function

Private Sub SortTextObject(ByRef ilstText() As Object)
    On Error Resume Next
    Dim lngI As Long: lngI = 0
    lngI = UBound(ilstText)
    On Error GoTo 0
    
    Dim i As Long
    For i = lngI To 1 Step -1
        Dim j As Long
        For j = 1 To i - 1
            If ilstText(j).name > ilstText(j + 1).name Then
                Dim objSwap As Object
                Set objSwap = ilstText(j)
                Set ilstText(j) = ilstText(j + 1)
                Set ilstText(j + 1) = objSwap
            End If
        Next j
    Next i
End Sub

Private Function fncIsNumberedName(ByVal istrName As String, ByVal istrKeyWord As String) As Boolean
    fncIsNumberedName = False
    
    If istrName = istrKeyWord Then
        fncIsNumberedName = True
        Exit Function
    End If
    
    istrKeyWord = istrKeyWord + "_"
    
    If istrName = istrKeyWord Then Exit Function
    If InStr(istrName, istrKeyWord) <> 1 Then Exit Function
    
    Dim lngNameLength As Long
    Dim lngKeyWordLength As Long
    lngNameLength = Len(istrName)
    lngKeyWordLength = Len(istrKeyWord)
    
    Dim strSuffix As String
    strSuffix = Right(istrName, lngNameLength - lngKeyWordLength)
    
    Dim i As Long
    For i = 1 To Len(strSuffix)
        Dim strBuff As String
        strBuff = ""
        strBuff = Mid(strSuffix, i, 1)
        If strBuff = "0" Then
        ElseIf strBuff = "1" Then
        ElseIf strBuff = "2" Then
        ElseIf strBuff = "3" Then
        ElseIf strBuff = "4" Then
        ElseIf strBuff = "5" Then
        ElseIf strBuff = "6" Then
        ElseIf strBuff = "7" Then
        ElseIf strBuff = "8" Then
        ElseIf strBuff = "9" Then
        Else
            Exit Function
        End If
    Next i
    
    fncIsNumberedName = True
End Function

Private Function fncGetDrawingParam(ByRef iobjDrawDoc As Object, ByVal istrParamName As String) As String
    fncGetDrawingParam = ""
    
    Dim objRootParamSet As Object, objParams As Object, objParam As Object
    
    Set objRootParamSet = iobjDrawDoc.Parameters.RootParameterSet
    If objRootParamSet Is Nothing Then Exit Function
    
    Set objParams = objRootParamSet.DirectParameters
    If objParams Is Nothing Then Exit Function
    
    On Error Resume Next
    Set objParam = Nothing
    Set objParam = objParams.ITEM(istrParamName)
    On Error GoTo 0
    
    If Not objParam Is Nothing Then fncGetDrawingParam = objParam.Value
End Function

Private Function fncGetExtendPropertiesFromDraw(ByRef iobjDrawDoc As Object, _
                                                ByRef ostrProperties() As String, _
                                                ByVal iblnLoad2dText As Boolean) As Boolean
    fncGetExtendPropertiesFromDraw = False
    Dim strTextName As String
    Dim strParamName As String
    Dim strTemp As String
    
    On Error Resume Next
    Dim lngCnt As Long
    lngCnt = UBound(modMain.gcurMainProperty)
    On Error GoTo 0
    
    Dim lstDrawText() As Object
    If iblnLoad2dText = True Then Call GetDrawText(iobjDrawDoc, lstDrawText)
    
    Dim blnDesignNoFromText As Boolean
    blnDesignNoFromText = False
    
    Dim i As Long
    For i = 1 To lngCnt
        strTemp = ""
        strTextName = ""
        strParamName = ""
        '/ TITLE_FILEDATATYP-Text Parameter
        If modMain.gcurMainProperty(i) = "File_Data_Type" Then
            strTemp = "CATDrawing"
        Else
            If iblnLoad2dText = True Then
                strTextName = getPLMpropertyData("Drawing_Text_Name", "Form_Name", modMain.gcurMainProperty(i))
            Else
                strParamName = getPLMpropertyData("Drawing_Parameter_Name", "Form_Name", modMain.gcurMainProperty(i))
            End If
            Dim blnGetText As Boolean: blnGetText = False
            If strTextName <> "" And iblnLoad2dText = True Then
                blnGetText = fncGetDrawingText(lstDrawText, strTextName, strTemp)
                strTemp = Replace(strTemp, vbLf, " ")
                If modMain.gcurMainProperty(i) = "Full_Design_No" Then blnDesignNoFromText = True
            ElseIf strParamName <> "" And iblnLoad2dText = False Then
                strTemp = fncGetDrawingParam(iobjDrawDoc, strParamName)
            Else
                strTemp = ""
            End If
        End If
        
        ReDim Preserve ostrProperties(i)
        ostrProperties(i) = strTemp
    Next i
    
    If blnDesignNoFromText = True Then
        Dim lngIndex_FullDesignNo As Long
        Dim lngIndex_DesignNo As Long
        Dim lngIndex_BranchNo As Long
        lngIndex_FullDesignNo = fncGetIndex("Full_Design_No")
        lngIndex_DesignNo = fncGetIndex("Design_No")
        lngIndex_BranchNo = fncGetIndex("Revision_No")
        Dim strDesignNo As String
        Dim strBranchNo As String
        strDesignNo = ostrProperties(lngIndex_FullDesignNo)
        Dim strDesignSplit() As String
        strDesignSplit = Split(strDesignNo, "&")
        On Error Resume Next
        Dim lngDesignSize As Long
        lngDesignSize = UBound(strDesignSplit)
        On Error GoTo 0
        If 0 < lngDesignSize Then strDesignNo = strDesignSplit(0)
        
        Dim strSplit() As String
        strSplit = Split(strDesignNo, "-")
        
        On Error Resume Next
        Dim lngSize As Long
        lngSize = UBound(strSplit)
        On Error GoTo 0
        
        If 0 < lngSize Then
            strDesignNo = strSplit(0)
            ostrProperties(lngIndex_DesignNo) = strSplit(0)
            strBranchNo = strSplit(1)
            For i = 2 To lngSize
                strBranchNo = strBranchNo & "-" & strSplit(i)
            Next i
            ostrProperties(lngIndex_BranchNo) = strBranchNo
        End If
        
        '/ Classification --------------------------------
        If strBranchNo <> "" Then
            '/NF_Classification
            Dim strRightBranch As String
            strRightBranch = Right(strBranchNo, 2)
            If IsNumeric(strRightBranch) = True Then
                Dim strClassVal As String
                If 0 <= strRightBranch And strRightBranch <= 79 Then
                    strClassVal = "Internal data"
                ElseIf 80 <= strRightBranch And strRightBranch <= 99 Then
                    strClassVal = "Submission data"
                Else
                    strClassVal = ""
                End If
                Dim lngIndex_Classification As Long
                lngIndex_Classification = fncGetIndex("Classification")
                ostrProperties(lngIndex_Classification) = strClassVal
            End If
        End If
    
        '/ Section
        '/ NF_DesignNo
        Dim strSectionCode As String
        Dim strSection As String
        On Error Resume Next
        strSectionCode = ""
        strSectionCode = Left(strDesignNo, 2)
        On Error GoTo 0
        strSection = getSectionData("Section", "Office_Code", strSectionCode)
        If strSection <> "" Then
            '/Section/DesignNo
            On Error Resume Next
            strDesignNo = Mid(strDesignNo, 3)
            On Error GoTo 0
            ostrProperties(lngIndex_DesignNo) = strDesignNo
        Else
            '/ Section
            On Error Resume Next
            strSectionCode = ""
            strSectionCode = Left(strDesignNo, 1)
            On Error GoTo 0
            strSection = getSectionData("Section", "Office_Code", strSectionCode)
            If strSection <> "" Then
                '/Section/DesignNo
                On Error Resume Next
                strDesignNo = Mid(strDesignNo, 2)
                On Error GoTo 0
                ostrProperties(lngIndex_DesignNo) = strDesignNo
            Else
                '/Section
                strSectionCode = ""
            End If
        End If
        
        Dim blnOldSection As Boolean
        If strSection <> "" Then
            '/ NF_Section
            Dim lngIndex_Section As Long
            lngIndex_Section = fncGetIndex("Section")
            ostrProperties(lngIndex_Section) = strSection
            If IsNumeric(strSectionCode) = True Then
                blnOldSection = False
            Else
                blnOldSection = True
            End If
        Else
            blnOldSection = True
        End If
    
        '/ Status --------------------------------
        '/ NF_DesignNoÇ/Status/DesignNo
        Dim strStatus As String
        Dim strStatusCode As String
        On Error Resume Next
        strStatusCode = ""
        strStatusCode = Left(strDesignNo, 1)
        If strStatusCode = "M" Or strStatusCode = "T" Then
            '/DesignNo
            On Error Resume Next
            strDesignNo = Mid(strDesignNo, 2)
            On Error GoTo 0
            ostrProperties(lngIndex_DesignNo) = strDesignNo
        Else
            '/Status
            strStatusCode = ""
        End If
        On Error GoTo 0
            
        If Trim(strStatusCode) <> "" Then
            If blnOldSection = False Then
                If strStatusCode = "M" Then
                    strStatus = "Mass production"
                ElseIf strStatusCode = "T" Then
                    strStatus = "Prototype/Study"
                Else
                    strStatus = ""
                End If
            Else
                If strStatusCode = "T" Then
                    strStatus = "Prototype/Study"
                    On Error Resume Next
                Else
                    strStatus = "Mass production"
                End If
            End If
        End If
        
        '/ Status
        Dim lngIndex_Status As Long
        lngIndex_Status = fncGetIndex("Current_Status")
        ostrProperties(lngIndex_Status) = strStatus
        '/DesignNo
        If blnOldSection = True Then ostrProperties(lngIndex_DesignNo) = strSectionCode & strDesignNo
    End If
    
    '/ NF_DesignerAliasName
    Dim lngIndex_Designer As Long
    lngIndex_Designer = fncGetIndex("Designer")
    
    Dim strDesigner As String
    strDesigner = ostrProperties(lngIndex_Designer)
    
    Dim strAliasName As String
    ostrProperties(lngIndex_Designer) = strDesigner
    fncGetExtendPropertiesFromDraw = True
End Function

Private Function fncGetPropertyFromProduct(ByRef iobjProduct As Object, ByVal iintLevel As Integer, _
                                           ByVal ilngParentIndex As Long, _
                                           ByRef oobjCatiaData As CATIAPropertyTable) As Boolean
    fncGetPropertyFromProduct = False
    
    If iobjProduct Is Nothing Then Exit Function

    Dim strFilePath As String, strFileName As String, strPartNumber As String, strInstanceName As String, strPropName As String, strModelID As String
    
    If fncGetDocPath(iobjProduct, strFilePath) = False Then Exit Function
    If fncGetDocName(iobjProduct, strFileName) = False Then Exit Function
    
    strFileName = fncSplitFileName(strFileName)
    strPartNumber = iobjProduct.partNumber
    strInstanceName = iobjProduct.name
    strPropName = getPLMpropertyData("Property_Name", "Form_Name", "ModelID/DrawingID")
    strModelID = fncGetUserRefProperty(iobjProduct, strPropName)
    
    Dim strProperties() As String
    Call fncGetUserRefProperties(iobjProduct, strProperties)
    
    Dim typRecord As Record
    Set typRecord.CatiaObject = iobjProduct
    typRecord.ParentIndex = ilngParentIndex
    typRecord.Level = iintLevel
    typRecord.Amount = 1
    typRecord.FilePath = strFilePath
    typRecord.FileName = strFileName
    typRecord.partNumber = strPartNumber
    typRecord.InstanceName = strInstanceName
    typRecord.ModelDrawingID = strModelID
    typRecord.Properties = strProperties

    Dim lngParentIndex As Long
    lngParentIndex = oobjCatiaData.fncAddRecord(typRecord)
    If lngParentIndex = 0 Then lngParentIndex = ilngParentIndex
    
    '/ Product
    Dim i As Long
    For i = 1 To iobjProduct.Products.count
        If fncGetPropertyFromProduct(iobjProduct.Products.ITEM(i), iintLevel + 1, lngParentIndex, oobjCatiaData) = False Then Exit Function
    Next i
    fncGetPropertyFromProduct = True
End Function

Private Function fncGetDocType(ByRef iobjProduct As Object, ByRef ostrDocType As String) As Boolean
    fncGetDocType = False
    ostrDocType = ""
    
    Dim strType As String
    
    '/ Parent/RootProduct
    Dim objParentProd As Object
    Set objParentProd = iobjProduct.Parent.Parent
    If TypeName(objParentProd) = "Application" Then
        '/ RootDocument
        Set objParentProd = iobjProduct.Parent
    
        If TypeName(objParentProd) = "ProductDocument" Then
            strType = "CATProduct"
        ElseIf TypeName(objParentProd) = "PartDocument" Then
            strType = "CATPart"
        End If
    Else
        Dim strParentPath As String, strPath As String
        If fncGetDocPath(objParentProd, strParentPath) = False Then Exit Function
        If fncGetDocPath(iobjProduct, strPath) = False Then Exit Function
        On Error GoTo 0
        
        '/Component
        Dim objDoc As Object
        Set objDoc = iobjProduct.ReferenceProduct.Parent
        
        If strParentPath = strPath Then
            strType = "Component"
        ElseIf TypeName(objDoc) = "ProductDocument" Then
            strType = "CATProduct"
        ElseIf TypeName(objDoc) = "PartDocument" Then
            strType = "CATPart"
        End If
    End If
    
    If strType = "" Then Exit Function
    
    ostrDocType = strType
    fncGetDocType = True
End Function

Private Function fncGetDocPath(ByRef iobjProduct As Object, ByRef ostrDocPath As String) As Boolean
    fncGetDocPath = False
    ostrDocPath = ""
    
    Dim strPath As String
    
    '/ Parent
    Dim objParentProd As Object
    Set objParentProd = iobjProduct.Parent.Parent
    If TypeName(objParentProd) = "Application" Then
        '/ RootDocumentÇÃèÍçáÇÕParentÇ™DocumentÇ…Ç»ÇÈ
        Set objParentProd = iobjProduct.Parent
        strPath = objParentProd.fullName
    Else
        '/ RootDocumentÇ≈Ç»Ç¢èÍçáÇÕÅAReferenceProductÇÃParentÇéÊìæÇ∑ÇÈ
        On Error Resume Next
        Dim objDoc As Object
        Set objDoc = iobjProduct.ReferenceProduct.Parent
        If objDoc Is Nothing Then Exit Function
        strPath = objDoc.fullName
        On Error GoTo 0
    End If
    
    If strPath = "" Then Exit Function
    
    ostrDocPath = strPath
    fncGetDocPath = True
End Function

Private Function fncGetDocName(ByRef iobjProduct As Object, ByRef ostrDocName As String) As Boolean
    fncGetDocName = False
    ostrDocName = ""
    
    Dim strName As String
    
    '/ Parent
    Dim objParentProd As Object
    Set objParentProd = iobjProduct.Parent.Parent
    If TypeName(objParentProd) = "Application" Then
        '/ RootDocumentÇÃèÍçáÇÕParentÇ™DocumentÇ…Ç»ÇÈ
        Set objParentProd = iobjProduct.Parent
        strName = objParentProd.name
    Else
        '/ RootDocument
        On Error Resume Next
        Dim objDoc As Object
        Set objDoc = iobjProduct.ReferenceProduct.Parent
        If objDoc Is Nothing Then Exit Function
        strName = objDoc.name
        On Error GoTo 0
    End If
    
    If strName = "" Then Exit Function
    
    ostrDocName = strName
    fncGetDocName = True
End Function

Private Function fncGetUserRefProperties(ByRef iobjProduct As Object, ByRef ostrProperties() As String) As Boolean
    fncGetUserRefProperties = False
    Dim strTemp As String
        
    '/ Main
    On Error Resume Next
    Dim lngCnt As Long
    lngCnt = UBound(modMain.gcurMainProperty)
    On Error GoTo 0
    
    Dim i As Long
    For i = 1 To lngCnt
        
        Dim strPropName As String
        strPropName = getPLMpropertyData("Property_Name", "Form_Name", modMain.gcurMainProperty(i))

        If modMain.gcurMainProperty(i) = "File_Data_Type" Then
            If fncGetDocType(iobjProduct, strTemp) = False Then Exit Function
        Else
            strTemp = ""
            strTemp = fncGetUserRefProperty(iobjProduct, strPropName)
        End If
        
        ReDim Preserve ostrProperties(i)
        ostrProperties(i) = strTemp
    Next i
    
    fncGetUserRefProperties = True
End Function

Private Function fncGetUserRefProperty(ByRef iobjProduct As Object, _
                                           ByVal istrPropertyName As String) As String

    fncGetUserRefProperty = ""
    
    If Trim(istrPropertyName) = "" Then Exit Function
    
    Dim objRefProd As Object
    Set objRefProd = iobjProduct.ReferenceProduct

    On Error Resume Next
    Dim objProperties As Object
    Set objProperties = Nothing
    Set objProperties = objRefProd.UserRefProperties
    On Error GoTo 0
    If objProperties Is Nothing Then Exit Function
    
    On Error Resume Next
    Dim objParam As Object
    Set objParam = Nothing
    Set objParam = objProperties.ITEM(istrPropertyName)
    On Error GoTo 0
        
    If Not objParam Is Nothing Then fncGetUserRefProperty = objParam.Value
End Function

Private Function fncCountInstance(ByRef iobjProduct As Object) As Long
    fncCountInstance = 0
    
    If iobjProduct Is Nothing Then Exit Function

    On Error Resume Next
    Dim objProducts As Object
    Set objProducts = iobjProduct.Parent
    On Error GoTo 0
    
    If TypeName(objProducts) <> "Products" Then
        '/ Products
        fncCountInstance = 1
        Exit Function
    End If
    
    '/ Product
    Dim strFilePath As String
    If fncGetDocPath(iobjProduct, strFilePath) = False Then Exit Function
    
    Dim i As Long
    For i = 1 To objProducts.count
        '/ Product
        Dim objChild As Object
        Set objChild = objProducts.ITEM(i)
        Dim strChildPath As String
        If fncGetDocPath(objChild, strChildPath) = True Then
            If strFilePath = strChildPath Then fncCountInstance = fncCountInstance + 1
        End If
    Next i
End Function

Private Sub AddNotFoundTextList(ByVal istrTextName As String)
    Dim lngCnt As Long

    On Error Resume Next
    lngCnt = UBound(gstrNotFoundText)
    On Error GoTo 0
    
    '/ Text
    Dim blnExist As Boolean
    blnExist = False
    Dim i As Integer
    For i = 1 To lngCnt
        If gstrNotFoundText(i) = istrTextName Then blnExist = True
    Next i
    
    If blnExist = True Then Exit Sub
    
    ReDim Preserve gstrNotFoundText(lngCnt + 1)
    gstrNotFoundText(lngCnt + 1) = istrTextName
End Sub

Public Function fncSetProperty(ByRef iobjCatiaData As CATIAPropertyTable, _
                               ByRef iobjExcelData As CATIAPropertyTable, _
                               ByVal iblnDrawingUpdate As Boolean, _
                               Optional ByVal iblnSetTitleBlock As Boolean = False) As Boolean

    fncSetProperty = False
    ReDim gstrNotFoundText(0)
    
    Dim lngCnt As Long
    lngCnt = iobjCatiaData.fncCount()
    
    Dim i As Long
    For i = 1 To lngCnt
        
        Dim typCATIARecord As Record
        Dim typExcelRecord As Record
        If iobjCatiaData.fncItem(i, typCATIARecord) = False Then
            Exit Function
        ElseIf iobjExcelData.fncItem(i, typExcelRecord) = False Then
            Exit Function
        End If
        
        If Trim(typExcelRecord.Sel) = True Then
            '/ SET PROPERTY
            If iblnSetTitleBlock = False Then
                If typCATIARecord.Properties(modMain.fncGetIndex("File_Data_Type") - 1) = "CATDrawing" Then
                    If fncSetPropertyToDrawing(typCATIARecord.CatiaObject, typExcelRecord) = False Then Exit Function
                Else
                    If fncSetPropertyToProduct(typCATIARecord.CatiaObject, typExcelRecord) = False Then Exit Function
                End If
            '/ SET title block
            Else
                '/ DrawingText
                If typCATIARecord.Properties(modMain.fncGetIndex("File_Data_Type") - 1) = "CATDrawing" Then
                    If fncSetPropertyToTitleBlock(typCATIARecord.CatiaObject, typExcelRecord) = False Then Exit Function
                End If
            End If
        End If
        
    Next i
    
    If iblnDrawingUpdate = True Then
        For i = 1 To lngCnt
            Dim typRecord As Record
            If iobjCatiaData.fncItem(i, typRecord) = False Then Exit Function
            If typRecord.Properties(modMain.fncGetIndex("File_Data_Type")) = "CATDrawing" Then Call typRecord.CatiaObject.Update
        Next i
    End If
    
    fncSetProperty = True
End Function

Private Function fncSetPropertyToDrawing(ByRef iobjDrawDoc As Object, _
                                         ByRef iobjRecord As Record) As Boolean

    fncSetPropertyToDrawing = False
    
    If iobjDrawDoc Is Nothing Then Exit Function
    
    Dim strParamName As String
    strParamName = ""

    '/ DrawingID
    strParamName = getPLMpropertyData("Drawing_Parameter_Name", "Form_Name", "ModelID/DrawingID")
    Call fncDeleteParam(iobjDrawDoc, strParamName)
    Call fncSetExtendPropertiesFromDraw(iobjDrawDoc, iobjRecord.Properties)
    
    fncSetPropertyToDrawing = True
End Function

Private Function fncSetDrawingParam(ByRef iobjDrawDoc As Object, ByVal istrParamName As String, _
                                   ByVal istrValue As String) As Boolean
                                   
    fncSetDrawingParam = False
    
    '/ RootParameterSetí
    Dim objRootParamSet As Object
    Set objRootParamSet = iobjDrawDoc.Parameters.RootParameterSet
    If objRootParamSet Is Nothing Then Exit Function
    
    Dim objParams As Object
    Set objParams = objRootParamSet.DirectParameters
    If objParams Is Nothing Then Exit Function
    
    '/Parameter
    On Error Resume Next
    Dim objParam As Object
    Set objParam = Nothing
    Set objParam = objParams.ITEM(istrParamName)
    On Error GoTo 0
    
    Dim strTemp As String
    strTemp = Replace(istrValue, vbLf, " ")
    
    If objParam Is Nothing Then
        Set objParam = objParams.CreateString(istrParamName, strTemp)
    Else
        If objParam.Value <> strTemp Then objParam.Value = strTemp
    End If
    
End Function

Private Function fncSetExtendPropertiesFromDraw(ByRef iobjDrawDoc As Object, _
                                                ByRef istrProperties() As String) As Boolean
    
    fncSetExtendPropertiesFromDraw = False
    
    Dim strParamName As String
    Dim lngIndex_Section As Long
    Dim lngIndex_Status As Long
    Dim lngIndex_Revision As Long
    lngIndex_Section = modMain.fncGetIndex("Section")
    lngIndex_Status = modMain.fncGetIndex("Current_Status")
    lngIndex_Revision = modMain.fncGetIndex("Revision_No")
    
    '/ Main
    On Error Resume Next
    Dim lngCnt As Long
    lngCnt = UBound(modMain.gcurMainProperty)
    On Error GoTo 0
    
    Dim i As Long
    For i = 1 To lngCnt
        strParamName = ""
        strParamName = getPLMpropertyData("Drawing_Parameter_Name", "Form_Name", modMain.gcurMainProperty(i))
        If strParamName <> "" Then Call fncSetDrawingParam(iobjDrawDoc, strParamName, istrProperties(i))
    Next i
    
    fncSetExtendPropertiesFromDraw = True
End Function

Private Function fncSetPropertyToTitleBlock(ByRef iobjDrawDoc As Object, _
                                            ByRef iobjRecord As Record) As Boolean
    fncSetPropertyToTitleBlock = False
       
    Dim strTextName As String
    Dim lngIndex_Section As Long
    Dim lngIndex_Status As Long
    Dim lngIndex_Revision As Long
    lngIndex_Section = modMain.fncGetIndex("Section")
    lngIndex_Status = modMain.fncGetIndex("Current_Status")
    lngIndex_Revision = modMain.fncGetIndex("Revision_No")
    
    Dim lstDrawText() As Object
    Call GetDrawText(iobjDrawDoc, lstDrawText)
    
    On Error Resume Next
    Dim lngCnt As Long
    lngCnt = UBound(modMain.gcurMainProperty)
    On Error GoTo 0
    
    Dim i As Long
    For i = 1 To lngCnt
        strTextName = ""
        strTextName = getPLMpropertyData("Drawing_Text_Name", "Form_Name", modMain.gcurMainProperty(i))
        
        Dim strValue As String
        strValue = Replace(iobjRecord.Properties(i), vbLf, " ")

        If strTextName <> "" Then Call fncSetDrawingText(lstDrawText, strTextName, strValue)
    Next i
    
    fncSetPropertyToTitleBlock = True
End Function

Private Sub GetDrawText(ByRef iobjDrawDoc As Object, ByRef olstDrawText() As Object)
    ReDim olstDrawText(0)
    Dim objSheet As Object
    Set objSheet = iobjDrawDoc.DrawingRoot.ActiveSheet
    Dim i As Long
    For i = 1 To objSheet.Views.count
        Dim objView As Object
        Set objView = objSheet.Views.ITEM(i)
        Dim j As Long
        For j = 1 To objView.Texts.count
            On Error Resume Next
            Dim lngSize As Long
            lngSize = UBound(olstDrawText)
            On Error GoTo 0
            ReDim Preserve olstDrawText(lngSize + 1)
            Set olstDrawText(lngSize + 1) = objView.Texts.ITEM(j)
        Next j
    Next i
End Sub

Private Function fncSetDrawingText(ByRef ilstDrawText() As Object, ByVal istrTextName As String, _
                                   ByVal istrValue As String) As Boolean
                                   
    fncSetDrawingText = False
    
    On Error Resume Next
    Dim lngSize As Long
    lngSize = UBound(ilstDrawText)
    On Error GoTo 0
    Dim lstTargetText() As Object
    ReDim lstTargetText(0)
    Dim i As Long
    For i = 1 To lngSize
        Dim objText As Object
        Set objText = ilstDrawText(i)
        If Not objText Is Nothing Then
            If fncIsNumberedName(objText.name, istrTextName) = True Then
                On Error Resume Next
                Dim lngCnt As Long: lngCnt = 0
                lngCnt = UBound(lstTargetText)
                On Error GoTo 0
                ReDim Preserve lstTargetText(lngCnt + 1)
                Set lstTargetText(lngCnt + 1) = objText
            End If
        End If
    Next i
    
    Call SortTextObject(lstTargetText)
    Dim lstSplit As Variant
    lstSplit = Split(istrValue, "&")
    On Error Resume Next
    Dim lngSplitSize As Long
    lngSplitSize = UBound(lstSplit)
    lngSplitSize = lngSplitSize + 1
    On Error GoTo 0
    
    On Error Resume Next
    Dim lngTargetSize As Long
    lngTargetSize = UBound(lstTargetText)
    On Error GoTo 0
    
    If lngSplitSize < lngTargetSize Then lngTargetSize = lngSplitSize
    
    Dim j As Long
    For j = 1 To lngTargetSize
        Dim objTargetText As Object: Set objTargetText = Nothing
        Set objTargetText = lstTargetText(j)
        If Not objTargetText Is Nothing Then objTargetText.Text = lstSplit(j - 1)
    Next j
    Call AddNotFoundTextList(istrTextName)
End Function

Private Function fncSetPropertyToProduct(ByRef iobjProduct As Object, _
                                         ByRef itypRecord As Record) As Boolean

    fncSetPropertyToProduct = False
    
    If iobjProduct Is Nothing Then Exit Function
    
    Dim strPropName As String
    strPropName = getPLMpropertyData("Property_Name", "Form_Name", "ModelID/DrawingID")
    Call fncDeleteUserRefProperty(iobjProduct, strPropName)
    Call fncSetUserRefProperties(iobjProduct, itypRecord.Properties)
    
    fncSetPropertyToProduct = True
End Function

Private Function fncSetUserRefProperties(ByRef iobjProduct As Object, ByRef istrProperties() As String) As Boolean
    fncSetUserRefProperties = False
    On Error Resume Next
    Dim lngCnt As Long
    lngCnt = UBound(modMain.gcurMainProperty)
    On Error GoTo 0
    
    Dim i As Long
    For i = 1 To lngCnt
        Dim strPropName As String
        strPropName = getPLMpropertyData("Property_Name", "Form_Name", modMain.gcurMainProperty(i))
        Call fncSetUserRefProperty(iobjProduct, strPropName, istrProperties(i))
    Next i
    
    fncSetUserRefProperties = True
End Function

Private Function fncSetUserRefProperty(ByRef iobjProduct As Object, ByVal istrPropertyName As String, _
                                       ByVal istrValue As String) As Boolean

    fncSetUserRefProperty = False
    
    If Trim(istrPropertyName) = "" Then Exit Function
    
    Dim objRefProd As Object
    Set objRefProd = Nothing
    Set objRefProd = iobjProduct.ReferenceProduct
    
    On Error Resume Next
    Dim objProperties As Object
    Set objProperties = Nothing
    Set objProperties = objRefProd.UserRefProperties
    On Error GoTo 0
    If objProperties Is Nothing Then Exit Function
    
    On Error Resume Next
    Dim objParam As Object
    Set objParam = Nothing
    Set objParam = objProperties.ITEM(istrPropertyName)
    On Error GoTo 0
    
    Dim strTemp As String
    strTemp = Replace(istrValue, vbLf, " ")
    
    If objParam Is Nothing Then
        Set objParam = objProperties.CreateString(istrPropertyName, strTemp)
    Else
        If objParam.Value <> strTemp Then objParam.Value = strTemp
    End If
        
    fncSetUserRefProperty = True
End Function

Public Function fncDeleteProperty(ByRef iobjCatiaData As CATIAPropertyTable) As Boolean
    fncDeleteProperty = False
    
    Dim lngCnt As Long
    lngCnt = iobjCatiaData.fncCount()
    
    Dim i As Long
    For i = 1 To lngCnt
        
        Dim typCATIARecord As Record
        If iobjCatiaData.fncItem(i, typCATIARecord) = False Then Exit Function
        
        If typCATIARecord.Properties(modMain.fncGetIndex("File_Data_Type") - 1) = "CATDrawing" Then
            If fncDeletePropertyOnDrawing(typCATIARecord.CatiaObject) = False Then Exit Function
        Else
            If fncDeletePropertyOnProduct(typCATIARecord.CatiaObject) = False Then Exit Function
        End If
    Next i
    
    fncDeleteProperty = True
End Function

Private Function fncDeletePropertyOnDrawing(ByRef iobjDrawDoc As Object) As Boolean
    fncDeletePropertyOnDrawing = False
    
    If iobjDrawDoc Is Nothing Then Exit Function
    
    Dim strParamName As String
    strParamName = getPLMpropertyData("Drawing_Parameter_Name", "Form_Name", "ModelID/DrawingID")
    If strParamName <> "" Then Call fncDeleteParam(iobjDrawDoc, strParamName)
    
    Call fncDeleteDrawingParameters(iobjDrawDoc)
    Call DeleteRootParamSet(iobjDrawDoc)
    Call iobjDrawDoc.DrawingRoot.Update
    
    fncDeletePropertyOnDrawing = True
End Function

Private Sub DeleteRootParamSet(ByRef iobjDrawDoc As Object)
    '/ RootParameterSet
    Dim objRootParamSet As Object
    Set objRootParamSet = iobjDrawDoc.Parameters.RootParameterSet
    If objRootParamSet Is Nothing Then Exit Sub
    
    '/ RootParameterSet
    Dim objParams As Object
    Set objParams = objRootParamSet.AllParameters
    If objParams Is Nothing Then Exit Sub
    
    '/ RootParameterSet
    If objParams.count = 0 Then
        Dim objSel As Object
        Set objSel = iobjDrawDoc.Selection
        Call objSel.clear
        Call objSel.Add(objRootParamSet)
        Call objSel.Delete
    End If
End Sub

Private Function fncDeleteParam(ByRef iobjDrawDoc As Object, _
                                ByVal istrParamName As String) As Boolean

    fncDeleteParam = False
     
    '/ RootParameterSet
    Dim objRootParamSet As Object
    Set objRootParamSet = iobjDrawDoc.Parameters.RootParameterSet
    If objRootParamSet Is Nothing Then Exit Function
    
    Dim objParams As Object
    Set objParams = objRootParamSet.DirectParameters
    If objParams Is Nothing Then Exit Function
    
    On Error Resume Next
    Dim objParam As Object
    Set objParam = objParams.ITEM(istrParamName)
    
    If Not objParam Is Nothing Then Call objParams.remove(istrParamName)
    On Error GoTo 0
    
    fncDeleteParam = True
End Function

Private Function fncDeleteDrawingParameters(ByRef iobjDrawDoc As Object) As Boolean
    fncDeleteDrawingParameters = False
        
    On Error Resume Next
    Dim lngCnt As Long
    lngCnt = UBound(modMain.gcurMainProperty)
    On Error GoTo 0
    
    Dim i As Long
    For i = 1 To lngCnt
        Dim strPropertyName As String
        strPropertyName = modMain.gcurMainProperty(i)
        Dim strParamName As String
        strParamName = getPLMpropertyData("Drawing_Parameter_Name", "Form_Name", strPropertyName)
        If strParamName <> "" Then Call fncDeleteParam(iobjDrawDoc, strParamName)
    Next i
    
    fncDeleteDrawingParameters = True
End Function

Private Function fncDeletePropertyOnProduct(ByRef iobjProduct As Object) As Boolean
    fncDeletePropertyOnProduct = False
    
    If iobjProduct Is Nothing Then Exit Function
    Dim strPropertyName As String
    strPropertyName = getPLMpropertyData("Property_Name", "Form_Name", "ModelID/DrawingID")
    
    Call fncDeleteUserRefProperty(iobjProduct, strPropertyName)
    Call fncDeleteUserRefProperties(iobjProduct)
    
    fncDeletePropertyOnProduct = True
End Function

Private Function fncDeleteUserRefProperty(ByRef iobjProduct As Object, _
                                       ByVal istrPropertyName As String) As Boolean

    fncDeleteUserRefProperty = False
    
    If Trim(istrPropertyName) = "" Then Exit Function
    
    Dim objRefProd As Object
    Set objRefProd = Nothing
    Set objRefProd = iobjProduct.ReferenceProduct
    
    On Error Resume Next
    Dim objProperties As Object
    Set objProperties = Nothing
    Set objProperties = objRefProd.UserRefProperties
    On Error GoTo 0
    If objProperties Is Nothing Then Exit Function
    
    On Error Resume Next
    Dim objParam As Object
    Set objParam = objProperties.ITEM(istrPropertyName)
    
    If Not objParam Is Nothing Then Call objProperties.remove(istrPropertyName)
    On Error GoTo 0
    
    fncDeleteUserRefProperty = True
End Function

Private Function fncDeleteUserRefProperties(ByRef iobjProduct As Object) As Boolean

    fncDeleteUserRefProperties = False
    
    On Error Resume Next
    Dim lngCnt As Long
    lngCnt = UBound(modMain.gcurMainProperty)
    On Error GoTo 0
    
    Dim i As Long
    For i = 1 To lngCnt
        Dim strPropName As String
        strPropName = getPLMpropertyData("Property_Name", "Form_Name", modMain.gcurMainProperty(i))
        Call fncDeleteUserRefProperty(iobjProduct, strPropName)
    Next i
    
    fncDeleteUserRefProperties = True
End Function

Public Function fncSaveData(ByRef iobjCatiaData As CATIAPropertyTable, _
                            ByVal istrSaveDir As String, _
                            ByRef iobjExcelData As CATIAPropertyTable) As Boolean

    fncSaveData = False
    
    Dim lndLastLevel As Long
    lndLastLevel = iobjExcelData.fncGetLastLevel()
    
    Dim lngCnt As Long
    lngCnt = iobjExcelData.fncCount()
    Dim i As Long
    For i = lndLastLevel To 0 Step -1
        
        Dim j As Long
        For j = 1 To lngCnt
            Dim typCATIARecord As Record
            Dim typExcelRecord As Record
            If iobjCatiaData.fncItem(j, typCATIARecord) = False Then Exit Function
            If iobjExcelData.fncItem(j, typExcelRecord) = False Then Exit Function
            
            If typExcelRecord.IsChildInstance = True Then GoTo continue
            If typExcelRecord.Sel <> True Then GoTo continue
            If typExcelRecord.Level <> i Then GoTo continue
                
            Dim strOldPath As String
            strOldPath = typExcelRecord.FilePath
            
            '/ 3DEX
            Dim lngFileNameIndex As Long
            Dim blnIn3DEX As Boolean
            lngFileNameIndex = modMain.fncGetIndex("File_Data_Name") - 1
            blnIn3DEX = fncSavedIn3dex(strOldPath)
            If blnIn3DEX = True And _
               typExcelRecord.FileName = typExcelRecord.Properties(lngFileNameIndex) Then
                GoTo continue
            End If
            
            Dim lngClassification As Long
            lngClassification = modMain.fncGetIndex("Classification") - 1
            Dim strClassification As String
            strClassification = typExcelRecord.Properties(lngClassification)
            
            '/ SaveAsNewName
            '/ Reference,SubProduct,Layout
            If blnIn3DEX = True And _
                Trim(modMain.gstrSaveAsNewName) = "0" And _
                (strClassification = "Reference" Or _
                 strClassification = "SubProduct" Or _
                 strClassification = "LayOut") Then
                GoTo continue
            End If

            Dim strNewPath As String
            Dim strNewName As String
            If fncSaveAs(typCATIARecord, istrSaveDir, strNewPath, strNewName, typExcelRecord) = False Then Exit Function

            Call iobjExcelData.UpdateModelID(j, typExcelRecord.ModelDrawingID)
            
continue:
        Next j
    Next i
    fncSaveData = True
End Function

Public Function fncSavedIn3dex(ByVal istrPath As String) As Boolean
    fncSavedIn3dex = False
    
    Dim lngIndex As Long
    lngIndex = InStr(UCase(istrPath), UCase(modMain.gstr3dexCacheDir))
    
    If 0 < lngIndex Then fncSavedIn3dex = True
End Function

Private Function fncSaveAs(ByRef itypCATIARecord As Record, ByVal istrSaveDir As String, _
                           ByRef ostrSavePath As String, ByRef ostrSaveName As String, _
                           ByRef itypExcelRecord As Record) As Boolean
    ostrSavePath = ""
    fncSaveAs = False
    
    Dim lngFileNameIndex As Long
    Dim lngFileTypeIndex As Long
    Dim lngClassificationIndex As Long
    lngFileNameIndex = modMain.fncGetIndex("File_Data_Name") + 1
    lngFileTypeIndex = modMain.fncGetIndex("File_Data_Type") - 1
    lngClassificationIndex = modMain.fncGetIndex("Classification") - 1
    
    '/ FileType
    Dim strExtension As String
    If StrConv(itypExcelRecord.Properties(lngFileTypeIndex), vbUpperCase) = StrConv("CATDrawing", vbUpperCase) Then
        strExtension = "CATDrawing"
    ElseIf StrConv(itypExcelRecord.Properties(lngFileTypeIndex), vbUpperCase) = StrConv("CATProduct", vbUpperCase) Then
        strExtension = "CATProduct"
    ElseIf StrConv(itypExcelRecord.Properties(lngFileTypeIndex), vbUpperCase) = StrConv("CATPart", vbUpperCase) Then
        strExtension = "CATPart"
    Else
        strExtension = itypExcelRecord.Properties(lngFileTypeIndex)
    End If
    
    '/SaveAsNewName-Classification-Reference/SubProduct/Layout
    Dim strClassification As String
    strClassification = itypExcelRecord.Properties(lngClassificationIndex)
    
    Dim strFileName As String
    If Trim(modMain.gstrSaveAsNewName) = "0" And _
      (strClassification = "Reference" Or _
       strClassification = "SubProduct" Or _
       strClassification = "LayOut") Then
        strFileName = itypExcelRecord.FileName
    Else
        strFileName = itypExcelRecord.Properties(lngFileNameIndex)
    End If
    
    strFileName = strFileName & "." & strExtension
    
    Dim strFullPath As String
    strFullPath = istrSaveDir & "\" & strFileName
    
    '/ Document
    Dim objDocument As Object
    If itypCATIARecord.Properties(lngFileTypeIndex) = "CATDrawing" Then
        Set objDocument = itypCATIARecord.CatiaObject
    Else
        On Error Resume Next
        Set objDocument = itypCATIARecord.CatiaObject.ReferenceProduct.Parent
        On Error GoTo 0
    End If
    
    If objDocument Is Nothing Then Exit Function
    
    
'R50-MODELID
    If itypCATIARecord.Properties(lngFileTypeIndex) = "CATDrawing" Then
        Dim strParamName As String
        strParamName = getPLMpropertyData("Drawing_Parameter_Name", "Form_Name", "ModelID/DrawingID")
        If Trim(strParamName) <> "" Then Call fncDeleteParam(itypCATIARecord.CatiaObject, strParamName)
    Else
        Dim strPropName As String
        strPropName = getPLMpropertyData("Property_Name", "Form_Name", "ModelID/DrawingID")
        If Trim(strPropName) <> "" Then Call fncDeleteUserRefProperty(itypCATIARecord.CatiaObject, strPropName)
    End If
    
    itypExcelRecord.ModelDrawingID = ""
    
    '/ SaveAs
    mobjCATIA.DisplayFileAlerts = False
    Call objDocument.SaveAs(strFullPath)
    mobjCATIA.DisplayFileAlerts = True
    
    ostrSavePath = strFullPath
    ostrSaveName = strFileName
    fncSaveAs = True

End Function

Public Function fncCheckBeforeSave(ByRef iobjExcelData As CATIAPropertyTable) As String

    fncCheckBeforeSave = ""
    
    Dim lngCnt As Long
    lngCnt = iobjExcelData.fncCount()
    
    Dim lngSaveCnt As Long
    lngSaveCnt = 0
    
    Dim i As Long
    For i = 1 To lngCnt
        
        Dim typExcelRecord As Record
        If iobjExcelData.fncItem(i, typExcelRecord) = False Then
            fncCheckBeforeSave = "E013"
            Exit Function
        End If
        
        If typExcelRecord.IsChildInstance = True Then GoTo CONTINUE1
        If typExcelRecord.Sel <> True Then GoTo CONTINUE1
        
        Dim lngClassification As Long
        Dim lngFileName As Long
        Dim lngFileType As Long
        lngClassification = modMain.fncGetIndex("Classification")
        lngFileName = modMain.fncGetIndex("File_Data_Name")
        lngFileType = modMain.fncGetIndex("File_Data_Type")
        
        If typExcelRecord.Properties(lngFileName) = "" Then
            fncCheckBeforeSave = "E046"
            Exit Function
        End If
        
        If typExcelRecord.Properties(lngFileType) = "" Then
            fncCheckBeforeSave = "E013"
            Exit Function
        End If
        
        Dim strFileName As String
        Dim strClassification As String
        strClassification = typExcelRecord.Properties(lngClassification)
        If Trim(modMain.gstrSaveAsNewName) = "0" And _
           (strClassification = "Reference" Or _
            strClassification = "SubProduct" Or _
            strClassification = "LayOut") Then
            strFileName = typExcelRecord.FileName
        Else
            strFileName = typExcelRecord.Properties(lngFileName) & "." & typExcelRecord.Properties(lngFileType)
        End If
        
        Dim j As Long
        For j = i + 1 To lngCnt
        
            Dim typTemp As Record
            If iobjExcelData.fncItem(j, typTemp) = False Then
                fncCheckBeforeSave = "E013"
                Exit Function
            End If
            
            If typTemp.Sel <> "*" Then GoTo CONTINUE2
            If typTemp.IsChildInstance = True Then GoTo CONTINUE2
            
            Dim strTempFileName As String
            Dim strTempClassification As String
            strTempClassification = typTemp.Properties(lngClassification)
            If Trim(modMain.gstrSaveAsNewName) = "0" And _
               (strTempClassification = "Reference" Or _
                strTempClassification = "SubProduct" Or _
                strTempClassification = "LayOut") Then
                strTempFileName = typTemp.FileName
            Else
                '/ Classification
                strTempFileName = typTemp.Properties(lngFileName)
            End If

            strTempFileName = strTempFileName & "." & typTemp.Properties(lngFileType)
            
            If strFileName = strTempFileName Then
                fncCheckBeforeSave = "E030"
                Exit Function
            End If
CONTINUE2:
        Next j
        lngSaveCnt = lngSaveCnt + 1
CONTINUE1:
    Next i
    If lngSaveCnt = 0 Then
        fncCheckBeforeSave = "E048"
        Exit Function
    End If
End Function

Public Function fncSplitFileName(ByVal istrFileName As String) As String
    fncSplitFileName = ""
    
    On Error Resume Next
    Dim lstSplit As Variant
    lstSplit = Split(istrFileName, ".")
    
    Dim lngSize As Long
    lngSize = 0
    lngSize = UBound(lstSplit)
    On Error GoTo 0
    
    If lngSize <= 0 Then
        fncSplitFileName = istrFileName
        Exit Function
    End If
    
    Dim i As Long
    For i = 0 To lngSize - 1
        fncSplitFileName = fncSplitFileName & lstSplit(i)
    Next i
End Function

Public Function getSectionData(outColumnName As String, inColumnName As String, matchVal As String) As String
    getSectionData = ""
    If matchVal = "" Then Exit Function
    If inColumnName = "" Then Exit Function
    If outColumnName = "" Then Exit Function
    
    Dim db As Database
    Dim rsPLMsections As Recordset
    Set db = CurrentDb()
    Set rsPLMsections = db.OpenRecordset("SELECT * FROM tblPLMsection WHERE " & inColumnName & " = '" & matchVal & "'")
    
    getSectionData = rsPLMsections(outColumnName)
    
    rsPLMsections.Close
    Set rsPLMsections = Nothing
    Set db = Nothing

End Function

Public Function getPLMpropertyData(outColumnName As String, inColumnName As String, matchVal As String) As String
    getPLMpropertyData = ""
    If matchVal = "" Then Exit Function
    If inColumnName = "" Then Exit Function
    If outColumnName = "" Then Exit Function
    
    rsPLMprops.FindFirst inColumnName & " = '" & matchVal & "'"
    
    If rsPLMprops.noMatch Then Exit Function
    
    getPLMpropertyData = Nz(rsPLMprops(outColumnName), "")
End Function