Option Explicit

Private gcurExcelTitle() As String
Private gcurProperty() As String
Private gcurDrawingTextName() As String
Private gcurDrawingParamName() As String
Private gcurOldDBAttrName() As String
Private gcurDataType() As String
Private gcurInputRequired() As String
Private gcurInputDisabled() As String
Private gcurReplaceLine() As String
Private gcurDesignerAlias() As String
Private gcurDesignerName() As String

Public Function fncRead() As Boolean
    fncRead = False
    
    Dim db As Database
    Dim rs1 As Recordset
    Set db = CurrentDb()
    Set rs1 = db.OpenRecordset("tblPLMproperties", dbOpenSnapshot)

    Dim lngCnt As Long
    lngCnt = 0
    
    Do While Not rs1.EOF
            ReDim Preserve gcurExcelTitle(lngCnt)
            ReDim Preserve gcurProperty(lngCnt)
            ReDim Preserve gcurDrawingTextName(lngCnt)
            ReDim Preserve gcurDrawingParamName(lngCnt)
            ReDim Preserve gcurOldDBAttrName(lngCnt)
            ReDim Preserve gcurDataType(lngCnt)
            ReDim Preserve gcurInputRequired(lngCnt)
            ReDim Preserve gcurInputDisabled(lngCnt)
            ReDim Preserve gcurReplaceLine(lngCnt)
            
            gcurExcelTitle(lngCnt) = Nz(rs1![Form_Name], "")
            gcurProperty(lngCnt) = Nz(rs1![Property_Name], "")
            gcurDrawingTextName(lngCnt) = Nz(rs1![Drawing_Text_Name], "")
            gcurDrawingParamName(lngCnt) = Nz(rs1![Drawing_Parameter_Name], "")
            gcurDataType(lngCnt) = Nz(rs1![Data_Type], "")
            gcurInputRequired(lngCnt) = Nz(rs1![Input_Required], "")
            gcurInputDisabled(lngCnt) = Nz(rs1![Input_Disabled], "")
            gcurReplaceLine(lngCnt) = 1
        lngCnt = lngCnt + 1
        rs1.MoveNext
    Loop

    rs1.Close
    Set rs1 = Nothing
    fncRead = True
End Function

Public Function fncCheck1() As Boolean
    fncCheck1 = False
    
    Dim blnIsFileName As Boolean: blnIsFileName = False
    Dim blnIsFileType As Boolean: blnIsFileType = False
    
    Dim lngCnt As Long
    On Error Resume Next
    lngCnt = UBound(gcurExcelTitle)
    On Error GoTo 0
    
    Dim i As Integer
    For i = 1 To lngCnt
    
        If gcurExcelTitle(i) = "File_Data_Name" Then
            blnIsFileName = True
        ElseIf gcurExcelTitle(i) = "File_Data_Type" Then
            blnIsFileType = True
        End If
        
    Next i
    
    If blnIsFileName = False Or blnIsFileType = False Then
        Exit Function
    End If
    
    fncCheck1 = True
End Function

Public Function fncCheck2() As Boolean
    fncCheck2 = False

    Dim lngCnt As Long
    On Error Resume Next
    lngCnt = UBound(gcurExcelTitle)
    On Error GoTo 0
    
    Dim i As Integer
    For i = 1 To lngCnt
        If gcurDrawingTextName(i) = "" And gcurDrawingParamName(i) = "" Then
            Exit Function
        End If
    Next i
    
    fncCheck2 = True
End Function

Public Sub Terminate()
    ReDim gcurExcelTitle(0)
    ReDim gcurProperty(0)
    ReDim gcurDrawingTextName(0)
    ReDim gcurDrawingParamName(0)
    ReDim gcurOldDBAttrName(0)
    ReDim gcurInputRequired(0)
    ReDim gcurInputDisabled(0)
    ReDim gcurDataType(0)
    ReDim gcurReplaceLine(0)
    ReDim gcurDesignerAlias(0)
    ReDim gcurDesignerName(0)
End Sub

Public Function fncGetPropertyName(ByVal istrExcelTitle As String) As String
    fncGetPropertyName = ""

    Dim lnCnt As Long
    On Error Resume Next
    lnCnt = UBound(gcurExcelTitle)
    On Error GoTo 0

    Dim i As Long
    For i = 1 To lnCnt
        If gcurExcelTitle(i) = istrExcelTitle Then
            fncGetPropertyName = gcurProperty(i)
            Exit Function
        End If
    Next i
End Function

Public Function fncGetDrawingTextName(ByVal istrExcelTitle As String) As String
    fncGetDrawingTextName = ""

    Dim lnCnt As Long
    On Error Resume Next
    lnCnt = UBound(gcurExcelTitle)
    On Error GoTo 0

    Dim i As Long
    For i = 1 To lnCnt
        If gcurExcelTitle(i) = istrExcelTitle Then
            fncGetDrawingTextName = gcurDrawingTextName(i)
            Exit Function
        End If
    Next i
End Function

Public Function fncGetDrawingParamName(ByVal istrExcelTitle As String) As String
    fncGetDrawingParamName = ""

    Dim lngCnt As Long
    On Error Resume Next
    lngCnt = UBound(gcurExcelTitle)
    On Error GoTo 0

    Dim i As Long
    For i = 1 To lngCnt
        If gcurExcelTitle(i) = istrExcelTitle Then
            fncGetDrawingParamName = gcurDrawingParamName(i)
            Exit Function
        End If
    Next i
End Function

Public Function fncGetOldDBAttrName(ByVal istrExcelTitle As String) As String
    fncGetOldDBAttrName = ""

    Dim lngCnt As Long
    On Error Resume Next
    lngCnt = UBound(gcurExcelTitle)
    On Error GoTo 0

    Dim i As Long
    For i = 1 To lngCnt
        If gcurExcelTitle(i) = istrExcelTitle Then
            fncGetOldDBAttrName = gcurOldDBAttrName(i)
            Exit Function
        End If
    Next i
End Function

Public Function fncGetDataType(ByVal istrExcelTitle As String) As String
    fncGetDataType = ""

    Dim lngCnt As Long
    On Error Resume Next
    lngCnt = UBound(gcurExcelTitle)
    On Error GoTo 0

    Dim i As Long
    For i = 1 To lngCnt
        If gcurExcelTitle(i) = istrExcelTitle Then
            fncGetDataType = gcurDataType(i)
            Exit Function
        End If
    Next i
End Function

Public Function fncGetInputRequired(ByVal istrExcelTitle As String) As String
    fncGetInputRequired = ""

    Dim lngCnt As Long
    On Error Resume Next
    lngCnt = UBound(gcurExcelTitle)
    On Error GoTo 0

    Dim i As Long
    For i = 1 To lngCnt
        If gcurExcelTitle(i) = istrExcelTitle Then
            fncGetInputRequired = gcurInputRequired(i)
            Exit Function
        End If
    Next i
End Function

Public Function fncGetInputDisabled(ByVal istrExcelTitle As String) As String
    fncGetInputDisabled = ""

    Dim lngCnt As Long
    On Error Resume Next
    lngCnt = UBound(gcurExcelTitle)
    On Error GoTo 0

    Dim i As Long
    For i = 1 To lngCnt
        If gcurExcelTitle(i) = istrExcelTitle Then
            fncGetInputDisabled = gcurInputDisabled(i)
            Exit Function
        End If
    Next i
End Function

Public Function fncGetReplaceLine(ByVal istrExcelTitle As String) As String
    fncGetReplaceLine = ""

    Dim lngCnt As Long
    On Error Resume Next
    lngCnt = UBound(gcurExcelTitle)
    On Error GoTo 0

    Dim i As Long
    For i = 1 To lngCnt
        If gcurExcelTitle(i) = istrExcelTitle Then
            fncGetReplaceLine = gcurReplaceLine(i)
            Exit Function
        End If
    Next i
End Function

Public Function fncGetDesignerName(ByVal istrAlias As String, ByRef ostrName As String) As Boolean
    fncGetDesignerName = False

    Dim lngCnt As Long
    On Error Resume Next
    lngCnt = UBound(gcurDesignerAlias)
    On Error GoTo 0

    Dim i As Long
    For i = 1 To lngCnt
        If gcurDesignerAlias(i) = istrAlias Then
            ostrName = gcurDesignerName(i)
            fncGetDesignerName = True
            Exit Function
        End If
    Next i
End Function