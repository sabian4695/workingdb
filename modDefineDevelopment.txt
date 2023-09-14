Option Explicit

Private gcurSelection() As String
Private gcurSection() As String
Private gcurOfficeCode() As String
Private gcurNumberingTable() As String
Public gstrSection As String
Public gstrOfficeCode As String
Public gstrNumberingTable As String

Public Function fncRead() As Boolean
    fncRead = False

    Dim db As Database
    Dim rs1 As Recordset
    Set db = CurrentDb()
    Set rs1 = db.OpenRecordset("tblPLMsection", dbOpenSnapshot)

    Dim lngCnt As Long
    lngCnt = 0
    Do While Not rs1.EOF
        ReDim Preserve gcurSelection(lngCnt)
        ReDim Preserve gcurSection(lngCnt)
        ReDim Preserve gcurOfficeCode(lngCnt)
        ReDim Preserve gcurNumberingTable(lngCnt)

        gcurSelection(lngCnt) = rs1![Selection]
        gcurSection(lngCnt) = rs1![Section]
        gcurOfficeCode(lngCnt) = rs1![Office_Code]
        gcurNumberingTable(lngCnt) = rs1![Numbering_Table]
        lngCnt = lngCnt + 1
        rs1.MoveNext
    Loop

    rs1.Close
    Set rs1 = Nothing

    fncRead = True
End Function

Public Function fncCheck() As String
    fncCheck = ""
    
    Dim lngSelCnt As Long
    lngSelCnt = 0
    
    Dim lngCnt As Long
    On Error Resume Next
    lngCnt = UBound(gcurSelection)
    On Error GoTo 0
    Dim i As Long
    For i = 1 To lngCnt
        If Trim(gcurNumberingTable(i)) = "" Then
            fncCheck = "E039"
            Exit Function
        End If
        
        If Trim(gcurOfficeCode(i)) = "" Then
            fncCheck = "E040"
            Exit Function
        End If
        
        If Trim(gcurSection(i)) = "" Then
            fncCheck = "E041"
            Exit Function
        End If
        
        If gcurSelection(i) <> "1" And gcurSelection(i) <> "0" Then
            fncCheck = "E042"
            Exit Function
        ElseIf gcurSelection(i) = "1" Then
            lngSelCnt = lngSelCnt + 1
            gstrSection = gcurSection(i)
            gstrOfficeCode = gcurOfficeCode(i)
            gstrNumberingTable = gcurNumberingTable(i)
        End If
        
    Next i

    If lngSelCnt <> 1 Then
        fncCheck = "E043"
        Exit Function
    End If
End Function

Public Function fncLastRow() As Long
    fncLastRow = 22
End Function

Public Function fncGetSectionFromOfficeCode(ByVal istrCode As String) As String
    fncGetSectionFromOfficeCode = ""

    Dim lngCnt As Long
    On Error Resume Next
    lngCnt = UBound(gcurOfficeCode)
    On Error GoTo 0

    Dim i As Long
    For i = 1 To lngCnt
        If gcurOfficeCode(i) = istrCode Then
            If gcurSection(i) = VALUE_FCSEPWR Then
                fncGetSectionFromOfficeCode = VALUE_FCS
            Else
                fncGetSectionFromOfficeCode = gcurSection(i)
            End If
            Exit Function
        End If
    Next i
End Function

Public Function fncGetOfficeCodeFromSection(ByVal istrSection As String) As String
    fncGetOfficeCodeFromSection = ""

    Dim strSection As String
    If istrSection = VALUE_FCS Then
        strSection = VALUE_FCSEPWR
    Else
        strSection = istrSection
    End If
    
    Dim lngCnt As Long
    On Error Resume Next
    lngCnt = UBound(gcurSection)
    On Error GoTo 0

    Dim i As Long
    For i = 1 To lngCnt
        If gcurSection(i) = strSection Then
            fncGetOfficeCodeFromSection = gcurOfficeCode(i)
            Exit Function
        End If
    Next i
End Function

Public Sub Terminate()
    ReDim gcurSelection(0)
    ReDim gcurSection(0)
    ReDim gcurOfficeCode(0)
    ReDim gcurNumberingTable(0)
    gstrSection = ""
    gstrOfficeCode = ""
    gstrNumberingTable = ""
End Sub