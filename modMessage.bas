Option Explicit

Public Function Show(ByVal istrMessageID As String, Optional ByVal msgDetail As String = "") As Boolean
    Show = False
    
    Dim msgTypeID As String
    msgTypeID = Left$(istrMessageID, 1)
    
    Dim msgStyle As VbMsgBoxStyle
    msgStyle = vbOKOnly
    Select Case msgTypeID
        Case "I"
            msgStyle = vbInformation + vbOKOnly
        Case "Q"
            msgStyle = vbQuestion + vbOKCancel
        Case "W"
            msgStyle = vbExclamation + vbOKOnly
        Case "E"
            msgStyle = vbCritical + vbOKOnly
    End Select
    
    Dim msg As String
    msg = GetMessage(istrMessageID)
    If msgDetail <> "" Then msg = msg & vbCrLf & msgDetail
    If msg = "" Then Exit Function
    
    Dim rc As VbMsgBoxResult
    rc = MsgBox(msg, msgStyle, "Notice")
    If rc = vbYes Or rc = vbOK Then
        Show = True
    Else
        Show = False
    End If
End Function

Public Function GetMessage(ByVal istrMessageID As String, Optional ByVal istrReplace As String = "") As String

    GetMessage = ""
    
    Select Case istrMessageID
        Case "W003"
            GetMessage = "The drawing linked to 3D is already numbered."
        Case "E001"
            GetMessage = "Setting Sheet is insufficient."
        Case "E002"
            GetMessage = "There is an undefined setting."
        Case "E003"
            GetMessage = "CATIA is not running."
        Case "E004"
            GetMessage = "Please open only one of CATDrawing or CATProduct(CATPart) for execution."
        Case "E005"
            GetMessage = "The document is not open. Make sure you only have one Catia session open."
        Case "E006"
            GetMessage = "An error occurred while acquiring CATIA attribute."
        Case "E007"
            GetMessage = "It is not related to drawing and 3D."
        Case "E008"
            GetMessage = "An error occurred during EXCEL export processing."
        Case "E009"
            GetMessage = "An error occurred in the numbering process."
        Case "E010"
            GetMessage = "The configuration differs between 3D and EXCEL."
        Case "E011"
            GetMessage = "An error occurred during ""Save As"" processing."
        Case "E012"
            GetMessage = "Failed to check structure of differences."
        Case "E013"
            GetMessage = "Saving could not be performed because required property is insufficient."
        Case "E014"
            GetMessage = "CATIA data save failed."
        Case "E015"
            GetMessage = "Required properties were not listed in the title column of the Main sheet."
        Case "E016"
            GetMessage = "Required properties were not listed in the DefineDrawing sheet."
        Case "E017"
            GetMessage = "Excel configuration data acquisition failed."
        Case "E018"
            GetMessage = "CATIA data reflection failed."
        Case "E020"
            GetMessage = "I can not find the target line for number assignment."
        Case "E021"
            GetMessage = "DB connection failed. Please check the latest numbering macro version."
        Case "E022"
            GetMessage = "An attempt to add a row to the ModelID table has failed."
        Case "E023"
            GetMessage = "I failed to add a row to the numbering table."
        Case "E024"
            GetMessage = "I found some missing text."
        Case "E026"
            GetMessage = "Failed to create save directory."
        Case "E027"
            GetMessage = "I can not find the target line for CATDrawing linked file."
        Case "E029"
            GetMessage = "A blank Product Name is found."
        Case "E030"
            GetMessage = "Saving could not be performed because the file name is duplicated."
        Case "E031"
            GetMessage = "There is no Design_No on the drawing link."
        Case "E032"
            GetMessage = "Prohibited characters are included. Please execute [LOAD MODEL] and confirm."
        Case "E033"
            GetMessage = "A blank ModelID/DrawingID is found."
        Case "E034"
            GetMessage = "A blank " & "Design_No" & " is found."
        Case "E035"
            GetMessage = "Failed to get information from Numbering DB."
        Case "E036"
            GetMessage = "Below " & "ModelID/DrawingID" & " was not found on old numbering DB."
        Case "E037"
            GetMessage = "3DEX cache directory is not found." & vbCrLf & "Please check below cell."
        Case "E038"
            GetMessage = "A blank " & istrReplace & " is found."
        Case "E039"
            GetMessage = "Blank NUMBERING TABLE is listed on the DefineDevelopment sheet."
        Case "E040"
            GetMessage = "Blank OFFICE CODE is listed on the DefineDevelopment sheet."
        Case "E041"
            GetMessage = "Blank SECTION is listed on the DefineDevelopment sheet."
        Case "E042"
            GetMessage = "Wrong Selection is listed on the DefineDevelopment sheet." & vbCrLf & "Please list Selection with 1 or 0."
        Case "E043"
            GetMessage = "More than two Selection is listed 1 on the DefineDevelopment sheet." & vbCrLf & "Please list Selection 1 only once."
        Case "E046"
            GetMessage = "Blank File_Data_Name is found." & vbCrLf & "Please execute SET PROPERTY first."
        Case "E047"
            GetMessage = "A blank Current_Status is found."
        Case "E048"
            GetMessage = "There is no file to be saved."
        Case "E049"
            GetMessage = "There is invalid date format in Designed Date." & vbCrLf & "Make sure that the date is in ""dd/mm/yy hh:mm:ss""."
        Case "E050"
            GetMessage = "There is invalid date format in Revised Date." & vbCrLf & "Make sure that the date is in ""dd/mm/yy hh:mm:ss""."
        Case "E999"
            GetMessage = "Unknown error has occured."
        Case "Q001"
            GetMessage = "Attributes are written in number assignment Excel. Do you want to overwrite?"
        Case "Q002"
            GetMessage = "Are you sure you want to delete attribute information?"
        Case "Q003"
            GetMessage = "Prohibited characters are included. Do you want to replace the string?"
    End Select

    GetMessage = istrMessageID & " : " & GetMessage
End Function

Public Function Show2(ByVal istrMessageID As String, ByRef icurAddition() As String) As Boolean
    Show2 = False
    Dim msgDetail As String, newLine As String, lngCnt As Long, i As Integer

    On Error Resume Next
    lngCnt = UBound(icurAddition)
    On Error GoTo 0
    For i = 1 To lngCnt
        msgDetail = msgDetail & newLine & icurAddition(i)
        newLine = vbCrLf
    Next i
    
    Show2 = Show(istrMessageID, msgDetail)
End Function