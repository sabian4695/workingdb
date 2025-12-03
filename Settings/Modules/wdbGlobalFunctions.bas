Option Compare Database
Option Explicit

Public bClone As Boolean

Declare PtrSafe Sub ChooseColor Lib "msaccess.exe" Alias "#53" (ByVal hwnd As LongPtr, rgb As Long)

Public Function setTheme(setForm As Form)
On Error Resume Next

Dim scalarBack As Double, scalarFront As Double, darkMode As Boolean
Dim backBase As Long, foreBase As Long, colorLevels(4), backSecondary As Long, btnXback As Long, btnXbackShade As Long

'IF NO THEME SET, APPLY DEFAULT THEME (for Dev mode)
If Nz(TempVars!themePrimary, "") = "" Then
    TempVars.Add "themePrimary", 3355443
    TempVars.Add "themeSecondary", 0
    TempVars.Add "themeMode", "Dark"
    TempVars.Add "themeColorLevels", "1.3,1.6,1.9,2.2"
End If

darkMode = TempVars!themeMode = "Dark"

If darkMode Then
    foreBase = 16777215
    btnXback = 4342397
    scalarBack = 1.3
    scalarFront = 0.9
Else
    foreBase = 657930
    btnXback = 8947896
    scalarBack = 1.1
    scalarFront = 0.3
End If

backBase = CLng(TempVars!themePrimary)
backSecondary = CLng(TempVars!themeSecondary)

Dim colorLevArr() As String
colorLevArr = Split(TempVars!themeColorLevels, ",")

If backSecondary <> 0 Then
    colorLevels(0) = backBase
    colorLevels(1) = shadeColor(backSecondary, CDbl(colorLevArr(0)))
    colorLevels(2) = shadeColor(backBase, CDbl(colorLevArr(1)))
    colorLevels(3) = shadeColor(backSecondary, CDbl(colorLevArr(2)))
    colorLevels(4) = shadeColor(backBase, CDbl(colorLevArr(3)))
Else
    colorLevels(0) = backBase
    colorLevels(1) = shadeColor(backBase, CDbl(colorLevArr(0)))
    colorLevels(2) = shadeColor(backBase, CDbl(colorLevArr(1)))
    colorLevels(3) = shadeColor(backBase, CDbl(colorLevArr(2)))
    colorLevels(4) = shadeColor(backBase, CDbl(colorLevArr(3)))
End If

setForm.FormHeader.BackColor = colorLevels(findColorLevel(setForm.FormHeader.tag))
setForm.Detail.BackColor = colorLevels(findColorLevel(setForm.Detail.tag))
If Len(setForm.Detail.tag) = 4 Then
    setForm.Detail.AlternateBackColor = colorLevels(findColorLevel(setForm.Detail.tag) + 1)
Else
    setForm.Detail.AlternateBackColor = setForm.Detail.BackColor
End If

setForm.FormFooter.BackColor = colorLevels(findColorLevel(setForm.FormFooter.tag))

'assuming form parts don't use tags for other uses

Dim ctl As Control, eachBtn As CommandButton
Dim classColor As String, fadeBack, fadeFore
Dim Level
Dim backCol As Long, levFore As Double
Dim disFore As Double
Dim foreLevInt As Long, maxLev As Long

For Each ctl In setForm.Controls
    If ctl.tag Like "*.L#*" Then
        Level = findColorLevel(ctl.tag)
        backCol = colorLevels(Level)
    Else
        GoTo nextControl
    End If
    foreLevInt = Level
    If foreLevInt > 3 Then foreLevInt = 3
    maxLev = Level + 1
    If maxLev > 4 Then maxLev = 4
    
    If darkMode Then
        foreLevInt = Level
        If foreLevInt > 3 Then foreLevInt = 3
        levFore = (1 / colorLevArr(foreLevInt)) + 0.2
        disFore = 1.4 - levFore
    Else
        levFore = (colorLevArr(foreLevInt))
        disFore = 15 - levFore
    End If

    Select Case ctl.ControlType
        Case acCommandButton, acToggleButton 'OPTIONS: cardBtn.L#, cardBtnContrastBorder.L#, btn.L#
            If Not (ctl.tag Like "*btn*") Then GoTo skipAhead0
            ctl.BackColor = backCol
            
            If (ctl.Picture = "") Then GoTo skipAhead0
            If darkMode Then
                If InStr(ctl.Picture, "\Core_theme_light\") Then ctl.Picture = Replace(ctl.Picture, "\Core_theme_light\", "\Core\")
            Else
                If InStr(ctl.Picture, "\Core\") Then ctl.Picture = Replace(ctl.Picture, "\Core\", "\Core_theme_light\")
            End If
            
skipAhead0:
            Select Case True
                Case ctl.tag Like "*cardBtn.L#*"
                    ctl.BorderColor = backCol
                Case ctl.tag Like "*cardBtnContrastBorder.L#*"
                    If ctl.BorderStyle <> 0 Then ctl.BorderColor = colorLevels(maxLev)
                Case ctl.tag Like "*btn.L#*"
                    If ctl.BorderStyle <> 0 Then ctl.BorderColor = backCol
                    
                    'fade the colors
                    fadeBack = shadeColor(backCol, scalarBack)
                    fadeFore = shadeColor(foreBase, levFore - 0.2)
                    
                    ctl.ForeColor = foreBase
                    ctl.HoverColor = fadeBack
                    ctl.PressedColor = fadeBack
                    ctl.HoverForeColor = fadeFore
                    ctl.PressedForeColor = fadeFore
                Case ctl.tag Like "*btnDis.L#*" 'for disabled look
                    If ctl.BorderStyle <> 0 Then ctl.BorderColor = backCol
                    
                    'fade the colors
                    fadeBack = shadeColor(backCol, scalarBack)
                    fadeFore = shadeColor(foreBase, disFore)
                    
                    ctl.ForeColor = fadeFore
                    ctl.HoverColor = fadeBack
                    ctl.PressedColor = fadeBack
                    ctl.HoverForeColor = fadeFore
                    ctl.PressedForeColor = fadeFore
                Case ctl.tag Like "*btnDisContrastBorder.L#*" 'for disabled look
                    If ctl.BorderStyle <> 0 Then ctl.BorderColor = colorLevels(maxLev)
                    
                    'fade the colors
                    fadeBack = shadeColor(backCol, scalarBack)
                    fadeFore = shadeColor(foreBase, disFore)
                    
                    ctl.ForeColor = fadeFore
                    ctl.HoverColor = fadeBack
                    ctl.PressedColor = fadeBack
                    ctl.HoverForeColor = fadeFore
                    ctl.PressedForeColor = fadeFore
                Case ctl.tag Like "*btnXdis.L#*" 'for disabled look
                    If ctl.BorderStyle <> 0 Then ctl.BorderColor = btnXback
                    ctl.ForeColor = foreBase
                    ctl.BackColor = btnXback
                    
                    'fade the colors
                    fadeBack = shadeColor(btnXback, scalarBack)
                    fadeFore = shadeColor(foreBase, disFore)
                    btnXbackShade = shadeColor(btnXback, (0.1 * Level) + scalarBack)
                    
                    ctl.BackColor = btnXbackShade
                    ctl.ForeColor = fadeFore
                    ctl.HoverColor = fadeBack
                    ctl.PressedColor = fadeBack
                    ctl.HoverForeColor = fadeFore
                    ctl.PressedForeColor = fadeFore
                Case ctl.tag Like "*btnX.L#*"
                    If ctl.BorderStyle <> 0 Then ctl.BorderColor = btnXback
                    ctl.ForeColor = foreBase
                    'fade the colors
                    fadeBack = shadeColor(btnXback, scalarBack)
                    fadeFore = shadeColor(foreBase, scalarFront)
                    btnXbackShade = shadeColor(btnXback, (0.1 * Level) + scalarBack)
                    
                    ctl.BackColor = btnXbackShade
                    ctl.HoverColor = fadeBack
                    ctl.PressedColor = fadeBack
                    ctl.HoverForeColor = fadeFore
                    ctl.PressedForeColor = fadeFore
                Case ctl.tag Like "*btnXcontrastBorder.L#*"
                    If ctl.BorderStyle <> 0 Then ctl.BorderColor = colorLevels(maxLev)
                    ctl.ForeColor = foreBase
                    'fade the colors
                    fadeBack = shadeColor(btnXback, scalarBack)
                    fadeFore = shadeColor(foreBase, scalarFront)
                    btnXbackShade = shadeColor(btnXback, (0.1 * Level) + scalarBack)
                    
                    ctl.BackColor = btnXbackShade
                    ctl.HoverColor = fadeBack
                    ctl.PressedColor = fadeBack
                    ctl.HoverForeColor = fadeFore
                    ctl.PressedForeColor = fadeFore
                Case ctl.tag Like "*btnContrastBorder.L#*"
                    If ctl.BorderStyle <> 0 Then ctl.BorderColor = colorLevels(maxLev)
                    ctl.ForeColor = foreBase
                    
                    'fade the colors
                    fadeBack = shadeColor(backCol, scalarBack)
                    fadeFore = shadeColor(foreBase, scalarFront)
                    
                    ctl.HoverColor = fadeBack
                    ctl.PressedColor = fadeBack
                    ctl.HoverForeColor = fadeFore
                    ctl.PressedForeColor = fadeFore
            End Select
        Case acLabel
            Select Case True
               Case ctl.tag Like "*lbl.L#*"
                   ctl.ForeColor = shadeColor(foreBase, levFore)
               Case ctl.tag Like "*lbl_wBack.L#*"
                   ctl.ForeColor = shadeColor(foreBase, levFore)
                   ctl.BackColor = backCol
                   If ctl.BorderStyle <> 0 Then ctl.BorderColor = backCol
            End Select
        Case acTextBox, acComboBox 'OPTIONS: txt.L#, txtBackBorder.L#, txtContrastBorder.L#
            If ctl.tag Like "*txt*" Then
                ctl.BackColor = backCol
                ctl.ForeColor = foreBase
            End If
            
            If ctl.FormatConditions.count = 1 Then 'special case for null value conditional formatting. Typically this is used for placeholder values
                If ctl.FormatConditions.ITEM(0).Expression1 Like "*IsNull*" Then
                    ctl.FormatConditions.ITEM(0).BackColor = backCol
                    ctl.FormatConditions.ITEM(0).ForeColor = foreBase
                End If
            End If
            
            Select Case True
                Case ctl.tag Like "*txtBackBorder*"
                    If ctl.BorderStyle <> 0 Then ctl.BorderColor = backCol
                Case ctl.tag Like "*txtContrastBorder*"
                    If ctl.BorderStyle <> 0 Then ctl.BorderColor = colorLevels(maxLev)
                Case ctl.tag Like "*txtTransFore*"
                    ctl.ForeColor = backCol
            End Select
        Case acRectangle, acSubform 'OPTIONS: cardBox.L#, cardBoxContrastBorder.L#
            If Not ctl.name Like "sfrm*" Then ctl.BackColor = backCol
            Select Case True
                Case ctl.tag Like "*cardBox.L#*"
                    ctl.BorderColor = backCol
                Case ctl.tag Like "*cardBoxContrastBorder.L#*"
                    ctl.BorderColor = colorLevels(Level + 1)
            End Select
        Case acTabCtl 'OPTIONS: tab.L#, tabContrastBorder.L#
            If ctl.tag Like "*tab*" Then
                If Level = 0 Then
                    ctl.BackColor = colorLevels(Level + 0)
                    ctl.BorderColor = backCol
                    ctl.PressedColor = backCol
                    
                    'fade the colors
                    fadeBack = shadeColor(CLng(colorLevels(Level - 1)), scalarBack)
                    fadeFore = shadeColor(foreBase, levFore - 0.6)
                    
                    ctl.HoverColor = fadeBack
                    ctl.ForeColor = fadeFore
                    
                    ctl.HoverForeColor = foreBase
                    ctl.PressedForeColor = foreBase
                Else
                    ctl.BackColor = colorLevels(Level - 1)
                    ctl.BorderColor = backCol
                    ctl.PressedColor = backCol
                    
                    'fade the colors
                    fadeBack = shadeColor(CLng(colorLevels(Level - 1)), scalarBack)
                    fadeFore = shadeColor(foreBase, levFore)
                    
                    ctl.HoverColor = fadeBack
                    ctl.ForeColor = fadeFore
                    
                    ctl.HoverForeColor = foreBase
                    ctl.PressedForeColor = foreBase
                End If
            End If
            If ctl.tag Like "*contrastBorder*" Then
                ctl.BorderColor = colorLevels(maxLev)
            End If
        Case acImage 'OPTIONS: pic.L#
            If ctl.tag Like "*pic*" Then ctl.BackColor = backCol
    End Select
    
nextControl:
Next

Exit Function
Err_Handler:
    Call handleError("wdbGlobalFunctions", "setTheme", Err.DESCRIPTION, Err.Number)
End Function

Function findColorLevel(tagText As String) As Long
On Error GoTo Err_Handler

findColorLevel = 0
If tagText = "" Then Exit Function

findColorLevel = Mid(tagText, InStr(tagText, ".L") + 2, 1)

Exit Function
Err_Handler:
    Call handleError("wdbGlobalFunctions", "setTheme", Err.DESCRIPTION, Err.Number)
End Function

Function shadeColor(inputColor As Long, scalar As Double) As Long
On Error GoTo Err_Handler

Dim tempHex, ioR, ioG, ioB

tempHex = Hex(inputColor)

If tempHex = "0" Then tempHex = "111111"

If Len(tempHex) = 1 Then tempHex = "0" & tempHex
If Len(tempHex) = 2 Then tempHex = "0" & tempHex
If Len(tempHex) = 3 Then tempHex = "0" & tempHex
If Len(tempHex) = 4 Then tempHex = "0" & tempHex
If Len(tempHex) = 5 Then tempHex = "0" & tempHex

ioR = val("&H" & Mid(tempHex, 5, 2)) * scalar
ioG = val("&H" & Mid(tempHex, 3, 2)) * scalar
ioB = val("&H" & Mid(tempHex, 1, 2)) * scalar

'Debug.Print ioR & " "; ioG & " " & ioB

If ioR > 255 Then ioR = 255
If ioG > 255 Then ioG = 255
If ioB > 255 Then ioB = 255

If ioR < 0 Then ioR = 0
If ioG < 0 Then ioG = 0
If ioB < 0 Then ioB = 0

shadeColor = rgb(ioR, ioG, ioB)

Exit Function
Err_Handler:
    Call handleError("wdbGlobalFunctions", "shadeColor", Err.DESCRIPTION, Err.Number)
End Function

Public Function colorPicker(Optional lngColor As Long) As Long
On Error GoTo Err_Handler
    'Static lngColor As Long
    ChooseColor Application.hWndAccessApp, lngColor
    colorPicker = lngColor
Exit Function
Err_Handler:
    Call handleError("wdbGlobalFunctions", "colorPicker", Err.DESCRIPTION, Err.Number)
End Function

Function doStuff()

Dim db As Database
Set db = CurrentDb()

Dim rsAppr As Recordset, rsGates As Recordset, rsSteps As Recordset

Set rsGates = db.OpenRecordset("SELECT * FROM tblPartGateTemplate WHERE projectTemplateId = 6")

Do While Not rsGates.EOF
    'add each gate
    db.Execute "INSERT INTO tblPartGateTemplate(projectTemplateId,gateTitle) VALUES (16,'" & rsGates!gateTitle & "')"
    TempVars.Add "gateId", db.OpenRecordset("SELECT @@identity")(0).Value

    Set rsSteps = db.OpenRecordset("SELECT * FROM tblPartStepTemplate WHERE gateTemplateId = " & rsGates!recordId)
    Do While Not rsSteps.EOF
        'add each step
        db.Execute "INSERT INTO tblPartStepTemplate(gateTemplateId,pillarStep,title,stepActionId,documentType,responsible,duration,durationDays,indexOrder) VALUES (" & _
            TempVars!gateId & "," & rsSteps!pillarStep & ",'" & StrQuoteReplace(rsSteps!Title) & "'," & Nz(rsSteps!stepActionId, 0) & "," & Nz(rsSteps!documentType, 0) & ",'" & rsSteps!responsible & "'," & rsSteps!duration & "," & Nz(rsSteps!durationDays, 0) & "," & rsSteps!indexOrder & ")"
        TempVars.Add "stepId", db.OpenRecordset("SELECT @@identity")(0).Value
    
        Set rsAppr = db.OpenRecordset("SELECT * FROM tblPartStepTemplateApprovals WHERE stepTemplateId = " & rsSteps!recordId)
        Do While Not rsAppr.EOF
            'add each approval
            db.Execute "INSERT INTO tblPartStepTemplateApprovals(stepTemplateId,dept,reqLevel) VALUES (" & TempVars!stepId & ",'" & rsAppr!dept & "','" & rsAppr!reqLevel & "')"
            
            rsAppr.MoveNext
        Loop
        rsSteps.MoveNext
    Loop
    rsGates.MoveNext
Loop


Set db = Nothing

End Function

Function dbExecute(sql As String)
On Error GoTo Err_Handler

Dim db As Database
Set db = CurrentDb()

db.Execute sql

Set db = Nothing

Exit Function
Err_Handler:
    Call handleError("wdbGlobalFunctions", "dbExecute", Err.DESCRIPTION, Err.Number, sql)
End Function

Public Function nowString() As String
On Error GoTo Err_Handler

nowString = Format(Now(), "yyyymmddTHHmmss")

Exit Function
Err_Handler:
    Call handleError("wdbGlobalFunctions", "nowString", Err.DESCRIPTION, Err.Number)
End Function

Public Function registerWdbUpdates(table As String, ID As Variant, column As String, oldVal As Variant, newVal As Variant, Optional tag0 As String = "", Optional tag1 As Variant = "")
On Error GoTo Err_Handler

Dim sqlColumns As String, sqlValues As String

If (VarType(oldVal) = vbDate) Then oldVal = Format(oldVal, "mm/dd/yyyy")
If (VarType(newVal) = vbDate) Then newVal = Format(newVal, "mm/dd/yyyy")

Dim db As Database
Set db = CurrentDb()
Dim rs1 As Recordset
Set rs1 = db.OpenRecordset("tblWdbUpdateTracking")

If Len(oldVal) > 255 Then oldVal = Left(oldVal, 255)
If Len(newVal) > 255 Then newVal = Left(newVal, 255)

With rs1
    .addNew
        !tableName = table
        !tableRecordId = ID
        !updatedBy = Environ("username")
        !updatedDate = Now()
        !columnName = column
        !previousData = StrQuoteReplace(CStr(Nz(oldVal, "")))
        !newData = StrQuoteReplace(CStr(Nz(newVal, "")))
        !dataTag0 = StrQuoteReplace(tag0)
        !dataTag1 = StrQuoteReplace(tag1)
    .Update
    .Bookmark = .lastModified
End With

rs1.Close
Set rs1 = Nothing
Set db = Nothing

Exit Function
Err_Handler:
    Call handleError("wdbGlobalFunctions", "registerWdbUpdates", Err.DESCRIPTION, Err.Number, table & " " & ID)
End Function

Public Function registerSalesUpdates(table As String, ID As Variant, column As String, oldVal As Variant, newVal As Variant, Optional tag0 As String = "", Optional tag1 As Variant = "")
On Error GoTo Err_Handler

Dim sqlColumns As String, sqlValues As String

If (VarType(oldVal) = vbDate) Then oldVal = Format(oldVal, "mm/dd/yyyy")
If (VarType(newVal) = vbDate) Then newVal = Format(newVal, "mm/dd/yyyy")

Dim db As Database
Set db = CurrentDb()
Dim rs1 As Recordset
Set rs1 = db.OpenRecordset("tblSalesUpdateTracking")

With rs1
    .addNew
        !tableName = table
        !tableRecordId = ID
        !updatedBy = Environ("username")
        !updatedDate = Now()
        !columnName = column
        !previousData = StrQuoteReplace(CStr(Nz(oldVal, "")))
        !newData = StrQuoteReplace(CStr(Nz(newVal, "")))
        !dataTag0 = StrQuoteReplace(tag0)
        !dataTag1 = StrQuoteReplace(tag1)
    .Update
    .Bookmark = .lastModified
End With

rs1.Close
Set rs1 = Nothing
Set db = Nothing

Exit Function
Err_Handler:
    Call handleError("wdbGlobalFunctions", "registerSalesUpdates", Err.DESCRIPTION, Err.Number)
End Function

Function getFullName() As String
On Error GoTo Err_Handler

Dim db As Database
Set db = CurrentDb()
Dim rs1 As Recordset
Set rs1 = db.OpenRecordset("SELECT firstName, lastName FROM tblPermissions WHERE User = '" & Environ("username") & "'", dbOpenSnapshot)
getFullName = rs1!firstName & " " & rs1!lastName
rs1.Close: Set rs1 = Nothing
Set db = Nothing

Exit Function
Err_Handler:
    Call handleError("wdbGlobalFunctions", "getFullName", Err.DESCRIPTION, Err.Number)
End Function

Function generateHTML(Title As String, subTitle As String, primaryMessage As String, detail1 As String, detail2 As String, detail3 As String, Optional Link As String = "") As String
On Error GoTo Err_Handler

Dim tblHeading As String, tblFooter As String, strHTMLBody As String

If Link <> "" Then
    primaryMessage = "<a href = '" & Link & "'>" & primaryMessage & "</a>"
End If

tblHeading = "<table style=""width: 100%; margin: 0 auto; padding: 2em 3em; text-align: center; background-color: #fafafa;"">" & _
                            "<tbody>" & _
                                "<tr><td><h2 style=""color: #414141; font-size: 28px; margin-top: 0;"">" & Title & "</h2></td></tr>" & _
                                "<tr><td><p style=""color: rgb(73, 73, 73);"">" & subTitle & "</p></td></tr>" & _
                                 "<tr><td><table style=""padding: 1em; text-align: center;"">" & _
                                     "<tr><td style=""padding: 1em 1.5em; background: #FF6B00; "">" & primaryMessage & "</td></tr>" & _
                                "</table></td></tr>" & _
                            "</tbody>" & _
                        "</table>"
                        
tblFooter = "<table style=""width: 100%; margin: 0 auto; padding: 3em; background: #2b2b2b; color: rgba(255,255,255,.5);"">" & _
                        "<tbody>" & _
                            "<tr style=""border-collapse: collapse;""><td style=""padding: 1em; color: #c9c9c9;"">Details</td></tr>" & _
                            "<tr style=""border-collapse: collapse;""><td style=""padding: .1em 2em;"">" & detail1 & "</td></tr>" & _
                            "<tr style=""border-collapse: collapse;""><td style=""padding: .1em 2em;"">" & detail2 & "</td></tr>" & _
                            "<tr style=""border-collapse: collapse;""><td style=""padding: .1em 2em 1em 2em;"">" & detail3 & "</td></tr>" & _
                        "</tbody>" & _
                    "</table>"

strHTMLBody = "" & _
"<!DOCTYPE html><html lang=""en"" xmlns=""http://www.w3.org/1999/xhtml"" xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:o=""urn:schemas-microsoft-com:office:office"">" & _
    "<head><meta charset=""utf-8""><title>Working DB Notification</title></head>" & _
    "<body style=""margin: 0 auto; Font-family: 'Montserrat', sans-serif; font-weight: 400; font-size: 15px; line-height: 1.8;"">" & _
        "<table style=""max-width: 600px; margin: 0 auto; text-align: center; "">" & _
            "<tbody>" & _
                "<tr><td>" & tblHeading & "</td></tr>" & _
                "<tr><td>" & tblFooter & "</td></tr>" & _
                "<tr><td><p style=""color: rgb(192, 192, 192); text-align: center;"">This email was created by  &copy; workingDB</p></td></tr>" & _
            "</tbody>" & _
        "</table>" & _
    "</body>" & _
"</html>"

generateHTML = strHTMLBody

Exit Function
Err_Handler:
    Call handleError("wdbGlobalFunctions", "generateHTML", Err.DESCRIPTION, Err.Number)
End Function

Function emailContentGen(subject As String, Title As String, subTitle As String, primaryMessage As String, detail1 As String, detail2 As String, detail3 As String) As String
On Error GoTo Err_Handler

emailContentGen = subject & "," & Title & "," & subTitle & "," & primaryMessage & "," & detail1 & "," & detail2 & "," & detail3

Exit Function
Err_Handler:
    Call handleError("wdbGlobalFunctions", "emailContentGen", Err.DESCRIPTION, Err.Number)
End Function

Function idNAM(inputVal As Variant, typeVal As Variant) As Variant
On Error Resume Next 'just skip in case Oracle Errors
Dim db As Database
Set db = CurrentDb()
Dim rs1 As Recordset
idNAM = ""

If inputVal = "" Then Exit Function

If typeVal = "ID" Then
    Set rs1 = db.OpenRecordset("SELECT SEGMENT1 FROM APPS_MTL_SYSTEM_ITEMS WHERE INVENTORY_ITEM_ID = " & inputVal, dbOpenSnapshot)
    If rs1.RecordCount = 0 Then GoTo exitFunction
    idNAM = rs1("SEGMENT1")
End If

If typeVal = "NAM" Then
    Set rs1 = db.OpenRecordset("SELECT INVENTORY_ITEM_ID FROM APPS_MTL_SYSTEM_ITEMS WHERE SEGMENT1 = '" & inputVal & "'", dbOpenSnapshot)
    If rs1.RecordCount = 0 Then GoTo exitFunction
    idNAM = rs1("INVENTORY_ITEM_ID")
End If

exitFunction:
rs1.Close
Set rs1 = Nothing
Set db = Nothing
End Function

Function sendNotification(sendTo As String, notType As Integer, notPriority As Integer, desc As String, emailContent As String, Optional appName As String = "", Optional appId As Long, Optional multiEmail As Boolean = False, Optional customEmail As Boolean = False) As Boolean
sendNotification = True

On Error GoTo Err_Handler

Dim db As Database
Set db = CurrentDb()

'has this person been notified about this thing today already?
Dim rsNotifications As Recordset
Set rsNotifications = db.OpenRecordset("SELECT * from tblNotificationsSP WHERE recipientUser = '" & sendTo & "' AND notificationDescription = '" & StrQuoteReplace(desc) & "' AND sentDate > #" & Date - 1 & "#")
If rsNotifications.RecordCount > 0 Then
    If rsNotifications!notificationType = 1 Then
        Dim msgTxt As String
        If rsNotifications!senderUser = Environ("username") Then
            msgTxt = "You already nudged this person today"
        Else
            msgTxt = sendTo & " has already been nudged about this today by " & rsNotifications!senderUser & ". Let's wait until tomorrow to nudge them again."
        End If
        MsgBox msgTxt, vbInformation, "Hold on a minute..."
        sendNotification = False
        Exit Function
    End If
End If

Dim strEmail
If customEmail = False Then
    Dim ITEM, sendToArr() As String
    If multiEmail Then
        sendToArr = Split(sendTo, ",")
        strEmail = ""
        For Each ITEM In sendToArr
            strEmail = strEmail & getEmail(CStr(ITEM)) & ";"
        Next ITEM
        strEmail = Left(strEmail, Len(strEmail) - 1)
    Else
        strEmail = getEmail(sendTo)
    End If
Else
    strEmail = sendTo
    sendTo = Split(sendTo, "@")(0)
End If

Dim strValues
strValues = "'" & sendTo & "','" & strEmail & "','" & Environ("username") & "','" & getEmail(Environ("username")) & "','" & Now() & "'," & notType & "," & notPriority & ",'" & StrQuoteReplace(desc) & "','" & appName & "'," & appId & ",'" & StrQuoteReplace(emailContent) & "'"

db.Execute "INSERT INTO tblNotificationsSP(recipientUser,recipientEmail,senderUser,senderEmail,sentDate,notificationType,notificationPriority,notificationDescription,appName,appId,emailContent) VALUES(" & strValues & ");"

On Error Resume Next
rsNotifications.Close
Set rsNotifications = Nothing
Set db = Nothing

Exit Function
Err_Handler:
sendNotification = False
    Call handleError("wdbGlobalFunctions", "sendNotification", Err.DESCRIPTION, Err.Number)
End Function

Function privilege(pref) As Boolean
On Error GoTo Err_Handler

privilege = DLookup("[" & pref & "]", "[tblPermissions]", "[User] = '" & Environ("username") & "'")
    
Exit Function
Err_Handler:
    Call handleError("wdbGlobalFunctions", "privilege", Err.DESCRIPTION, Err.Number)
End Function

Function userData(data) As String
On Error GoTo Err_Handler

userData = Nz(DLookup("[" & data & "]", "[tblPermissions]", "[User] = '" & Environ("username") & "'"))

Exit Function
Err_Handler:
    Call handleError("wdbGlobalFunctions", "replaceDriveLetters", Err.DESCRIPTION, Err.Number)
End Function

Function restrict(userName As String, dept As String, Optional reqLevel As String = "", Optional orAbove As Boolean = False) As Boolean
On Error GoTo Err_Handler

If (CurrentProject.Path = "H:\dev") Then
    If userData("Developer") Then
        restrict = False
        Exit Function
    End If
End If

Dim db As Database
Set db = CurrentDb()
Dim d As Boolean, l As Boolean, rsPerm As Recordset
d = False
l = False

Set rsPerm = db.OpenRecordset("SELECT * FROM tblPermissions WHERE user = '" & userName & "'")
'restrict = true means you cannot access
'set No Access first, then allow as it is OK
d = True
l = True

If Nz(rsPerm!dept) = "" Or Nz(rsPerm("level")) = "" Then GoTo setRestrict 'if person isnt fully set up, do not allow access

If rsPerm!dept = dept Then d = False 'if correct department, set d to false

Select Case True 'figure out level
    Case reqLevel = "" 'if level isn't specified, this doesn't matter! - allow
        l = False
    Case rsPerm("level") = reqLevel 'if the level matches perfectly, allow
        l = False
    Case orAbove And reqLevel = "Supervisor" 'if supervisor and above check level and both supervisors and managers
        If rsPerm("level") = "Supervisor" Or rsPerm("level") = "Manager" Then l = False
    Case orAbove And reqLevel = "Engineer" 'if engineer and above, check level
        If rsPerm("level") = "Engineer" Or rsPerm("level") = "Supervisor" Or rsPerm("level") = "Manager" Then l = False
End Select

setRestrict:
restrict = d Or l

rsPerm.Close
Set rsPerm = Nothing
Set db = Nothing

Exit Function
Err_Handler:
    Call handleError("wdbGlobalFunctions", "restrict", Err.DESCRIPTION, Err.Number)
End Function

Function getEmail(userName As String) As String
On Error GoTo Err_Handler

getEmail = ""
On Error GoTo tryOracle
Dim db As Database
Set db = CurrentDb()
Dim rsPermissions As Recordset
Set rsPermissions = db.OpenRecordset("SELECT * from tblPermissions WHERE user = '" & userName & "'")
getEmail = Nz(rsPermissions!userEmail, "")
rsPermissions.Close
Set rsPermissions = Nothing

GoTo exitFunc

tryOracle:
Dim rsEmployee As Recordset
Set rsEmployee = db.OpenRecordset("SELECT FIRST_NAME, LAST_NAME, EMAIL_ADDRESS FROM APPS_XXCUS_USER_EMPLOYEES_V WHERE USER_NAME = '" & StrConv(userName, vbUpperCase) & "'")
getEmail = Nz(rsEmployee!EMAIL_ADDRESS, "")
rsEmployee.Close
Set rsEmployee = Nothing

exitFunc:
Set db = Nothing

Exit Function
Err_Handler:
    Call handleError("wdbGlobalFunctions", "getEmail", Err.DESCRIPTION, Err.Number)
End Function

Function splitString(a, b, c) As String
    On Error GoTo errorCatch
    splitString = Split(a, b)(c)
    Exit Function
errorCatch:
    splitString = ""
End Function

Public Function StrQuoteReplace(strValue)
On Error GoTo Err_Handler

StrQuoteReplace = Replace(Nz(strValue, ""), "'", "''")

Exit Function
Err_Handler:
    Call handleError("wdbGlobalFunctions", "StrQuoteReplace", Err.DESCRIPTION, Err.Number)
End Function

Public Function wdbEmail(ByVal strTo As String, ByVal strCC As String, ByVal strSubject As String, body As String) As Boolean
On Error GoTo Err_Handler
wdbEmail = True
Dim SendItems As New clsOutlookCreateItem
Set SendItems = New clsOutlookCreateItem

If IsNull(strCC) Then strCC = ""

SendItems.CreateMailItem sendTo:=strTo, _
                             CC:=strCC, _
                             subject:=strSubject, _
                             htmlBody:=body
    Set SendItems = Nothing
    
Exit Function
Err_Handler:
wdbEmail = False
    Call handleError("wdbGlobalFunctions", "wdbEmail", Err.DESCRIPTION, Err.Number)
End Function