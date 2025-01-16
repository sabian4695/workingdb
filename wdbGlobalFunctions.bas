Option Compare Database
Option Explicit

Public bClone As Boolean

Declare PtrSafe Sub ChooseColor Lib "msaccess.exe" Alias "#53" (ByVal hwnd As LongPtr, rgb As Long)

Declare PtrSafe Function LoadCursorFromFile Lib "user32" Alias "LoadCursorFromFileA" (ByVal lpFileName As String) As Long
Declare PtrSafe Function setCursor Lib "user32" Alias "SetCursor" (ByVal hCursor As Long) As Long

Function doStuff1()

'dbExecute ("UPDATE tblPartSteps SET stepActionId = 27 WHERE stepType = 'Upload WI in Q-Pulse' AND status <> 'Closed'")

End Function

Function doStuff()

Dim db As Database
Set db = CurrentDb()

Dim rs1 As Recordset
Set rs1 = db.OpenRecordset("")

Do While Not rs1.EOF
    

    rs1.MoveNext
Loop


Set db = Nothing

End Function

Function setCustomCursor()
Dim lngRet As Long
lngRet = LoadCursorFromFile("\\data\mdbdata\WorkingDB\Pictures\Theme_Pictures\cursor.cur")
lngRet = setCursor(lngRet)
End Function

Public Function setTheme(setForm As Form)
On Error Resume Next

Dim scalarBack As Double, scalarFront As Double, darkMode As Boolean
Dim backBase As Long, foreBase As Long, colorLevels(4), backSecondary As Long, btnXback As Long

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

For Each eachBtn In setForm.Controls
    
Next eachBtn

For Each ctl In setForm.Controls
    If ctl.tag Like "*.L#*" Then
        Level = findColorLevel(ctl.tag)
        backCol = colorLevels(Level)
    Else
        GoTo nextControl
    End If
    If darkMode Then
        levFore = (1 / colorLevArr(Level)) + 0.2
    Else
        levFore = colorLevArr(Level) * 7
    End If

    Select Case ctl.ControlType
        Case acCommandButton, acToggleButton 'OPTIONS: cardBtn.L#, cardBtnContrastBorder.L#, btn.L#
            If ctl.tag Like "*btn*" Then ctl.BackColor = backCol
            Select Case True
                Case ctl.tag Like "*cardBtn.L#*"
                    ctl.BorderColor = backCol
                Case ctl.tag Like "*cardBtnContrastBorder.L#*"
                    If ctl.BorderStyle <> 0 Then ctl.BorderColor = colorLevels(Level + 1)
                Case ctl.tag Like "*btn.L#*"
                    If ctl.BorderStyle <> 0 Then ctl.BorderColor = backCol
                    
                    'fade the colors
                    fadeBack = shadeColor(backCol, scalarBack)
                    fadeFore = shadeColor(foreBase, scalarFront)
                    
                    ctl.ForeColor = foreBase
                    ctl.HoverColor = fadeBack
                    ctl.PressedColor = fadeBack
                    ctl.HoverForeColor = fadeFore
                    ctl.PressedForeColor = fadeFore
                Case ctl.tag Like "*btnDis.L#*" 'for disabled look
                    If ctl.BorderStyle <> 0 Then ctl.BorderColor = backCol
                    
                    'fade the colors
                    fadeBack = shadeColor(backCol, scalarBack)
                    fadeFore = shadeColor(foreBase, levFore - 0.2)
                    
                    ctl.ForeColor = fadeFore
                    ctl.HoverColor = fadeBack
                    ctl.PressedColor = fadeBack
                    ctl.HoverForeColor = fadeFore
                    ctl.PressedForeColor = fadeFore
                Case ctl.tag Like "*btnDisContrastBorder.L#*" 'for disabled look
                    If ctl.BorderStyle <> 0 Then ctl.BorderColor = colorLevels(Level + 1)
                    
                    'fade the colors
                    fadeBack = shadeColor(backCol, scalarBack)
                    fadeFore = shadeColor(foreBase, levFore - 0.2)
                    
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
                    fadeFore = shadeColor(foreBase, levFore - 0.2)
                    
                    ctl.ForeColor = fadeFore
                    ctl.HoverColor = fadeBack
                    ctl.PressedColor = fadeBack
                    ctl.HoverForeColor = fadeFore
                    ctl.PressedForeColor = fadeFore
                Case ctl.tag Like "*btnX.L#*"
                    If ctl.BorderStyle <> 0 Then ctl.BorderColor = btnXback
                    ctl.ForeColor = foreBase
                    ctl.BackColor = btnXback
                    'fade the colors
                    fadeBack = shadeColor(btnXback, scalarBack)
                    fadeFore = shadeColor(foreBase, scalarFront)
                    
                    ctl.HoverColor = fadeBack
                    ctl.PressedColor = fadeBack
                    ctl.HoverForeColor = fadeFore
                    ctl.PressedForeColor = fadeFore
                Case ctl.tag Like "*btnXcontrastBorder.L#*"
                    If ctl.BorderStyle <> 0 Then ctl.BorderColor = colorLevels(Level + 1)
                    ctl.ForeColor = foreBase
                    ctl.BackColor = btnXback
                    'fade the colors
                    fadeBack = shadeColor(btnXback, scalarBack)
                    fadeFore = shadeColor(foreBase, scalarFront)
                    
                    ctl.HoverColor = fadeBack
                    ctl.PressedColor = fadeBack
                    ctl.HoverForeColor = fadeFore
                    ctl.PressedForeColor = fadeFore
                Case ctl.tag Like "*btnContrastBorder.L#*"
                    If ctl.BorderStyle <> 0 Then ctl.BorderColor = colorLevels(Level + 1)
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
                    If ctl.BorderStyle <> 0 Then ctl.BorderColor = colorLevels(Level + 1)
                Case ctl.tag Like "*txtTransFore*"
                    ctl.ForeColor = backCol
            End Select
        Case acRectangle, acSubform 'OPTIONS: cardBox.L#, cardBoxContrastBorder.L#
            ctl.BackColor = backCol
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
                ctl.BorderColor = colorLevels(Level + 1)
            End If
        Case acImage 'OPTIONS: pic.L#
            If ctl.tag Like "*pic*" Then ctl.BackColor = backCol
    End Select
    
nextControl:
Next

Exit Function
err_handler:
    Call handleError("wdbGlobalFunctions", "setTheme", Err.DESCRIPTION, Err.number)
End Function

Function findColorLevel(tagText As String) As Long
On Error GoTo err_handler

findColorLevel = 0
If tagText = "" Then Exit Function

findColorLevel = Mid(tagText, InStr(tagText, ".L") + 2, 1)

Exit Function
err_handler:
    Call handleError("wdbGlobalFunctions", "setTheme", Err.DESCRIPTION, Err.number)
End Function

Function shadeColor(inputColor As Long, scalar As Double) As Long
On Error GoTo err_handler

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

If ioR > 255 Then ioR = 255
If ioG > 255 Then ioG = 255
If ioB > 255 Then ioB = 255

shadeColor = rgb(ioR, ioG, ioB)

Exit Function
err_handler:
    Call handleError("wdbGlobalFunctions", "shadeColor", Err.DESCRIPTION, Err.number)
End Function

Public Function colorPicker() As Long
On Error GoTo err_handler
    Static lngColor As Long
    ChooseColor Application.hWndAccessApp, lngColor
    colorPicker = lngColor
Exit Function
err_handler:
    Call handleError("wdbGlobalFunctions", "colorPicker", Err.DESCRIPTION, Err.number)
End Function

Function dbExecute(sql As String)
On Error GoTo err_handler

Dim db As Database
Set db = CurrentDb()

db.Execute sql

Set db = Nothing

Exit Function
err_handler:
    Call handleError("wdbGlobalFunctions", "dbExecute", Err.DESCRIPTION, Err.number, sql)
End Function

Function findDescription(partNumber As String) As String
On Error GoTo err_handler

findDescription = ""

'first, check Oracle, then check SIFs

Dim db As Database
Dim rs1 As Recordset
Set db = CurrentDb
Set rs1 = db.OpenRecordset("SELECT SEGMENT1, DESCRIPTION FROM APPS_MTL_SYSTEM_ITEMS WHERE SEGMENT1 = '" & partNumber & "'", dbOpenSnapshot)
If rs1.RecordCount = 0 Then 'not in main Oracle table, now look through SIFs
    If DCount("[ROW_ID]", "APPS_Q_SIF_NEW_ASSEMBLED_PART_V", "[NIFCO_PART_NUMBER] = '" & partNumber & "'") > 0 Then 'is it in assy table?
        Set rs1 = db.OpenRecordset("SELECT SIFNUM, PART_DESCRIPTION FROM APPS_Q_SIF_NEW_ASSEMBLED_PART_V WHERE NIFCO_PART_NUMBER = '" & partNumber & "'", dbOpenSnapshot)
        rs1.MoveLast
        findDescription = rs1!Part_Description
    ElseIf DCount("[ROW_ID]", "APPS_Q_SIF_NEW_MOLDED_PART_V ", "[NIFCO_PART_NUMBER] = '" & partNumber & "'") > 0 Then 'is it in molded table?
        Set rs1 = db.OpenRecordset("SELECT SIFNUM, PART_DESCRIPTION FROM APPS_Q_SIF_NEW_MOLDED_PART_V WHERE NIFCO_PART_NUMBER = '" & partNumber & "'", dbOpenSnapshot)
        rs1.MoveLast
        findDescription = rs1!Part_Description
    End If
    Exit Function
End If

findDescription = rs1("DESCRIPTION")

rs1.Close
Set rs1 = Nothing

Exit Function
err_handler:
    Call handleError("wdbGlobalFunctions", "findDescription", Err.DESCRIPTION, Err.number)
End Function

Function gramsToLbs(gramsValue) As Double
On Error GoTo err_handler

gramsToLbs = gramsValue * 0.00220462

Exit Function
err_handler:
    Call handleError("wdbGlobalFunctions", "gramsToLbs", Err.DESCRIPTION, Err.number)
End Function

Function applyToAllForms()
Dim obj As AccessObject, dbs As Object
Set dbs = Application.CurrentProject
' Search for open AccessObject objects in AllForms collection.
For Each obj In dbs.AllForms
    If Left(obj.name, 1) = "f" Then
        If obj.name = "frmSearchHistory" Then GoTo nextOne
        DoCmd.OpenForm obj.name, acDesign
        
        If forms(obj.name).DefaultView = 1 Then
            forms(obj.name).BorderStyle = 2
        End If
        DoCmd.Close acForm, obj.name, acSaveYes
    End If
nextOne:
Next obj

End Function

Public Function exportSQL(sqlString As String, FileName As String)
On Error Resume Next
Dim db As Database
Set db = CurrentDb()
db.QueryDefs.Delete "myExportQueryDef"
On Error GoTo err_handler

Dim qExport As DAO.QueryDef
Set qExport = db.CreateQueryDef("myExportQueryDef", sqlString)

DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, "myExportQueryDef", FileName, True
If MsgBox("Export Complete. File path: " & FileName & vbNewLine & "Do you want to open this file?", vbYesNo, "Notice") = vbYes Then openPath (FileName)

db.QueryDefs.Delete "myExportQueryDef"

Set db = Nothing
Exit Function
err_handler:
    Call handleError("wdbGlobalFunctions", "exportSQL", Err.DESCRIPTION, Err.number)
End Function

Public Function nowString() As String
On Error GoTo err_handler

nowString = Format(Now(), "yyyymmddTHHmmss")

Exit Function
err_handler:
    Call handleError("wdbGlobalFunctions", "nowString", Err.DESCRIPTION, Err.number)
End Function

Public Function snackBox(sType As String, sTitle As String, sMessage As String, refForm As String, Optional centerBool As Boolean = False, Optional autoClose As Boolean = True)
On Error GoTo err_handler

TempVars.Add "snackType", sType
TempVars.Add "snackTitle", sTitle
TempVars.Add "snackMessage", sMessage
TempVars.Add "snackAutoClose", autoClose

If centerBool Then
    TempVars.Add "snackCenter", "True"
    TempVars.Add "snackLeft", forms(refForm).WindowLeft + forms(refForm).WindowWidth / 2 - 3393
    TempVars.Add "snackTop", forms(refForm).WindowTop + forms(refForm).WindowHeight / 2 - 500
Else
    TempVars.Add "snackCenter", "False"
    TempVars.Add "snackLeft", forms(refForm).WindowLeft + 200
    TempVars.Add "snackTop", forms(refForm).WindowTop + forms(refForm).WindowHeight - 1250
End If

DoCmd.OpenForm "frmSnack"

Exit Function
err_handler:
    Call handleError("wdbGlobalFunctions", "snackBox", Err.DESCRIPTION, Err.number)
End Function

Public Function labelUpdate(oldLabel As String)
On Error GoTo err_handler

Select Case True
    Case InStr(oldLabel, "-") <> 0
        labelUpdate = Replace(oldLabel, "-", ">")
    Case InStr(oldLabel, ">") <> 0
        labelUpdate = Replace(oldLabel, ">", "<")
    Case InStr(oldLabel, "<") <> 0
        labelUpdate = Replace(oldLabel, "<", "-")
End Select

Exit Function
err_handler:
    Call handleError("wdbGlobalFunctions", "labelUpdate", Err.DESCRIPTION, Err.number)
End Function

Public Function labelDirection(label As String)
On Error GoTo err_handler
If InStr(label, ">") <> 0 Then
    labelDirection = "DESC"
Else
    labelDirection = ""
End If
Exit Function
err_handler:
    Call handleError("wdbGlobalFunctions", "labelDirection", Err.DESCRIPTION, Err.number)
End Function

Public Function registerWdbUpdates(table As String, ID As Variant, column As String, oldVal As Variant, newVal As Variant, Optional tag0 As String = "", Optional tag1 As Variant = "")
On Error GoTo err_handler

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
err_handler:
    Call handleError("wdbGlobalFunctions", "registerWdbUpdates", Err.DESCRIPTION, Err.number, table & " " & ID)
End Function

Public Function registerSalesUpdates(table As String, ID As Variant, column As String, oldVal As Variant, newVal As Variant, Optional tag0 As String = "", Optional tag1 As Variant = "")
On Error GoTo err_handler

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
err_handler:
    Call handleError("wdbGlobalFunctions", "registerSalesUpdates", Err.DESCRIPTION, Err.number)
End Function

Function checkTime(whatIsHappening As String)

DoEvents
Debug.Print Format$((Timer - TempVars!tStamp) * 100!, "0.00 " & whatIsHappening)
TempVars.Add "tStamp", Timer

End Function

Public Function addWorkdays(dateInput As Date, daysToAdd As Long) As Date
On Error GoTo err_handler

Dim db As Database
Set db = CurrentDb()
Dim I As Long, testDate As Date, daysLeft As Long, rsHolidays As Recordset, intDirection
testDate = dateInput
daysLeft = Abs(daysToAdd)
intDirection = 1
If daysToAdd < 0 Then intDirection = -1

Set rsHolidays = db.OpenRecordset("tblHolidays")

Do While daysLeft > 0
    testDate = testDate + intDirection
    If Weekday(testDate) = 7 Or Weekday(testDate) = 1 Then ' IF WEEKEND -> skip
        testDate = testDate + intDirection
        GoTo skipDate
    End If
    
    rsHolidays.FindFirst "holidayDate = #" & testDate & "#"
    If Not rsHolidays.NoMatch Then GoTo skipDate ' IF HOLIDAY -> skip to next da

     daysLeft = daysLeft - 1
skipDate:
Loop

addWorkdays = testDate

On Error Resume Next
rsHolidays.Close
Set rsHolidays = Nothing
Set db = Nothing

Exit Function
err_handler:
    Call handleError("wdbGlobalFunctions", "addWorkdays", Err.DESCRIPTION, Err.number)
End Function

Public Function countWorkdays(oldDate As Date, newDate As Date) As Long
On Error GoTo err_handler

Dim total, sunday, saturday, weekdays, holidays

total = DateDiff("d", [oldDate], [newDate], vbSunday)
sunday = DateDiff("ww", [oldDate], [newDate], 1)
saturday = DateDiff("ww", [oldDate], [newDate], 7)
holidays = DCount("recordId", "tblHolidays", "holidayDate > #" & oldDate - 1 & "# AND holidayDate < #" & newDate & "#")
countWorkdays = total - sunday - saturday - holidays

Exit Function
err_handler:
    Call handleError("wdbGlobalFunctions", "countWorkdays", Err.DESCRIPTION, Err.number)
End Function

Function getFullName(Optional userName As String = "", Optional firstOnly As Boolean = False) As String
On Error GoTo err_handler

If userName = "" Then userName = Environ("username")

Dim db As Database
Set db = CurrentDb()
Dim rs1 As Recordset
Set rs1 = db.OpenRecordset("SELECT firstName, lastName FROM tblPermissions WHERE User = '" & userName & "'", dbOpenSnapshot)

If firstOnly Then
    getFullName = rs1!firstName
Else
    getFullName = rs1!firstName & " " & rs1!lastName
End If

rs1.Close: Set rs1 = Nothing
Set db = Nothing

Exit Function
err_handler:
    Call handleError("wdbGlobalFunctions", "getFullName", Err.DESCRIPTION, Err.number)
End Function

Function notificationsCount()
On Error Resume Next

Dim db As Database
Set db = CurrentDb()
Dim rsNoti As Recordset
Set rsNoti = db.OpenRecordset("SELECT count(ID) as unRead FROM tblNotificationsSP WHERE recipientUser = '" & Environ("username") & "' AND readDate is null")

Select Case rsNoti!unRead
    Case Is > 9
        Form_DASHBOARD.Form.notifications.Caption = "9+"
        Form_DASHBOARD.Form.notifications.BackColor = rgb(230, 0, 0)
    Case 0
        Form_DASHBOARD.Form.notifications.Caption = CStr(rsNoti!unRead)
        Form_DASHBOARD.Form.notifications.BackColor = rgb(60, 170, 60)
    Case Else
        Form_DASHBOARD.Form.notifications.Caption = CStr(rsNoti!unRead)
        Form_DASHBOARD.Form.notifications.BackColor = rgb(230, 0, 0)
End Select

rsNoti.Close
Set rsNoti = Nothing
Set db = Nothing

End Function

Function loadECOtype(changeNotice As String) As String
On Error GoTo err_handler

loadECOtype = ""

Dim db As Database
Set db = CurrentDb()
Dim rs1 As Recordset
Set rs1 = db.OpenRecordset("SELECT [CHANGE_ORDER_TYPE_ID] from ENG_ENG_ENGINEERING_CHANGES where [CHANGE_NOTICE] = '" & changeNotice & "'", dbOpenSnapshot)

If rs1.RecordCount > 0 Then loadECOtype = DLookup("[ECO_Type]", "[tblOracleDropDowns]", "[ECO_Type_ID]=" & rs1!CHANGE_ORDER_TYPE_ID)

rs1.Close
Set rs1 = Nothing
Set db = Nothing

Exit Function
err_handler:
    Call handleError("wdbGlobalFunctions", "loadECOtype", Err.DESCRIPTION, Err.number)
End Function

Function randomNumber(low As Long, high As Long) As Long
On Error GoTo err_handler

Randomize
randomNumber = Int((high - low + 1) * Rnd() + low)

Exit Function
err_handler:
    Call handleError("wdbGlobalFunctions", "randomNumber", Err.DESCRIPTION, Err.number)
End Function

Function getAPI(url, header1, header2)
On Error GoTo err_handler

Dim reader As New XMLHTTP60
    reader.open "GET", url, False
    reader.setRequestHeader header1, header2
    reader.send
        Do Until reader.ReadyState = 4
            DoEvents
        Loop
If reader.status = 200 Then
    getAPI = reader.responseText
Else
    MsgBox reader.status
End If

Exit Function
err_handler:
    Call handleError("wdbGlobalFunctions", "getAPI", Err.DESCRIPTION, Err.number)
End Function

Function generateHTML(Title As String, subTitle As String, primaryMessage As String, detail1 As String, detail2 As String, detail3 As String, Optional Link As String = "") As String
On Error GoTo err_handler

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
err_handler:
    Call handleError("wdbGlobalFunctions", "generateHTML", Err.DESCRIPTION, Err.number)
End Function

Function dailySummary(Title As String, subTitle As String, lates() As String, todays() As String, nexts() As String, lateCount As Long, todayCount As Long, nextCount As Long) As String
On Error GoTo err_handler

Dim tblHeading As String, tblStepOverview As String, strHTMLBody As String

tblHeading = "<table style=""width: 100%; margin: 0 auto; padding: 2em 2em 1em 2em; text-align: center; background-color: #fafafa;"">" & _
                            "<tbody>" & _
                                "<tr><td><h2 style=""color: #414141; font-size: 28px; margin-top: 0;"">" & Title & "</h2></td></tr>" & _
                                "<tr><td><p style=""color: rgb(73, 73, 73);"">Here is what you have happening...</p></td></tr>" & _
                            "</tbody>" & _
                        "</table>"
                        
Dim I As Long, lateTable As String, todayTable As String, nextTable As String, varStr As String, varStr1 As String, seeMore As String
seeMore = "<tr style=""border-collapse: collapse;""><td style=""padding: .1em 2em; font-style: italic;"" colspan=""3"">see the rest in the workingdb...</td></tr>"
I = 0
tblStepOverview = ""

varStr = ""
varStr1 = ""
If lates(0) <> "" Then
    For I = 0 To UBound(lates)
        lateTable = lateTable & "<tr style=""border-collapse: collapse;"">" & _
                                                "<td style=""padding: .1em 2em;"">" & Split(lates(I), ",")(0) & "</td>" & _
                                                "<td style=""padding: .1em 2em;"">" & Split(lates(I), ",")(1) & "</td>" & _
                                                "<td style=""padding: .1em 2em;  color: rgb(255,195,195);"">" & Split(lates(I), ",")(2) & "</td></tr>"
    Next I
    If lateCount > 1 Then varStr = "s"
    If lateCount > 15 Then varStr1 = seeMore
    tblStepOverview = tblStepOverview & "<table style=""width: 100%; margin: 0 auto; background: #2b2b2b; color: rgb(255,255,255);""><tr><th style=""padding: 1em; font-size: 20px; color: rgb(255,150,150); display: table-header-group;"" colspan=""3"">You have " & _
                                                                lateCount & " item" & varStr & " overdue</th></tr><tbody>" & _
                                                            "<tr style=""padding: .1em 2em;""><th style=""text-align: left"">Part#</th><th style=""text-align: left"">Item</th><th style=""text-align: left"">Due</th></tr>" & lateTable & varStr1 & "</tbody></table>"
End If

varStr = ""
varStr1 = ""
If todays(0) <> "" Then
    For I = 0 To UBound(todays)
        todayTable = todayTable & "<tr style=""border-collapse: collapse;"">" & _
                                                "<td style=""padding: .1em 2em;"">" & Split(todays(I), ",")(0) & "</td>" & _
                                                "<td style=""padding: .1em 2em;"">" & Split(todays(I), ",")(1) & "</td>" & _
                                                "<td style=""padding: .1em 2em;"">" & Split(todays(I), ",")(2) & "</td></tr>"
    Next I
    If todayCount > 1 Then varStr = "s"
    If todayCount > 15 Then varStr1 = seeMore
    tblStepOverview = tblStepOverview & "<table style=""width: 100%; margin: 0 auto; background: #2b2b2b; color: rgb(255,255,255);""><tr><th style=""padding: 1em; font-size: 20px; color: rgb(235,200,200); display: table-header-group;"" colspan=""3"">You have " & _
                                                                todayCount & " item" & varStr & " due today</th></tr><tbody>" & _
                                                            "<tr style=""padding: .1em 2em;""><th style=""text-align: left"">Part#</th><th style=""text-align: left"">Item</th><th style=""text-align: left"">Due</th></tr>" & todayTable & varStr1 & "</tbody></table>"
End If

varStr = ""
varStr1 = ""
If nexts(0) <> "" Then
    For I = 0 To UBound(nexts)
        nextTable = nextTable & "<tr style=""border-collapse: collapse;"">" & _
                                                "<td style=""padding: .1em 2em;"">" & Split(nexts(I), ",")(0) & "</td>" & _
                                                "<td style=""padding: .1em 2em;"">" & Split(nexts(I), ",")(1) & "</td>" & _
                                                "<td style=""padding: .1em 2em;"">" & Split(nexts(I), ",")(2) & "</td></tr>"
    Next I
    If nextCount > 1 Then varStr = "s"
    If nextCount > 15 Then varStr1 = seeMore
    tblStepOverview = tblStepOverview & "<table style=""width: 100%; margin: 0 auto; background: #2b2b2b; color: rgb(255,255,255);""><tr><th style=""padding: 1em; font-size: 20px; color: rgb(235,235,235); display: table-header-group;"" colspan=""3"">You have " & _
                                                                nextCount & " item" & varStr & " due soon</th></tr><tbody>" & _
                                                            "<tr style=""padding: .1em 2em;""><th style=""text-align: left"">Part#</th><th style=""text-align: left"">Item</th><th style=""text-align: left"">Due</th></tr>" & nextTable & varStr1 & "</tbody></table>"
End If

strHTMLBody = "" & _
"<!DOCTYPE html><html lang=""en"" xmlns=""http://www.w3.org/1999/xhtml"" xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:o=""urn:schemas-microsoft-com:office:office"">" & _
    "<head><meta charset=""utf-8""><title>Working DB Notification</title></head>" & _
    "<body style=""margin: 0 auto; Font-family: 'Montserrat', sans-serif; font-weight: 400; font-size: 15px; line-height: 1.8;"">" & _
        "<table style=""max-width: 600px; margin: 0 auto; text-align: center; "">" & _
            "<tbody>" & _
                "<tr><td>" & tblHeading & "</td></tr>" & _
                "<tr><td>" & tblStepOverview & "</td></tr>" & _
                "<tr><td><p style=""color: rgb(192, 192, 192); text-align: center;"">If you wish to no longer receive these emails,  go into your account menu in the workingDB to disable daily summary notifications</p></td></tr>" & _
                "<tr><td><p style=""color: rgb(192, 192, 192); text-align: center;"">This email was created by  &copy; workingDB</p></td></tr>" & _
            "</tbody>" & _
        "</table>" & _
    "</body>" & _
"</html>"

dailySummary = strHTMLBody

Exit Function
err_handler:
    Call handleError("wdbGlobalFunctions", "dailySummary", Err.DESCRIPTION, Err.number)
End Function

Function emailContentGen(subject As String, Title As String, subTitle As String, primaryMessage As String, detail1 As String, detail2 As String, detail3 As String) As String
On Error GoTo err_handler

emailContentGen = subject & "," & Title & "," & subTitle & "," & primaryMessage & "," & detail1 & "," & detail2 & "," & detail3

Exit Function
err_handler:
    Call handleError("wdbGlobalFunctions", "emailContentGen", Err.DESCRIPTION, Err.number)
End Function

Function sendNotification(sendTo As String, notType As Integer, notPriority As Integer, desc As String, emailContent As String, Optional appName As String = "", Optional appId As Long, Optional multiEmail As Boolean = False, Optional customEmail As Boolean = False) As Boolean
sendNotification = True

On Error GoTo err_handler

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
            If ITEM = "" Then GoTo nextItem
            strEmail = strEmail & getEmail(CStr(ITEM)) & ";"
nextItem:
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
err_handler:
sendNotification = False
    Call handleError("wdbGlobalFunctions", "sendNotification", Err.DESCRIPTION, Err.number)
End Function

Function privilege(pref) As Boolean
On Error GoTo err_handler

privilege = DLookup("[" & pref & "]", "[tblPermissions]", "[User] = '" & Environ("username") & "'")
    
Exit Function
err_handler:
    Call handleError("wdbGlobalFunctions", "privilege", Err.DESCRIPTION, Err.number)
End Function

Function userData(data) As String
On Error GoTo err_handler

userData = Nz(DLookup("[" & data & "]", "[tblPermissions]", "[User] = '" & Environ("username") & "'"))

Exit Function
err_handler:
    Call handleError("wdbGlobalFunctions", "replaceDriveLetters", Err.DESCRIPTION, Err.number)
End Function

Public Function getTotalPackingListWeight(packId As Long) As Double
On Error Resume Next
getTotalPackingListWeight = 0

Dim db As Database
Set db = CurrentDb()
Dim rs1 As Recordset
Set rs1 = db.OpenRecordset("SELECT sum(unitWeight*quantity) as total FROM tblPackListChild WHERE packListId = " & packId & " GROUP BY packListId")

getTotalPackingListWeight = rs1!total

rs1.Close
Set rs1 = Nothing
Set db = Nothing

End Function

Public Function getTotalPackingListCost(packId As Long) As Double
On Error Resume Next
getTotalPackingListCost = 0

Dim db As Database
Set db = CurrentDb()
Dim rs1 As Recordset
Set rs1 = db.OpenRecordset("SELECT sum(unitCost*quantity) as total FROM tblPackListChild WHERE packListId = " & packId & " GROUP BY packListId")

getTotalPackingListCost = rs1!total

rs1.Close
Set rs1 = Nothing
Set db = Nothing

End Function

Function restrict(userName As String, dept As String, Optional reqLevel As String = "", Optional orAbove As Boolean = False) As Boolean
On Error GoTo err_handler

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
err_handler:
    Call handleError("wdbGlobalFunctions", "restrict", Err.DESCRIPTION, Err.number)
End Function

Public Sub checkForFirstTimeRun()
On Error GoTo err_handler

Dim db As Database
Set db = CurrentDb()
Dim rsAnalytics As Recordset, rsRefreshReports As Recordset, rsSummaryEmail As Recordset

Set rsAnalytics = db.OpenRecordset("SELECT max(dateUsed) as anaDate from tblAnalytics WHERE module = 'firstTimeRun'")
If Not Format(rsAnalytics!anaDate, "mm/dd/yyyy") >= Format(Date, "mm/dd/yyyy") Then
    'if max date is today, then this has already ran.
    Call checkProgramEvents
    Call scanSteps("all", "firstTimeRun")
    db.Execute "INSERT INTO tblAnalytics (module,form,userName,dateUsed) VALUES ('firstTimeRun','Form_frmSplash','" & Environ("username") & "','" & Now() & "')"
End If

Set rsRefreshReports = db.OpenRecordset("SELECT max(dateUsed) as anaDate from tblAnalytics WHERE module = 'refreshReports'")
If Not Format(rsRefreshReports!anaDate, "mm/dd/yyyy") >= Format(Date, "mm/dd/yyyy") Then Call openPath("\\data\mdbdata\WorkingDB\build\Commands\Misc_Commands\refreshReports.vbs")

Set rsSummaryEmail = db.OpenRecordset("SELECT max(dateUsed) as anaDate from tblAnalytics WHERE module = 'summaryEmail'")
If Not Format(rsSummaryEmail!anaDate, "mm/dd/yyyy") >= Format(Date, "mm/dd/yyyy") Then Call openPath("\\data\mdbdata\WorkingDB\build\Commands\Misc_Commands\summaryEmail.vbs")

On Error Resume Next
rsAnalytics.Close: Set rsAnalytics = Nothing
rsRefreshReports.Close: Set rsRefreshReports = Nothing
rsSummaryEmail.Close: Set rsSummaryEmail = Nothing
Set db = Nothing

Exit Sub
err_handler:
    Call handleError("wdbGlobalFunctions", "checkForFirstTimeRun", Err.DESCRIPTION, Err.number)
End Sub

Function grabSummaryInfo(Optional specificUser As String = "") As Boolean
On Error GoTo err_handler

grabSummaryInfo = False

Dim db As Database
Set db = CurrentDb()
Dim rsPeople As Recordset, rsOpenSteps As Recordset, rsOpenWOs As Recordset, rsNoti As Recordset, rsAnalytics As Recordset
Dim lateSteps() As String, todaySteps() As String, nextSteps() As String
Dim li As Long, ti As Long, ni As Long
Dim strQry, ranThisWeek As Boolean

Set rsAnalytics = db.OpenRecordset("SELECT max(dateUsed) as anaDate from tblAnalytics WHERE module = 'firstTimeRun'")
ranThisWeek = Format(rsAnalytics!anaDate, "ww", vbMonday, vbFirstFourDays) = Format(Date, "ww", vbMonday, vbFirstFourDays)

strQry = ""
If specificUser <> "" Then strQry = " AND user = '" & specificUser & "'"

Set rsPeople = db.OpenRecordset("SELECT * from tblPermissions WHERE Inactive = False" & strQry)
    li = 0
    ti = 0
    ni = 0
    ReDim Preserve lateSteps(li)
    ReDim Preserve todaySteps(ti)
    ReDim Preserve nextSteps(ni)

Do While Not rsPeople.EOF 'go through every active person
    If rsPeople!notifications = 1 And specificUser = "" Then GoTo nextPerson 'this person wants no notifications
    If rsPeople!notifications = 2 And ranThisWeek And specificUser = "" Then GoTo nextPerson 'this person only wants weekly notifications
    
    li = 0
    ti = 0
    ni = 0
    Erase lateSteps, todaySteps, nextSteps
    ReDim lateSteps(li)
    ReDim todaySteps(ti)
    ReDim nextSteps(ni)

    Set rsOpenSteps = db.OpenRecordset("SELECT * from qryStepApprovalTracker " & _
                                "WHERE person = '" & rsPeople!User & "' AND dueDate <= Date()+7")
    
    Do While (Not rsOpenSteps.EOF And Not (ti > 15 And li > 15 And ni > 15))
        Select Case rsOpenSteps!dueDate
            Case Date 'due today
                If ti > 15 Then
                    ti = ti + 1
                    GoTo nextStep
                End If
                ReDim Preserve todaySteps(ti)
                todaySteps(ti) = rsOpenSteps!partNumber & "," & rsOpenSteps!Action & ",Today"
                ti = ti + 1
            Case Is < Date 'over due
                If li > 15 Then
                    li = li + 1
                    GoTo nextStep
                End If
                ReDim Preserve lateSteps(li)
                lateSteps(li) = rsOpenSteps!partNumber & "," & rsOpenSteps!Action & "," & Format(rsOpenSteps!dueDate, "mm/dd/yyyy")
                li = li + 1
            Case Is <= (Date + 7) 'due in next week
                If ni > 15 Then
                    ni = ni + 1
                    GoTo nextStep
                End If
                ReDim Preserve nextSteps(ni)
                nextSteps(ni) = rsOpenSteps!partNumber & "," & rsOpenSteps!Action & "," & Format(rsOpenSteps!dueDate, "mm/dd/yyyy")
                ni = ni + 1
        End Select
nextStep:
        rsOpenSteps.MoveNext
    Loop
    rsOpenSteps.Close
    Set rsOpenSteps = Nothing
    
    If ti + li + ni > 0 Then
        Set rsNoti = db.OpenRecordset("tblNotificationsSP")
        With rsNoti
            .addNew
            !recipientUser = rsPeople!User
            !recipientEmail = rsPeople!userEmail
            !senderUser = "workingDB"
            !senderEmail = "workingDB@us.nifco.com"
            !sentDate = Now()
            !readDate = Now()
            !notificationType = 9
            !notificationPriority = 2
            !notificationDescription = "Summary Email"
            !emailContent = StrQuoteReplace(dailySummary("Hi " & rsPeople!firstName, "Here is what you have going on...", lateSteps(), todaySteps(), nextSteps(), li, ti, ni))
            .Update
        End With
        rsNoti.Close
        Set rsNoti = Nothing
    End If
    
nextPerson:
    rsPeople.MoveNext
Loop
Set db = Nothing
grabSummaryInfo = True

Exit Function
err_handler:
    Call handleError("wdbGlobalFunctions", "grabSummaryInfo", Err.DESCRIPTION, Err.number)
End Function

Function checkProgramEvents() As Boolean
On Error GoTo err_handler

Dim db As Database
Set db = CurrentDb()

Dim rsProgram As Recordset, rsEvents As Recordset, rsWO As Recordset, rsComments As Recordset, rsPeople As Recordset, rsNoti As Recordset
Dim controlNum As Long, Comments As String, dueDate, body As String, strValues

dueDate = addWorkdays(Date, 5)

Set rsEvents = db.OpenRecordset("SELECT * from tblProgramEvents WHERE designWOcreated = False AND eventDate BETWEEN #" & Date & "# AND #" & Date + 50 & "#")
Set rsPeople = db.OpenRecordset("SELECT * from tblPermissions WHERE designWOid = 1 AND InActive = FALSE")

Do While Not rsEvents.EOF
    Set rsProgram = db.OpenRecordset("SELECT * from tblPrograms WHERE ID = " & rsEvents!programId)
    
    Set rsWO = db.OpenRecordset("dbo_tblDRS")
    rsWO.addNew
        With rsWO
            !Issue_Date = Date
            !Approval_Status = 1
            !Requester = "automated"
            !DR_Level = 1
            !Request_Type = 23
            !Design_Level = 4 'ETA
            !Due_Date = dueDate
            !Part_Number = "D8157"
            !Part_Description = "Program Review"
            !Model_Code = rsProgram!modelCode
        End With
    rsWO.Update
    
    controlNum = db.OpenRecordset("SELECT @@identity")(0).Value
    Comments = "'Hold program review for " & rsProgram!modelCode & " " & rsEvents!eventTitle & "'"
    
    db.Execute "INSERT INTO dbo_tblComments(Control_Number, Comments) VALUES(" & controlNum & "," & Comments & ")"
    
    body = emailContentGen("Program Review WO", "WO Notice", "WO Auto-Created for " & rsProgram!modelCode & " Program Review", "Event: " & rsEvents!eventTitle, "WO#" & controlNum, "Due: " & dueDate, "Sent On: " & CStr(Now()))
    
    rsEvents.Edit
    rsEvents!designWOcreated = True
    rsEvents.Update
    
    Set rsNoti = db.OpenRecordset("tblNotificationsSP")
    rsPeople.MoveFirst
    Do While Not rsPeople.EOF
        With rsNoti
            .addNew
            !recipientUser = rsPeople!User
            !recipientEmail = rsPeople!userEmail
            !senderUser = Environ("username")
            !senderEmail = getEmail(Environ("username"))
            !sentDate = Now()
            !notificationType = 10
            !notificationPriority = 2
            !notificationDescription = "WO Auto-Created for " & rsProgram!modelCode & " Program Review"
            !appName = "Design WO"
            !appId = controlNum
            !emailContent = body
            .Update
        End With
        rsPeople.MoveNext
    Loop
    
    rsNoti.Close
    Set rsNoti = Nothing
        
    rsProgram.Close
    Set rsProgram = Nothing
    
    rsEvents.MoveNext
Loop

rsEvents.Close
Set rsEvents = Nothing

rsPeople.Close
Set rsPeople = Nothing

Set db = Nothing

Exit Function
err_handler:
    Call handleError("wdbGlobalFunctions", "checkProgramEvents", Err.DESCRIPTION, Err.number)
End Function

Function getEmail(userName As String) As String
On Error GoTo err_handler

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
err_handler:
    Call handleError("wdbGlobalFunctions", "getEmail", Err.DESCRIPTION, Err.number)
End Function

Function splitString(a, B, c) As String
    On Error GoTo errorCatch
    splitString = Split(a, B)(c)
    Exit Function
errorCatch:
    splitString = ""
End Function

Function labelCycle(checkLabel As String, nameLabel As String, Optional controlSourceVal As String = "") As String()
On Error GoTo err_handler

    Dim returnVal(0 To 1) As String, sortLabel As String
    If controlSourceVal = "" Then
        sortLabel = nameLabel
    Else
        sortLabel = controlSourceVal
    End If
    Select Case True
        Case InStr(checkLabel, "-") > 0
            returnVal(0) = nameLabel & " >"
            returnVal(1) = sortLabel & " DESC"
        Case InStr(checkLabel, ">") > 0
            returnVal(0) = nameLabel & " <"
            returnVal(1) = sortLabel & " ASC"
        Case Else
            returnVal(0) = nameLabel & " -"
            returnVal(1) = sortLabel & " ASC"
    End Select
    labelCycle = returnVal
    
Exit Function
err_handler:
    Call handleError("wdbGlobalFunctions", "labelCycle", Err.DESCRIPTION, Err.number)
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

Function getDescriptionFromId(inventId As Long) As String
On Error GoTo err_handler

Dim db As Database
Set db = CurrentDb()
Dim rs1 As Recordset

getDescriptionFromId = ""
If IsNull(inventId) Then Exit Function
On Error Resume Next

Set rs1 = db.OpenRecordset("SELECT DESCRIPTION FROM APPS_MTL_SYSTEM_ITEMS WHERE INVENTORY_ITEM_ID = " & inventId, dbOpenSnapshot)
If rs1.RecordCount = 0 Then GoTo exitFunction
getDescriptionFromId = rs1("DESCRIPTION")

exitFunction:
rs1.Close
Set rs1 = Nothing
Set db = Nothing

Exit Function
err_handler:
    Call handleError("wdbGlobalFunctions", "getDescriptionFromId", Err.DESCRIPTION, Err.number)
End Function

Public Function StrQuoteReplace(strValue)
On Error GoTo err_handler

StrQuoteReplace = Replace(Nz(strValue, ""), "'", "''")

Exit Function
err_handler:
    Call handleError("wdbGlobalFunctions", "StrQuoteReplace", Err.DESCRIPTION, Err.number)
End Function

Public Function wdbEmail(ByVal strTo As String, ByVal strCC As String, ByVal strSubject As String, body As String) As Boolean
On Error GoTo err_handler
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
err_handler:
wdbEmail = False
    Call handleError("wdbGlobalFunctions", "wdbEmail", Err.DESCRIPTION, Err.number)
End Function

Function removeReferenceString(stringWithReference As String, Optional addBetween As String = "") As String
On Error GoTo err_handler

Dim tempString As String
tempString = stringWithReference

If InStr(stringWithReference, "(") Then tempString = Trim(Split(stringWithReference, "(")(0))
If InStr(stringWithReference, ")") Then tempString = tempString & addBetween & Trim(Split(stringWithReference, ")")(1))

removeReferenceString = tempString

Exit Function
err_handler:
    Call handleError("wdbGlobalFunctions", "removeReferenceString", Err.DESCRIPTION, Err.number)
End Function