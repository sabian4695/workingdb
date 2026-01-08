Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare PtrSafe Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare PtrSafe Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare PtrSafe Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

Dim CATIA As Object

Public Function addCatiaRefs()

'IN PROGRESS

Dim fso, oFile As Object

Set fso = CreateObject("Scripting.FileSystemObject")

For Each oFile In fso.GetFolder("").Files 'delete all temp files
    References.AddFromFile "C:\WINNT\system32\scrrun.dll"
Next

End Function

Private Function PixelTest(objPict As Object, ByVal X As Long, ByVal Y As Long) As Long
 Dim lDC As Variant
 lDC = CreateCompatibleDC(0)
 SelectObject lDC, objPict.Handle
 PixelTest = GetPixel(lDC, X, Y)
 DeleteDC lDC
End Function

Function formStatus(inWork As Boolean)

If inWork Then
    Me.Detail.tag = ".L4"
    Me.lblTitle.Caption = "Press Esc (in Catia) to exit the macro!"
Else
    Me.Detail.tag = ".L0"
    Me.lblTitle.Caption = "Only have one Catia open at a time!"
End If

Me.codeRunning.Visible = inWork

Call setTheme(Me)

End Function

Private Sub addDrawingItems_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)

If catiaOpen = False Then
    MsgBox "CATIA must be open to use this function.", vbExclamation, "CATIA Not Open"
    Exit Sub
End If

formStatus (True)

Dim rs As Recordset
Dim oSelection
Dim masterDrawing, alreadyOpen As Boolean
Dim detailSheet
Dim bool3DMaster As Boolean: bool3DMaster = False
Dim arr3DMaster
Dim myView
Dim myElement
Dim i As Long

Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblSessionVariables WHERE drawingItemSelect = TRUE", dbOpenSnapshot)
Set oSelection = CATIA.ActiveDocument.Selection

If InStr(Right(CATIA.ActiveDocument.name, 10), "CATDrawing") = 0 Then
    MsgBox "This function can only be run on a CATDrawing.", vbOKOnly & vbExclamation, "Not a CATDrawing"
    alreadyOpen = True
    GoTo exit_handler
End If

i = 1
Do While IsEmpty(masterDrawing)
    If CATIA.Documents.ITEM(i).name = "Internal_Master.CATDrawing" Then
        Set masterDrawing = CATIA.Documents.ITEM(i)
        alreadyOpen = True
    End If
    If i = CATIA.Documents.count And IsEmpty(masterDrawing) Then Set masterDrawing = CATIA.Documents.Read("\\design\data\Catia\nifcoV5\modelV5\Title_blocks\Internal_Master.CATDrawing")
    i = i + 1
Loop

If masterDrawing Is CATIA.ActiveDocument Then
    MsgBox "This function cannot be run on the NAM standard title block drawing." & vbCrLf & "Please activate a different drawing.", vbOKOnly & vbExclamation, "NAM Title Block Active"
    GoTo exit_handler
End If

i = 1
Do While IsEmpty(detailSheet)
    If CATIA.ActiveDocument.Sheets.ITEM(i).name = "Detail" Then Set detailSheet = CATIA.ActiveDocument.Sheets.ITEM(i)
    If i = CATIA.ActiveDocument.Sheets.count And IsEmpty(detailSheet) Then Set detailSheet = CATIA.ActiveDocument.Sheets.AddDetail("Detail")
    i = i + 1
Loop

oSelection.clear
Do While Not rs.EOF
    If rs!drawingItem = "3D Is Master" Then '3D Master symbol doesn't have it's own view in the NAM standard title block, so we have to build it from scratch
        bool3DMaster = True
    Else
        oSelection.Add masterDrawing.Sheets.ITEM("Detail").Views.ITEM((rs!drawingItem))
    End If
    rs.MoveNext
Loop

If oSelection.count > 0 Then
    oSelection.Copy
    oSelection.clear
    oSelection.Add detailSheet
    oSelection.Paste
End If

If bool3DMaster Then
    arr3DMaster = Split("0,13.77,7.5,16.5,7.5,16.5,15,13.77,15,13.77,7.5,11.04,7.5,11.04,0,13.77,0,13.77,0,4.77,7.5,11.04,7.5,2.04,15,13.77,15,4.77,0,4.77,7.5,2.04,7.5,2.04,15,4.77,15,8.5,17,8.5,17,8.5,17,0,17,0,5,0,5,0,5,2.95", ",")
    Set myView = detailSheet.Views.Add("3D Is Master")
    myView.Activate
    myView.X = 336
    myView.Y = 396

    i = 0
    Do While i < UBound(arr3DMaster)
        Set myElement = myView.Factory2D.CreateLine(arr3DMaster(i), arr3DMaster(i + 1), arr3DMaster(i + 2), arr3DMaster(i + 3))
        i = i + 4
    Loop

    oSelection.clear
    For i = 2 To myView.GeometricElements.count
        oSelection.Add myView.GeometricElements.ITEM(i)
    Next i
    oSelection.VisProperties.SetRealWidth 3, 1
    oSelection.clear

    myView.Texts.Add "3D", 0.785267665423703, 7.86563927710722
    myView.Texts.Add "2D", 11.9892676654234, 2.01363927710725

    myView.Texts.ITEM(1).SetFontName 1, 2, "SSS1"
    myView.Texts.ITEM(1).SetParameterOnSubString 7, 1, 2, 2800
    myView.Texts.ITEM(1).SetParameterOnSubString 14, 1, 2, 100
    myView.Texts.ITEM(1).SetParameterOnSubString 15, 1, 2, 62
    myView.Texts.ITEM(1).AnchorPosition = 2

    myView.Texts.ITEM(2).SetFontName 1, 2, "SSS1"
    myView.Texts.ITEM(2).SetParameterOnSubString 7, 1, 2, 2100
    myView.Texts.ITEM(2).SetParameterOnSubString 14, 1, 2, 100
    myView.Texts.ITEM(2).SetParameterOnSubString 15, 1, 2, 23
    myView.Texts.ITEM(2).AnchorPosition = 2
End If

exit_handler:
Set rs = Nothing
Set oSelection = Nothing
If Not alreadyOpen Then masterDrawing.CLOSE
Set masterDrawing = Nothing
Set detailSheet = Nothing
formStatus (False)
Set CATIA = Nothing

Exit Sub
Err_Handler:
    MsgBox Err.DESCRIPTION, vbOKOnly, "Error Code: " & Err.number
End Sub

Private Sub addParentheses_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)

If catiaOpen = False Then
    MsgBox "CATIA must be open to use this function.", vbExclamation, "CATIA Not Open"
    Exit Sub
End If

formStatus (True)

Dim oSelection
Dim InputObject(2)
Dim oStatus
Dim myObject
Dim oBefore, oAfter, oUpper, oLower

Set oSelection = CATIA.ActiveDocument.Selection

restart:
oSelection.clear
InputObject(0) = "DrawingText"
InputObject(1) = "DrawingDimension"
InputObject(2) = "DrawingComponent"
oStatus = oSelection.SelectElement2(InputObject, "Select a dimension or text box", True)

Do While oStatus <> "Cancel"
    Set myObject = oSelection.ITEM(1).Value
    Select Case typeName(myObject)
        Case "DrawingText"
            If (Left(myObject.Text, 1) = "(" And Right(myObject.Text, 1) = ")") = False Then myObject.Text = "(" & myObject.Text & ")"
        Case "DrawingDimension"
            Call myObject.getValue.GetBaultText(1, oBefore, oAfter, oUpper, oLower)
            If InStr(oBefore, "(") = 0 And InStr(oAfter, ")") = 0 Then Call myObject.getValue.SetBaultText(1, "(" & oBefore, oAfter & ")", oUpper, oLower)
        Case "DrawingComponent"
            If myObject.CompRef.Texts.count = 1 Then
                Set myObject = myObject.CompRef.Texts.ITEM(1)
                If (Left(myObject.Text, 1) = "(" And Right(myObject.Text, 1) = ")") = False Then myObject.Text = "(" & myObject.Text & ")"
            ElseIf myObject.CompRef.Texts.count > 1 Then
                MsgBox "This instantiated component has more than 1 text box." & vbCrLf & "Please select the individual text boxes directly.", vbExclamation, "Invalid Object Selected"
                GoTo restart
            Else
                MsgBox "This instantiated component has 0 text boxes.", vbExclamation, "Invalid Object Selected"
                GoTo restart
            End If
        Case Else
            MsgBox "This function only works on text boxes and dimensions.", vbExclamation, "Invalid Object Selected"
            GoTo restart
    End Select
    oSelection.clear
    oStatus = oSelection.SelectElement2(InputObject, "Select a text box or dimension", True)
Loop

exit_handler:
formStatus (False)
Set CATIA = Nothing

Exit Sub
Err_Handler:
    MsgBox Err.DESCRIPTION, vbOKOnly, "Error Code: " & Err.number
End Sub

Private Sub addSignatures_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)

If catiaOpen = False Then
    MsgBox "CATIA must be open to use this function.", vbExclamation, "CATIA Not Open"
    Exit Sub
End If

formStatus (True)

'First, check if DRS dash is open
If CurrentProject.AllForms("frmDRSdashboard").IsLoaded = False Then
    MsgBox "Need to have a WO open to use this", vbCritical, "No can do"
    GoTo exit_handler
End If

Dim oSelection
Dim Checker As String, approver As String
Dim chkDate As String, appDate As String
Dim mySheets, mySheet, myView, previousSheet
Dim chkTextBox, appTextBox
Dim chkDateTextBox, appDateTextBox
Dim myGeoElem, myDirection(1)
Dim myLine
Dim i As Long

Set oSelection = CATIA.ActiveDocument.Selection
oSelection.clear
Checker = Nz(Form_frmDRSdashboard.Checker1, "")
approver = Nz(Form_frmDRSdashboard.Checker2, "")
chkDate = Nz(Form_frmDRSdashboard.checker1SignDate, "")
appDate = Nz(Form_frmDRSdashboard.checker2SignDate, "")

If appDate = "" Then
    MsgBox "This WO has not been approved yet.", vbCritical, "Not Approved"
    GoTo exit_handler
End If

Set mySheets = CATIA.ActiveDocument.Sheets
For i = 1 To mySheets.count
    If mySheets.ITEM(i).name = "Internal Title Block" Then Set mySheet = mySheets.ITEM(i)
Next

If IsEmpty(mySheet) Then
    MsgBox "This function only works on the NAM standard drawing template.", vbCritical, "Internal Title Block Tab Not Found"
    GoTo exit_handler
End If
Set myView = mySheet.Views.ITEM(3)

For i = 1 To myView.Texts.count
    Select Case myView.Texts.ITEM(i).name
        Case "Checker"
            Set chkTextBox = myView.Texts.ITEM(i)
        Case "Approver"
            Set appTextBox = myView.Texts.ITEM(i)
        Case "Chk_Date"
            Set chkDateTextBox = myView.Texts.ITEM(i)
        Case "Resp_Date"
            Set appDateTextBox = myView.Texts.ITEM(i)
    End Select
Next

If Checker = "" Then
    'Hide the checker text boxes
    oSelection.Add chkTextBox
    oSelection.Add chkDateTextBox
    oSelection.VisProperties.SetShow 1
    oSelection.clear
    
    'Look for an existing line that crosses out the checker box
    For i = 1 To myView.GeometricElements.count
        Set myGeoElem = myView.GeometricElements.ITEM(i)
        If typeName(myGeoElem) = "Line2D" Then
            myGeoElem.GetDirection myDirection
            myDirection(0) = Format(Abs(myDirection(0)), "Standard")
            myDirection(1) = Format(Abs(myDirection(1)), "Standard")
            If myDirection(0) > 0.1 And myDirection(0) < 0.9 And myDirection(1) > 0.1 And myDirection(1) < 0.9 Then Set myLine = myGeoElem
        End If
    Next
    If IsEmpty(myLine) Then
        Set previousSheet = CATIA.ActiveDocument.Sheets.ActiveSheet
        mySheet.Activate
        Set myLine = myView.Factory2D.CreateLine(-135, 30, -160, 0)
        oSelection.Add myLine
        oSelection.VisProperties.SetRealWidth 1, 1
        oSelection.clear
        previousSheet.Activate
    Else
        oSelection.Add myLine
        oSelection.VisProperties.SetShow 0
        oSelection.clear
    End If
Else
    'Add checker signature and date, unhide if hidden
    chkTextBox.Text = UCase(getFullName(Checker))
    chkDateTextBox.Text = Format(chkDate, "mm/dd/yy")
    'Unhide if hidden
    oSelection.Add chkTextBox
    oSelection.Add chkDateTextBox
    oSelection.VisProperties.SetShow 0
    oSelection.clear
End If

'Add approver signature and date
appTextBox.Text = UCase(getFullName(approver))
appDateTextBox.Text = Format(appDate, "mm/dd/yy")
'Unhide if hidden
oSelection.Add appTextBox
oSelection.Add appDateTextBox
oSelection.VisProperties.SetShow 0
oSelection.clear

exit_handler:
formStatus (False)
Set CATIA = Nothing

Exit Sub
Err_Handler:
    MsgBox Err.DESCRIPTION, vbOKOnly, "Error Code: " & Err.number
End Sub

Private Sub autoNumberNotes_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)

If catiaOpen = False Then
    MsgBox "CATIA must be open to use this function.", vbExclamation, "CATIA Not Open"
    Exit Sub
End If

formStatus (True)

Dim oSelection
Dim InputObject(0)
Dim oStatus
Dim Notes
Dim lineArr() As String
Dim i As Long, j As Long, k As Long
Dim firstChar As Long

Set oSelection = CATIA.ActiveDocument.Selection
oSelection.clear
InputObject(0) = "DrawingText"
oStatus = oSelection.SelectElement2(InputObject, "Select the NOTES box", True)
If oStatus = "Cancel" Then GoTo exit_handler
Set Notes = oSelection.ITEM(1).Value

lineArr = Split(Notes.Text, Chr$(10))
j = 1
For i = 0 To UBound(lineArr)
    If IsNumeric(Left(lineArr(i), 3)) Or InStr(Left(lineArr(i), 3), "X") Then
        For k = 3 To Len(lineArr(i))
            If Mid(lineArr(i), k, 1) Like "[A-Z]" Then
                firstChar = k
                Exit For
            End If
        Next k
        
        If j < 10 Then
            lineArr(i) = " " & j & ". " & Right(lineArr(i), Len(lineArr(i)) - firstChar + 1)
        Else
            lineArr(i) = j & ". " & Right(lineArr(i), Len(lineArr(i)) - firstChar + 1)
        End If
        j = j + 1
    End If
Next i

Notes.Text = Join(lineArr, vbCrLf)
oSelection.clear

exit_handler:
formStatus (False)
Set CATIA = Nothing

Exit Sub
Err_Handler:
    MsgBox Err.DESCRIPTION, vbOKOnly, "Error Code: " & Err.number
End Sub

Private Sub btnAnchor_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)

If catiaOpen = False Then
    MsgBox "CATIA must be open to use this function.", vbExclamation, "CATIA Not Open"
    Exit Sub
End If

formStatus (True)

DoCmd.OpenForm "frm3DTextAnchor"

exit_handler:
formStatus (False)
Set CATIA = Nothing

Exit Sub
Err_Handler:
    MsgBox Err.DESCRIPTION, vbOKOnly, "Error Code: " & Err.number
End Sub

Private Sub checkFontSize_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)

If catiaOpen = False Then
    MsgBox "CATIA must be open to use this function.", vbExclamation, "CATIA Not Open"
    Exit Sub
End If

formStatus (True)

Dim oSelection
Dim visDimensions As Object
Dim visTexts As Object
Dim fontSizeDict As Object
Dim searchAttempts As Long
Dim tempText
Dim myDimension
Dim myDimValue
Dim myFontSize As Double
Dim dimCollection As Collection
Dim myText
Dim fontSizes
Dim maxCount As Double
Dim mostCommonFont As Double
Dim message As String
Dim i As Long, j As Long, k As Long

Set oSelection = CATIA.ActiveDocument.Selection
Set visDimensions = New Collection
Set visTexts = New Collection
Set fontSizeDict = CreateObject("Scripting.Dictionary")

CATIA.ActiveWindow.ActiveViewer.Reframe

oSelection.clear
On Error GoTo error_loop
error_loop:
searchAttempts = searchAttempts + 1
If searchAttempts > 2 Then
    On Error GoTo Err_Handler
End If
oSelection.Search "Type='Dimension'+'Text',scr"
On Error GoTo Err_Handler
For i = 1 To oSelection.Count2
    If typeName(oSelection.ITEM(i).Value) = "DrawingDimension" Then
        visDimensions.Add oSelection.ITEM(i).Value
    End If
    If typeName(oSelection.ITEM(i).Value) = "DrawingText" Then
        visTexts.Add oSelection.ITEM(i).Value
    End If
Next i
oSelection.clear

Set tempText = CATIA.ActiveDocument.Sheets.ActiveSheet.Views.ITEM(3).Texts.Add("TEMP", 0, 0)
oSelection.Add tempText
oSelection.VisProperties.SetShow 1
oSelection.clear
For i = 1 To visDimensions.count
    Set myDimension = visDimensions.ITEM(i)
    Set myDimValue = myDimension.getValue
    oSelection.clear
    oSelection.Add tempText
    CATIA.StartCommand ("Copy Object Format")
    If myDimension.DimType = 4 Then
        oSelection.Search "Drafting.Dimension.Name='" & myDimension.name & "'&Drafting.Dimension.Value='" & myDimValue.Value * (180 / (4 * Atn(1))) & "deg',scr"
    Else
        oSelection.Search "Drafting.Dimension.Name='" & myDimension.name & "'&Drafting.Dimension.Value='" & Format(myDimValue.Value, "Standard") & "mm',scr"
    End If
    If oSelection.count = 1 Then
        myFontSize = tempText.GetFontSize(0, 0)
        If Not fontSizeDict.exists(myFontSize) Then
            Set dimCollection = New Collection
            Set fontSizeDict(myFontSize) = dimCollection
        End If
        fontSizeDict(myFontSize).Add myDimension
    End If
    oSelection.clear
Next i
oSelection.clear
CATIA.ActiveDocument.Sheets.ActiveSheet.Views.ITEM(3).Texts.remove (CATIA.ActiveDocument.Sheets.ActiveSheet.Views.ITEM(3).Texts.count)

For i = 1 To visTexts.count
    Set myText = visTexts.ITEM(i)
    If myText.GetParameterOnSubString(0, 1, 1) = 0 And myText.GetParameterOnSubString(2, 1, 1) = 0 Then
        myFontSize = myText.GetFontSize(0, 0)
        If Not fontSizeDict.exists(myFontSize) Then
            Set dimCollection = New Collection
            Set fontSizeDict(myFontSize) = dimCollection
        End If
        fontSizeDict(myFontSize).Add myText
    End If
Next i

fontSizes = fontSizeDict.Keys
For i = 0 To UBound(fontSizes)
    If fontSizeDict(fontSizes(i)).count > maxCount Then
        maxCount = fontSizeDict(fontSizes(i)).count
        mostCommonFont = fontSizes(i)
    End If
Next i

For i = 0 To UBound(fontSizes)
    If fontSizes(i) <> mostCommonFont Then
        For j = 1 To fontSizeDict(fontSizes(i)).count
            oSelection.Add fontSizeDict(fontSizes(i)).ITEM(j)
        Next j
    End If
Next i
oSelection.VisProperties.SetRealColor 255, 0, 0, 1

For i = 0 To UBound(fontSizes)
    If fontSizeDict(fontSizes(i)).count = 1 Then
        message = message & fontSizeDict(fontSizes(i)).count & " text with font size " & fontSizes(i) & vbCrLf
    Else
        message = message & fontSizeDict(fontSizes(i)).count & " texts with font size " & fontSizes(i) & vbCrLf
    End If
Next i

If oSelection.count = 0 Then
    MsgBox "No inconsistencies found." & vbCrLf & "Font size is: " & mostCommonFont, vbOKOnly & vbInformation, "No Inconsistencies Found"
Else
    MsgBox message, vbOKOnly & vbExclamation, oSelection.count & " Inconsistencies Found"
End If

exit_handler:
formStatus (False)
Set CATIA = Nothing

Exit Sub
Err_Handler:
    MsgBox Err.DESCRIPTION, vbOKOnly, "Error Code: " & Err.number
End Sub

Private Sub clearSelection_Click()
On Error GoTo Err_Handler

Dim rs As Recordset
Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblSessionVariables WHERE drawingItemSelect = TRUE")

Do While Not rs.EOF
    With rs
        .Edit
        !drawingItemSelect = False
        .Update
        .MoveNext
    End With
Loop

resetDrawingItemFilter
Set rs = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub resetDrawingItemFilter()
Me.itemSearch.Value = ""
Me.drawingItems.Form.filter = ""
Me.drawingItems.Form.FilterOn = True
Me.itemImage.Picture = ""

If Me.drawingItems.Form.ScrollBars <> 2 Then
    Me.drawingItems.Form.ScrollBars = 2
    Me.drawingItems.Form.itemName.Width = 1380
    Me.drawingItems.Form.itemLabel.Width = 1380
    Me.drawingItems.Form.Width = 1560
End If
End Sub

Private Sub cmbFont_AfterUpdate()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)

Me.cmbFont.FontName = Me.cmbFont.column(1)
Me.txtInput.FontName = Me.cmbFont.column(1)

Exit Sub
Err_Handler:
    MsgBox Err.DESCRIPTION, vbOKOnly, "Error Code: " & Err.number
End Sub

Private Sub create3DText_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)

If catiaOpen = False Then
    MsgBox "CATIA must be open to use this function.", vbExclamation, "CATIA Not Open"
    Exit Sub
End If

formStatus (True)
Dim objShell As Object
Set objShell = CreateObject("Wscript.Shell")
'objShell.Run "\\data\mdbdata\WorkingDB\_docs\Catia_Custom_Files\Macro_Data\Scripts\start.CATScript"

Dim inputText As String, inputTextArr() As String
Dim charHeight As Double, charSpacing As Double, charAngle As Double
Dim rs As Recordset
Dim myDictionary As Object
Dim myPart
Dim mySketch
Dim pointCount As Long
Dim oSelection
Dim InputObject(0)
Dim oStatus
Dim myPoint
Dim myPointCoords(1)
Dim referenceLine
Dim refLineDirection(1)
Dim response
Dim textWidth As Double, charWidth As Double
Dim leftPoint As Double, rightPoint As Double
Dim mirror As Long
Dim anchorX As Double, anchorY As Double
Dim tempPoint(1) As Double
Dim firstPoint(1) As Double
Dim startPoint(1) As Double, midPoint(1) As Double, endPoint(1) As Double
Dim vector1(1) As Double, vector2(1) As Double
Dim magnitude1 As Double, magnitude2 As Double
Dim X As Double, myAngle As Double
Dim splinePoints()
Dim mySpline
Dim myLine
Dim i As Long, j As Long

inputText = Me.txtInput
charHeight = Me.txtHeight
charSpacing = Me.txtSpacing

'Assign selected font dictionary
Set rs = CurrentDb.OpenRecordset("SELECT * from tblCatiaFont WHERE fontName = " & Me.cmbFont, dbOpenSnapshot)
Set myDictionary = CreateObject("Scripting.Dictionary")
Do While Not rs.EOF
    myDictionary((rs!characterName)) = Split(rs!characterArray, ",")
    rs.MoveNext
Loop

'Create array of characters based on input text
If inputText = "" Then
    MsgBox "Please enter some text.", vbOKOnly & vbExclamation, "No Text Entered"
    GoTo exit_handler
End If
ReDim inputTextArray(Len(inputText) - 1)
For i = 0 To UBound(inputTextArray)
    inputTextArray(i) = Mid(inputText, i + 1, 1)
    If myDictionary.exists(inputTextArray(i)) = False Then
        If inputTextArray(i) <> " " Then
            MsgBox "The character '" & inputTextArray(i) & "' cannot be used.", vbOKOnly & vbExclamation, "Invalid Character"
            GoTo exit_handler
        End If
    End If
Next i

'Get active sketch
Set myPart = CATIA.ActiveDocument.Part
Set mySketch = myPart.InWorkObject
If typeName(mySketch) <> "Sketch" Then
    MsgBox "Please create/activate a sketch for the text to be added.", vbOKOnly & vbExclamation, "No Active Sketch"
    GoTo exit_handler
End If

'Look for points in the active sketch
For i = 1 To mySketch.GeometricElements.count
    If typeName(mySketch.GeometricElements.ITEM(i)) = "Point2D" Then
        GoTo select_point
    End If
Next i

'If no points are found in the active sketch
MsgBox "A reference point is required." & vbCrLf & "Please create a point and try again.", vbOKOnly & vbExclamation, "No Points Found"
GoTo exit_handler

'If more than one point is found in the active sketch, ask the user to select one
select_point:
Set oSelection = CATIA.ActiveDocument.Selection
oSelection.clear
InputObject(0) = "Point2D"
oStatus = oSelection.SelectElement2(InputObject, "Select a reference point", True)
If oStatus = "Cancel" Then GoTo exit_handler
Set myPoint = oSelection.ITEM(1).Value
oSelection.clear

'Find the coordinates of the reference point
myPoint.GetCoordinates myPointCoords
myPoint.Construction = True

'If reference line is selected, ask the user to select a line
If Me.chkReferenceLine Then
    Set oSelection = CATIA.ActiveDocument.Selection
    oSelection.clear
    InputObject(0) = "Line2D"
    oStatus = oSelection.SelectElement2(InputObject, "Select a reference line", True)
    If oStatus = "Cancel" Then GoTo exit_handler
    Set referenceLine = oSelection.ITEM(1).Value
    oSelection.clear
    referenceLine.GetDirection refLineDirection
    'Change angle calculation based on vector direction (I don't really know why this works but it does)
    Select Case True
        Case refLineDirection(0) >= 0 And refLineDirection(1) < 0
            charAngle = Asin(refLineDirection(1))
        Case refLineDirection(0) < 0 And refLineDirection(1) >= 0
            charAngle = -Acos(Abs(refLineDirection(0)))
        Case Else
            charAngle = Acos(Abs(refLineDirection(0)))
    End Select
    referenceLine.Construction = True
Else
    ' charAngle = Me.txtAngle * (pi / 180)
End If

objShell.Run "\\data\mdbdata\WorkingDB\_docs\Catia_Custom_Files\Macro_Data\Scripts\start.CATScript"

'Flip charAngle if flip is selected
If Me.chkFlip Then charAngle = charAngle + pi

'Look for already created 3D Text and ask user if they want to delete it or not
For i = 1 To mySketch.GeometricElements.count
    If InStr(mySketch.GeometricElements.ITEM(i).name, "3DText") Then
        response = MsgBox("Previously generated 3D text was found." & vbCrLf & "Do you want to remove it?", vbQuestion & vbYesNoCancel, "3D Text Found")
        If response = vbYes Then
            Set oSelection = CATIA.ActiveDocument.Selection
            oSelection.clear
            oSelection.Search ("Name=3DText*,all")
            oSelection.Delete
            oSelection.clear
            GoTo exit_loop
        ElseIf response = vbNo Then
            GoTo exit_loop
        Else
            GoTo exit_handler
        End If
    End If
Next i
exit_loop:

'Calculate total width of the final text
For i = 0 To UBound(inputTextArray)
    If inputTextArray(i) = " " Then
        textWidth = textWidth + (1.5 * charHeight) - charSpacing
        i = i + 1
    End If
    j = 0
    leftPoint = 0
    rightPoint = 0
    Do While j <= UBound(myDictionary(inputTextArray(i)))
        If (myDictionary(inputTextArray(i))(j) * charHeight) < leftPoint Then leftPoint = (myDictionary(inputTextArray(i))(j) * charHeight)
        If (myDictionary(inputTextArray(i))(j) * charHeight) > rightPoint Then rightPoint = (myDictionary(inputTextArray(i))(j) * charHeight)
        j = j + 2
    Loop
    textWidth = textWidth + (rightPoint - leftPoint)
    If i < UBound(inputTextArray) Then textWidth = textWidth + charSpacing
Next i

'Define mirror value and reverse it and textWidth if Mirror is selected
mirror = 1
If Me.chkMirror = True Then
     mirror = -1
     textWidth = -textWidth
End If

'Define anchorX
Select Case Me.btnAnchor.tag
    Case "I_TopLeftAnchor", "I_MiddleLeftAnchor", "I_BottomLeftAnchor"
        anchorX = 0
    Case "I_TopCenterAnchor", "I_MiddleCenterAnchor", "I_BottomCenterAnchor"
        anchorX = 0.5
    Case "I_TopRightAnchor", "I_MiddleRightAnchor", "I_BottomRightAnchor"
        anchorX = 1
End Select

'Define anchorY
Select Case Me.btnAnchor.tag
    Case "I_TopLeftAnchor", "I_TopCenterAnchor", "I_TopRightAnchor"
        anchorY = 1
    Case "I_MiddleLeftAnchor", "I_MiddleCenterAnchor", "I_MiddleRightAnchor"
        anchorY = 0.5
    Case "I_BottomLeftAnchor", "I_BottomCenterAnchor", "I_BottomRightAnchor"
        anchorY = 0
End Select

'Open sketch and loop through every character in array
mySketch.OpenEdition
For i = 0 To UBound(inputTextArray)
    'Handle spaces
    If inputTextArray(i) = " " Then
        charWidth = charWidth + (1.5 * charHeight) - (charSpacing * i)
        i = i + 1
    End If
    
    'Define the first point to be used later
    tempPoint(0) = ((myDictionary(inputTextArray(i))(0) * charHeight) + charWidth + (charSpacing * i)) * mirror - (textWidth * anchorX)
    tempPoint(1) = myDictionary(inputTextArray(i))(1) * charHeight - (charHeight * anchorY)
    firstPoint(0) = myPointCoords(0) + (tempPoint(0) * Cos(charAngle) - tempPoint(1) * Sin(charAngle))
    firstPoint(1) = myPointCoords(1) + (tempPoint(0) * Sin(charAngle) + tempPoint(1) * Cos(charAngle))
    
    j = 0
    Do While j <= UBound(myDictionary(inputTextArray(i))) - 5
        'Define three points at a time and calculate the angle that they create
        tempPoint(0) = ((myDictionary(inputTextArray(i))(j) * charHeight) + charWidth + (charSpacing * i)) * mirror - (textWidth * anchorX)
        tempPoint(1) = myDictionary(inputTextArray(i))(j + 1) * charHeight - (charHeight * anchorY)
        startPoint(0) = myPointCoords(0) + (tempPoint(0) * Cos(charAngle) - tempPoint(1) * Sin(charAngle))
        startPoint(1) = myPointCoords(1) + (tempPoint(0) * Sin(charAngle) + tempPoint(1) * Cos(charAngle))
        
        tempPoint(0) = ((myDictionary(inputTextArray(i))(j + 2) * charHeight) + charWidth + (charSpacing * i)) * mirror - (textWidth * anchorX)
        tempPoint(1) = myDictionary(inputTextArray(i))(j + 3) * charHeight - (charHeight * anchorY)
        midPoint(0) = myPointCoords(0) + (tempPoint(0) * Cos(charAngle) - tempPoint(1) * Sin(charAngle))
        midPoint(1) = myPointCoords(1) + (tempPoint(0) * Sin(charAngle) + tempPoint(1) * Cos(charAngle))
        
        tempPoint(0) = ((myDictionary(inputTextArray(i))(j + 4) * charHeight) + charWidth + (charSpacing * i)) * mirror - (textWidth * anchorX)
        tempPoint(1) = myDictionary(inputTextArray(i))(j + 5) * charHeight - (charHeight * anchorY)
        endPoint(0) = myPointCoords(0) + (tempPoint(0) * Cos(charAngle) - tempPoint(1) * Sin(charAngle))
        endPoint(1) = myPointCoords(1) + (tempPoint(0) * Sin(charAngle) + tempPoint(1) * Cos(charAngle))
        
        vector1(0) = midPoint(0) - startPoint(0)
        vector1(1) = midPoint(1) - startPoint(1)
        vector2(0) = midPoint(0) - endPoint(0)
        vector2(1) = midPoint(1) - endPoint(1)
        magnitude1 = Sqr((vector1(0) ^ 2) + (vector1(1) ^ 2))
        magnitude2 = Sqr((vector2(0) ^ 2) + (vector2(1) ^ 2))
        X = ((vector1(0) * vector2(0)) + (vector1(1) * vector2(1))) / (Abs(magnitude1) * Abs(magnitude2))
        myAngle = Acos(X) * (180 / pi)
        
        'If the angle is close to 180 and the lines are less then 0.2mm, use these points to create a spline
        If Abs(myAngle) > 170 And Abs(myAngle) < 190 And magnitude1 < (0.05 * charHeight) And magnitude2 < (0.05 * charHeight) Then
            If (Not Not splinePoints) = 0 Then
                ReDim splinePoints(2)
                Set splinePoints(UBound(splinePoints) - 2) = mySketch.Factory2D.CreatePoint(startPoint(0), startPoint(1))
                splinePoints(UBound(splinePoints) - 2).name = "3DText." & inputText & "." & splinePoints(UBound(splinePoints) - 2).name
                Set splinePoints(UBound(splinePoints) - 1) = mySketch.Factory2D.CreatePoint(midPoint(0), midPoint(1))
                splinePoints(UBound(splinePoints) - 1).name = "3DText." & inputText & "." & splinePoints(UBound(splinePoints) - 1).name
                Set splinePoints(UBound(splinePoints)) = mySketch.Factory2D.CreatePoint(endPoint(0), endPoint(1))
                splinePoints(UBound(splinePoints)).name = "3DText." & inputText & "." & splinePoints(UBound(splinePoints)).name
            Else
                ReDim Preserve splinePoints(UBound(splinePoints) + 1)
                Set splinePoints(UBound(splinePoints)) = mySketch.Factory2D.CreatePoint(endPoint(0), endPoint(1))
                splinePoints(UBound(splinePoints)).name = "3DText." & inputText & "." & splinePoints(UBound(splinePoints)).name
            End If
            
        'If not, create a spline from the previously collected points, or create a straight line
        Else
            If (Not Not splinePoints) = 0 Then
                Set myLine = mySketch.Factory2D.CreateLine(startPoint(0), startPoint(1), midPoint(0), midPoint(1))
                myLine.name = "3DText." & inputText & "." & myLine.name
            Else
                Set mySpline = mySketch.Factory2D.CreateSpline(splinePoints)
                mySpline.name = "3DText." & inputText & "." & mySpline.name
                Erase splinePoints()
            End If
        End If
        
        'If the first point and the current end point are the same, close the shape and jump to the next loop
        If firstPoint(0) = endPoint(0) And firstPoint(1) = endPoint(1) Then
            If (Not Not splinePoints) = 0 Then
                Set myLine = mySketch.Factory2D.CreateLine(midPoint(0), midPoint(1), endPoint(0), endPoint(1))
                myLine.name = "3DText." & inputText & "." & myLine.name
            Else
                Set mySpline = mySketch.Factory2D.CreateSpline(splinePoints)
                mySpline.name = "3DText." & inputText & "." & mySpline.name
                Erase splinePoints()
            End If
            If j + 7 <= UBound(myDictionary(inputTextArray(i))) - 5 Then
                'Redefine the first point of the next closed shape
                tempPoint(0) = ((myDictionary(inputTextArray(i))(j + 6) * charHeight) + charWidth + (charSpacing * i)) * mirror - (textWidth * anchorX)
                tempPoint(1) = myDictionary(inputTextArray(i))(j + 7) * charHeight - (charHeight * anchorY)
                firstPoint(0) = myPointCoords(0) + (tempPoint(0) * Cos(charAngle) - tempPoint(1) * Sin(charAngle))
                firstPoint(1) = myPointCoords(1) + (tempPoint(0) * Sin(charAngle) + tempPoint(1) * Cos(charAngle))
            End If
            j = j + 4
        End If
        j = j + 2
    Loop
    
    'Calculate where to place the next character based on the width of the character that was just placed
    j = 0
    leftPoint = 0
    rightPoint = 0
    Do While j <= UBound(myDictionary(inputTextArray(i)))
        If (myDictionary(inputTextArray(i))(j) * charHeight) < leftPoint Then leftPoint = (myDictionary(inputTextArray(i))(j) * charHeight)
        If (myDictionary(inputTextArray(i))(j) * charHeight) > rightPoint Then rightPoint = (myDictionary(inputTextArray(i))(j) * charHeight)
        j = j + 2
    Loop
    charWidth = charWidth + (rightPoint - leftPoint)
Next i

'Cleanup
mySketch.CloseEdition
Set myDictionary = Nothing
'myPart.Update
rs.CLOSE
Set rs = Nothing

exit_handler:
objShell.Run "\\data\mdbdata\WorkingDB\_docs\Catia_Custom_Files\Macro_Data\Scripts\finish.CATScript"
Set objShell = Nothing
formStatus (False)
Set CATIA = Nothing

Exit Sub
Err_Handler:
    MsgBox Err.DESCRIPTION, vbOKOnly, "Error Code: " & Err.number
End Sub

Sub generateLetterMap()

Dim oSelection
Dim charsToMap, charsToMapArr()
Dim myLine
Dim startCoords(1), endCoords(1)
Dim firstPoint()
Dim myLetter(), myShape()
Dim letterLeft, letterRight
Dim shapeLeft, shapeRight
Dim leftMostPoint
Dim rs As Recordset
Dim i As Long, j As Long

Set CATIA = GetObject(, "CATIA.Application")
Set oSelection = CATIA.ActiveDocument.Selection

charsToMap = "AaBbCcDdEeFfGgHhIiJjKkLlMmNnOoPpQqRrSsTtUuVvWwXxYyZz`1234567890-=~!@#$%^&*()_+[]\{}|;':" & Chr(34) & ",./<>?"
ReDim charsToMapArr(Len(charsToMap) - 1)
For i = 0 To UBound(charsToMapArr)
    charsToMapArr(i) = Mid(charsToMap, i + 1, 1)
Next i

For i = 1 To oSelection.count
    Do While j < UBound(charsToMapArr)
        If typeName(oSelection.ITEM(i).Value) = "Line2D" Then
            Set myLine = oSelection.ITEM(i).Value
            myLine.startPoint.GetCoordinates startCoords
            myLine.endPoint.GetCoordinates endCoords
            
            'If first point is empty
            If (Not Not firstPoint) = 0 Then
                ReDim firstPoint(1)
                firstPoint(0) = startCoords(0)
                firstPoint(1) = startCoords(1)
                
                ' Add coordinates of the line's start point
                If (Not Not myLetter) = 0 Then
                    ReDim myLetter(1)
                Else
                    ReDim Preserve myLetter(UBound(myLetter) + 2)
                End If
                myLetter(UBound(myLetter) - 1) = startCoords(0)
                myLetter(UBound(myLetter)) = startCoords(1)
            End If
            
            ' Add the coordinates of the line's end point
            ReDim Preserve myLetter(UBound(myLetter) + 2)
            myLetter(UBound(myLetter) - 1) = endCoords(0)
            myLetter(UBound(myLetter)) = endCoords(1)
            
            If Join(firstPoint, "") = Join(endCoords, "") Then
                Erase firstPoint()
            End If
        End If
        j = j + 1
    Loop
Next i

i = 0
Do While i < UBound(myLetter)
    If IsEmpty(leftMostPoint) Then
        leftMostPoint = myLetter(i)
    Else
        If myLetter(i) < leftMostPoint Then leftMostPoint = myLetter(i)
    End If
    i = i + 2
Loop

i = 0
Do While i < UBound(myLetter)
    myLetter(i) = myLetter(i) - leftMostPoint
    i = i + 2
Loop

Set rs = CurrentDb.OpenRecordset("tblCatiaFont")
With rs
    .addNew
    !FontName = "3"
    !characterName = InputBox("What character is this?", "Character")
    !characterArray = Join(myLetter, ",")
    .Update
End With

End Sub

Private Sub btnBlack_Click()
setBackgroundColor (Me.ActiveControl.Caption)
End Sub

Private Sub btnGray_Click()
setBackgroundColor (Me.ActiveControl.Caption)
End Sub

Private Sub btnBlue_Click()
setBackgroundColor (Me.ActiveControl.Caption)
End Sub

Private Sub btnWhite_Click()
setBackgroundColor (Me.ActiveControl.Caption)
End Sub

Private Sub btnCustom_Click()

If isHex("#" & Nz(Me.txtCustom.Value)) = False Then
    MsgBox "You must enter a valid HEX code.", vbExclamation, "Invalid HEX Code"
    Exit Sub
End If

setBackgroundColor (Nz(Me.txtCustom.Value))
End Sub

Private Sub setBackgroundColor(setColor As String)
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name, setColor)

If catiaOpen = False Then
    MsgBox "CATIA must be open to use this function.", vbExclamation, "CATIA Not Open"
    Exit Sub
End If

formStatus (True)

Dim visualizationSettingAtt1
Dim ioR As Long, ioG As Long, ioB As Long
Set visualizationSettingAtt1 = CATIA.SettingControllers.ITEM("CATVizVisualizationSettingCtrl")

Select Case setColor
    Case "Black"
        visualizationSettingAtt1.SetBackgroundRGB 0, 0, 0
        Me.dispBlack.Caption = "P"
        Me.dispGray.Caption = ""
        Me.dispBlue.Caption = ""
        Me.dispWhite.Caption = ""
        Me.dispCustom.Caption = ""
    Case "Gray"
        visualizationSettingAtt1.SetBackgroundRGB 64, 64, 64
        Me.dispBlack.Caption = ""
        Me.dispGray.Caption = "P"
        Me.dispBlue.Caption = ""
        Me.dispWhite.Caption = ""
        Me.dispCustom.Caption = ""
    Case "Blue"
        visualizationSettingAtt1.SetBackgroundRGB 51, 51, 102
        Me.dispBlack.Caption = ""
        Me.dispGray.Caption = ""
        Me.dispBlue.Caption = "P"
        Me.dispWhite.Caption = ""
        Me.dispCustom.Caption = ""
    Case "White"
        visualizationSettingAtt1.SetBackgroundRGB 255, 255, 255
        Me.dispBlack.Caption = ""
        Me.dispGray.Caption = ""
        Me.dispBlue.Caption = ""
        Me.dispWhite.Caption = "P"
        Me.dispCustom.Caption = ""
    Case Else
        ioR = val("&H" & Mid(setColor, 1, 2))
        ioG = val("&H" & Mid(setColor, 3, 2))
        ioB = val("&H" & Mid(setColor, 5, 2))
        visualizationSettingAtt1.SetBackgroundRGB ioR, ioG, ioB
        Me.dispBlack.Caption = ""
        Me.dispGray.Caption = ""
        Me.dispBlue.Caption = ""
        Me.dispWhite.Caption = ""
        Me.dispCustom.Caption = "P"
End Select

visualizationSettingAtt1.SaveRepository

exit_handler:
formStatus (False)
Set CATIA = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub exportSheetDocHis_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)

If catiaOpen = False Then
    MsgBox "CATIA must be open to use this function.", vbExclamation, "CATIA Not Open"
    Exit Sub
End If

formStatus (True)

'First, check if DRS dash is open
If CurrentProject.AllForms("frmDRSdashboard").IsLoaded = False Then
    MsgBox "Need to have a WO open to use this", vbCritical, "No can do"
    GoTo exit_handler
End If

If Form_frmDRSdashboard.Check_In_Prog <> "Approved" Then
    MsgBox "WO must be approved to publish drawing", vbCritical, "No can do"
    GoTo exit_handler
End If

Call CATIA_exportCurrentSheet(openDocumentHistoryFolder(Form_frmDRSdashboard.Part_Number, False), , Nz(Me.publishDwgName))

exit_handler:
formStatus (False)
Set CATIA = Nothing
Exit Sub

Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
    GoTo exit_handler
End Sub

Private Sub exportToChkFold_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)

If catiaOpen = False Then
    MsgBox "CATIA must be open to use this function.", vbExclamation, "CATIA Not Open"
    Exit Sub
End If

formStatus (True)

'First, check if DRS dash is open
If CurrentProject.AllForms("frmDRSdashboard").IsLoaded = False Then
    MsgBox "Need to have a WO open to use this", vbCritical, "No can do"
    GoTo exit_handler
End If

Dim chkFold As String
chkFold = Nz(Form_frmDRSdashboard.Check_Folder, "")

If chkFold = "" Then
    MsgBox "Need to have a check folder saved to do this", vbCritical, "No can do"
    GoTo exit_handler
End If

Call CATIA_exportCurrentSheet(chkFold, Nz(Me.WOfilePrefix, ""))

exit_handler:
formStatus (False)
Set CATIA = Nothing
Exit Sub

Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
    GoTo exit_handler
End Sub

Private Sub forceIsolate_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)

If catiaOpen = False Then
    MsgBox "CATIA must be open to use this function.", vbExclamation, "CATIA Not Open"
    Exit Sub
End If

formStatus (True)

Dim USel
Dim InputObject(0)
Dim oStatus
Dim myView

Set USel = CATIA.ActiveDocument.Selection
InputObject(0) = "DrawingView"

oStatus = USel.SelectElement2(InputObject, "Select a view to isolate", True)

Do While oStatus <> "Cancel"
    Set myView = USel.ITEM(1).Value
    myView.GenerativeLinks.RemoveAllLinks
    USel.clear
    
    If myView.name = CATIA.ActiveDocument.Sheets.ActiveSheet.Views.ActiveView.name Then CATIA.ActiveDocument.Sheets.ActiveSheet.Views.ITEM(1).Activate
    myView.Activate

    oStatus = USel.SelectElement2(InputObject, "Select a view to isolate", True)
Loop

CATIA.ActiveDocument.Sheets.ActiveSheet.Activate

exit_handler:
formStatus (False)
Set CATIA = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

formStatus (True)

Call setTheme(Me)

Dim ioR1 As Long, ioG1 As Long, ioB1 As Long

DoCmd.applyFilter , "User = '" & Environ("username") & "'"
If Me.catiaCustomColor <> "" Then
    ioR1 = val("&H" & Mid(Me.txtCustom.Value, 1, 2))
    ioG1 = val("&H" & Mid(Me.txtCustom.Value, 3, 2))
    ioB1 = val("&H" & Mid(Me.txtCustom.Value, 5, 2))
    Me.dispCustom.BackColor = rgb(ioR1, ioG1, ioB1)
End If

If catiaOpen = False Then
    Me.txtActiveCATIA = "None"
    GoTo exit_handler
End If
Me.txtActiveCATIA = CATIA.Application.Caption

Dim visualizationSettingAtt1
Dim ioR2 As Long, ioG2 As Long, ioB2 As Long
Set visualizationSettingAtt1 = CATIA.SettingControllers.ITEM("CATVizVisualizationSettingCtrl")

visualizationSettingAtt1.GetBackgroundRGB ioR2, ioG2, ioB2
Select Case True
    Case ioR2 = 0 And ioG2 = 0 And ioB2 = 0
        Me.dispBlack.Caption = "P"
        Me.dispGray.Caption = ""
        Me.dispBlue.Caption = ""
        Me.dispWhite.Caption = ""
        Me.dispCustom.Caption = ""
    Case ioR2 = 64 And ioG2 = 64 And ioB2 = 64
        Me.dispBlack.Caption = ""
        Me.dispGray.Caption = "P"
        Me.dispBlue.Caption = ""
        Me.dispWhite.Caption = ""
        Me.dispCustom.Caption = ""
    Case ioR2 = 51 And ioG2 = 51 And ioB2 = 102
        Me.dispBlack.Caption = ""
        Me.dispGray.Caption = ""
        Me.dispBlue.Caption = "P"
        Me.dispWhite.Caption = ""
        Me.dispCustom.Caption = ""
    Case ioR2 = 255 And ioG2 = 255 And ioB2 = 255
        Me.dispBlack.Caption = ""
        Me.dispGray.Caption = ""
        Me.dispBlue.Caption = ""
        Me.dispWhite.Caption = "P"
        Me.dispCustom.Caption = ""
    Case ioR1 = ioR2 And ioG1 = ioG2 And ioB1 = ioB2
        Me.dispBlack.Caption = ""
        Me.dispGray.Caption = ""
        Me.dispBlue.Caption = ""
        Me.dispWhite.Caption = ""
        Me.dispCustom.Caption = "P"
    Case Else
        Me.dispBlack.Caption = ""
        Me.dispGray.Caption = ""
        Me.dispBlue.Caption = ""
        Me.dispWhite.Caption = ""
        Me.dispCustom.Caption = ""
End Select

If visualizationSettingAtt1.ColorBackgroundMode = True Then
    Me.btnGradient.Value = True
    Me.dispBlack.Gradient = 15
    Me.dispGray.Gradient = 15
    Me.dispBlue.Gradient = 15
    Me.dispWhite.Gradient = 15
    Me.dispCustom.Gradient = 15
Else
    Me.btnGradient.Value = False
    Me.dispBlack.Gradient = 0
    Me.dispGray.Gradient = 0
    Me.dispBlue.Gradient = 0
    Me.dispWhite.Gradient = 0
    Me.dispCustom.Gradient = 0
End If

exit_handler:
formStatus (False)
Set CATIA = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub catiaMacrosHelp_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)

Call openPath(mainFolder(Me.ActiveControl.name))

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub btnGradient_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)

If catiaOpen = False Then
    MsgBox "CATIA must be open to use this function.", vbExclamation, "CATIA Not Open"
    Me.btnGradient.Value = Null
    Exit Sub
End If

formStatus (True)

Dim visualizationSettingAtt1
Set visualizationSettingAtt1 = CATIA.SettingControllers.ITEM("CATVizVisualizationSettingCtrl")

If Me.btnGradient.Value = False Then
    Me.btnGradient.Value = False
    visualizationSettingAtt1.ColorBackgroundMode = False
    Me.dispBlack.Gradient = 0
    Me.dispGray.Gradient = 0
    Me.dispBlue.Gradient = 0
    Me.dispWhite.Gradient = 0
    Me.dispCustom.Gradient = 0
Else
    Me.btnGradient.Value = True
    visualizationSettingAtt1.ColorBackgroundMode = True
    Me.dispBlack.Gradient = 15
    Me.dispGray.Gradient = 15
    Me.dispBlue.Gradient = 15
    Me.dispWhite.Gradient = 15
    Me.dispCustom.Gradient = 15
End If

visualizationSettingAtt1.SaveRepository

exit_handler:
formStatus (False)
Set CATIA = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub copyPaste_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)

If catiaOpen = False Then
    MsgBox "CATIA must be open to use this function.", vbExclamation, "CATIA Not Open"
    Exit Sub
End If

formStatus (True)

Dim USelCopy, USelPaste
Dim InputObject(0)
Dim oStatusCopy, oStatusPaste
Dim X, oStatus

InputObject(0) = "AnyObject"
Set USelCopy = CATIA.ActiveDocument.Selection
Set USelPaste = CATIA.ActiveDocument.Selection

On Error Resume Next
Do While oStatus <> "Cancel"
    oStatusCopy = USelCopy.SelectElement2(InputObject, "Select COPY element", True)

    X = USelCopy.ITEM(1).Value.Text
    USelCopy.clear
    
    oStatusPaste = USelPaste.SelectElement2(InputObject, "Select PASTE element", True)
    If (oStatusPaste = "Cancel") Then GoTo exit_handler

    USelPaste.ITEM(1).Value.Text = X
    USelPaste.clear
Loop

exit_handler:
formStatus (False)
Set CATIA = Nothing
    
Exit Sub
Err_Handler:
If Err.number = -2147467259 Then GoTo exit_handler
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub countRevs_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)

If catiaOpen = False Then
    MsgBox "CATIA must be open to use this function.", vbExclamation, "CATIA Not Open"
    Exit Sub
End If

formStatus (True)

Dim oSelection
Dim myComponent, myComponent2
Dim showState As Long
Dim revLevel, revCount
Dim i As Long, j As Long

revLevel = Nz(Me.revLevel, "")
If revLevel = "" Then
    MsgBox "No number entered", vbExclamation, "Cancel Pressed"
    GoTo exit_handler
End If

Set oSelection = CATIA.ActiveDocument.Selection
CATIA.ActiveWindow.ActiveViewer.Reframe
revCount = 0

oSelection.clear
oSelection.Search "Type='2D Component Instance',scr"
For i = 1 To oSelection.count
    Set myComponent = oSelection.ITEM(i).Value
    If InStr(myComponent.name, "REV" & revLevel) > 0 Or Left(myComponent.name, 3) = "[" & revLevel & "]" Then
        revCount = revCount + 1
    End If
    For j = 1 To myComponent.CompRef.Components.count
        Set myComponent2 = myComponent.CompRef.Components.ITEM(j)
        If InStr(myComponent2.name, "REV" & revLevel) > 0 Or Left(myComponent2.name, 3) = "[" & revLevel & "]" Then
            revCount = revCount + 1
        End If
    Next j
Next i
oSelection.clear

MsgBox "#" & revLevel & " rev triangles: " & revCount, vbInformation, "Here ya go"

exit_handler:
formStatus (False)
Set CATIA = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub createPQS_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)

If catiaOpen = False Then
    MsgBox "CATIA must be open to use this function.", vbExclamation, "CATIA Not Open"
    Exit Sub
End If

formStatus (True)

Dim oSelection
Dim dimension
Dim gdt
Dim instantCrit
Dim i As Integer
Dim oBefore As String
Dim oPrefix As String
Dim upperText
Dim critCode As String
Dim critAmount As Long
Dim response1 As Long
Dim response2 As Long
Dim textBox, viewArray

CATIA.ActiveWindow.ActiveViewer.Reframe

Set viewArray = CreateObject("System.Collections.ArrayList")
Set oSelection = CATIA.ActiveDocument.Selection

'Select all dimensions and check for critical symbols
oSelection.Search "Name=Dimension*,scr"
For i = 1 To oSelection.count
    Set dimension = oSelection.ITEM(i).Value
    Call dimension.getValue.GetBaultText(1, oBefore, "", "", "")
    Call dimension.getValue.GetPSText(1, oPrefix, "")
    critCode = StrConv(oBefore, vbUnicode)
    
    'Add views with critical dimensions to viewArray
    If (InStr(critCode, "¼%") > 0 And InStr(critCode, "(") = 0) Or oPrefix = "<Black Triangle Down>" Then
        If viewArray.Contains(dimension.Parent.Parent) = False Then viewArray.Add dimension.Parent.Parent
        If InStr(dimension.Parent.Parent.name, "SEC") > 0 Or InStr(dimension.Parent.Parent.name, "DETAIL") > 0 Then
            If viewArray.Contains(dimension.Parent.Parent.ReferenceView) = False Then viewArray.Add dimension.Parent.Parent.ReferenceView
            If viewArray.Contains(dimension.Parent.Parent.GenerativeBehavior.ParentView) = False Then viewArray.Add dimension.Parent.Parent.GenerativeBehavior.ParentView
        End If
        critAmount = critAmount + 1
    End If
Next
oSelection.clear

'Select all GD&T and check for critical symbols
oSelection.Search "Name='Geometrical Tolerance*',scr"
For i = 1 To oSelection.count
    Set gdt = oSelection.ITEM(i).Value
    Set upperText = gdt.GetTextRange(0, 0)
    critCode = StrConv(oBefore, vbUnicode)
    
    'Add views with critical dimensions to viewArray
    If InStr(critCode, "¼%") > 0 Then
        If viewArray.Contains(gdt.Parent.Parent) = False Then viewArray.Add gdt.Parent.Parent
        If InStr(gdt.Parent.Parent.name, "SEC") > 0 Or InStr(gdt.Parent.Parent.name, "DETAIL") > 0 Then
            If viewArray.Contains(gdt.Parent.Parent.ReferenceView) = False Then viewArray.Add gdt.Parent.Parent.ReferenceView
            If viewArray.Contains(gdt.Parent.Parent.GenerativeBehavior.ParentView) = False Then viewArray.Add gdt.Parent.Parent.GenerativeBehavior.ParentView
        End If
        critAmount = critAmount + 1
    End If
Next
oSelection.clear

'Select all instantiated critical marks
oSelection.Search "Name='Critical Mark*',scr"
For i = 1 To oSelection.count
    Set instantCrit = oSelection.ITEM(i).Value
    
    'Add views with critical dimensions to viewArray
    If viewArray.Contains(instantCrit.Parent.Parent) = False Then viewArray.Add instantCrit.Parent.Parent
    If InStr(instantCrit.Parent.Parent.name, "SEC") > 0 Or InStr(instantCrit.Parent.Parent.name, "DETAIL") > 0 Then
        If viewArray.Contains(instantCrit.Parent.Parent.ReferenceView) = False Then viewArray.Add instantCrit.Parent.Parent.GenerativeBehavior.ParentView
        If viewArray.Contains(instantCrit.Parent.Parent.GenerativeBehavior.ParentView) = False Then viewArray.Add instantCrit.Parent.Parent.GenerativeBehavior.ParentView
    End If
    critAmount = critAmount + 1
Next
oSelection.clear

'Message box to display critical dimension count
Select Case critAmount
    Case 0
        MsgBox "There are no critical dimensions on this drawing." & vbCrLf & "Please check if you are on the correct sheet.", , "No Critical Dimensions"
        GoTo exit_handler
    Case 1
        response1 = MsgBox(critAmount & " critical dimension found." & vbCrLf & "Do you want to create a PQS?", vbYesNo, "Critical Dimension Count")
    Case Else
        response1 = MsgBox(critAmount & " critical dimensions found." & vbCrLf & "Do you want to create a PQS?", vbYesNo, "Critical Dimension Count")
End Select
response2 = MsgBox("Do you want to update the critical quantity mark?", vbYesNo, "Critical Dimension Count")

'Update critical quantity mark
If response2 = vbYes Then
    CATIA.ActiveDocument.Sheets.ITEM("Detail").Views.ITEM("Critical Quantity").Texts.ITEM(1).Text = "x " & critAmount
End If
If response1 = vbNo Then GoTo exit_handler

'Create PQS drawing sheet
Dim View
Dim oDrawingSheets
Dim oDrawingSheet
Set oDrawingSheets = CATIA.ActiveDocument.Sheets
Set oDrawingSheet = oDrawingSheets.Add("PQS_1000")

'Copy views containing critical dimensions
For Each View In viewArray
    If IsObject(View) Then oSelection.Add View
Next
oSelection.Copy
oSelection.clear

'Paste copied views into PQS sheet
oSelection.Add oDrawingSheet
oSelection.Paste
oDrawingSheet.Activate
oSelection.clear

'Select all dimensions and hide non critical
oSelection.Search "Name=Dimension*,scr"
i = 1
Do While i <= oSelection.count
    Set dimension = oSelection.ITEM(i).Value
    Call dimension.getValue.GetBaultText(1, oBefore, "", "", "")
    Call dimension.getValue.GetPSText(1, oPrefix, "")
    critCode = StrConv(oBefore, vbUnicode)
    
    'Deselect critical dimensions
    If (InStr(critCode, "¼%") > 0 And InStr(critCode, "(") = 0) Or oPrefix = "<Black Triangle Down>" Then
        oSelection.remove i
    Else
        i = i + 1
    End If
Loop
oSelection.VisProperties.SetShow 1
oSelection.clear

'Select all GD&T and hide non critical
oSelection.Search "Name='Geometrical Tolerance*',scr"
i = 1
Do While i <= oSelection.count
    Set gdt = oSelection.ITEM(i).Value
    Set upperText = gdt.GetTextRange(0, 0)
    critCode = StrConv(oBefore, vbUnicode)
    
    'Deselect critical GD&T
    If InStr(critCode, "¼%") > 0 Then
        oSelection.remove i
    Else
        i = i + 1
    End If
Loop
oSelection.VisProperties.SetShow 1
oSelection.clear

'Select all text boxes and hide unimportant text
oSelection.Search "Name=Text*,scr"
i = 1
Do While i <= oSelection.count
    Set textBox = oSelection.ITEM(i).Value
    Select Case True
        Case InStr(textBox.Text, "SEC") > 0 Or InStr(textBox.Text, "DETAIL") > 0 Or InStr(textBox.Text, "SCALE") > 0 Or textBox.TextProperties.Bold = 1
            oSelection.remove i
        Case InStr(textBox.Text, "PLACE") > 0 And textBox.Leaders.count = 0
            If IsError(textBox.AssociativeElement.Leaders.count) Then
                oSelection.remove i
            Else
                If textBox.AssociativeElement.Leaders.count = 0 Then
                    oSelection.remove i
                Else
                    i = i + 1
                End If
            End If
        Case Else
            i = i + 1
    End Select
Loop
oSelection.VisProperties.SetShow 1
oSelection.clear

'Hide all other elements unimportant to PQS
'Others to potentially hide:
'Name='Line*'
'Name='Area Fill*'
oSelection.Search "Name='Datum Feature*'+Name='Circle Dot*'+Name='Point*'+Name='REV*'+Name='DOT*',scr"
oSelection.VisProperties.SetShow 1
oSelection.clear

'Final message box
MsgBox "Please double check PQS for possible errors.", , "PQS Complete"

exit_handler:
formStatus (False)
Set CATIA = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub deleteFirstLine_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)

If catiaOpen = False Then
    MsgBox "CATIA must be open to use this function.", vbExclamation, "CATIA Not Open"
    Exit Sub
End If

formStatus (True)

Dim USel
Dim InputObject(0)
Dim oStatus
Dim X As String

InputObject(0) = "AnyObject"
Set USel = CATIA.ActiveDocument.Selection

Do While oStatus <> "Cancel"
    oStatus = USel.SelectElement2(InputObject, "Select element", True)

    X = USel.ITEM(1).Value.Text
    
    Dim arrLines() As String

    X = Replace(X, vbCrLf, vbCr)
    X = Replace(X, vbLf, vbCr)
    arrLines = Split(X, vbCr)
    X = arrLines(1)
    
    If UBound(arrLines) > 1 Then
        Dim j As Long
        For j = 2 To UBound(arrLines)
           X = X & vbNewLine & arrLines(j)
        Next
    End If

    USel.ITEM(1).Value.Text = StrConv(X, vbUpperCase)
    USel.clear
Loop

exit_handler:
formStatus (False)
Set CATIA = Nothing
    
Exit Sub
Err_Handler:
If Err.number = -2147467259 Then GoTo exit_handler
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub exportCurrentSheet_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)

If catiaOpen = False Then
    MsgBox "CATIA must be open to use this function.", vbExclamation, "CATIA Not Open"
    Exit Sub
End If

formStatus (True)

Call CATIA_exportCurrentSheet

exit_handler:
formStatus (False)
Set CATIA = Nothing
Exit Sub

Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
    GoTo exit_handler
End Sub

Function CATIA_exportCurrentSheet(Optional location As String = "", Optional prefix As String = "", Optional customFileName As String = "")
On Error GoTo Err_Handler

If catiaOpen = False Then
    MsgBox "CATIA must be open to use this function.", vbExclamation, "CATIA Not Open"
    Exit Function
End If

formStatus (True)

Dim fileNameFull As String, filePathFull As String, sheetName As String, newFileName As String, folderLoc As String
fileNameFull = CATIA.ActiveDocument.name

If Right(fileNameFull, 11) <> ".CATDrawing" Then
    MsgBox "Must have a CATDrawing active to export it", vbCritical, "Hmm...?"
    GoTo exit_handler
End If
fileNameFull = Left(fileNameFull, Len(fileNameFull) - 11)

sheetName = CATIA.ActiveDocument.Sheets.ActiveSheet.name

If InStr(sheetName, ".") Then
    MsgBox "Sheet name cannot have a period in it, please modify then try again", vbInformation, "Try again"
    GoTo exit_handler
End If

folderLoc = "H:\CATIA_temp\"
If Dir(folderLoc, vbDirectory) = "" Then MkDir (folderLoc)

Dim objDoc As Object, objParam As Object, objItems As Object 'find the parameter for part number
'Set objDoc = CATIA.Windows.Item(1).Parent
Set objDoc = CATIA.ActiveDocument
Set objParam = objDoc.Parameters.RootParameterSet
On Error GoTo paramError
Set objItems = objParam.DirectParameters.ITEM("Drawing\NF_Part_No")
On Error GoTo Err_Handler

Dim partNum As String, dashLoc As Integer 'grab part number
If Left(objItems.Value, 1) = "(" Then
    partNum = Right(objItems.Value, 8)
Else
    partNum = Left(objItems.Value, 8)
End If
partNum = Replace(partNum, "-", "_")

newFileName = partNum & "_" & sheetName & ".pdf"

If location = "" And Left(fileNameFull, 3) = "30M" Then
    location = "H:\Documents\"
ElseIf location = "" Then
    location = CATIA.ActiveDocument.Path
End If

location = addLastSlash(location)

filePathFull = location & prefix & newFileName
If customFileName <> "" Then filePathFull = location & customFileName

CATIA.ActiveDocument.ExportData folderLoc & partNum, "pdf"

Dim fso, oFile As Object
Set fso = CreateObject("Scripting.FileSystemObject")
Call fso.CopyFile(folderLoc & newFileName, filePathFull)

For Each oFile In fso.GetFolder(folderLoc).Files 'delete all temp files
    oFile.Delete
Next

If MsgBox("File: " & newFileName & vbNewLine & "Exported to: " & filePathFull & vbNewLine & vbNewLine & "Do you want to open " & location & "?", vbYesNo, "Export Successful") = vbYes Then openPath (location)

exit_handler:
formStatus (False)
Set CATIA = Nothing
Exit Function

Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
    GoTo exit_handler
    
paramError:
    MsgBox "You need to have 3Dex Properties set up to use this function", vbInformation, "Error!"
End Function

Private Sub exportPicture_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)

If catiaOpen = False Then
    MsgBox "CATIA must be open to use this function.", vbExclamation, "CATIA Not Open"
    Exit Sub
End If

formStatus (True)

Dim objViewer3D, CatCaptureFormatJPEG, tempPic As String

CATIA.ActiveDocument.Selection.clear 'clear selection
CATIA.StartCommand ("Specifications") 'hide spec tree
CATIA.StartCommand ("Compass") 'hide compass

Set objViewer3D = CATIA.ActiveWindow.ActiveViewer

Dim dblWhiteArray(2)
dblWhiteArray(0) = 1
dblWhiteArray(1) = 1
dblWhiteArray(2) = 1
objViewer3D.PutBackgroundColor (dblWhiteArray) 'set background to white

objViewer3D.FullScreen = True 'best resolution

If Me.partPictureFitAll Then objViewer3D.Reframe 'Fit all in screen

tempPic = Environ("temp") & "\tempPic" & nowString & ".JPEG"

Call objViewer3D.CaptureToFile(5, tempPic) 'save to temp file

CATIA.StartCommand ("Specifications") 'show spec tree
CATIA.StartCommand ("Compass") 'show compass
objViewer3D.FullScreen = False
         
'Grab part number from parameters
Dim prodName, objParam
Dim partNum As String 'grab part number

On Error GoTo paramError
prodName = CATIA.ActiveDocument.Product.name
Set objParam = CATIA.ActiveDocument.Product.Parameters.ITEM(prodName & "\Properties\NF_Part_No")
On Error GoTo Err_Handler

If Left(objParam.Value, 1) = "(" Then
    partNum = Right(objParam.Value, 5)
Else
    partNum = Left(objParam.Value, 5)
End If

GoTo paramOK

paramError:
Dim X
X = InputBox("Enter Part Number", "Enter Part Number (not found in properties)")
If X = "" Or X = vbCancel Then Exit Sub
partNum = CStr(X)

paramOK:
Dim ppt As New PowerPoint.Application
Dim pptPres As PowerPoint.Presentation
Dim curSlide As PowerPoint.Slide
Dim pptLayout As CustomLayout
Dim shp As PowerPoint.Shape

ppt.Presentations.Add
Set pptPres = ppt.ActivePresentation
Set pptLayout = pptPres.Designs(1).SlideMaster.CustomLayouts(7)
Set curSlide = pptPres.Slides.AddSlide(1, pptLayout)

Set shp = curSlide.Shapes.AddPicture(tempPic, msoFalse, msoTrue, 0, 0)

'CROP THE IMAGE
Dim objImage As Object, objPict As Object
Dim endingPoint, midPoint, margin, marginScan, preCroppedPic, reverseScan
Set objImage = CreateObject("WIA.ImageFile")
objImage.LoadFile tempPic
Set objPict = LoadPicture(tempPic)
endingPoint = objImage.Width
midPoint = (0.5 * objImage.Height)

For marginScan = 1 To endingPoint
    On Error Resume Next
    If Not (PixelTest(objPict, marginScan, midPoint) Like "1677*") Then
        margin = marginScan
        shp.PictureFormat.CropLeft = margin * 0.7
        Exit For
    End If
Next
For marginScan = 1 To endingPoint
    reverseScan = endingPoint - marginScan
    If Not (PixelTest(objPict, reverseScan, midPoint) Like "1677*") Then
        margin = marginScan
        shp.PictureFormat.CropRight = margin * 0.7
        Exit For
    End If
Next
endingPoint = objImage.Height
midPoint = (0.5 * objImage.Width)
For marginScan = 1 To endingPoint
    If Not (PixelTest(objPict, midPoint, marginScan) Like "1677*") Then
        margin = marginScan
        shp.PictureFormat.CropTop = margin * 0.3
        Exit For
    End If
Next
For marginScan = 1 To endingPoint
    reverseScan = endingPoint - marginScan
    If Not (PixelTest(objPict, midPoint, reverseScan) Like "1677*") Then
        margin = marginScan
        shp.PictureFormat.CropBottom = margin * 0.3
        Exit For
    End If
Next

MsgBox "Please check the crop of the image, then click OK", vbInformation, "Check it out"

shp.PictureFormat.TransparencyColor = rgb(255, 255, 255)
shp.export "\\data\mdbdata\WorkingDB\_docs\Part_Pictures\" & partNum & ".png", ppShapeFormatPNG

On Error Resume Next
pptPres.CLOSE
ppt.Quit
Set CATIA = Nothing
Set ppt = Nothing
Set pptPres = Nothing
Set curSlide = Nothing

Call snackBox("success", "Success!", "Part Picture Uploaded.", Me.name)

exit_handler:
formStatus (False)
Set CATIA = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub findReplaceText_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)

If catiaOpen = False Then
    MsgBox "CATIA must be open to use this function.", vbExclamation, "CATIA Not Open"
    Exit Sub
End If

formStatus (True)

'--Zoom out to fit view on screen
CATIA.ActiveWindow.ActiveViewer.Reframe

'--Set variables
Dim findText As String
Dim replaceText As String
Dim i As Integer, oDoc, oSheets, oViews, oView, oTexts, srcText

findText = Nz(Me.findText, "")
    
If findText = "" Then
    MsgBox "Find text can't be blank", vbInformation, "Woops"
    GoTo exit_handler
End If

replaceText = Nz(Me.replaceText, "")

Set oDoc = CATIA.ActiveDocument
Set oSheets = oDoc.Sheets
Set oViews = oSheets.ActiveSheet.Views
Set oView = oViews.ActiveView
Set oTexts = oView.Texts

If Me.findReplaceExact.Value = False Then GoTo wildCardSearch

On Error Resume Next
For i = 1 To oViews.count
    Set oView = oViews.ITEM(i)
    Set oTexts = oView.Texts
    For Each srcText In oTexts
        If StrConv(srcText.Text, vbUpperCase) = findText Then srcText.Text = replaceText
    Next
Next
GoTo exit_handler

wildCardSearch:
Dim testText As String

On Error Resume Next
For i = 1 To oViews.count
    Set oView = oViews.ITEM(i)
    Set oTexts = oView.Texts
    For Each srcText In oTexts
        testText = StrConv(srcText.Text, vbUpperCase)
        If InStr(testText, findText) Then 'if the find text is in the converted string...
            srcText.Text = Replace(testText, findText, replaceText) 'replace just that portion
        End If
    Next
Next

exit_handler:
formStatus (False)
Set CATIA = Nothing

Exit Sub
Err_Handler:
If Err.number = -2147467259 Then GoTo exit_handler
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub forceLink_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)

If catiaOpen = False Then
    MsgBox "CATIA must be open to use this function.", vbExclamation, "CATIA Not Open"
    Exit Sub
End If

formStatus (True)

Dim myViews
Dim strFilePath As String
Dim newLink
Dim i
Dim myView

Set myViews = CATIA.ActiveDocument.Sheets.ActiveSheet.Views
If myViews.count < 3 Then
    MsgBox ("There are no views to relink in this drawing.")
    GoTo exit_handler
End If

strFilePath = CATIA.FileSelectionBox("Select the file to relink to this drawing", "*.*", 0)
If strFilePath = "" Then GoTo exit_handler
Set newLink = CATIA.Documents.Read(strFilePath)

For i = 3 To myViews.count
    Set myView = myViews.ITEM(i)
    If myView.LockStatus = True Then GoTo nextLoop
    If InStr(myView.name, "Drawing Frame") > 0 Then GoTo nextLoop
    If InStr(myView.name, "zuwaku") > 0 Then GoTo nextLoop
    
    myView.GenerativeLinks.RemoveAllLinks
    myView.GenerativeLinks.AddLink newLink.Product
    myView.Activate
nextLoop:
Next

exit_handler:
formStatus (False)
Set CATIA = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub formatText_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)

If catiaOpen = False Then
    MsgBox "CATIA must be open to use this function.", vbExclamation, "CATIA Not Open"
    Exit Sub
End If

formStatus (True)

Dim oSelection
Dim InputObject(0)
Dim oStatus
Dim fontSize As Double, ratio As Double, spacing As Double
Dim myObject

Set oSelection = CATIA.ActiveDocument.Selection

fontSize = Nz(Me.fontSize, "")
ratio = Nz(Me.ratio, "")
spacing = Nz(Me.spacing, "")

restart:
oSelection.clear
InputObject(0) = "AnyObject"
oStatus = oSelection.SelectElement2(InputObject, "Select a text box", True)

Do While oStatus <> "Cancel"
    Set myObject = oSelection.ITEM(1).Value
    Select Case typeName(myObject)
        Case "DrawingText"
            myObject.SetParameterOnSubString 7, 1, Len(myObject.Text), fontSize * 1000
            myObject.SetParameterOnSubString 14, 1, Len(myObject.Text), ratio
            myObject.SetParameterOnSubString 15, 1, Len(myObject.Text), spacing
        Case "DrawingComponent"
            If myObject.CompRef.Texts.count = 1 Then
                Set myObject = myObject.CompRef.Texts.ITEM(1)
                myObject.SetParameterOnSubString 7, 1, Len(myObject.Text), fontSize * 1000
                myObject.SetParameterOnSubString 14, 1, Len(myObject.Text), ratio
                myObject.SetParameterOnSubString 15, 1, Len(myObject.Text), spacing
            ElseIf myObject.CompRef.Texts.count > 1 Then
                MsgBox "This instantiated component has more than 1 text box." & vbCrLf & "Please select the individual text boxes directly.", vbExclamation, "Invalid Object Selected"
                GoTo restart
            Else
                MsgBox "This instantiated component has 0 text boxes.", vbExclamation, "Invalid Object Selected"
                GoTo restart
            End If
        Case Else
            MsgBox "This function only works on text boxes (for now).", vbExclamation, "Invalid Object Selected"
            GoTo restart
    End Select
    oSelection.clear
    oStatus = oSelection.SelectElement2(InputObject, "Select a text box", True)
Loop

exit_handler:
formStatus (False)
Set CATIA = Nothing

Exit Sub
Err_Handler:
    MsgBox Err.DESCRIPTION, vbOKOnly, "Error Code: " & Err.number
End Sub

Private Sub funkyTextFixer_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)

If catiaOpen = False Then
    MsgBox "CATIA must be open to use this function.", vbExclamation, "CATIA Not Open"
    Exit Sub
End If

formStatus (True)

Dim USel
Dim InputObject(0)
Dim oStatus
Dim X As String
Dim myDimension

Dim iIndex As Long
Dim oBefore As String
Dim oAfter As String
Dim oUpper As String
Dim oLower As String
iIndex = 1

InputObject(0) = "AnyObject"
Set USel = CATIA.ActiveDocument.Selection

Do While oStatus <> "Cancel"
    On Error GoTo exit_handler
    oStatus = USel.SelectElement2(InputObject, "Select element", True)

    Set myDimension = USel.ITEM(1).Value
    
    On Error GoTo CheckText
    Call myDimension.getValue.GetBaultText(iIndex, oBefore, oAfter, oUpper, oLower)
    Call myDimension.getValue.SetBaultText(iIndex, StrConv(oBefore, vbUpperCase), StrConv(oAfter, vbUpperCase), StrConv(oUpper, vbUpperCase), StrConv(oLower, vbUpperCase))
    GoTo nextOne
    
CheckText:
    X = USel.ITEM(1).Value.Text
    USel.ITEM(1).Value.Text = StrConv(X, vbUpperCase)
    
nextOne:
    USel.clear
Loop

exit_handler:
formStatus (False)
Set CATIA = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub hideLines_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)

If catiaOpen = False Then
    MsgBox "CATIA must be open to use this function.", vbExclamation, "CATIA Not Open"
    Exit Sub
End If

formStatus (True)

Dim InputObject(0)
Dim USel
Dim oStatus
Dim oRed, oGreen, oBlue

InputObject(0) = "AnyObject"
Set USel = CATIA.ActiveDocument.Selection

oStatus = USel.SelectElement2(InputObject, "Select a color from an element", True)
Do While oStatus <> "Cancel"
    USel.VisProperties.GetVisibleColor oRed, oGreen, oBlue
    USel.clear
    oStatus = "Cancel"
Loop

InputObject(0) = "DrawingView"
oStatus = USel.SelectElement2(InputObject, "Select a drawing view", True)
Do While oStatus <> "Cancel"
    USel.Search "Color='(" & oRed & "," & oGreen & "," & oBlue & ")'&Drafting.Circle+Drafting.Curve+Drafting.Ellipse+Drafting.Line+Drafting.Point+Drafting.Spline-Drafting.'Generated Item',sel"
    USel.VisProperties.SetShow 1
    USel.clear
    oStatus = USel.SelectElement2(InputObject, "Select a drawing view", True)
Loop

exit_handler:
formStatus (False)
Set CATIA = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub itemSearch_Change()
On Error GoTo Err_Handler

Me.Dirty = False
Me.itemSearch.SelStart = Len(Nz(Me.itemSearch))
Me.drawingItems.Form.filter = "drawingItem Like '*" & Me.itemSearch & "*'"
Me.drawingItems.Form.FilterOn = True

If Me.drawingItems.Form.Recordset.RecordCount < 6 Then
    If Me.drawingItems.Form.ScrollBars <> 1 Then
        Me.drawingItems.Form.ScrollBars = 1
        Me.drawingItems.Form.itemName.Width = 1620
        Me.drawingItems.Form.itemLabel.Width = 1620
    End If
Else
    If Me.drawingItems.Form.ScrollBars <> 2 Then
        Me.drawingItems.Form.ScrollBars = 2
        Me.drawingItems.Form.itemName.Width = 1380
        Me.drawingItems.Form.itemLabel.Width = 1380
        Me.drawingItems.Form.Width = 1560
    End If
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lockViews_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)

If catiaOpen = False Then
    MsgBox "CATIA must be open to use this function.", vbExclamation, "CATIA Not Open"
    Exit Sub
End If

formStatus (True)

Dim myViews
Dim myView
Dim i As Long

Set myViews = CATIA.ActiveDocument.Sheets.ActiveSheet.Views
If myViews.count < 3 Then
    MsgBox ("This drawing has no views.")
    GoTo exit_handler
End If

For i = 3 To myViews.count
    Set myView = myViews.ITEM(i)
    myView.LockStatus = True
Next

exit_handler:
formStatus (False)
Set CATIA = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub tabCatiaMacros_Change()
On Error GoTo Err_Handler

If Me.tabCatiaMacros.TabIndex = 3 Then
    Dim folderName As String
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    
    Dim db As Database
    Dim rs As Recordset
    
    Set db = CurrentDb()
    
    folderName = "\\data\mdbdata\WorkingDB\Pictures\CATIA_Drawing_Items"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderName)
    
    db.Execute "DELETE * FROM tblSessionVariables WHERE drawingItem IS NOT NULL", dbFailOnError
    
    Set rs = db.OpenRecordset("tblSessionVariables")
    For Each file In folder.Files
        With rs
            .addNew
            !drawingItem = Left(file.name, Len(file.name) - 4)
            .Update
        End With
    Next
    
    rs.CLOSE
    Set rs = Nothing
    Set db = Nothing
    
    resetDrawingItemFilter
    Set fso = Nothing
    Set folder = Nothing
    
    Me.drawingItems.Requery
    
    Me.itemImage.Picture = "\\data\mdbdata\WorkingDB\Pictures\CATIA_Drawing_Items\Placeholder\placeholder.png"
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub unlockViews_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)

If catiaOpen = False Then
    MsgBox "CATIA must be open to use this function.", vbExclamation, "CATIA Not Open"
    Exit Sub
End If

formStatus (True)

Dim myViews
Dim myView
Dim i As Long

Set myViews = CATIA.ActiveDocument.Sheets.ActiveSheet.Views
If myViews.count < 3 Then
    MsgBox ("This drawing has no views.")
    GoTo exit_handler
End If

For i = 3 To myViews.count
    Set myView = myViews.ITEM(i)
    myView.LockStatus = False
Next

exit_handler:
formStatus (False)
Set CATIA = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub multiTextChange_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)

If catiaOpen = False Then
    MsgBox "CATIA must be open to use this function.", vbExclamation, "CATIA Not Open"
    Exit Sub
End If

formStatus (True)

Dim newText
Dim InputObject(0)
Dim USel 'As Selection
Dim oStatus

newText = Nz(Me.multiText, "")

InputObject(0) = "AnyObject"
Set USel = CATIA.ActiveDocument.Selection

oStatus = USel.SelectElement2(InputObject, "Select something in Specification Tree", True)

Do While oStatus <> "Cancel"
    USel.ITEM(1).Value.Text = newText
    USel.clear

    oStatus = USel.SelectElement2(InputObject, "Select something in Specification Tree", True)
Loop

exit_handler:
formStatus (False)
Set CATIA = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub pickCustom_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)

Dim r, G, B, pickerVal

pickerVal = Hex(colorPicker())

r = Mid(pickerVal, 5, 2)
G = Mid(pickerVal, 3, 2)
B = Mid(pickerVal, 1, 2)

Me.catiaCustomColor = r & G & B
Me.refresh

If isHex("#" & Nz(Me.txtCustom.Value)) = False Then
    MsgBox "You must enter a valid HEX code.", vbExclamation, "Invalid HEX Code"
    Exit Sub
End If

Dim ioR As Long, ioG As Long, ioB As Long

ioR = val("&H" & Mid(Me.txtCustom.Value, 1, 2))
ioG = val("&H" & Mid(Me.txtCustom.Value, 3, 2))
ioB = val("&H" & Mid(Me.txtCustom.Value, 5, 2))

Me.dispCustom.BackColor = rgb(ioR, ioG, ioB)

If Me.btnGradient.Value = True Then Me.dispCustom.Gradient = 15

exit_handler:
formStatus (False)
Set CATIA = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub publishDwgName_Enter()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)

If catiaOpen = False Then
    MsgBox "CATIA must be open to use this function.", vbExclamation, "CATIA Not Open"
    Exit Sub
End If

formStatus (True)

Dim fileNameFull As String, filePathFull As String, sheetName As String, newFileName As String, folderLoc As String
fileNameFull = CATIA.ActiveDocument.name

If Right(fileNameFull, 11) <> ".CATDrawing" Then
    MsgBox "Must have a CATDrawing active to export it", vbCritical, "Hmm...?"
    GoTo exit_handler
End If

fileNameFull = Left(fileNameFull, Len(fileNameFull) - 11)

sheetName = CATIA.ActiveDocument.Sheets.ActiveSheet.name

If InStr(sheetName, ".") Then
    MsgBox "Sheet name cannot have a period in it, please modify then try again", vbInformation, "Try again"
    GoTo exit_handler
End If

Dim objDoc As Object, objParam As Object, objItems As Object 'find the parameter for part number
Set objDoc = CATIA.ActiveDocument
Set objParam = objDoc.Parameters.RootParameterSet
On Error GoTo paramError
Set objItems = objParam.DirectParameters.ITEM("Drawing\NF_Part_No")
On Error GoTo Err_Handler

Dim partNum As String, dashLoc As Integer 'grab part number
If Left(objItems.Value, 1) = "(" Then
    partNum = Right(objItems.Value, 8)
Else
    partNum = Left(objItems.Value, 8)
End If
partNum = Replace(partNum, "-", "_")
newFileName = partNum & "_" & sheetName & ".pdf"

Me.publishDwgName = newFileName

exit_handler:
formStatus (False)
Set CATIA = Nothing
Exit Sub

Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
    GoTo exit_handler
    
paramError:
    MsgBox "You need to have 3Dex Properties set up to use this function", vbInformation, "Error!"
End Sub

Private Sub setupRevBlock_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)

If catiaOpen = False Then
    MsgBox "CATIA must be open to use this function.", vbExclamation, "CATIA Not Open"
    Exit Sub
End If

formStatus (True)

Dim drawingDocument1
Dim selection1
Set drawingDocument1 = CATIA.ActiveDocument
Set selection1 = drawingDocument1.Selection

selection1.Search ("Name=Text.17,scr") 'revBlock Date
If selection1.count = 0 Then
    MsgBox "Please activate the revision block and try again.", vbExclamation, "Revision Block Not Active"
Else
    selection1.ITEM(1).Value.Text = Format(Date, "mm/dd/yy")
    selection1.clear
    
    selection1.Search ("Name=Text.8,scr") 'revBlock Revr.
    selection1.ITEM(1).Value.Text = UCase(getFullName)
    selection1.clear
    MsgBox "Revision block setup complete.", , "Success"
End If

exit_handler:
formStatus (False)
Set CATIA = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub unbalTolSearch_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)

If catiaOpen = False Then
    MsgBox "CATIA must be open to use this function.", vbExclamation, "CATIA Not Open"
    Exit Sub
End If

formStatus (True)

Dim oSelection
Dim dimension
Dim i As Integer
Dim odUpTol As Double
Dim odLowTol As Double
Dim gdt
Dim tolValue
Dim tolCode As String, dimArray, dimItem

CATIA.ActiveWindow.ActiveViewer.Reframe

Set dimArray = CreateObject("System.Collections.ArrayList")
Set oSelection = CATIA.ActiveDocument.Selection

'Select all dimensions and check for unbalanced tolerances
oSelection.Search "Name=Dimension*,scr"
For i = 1 To oSelection.count
    Set dimension = oSelection.ITEM(i).Value
    Call dimension.GetTolerances(0, "", "", "", odUpTol, odLowTol, 0)
    If odUpTol + odLowTol <> 0 Then
        If dimArray.Contains(dimension) = False Then dimArray.Add dimension
    End If
    odUpTol = 0
    odLowTol = 0
Next
oSelection.clear

'Select all GD&T and check for unilateral symbol
oSelection.Search "Name='Geometrical Tolerance*',scr"
For i = 1 To oSelection.count
    Set gdt = oSelection.ITEM(i).Value
    Set tolValue = gdt.GetTextRange(1, 0)
    tolCode = StrConv(tolValue.Text, vbUnicode)
    If InStr(tolCode, "Ê$") > 0 Then
        If dimArray.Contains(gdt) = False Then dimArray.Add gdt
    End If
Next
oSelection.clear

'If there are no unbalanced tolerances
If dimArray.count = 0 Then
    MsgBox "There are no unbalanced tolerances.", vbInformation, "Unbalanced Tolerance Search"
    GoTo exit_handler
End If

'Select all unbalanced tolerances and change to red
For Each dimItem In dimArray
    If IsObject(dimItem) Then oSelection.Add dimItem
Next
oSelection.VisProperties.SetRealColor 255, 0, 0, 1
oSelection.clear

'Final message box
If dimArray.count = 1 Then
    MsgBox dimArray.count & " unbalanced tolerance found." & vbCrLf & "Dimension is highlighted red.", vbExclamation, "Unbalanced Tolerance Search"
Else
    MsgBox dimArray.count & " unbalanced tolerances found." & vbCrLf & "Dimensions are highlighted red.", vbExclamation, "Unbalanced Tolerance Search"
End If

exit_handler:
formStatus (False)
Set CATIA = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub createBalloonPrint_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)

If catiaOpen = False Then
    MsgBox "CATIA must be open to use this function.", vbExclamation, "CATIA Not Open"
    Exit Sub
End If

formStatus (True)

'RED = -16776961
'BLUE = 65535

Dim CATIA As Object
Dim oSelection
Dim i As Long, j As Long, k As Long
Dim arrGDT() As String
Dim oView
Dim oDimension
Dim showState As Long
Dim tolValue As String, refCount As Long
Dim balloonCount As Long
Dim distFromDim As Double
Dim balloon
Dim oBefore As String
Dim oValues(7)
Dim oGeomInfos(3)
Dim leftSide As Double
Dim rightSide As Double
Dim x1 As Double, x2 As Double, x3 As Double
Dim y1 As Double, y2 As Double, y3 As Double
Dim dimLength As Double

Set CATIA = GetObject(, "CATIA.Application")
Set oSelection = CATIA.ActiveDocument.Selection

''Create a blank drawing sheet for the balloon print
'oSelection.Add CATIA.ActiveDocument.Sheets.ActiveSheet
'oSelection.Copy
'oSelection.clear
'oSelection.Add CATIA.ActiveDocument.DrawingRoot
'oSelection.Paste
'oSelection.clear
'CATIA.ActiveDocument.Sheets.ActiveSheet.name = "BP_1000"

CATIA.ActiveWindow.ActiveViewer.Reframe

'Dim fso As Object
'Dim xlApp As Excel.Application
'Dim wb As Excel.Workbook
'Dim ws As Excel.Worksheet
'Dim objParam As Object
'
''on error goto paramError UNCOMMENT LATER
'Set objParam = CATIA.ActiveDocument.Parameters.RootParameterSet.DirectParameters
''on error goto err_handler UNCOMMENT LATER
'
'Set fso = CreateObject("Scripting.FileSystemObject")
'fso.CopyFile "\\nas01\lab\lab\LABDOCS\Lab Forms\Blank ISIR.xls", "H:\Documents\" & Left(CATIA.ActiveDocument.name, 5) & "_ISIR.xls"
'Set xlApp = Excel.Application
'xlApp.Visible = True 'COMMENT LATER
'Set wb = xlApp.Workbooks.open("H:\Documents\" & Left(CATIA.ActiveDocument.name, 5) & "_ISIR.xls", 0)
'Set ws = wb.Worksheets("ISIR")
'
''Fill out the basic info at the top of the ISIR
'ws.Range("E6").Value = removeReferenceString(objParam.ITEM("Drawing\NF_Part_No").Value)
'ws.Range("E7").Value = objParam.ITEM("Drawing\NF_Product_Name").Value
'ws.Range("E8").Value = removeReferenceString(objParam.ITEM("Drawing\NF_Material_Symbol").Value)
'ws.Range("E9").Value = objParam.ITEM("Drawing\NF_Customer").Value
'ws.Range("E10").Value = removeReferenceString(objParam.ITEM("Drawing\NF_Customer_Part_No").Value)
'ws.Range("I7").Value = Right(Split(CATIA.ActiveDocument.name, ".")(0), 4)
'ws.Range("K7").Value = fso.GetFile(CATIA.ActiveDocument.fullName).DateLastModified

'Loop through visible GD&T and add to an array
oSelection.Search "Name='Geometrical Tolerance*',scr"
For i = 1 To oSelection.count
    ReDim Preserve arrGDT(i - 1)
    arrGDT(i - 1) = oSelection.ITEM(i).Value.name
Next
oSelection.clear

'Loop through all visible views
oSelection.Search "CATDrwSearch.DrwView,scr"
For i = 1 To oSelection.count
    Set oView = oSelection.ITEM(i).Value
    
    'Loop through GD&T in each view
    For j = 1 To oView.GDTs.count
        Set oDimension = oView.GDTs.ITEM(j)
        
        'Loop through array of visible GD&T (defined previously) and add balloon
        'This is to ensure hidden GD&T do not get ballooned
        For k = 0 To UBound(arrGDT)
            If oDimension.name = arrGDT(k) Then
                tolValue = oDimension.GetTextRange(1, 0).Text
                refCount = oDimension.GetReferenceNumber(1)
                
                balloonCount = balloonCount + 1
                If oDimension.Leaders.ITEM(1).AnchorPoint < 10 Then
                    distFromDim = (2.5 + Len(CStr(balloonCount)) + 7 + (3.5 * Len(tolValue)) + (6.3 * refCount)) / oView.Scale
                    'This number ---^ controls the distance from the right edge of the GD&T box to the center of the balloon
                    Set balloon = oView.Texts.Add(CStr(balloonCount), oDimension.X + distFromDim, oDimension.Y - (3.5 / oView.Scale))
                Else
                    distFromDim = (2.2 + Len(CStr(balloonCount))) / oView.Scale
                    'This number ---^ controls the distance from the left edge of the GD&T box to the center of the balloon
                    Set balloon = oView.Texts.Add(CStr(balloonCount), oDimension.X - distFromDim, oDimension.Y - (3.5 / oView.Scale))
                End If
                
                balloon.AnchorPosition = 5
                balloon.name = "Bubble." & balloonCount
                balloon.FrameType = 3
                If balloonCount < 10 Then balloon.SetFontSize 0, 0, 3.5
                balloon.TextProperties.Update
            End If
        Next
    Next
Next
oSelection.clear

'Loop through all visible dimensions and add balloon
oSelection.Search "Name=Dimension*,scr"
For i = 1 To oSelection.count
    Set oDimension = oSelection.ITEM(i).Value
    Call oDimension.getValue.GetBaultText(1, oBefore, "", "", "")
    If (InStr(oBefore, "(") > 0 Or InStr(StrConv(oBefore, vbUnicode), "ÿ") > 0) And InStr(StrConv(oBefore, vbUnicode), "&") = 0 Then GoTo nextLoop
    If oDimension.getValue.FakeDimType <> 0 And (oDimension.getValue.GetFakeDimValue(1) = "" Or oDimension.getValue.GetFakeDimValue(1) = " " Or InStr(StrConv(oDimension.getValue.GetFakeDimValue(1), vbUnicode), "; ") > 0) Then GoTo nextLoop
    Set oView = oDimension.Parent.Parent
    
    oDimension.GetBoundaryBox oValues
    oDimension.GetDimLine.GetGeomInfo oGeomInfos
    
    leftSide = Sqr((oValues(0) - oGeomInfos(0)) ^ 2 + (oValues(1) - oGeomInfos(1)) ^ 2)
    rightSide = Sqr((oValues(2) - oGeomInfos(0)) ^ 2 + (oValues(3) - oGeomInfos(1)) ^ 2)
    
    x1 = (oValues(0) + oValues(4)) / 2
    y1 = (oValues(1) + oValues(5)) / 2
    x2 = (oValues(2) + oValues(6)) / 2
    y2 = (oValues(3) + oValues(7)) / 2
    dimLength = Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2)
    
    distFromDim = (2 + Len(CStr(balloonCount))) / oView.Scale
    'This number --^ controls the distance from the edge of the dimension bounding box to the center of the balloon
    
    If leftSide < rightSide Then
        x3 = x2 + (x2 - x1) / dimLength * distFromDim
        y3 = y2 + (y2 - y1) / dimLength * distFromDim
    Else
        x3 = x1 + (x1 - x2) / dimLength * distFromDim
        y3 = y1 + (y1 - y2) / dimLength * distFromDim
    End If
    
    balloonCount = balloonCount + 1
    Set balloon = oView.Texts.Add(CStr(balloonCount), x3, y3)
    balloon.AnchorPosition = 5
    
    balloon.name = "Bubble." & balloonCount
    balloon.FrameType = 3
    If balloonCount < 10 Then balloon.SetFontSize 0, 0, 3.5
    balloon.TextProperties.Update
nextLoop:
Next
oSelection.clear

'Select all balloons again and change color
oSelection.Search "Name=Bubble*,scr"
oSelection.VisProperties.SetRealColor 255, 0, 0, 1
For i = 1 To oSelection.count
    oSelection.ITEM(i).Value.TextProperties.Color = 65535
    oSelection.ITEM(i).Value.TextProperties.Update
Next
oSelection.clear

'Final confirmation message
MsgBox balloonCount & " dimensions ballooned." & vbCrLf & "Please review for readability.", , "Balloon Print Complete"

exit_handler:
formStatus (False)
Set CATIA = Nothing

Exit Sub
paramError:
    MsgBox "You need to have parameters set up to use this function.", vbInformation, "No Parameters"
    Exit Sub

Err_Handler:
    MsgBox Err.DESCRIPTION, vbOKOnly, "Error Code: " & Err.number
End Sub

Private Function catiaOpen() As Boolean
On Error GoTo notOpen

Set CATIA = GetObject(, "CATIA.Application")

catiaOpen = True
Exit Function

notOpen:
catiaOpen = False
End Function

Private Function isHex(colorCode As String) As Boolean

With New RegExp
    .Pattern = "^#[0-9A-F]{1,6}$"
    .IgnoreCase = True
    isHex = .test(colorCode)
End With

End Function

Private Sub viewCalloutFixer_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)

If catiaOpen = False Then
    MsgBox "CATIA must be open to use this function.", vbExclamation, "CATIA Not Open"
    Exit Sub
End If

formStatus (True)

Dim oSelection
Dim InputObject(0)
Dim oStatus
Dim myText, myText2
Dim showState As Long
Dim charCount As Long
Dim myChar As String
Dim splitPoint As Long
Dim i As Long, k As Long

Set oSelection = CATIA.ActiveDocument.Selection

restart:
oSelection.clear
InputObject(0) = "AnyObject"
oStatus = oSelection.SelectElement2(InputObject, "Select view title text", True)

Do While oStatus <> "Cancel"
    Set myText = oSelection.ITEM(1).Value
    If typeName(myText) = "DrawingComponent" Then
        Set myText = myText.CompRef.Texts.ITEM(1)
    ElseIf typeName(myText) <> "DrawingText" Then
        MsgBox "This function only works on text boxes", vbExclamation, "Please Select a Text Box"
        GoTo restart
    End If
    
    For i = 1 To myText.Parent.count
        Set myText2 = myText.Parent.ITEM(i)
        On Error Resume Next
        If myText2.AssociativeElement Is Nothing Then
            GoTo nextLoop
        ElseIf myText2.AssociativeElement.name = myText.name Then
            On Error GoTo Err_Handler
            oSelection.clear
            oSelection.Add myText2
            oSelection.VisProperties.GetShow showState
            If showState = 0 Then
                myText.Text = myText.Text & myText2.Text
                oSelection.VisProperties.SetShow 1
            End If
        End If
        On Error GoTo Err_Handler
nextLoop:
    Next i
    oSelection.clear
    
    If InStr(myText.Parent.Parent.name, "SEC") > 0 Then
        If InStr(myText.Text, "SECTION") > 0 Then myText.Text = Replace(myText.Text, "SECTION", "SEC")
        If InStr(myText.Text, "SEC") = 0 Then myText.Text = "SEC " & myText.Text
    End If
    
    i = 1
    charCount = Len(myText.Text)
    Do While i <= charCount
        myChar = Mid(myText.Text, i, 1)
        Select Case myChar
            Case "-"
                myText.Text = Left(myText.Text, i - 1) & Right(myText.Text, Len(myText.Text) - i)
                i = i - 1
                charCount = charCount - 1
            Case "("
                Do While Mid(myText.Text, i - 1, 1) = " "
                    myText.Text = Left(myText.Text, i - 2) & Right(myText.Text, Len(myText.Text) - i + 1)
                    i = i - 1
                    charCount = charCount - 1
                Loop
                If Mid(myText.Text, i - 1, 1) <> Chr(10) Then
                    myText.Text = Left(myText.Text, i - 1) & vbCrLf & Right(myText.Text, Len(myText.Text) - i + 1)
                    i = i + 1
                    charCount = charCount + 1
                End If
                If InStr(myText.Text, "SCALE") = 0 Then
                    myText.Text = Left(myText.Text, i) & "SCALE " & Right(myText.Text, Len(myText.Text) - i)
                    i = i + 6
                    charCount = charCount + 6
                End If
        End Select
        i = i + 1
    Loop
    
    splitPoint = InStr(myText.Text, Chr(10))
    If splitPoint > 0 Then
        myText.SetParameterOnSubString 0, 1, splitPoint, 1
        myText.SetParameterOnSubString 0, splitPoint, Len(myText.Text), 0
        myText.SetParameterOnSubString 2, 1, splitPoint, 1
        myText.SetParameterOnSubString 2, splitPoint, Len(myText.Text), 0
        myText.SetParameterOnSubString 7, 1, splitPoint, 5000
        myText.SetParameterOnSubString 7, splitPoint, Len(myText.Text), 3500
    Else
        myText.SetParameterOnSubString 0, 1, Len(myText.Text), 1
        myText.SetParameterOnSubString 2, 1, Len(myText.Text), 1
        myText.SetParameterOnSubString 7, 1, Len(myText.Text), 5000
    End If
    
    myText.SetFontName 1, Len(myText.Text), "KANJ"
    myText.SetParameterOnSubString 13, 1, Len(myText.Text), 1
    myText.SetParameterOnSubString 14, 1, Len(myText.Text), 75
    myText.SetParameterOnSubString 15, 1, Len(myText.Text), 25
    
    oSelection.clear
    oStatus = oSelection.SelectElement2(InputObject, "Select view title text", True)
Loop

exit_handler:
formStatus (False)
Set CATIA = Nothing

Exit Sub
Err_Handler:
    MsgBox Err.DESCRIPTION, vbOKOnly, "Error Code: " & Err.number
End Sub

Private Sub countSpecialCharacteristics_Click()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)

If catiaOpen = False Then
    MsgBox "CATIA must be open to use this function.", vbExclamation, "CATIA Not Open"
    Exit Sub
End If

formStatus (True)

Dim criticalCount As Long
Dim fMarks() As String
Dim fMarksCount() As Long
Dim fQuantitySymbs As Collection
Dim message, question, buttons, Title, response
Dim i As Long

criticalCount = countCriticalMarks
countFMarks fMarks, fMarksCount, fQuantitySymbs

question = vbCrLf & "No special characteristics found."
buttons = vbOKOnly
Title = "No Special Characteristics Found"

message = "Critical Marks" & vbTab & "=" & vbTab & criticalCount & vbCrLf
If criticalCount > 0 Then
    question = vbCrLf & "Do you want to add quantity symbols?"
    buttons = vbYesNo
    Title = "Special Characteristics Found"
End If

For i = 0 To UBound(fMarksCount)
    message = message & fMarks(i) & vbTab & vbTab & "=" & vbTab & fMarksCount(i) & vbTab & vbCrLf
    If fMarksCount(i) > 0 Then
        question = vbCrLf & "Do you want to add quantity symbols?"
        buttons = vbYesNo
        Title = "Special Characteristics Found"
    End If
Next i

response = MsgBox(message & question, buttons, Title)

If response = vbYes Then Call addQuantitySymbols(criticalCount, fMarks, fMarksCount, fQuantitySymbs)

formStatus (False)

Exit Sub
Err_Handler:
    MsgBox Err.DESCRIPTION, vbOKOnly, "Error Code: " & Err.number
    formStatus (False)
End Sub

Function countCriticalMarks() As Long
On Error GoTo Err_Handler

Dim oSelection
Dim myDimension
Dim oBefore As String, oPrefix As String
Dim myGDT
Dim upperText
Dim myComponent
Dim criticalCount As Long
Dim i As Long

Set oSelection = CATIA.ActiveDocument.Selection
CATIA.ActiveWindow.ActiveViewer.Reframe
oSelection.clear

oSelection.Search "Name=Dimension*,scr"
For i = 1 To oSelection.count
    Set myDimension = oSelection.ITEM(i).Value
    Call myDimension.getValue.GetBaultText(1, oBefore, "", "", "")
    Call myDimension.getValue.GetPSText(1, oPrefix, "")
    If (InStr(StrConv(oBefore, vbUnicode), "¼%") And InStr(StrConv(oBefore, vbUnicode), "(") = 0) Then criticalCount = criticalCount + 1
    If oPrefix = "<Black Triangle Down>" Then criticalCount = criticalCount + 1
Next i
oSelection.clear

oSelection.Search "Name='Geometrical Tolerance*',scr"
For i = 1 To oSelection.count
    Set myGDT = oSelection.ITEM(i).Value
    Set upperText = myGDT.GetTextRange(0, 0)
    If InStr(StrConv(upperText.Text, vbUnicode), "¼%") > 0 Then criticalCount = criticalCount + 1
Next i
oSelection.clear

oSelection.Search "Name='Critical Mark*',scr"
For i = 1 To oSelection.count
    Set myComponent = oSelection.ITEM(i).Value
    If InStr(StrConv(myComponent.name, vbLowerCase), "critical mark") Then criticalCount = criticalCount + 1
Next i
oSelection.clear

countCriticalMarks = criticalCount

Exit Function
Err_Handler:
    MsgBox Err.DESCRIPTION, vbOKOnly, "Error Code: " & Err.number
End Function

Function countFMarks(ByRef fMarks() As String, ByRef fMarksCount() As Long, ByRef fQuantitySymbs As Collection)
On Error GoTo Err_Handler

Dim oSelection
Dim visComponents As Collection
Dim showState As Long
Dim myComponent
Dim myComponent2
Dim i As Long, j As Long, k As Long

Set oSelection = CATIA.ActiveDocument.Selection
Set visComponents = New Collection
fMarks = Split("fa,fb,fe,fr,fs", ",")
ReDim Preserve fMarksCount(4)
Set fQuantitySymbs = New Collection

CATIA.ActiveWindow.ActiveViewer.Reframe

oSelection.clear
oSelection.Search "Type='2D Component Instance',scr"
For i = 1 To oSelection.count
    visComponents.Add oSelection.ITEM(i).Value
Next i
oSelection.clear

For i = 1 To visComponents.count
    Set myComponent = visComponents.ITEM(i)
    For j = 0 To UBound(fMarks)
        If InStr(StrConv(myComponent.name, vbLowerCase), fMarks(j)) And InStr(StrConv(myComponent.name, vbLowerCase), "frame") = 0 And InStr(StrConv(myComponent.name, vbLowerCase), "reference") = 0 Then
            If InStr(StrConv(myComponent.name, vbLowerCase), "*") = 0 And InStr(StrConv(myComponent.name, vbLowerCase), "quantity") = 0 Then
                If withinNotes(myComponent) = False Then fMarksCount(j) = fMarksCount(j) + 1
            Else
                fQuantitySymbs.Add myComponent
            End If
        End If
        For k = 1 To myComponent.CompRef.Components.count
            Set myComponent2 = myComponent.CompRef.Components.ITEM(k)
            If InStr(StrConv(myComponent2.name, vbLowerCase), fMarks(j)) Then
                If withinNotes(myComponent2) = False Then
                    oSelection.Add myComponent2
                    oSelection.VisProperties.GetShow showState
                    oSelection.clear
                    If showState = 0 Then fMarksCount(j) = fMarksCount(j) + 1
                End If
            End If
        Next k
    Next j
Next i

Exit Function
Err_Handler:
    MsgBox Err.DESCRIPTION, vbOKOnly, "Error Code: " & Err.number
End Function

Function withinNotes(fMark) As Boolean
On Error GoTo Err_Handler

Dim oSelection
Dim markX As Double
Dim myView
Dim myText
Dim noteX As Double
Dim showState As Long
Dim i As Long

Set oSelection = CATIA.ActiveDocument.Selection
Set myView = fMark.Parent.Parent
markX = fMark.X
noteX = markX

For i = 1 To myView.Texts.count
    Set myText = myView.Texts.ITEM(i)
    If (InStr(Left(myText.Text, InStr(myText.Text, Chr(10))), "NOTE") > 0 _
     Or InStr(Left(myText.Text, InStr(myText.Text, Chr(10))), "SPEC") > 0 _
     Or (InStr(StrConv(Left(myText.Text, InStr(myText.Text, Chr(10))), vbUnicode), "ÕN") > 0 _
     And InStr(StrConv(Left(myText.Text, InStr(myText.Text, Chr(10))), vbUnicode), "Øi") > 0)) _
     And InStr(myText.Text, Chr(10)) > 1 Then
        oSelection.clear
        oSelection.Add myText
        oSelection.VisProperties.GetShow showState
        If showState = 0 Then
            noteX = myText.X
        End If
    End If
Next i

If Abs(markX - noteX) > 5 And markX > noteX Then
    withinNotes = True
Else
    withinNotes = False
End If

Exit Function
Err_Handler:
    MsgBox Err.DESCRIPTION, vbOKOnly, "Error Code: " & Err.number
End Function

Sub addQuantitySymbols(criticalCount As Long, fMarks() As String, fMarksCount() As Long, fQuantitySymbs As Collection)
On Error GoTo Err_Handler

Dim oSelection
Dim mySheet
Dim myView
Dim myComponentInst
Dim sheetSize As String
Dim detailSheet
Dim drawingFrame
Dim instX As Double, instY As Double
Dim fQuantitySymbol
Dim critQuantitySymbol
Dim assembly As Boolean
Dim i As Long, j As Long, k As Long

Set oSelection = CATIA.ActiveDocument.Selection
Set mySheet = CATIA.ActiveDocument.Sheets.ActiveSheet
Set detailSheet = CATIA.ActiveDocument.Sheets.ITEM("Detail")
sheetSize = CATIA.ActiveDocument.Sheets.ActiveSheet.PaperName
oSelection.clear

oSelection.Search "Name='Critical Quantity'*,scr"
If oSelection.count > 0 Then oSelection.Delete
oSelection.clear

For i = 1 To fQuantitySymbs.count
    oSelection.Add fQuantitySymbs(i)
    oSelection.Delete
Next i
oSelection.clear

For i = 1 To mySheet.Views.count
    If InStr(mySheet.Views.ITEM(i).name, "zuwaku") > 0 Or InStr(mySheet.Views.ITEM(i).name, "zuwaku") > 0 Then
        Set myView = mySheet.Views.ITEM(i)
        GoTo continue
    End If
Next i

continue:
If myView.name = "Drawing Frame" Then
    instY = 72
    Select Case True
        Case InStr(sheetSize, "A0")
            Set drawingFrame = detailSheet.Views.ITEM("A0 Frame")
            instX = 939
        Case InStr(sheetSize, "A1")
            Set drawingFrame = detailSheet.Views.ITEM("A1 Frame")
            instX = 591
        Case InStr(sheetSize, "A2")
            Set drawingFrame = detailSheet.Views.ITEM("A2 Frame")
            instX = 344
        Case InStr(sheetSize, "A3")
            Set drawingFrame = detailSheet.Views.ITEM("A3 Frame")
            instX = 169.68
    End Select
    For i = 1 To drawingFrame.Components.count
        If InStr(drawingFrame.Components.ITEM(i).name, "Additional Material Block") Then
            instX = instX + 30
            instY = instY + 12
        End If
        If InStr(drawingFrame.Components.ITEM(i).name, "BOM") Then assembly = True
        If InStr(drawingFrame.Components.ITEM(i).name, "PARTS_LIST") Then assembly = True
    Next i
Else
    instX = -240
    instY = 62
    oSelection.Search "Name='Material_Symbol'*,scr"
    If oSelection.count > 1 Then
        instX = instX + (30 * (oSelection.count - 1))
        instY = instY + (12 * (oSelection.count - 1))
    End If
    oSelection.Search "Name='PARTS_LIST'*,scr"
    If oSelection.count > 0 Then assembly = True
    oSelection.clear
End If

k = 0
For i = 0 To UBound(fMarksCount)
    If Not IsEmpty(fQuantitySymbol) Then Set fQuantitySymbol = Nothing
    If fMarksCount(i) > 0 Then
        For j = 1 To detailSheet.Views.count
            If InStr(detailSheet.Views.ITEM(j).name, fMarks(i) & " Quantity") Then
                Set fQuantitySymbol = detailSheet.Views.ITEM(j)
                fQuantitySymbol.Texts.ITEM(1).Text = fMarks(i) & " x " & fMarksCount(i)
            End If
        Next j
        If IsEmpty(fQuantitySymbol) Then
            Set fQuantitySymbol = detailSheet.Views.Add(fMarks(i) & " Quantity")
            fQuantitySymbol.X = 151
            fQuantitySymbol.Y = 124 - (18 * i)
            fQuantitySymbol.Texts.Add fMarks(i) & " x " & fMarksCount(i), 0.94287109375, 5.67923736572266
        End If
        fQuantitySymbol.Texts.ITEM(1).SetParameterOnSubString 12, 1, 2, 3
        fQuantitySymbol.Texts.ITEM(1).SetParameterOnSubString 12, 3, Len(fQuantitySymbol.Texts.ITEM(1).Text), 0
        fQuantitySymbol.Texts.ITEM(1).SetParameterOnSubString 7, 1, Len(fQuantitySymbol.Texts.ITEM(1).Text), 3500
        fQuantitySymbol.Texts.ITEM(1).SetParameterOnSubString 14, 1, Len(fQuantitySymbol.Texts.ITEM(1).Text), 100
        fQuantitySymbol.Texts.ITEM(1).SetParameterOnSubString 15, 1, Len(fQuantitySymbol.Texts.ITEM(1).Text), 20
        fQuantitySymbol.Texts.ITEM(1).AnchorPosition = 2
        If assembly Then
            Set myComponentInst = myView.Components.Add(fQuantitySymbol, instX, instY + (10 * k))
        Else
            Set myComponentInst = myView.Components.Add(fQuantitySymbol, instX + (28 * k), instY)
        End If
        k = k + 1
    End If
Next i

instX = instX + 8.9
If criticalCount > 0 Then
    For i = 1 To detailSheet.Views.count
        If InStr(detailSheet.Views.ITEM(i).name, "Critical Quantity") Then
            Set critQuantitySymbol = detailSheet.Views.ITEM(i)
            critQuantitySymbol.Texts.ITEM(1).Text = "x " & criticalCount
        End If
    Next i
    If IsEmpty(critQuantitySymbol) Then
        Set critQuantitySymbol = detailSheet.Views.Add("Critical Quantity")
        critQuantitySymbol.X = 160.31
        critQuantitySymbol.Y = 138.46
        critQuantitySymbol.Texts.Add "x " & criticalCount, 3.94, 5.82
        critQuantitySymbol.Texts.ITEM(1).SetParameterOnSubString 7, 1, Len(critQuantitySymbol.Texts.ITEM(1).Text), 3500
        critQuantitySymbol.Texts.ITEM(1).SetParameterOnSubString 14, 1, Len(critQuantitySymbol.Texts.ITEM(1).Text), 100
        critQuantitySymbol.Texts.ITEM(1).SetParameterOnSubString 15, 1, Len(critQuantitySymbol.Texts.ITEM(1).Text), 20
        critQuantitySymbol.Texts.ITEM(1).AnchorPosition = 2
        critQuantitySymbol.Texts.Add ChrW(9660), -7.4, 10.29
        critQuantitySymbol.Texts.ITEM(2).SetParameterOnSubString 7, 1, 1, 10000
        critQuantitySymbol.Texts.ITEM(2).SetParameterOnSubString 14, 1, 1, 75
        critQuantitySymbol.Texts.ITEM(2).SetParameterOnSubString 15, 1, 1, 25
        critQuantitySymbol.Texts.ITEM(2).AnchorPosition = 1
    End If
    If assembly Then
        Set myComponentInst = myView.Components.Add(critQuantitySymbol, instX, instY + (10 * k))
    Else
        Set myComponentInst = myView.Components.Add(critQuantitySymbol, instX + (28 * k), instY)
    End If
    k = k + 1
End If

MsgBox "Quantity symbols added successfully." & vbCrLf & "Please double check positioning.", , "Success"

Exit Sub
Err_Handler:
    MsgBox Err.DESCRIPTION, vbOKOnly, "Error Code: " & Err.number
End Sub

Sub test()

Dim objShell
Set objShell = CreateObject("Wscript.Shell")

objShell.Run "\\data\mdbdata\WorkingDB\_docs\Catia_Custom_Files\Macro_Data\Test\addParentheses.vbs"

Set objShell = Nothing

End Sub
