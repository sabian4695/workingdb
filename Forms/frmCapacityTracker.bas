Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'Private Sub btnDetails_Click()
'On Error GoTo Err_Handler
'
'If CurrentProject.AllForms("frmCPC_Dashboard").IsLoaded = True Then DoCmd.CLOSE acForm, "frmCPC_Dashboard"
'DoCmd.OpenForm "frmCPC_Dashboard", , , "projectNumber = '" & Me.projectNumber & "'"
'
'Exit Sub
'Err_Handler:
'    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
'End Sub
'
'Private Sub cboPartNumber_AfterUpdate()
'On Error GoTo Err_Handler
'Filter_Form
'Exit Sub
'Err_Handler:
'    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
'End Sub
'
'Private Sub cpcTrackerHelp_Click()
'On Error GoTo Err_Handler
'
'Call openPath(mainFolder(Me.ActiveControl.name))
'Exit Sub
'Err_Handler:
'    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
'End Sub
'
'Private Sub lblProjectNumber_Click()
'On Error GoTo Err_Handler
'
'Dim newLabel As String
'newLabel = labelUpdate(Me.lblProjectNumber.Caption)
'
'Reset_Labels
'Me.lblProjectNumber.Caption = newLabel
'
'Me.OrderBy = "projectNumber " & direction(newLabel)
'
'Exit Sub
'Err_Handler:
'    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
'End Sub
'
'Private Sub lblPN_Click()
'On Error GoTo Err_Handler
'
'Dim newLabel As String
'newLabel = labelUpdate(Me.lblPN.Caption)
'
'Reset_Labels
'Me.lblPN.Caption = newLabel
'
'Me.OrderBy = "partNumber " & direction(newLabel)
'
'Exit Sub
'Err_Handler:
'    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
'End Sub
'
'Private Sub lblOwner_Click()
'On Error GoTo Err_Handler
'
'Dim newLabel As String
'newLabel = labelUpdate(Me.lblOwner.Caption)
'
'Reset_Labels
'Me.lblOwner.Caption = newLabel
'
'Me.OrderBy = "owner " & direction(newLabel)
'
'Exit Sub
'Err_Handler:
'    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
'End Sub
'
'Private Sub lblPriority_Click()
'On Error GoTo Err_Handler
'
'Dim newLabel As String
'newLabel = labelUpdate(Me.lblPriority.Caption)
'
'Reset_Labels
'Me.lblPriority.Caption = newLabel
'
'If InStr(newLabel, ">") <> 0 Then
'    Me.OrderBy = "IIf([priority]='Urgent',1,IIf([priority]='High',2,IIf([priority]='Medium',3,IIf([priority]='Low',4,5)))) DESC"
'Else
'    Me.OrderBy = "IIf([priority]='Urgent',1,IIf([priority]='High',2,IIf([priority]='Medium',3,IIf([priority]='Low',4,5))))"
'End If
'
'Exit Sub
'Err_Handler:
'    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
'End Sub
'
'Private Sub lblLocation_Click()
'On Error GoTo Err_Handler
'
'Dim newLabel As String
'newLabel = labelUpdate(Me.lblLocation.Caption)
'
'Reset_Labels
'Me.lblLocation.Caption = newLabel
'
'Me.OrderBy = "location " & direction(newLabel)
'
'Exit Sub
'Err_Handler:
'    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
'End Sub
'
'Private Sub lblDaysOpen_Click()
'On Error GoTo Err_Handler
'
'Dim newLabel As String
'newLabel = labelUpdate(Me.lblDaysOpen.Caption)
'
'Reset_Labels
'Me.lblDaysOpen.Caption = newLabel
'
'Me.OrderBy = "dateCreated " & direction(newLabel)
'
'Exit Sub
'Err_Handler:
'    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
'End Sub
'
'Private Sub lblLastModified_Click()
'On Error GoTo Err_Handler
'
'Dim newLabel As String
'newLabel = labelUpdate(Me.lblLastModified.Caption)
'
'Reset_Labels
'Me.lblLastModified.Caption = newLabel
'
'Me.OrderBy = "lastModified " & direction(newLabel)
'
'Exit Sub
'Err_Handler:
'    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
'End Sub
'
'Sub Reset_Labels()
'On Error GoTo Err_Handler
'
'Dim ctrl As Control
'
'For Each ctrl In Me.Controls
'    If TypeOf ctrl Is label Then
'        ctrl.Caption = Replace(ctrl.Caption, ">", "-")
'        ctrl.Caption = Replace(ctrl.Caption, "<", "-")
'    End If
'Next ctrl
'
'Exit Sub
'Err_Handler:
'    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
'End Sub
'
'Private Function labelUpdate(oldLabel As String)
'On Error GoTo Err_Handler
'Select Case True
'    Case InStr(oldLabel, "-") <> 0
'        labelUpdate = Replace(oldLabel, "-", ">")
'    Case InStr(oldLabel, ">") <> 0
'        labelUpdate = Replace(oldLabel, ">", "<")
'    Case InStr(oldLabel, "<") <> 0
'        labelUpdate = Replace(oldLabel, "<", ">")
'End Select
'Exit Function
'Err_Handler:
'    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
'End Function
'
'Private Function direction(label As String)
'On Error GoTo Err_Handler
'If InStr(label, ">") <> 0 Then
'    direction = "DESC"
'Else
'    direction = ""
'End If
'Exit Function
'Err_Handler:
'    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
'End Function
'
'Private Sub cboUserName_AfterUpdate()
'On Error GoTo Err_Handler
'Filter_Form
'Exit Sub
'Err_Handler:
'    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
'End Sub
'
'Private Sub Form_Load()
'On Error GoTo Err_Handler
'Dim owner As String
'
'Call setTheme(Me)
'
'Reset_Labels
'Reset_Filters
'
'If DCount("[id]", "tblCPC_Projects", "[userName] = '" & Environ("username") & "'") > 0 Then
'    Me.cboUserName = Environ("username")
'Else
'    Me.cboUserName = ""
'End If
'
'Me.cboYear.RowSource = "SELECT DISTINCT Year(dateCreated) as Year from tblCPC_Projects"
'
'Filter_Form
'
'If Me.Form.OrderBy <> "projectNumber DESC" Then
'    Me.Form.OrderBy = "projectNumber DESC"
'End If
'
'Me.Form.FilterOn = True
'Me.Form.OrderByOn = True
'
'Me.lblTotal.Caption = "Total" & vbCrLf & Me.Form.Recordset.RecordCount
'
'Exit Sub
'Err_Handler:
'    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
'End Sub
'
'Private Sub btnNewProject_Click()
'On Error GoTo Err_Handler
'
'DoCmd.OpenForm "frmCPC_NewProject"
'
'Exit Sub
'Err_Handler:
'    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
'End Sub
'
'Private Sub btnRefresh_Click()
'On Error GoTo Err_Handler
'Dim owner As String
'
'Reset_Labels
'Reset_Filters
'
'If DCount("[id]", "tblCPC_Projects", "[userName] = '" & Environ("username") & "'") > 0 Then
'    Me.cboUserName = Environ("username")
'Else
'    Me.cboUserName = ""
'End If
'
'Filter_Form
'
'If Me.Form.OrderBy <> "projectNumber DESC" Then Me.Form.OrderBy = "projectNumber DESC"
'
'Me.Requery
'
'Me.lblTotal.Caption = "Total" & vbCrLf & Me.Form.Recordset.RecordCount
'
'Exit Sub
'Err_Handler:
'    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
'End Sub
'
'Private Sub cboStatus_AfterUpdate()
'On Error GoTo Err_Handler
'Filter_Form
'Exit Sub
'Err_Handler:
'    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
'End Sub
'
'Private Sub cboYear_AfterUpdate()
'On Error GoTo Err_Handler
'Filter_Form
'Exit Sub
'Err_Handler:
'    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
'End Sub
'
'Private Sub cboType_AfterUpdate()
'On Error GoTo Err_Handler
'Filter_Form
'Exit Sub
'Err_Handler:
'    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
'End Sub
'
'Private Sub cboLocation_AfterUpdate()
'On Error GoTo Err_Handler
'Filter_Form
'Exit Sub
'Err_Handler:
'    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
'End Sub
'
'Private Sub Filter_Form()
'On Error GoTo Err_Handler
'
'Dim filter As String
'
'If IsNull(Me.cboStatus) Then
'    filter = "[status] <> 'Deleted'"
'Else
'    filter = "[status]='" & Me.cboStatus & "'"
'End If
'
'If Not IsNull(Me.cboYear) Then filter = filter & " AND Year([dateCreated])='" & Me.cboYear & "'"
'If Not IsNull(Me.cboType) Then filter = filter & " AND [projectType]='" & Me.cboType & "'"
'If Not IsNull(Me.cboLocation) Then filter = filter & " AND [location]='" & Me.cboLocation & "'"
'If Not IsNull(Me.cboUserName) Then filter = filter & " AND [userName]='" & Me.cboUserName & "'"
'If Not IsNull(Me.cboPartNumber) Then filter = filter & " AND [projectNumber] IN (SELECT [projectNumber] FROM tblCPC_Parts WHERE [partNumber] = '" & Me.cboPartNumber & "')"
'
'Me.Form.filter = filter
'Me.Form.FilterOn = True
'
'Me.lblTotal.Caption = "Total" & vbCrLf & Me.Form.Recordset.RecordCount
'
'Exit Sub
'Err_Handler:
'    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
'End Sub
'
'Private Sub Reset_Filters()
'On Error GoTo Err_Handler
'Dim ctrl As Control
'
'For Each ctrl In Me.Controls
'    If TypeOf ctrl Is ComboBox Then
'        If ctrl.name <> "cboStatus" And ctrl <> "All" Then
'            ctrl = ""
'        ElseIf ctrl.name = "cboStatus" And ctrl <> "Open" Then
'            ctrl = "Open"
'        End If
'    End If
'Next ctrl
'
'Exit Sub
'Err_Handler:
'    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
'End Sub
