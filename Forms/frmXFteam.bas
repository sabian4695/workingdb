Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Load()
On Error Resume Next
Dim X
Dim TE As String
Dim rs As DAO.Recordset, rsEmployee As Recordset, rsFiltered As Recordset
Dim db As DAO.Database
X = Form_DASHBOARD.partNumberSearch

Call setTheme(Me)

Dim projId
projId = Nz(DLookup("projectId", "tblPartProjectPartNumbers", "childPartNumber = '" & X & "'"))

If projId <> "" Then
    X = DLookup("partNumber", "tblPartProject", "recordId = " & projId)
End If

Me.partNumber = X

Me.sfrmPartTeam.Form.filter = "partNumber = '" & X & "'"
Me.sfrmPartTeam.Form.FilterOn = True

Me.filter = "primaryPN = '" & X & "' OR relatedPN = '" & X & "'"
Me.FilterOn = True

Me.currentPN = X

Set db = CurrentDb

If X Like "D*" Then
    Me.toolingEngineer = "Not Found"
    Exit Sub
End If

TE = Nz(DLookup("[Tooling Engineer]", "tblToolsMain", "[Tool Number] = '" & X & "'"))
If IsNull(TE) Or TE = "" Then
    Me.lblDB.Caption = "CNL Project DB"
    Set rs = db.OpenRecordset("SELECT Design_Engineer_ID, Project_Engineer_ID, Account_Manager_ID, Manufacturing_Engineer_ID, Quality_Engineer_ID, Tooling_Engineer_ID FROM [Main Table1] WHERE [Part Number] = '" & X & "'", dbOpenSnapshot)
    If rs.RecordCount = 0 Then Exit Sub
    Set rsEmployee = db.OpenRecordset("SELECT EMPLOYEE, PERSON_ID from APPS_XXCUS_USER_EMPLOYEES_V")
    
    rsEmployee.filter = "Person_ID = " & rs![Project_Engineer_ID]
    Set rsFiltered = rsEmployee.OpenRecordset
    Me.projectEngineer = StrConv(rsFiltered![Employee], vbProperCase)
    
    rsEmployee.filter = "Person_ID = " & rs![Account_Manager_ID]
    Set rsFiltered = rsEmployee.OpenRecordset
    Me.sales = StrConv(rsFiltered![Employee], vbProperCase)
    
    rsEmployee.filter = "Person_ID = " & rs![Manufacturing_Engineer_ID]
    Set rsFiltered = rsEmployee.OpenRecordset
    Me.mfgEngineer = StrConv(rsFiltered![Employee], vbProperCase)
    
    rsEmployee.filter = "Person_ID = " & rs![Quality_Engineer_ID]
    Set rsFiltered = rsEmployee.OpenRecordset
    Me.qualityEngineer = StrConv(rsFiltered![Employee], vbProperCase)
    
    rsEmployee.filter = "Person_ID = " & rs![Tooling_Engineer_ID]
    Set rsFiltered = rsEmployee.OpenRecordset
    Me.toolingEngineer = StrConv(rsFiltered![Employee], vbProperCase)
    
    rsEmployee.filter = "Person_ID = " & rs![Design_Engineer_ID]
    Set rsFiltered = rsEmployee.OpenRecordset
    Me.designEngineer = StrConv(rsFiltered![Employee], vbProperCase)
Else
    Me.lblDB.Caption = "SLB Tooling DB"
    Me.toolingEngineer = StrConv(TE, vbProperCase)
    Me.qualityEngineer = StrConv(DLookup("[Quality Engineer]", "tblToolsMain", "[Tool Number] = '" & X & "'"), vbProperCase)
    Me.projectEngineer = StrConv(DLookup("[Project Engineer]", "tblToolsMain", "[Tool Number] = '" & X & "'"), vbProperCase)
    Me.designEngineer = StrConv(DLookup("[DevDesignEng]", "tblToolsMain", "[Tool Number] = '" & X & "'"), vbProperCase)
    Me.mfgEngineer = ""
    Me.sales = StrConv(DLookup("[Sales]", "tblToolsMain", "[Tool Number] = '" & X & "'"), vbProperCase)
End If

On Error Resume Next
rsEmployee.CLOSE: Set rsEmployee = Nothing
rs.CLOSE: Set rs = Nothing
rsFiltered.CLOSE: Set rsFiltered = Nothing
Set db = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub partNumber_Click()
On Error GoTo Err_Handler

Me.sfrmPartTeam.Form.filter = "partNumber = '" & Me.partNumber & "'"
Me.sfrmPartTeam.Form.FilterOn = True
Me.currentPN = Me.partNumber

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub primaryPN_Click()
On Error GoTo Err_Handler

Me.sfrmPartTeam.Form.filter = "partNumber = '" & Me.primaryPN & "'"
Me.sfrmPartTeam.Form.FilterOn = True
Me.currentPN = Me.primaryPN

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub relatedPN_Click()
On Error GoTo Err_Handler

Me.sfrmPartTeam.Form.filter = "partNumber = '" & Me.relatedPN & "'"
Me.sfrmPartTeam.Form.FilterOn = True
Me.currentPN = Me.relatedPN

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
