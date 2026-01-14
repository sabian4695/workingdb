Option Compare Database
Option Explicit

Function doStuffFiles()

Dim folderName As String
Dim fso As Object
Dim folder As Object
Dim file As Object

Dim bit As String
bit = "64"

folderName = "\\data\mdbdata\WorkingDB\Pictures\Core\" & bit & "\"
Set fso = CreateObject("Scripting.FileSystemObject")
Set folder = fso.GetFolder(folderName)

Dim newFold As String
newFold = "\\data\mdbdata\WorkingDB\Pictures\SVG_theme_light\"

On Error GoTo checkThis

For Each file In folder.Files
    Dim FileName As String, newFile As String
    FileName = Replace(file.name, ".ico", ".svg")
    newFile = FileName
    
    If InStr(FileName, "_" & bit & "px") Then
        FileName = Replace(newFile, "_" & bit & "px", "")
    End If
    
    Call fso.CopyFile("\\data\mdbdata\WorkingDB\Pictures\SVG_theme_light\" & FileName, "\\data\mdbdata\WorkingDB\Pictures\SVG_theme_light\" & bit & "\" & newFile)
    
    GoTo skipCheck
checkThis:
    Debug.Print file.name
    Err.clear
skipCheck:
Next

Set fso = Nothing
Set folder = Nothing
    
End Function

Function doStuffApprovals()

Dim db As Database
Set db = CurrentDb()

Dim rs1 As Recordset
Set rs1 = db.OpenRecordset("SELECT * FROM tblPartSteps WHERE stepType = 'Upload'")

Dim rsApprovals As Recordset

Do While Not rs1.EOF
    Set rsApprovals = db.OpenRecordset("SELECT * FROM tblPartTrackingApprovals WHERE approvedOn is null AND tableRecordId = " & rs1!recordId)
    
    Do While Not rsApprovals.EOF
        rsApprovals.Delete
        rsApprovals.MoveNext
    Loop
    
    rsApprovals.CLOSE
    Set rsApprovals = Nothing
    rs1.MoveNext
Loop

rs1.CLOSE
Set rs1 = Nothing
Set db = Nothing

End Function

Public Function addPostgreSQL_DATA()

Dim db As Database
Set db = CurrentDb()

Dim tbl As TableDef

'For Each tbl In db.TableDefs
'    If InStr(tbl.name, "_old") Then
'        Debug.Print tbl.name & " START"
'
'        db.Execute "DELETE * FROM " & Replace(tbl.name, "_old", "")
'        DoEvents
'
'        moveRecords (Replace(tbl.name, "_old", ""))
'    End If
'Next tbl

Dim tblName As String
tblName = "tblpartsteps"

'db.Execute "DELETE * FROM " & tblName
'DoEvents

'moveRecords (tblName)

End Function

Public Function setItUp(tableName As String)

Dim serverCon As String, tableNameTo As String

Dim dbName As String
Dim serverName As String
Dim Uid As String
Dim Pwd As String
Dim schemaName As String

schemaName = "design."
dbName = "dm2"
serverName = "dw1v2-cluster.cluster-ro-c1aekkohw3x2.us-west-2.rds.amazonaws.com"
Uid = "uDesign"
'Uid = "npostgres"
Pwd = "zNQG6230^b7-"
'Pwd = "Khkdbh!01"

serverCon = "DATABASE=" & dbName & ";SERVER=" & serverName & ";PORT=5432;Uid=" & Uid & ";Pwd=" & Pwd & ";"

tableName = schemaName & LCase(tableName)
tableNameTo = tableName

Call Link_ODBCTbl(serverCon, tableName, tableNameTo, CurrentDb())

'Call moveRecords(tableNameTo)

End Function

Public Sub Link_ODBCTbl(serverConn As String, rstrTblSrc As String, rstrTblDest As String, db As DAO.Database)

'on error goto err_handler

    Dim tdf As TableDef
    Dim connOptions As String
    Dim myConn As String
    Dim myLen As Integer
    Dim bNoErr As Boolean

    bNoErr = True

    Set tdf = db.CreateTableDef(rstrTblDest)

' ***WORKAROUND*** Tested Access 2000 on Win2k, PostgreSQL 7.1.3 on Red Hat 7.2
'
'
'   PG_ODBC_PARAMETER           ACCESS_PARAMETER
'   *********************************************
'   READONLY                    A0
'   PROTOCOL                    A1
'   FAKEOIDINDEX                A2  'A2 must be 0 unless A3=1
'   SHOWOIDCOLUMN               A3
'   ROWVERSIONING               A4
'   SHOWSYSTEMTABLES            A5
'   CONNSETTINGS                A6
'   FETCH                       A7
'   SOCKET                      A8
'   UNKNOWNSIZES                A9  ' range [0-2]
'   MAXVARCHARSIZE              B0
'   MAXLONGVARCHARSIZE          B1
'   DEBUG                       B2
'   COMMLOG                     B3
'   OPTIMIZER                   B4  ' note that 1 = _cancel_ generic optimizer...
'   KSQO                        B5
'   USEDECLAREFETCH             B6
'   TEXTASLONGVARCHAR           B7
'   UNKNOWNSASLONGVARCHAR       B8
'   BOOLSASCHAR                 B9
'   PARSE                       C0
'   CANCELASFREESTMT            C1
'   EXTRASYSTABLEPREFIXES       C2

'myConn = "ODBC;DRIVER={PostgreSQL35W};" & serverConn & _
            "A0=0;A1=6.4;A2=0;A3=0;A4=0;A5=0;A6=;A7=100;A8=4096;A9=0;" & _
            "B0=254;B1=8190;B2=0;B3=0;B4=1;B5=1;B6=0;B7=1;B8=0;B9=1;" & _
            "C0=0;C1=0;C2=dd_"
            
myConn = "ODBC;DRIVER={PostgreSQL Unicode};" & serverConn & _
            "CA=d;A7=100;B0=255;B1=8190;BI=0;C2=;D6=-101;CX=1c305008b;A1=7.4"

Debug.Print myConn
'Exit Sub
    tdf.Connect = myConn
    tdf.SourceTableName = rstrTblSrc
    db.TableDefs.Append tdf
    db.TableDefs.refresh

    ' If we made it this far without errors, table was linked...
    If bNoErr Then
        MsgBox "Form_Login.Link_ODBCTbl: Linked new relation: " & _
                 rstrTblSrc
    End If

    'Debug.Print "Linked new relation: " & rstrTblSrc ' Link new relation

    Set tdf = Nothing

Exit Sub

Err_Handler:
    bNoErr = False
    Debug.Print Err.number & " : " & Err.DESCRIPTION
    If Err.number <> 0 Then MsgBox Err.number, Err.DESCRIPTION, "TEST" & _
                                     ": Form_Login.Link_ODBCTbl"
    Resume Next

End Sub

Public Sub UnLink_ODBCTbl(rstrTblName As String, db As DAO.Database)

MsgBox "Entering " & "TEST" & ": Form_Login.UnLink_ODBCTbbl"

On Error GoTo Err_Handler

    db.TableDefs.Delete rstrTblName
    db.TableDefs.refresh

    Debug.Print "Removed revoked relation: " & rstrTblName

Exit Sub

Err_Handler:
    Debug.Print Err.number & " : " & Err.DESCRIPTION
    If Err.number <> 0 Then MsgBox Err.number, Err.DESCRIPTION, "TEST" & _
                                     ": Form_Login.UnLink_ODBCTbl"
    Resume Next

End Sub


Function moveRecords(tableName As String)

Dim db As Database
Set db = CurrentDb()

Dim rs As Recordset, rsOld As Recordset

Set rs = db.OpenRecordset(tableName)
Set rsOld = db.OpenRecordset(tableName & "_old")

Dim fld As DAO.Field

Do While Not rsOld.EOF
    
    rs.addNew
    
    For Each fld In rsOld.Fields
        Select Case LCase(fld.name)
            Case "recordid"
                'Debug.Print fld.Value
                 rs("recordid") = rsOld(fld.name).Value
            Case "id"
                'Debug.Print fld.Value
                 rs("recordid") = rsOld(fld.name).Value
            Case "quantity", "pieceweight", "matnum1pieceweight", "outsourcecost"
                rs(fld.name) = Round(rsOld(fld.name).Value, 5)
            Case "3Dweight"
                    rs("_" & fld.name) = rsOld(fld.name).Value
           Case "shotsperhour", "piecesperhour"
                rs(fld.name) = Round(rsOld(fld.name).Value, 2)
            Case "toolnumber"
                If IsNull(fld.Value) Then GoTo nextOne
            Case "length"
                rs("partlength") = rsOld(fld.name).Value
            Case "width"
                rs("partwidth") = rsOld(fld.name).Value
            Case "packagingteststatus", "NPIFstatus", "customerapprovalstatus", "tooltips", "training_mode", "catiacustomcolor", "designwoid", "autoposition", "notifications"
                GoTo nextField
            Case "user"
                rs("username") = rsOld(fld.name).Value
            Case "edit", "beta", "admin", "developer"
                rs(fld.name & "permission") = rsOld(fld.name).Value
            Case "dept"
                rs("department") = rsOld(fld.name).Value
            Case "level"
                rs("englevel") = rsOld(fld.name).Value
            Case "designwopermissions"
                rs("designwopermission") = rsOld(fld.name).Value
            Case "type"
                rs("unittype") = rsOld(fld.name).Value
            Case Else
                rs(fld.name) = rsOld(fld.name).Value
        End Select
nextField:
    Next

    rs.Update
nextOne:
    rsOld.MoveNext
Loop

Set db = Nothing
End Function

Public Function FieldExists(ByVal rs As Recordset, ByVal fieldName As String) As Boolean
    On Error GoTo merr

    FieldExists = rs.Fields(fieldName).name <> ""
    Exit Function

merr:
    FieldExists = False
End Function

Function doToAllFormsAndThings()

Dim obj As AccessObject, dbs As Object, testString As String
Dim ctl As Control

Dim frm As Form

Set dbs = Application.CurrentProject
' Search for open AccessObject objects in AllForms collection.

For Each obj In dbs.AllForms
    DoCmd.OpenForm obj.name, acDesign
    
    Set frm = forms(obj.name)
    
    testString = frm.RecordSource
    
    If InStr(frm.RecordSource, "tblDropDowns.ID") Then
'        frm.RecordSource = Replace(Forms(obj.name).RecordSource, "tblDropDownsSP.ID", "tblDropDownsSP.recordid")
'        frm.RecordSource = Replace(Forms(obj.name).RecordSource, "tblDropDownsSP_1.ID", "tblDropDownsSP_1.recordid")
'        frm.RecordSource = Replace(Forms(obj.name).RecordSource, "tblDropDownsSP_2.ID", "tblDropDownsSP_2.recordid")
        
        Debug.Print testString
        testString = Replace(testString, "tblDropDowns", "tblDropDownsSP")
        testString = Replace(testString, "tblDropDownsSP.ID", "tblDropDownsSP.recordid")
        testString = Replace(testString, "DRSadjustReasons", "drs_adjustreasons")
        testString = Replace(testString, "DRSapprovalStatus", "drs_status")
        testString = Replace(testString, "DRSdesignResponsibility", "drs_designresponsibility")
        testString = Replace(testString, "DFMEA", "drs_dfmea")
        testString = Replace(testString, "DRSdrLevels", "drs_drlevel")
        testString = Replace(testString, "DRSpartComplexity", "drs_partcomplexity")
        testString = Replace(testString, "DRStoolingDept", "drs_toolingdept")
        testString = Replace(testString, "DRSunit12Location", "drs_unit12location")
        testString = Replace(testString, "DRSworkTypes", "drs_worktype")
        testString = Replace(testString, "DRSeta", "drs_eta")
        testString = Replace(testString, "DRSpermissionLevels", "drs_permissionlevel")
        testString = Replace(testString, "DRSpartComplexitySort", "drs_partcomplexitysort")
        testString = Replace(testString, "tblDRStrackerExtras1", "drs_status")
        testString = Replace(testString, "DRStype", "drs_type")
        Debug.Print testString
        
        frm.RecordSource = testString
    End If
    
'    For Each ctl In Forms(obj.name).Controls
'        If ctl.ControlType = acComboBox Then
'            If InStr(ctl.RowSource, "tblDropDowns.") Then
'                Debug.Print ctl.RowSource
'                ctl.RowSource = Replace(ctl.RowSource, "tblDropDowns", "tblDropDownsSP")
'                ctl.RowSource = Replace(ctl.RowSource, "tblDropDownsSP.ID", "tblDropDownsSP.recordid")
'                ctl.RowSource = Replace(ctl.RowSource, "DRSadjustReasons", "drs_adjustreasons")
'                ctl.RowSource = Replace(ctl.RowSource, "DRSapprovalStatus", "drs_status")
'                ctl.RowSource = Replace(ctl.RowSource, "DRSdesignResponsibility", "drs_designresponsibility")
'                ctl.RowSource = Replace(ctl.RowSource, "DFMEA", "drs_dfmea")
'                ctl.RowSource = Replace(ctl.RowSource, "DRSdrLevels", "drs_drlevel")
'                ctl.RowSource = Replace(ctl.RowSource, "DRSpartComplexity", "drs_partcomplexity")
'                ctl.RowSource = Replace(ctl.RowSource, "DRStoolingDept", "drs_toolingdept")
'                ctl.RowSource = Replace(ctl.RowSource, "DRSunit12Location", "drs_unit12location")
'                ctl.RowSource = Replace(ctl.RowSource, "DRSworkTypes", "drs_worktype")
'                ctl.RowSource = Replace(ctl.RowSource, "DRSeta", "drs_eta")
'                ctl.RowSource = Replace(ctl.RowSource, "DRSpermissionLevels", "drs_permissionlevel")
'                ctl.RowSource = Replace(ctl.RowSource, "DRSpartComplexitySort", "drs_partcomplexitysort")
'                ctl.RowSource = Replace(ctl.RowSource, "tblDRStrackerExtras1", "drs_status")
'                ctl.RowSource = Replace(ctl.RowSource, "DRStype", "drs_type")
'                Debug.Print ctl.RowSource
'            End If
'        End If
'    Next
    
    Set frm = Nothing
    DoCmd.CLOSE acForm, obj.name, acSaveYes
nextOne:
Next obj


End Function