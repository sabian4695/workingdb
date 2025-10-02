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

Function doStuff()

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

Public Function setItUp()

Dim serverCon As String, tableName As String, tableNameTo As String

Dim dbName As String
Dim serverName As String
Dim Uid As String
Dim Pwd As String

dbName = "dm2"
serverName = "dw1v2-cluster.cluster-ro-c1aekkohw3x2.us-west-2.rds.amazonaws.com"
Uid = "uDesign"
'Uid = "npostgres"
Pwd = "zNQG6230^b7-"
'Pwd = "Khkdbh!01"

serverCon = "DATABASE=" & dbName & ";SERVER=" & serverName & ";PORT=5432;Uid=" & Uid & ";Pwd=" & Pwd & ";"

tableName = "design.tbltasks"
tableNameTo = "tblTasks"

Call Link_ODBCTbl(serverCon, tableName, tableNameTo, CurrentDb())

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


Function moveRecords()

Dim db As Database
Set db = CurrentDb()

Dim rs As Recordset, rsOld As Recordset
Dim tableName As String

tableName = "tblTasks"

Set rs = db.OpenRecordset(tableName)
Set rsOld = db.OpenRecordset(tableName & "_old")

Dim fld As DAO.Field

Do While Not rsOld.EOF
    Debug.Print rsOld!recordId
    rs.addNew
    
    For Each fld In rsOld.Fields
        rs(fld.name) = rsOld(fld.name).Value
    Next
    
    rs.Update
    rsOld.MoveNext
Loop

Set db = Nothing
End Function