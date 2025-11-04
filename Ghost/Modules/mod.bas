Option Compare Database

Public Function emailUsers()

Dim strTo As String

    Dim db As Database
    Dim rs1 As Recordset
    Set db = CurrentDb()
    Set rs1 = db.OpenRecordset("openSessions", dbOpenSnapshot)
    strTo = ""

    Dim lngCnt As Long
    lngCnt = 0
    Do While Not rs1.EOF
        strTo = strTo & getEmail(rs1![User]) & "; "
        lngCnt = lngCnt + 1
        rs1.MoveNext
    Loop

    rs1.Close
    Set rs1 = Nothing
    Set db = Nothing
    
    
    Dim objEmail As Object

Set objEmail = CreateObject("outlook.Application")
Set objEmail = objEmail.CreateItem(0)

With objEmail
    .To = ""
    .CC = ""
    .BCC = strTo
    .subject = "Working DB"
    .htmlBody = body
    .display
End With

Set objEmail = Nothing

End Function

Function getEmail(userName As String) As String

getEmail = ""
Dim db As Database
Set db = OpenDatabase("\\data\mdbdata\WorkingDB\build\Code_Review\WorkingDB_Connection.accdb")
Dim rsPermissions As Recordset
Set rsPermissions = db.OpenRecordset("SELECT * from tblPermissions WHERE user = '" & userName & "'")
getEmail = Nz(rsPermissions!userEmail, "")
rsPermissions.Close
Set rsPermissions = Nothing

GoTo exitFunc

exitFunc:
db.Close
Set db = Nothing
End Function