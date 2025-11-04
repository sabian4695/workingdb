Option Compare Database
Option Explicit

Public Function runAll()
On Error Resume Next

CurrentDb.Execute "INSERT INTO tblAnalytics (module,form,userName,dateUsed) VALUES ('refreshReports','Form_frmSplash','" & Environ("username") & "','" & Now() & "')"

Call refreshSP

Application.Quit

End Function

Public Function refreshSP()
On Error Resume Next

Dim db As Database
Set db = CurrentDb()

Dim tdf As DAO.TableDef

For Each tdf In db.TableDefs
    If Left(tdf.Name, 2) = "cb" Then
        Debug.Print "ex_" & tdf.Name
        db.Execute "DELETE * FROM " & tdf.Name
        db.Execute "INSERT INTO " & tdf.Name & " Select ex_" & tdf.Name & ".* From ex_" & tdf.Name
    End If
Next

Set db = Nothing

End Function