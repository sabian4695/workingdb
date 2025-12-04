On Error Resume Next
dim wdb
dim location
Dim fso

location = "C:\workingdb\WorkingDB"
'location = "H:\dev\repo\Front_End\WorkingDB_FE"

Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FileExists(location & ".laccdb") Then 'if the file exists, try to delete it.
	fso.Deletefile location & ".laccdb" 'if it does not let you delete it, that means the database is active and in use.
	If Err.Number <> 0 Then
		Set wdb = GetObject(location & ".accdb")
		wdb.Run("openNotificationFromEmail")
		Set wdb = Nothing
	end if
End If

Set fso = Nothing