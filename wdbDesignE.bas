Option Compare Database
Option Explicit

Function createDnumber() As String
Dim rs1 As Recordset
Dim strInsert
Set rs1 = CurrentDb().OpenRecordset("tblDnumbers", dbOpenSnapshot)

Dim dNum

rs1.FindFirst "dNumber = 9999"
If rs1.NoMatch Then
    rs1.filter = "dNumber < 10000"
    Set rs1 = rs1.OpenRecordset
End If

rs1.Sort = "dNumber"
Set rs1 = rs1.OpenRecordset
rs1.MoveLast
dNum = rs1!dNumber + 1

strInsert = "INSERT INTO tblDnumbers(dNumber,createdBy,createdDate) VALUES (" & dNum & ",'" & Environ("username") & "','" & Now() & "')"
CurrentDb().Execute strInsert, dbFailOnError

createDnumber = "D" & dNum

rs1.Close
Set rs1 = Nothing
End Function

Sub SetNavButtons(ByRef frmSomeForm As Form)
On Error GoTo SetNavButtons_Error
'-- enable/disable buttons depending on record position
With frmSomeForm
    If .Recordset.RecordCount <= 1 Or .CurrentRecord > .Recordset.RecordCount Then
       .cmdFirst.Enabled = True
       .cmdPrevious.Enabled = True
       .cmdPrevious.SetFocus
       .cmdNext.Enabled = False
       .cmdLast.Enabled = False
   ElseIf .CurrentRecord = 1 Then
       .cmdNext.Enabled = True
       .cmdLast.Enabled = True
       .cmdNext.SetFocus
       .cmdFirst.Enabled = False
       .cmdPrevious.Enabled = False
    Else
       .cmdFirst.Enabled = True
       .cmdPrevious.Enabled = True
       .cmdNext.Enabled = True
       .cmdLast.Enabled = True
    End If
End With
SetNavButtons_Exit:
On Error Resume Next
Exit Sub
SetNavButtons_Error:
MsgBox "Error " & Err.number & " (" & Err.description & _
") in procedure SetNavButtons of Module modFormOperations"
GoTo SetNavButtons_Exit
End Sub

Function getCheckFolder(controlNum As Long)
Dim chkFold
chkFold = DLookup("[Check_Folder]", "tblDRStrackerExtras", "[Control_Number] = " & controlNum)

If IsNull(chkFold) Then
    If MsgBox("No check folder found yet. Would you like to add one?", vbYesNo, "No Folder Found") = vbYes Then
        Dim x
        x = InputBox("Paste Link to Check Folder Here", "Add Check Folder Link")
        Select Case x
            Case vbCancel
                Exit Function
            Case ""
                Exit Function
        End Select
        x = Replace(x, "'", "''")
        Call registerDRSUpdates("tblDRStrackerExtras", controlNum, "Check_Folder", "", x)
        
        CurrentDb().Execute ("UPDATE tblDRStrackerExtras SET [tblDRStrackerExtras].[Check_Folder] = '" & x & "' WHERE [tblDRStrackerExtras].[Control_Number] = " & controlNum)
    Else
        Exit Function
    End If
End If
chkFold = DLookup("[Check_Folder]", "tblDRStrackerExtras", "[Control_Number] = " & controlNum)

If InStr(Left(chkFold, 10), "file") Then
    chkFold = Replace(chkFold, "%20", " ")
    chkFold = Right(Left(chkFold, Len(chkFold) - 1), Len(chkFold) - 10)
End If

If Not Right(chkFold, 1) = "\" Then
    chkFold = chkFold & "\"
End If

getCheckFolder = chkFold
End Function

Public Sub registerDRSUpdates(table As String, ID As Variant, column As String, oldVal As Variant, newVal As Variant, Optional tag0 As String, Optional tag1 As String)
Dim sqlColumns As String, sqlValues As String

If (VarType(oldVal) = vbDate) Then
    oldVal = Format(oldVal, "mm/dd/yyyy")
End If

If (VarType(newVal) = vbDate) Then
    newVal = Format(newVal, "mm/dd/yyyy")
End If

If (IsNull(oldVal)) Then
    oldVal = ""
End If

If (IsNull(newVal)) Then
    newVal = ""
End If

sqlColumns = "(tableName,tableRecordId,updatedBy,updatedDate,columnName,previousData,newData,dataTag0"
                    
sqlValues = " values ('" & table & "', '" & ID & "', '" & Environ("username") & "', '" & Now() & "', '" & column & "', '" & StrQuoteReplace(CStr(oldVal)) & "', '" & StrQuoteReplace(CStr(newVal)) & "','" & tag0 & "'"

If (IsNull(tag1)) Then
    sqlColumns = sqlColumns & ")"
    sqlValues = sqlValues & ");"
Else
    sqlColumns = sqlColumns & ",dataTag1)"
    sqlValues = sqlValues & ",'" & tag1 & "');"
End If


CurrentDb().Execute "INSERT INTO tblDRSUpdateTracking" & sqlColumns & sqlValues

End Sub

Function DRShistoryGrabReference(columnName As String, inputVal As Variant) As String

DRShistoryGrabReference = inputVal

On Error GoTo exitFunc
inputVal = CDbl(inputVal)

Dim lookup As String

Select Case columnName
    Case "Request_Type"
        lookup = "DRStype"
    Case "DR_Level"
        lookup = "DRSdrLevels"
    Case "Design_Responsibility"
        lookup = "DRSdesignResponsibility"
    Case "Part_Complexity"
        lookup = "DRSpartComplexity"
    Case "DRS_Location"
        lookup = "DRSdesignGroup"
    Case "Assignee"
        GoTo personLookup
    Case "cboChecker1"
        GoTo personLookup
    Case "cboChecker2"
        GoTo personLookup
    Case "Dev_Responsibility"
        GoTo personLookup
    Case "Project_Location"
        lookup = "DRSunit12Location"
    Case "Tooling_Department"
        lookup = "DRStoolingDept"
    Case "Customer"
        DRShistoryGrabReference = DLookup("[CUSTOMER_NAME]", "APPS_XXCUS_CUSTOMERS", "[CUSTOMER_ID] = " & inputVal)
    Case "Adjusted_Reason"
        lookup = "DRSadjustReasons"
    Case "Delay_Reason"
        lookup = "DRSadjustReasons"
    Case "cboApprovalStatus"
        lookup = "DRSapprovalStatus"
    Case "assigneeSign"
        GoTo trueFalse
    Case "checker1Sign"
        GoTo trueFalse
    Case "checker2Sign"
        GoTo trueFalse
    Case Else
        Exit Function
End Select

DRShistoryGrabReference = DLookup("[" & lookup & "]", "tblDropDowns", "ID = " & inputVal)

Exit Function
personLookup:
DRShistoryGrabReference = DLookup("[user]", "tblPermissions", "ID = " & inputVal)

Exit Function
trueFalse:
If (inputVal = 0) Then
    DRShistoryGrabReference = "False"
Else
    DRShistoryGrabReference = "True"
End If

exitFunc:
End Function

Function progressPercent(controlNum As Long)
Dim total
Dim checked

total = DCount("[Task_ID]", "[tblTaskTracker]", "[Control_Number] = " & controlNum)
checked = DCount("[Task_ID]", "[tblTaskTracker]", "[Control_Number] = " & controlNum & "AND [cbClosed] = TRUE")

progressPercent = checked / total
End Function