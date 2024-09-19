Option Compare Database
Option Explicit

Function createDnumber() As String
On Error GoTo err_handler

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

Exit Function
err_handler:
    Call handleError("wdbDesignE", "createDnumber", Err.DESCRIPTION, Err.number)
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
MsgBox "Error " & Err.number & " (" & Err.DESCRIPTION & _
") in procedure SetNavButtons of Module modFormOperations"
GoTo SetNavButtons_Exit
End Sub

Public Sub registerDRSUpdates(table As String, ID As Variant, column As String, oldVal As Variant, newVal As Variant, Optional tag0 As String, Optional tag1 As String)
On Error GoTo err_handler

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

If Len(oldVal) > 255 Then oldVal = Left(oldVal, 255)
If Len(newVal) > 255 Then newVal = Left(newVal, 255)

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

Exit Sub
err_handler:
    Call handleError("wdbDesignE", "registerDRSUpdates", Err.DESCRIPTION, Err.number)
End Sub

Function DRShistoryGrabReference(columnName As String, inputVal As Variant) As String

DRShistoryGrabReference = inputVal

On Error GoTo exitFunc
inputVal = CDbl(inputVal)

Dim lookup As String

Select Case columnName
    Case "Request_Type", "cboRequestType"
        lookup = "DRStype"
    Case "DR_Level"
        lookup = "DRSdrLevels"
    Case "Design_Responsibility", "cboDesignResponsibility"
        lookup = "DRSdesignResponsibility"
    Case "Part_Complexity", "cboComplexity"
        lookup = "DRSpartComplexity"
    Case "DRS_Location"
        lookup = "DRSdesignGroup"
    Case "Assignee", "cboAssignee"
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
    Case "Adjusted_Reason", "cboAdjustedReason"
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
On Error GoTo err_handler

Dim total
Dim checked

total = DCount("[Task_ID]", "[tblTaskTracker]", "[Control_Number] = " & controlNum)
checked = DCount("[Task_ID]", "[tblTaskTracker]", "[Control_Number] = " & controlNum & "AND [cbClosed] = TRUE")

progressPercent = checked / total

Exit Function
err_handler:
    Call handleError("wdbDesignE", "progressPercent", Err.DESCRIPTION, Err.number)
End Function