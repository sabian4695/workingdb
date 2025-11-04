Option Compare Database
Option Explicit

Public Sub registerDRSUpdates(table As String, ID As Variant, column As String, oldVal As Variant, newVal As Variant, Optional tag0 As String, Optional tag1 As String)
On Error GoTo Err_Handler

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

Dim db As Database
Set db = CurrentDb()
db.Execute "INSERT INTO tblDRSUpdateTracking" & sqlColumns & sqlValues
Set db = Nothing

Exit Sub
Err_Handler:
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