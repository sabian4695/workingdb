Option Compare Database
Option Explicit

Function populateDCR(partNumber As String, Optional changeType As String = ".", Optional specificECO As String = "") As Boolean
On Error GoTo err_handler

populateDCR = False

Dim db As Database
Set db = CurrentDb()

Dim docHis As String, partPath As String, dcrFold As String, dcrPath As String
docHis = addLastSlash(openDocumentHistoryFolder(partNumber, False))
partPath = docHis & "Misc\"
dcrFold = partPath & "DCR\"
dcrPath = dcrFold & partNumber & "_DCR.pptx"

'make sure DCR folder exists
If FolderExists(partPath) = False Then MkDir (partPath)
If FolderExists(dcrFold) = False Then MkDir (dcrFold)

'check for existing file. If exists, then kill this whole thing
Dim fso As Object
Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FileExists(dcrPath) Then Exit Function
Call fso.CopyFile(mainFolder("DCR"), dcrPath) 'copy template file

Dim ppt As New PowerPoint.Application
Dim pptPres As PowerPoint.Presentation
Dim curSlide As PowerPoint.Slide

ppt.Presentations.open dcrPath
Set pptPres = ppt.ActivePresentation
Set curSlide = pptPres.Slides(1)

Dim shp As PowerPoint.Shape

curSlide.Shapes.Range(Array("partNumber")).TextFrame.TextRange.Text = partNumber
curSlide.Shapes.Range(Array("purpose")).TextFrame.TextRange.Text = changeType
curSlide.Shapes.Range(Array("issueOpened")).TextFrame.TextRange.Text = Date
curSlide.Shapes.Range(Array("ECOimp")).TextFrame.TextRange.Text = "ASAP"
curSlide.Shapes("tblSignatures").table.Cell(2, 1).Shape.TextFrame.TextRange.Text = getFullName
curSlide.Shapes("tblSignatures").table.Cell(3, 1).Shape.TextFrame.TextRange.Text = Date

Dim ecoText As String
Dim rsECOs As Recordset
ecoText = specificECO

If specificECO = "" Then
    Set rsECOs = db.OpenRecordset("SELECT * FROM ENG_ENG_REVISED_ITEMS WHERE REVISED_ITEM_ID = " & idNAM(partNumber, "NAM"), dbOpenSnapshot)
    rsECOs.MoveLast
    ecoText = rsECOs!CHANGE_NOTICE
    rsECOs.Close
End If

curSlide.Shapes.Range(Array("ECO")).TextFrame.TextRange.Text = ecoText
curSlide.Shapes.Range(Array("partDesc")).TextFrame.TextRange.Text = findDescription(partNumber)

populateDCR = True

Set rsECOs = Nothing

Set db = Nothing
Set curSlide = Nothing
Set pptPres = Nothing
Set ppt = Nothing

Exit Function
err_handler:
    Call handleError("wdbDesignE", "populateDCR", Err.DESCRIPTION, Err.number)
End Function

Function populateETAs(issueDate As Date, dueDate As Date)
On Error GoTo err_handler

Dim db As Database
Set db = CurrentDb()
Dim rsWorkloadTbl As Recordset, rsWorkloadTbl1 As Recordset
Dim rsSessVar As Recordset, rsPerm As Recordset

Set rsPerm = db.OpenRecordset("SELECT * FROM tblPermissions WHERE designWOpermissions <> 3 AND Inactive = False")
Set rsSessVar = db.OpenRecordset("SELECT * FROM tblSessionVariables WHERE userName Is Not Null")

Do While Not rsSessVar.EOF
    rsSessVar.Delete
    rsSessVar.MoveNext
Loop

Do While Not rsPerm.EOF
    Set rsWorkloadTbl = db.OpenRecordset("SELECT Round(Sum([hours]),2) AS totalHours FROM tblWorkloadRanking WHERE " & _
        "userName = '" & rsPerm!User & "' AND hoursDate < #" & Date & "#")
    Set rsWorkloadTbl1 = db.OpenRecordset("SELECT Round(Sum([hours]),2) AS totalHours FROM tblWorkloadRanking WHERE " & _
        "userName = '" & rsPerm!User & "' AND hoursDate >= #" & issueDate & "# AND hoursDate <= #" & dueDate & "#")

    rsSessVar.addNew
    
    rsSessVar!userName = rsPerm!User
    rsSessVar!overdueETA = Nz(rsWorkloadTbl!totalHours, 0)
    rsSessVar!etaBetween = Nz(rsWorkloadTbl1!totalHours, 0)
    
    rsSessVar.Update
    rsPerm.MoveNext
Loop

rsSessVar.Close: Set rsSessVar = Nothing
rsPerm.Close: Set rsPerm = Nothing
rsWorkloadTbl.Close: Set rsWorkloadTbl = Nothing
rsWorkloadTbl1.Close: Set rsWorkloadTbl1 = Nothing
Set db = Nothing

Exit Function
err_handler:
    Call handleError("wdbDesignE", "populateETAs", Err.DESCRIPTION, Err.number)
End Function

Function populateWorkload() As Boolean
On Error GoTo err_handler

TempVars.Add "tStamp", Timer

Dim db As Database
Set db = CurrentDb()
Dim rsUsers As Recordset
Dim rsWO As Recordset
Dim rsWorkloadTbl As Recordset
Dim rsHolidays As Recordset

Dim timeSum As Double
Dim availabilityRank As Long
Dim overdue As Double
Dim hours7Days As Double
Dim hours30Days As Double
Dim dueDate As Date
Dim eta As Double
Dim woPerDay As Double
Dim forDate As Date

Set rsUsers = db.OpenRecordset("SELECT * FROM tblPermissions WHERE designWOpermissions <> 3 AND Inactive = False")
Set rsWorkloadTbl = db.OpenRecordset("tblWorkloadRanking")
Set rsHolidays = db.OpenRecordset("tblHolidays")

'Clear Table
If rsWorkloadTbl.RecordCount > 0 Then
    Do While Not rsWorkloadTbl.EOF
        rsWorkloadTbl.Delete
        rsWorkloadTbl.MoveNext
    Loop
End If

Do While Not rsUsers.EOF
    'select all open WOs for this assignee
    Set rsWO = db.OpenRecordset("SELECT * FROM dbo_tblDRS WHERE Assignee = " & rsUsers!ID & " AND Approval_Status = 2 AND Completed_Date Is Null")
    
    Do While Not rsWO.EOF
        dueDate = Nz(rsWO!Adjusted_Due_Date, rsWO!Due_Date)
        
        'calculate eta per day
        woPerDay = rsWO!Design_Level / (countWorkdays(rsWO!Issue_Date, dueDate) + 1) 'include due date in calc, so add 1
        
        'add that ETA per day into each record between issue and due dates
        For forDate = rsWO!Issue_Date To dueDate
            'Set rsHolidays = db.OpenRecordset("SELECT * FROM tblHolidays WHERE holidayDate = #" & forDate & "#")
            'If rsHolidays.RecordCount > 0 Then GoTo nextDate
            
            rsHolidays.FindFirst "holidayDate = #" & forDate & "#"
            If Not rsHolidays.noMatch Then GoTo nextDate
            
            If Weekday(forDate) = 7 Or Weekday(forDate) = 1 Then GoTo nextDate
            
            rsWorkloadTbl.addNew
            With rsWorkloadTbl
                !userName = rsUsers!User
                !hoursDate = forDate
                !hours = woPerDay
            End With
            rsWorkloadTbl.Update
nextDate:
        Next forDate
        
        rsWO.MoveNext
    Loop
    
    rsUsers.MoveNext
Loop

populateWorkload = True

rsWorkloadTbl.Close: Set rsWorkloadTbl = Nothing
rsHolidays.Close: Set rsHolidays = Nothing
rsUsers.Close: Set rsUsers = Nothing
rsWO.Close: Set rsWO = Nothing
Set db = Nothing

Exit Function
err_handler:
    Call handleError("wdbDesignE", "populateWorkload", Err.DESCRIPTION, Err.number)
End Function

Function createDnumber() As String
On Error GoTo err_handler

Dim db As Database
Set db = CurrentDb()
Dim rs1 As Recordset
Dim strInsert
Set rs1 = db.OpenRecordset("tblDnumbers", dbOpenSnapshot)

Dim dNum

rs1.FindFirst "dNumber = 9999"
If rs1.noMatch Then
    rs1.filter = "dNumber < 10000"
    Set rs1 = rs1.OpenRecordset
End If

rs1.Sort = "dNumber"
Set rs1 = rs1.OpenRecordset
rs1.MoveLast
dNum = rs1!dNumber + 1

strInsert = "INSERT INTO tblDnumbers(dNumber,createdBy,createdDate) VALUES (" & dNum & ",'" & Environ("username") & "','" & Now() & "')"
db.Execute strInsert, dbFailOnError

createDnumber = "D" & dNum

rs1.Close
Set rs1 = Nothing
Set db = Nothing

Exit Function
err_handler:
    Call handleError("wdbDesignE", "createDnumber", Err.DESCRIPTION, Err.number)
End Function

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

Dim db As Database
Set db = CurrentDb()
db.Execute "INSERT INTO tblDRSUpdateTracking" & sqlColumns & sqlValues
Set db = Nothing

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

progressPercent = 0

Dim total
Dim checked

total = DCount("[Task_ID]", "[tblTaskTracker]", "[Control_Number] = " & controlNum)
checked = DCount("[Task_ID]", "[tblTaskTracker]", "[Control_Number] = " & controlNum & "AND [cbClosed] = TRUE")

If total <> 0 Then progressPercent = checked / total

Exit Function
err_handler:
    Call handleError("wdbDesignE", "progressPercent", Err.DESCRIPTION, Err.number)
End Function