Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim iUser As Integer
Dim strPrevYr As String
Dim strCurYr As String

Private Sub btnApply_Click()
On Error GoTo Err_Handler

setScorecardData

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub filtAdjustedQ1_Click()
On Error GoTo Err_Handler

DoCmd.applyFilter , "[Assignee] = " & iUser & " AND [Completed_Date] Between #10/1/" & CStr(Me.fltYear - 1) & "# And #12/31/" & CStr(Me.fltYear - 1) & "# AND Adjusted = True"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub filtAdjustedQ2_Click()
On Error GoTo Err_Handler

DoCmd.applyFilter , "[Assignee] = " & iUser & " AND [Completed_Date] Between #1/1/" & CStr(Me.fltYear) & "# And #3/31/" & CStr(Me.fltYear) & "# AND Adjusted = True"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub filtAdjustedQ3_Click()
On Error GoTo Err_Handler

DoCmd.applyFilter , "[Assignee] = " & iUser & " AND [Completed_Date] Between #4/1/" & CStr(Me.fltYear) & "# And #6/30/" & CStr(Me.fltYear) & "# AND Adjusted = True"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub filtAdjustedQ4_Click()
On Error GoTo Err_Handler

DoCmd.applyFilter , "[Assignee] = " & iUser & " AND [Completed_Date] Between #7/1/" & CStr(Me.fltYear) & "# And #9/30/" & CStr(Me.fltYear) & "# AND Adjusted = True"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub fltQtr1_Click()
On Error GoTo Err_Handler

DoCmd.applyFilter , "[Assignee] = " & iUser & " AND [Completed_Date] Between #10/1/" & CStr(Me.fltYear - 1) & "# And #12/31/" & CStr(Me.fltYear - 1) & "# AND Late = True"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub fltQtr2_Click()
On Error GoTo Err_Handler

DoCmd.applyFilter , "[Assignee] = " & iUser & " AND [Completed_Date] Between #1/1/" & CStr(Me.fltYear) & "# And #3/31/" & CStr(Me.fltYear) & "# AND Late = True"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub fltQtr3_Click()
On Error GoTo Err_Handler

DoCmd.applyFilter , "[Assignee] = " & iUser & " AND [Completed_Date] Between #4/1/" & CStr(Me.fltYear) & "# And #6/30/" & CStr(Me.fltYear) & "# AND Late = True"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub fltQtr4_Click()
On Error GoTo Err_Handler

DoCmd.applyFilter , "[Assignee] = " & iUser & " AND [Completed_Date] Between #7/1/" & CStr(Me.fltYear) & "# And #9/30/" & CStr(Me.fltYear) & "# AND Late = True"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Function setScorecardData()
'On Error GoTo Err_Handler

Me.lblWho.Caption = Form_frmDRSworkTracker.fltAssignee
iUser = Nz(DLookup("[ID]", "tblPermissions", "[user] = '" & Form_frmDRSworkTracker.fltAssignee & "'"), 0)
    
strPrevYr = CStr(Me.fltYear - 1)
strCurYr = CStr(Me.fltYear)
    
Me.filter = "[Assignee] = " & iUser & " AND [Completed_Date] Between #01/1/" & strCurYr & "# And #12/31/" & strCurYr & "# AND Late = True"
Me.FilterOn = True


Dim Q As New Collection
Q.Add "Q1"
Q.Add "Q2"
Q.Add "Q3"
Q.Add "Q4"

Dim qFilt As String, all As String, assigneeFilt As String, adjFlt As String, jdgFlt As String, qPrevFilt As String
Dim cNum As String
cNum = "[Control_Number]"
assigneeFilt = "[Assignee] = " & iUser
all = "qryApprovedAll"
adjFlt = "Adjusted_Due_Date is not null"
jdgFlt = "[Judgment] = 'Late'"

Dim ITEM

For Each ITEM In Q
    
    Select Case ITEM
        Case "Q1"
            qFilt = "[Completed_Date] Between #10/1/" & strCurYr - 1 & "# And #12/31/" & strCurYr & "#"
            qPrevFilt = "[Completed_Date] Between #10/1/" & strPrevYr - 1 & "# And #12/31/" & strPrevYr & "#"
        Case "Q2"
            qFilt = "[Completed_Date] Between #1/1/" & strCurYr & "# And #3/31/" & strCurYr & "#"
            qPrevFilt = "[Completed_Date] Between #1/1/" & strPrevYr & "# And #3/31/" & strPrevYr & "#"
        Case "Q3"
            qFilt = "[Completed_Date] Between #4/1/" & strCurYr & "# And #6/30/" & strCurYr & "#"
            qPrevFilt = "[Completed_Date] Between #4/1/" & strPrevYr & "# And #6/30/" & strPrevYr & "#"
        Case "Q4"
            qFilt = "[Completed_Date] Between #7/1/" & strCurYr & "# And #9/30/" & strCurYr & "#"
            qPrevFilt = "[Completed_Date] Between #7/1/" & strPrevYr & "# And #9/30/" & strPrevYr & "#"
    End Select

    Me.Controls("txt" & ITEM & "Prev") = _
        Nz(DCount(cNum, all, assigneeFilt & " AND " & qPrevFilt), 0)
        
    Me.Controls("txt" & ITEM & "Cur") = _
        Nz(DCount(cNum, all, assigneeFilt & " AND " & qFilt), 0)
        
    Me.Controls("txt" & ITEM & "TKO") = _
        Nz(DCount(cNum, "qryApprovedTKO", assigneeFilt & " AND " & qFilt), 0)
        
    Me.Controls("txt" & ITEM & "CusMeet") = _
        Nz(DCount(cNum, "qryApprovedCustMeet", assigneeFilt & " AND " & qFilt), 0)
        
    Me.Controls("txt" & ITEM & "ExtCus") = _
        Nz(DCount(cNum, "qryApprovedExtCust", assigneeFilt & " AND " & qFilt), 0)
        
    Me.Controls("txt" & ITEM & "Int") = _
        Nz(DCount(cNum, "qryApprovedInternal", assigneeFilt & " AND " & qFilt), 0)
        
    Me.Controls("txt" & ITEM & "PrevLate") = _
        Nz(DCount(cNum, all, assigneeFilt & " AND " & jdgFlt & " AND " & qPrevFilt), 0)
        
    Me.Controls("txt" & ITEM & "CurLate") = _
        Nz(DCount(cNum, "qryApprovedAll", assigneeFilt & " AND " & jdgFlt & " AND " & qFilt), 0)
        
    Me.Controls("txt" & ITEM & "TKOLate") = _
        Nz(DCount(cNum, "qryApprovedTKO", assigneeFilt & " AND " & jdgFlt & " AND  " & qFilt), 0)
        
    Me.Controls("txt" & ITEM & "CusMeetLate") = _
        Nz(DCount(cNum, "qryApprovedCustMeet", assigneeFilt & " AND " & jdgFlt & " AND " & qFilt), 0)
        
    Me.Controls("txt" & ITEM & "ExtCusLate") = _
        Nz(DCount(cNum, "qryApprovedExtCust", assigneeFilt & " AND " & jdgFlt & " AND " & qFilt), 0)
        
    Me.Controls("txt" & ITEM & "IntLate") = _
        Nz(DCount(cNum, "qryApprovedInternal", assigneeFilt & " AND " & jdgFlt & " AND " & qFilt), 0)
        
    Me.Controls("txt" & ITEM & "PrevPct") = _
        Format(1 - (Me.Controls("txt" & ITEM & "PrevLate") / _
        IIf(Me.Controls("txt" & ITEM & "Prev") = 0, 1, Me.Controls("txt" & ITEM & "Prev"))), "Percent")
        
    Me.Controls("txt" & ITEM & "CurPct") = _
        Format(1 - (Me.Controls("txt" & ITEM & "CurLate") / _
        IIf(Me.Controls("txt" & ITEM & "Cur") = 0, 1, Me.Controls("txt" & ITEM & "Cur"))), "Percent")
    
    Me.Controls("txt" & ITEM & "TKOPct") = _
        Format(1 - (Me.Controls("txt" & ITEM & "TKOLate") / _
        IIf(Me.Controls("txt" & ITEM & "TKO") = 0, 1, Me.Controls("txt" & ITEM & "TKO"))), "Percent")
    
    Me.Controls("txt" & ITEM & "CusMeetPct") = _
        Format(1 - (Me.Controls("txt" & ITEM & "CusMeetLate") / _
        IIf(Me.Controls("txt" & ITEM & "CusMeet") = 0, 1, Me.Controls("txt" & ITEM & "CusMeet"))), "Percent")
    
    Me.Controls("txt" & ITEM & "ExtCusPct") = _
        Format(1 - (Me.Controls("txt" & ITEM & "ExtCusLate") / _
        IIf(Me.Controls("txt" & ITEM & "ExtCus") = 0, 1, Me.Controls("txt" & ITEM & "ExtCus"))), "Percent")
    
    Me.Controls("txt" & ITEM & "IntPct") = _
        Format(1 - (Me.Controls("txt" & ITEM & "IntLate") / _
        IIf(Me.Controls("txt" & ITEM & "Int") = 0, 1, Me.Controls("txt" & ITEM & "Int"))), "Percent")
    
    Me.Controls("txt" & ITEM & "PrevAdj") = _
        Nz(DCount(cNum, all, assigneeFilt & " AND " & adjFlt & " AND " & qPrevFilt), 0)
    
    Me.Controls("txt" & ITEM & "CurAdj") = _
        Nz(DCount(cNum, all, assigneeFilt & " AND " & adjFlt & " AND " & qFilt), 0)
    
    Me.Controls("txt" & ITEM & "TKOAdj") = _
        Nz(DCount(cNum, "qryApprovedTKO", assigneeFilt & " AND " & adjFlt & " AND " & qFilt), 0)
    
    Me.Controls("txt" & ITEM & "CusMeetAdj") = _
        Nz(DCount(cNum, "qryApprovedCustMeet", assigneeFilt & " AND " & adjFlt & " AND " & qFilt), 0)
    
    Me.Controls("txt" & ITEM & "ExtCusAdj") = _
        Nz(DCount(cNum, "qryApprovedExtCust", assigneeFilt & " AND " & adjFlt & " AND " & qFilt), 0)
    
    Me.Controls("txt" & ITEM & "IntAdj") = _
        Nz(DCount(cNum, "qryApprovedInternal", assigneeFilt & " AND " & adjFlt & " AND " & qFilt), 0)
    
Next ITEM
    
'-- set Personal Analytics Values
Me.allCompleted = Nz(DCount(cNum, "qryApprovedAll", assigneeFilt & " AND [Completed_Date] is not null"), 0)
Me.allTime = Nz(DSum("TimeTrack_Work_Hours", "dbo_tblTimeTrackChild", "Associate_ID = " & iUser))
Me.allLate = Nz(DCount(cNum, "qryApprovedAll", assigneeFilt & " AND " & jdgFlt & " AND [Completed_Date] is not null"), 0)
Me.allDeclined = Nz(DCount(cNum, "dbo_tblDRS", assigneeFilt & " AND Approval_Status = 3"))
Me.allCancelled = Nz(DCount(cNum, "dbo_tblDRS", assigneeFilt & " AND Delay_Reason = 11"))
Me.allAdjusted = Nz(DCount(cNum, "qryApprovedAll", assigneeFilt & " AND " & adjFlt & " AND [Completed_Date] is not null"), 0)

Me.latePerc = Format(Me.allLate / Me.allCompleted, "Percent")
Me.declinedPerc = Format(Me.allCancelled / Me.allCompleted, "Percent")
Me.cancelledPerc = Format(Me.allCancelled / Me.allCompleted, "Percent")
Me.adjustedPerc = Format(Me.allAdjusted / Me.allCompleted, "Percent")

Me.allTKOs = Nz(DCount(cNum, "qryApprovedTKO", assigneeFilt & " AND [Completed_Date] is not null"), 0)
Me.TKOsPerc = Format(Me.allTKOs / Me.allCompleted, "Percent")

Dim db As Database
Set db = CurrentDb()
Dim rs As Recordset

Set rs = db.OpenRecordset("SELECT Sum(TimeTrack_Work_Hours) as sumTime " & _
    "FROM dbo_tblTimeTrackChild WHERE [Associate_ID] = " & iUser & _
    " AND Control_Number IN (SELECT Control_Number FROM qryApprovedTKO WHERE [Assignee] = " & iUser & ")")

Me.avgTKO = rs!sumTime / Me.allTKOs
Me.tkoHrPerc = Format(rs!sumTime / Me.allTime, "Percent")

rs.CLOSE
Set rs = Nothing
Set db = Nothing

Exit Function
Err_Handler:
    Call handleError(Me.name, "setScorecardData", Err.DESCRIPTION, Err.number)
End Function


Private Sub Form_Load()
'On Error GoTo Err_Handler

Call setTheme(Me)

Me.fltYear = Year(Date)
    
setScorecardData
    
Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub
