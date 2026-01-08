Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim iUser As Integer
Dim strPrevYr As String
Dim strCurYr As String

Private Sub filtAdjustedQ1_Click()
On Error GoTo Err_Handler

DoCmd.applyFilter , "[Assignee] = " & iUser & " AND [Completed_Date] Between #1/1/" & strCurYr & "# And #3/31/" & strCurYr & "# AND Adjusted = True"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub filtAdjustedQ2_Click()
On Error GoTo Err_Handler

DoCmd.applyFilter , "[Assignee] = " & iUser & " AND [Completed_Date] Between #4/1/" & strCurYr & "# And #6/30/" & strCurYr & "# AND Adjusted = True"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub filtAdjustedQ3_Click()
On Error GoTo Err_Handler

DoCmd.applyFilter , "[Assignee] = " & iUser & " AND [Completed_Date] Between #7/1/" & strCurYr & "# And #9/30/" & strCurYr & "# AND Adjusted = True"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub filtAdjustedQ4_Click()
On Error GoTo Err_Handler

DoCmd.applyFilter , "[Assignee] = " & iUser & " AND [Completed_Date] Between #10/1/" & strCurYr & "# And #12/31/" & strCurYr & "# AND Adjusted = True"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub fltQtr1_Click()
On Error GoTo Err_Handler

DoCmd.applyFilter , "[Assignee] = " & iUser & " AND [Completed_Date] Between #1/1/" & strCurYr & "# And #3/31/" & strCurYr & "# AND Late = True"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub fltQtr2_Click()
On Error GoTo Err_Handler

DoCmd.applyFilter , "[Assignee] = " & iUser & " AND [Completed_Date] Between #4/1/" & strCurYr & "# And #6/30/" & strCurYr & "# AND Late = True"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub fltQtr3_Click()
On Error GoTo Err_Handler

DoCmd.applyFilter , "[Assignee] = " & iUser & " AND [Completed_Date] Between #7/1/" & strCurYr & "# And #9/30/" & strCurYr & "# AND Late = True"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub fltQtr4_Click()
On Error GoTo Err_Handler

DoCmd.applyFilter , "[Assignee] = " & iUser & " AND [Completed_Date] Between #10/1/" & strCurYr & "# And #12/31/" & strCurYr & "# AND Late = True"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

'-- dimension the variables
    
'-- define the variables
    Me.lblWho.Caption = Form_frmDRSworkTracker.fltAssignee
    iUser = Nz(DLookup("[ID]", "tblPermissions", "[user] = '" & Form_frmDRSworkTracker.fltAssignee & "'"), 0)
    strPrevYr = CStr(Format(DateAdd("yyyy", -1, Date), "yyyy"))
    strCurYr = CStr(Format(Date, "yyyy"))
    
    Me.filter = "[Assignee] = " & iUser & " AND [Completed_Date] Between #01/1/" & strCurYr & "# And #12/31/" & strCurYr & "# AND Late = True"
    Me.FilterOn = True
    
'-- set q1 values
    Me.txtQ1Prev = Nz(DCount("[Control_Number]", "qryApprovedAll", "[Assignee] = " & iUser & " AND [Completed_Date] Between #1/1/" & strPrevYr & "# And #3/31/" & strPrevYr & "#"), 0)
    Me.txtQ1Cur = Nz(DCount("[Control_Number]", "qryApprovedAll", "[Assignee] = " & iUser & " AND [Completed_Date] Between #1/1/" & strCurYr & "# And #3/31/" & strCurYr & "#"), 0)
    Me.txtQ1TKO = Nz(DCount("[Control_Number]", "qryApprovedTKO", "[Assignee] = " & iUser & " AND [Completed_Date] Between #1/1/" & strCurYr & "# And #3/31/" & strCurYr & "#"), 0)
    Me.txtQ1CusMeet = Nz(DCount("[Control_Number]", "qryApprovedCustMeet", "[Assignee] = " & iUser & " AND [Completed_Date] Between #1/1/" & strCurYr & "# And #3/31/" & strCurYr & "#"), 0)
    Me.txtQ1ExtCus = Nz(DCount("[Control_Number]", "qryApprovedExtCust", "[Assignee] = " & iUser & " AND [Completed_Date] Between #1/1/" & strCurYr & "# And #3/31/" & strCurYr & "#"), 0)
    Me.txtQ1Int = Nz(DCount("[Control_Number]", "qryApprovedInternal", "[Assignee] = " & iUser & " AND [Completed_Date] Between #1/1/" & strCurYr & "# And #3/31/" & strCurYr & "#"), 0)
    Me.txtQ1PrevLate = Nz(DCount("[Control_Number]", "qryApprovedAll", "[Assignee] = " & iUser & " AND [Judgment] = 'Late' AND [Completed_Date] Between #1/1/" & strPrevYr & "# And #3/31/" & strPrevYr & "#"), 0)
    Me.txtQ1CurLate = Nz(DCount("[Control_Number]", "qryApprovedAll", "[Assignee] = " & iUser & " AND [Judgment] = 'Late' AND [Completed_Date] Between #1/1/" & strCurYr & "# And #3/31/" & strCurYr & "#"), 0)
    Me.txtQ1TKOLate = Nz(DCount("[Control_Number]", "qryApprovedTKO", "[Assignee] = " & iUser & " AND [Judgment] = 'Late' AND [Completed_Date] Between #1/1/" & strCurYr & "# And #3/31/" & strCurYr & "#"), 0)
    Me.txtQ1CusMeetLate = Nz(DCount("[Control_Number]", "qryApprovedCustMeet", "[Assignee] = " & iUser & " AND [Judgment] = 'Late' AND [Completed_Date] Between #1/1/" & strCurYr & "# And #3/31/" & strCurYr & "#"), 0)
    Me.txtQ1ExtCusLate = Nz(DCount("[Control_Number]", "qryApprovedExtCust", "[Assignee] = " & iUser & " AND [Judgment] = 'Late' AND [Completed_Date] Between #1/1/" & strCurYr & "# And #3/31/" & strCurYr & "#"), 0)
    Me.txtQ1IntLate = Nz(DCount("[Control_Number]", "qryApprovedInternal", "[Assignee] = " & iUser & " AND [Judgment] = 'Late' AND [Completed_Date] Between #1/1/" & strCurYr & "# And #3/31/" & strCurYr & "#"), 0)
    Me.txtQ1PrevPct = Format(1 - (Me.txtQ1PrevLate / IIf(Me.txtQ1Prev = 0, 1, Me.txtQ1Prev)), "Percent")
    Me.txtQ1CurPct = Format(1 - (Me.txtQ1CurLate / IIf(Me.txtQ1Cur = 0, 1, Me.txtQ1Cur)), "Percent")
    Me.txtQ1TKOPct = Format(1 - (Me.txtQ1TKOLate / IIf(Me.txtQ1TKO = 0, 1, Me.txtQ1TKO)), "Percent")
    Me.txtQ1CusMeetPct = Format(1 - (Me.txtQ1CusMeetLate / IIf(Me.txtQ1CusMeet = 0, 1, Me.txtQ1CusMeet)), "Percent")
    Me.txtQ1ExtCusPct = Format(1 - (Me.txtQ1ExtCusLate / IIf(Me.txtQ1ExtCus = 0, 1, Me.txtQ1ExtCus)), "Percent")
    Me.txtQ1IntPct = Format(1 - (Me.txtQ1IntLate / IIf(Me.txtQ1Int = 0, 1, Me.txtQ1Int)), "Percent")
    Me.txtQ1PrevAdj = Nz(DCount("[Control_Number]", "qryApprovedAll", "[Assignee] = " & iUser & " AND Adjusted_Due_Date is not null AND [Completed_Date] Between #1/1/" & strPrevYr & "# And #3/31/" & strPrevYr & "#"), 0)
    Me.txtQ1CurAdj = Nz(DCount("[Control_Number]", "qryApprovedAll", "[Assignee] = " & iUser & " AND Adjusted_Due_Date is not null AND [Completed_Date] Between #1/1/" & strCurYr & "# And #3/31/" & strCurYr & "#"), 0)
    Me.txtQ1TKOAdj = Nz(DCount("[Control_Number]", "qryApprovedTKO", "[Assignee] = " & iUser & " AND Adjusted_Due_Date is not null AND [Completed_Date] Between #1/1/" & strCurYr & "# And #3/31/" & strCurYr & "#"), 0)
    Me.txtQ1CusMeetAdj = Nz(DCount("[Control_Number]", "qryApprovedCustMeet", "[Assignee] = " & iUser & " AND Adjusted_Due_Date is not null AND [Completed_Date] Between #1/1/" & strCurYr & "# And #3/31/" & strCurYr & "#"), 0)
    Me.txtQ1ExtCusAdj = Nz(DCount("[Control_Number]", "qryApprovedExtCust", "[Assignee] = " & iUser & " AND Adjusted_Due_Date is not null AND [Completed_Date] Between #1/1/" & strCurYr & "# And #3/31/" & strCurYr & "#"), 0)
    Me.txtQ1IntAdj = Nz(DCount("[Control_Number]", "qryApprovedInternal", "[Assignee] = " & iUser & " AND Adjusted_Due_Date is not null AND [Completed_Date] Between #1/1/" & strCurYr & "# And #3/31/" & strCurYr & "#"), 0)

'-- set q2 values
    Me.txtQ2Prev = Nz(DCount("[Control_Number]", "qryApprovedAll", "[Assignee] = " & iUser & " AND [Completed_Date] Between #4/1/" & strPrevYr & "# And #6/30/" & strPrevYr & "#"), 0)
    Me.txtQ2Cur = Nz(DCount("[Control_Number]", "qryApprovedAll", "[Assignee] = " & iUser & " AND [Completed_Date] Between #4/1/" & strCurYr & "# And #6/30/" & strCurYr & "#"), 0)
    Me.txtQ2TKO = Nz(DCount("[Control_Number]", "qryApprovedTKO", "[Assignee] = " & iUser & " AND [Completed_Date] Between #4/1/" & strCurYr & "# And #6/30/" & strCurYr & "#"), 0)
    Me.txtQ2CusMeet = Nz(DCount("[Control_Number]", "qryApprovedCustMeet", "[Assignee] = " & iUser & " AND [Completed_Date] Between #4/1/" & strCurYr & "# And #6/30/" & strCurYr & "#"), 0)
    Me.txtQ2ExtCus = Nz(DCount("[Control_Number]", "qryApprovedExtCust", "[Assignee] = " & iUser & " AND [Completed_Date] Between #4/1/" & strCurYr & "# And #6/30/" & strCurYr & "#"), 0)
    Me.txtQ2Int = Nz(DCount("[Control_Number]", "qryApprovedInternal", "[Assignee] = " & iUser & " AND [Completed_Date] Between #4/1/" & strCurYr & "# And #6/30/" & strCurYr & "#"), 0)
    Me.txtQ2PrevLate = Nz(DCount("[Control_Number]", "qryApprovedAll", "[Assignee] = " & iUser & " AND [Judgment] = 'Late' AND [Completed_Date] Between #4/1/" & strPrevYr & "# And #6/30/" & strPrevYr & "#"), 0)
    Me.txtQ2CurLate = Nz(DCount("[Control_Number]", "qryApprovedAll", "[Assignee] = " & iUser & " AND [Judgment] = 'Late' AND [Completed_Date] Between #4/1/" & strCurYr & "# And #6/30/" & strCurYr & "#"), 0)
    Me.txtQ2TKOLate = Nz(DCount("[Control_Number]", "qryApprovedTKO", "[Assignee] = " & iUser & " AND [Judgment] = 'Late' AND [Completed_Date] Between #4/1/" & strCurYr & "# And #6/30/" & strCurYr & "#"), 0)
    Me.txtQ2CusMeetLate = Nz(DCount("[Control_Number]", "qryApprovedCustMeet", "[Assignee] = " & iUser & " AND [Judgment] = 'Late' AND [Completed_Date] Between #4/1/" & strCurYr & "# And #6/30/" & strCurYr & "#"), 0)
    Me.txtQ2ExtCusLate = Nz(DCount("[Control_Number]", "qryApprovedExtCust", "[Assignee] = " & iUser & " AND [Judgment] = 'Late' AND [Completed_Date] Between #4/1/" & strCurYr & "# And #6/30/" & strCurYr & "#"), 0)
    Me.txtQ2IntLate = Nz(DCount("[Control_Number]", "qryApprovedInternal", "[Assignee] = " & iUser & " AND [Judgment] = 'Late' AND [Completed_Date] Between #4/1/" & strCurYr & "# And #6/30/" & strCurYr & "#"), 0)
    Me.txtQ2PrevPct = Format(1 - (Me.txtQ2PrevLate / IIf(Me.txtQ2Prev = 0, 1, Me.txtQ2Prev)), "Percent")
    Me.txtQ2CurPct = Format(1 - (Me.txtQ2CurLate / IIf(Me.txtQ2Cur = 0, 1, Me.txtQ2Cur)), "Percent")
    Me.txtQ2TKOPct = Format(1 - (Me.txtQ2TKOLate / IIf(Me.txtQ2TKO = 0, 1, Me.txtQ2TKO)), "Percent")
    Me.txtQ2CusMeetPct = Format(1 - (Me.txtQ2CusMeetLate / IIf(Me.txtQ2CusMeet = 0, 1, Me.txtQ2CusMeet)), "Percent")
    Me.txtQ2ExtCusPct = Format(1 - (Me.txtQ2ExtCusLate / IIf(Me.txtQ2ExtCus = 0, 1, Me.txtQ2ExtCus)), "Percent")
    Me.txtQ2IntPct = Format(1 - (Me.txtQ2IntLate / IIf(Me.txtQ2Int = 0, 1, Me.txtQ2Int)), "Percent")
    Me.txtQ2PrevAdj = Nz(DCount("[Control_Number]", "qryApprovedAll", "[Assignee] = " & iUser & " AND Adjusted_Due_Date is not null AND [Completed_Date] Between #4/1/" & strPrevYr & "# And #6/30/" & strPrevYr & "#"), 0)
    Me.txtQ2CurAdj = Nz(DCount("[Control_Number]", "qryApprovedAll", "[Assignee] = " & iUser & " AND Adjusted_Due_Date is not null AND [Completed_Date] Between #4/1/" & strCurYr & "# And #6/30/" & strCurYr & "#"), 0)
    Me.txtQ2TKOAdj = Nz(DCount("[Control_Number]", "qryApprovedTKO", "[Assignee] = " & iUser & " AND Adjusted_Due_Date is not null AND [Completed_Date] Between #4/1/" & strCurYr & "# And #6/30/" & strCurYr & "#"), 0)
    Me.txtQ2CusMeetAdj = Nz(DCount("[Control_Number]", "qryApprovedCustMeet", "[Assignee] = " & iUser & " AND Adjusted_Due_Date is not null AND [Completed_Date] Between #4/1/" & strCurYr & "# And #6/30/" & strCurYr & "#"), 0)
    Me.txtQ2ExtCusAdj = Nz(DCount("[Control_Number]", "qryApprovedExtCust", "[Assignee] = " & iUser & " AND Adjusted_Due_Date is not null AND [Completed_Date] Between #4/1/" & strCurYr & "# And #6/30/" & strCurYr & "#"), 0)
    Me.txtQ2IntAdj = Nz(DCount("[Control_Number]", "qryApprovedInternal", "[Assignee] = " & iUser & " AND Adjusted_Due_Date is not null AND [Completed_Date] Between #4/1/" & strCurYr & "# And #6/30/" & strCurYr & "#"), 0)

'-- set q3 values
    Me.txtQ3Prev = Nz(DCount("[Control_Number]", "qryApprovedAll", "[Assignee] = " & iUser & " AND [Completed_Date] Between #7/1/" & strPrevYr & "# And #9/30/" & strPrevYr & "#"), 0)
    Me.txtQ3Cur = Nz(DCount("[Control_Number]", "qryApprovedAll", "[Assignee] = " & iUser & " AND [Completed_Date] Between #7/1/" & strCurYr & "# And #9/30/" & strCurYr & "#"), 0)
    Me.txtQ3TKO = Nz(DCount("[Control_Number]", "qryApprovedTKO", "[Assignee] = " & iUser & " AND [Completed_Date] Between #7/1/" & strCurYr & "# And #9/30/" & strCurYr & "#"), 0)
    Me.txtQ3CusMeet = Nz(DCount("[Control_Number]", "qryApprovedCustMeet", "[Assignee] = " & iUser & " AND [Completed_Date] Between #7/1/" & strCurYr & "# And #9/30/" & strCurYr & "#"), 0)
    Me.txtQ3ExtCus = Nz(DCount("[Control_Number]", "qryApprovedExtCust", "[Assignee] = " & iUser & " AND [Completed_Date] Between #7/1/" & strCurYr & "# And #9/30/" & strCurYr & "#"), 0)
    Me.txtQ3Int = Nz(DCount("[Control_Number]", "qryApprovedInternal", "[Assignee] = " & iUser & " AND [Completed_Date] Between #7/1/" & strCurYr & "# And #9/30/" & strCurYr & "#"), 0)
    Me.txtQ3PrevLate = Nz(DCount("[Control_Number]", "qryApprovedAll", "[Assignee] = " & iUser & " AND [Judgment] = 'Late' AND [Completed_Date] Between #7/1/" & strPrevYr & "# And #9/30/" & strPrevYr & "#"), 0)
    Me.txtQ3CurLate = Nz(DCount("[Control_Number]", "qryApprovedAll", "[Assignee] = " & iUser & " AND [Judgment] = 'Late' AND [Completed_Date] Between #7/1/" & strCurYr & "# And #9/30/" & strCurYr & "#"), 0)
    Me.txtQ3TKOLate = Nz(DCount("[Control_Number]", "qryApprovedTKO", "[Assignee] = " & iUser & " AND [Judgment] = 'Late' AND [Completed_Date] Between #7/1/" & strCurYr & "# And #9/30/" & strCurYr & "#"), 0)
    Me.txtQ3CusMeetLate = Nz(DCount("[Control_Number]", "qryApprovedCustMeet", "[Assignee] = " & iUser & " AND [Judgment] = 'Late' AND [Completed_Date] Between #7/1/" & strCurYr & "# And #9/30/" & strCurYr & "#"), 0)
    Me.txtQ3ExtCusLate = Nz(DCount("[Control_Number]", "qryApprovedExtCust", "[Assignee] = " & iUser & " AND [Judgment] = 'Late' AND [Completed_Date] Between #7/1/" & strCurYr & "# And #9/30/" & strCurYr & "#"), 0)
    Me.txtQ3IntLate = Nz(DCount("[Control_Number]", "qryApprovedInternal", "[Assignee] = " & iUser & " AND [Judgment] = 'Late' AND [Completed_Date] Between #7/1/" & strCurYr & "# And #9/30/" & strCurYr & "#"), 0)
    Me.txtQ3PrevPct = Format(1 - (Me.txtQ3PrevLate / IIf(Me.txtQ3Prev = 0, 1, Me.txtQ3Prev)), "Percent")
    Me.txtQ3CurPct = Format(1 - (Me.txtQ3CurLate / IIf(Me.txtQ3Cur = 0, 1, Me.txtQ3Cur)), "Percent")
    Me.txtQ3TKOPct = Format(1 - (Me.txtQ3TKOLate / IIf(Me.txtQ3TKO = 0, 1, Me.txtQ3TKO)), "Percent")
    Me.txtQ3CusMeetPct = Format(1 - (Me.txtQ3CusMeetLate / IIf(Me.txtQ3CusMeet = 0, 1, Me.txtQ3CusMeet)), "Percent")
    Me.txtQ3ExtCusPct = Format(1 - (Me.txtQ3ExtCusLate / IIf(Me.txtQ3ExtCus = 0, 1, Me.txtQ3ExtCus)), "Percent")
    Me.txtQ3IntPct = Format(1 - (Me.txtQ3IntLate / IIf(Me.txtQ3Int = 0, 1, Me.txtQ3Int)), "Percent")
    Me.txtQ3PrevAdj = Nz(DCount("[Control_Number]", "qryApprovedAll", "[Assignee] = " & iUser & " AND Adjusted_Due_Date is not null AND [Completed_Date] Between #7/1/" & strPrevYr & "# And #9/30/" & strPrevYr & "#"), 0)
    Me.txtQ3CurAdj = Nz(DCount("[Control_Number]", "qryApprovedAll", "[Assignee] = " & iUser & " AND Adjusted_Due_Date is not null AND [Completed_Date] Between #7/1/" & strCurYr & "# And #9/30/" & strCurYr & "#"), 0)
    Me.txtQ3TKOAdj = Nz(DCount("[Control_Number]", "qryApprovedTKO", "[Assignee] = " & iUser & " AND Adjusted_Due_Date is not null AND [Completed_Date] Between #7/1/" & strCurYr & "# And #9/30/" & strCurYr & "#"), 0)
    Me.txtQ3CusMeetAdj = Nz(DCount("[Control_Number]", "qryApprovedCustMeet", "[Assignee] = " & iUser & " AND Adjusted_Due_Date is not null AND [Completed_Date] Between #7/1/" & strCurYr & "# And #9/30/" & strCurYr & "#"), 0)
    Me.txtQ3ExtCusAdj = Nz(DCount("[Control_Number]", "qryApprovedExtCust", "[Assignee] = " & iUser & " AND Adjusted_Due_Date is not null AND [Completed_Date] Between #7/1/" & strCurYr & "# And #9/30/" & strCurYr & "#"), 0)
    Me.txtQ3IntAdj = Nz(DCount("[Control_Number]", "qryApprovedInternal", "[Assignee] = " & iUser & " AND Adjusted_Due_Date is not null AND [Completed_Date] Between #7/1/" & strCurYr & "# And #9/30/" & strCurYr & "#"), 0)

'-- set q4 values
    Me.txtQ4Prev = Nz(DCount("[Control_Number]", "qryApprovedAll", "[Assignee] = " & iUser & " AND [Completed_Date] Between #10/1/" & strPrevYr & "# And #12/31/" & strPrevYr & "#"), 0)
    Me.txtQ4Cur = Nz(DCount("[Control_Number]", "qryApprovedAll", "[Assignee] = " & iUser & " AND [Completed_Date] Between #10/1/" & strCurYr & "# And #12/31/" & strCurYr & "#"), 0)
    Me.txtQ4TKO = Nz(DCount("[Control_Number]", "qryApprovedTKO", "[Assignee] = " & iUser & " AND [Completed_Date] Between #10/1/" & strCurYr & "# And #12/31/" & strCurYr & "#"), 0)
    Me.txtQ4CusMeet = Nz(DCount("[Control_Number]", "qryApprovedCustMeet", "[Assignee] = " & iUser & " AND [Completed_Date] Between #10/1/" & strCurYr & "# And #12/31/" & strCurYr & "#"), 0)
    Me.txtQ4ExtCus = Nz(DCount("[Control_Number]", "qryApprovedExtCust", "[Assignee] = " & iUser & " AND [Completed_Date] Between #10/1/" & strCurYr & "# And #12/31/" & strCurYr & "#"), 0)
    Me.txtQ4Int = Nz(DCount("[Control_Number]", "qryApprovedInternal", "[Assignee] = " & iUser & " AND [Completed_Date] Between #10/1/" & strCurYr & "# And #12/31/" & strCurYr & "#"), 0)
    Me.txtQ4PrevLate = Nz(DCount("[Control_Number]", "qryApprovedAll", "[Assignee] = " & iUser & " AND [Judgment] = 'Late' AND [Completed_Date] Between #10/1/" & strPrevYr & "# And #12/31/" & strPrevYr & "#"), 0)
    Me.txtQ4CurLate = Nz(DCount("[Control_Number]", "qryApprovedAll", "[Assignee] = " & iUser & " AND [Judgment] = 'Late' AND [Completed_Date] Between #10/1/" & strCurYr & "# And #12/31/" & strCurYr & "#"), 0)
    Me.txtQ4TKOLate = Nz(DCount("[Control_Number]", "qryApprovedTKO", "[Assignee] = " & iUser & " AND [Judgment] = 'Late' AND [Completed_Date] Between #10/1/" & strCurYr & "# And #12/31/" & strCurYr & "#"), 0)
    Me.txtQ4CusMeetLate = Nz(DCount("[Control_Number]", "qryApprovedCustMeet", "[Assignee] = " & iUser & " AND [Judgment] = 'Late' AND [Completed_Date] Between #10/1/" & strCurYr & "# And #12/31/" & strCurYr & "#"), 0)
    Me.txtQ4ExtCusLate = Nz(DCount("[Control_Number]", "qryApprovedExtCust", "[Assignee] = " & iUser & " AND [Judgment] = 'Late' AND [Completed_Date] Between #10/1/" & strCurYr & "# And #12/31/" & strCurYr & "#"), 0)
    Me.txtQ4IntLate = Nz(DCount("[Control_Number]", "qryApprovedInternal", "[Assignee] = " & iUser & " AND [Judgment] = 'Late' AND [Completed_Date] Between #10/1/" & strCurYr & "# And #12/31/" & strCurYr & "#"), 0)
    Me.txtQ4PrevPct = Format(1 - (Me.txtQ4PrevLate / IIf(Me.txtQ4Prev = 0, 1, Me.txtQ4Prev)), "Percent")
    Me.txtQ4CurPct = Format(1 - (Me.txtQ4CurLate / IIf(Me.txtQ4Cur = 0, 1, Me.txtQ4Cur)), "Percent")
    Me.txtQ4TKOPct = Format(1 - (Me.txtQ4TKOLate / IIf(Me.txtQ4TKO = 0, 1, Me.txtQ4TKO)), "Percent")
    Me.txtQ4CusMeetPct = Format(1 - (Me.txtQ4CusMeetLate / IIf(Me.txtQ4CusMeet = 0, 1, Me.txtQ4CusMeet)), "Percent")
    Me.txtQ4ExtCusPct = Format(1 - (Me.txtQ4ExtCusLate / IIf(Me.txtQ4ExtCus = 0, 1, Me.txtQ4ExtCus)), "Percent")
    Me.txtQ4IntPct = Format(1 - (Me.txtQ4IntLate / IIf(Me.txtQ4Int = 0, 1, Me.txtQ4Int)), "Percent")
    Me.txtQ4PrevAdj = Nz(DCount("[Control_Number]", "qryApprovedAll", "[Assignee] = " & iUser & " AND Adjusted_Due_Date is not null AND [Completed_Date] Between #10/1/" & strPrevYr & "# And #12/31/" & strPrevYr & "#"), 0)
    Me.txtQ4CurAdj = Nz(DCount("[Control_Number]", "qryApprovedAll", "[Assignee] = " & iUser & " AND Adjusted_Due_Date is not null AND [Completed_Date] Between #10/1/" & strCurYr & "# And #12/31/" & strCurYr & "#"), 0)
    Me.txtQ4TKOAdj = Nz(DCount("[Control_Number]", "qryApprovedTKO", "[Assignee] = " & iUser & " AND Adjusted_Due_Date is not null AND [Completed_Date] Between #10/1/" & strCurYr & "# And #12/31/" & strCurYr & "#"), 0)
    Me.txtQ4CusMeetAdj = Nz(DCount("[Control_Number]", "qryApprovedCustMeet", "[Assignee] = " & iUser & " AND Adjusted_Due_Date is not null AND [Completed_Date] Between #10/1/" & strCurYr & "# And #12/31/" & strCurYr & "#"), 0)
    Me.txtQ4ExtCusAdj = Nz(DCount("[Control_Number]", "qryApprovedExtCust", "[Assignee] = " & iUser & " AND Adjusted_Due_Date is not null AND [Completed_Date] Between #10/1/" & strCurYr & "# And #12/31/" & strCurYr & "#"), 0)
    Me.txtQ4IntAdj = Nz(DCount("[Control_Number]", "qryApprovedInternal", "[Assignee] = " & iUser & " AND Adjusted_Due_Date is not null AND [Completed_Date] Between #10/1/" & strCurYr & "# And #12/31/" & strCurYr & "#"), 0)
    
'-- set Personal Analytics Values
Me.allCompleted = Nz(DCount("[Control_Number]", "qryApprovedAll", "[Assignee] = " & iUser & " AND [Completed_Date] is not null"), 0)
Me.allTime = Nz(DSum("TimeTrack_Work_Hours", "dbo_tblTimeTrackChild", "Associate_ID = " & iUser))
Me.allLate = Nz(DCount("[Control_Number]", "qryApprovedAll", "[Assignee] = " & iUser & " AND [Judgment] = 'Late' AND [Completed_Date] is not null"), 0)
Me.allDeclined = Nz(DCount("[Control_Number]", "dbo_tblDRS", "[Assignee] = " & iUser & " AND Approval_Status = 3"))
Me.allCancelled = Nz(DCount("[Control_Number]", "dbo_tblDRS", "[Assignee] = " & iUser & " AND Delay_Reason = 11"))
Me.allAdjusted = Nz(DCount("[Control_Number]", "qryApprovedAll", "[Assignee] = " & iUser & " AND Adjusted_Due_Date is not null AND [Completed_Date] is not null"), 0)

Me.latePerc = Format(Me.allLate / Me.allCompleted, "Percent")
Me.declinedPerc = Format(Me.allCancelled / Me.allCompleted, "Percent")
Me.cancelledPerc = Format(Me.allCancelled / Me.allCompleted, "Percent")
Me.adjustedPerc = Format(Me.allAdjusted / Me.allCompleted, "Percent")

Me.allTKOs = Nz(DCount("[Control_Number]", "qryApprovedTKO", "[Assignee] = " & iUser & " AND [Completed_Date] is not null"), 0)
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
    
Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub
