Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdApply_Click()
On Error GoTo Err_Handler

Dim strSQL As String
    Dim strSQL1 As String               ' account for approval status
    Dim strSQL2 As String               ' account for issue date
    Dim strSQL3 As String               ' acount for assignee
    Dim strSQL4 As String               ' account for requester
    Dim strSQLMid As String
    
'-- approval status
    If IsNull(Me.cboApprovalStatus) Then
        strSQL1 = ""
    Else
        strSQL1 = "[Approval_Status] = " & Me.cboApprovalStatus & " AND "
    End If
    
'-- issue date
    If IsNull(Me.Issue_Date) Then
        strSQL2 = ""
    Else
        strSQL2 = "[Issue_Date] " & Me.cboOperator & " #" & Me.Issue_Date & "# AND "
    End If
    
'-- assignee
    If IsNull(Me.cboAssignee) Then
        strSQL3 = ""
    Else
        strSQL3 = "[Assignee] = " & Me.cboAssignee & " AND "
    End If
    
'-- requester
    If IsNull(Me.cboRequester) Then
        strSQL4 = ""
    Else
        strSQL4 = "[Requester] = " & Me.cboRequester & " AND "
    End If
    
'-- get rid of 'and' at the end
    strSQLMid = strSQL1 & strSQL2 & strSQL3 & strSQL4
    strSQL = Left(strSQLMid, Len(strSQLMid) - 4)
    
'-- filter
    Form_frmApproveDRS.filter = strSQL
    Form_frmApproveDRS.FilterOn = True
    Form_frmApproveDRS.lblFiltered.Visible = True
    Form_frmApproveDRS.cmdFilter.ControlTipText = "Remove Filter"
    DoCmd.CLOSE
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub cmdCancel_Click()
On Error GoTo Err_Handler

    Form_frmApproveDRS.cmdFilter.ControlTipText = "Apply Filter"
    Form_frmApproveDRS.lblFiltered.Visible = False
    DoCmd.CLOSE acForm, Me.name
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
