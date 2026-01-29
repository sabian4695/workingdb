Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub email_Click()
On Error GoTo Err_Handler

Call wdbEmail(Me.userEmail, "", "", "")

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Current()
On Error GoTo Err_Handler
Me.locationFull = ""
Select Case Me.permissionslocation
    Case "CNL"
        Me.locationFull = "Canal Winchester, Ohio"
    Case "SLB"
        Me.locationFull = "Shelbyville, Kentucky"
    Case "LVG"
        Me.locationFull = "La Vergne, Tennessee"
    Case "CUU"
        Me.locationFull = "Chihuahua, Mexico"
    Case "NCM"
        Me.locationFull = "Irapuato, Mexico"
    Case Else
        Me.locationFull = ""
End Select

Dim db As Database
Set db = CurrentDb()

Dim rs As Recordset, rsCount As Recordset
Set rs = db.OpenRecordset("SELECT * FROM tblPartUpdateTracking WHERE updatedBy = '" & Me.User & "'")

'steps closed
rs.filter = "newData = 'Closed' AND tableName = 'tblPartSteps'"
Set rsCount = rs.OpenRecordset

If rsCount.EOF Then
   Me.stepsClosed = 0
Else
   rsCount.MoveLast
   Me.stepsClosed = rsCount.RecordCount
End If

'steps approved
rs.filter = "newData <> 'Deleted' AND newData <> 'Created' AND tableName = 'tblPartTrackingApprovals'"
Set rsCount = rs.OpenRecordset

If rsCount.EOF Then
   Me.stepsApproved = 0
Else
   rsCount.MoveLast
   Me.stepsApproved = rsCount.RecordCount
End If

'files uploaded
rs.filter = "tableName = 'tblPartAttachmentsSP' AND newData = 'Uploaded'"
Set rsCount = rs.OpenRecordset

If rsCount.EOF Then
   Me.filesUploaded = 0
Else
   rsCount.MoveLast
   Me.filesUploaded = rsCount.RecordCount
End If

rs.CLOSE
Set rs = Nothing
Set db = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Current", Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub Label268_Click()
On Error GoTo Err_Handler

Me.dept.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblFullName_Click()
On Error GoTo Err_Handler

Me.rowFullName.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblUser_Click()
On Error GoTo Err_Handler

Me.User.SetFocus
DoCmd.RunCommand acCmdFilterMenu

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub removeFilter_Click()
On Error GoTo Err_Handler

Me.FilterOn = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
