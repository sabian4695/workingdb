Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub clear_Click()
On Error GoTo Err_Handler
Me.ECOsrch.SetFocus
Me.ECOsrch = ""
Me.PEsrch.BackColor = rgb(64, 64, 64)
Me.PEsrch.ForeColor = vbWhite
Me.DEsrch.BackColor = rgb(64, 64, 64)
Me.DEsrch.ForeColor = vbWhite
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub DEsrch_Click()
On Error GoTo Err_Handler

Me.Form.filter = "[DE] = '" & UCase(Me.ECOsrch) & "'"
Me.Form.FilterOn = True
Me.DEsrch.BackColor = rgb(222, 174, 0)
Me.DEsrch.ForeColor = vbBlack
Me.PEsrch.BackColor = rgb(64, 64, 64)
Me.PEsrch.ForeColor = vbWhite
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Me.ECOsrch.SetFocus

Dim db As Database
Dim rs1 As Recordset
Set db = CurrentDb()
Set rs1 = db.OpenRecordset("APPS_XXCUS_USER_EMPLOYEES_V", dbOpenSnapshot)

rs1.filter = "USER_NAME = '" & UCase(Environ("username")) & "'"
Set rs1 = rs1.OpenRecordset

Me.ECOsrch = rs1!Last_Name & ", " & rs1!First_Name

rs1.CLOSE
Set rs1 = Nothing
Set db = Nothing

Me.Form.filter = "[DE] = '" & UCase(Me.ECOsrch) & "'"
Me.Form.FilterOn = True

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub PEsrch_Click()
On Error GoTo Err_Handler

Me.Form.filter = "[PE] = '" & UCase(Me.ECOsrch) & "'"
Me.Form.FilterOn = True
Me.PEsrch.BackColor = rgb(166, 166, 166)
Me.PEsrch.ForeColor = vbBlack
Me.DEsrch.BackColor = rgb(64, 64, 64)
Me.DEsrch.ForeColor = vbWhite
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub unAppflt_Click()
On Error GoTo Err_Handler

If Me.PEsrch.BackColor = rgb(166, 166, 166) Then
    Me.Form.filter = "[PE] = '" & UCase(Me.ECOsrch) & "'" & " AND [Approval_Date] IS NULL"
    Me.Form.FilterOn = True
Else
    Me.Form.filter = "[DE] = '" & UCase(Me.ECOsrch) & "'" & " AND [Approval_Date] IS NULL"
    Me.Form.FilterOn = True
End If
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub unImpflt_Click()
On Error GoTo Err_Handler

If Me.PEsrch.BackColor = rgb(166, 166, 166) Then
    Me.Form.filter = "[PE] = '" & UCase(Me.ECOsrch) & "'" & " AND [Implementation_Date] IS NULL" & " AND [Approval_Date] IS NOT NULL"
    Me.Form.FilterOn = True
Else
    Me.Form.filter = "[DE] = '" & UCase(Me.ECOsrch) & "'" & " AND [Implementation_Date] IS NULL" & " AND [Approval_Date] IS NOT NULL"
    Me.Form.FilterOn = True
End If
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
