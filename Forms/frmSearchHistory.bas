Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnDelete_Click()
On Error GoTo Err_Handler

dbExecute "DELETE * FROM tblSessionVariables WHERE ID = " & Me.ID
Me.Requery

Dim dNum As Long, finalHeight As Long, i As Long
dNum = DCount("ID", "tblSessionVariables", "searchHistory is not null")

If dNum > 7 Then
    finalHeight = 360 * 7 + 120 + 300
Else
    finalHeight = 360 * dNum + 120 + 300
End If

Me.InsideHeight = finalHeight

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub clickLink_Click()
On Error GoTo Err_Handler

Form_DASHBOARD.partNumberSearch = Me.searchHistory
Form_DASHBOARD.filterbyPN_Click
Form_DASHBOARD.SetFocus

DoCmd.CLOSE acForm, "frmSearchHistory"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub closeForm_Click()
On Error GoTo Err_Handler

DoCmd.CLOSE acForm, "frmSearchHistory"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Application.Echo False
Me.Painting = False

Me.Move Form_DASHBOARD.WindowLeft + Form_DASHBOARD.Command75.Left - 25, Form_DASHBOARD.WindowTop + Form_DASHBOARD.Command75.Top + 310 'set position

Dim dNum As Long, finalHeight As Long, i As Long
dNum = DCount("ID", "tblSessionVariables", "searchHistory is not null")

If dNum > 7 Then
    finalHeight = 360 * 7 + 120 + 300
    Me.ScrollBars = 2
Else
    finalHeight = 360 * dNum + 120 + 300
    Me.ScrollBars = 1
End If

Application.Echo True
Me.Painting = True

Me.SetFocus
Do While i < finalHeight
    Me.InsideHeight = i
    Me.Repaint
    Me.refresh
    i = i + 500
Loop

Me.InsideHeight = finalHeight

Exit Sub
Err_Handler:
Application.Echo True
Me.Painting = True
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Err_Handler

Dim dNum As Long, finalHeight As Long, i As Long
dNum = DCount("ID", "tblSessionVariables", "searchHistory is not null")

If dNum > 7 Then
    finalHeight = 360 * 7 + 120 + 300
Else
    finalHeight = 360 * dNum + 120 + 300
End If

Me.SetFocus
i = finalHeight
Do While i > 0
    Me.InsideHeight = i
    Me.Repaint
    i = i - 350
Loop

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Unload", Err.DESCRIPTION, Err.number)
End Sub
