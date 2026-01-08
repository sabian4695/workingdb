Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub backBtn_Click()
On Error GoTo Err_Handler
DoCmd.CLOSE
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler
Me.ShortcutMenu = False

Me.progBar.Width = 6800

Me.lblTitle.tag = TempVars!snackType 'set icon
Me.lblTitle.Caption = TempVars!snackTitle 'set title

Select Case TempVars!snackType 'set progress bar color
    Case "success"
        Me.progBar.BorderColor = rgb(140, 150, 100)
        Me.bxOutline.BorderColor = rgb(140, 150, 100)
    Case "error"
        Me.progBar.BorderColor = rgb(150, 100, 100)
        Me.bxOutline.BorderColor = rgb(150, 100, 100)
    Case "info"
        Me.progBar.BorderColor = rgb(110, 120, 130)
        Me.bxOutline.BorderColor = rgb(110, 120, 130)
End Select

Me.lblMessage = TempVars!snackMessage
If Len(Me.lblMessage) < 48 Then Me.lblMessage.TopMargin = 72 'if it's only one line, add top margin to the text box
If Len(Me.lblMessage) > 112 Then Me.lblMessage.fontSize = 8 'if it's longer than two lines (ish)

Me.Move TempVars!snackLeft, TempVars!snackTop 'set position

Dim arr() As String, subAm
arr = VBA.Split(TempVars!snackMessage, " ")

subAm = 2 * 274.44 / ((UBound(arr) - LBound(arr) + 1)) '250 wpm / 60 seconds

If subAm < 70 Then subAm = 50
If subAm > 200 Then subAm = 200

TempVars.Add "snackSubtract", subAm

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Timer()

If TempVars!snackAutoClose = False Then Exit Sub

If Me.progBar.Width < TempVars!snackSubtract Then
    DoCmd.CLOSE
    Exit Sub
End If
Me.progBar.Width = Me.progBar.Width - Nz(TempVars!snackSubtract, 0)

End Sub
