Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Me.Move Form_frmCatiaMacros.WindowLeft + Form_frmCatiaMacros.btnAnchor.Left + Form_frmCatiaMacros.btnAnchor.Width, Form_frmCatiaMacros.WindowTop + Form_frmCatiaMacros.btnAnchor.Top + Form_frmCatiaMacros.btnAnchor.Height

End Sub

Function setAnchor()
On Error GoTo Err_Handler
Call logClick(Me.ActiveControl.name, Me.name)

Form_frmCatiaMacros.btnAnchor.Picture = "\\data\mdbdata\WorkingDB\Pictures\Colored_Icons\" & Me.ActiveControl.name & ".bmp"
Form_frmCatiaMacros.btnAnchor.tag = Me.ActiveControl.name

DoCmd.CLOSE acForm, "frm3DTextAnchor"

Exit Function
Err_Handler:
    MsgBox Err.DESCRIPTION, vbOKOnly, "Error Code: " & Err.number
End Function
