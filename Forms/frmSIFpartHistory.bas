Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnSearch_Click()
On Error GoTo Err_Handler

Dim partNum
partNum = Me.srchBox

If partNum = "" Then
    Me.FilterOn = False
Else
    Me.filter = "[Nifco_Part_Number] = '" & partNum & "'"
    Me.FilterOn = True
End If

Me.srchBox.SetFocus

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub btnSIFDetails_Click()
On Error GoTo Err_Handler

If CurrentProject.AllForms("frmSIF").IsLoaded = False Then
    DoCmd.OpenForm "frmSIF"
End If

Form_frmSIF.srchBox = Me.sifNum
Form_frmSIF.srch_Click

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub clear_Click()
On Error GoTo Err_Handler
Me.srchBox.SetFocus
Me.srchBox = ""
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
