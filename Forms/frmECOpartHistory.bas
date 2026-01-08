Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnECODetails_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmECOs", , , "[CHANGE_NOTICE] = '" & UCase(Me.CHANGE_NOTICE) & "'"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub clear_Click()
On Error GoTo Err_Handler
Me.ecobyRev.SetFocus
Me.ecobyRev = ""
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub revItemSrch_Click()
On Error GoTo Err_Handler

Dim partNum
partNum = Me.ecobyRev
If partNum <> "" Then partNum = idNAM(partNum, "NAM")

If partNum = "" Then
    Me.FilterOn = False
    MsgBox "Part number not found", vbInformation, "Huh..."
Else
    Me.filter = "[REVISED_ITEM_ID] = " & partNum
    Me.FilterOn = True
End If

Me.ecobyRev.SetFocus

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Dim partNum
partNum = Form_DASHBOARD.partNumberSearch
Me.ecobyRev = partNum
If partNum <> "" Then partNum = idNAM(partNum, "NAM")

If partNum = "" Or IsNull(partNum) Then
    Me.FilterOn = False
Else
    Me.filter = "[REVISED_ITEM_ID] = " & partNum
    Me.FilterOn = True
End If

Me.ecobyRev.SetFocus

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub
