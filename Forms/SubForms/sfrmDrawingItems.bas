Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub itemName_Click()
changeItemImage (Me.itemName)
End Sub

Private Sub itemSelect_Click()
changeItemImage (Me.itemName)
End Sub

Private Sub changeItemImage(itemName As String)
On Error GoTo Err_Handler

Dim folderPath As String
folderPath = "\\data\mdbdata\WorkingDB\Pictures\CATIA_Drawing_Items\"

Form_frmCatiaMacros.itemImage.Picture = folderPath & itemName & ".png"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
