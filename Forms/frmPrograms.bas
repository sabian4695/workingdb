Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub allEvents_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmProgramEvents"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub btnAddProgram_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmAddProgram"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub btnProgramDetails_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmProgramReview"

Dim Program
Program = Me.modelCode

Form_frmProgramReview.txtFilterInput.Value = Program
Form_frmProgramReview.filterByProgram_Click

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Function applyFilter()
On Error GoTo Err_Handler
Dim filt
filt = ""

'year is the only one that you can use with other filters

If Me.filtYear <> "" Then filt = "modelYear = " & Me.filtYear

If Me.ActiveControl.name = "filtOEM" Then Me.filtModel = Null
If Me.ActiveControl.name = "filtModel" Then Me.filtOEM = Null

If Not IsNull(Me.filtOEM) Then
    If filt <> "" Then filt = filt & " AND "
    filt = filt & "OEM = '" & Me.filtOEM & "'"
    Me.filtModel = ""
    GoTo filtNow
End If

If Not IsNull(Me.filtModel) Then
    If filt <> "" Then filt = filt & " AND "
    filt = filt & "modelCode LIKE '*" & Me.filtModel & "*'"
    Me.filtOEM = ""
    GoTo filtNow
End If

filtNow:
Me.filter = filt
Me.FilterOn = filt <> ""

Exit Function
Err_Handler:
    Call handleError(Me.name, "applyTheFilters", Err.DESCRIPTION, Err.number)
End Function

Private Sub filtOEM_AfterUpdate()
On Error GoTo Err_Handler

Call applyFilter

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub filtModel_AfterUpdate()
On Error GoTo Err_Handler

Call applyFilter

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub filtYear_AfterUpdate()
On Error GoTo Err_Handler

Call applyFilter

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Dim bool1 As Boolean, bool2 As Boolean

bool1 = Not restrict(Environ("username"), "Project")
bool2 = Not restrict(Environ("username"), "Design")

Me.btnAddProgram.Visible = bool1 Or bool2

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub imgUser_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmUserProfile", , , "user = '" & Me.peChampion & "'"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblChangeCode_Click()
    Me.changeCode.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
End Sub

Private Sub lblClassification_Click()
    Me.programClassification.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
End Sub

Private Sub lblModelName_Click()
    Me.modelName.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
End Sub

Private Sub lblModelYear_Click()
    Me.modelYear.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
End Sub

Private Sub lblOem_Click()
    Me.OEM.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
End Sub

Private Sub lblPEchampion_Click()
    Me.peChampion.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
End Sub

Private Sub lblProgram_Click()
    Me.modelCode.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
End Sub

Private Sub lblProjCount_Click()
    Me.projCount.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
End Sub

Private Sub lblSOP_Click()
    Me.SOPdate.SetFocus
    DoCmd.RunCommand acCmdFilterMenu
End Sub
