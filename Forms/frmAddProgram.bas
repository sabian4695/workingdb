Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Load()
On Error GoTo Err_Handler
Call setTheme(Me)
DoCmd.GoToRecord , , acNewRec

Exit Sub
Err_Handler:: Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number): End Sub

Private Sub btnImportPhoto_Click()
On Error GoTo Err_Handler

If Nz(Me.modelCode) = "" Then
    MsgBox "You must enter a Model Code before importing a photo.", vbOKOnly, "Please do as I tell you"
    Exit Sub
End If

Dim fd As FileDialog
Dim FileName As String
    
Set fd = Application.FileDialog(msoFileDialogOpen)
With fd
    .Filters.clear
    .Filters.Add "PNG Files", "*.png"
    .InitialFileName = "C:\Users\Public"
End With
    
fd.Show
On Error GoTo errorCatch
FileName = fd.SelectedItems(1)

Dim general, Program

general = "\\data\mdbdata\WorkingDB\_docs\Program_Review_Docs\"
Program = Me.modelCode

If FolderExists(general & Program) = True Then
    GoTo pathMade
Else
    MkDir (general & Program)
    GoTo pathMade
End If

pathMade:
    Dim fso, FilePath
    Set fso = CreateObject("Scripting.FileSystemObject")
    FilePath = general & Program & "\" & Program & ".png"
    Call fso.CopyFile(FileName, FilePath)
    Me.CarPicture = FilePath

Me.imgCarPicture.Picture = FilePath
Me.btnImportPhoto.Visible = False

errorCatch:
Exit Sub
Err_Handler:: Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number): End Sub

Function validate() As String

validate = ""

Select Case True
    Case Nz(Me.modelCode) = ""
        validate = "Please enter a model code."
    Case Nz(Me.modelYear) = ""
        validate = "Please enter a model year."
    Case Nz(Me.OEM) = ""
        validate = "Please select an OEM."
    Case Nz(Me.modelName) = ""
        validate = "Please enter a model name."
    Case Len(Me.modelYear) <> 4
        validate = "Please adjust your model year to be a 4 digit number"
End Select

If (Me.OEM = "Honda" Or Me.OEM = "Acura") And Len(Me.modelCode) <> 4 Then
    validate = "For Honda, please put the 4 letter code in. The 3 letter code should be the Change Code"
End If

If LCase(Me.modelCode) Like "*multi*" Then
    validate = "Please enter an individual model code. Multi is not an option."
End If

End Function

Private Sub btnSave_Click()
On Error GoTo Err_Handler

Dim val As String
val = validate
If val <> "" Then
    MsgBox val, vbInformation, "Fix it"
    Exit Sub
End If

If Me.Dirty Then Me.Dirty = False

Call registerPartUpdates("tblPrograms", Me.ID, "Created", "", Me.modelCode, "", Me.modelCode)

DoCmd.CLOSE

If CurrentProject.AllForms("frmPrograms").IsLoaded Then
    Form_frmPrograms.Requery
End If

If CurrentProject.AllForms("frmApproveDRS").IsLoaded Then
    Form_frmApproveDRS.cboModelCode.Requery
End If

If CurrentProject.AllForms("frmPartInformation").IsLoaded Then
    Form_frmPartInformation.programId.Requery
End If

Exit Sub
Err_Handler:: Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number): End Sub

Private Sub txtModelCode_AfterUpdate()
On Error GoTo Err_Handler

If DCount("ID", "tblPrograms", "modelCode = '" & Me.modelCode & "'") <> 0 Then
    MsgBox "This model code already exists. This form is for creating new ones.", vbOKOnly, "Duplicate Model Code"
End If

Exit Sub
Err_Handler:: Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number): End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Err_Handler

If validate <> "" Then
    If MsgBox("Are you sure?" & vbNewLine & "Your current record will be deleted.", vbYesNo, "Please confirm") <> vbYes Then
        Cancel = True
        Exit Sub
    End If

    DoCmd.SetWarnings False
    If Nz(Me.modelCode) <> "" Then DoCmd.RunCommand acCmdDeleteRecord
    DoCmd.SetWarnings True
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
