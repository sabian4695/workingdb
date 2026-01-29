Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnMasterSchedule_Click()
On Error GoTo Err_Handler

If Nz(Me.masterSchedule) <> "" Then openPath (Me.masterSchedule)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub changeCode_AfterUpdate()
On Error GoTo Err_Handler
Call registerPartUpdates("tblPrograms", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "", Me.modelCode)
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub changeType_AfterUpdate()
On Error GoTo Err_Handler
Call registerPartUpdates("tblPrograms", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "", Me.modelCode)
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub editMasterSchedule_Click()
On Error GoTo Err_Handler
Dim masterSched, X As String

masterSched = Nz(Me.masterSchedule, "")

X = InputBox("Paste Link to Master Schedule Here", "Edit Master Schedule Link", masterSched)
If StrPtr(X) = 0 Then Exit Sub
If X = "" Then If MsgBox("Nothing entered. Would you like to clear the master schedule?", vbYesNo, "You didn't type anything") = vbNo Then Exit Sub
Me.masterSchedule = X
Call registerPartUpdates("tblPrograms", Me.ID, Me.masterSchedule.name, Me.masterSchedule.OldValue, Me.masterSchedule, "", Me.modelCode)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Public Sub filterByProgram_Click()
On Error GoTo Err_Handler

Dim Program
Program = Me.txtFilterInput.Value

Me.filter = "[modelCode] = '" & Program & "'"
Me.FilterOn = True
If Me.RecordsetClone.RecordCount = 0 Then
    Me.FilterOn = False
    MsgBox "There are no records" & vbCrLf & "that match this filter.", vbInformation + vbOKOnly, "No records returned"
    Exit Sub
End If

Dim allowIt As Boolean

'PE supervisors/managers + PE Champion can edit
allowIt = (Not restrict(Environ("username"), "Project", "Supervisor", True)) Or (Environ("username") = Nz(Me.peChampion, ""))

Me.allowEdits = allowIt
Me.editMasterSchedule.Enabled = allowIt
Form_sfrmProgramEvents.deletebtn.Enabled = allowIt
Form_sfrmProgramEvents.allowEdits = allowIt

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub gentaniActual_AfterUpdate()
On Error GoTo Err_Handler
Call registerPartUpdates("tblPrograms", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "", Me.modelCode)
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub gentaniProjected_AfterUpdate()
On Error GoTo Err_Handler
Call registerPartUpdates("tblPrograms", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "", Me.modelCode)
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub history_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmHistory", acNormal, , "[tableRecordId] = " & Me.ID & " AND ([tableName] = 'tblPrograms' OR [tableName] = 'tblProgramEvents')"

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

Private Sub importPhoto_Click()
On Error GoTo Err_Handler

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

If Not FolderExists(general & Program) Then MkDir (general & Program)

Dim fso, FilePath
Set fso = CreateObject("Scripting.FileSystemObject")
FilePath = general & Program & "\" & Program & ".png"
Call fso.CopyFile(FileName, FilePath)

Call registerPartUpdates("tblPrograms", Me.ID, "Photo", Me.CarPicture, FilePath, "", Me.modelCode)

dbExecute "UPDATE tblPrograms SET CarPicture = '" & FilePath & "' Where [modelCode] = '" & Program & "'"

Me.Requery

errorCatch:
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub manufacturer_AfterUpdate()
On Error GoTo Err_Handler
Call registerPartUpdates("tblPrograms", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "", Me.modelCode)
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub masterSchedule_AfterUpdate()
On Error GoTo Err_Handler
Call registerPartUpdates("tblPrograms", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "", Me.modelCode)
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub modelCode_AfterUpdate()
On Error GoTo Err_Handler
Call registerPartUpdates("tblPrograms", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "", Me.modelCode)
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub modelName_AfterUpdate()
On Error GoTo Err_Handler
Call registerPartUpdates("tblPrograms", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "", Me.modelCode)
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub modelYear_AfterUpdate()
On Error GoTo Err_Handler
Call registerPartUpdates("tblPrograms", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "", Me.modelCode)
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub OEM_AfterUpdate()
On Error GoTo Err_Handler
Call registerPartUpdates("tblPrograms", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "", Me.modelCode)
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub peChampion_AfterUpdate()
On Error GoTo Err_Handler
Call registerPartUpdates("tblPrograms", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "", Me.modelCode)
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub SOPdate_AfterUpdate()
On Error GoTo Err_Handler
Call registerPartUpdates("tblPrograms", Me.ID, Me.ActiveControl.name, Me.ActiveControl.OldValue, Me.ActiveControl, "", Me.modelCode)
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
