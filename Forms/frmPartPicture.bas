Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub clear_Click()
On Error GoTo Err_Handler

Me.srchBox = ""
Me.srchBox.SetFocus

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub download_Click()
On Error GoTo Err_Handler

If Len(Me.srchBox) < 5 Then Exit Sub

Dim partPic, partPicDir
partPicDir = "\\data\mdbdata\WorkingDB\_docs\Part_Pictures\"
partPic = Dir((partPicDir & Me.srchBox & "*"))
If Len(partPic) > 0 Then
    Dim fso, FilePath, tempFold
    Set fso = CreateObject("Scripting.FileSystemObject")
    tempFold = getTempFold
    If FolderExists(tempFold) = False Then MkDir (tempFold)
    FilePath = tempFold & partPic
    Call fso.CopyFile(partPicDir & partPic, FilePath)
    MsgBox "Done! Opening downloads folder now. File name is: " & partPic, vbOKOnly, "That was easy."
    openPath (tempFold)
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Me.imgPartPicture.Visible = False
Me.srchBox = Form_DASHBOARD.partNumberSearch
Call srch_Click

Me.btnImportPhoto.Visible = userData("Dept") = "Design"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub btnImportPhoto_Click()
On Error GoTo Err_Handler

Dim partNum

partNum = Nz(Me.srchBox)
If IsNull(partNum) Then
    MsgBox "You must enter a part number before importing a photo.", vbOKOnly, "Please do as I tell you"
    Exit Sub
End If
If Len(partNum) < 5 Then
    MsgBox "You must enter a 5 digit part number before importing a photo.", vbOKOnly, "Please do as I tell you"
    Exit Sub
End If

Dim fd As FileDialog
Dim FileName As String
    
Set fd = Application.FileDialog(msoFileDialogOpen)
With fd
    .Filters.clear
    .Filters.Add "PNG Files", "*.png"
End With
    
fd.Show
On Error GoTo errorCatch
FileName = fd.SelectedItems(1)

Dim general

general = "\\data\mdbdata\WorkingDB\_docs\Part_Pictures\"

    Dim fso, FilePath
    Set fso = CreateObject("Scripting.FileSystemObject")
    FilePath = general & "\" & partNum & ".png"
    Call fso.CopyFile(FileName, FilePath)

Me.imgPartPicture.Picture = FilePath
Me.imgPartPicture.Visible = True

Call registerWdbUpdates("tblPartPictures", partNum, "partPicture", "Old Picture", "New Picture")

MsgBox "Uploaded!", vbInformation, "Nice"

errorCatch:
    Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub partPicturesHelp_Click()
On Error GoTo Err_Handler

Call openPath(mainFolder(Me.ActiveControl.name))

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub srch_Click()
On Error GoTo Err_Handler

Me.imgPartPicture.Visible = False

If Len(Me.srchBox) < 5 Then Exit Sub

Dim partPic, partPicDir
partPicDir = "\\data\mdbdata\WorkingDB\_docs\Part_Pictures\"
partPic = Dir((partPicDir & Me.srchBox & "*"))
If Len(partPic) > 0 Then
    Me.imgPartPicture.Picture = partPicDir & partPic
    Me.imgPartPicture.Visible = True
Else
    
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
