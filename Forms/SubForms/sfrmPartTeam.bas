Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub emailCFteam_Click()
On Error GoTo Err_Handler

Dim objEmail As Object

Set objEmail = CreateObject("outlook.Application")
Set objEmail = objEmail.CreateItem(0)

With objEmail
    .To = grabPartTeam(Me.partNumber, True)
    .subject = CStr(Me.partNumber.Value)
    .display
End With

Set objEmail = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Dim allowEdit As Boolean

allowEdit = False
If Not restrict(Environ("username"), "Project") Or Not restrict(Environ("username"), "Service") Or userData("Level") = "Manager" Then allowEdit = True

Me.allowEdits = allowEdit
Me.AllowAdditions = allowEdit
Me.remove.Visible = allowEdit

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub imgPart_Click()
On Error GoTo Err_Handler

DoCmd.OpenForm "frmUserProfile", , , "user = '" & Me.person & "'"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub person_AfterUpdate()
On Error GoTo Err_Handler

Call registerPartUpdates("tblPartInfo", Me.recordId, Me.ActiveControl.name, Nz(Me.ActiveControl.OldValue, ""), Me.ActiveControl, Me.partNumber, Me.name)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name & "_AfterUpdate", Err.DESCRIPTION, Err.number)
End Sub

Private Sub person_BeforeUpdate(Cancel As Integer)
On Error GoTo Err_Handler

Dim errorMsg As String
Dim projOwner As Boolean
Dim deptSuper As Boolean

'check for duplicate. don't allow overrides
If DCount("person", "tblPartTeam", "partNumber = '" & Me.partNumber & "' AND person = '" & Nz(Me.person, "EMPTY") & "'") > 0 Then errorMsg = "This team member is already present, let's not put them on twice"

'check if in department of project owner
projOwner = Not restrict(Environ("username"), TempVars!projectOwner)
deptSuper = Not restrict(Environ("username"), DLookup("Dept", "tblPermissions", "user = '" & Me.person & "'"), "Supervisor", True)

If Not (projOwner Or deptSuper) Then errorMsg = "You must be project owner or Supervisor in the member's dept to make this change"

If errorMsg <> "" Then
    Cancel = True
    Me.Undo
    
    Call snackBox("error", "Woops", errorMsg, Me.Parent.name)
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name & "_BeforeUpdate", Err.DESCRIPTION, Err.number)
End Sub

Private Sub remove_Click()
On Error GoTo Err_Handler
Me.Dirty = False

If IsNull(Me.recordId) Then
    MsgBox "This is an empty record, don't worry about deleting this.", vbInformation, "Can't do that"
    Exit Sub
End If

If Not restrict(Environ("username"), Nz(TempVars!projectOwner, "")) Then GoTo deleteThis
If userData("Dept") = DLookup("Dept", "tblPermissions", "user = '" & Me.person & "'") Then GoTo deleteThis

Call snackBox("error", "Woops", "You must be PE or manager in that dept to make that change", Me.Parent.name)
Exit Sub

deleteThis:
If MsgBox("Are you sure you want to delete this?", vbYesNo, "Please confirm") = vbYes Then
    dbExecute ("DELETE FROM tblPartTeam WHERE [recordId] = " & Me.recordId)
    Me.Requery
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub XFmeeting_Click()
On Error GoTo Err_Handler

Dim obj0App As Object
Dim objAppt As Object

Set obj0App = CreateObject("outlook.Application")
Set objAppt = obj0App.CreateItem(1)

With objAppt
    .RequiredAttendees = grabPartTeam(Me.partNumber, True)
    .subject = Me.partNumber & " Team Meeting"
    .ReminderMinutesBeforeStart = 5
    .Meetingstatus = 1
    .responserequested = True
    .display
End With

Set obj0App = Nothing
Set objAppt = Nothing

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
