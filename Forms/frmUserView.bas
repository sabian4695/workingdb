Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub apiOption_AfterUpdate()
On Error GoTo Err_Handler

If Me.apiOption = "initials" Then
    Dim fi As String, li As String
    
    fi = Left(userData("firstName"), 1)
    li = Left(userData("lastName"), 1)
    
    Call getAvatar(Environ("username"), fi & li)
Else
    Call getAvatarAPI(Me.apiOption)
End If

Me.userPic.Picture = "\\data\mdbdata\WorkingDB\Pictures\Avatars\" & Environ("username") & ".png"

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub btnGenerateSummary_Click()
On Error GoTo Err_Handler

If grabSummaryInfo(Environ("username")) Then
    MsgBox "Email arriving soon! (if you have open items)", vbInformation, "Success"
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub donotclick_Click()
On Error GoTo Err_Handler

Select Case Me.donotclick.Caption
    Case "Do Not Click"
        Me.donotclick.Caption = "Seriously, don't"
    Case "Seriously, don't"
        Me.donotclick.Caption = "For real? Stop"
    Case "For real? Stop"
        Me.donotclick.Caption = "One more, I dare you"
    Case "One more, I dare you"
        Me.donotclick.Caption = "Maybe double click?"
End Select
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub donotclick_DblClick(Cancel As Integer)
On Error GoTo Err_Handler
If Me.donotclick.Caption = "Maybe double click?" Then
    Me.donotclick.Caption = "bye"
    Call openPath(mainFolder(Me.ActiveControl.name))
    DoCmd.CLOSE
End If
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Me.donotclick.Caption = "Do Not Click"
DoCmd.applyFilter , "[tblPermissions].User = '" & Environ("username") & "'"
Me.lblVersion.Caption = Nz(TempVars!wdbVersion, "")

Me.userPic.Picture = "\\data\mdbdata\WorkingDB\Pictures\Avatars\" & Environ("username") & ".png"

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Err_Handler

Form_DASHBOARD.userPic.Picture = "\\data\mdbdata\WorkingDB\Pictures\Avatars\" & Environ("username") & ".png"

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub smallScreenMode_Click()
On Error GoTo Err_Handler

If TempVars!smallScreen = "True" Then
    TempVars.Add "smallScreen", "False"
    Form_DASHBOARD.smallScreenMode (False)
Else
    TempVars.Add "smallScreen", "True"
    Form_DASHBOARD.smallScreenMode (True)
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
