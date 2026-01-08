Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim userType As String

Private Sub btnOpen_Click()
On Error GoTo Err_Handler

openPath (Me.Link)
Me.clickCount = Me.clickCount + 1
Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub clear_Click()
On Error GoTo Err_Handler
Dim w
Me.srchBox = ""
w = "obsolete = False"
If userType <> "Design" Then w = w & " AND Restricted = False"

Me.Form.filter = w
Me.Form.FilterOn = True
Me.srchBox.SetFocus
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub favorite_Click()
On Error GoTo Err_Handler

'check favorites bar. Are there any open?
'If there are any open, add this link to that one
'if not, msgbox to tell them to "reset" a favorites button, then try again

Dim db As Database
Set db = CurrentDb()
Dim rs1 As Recordset

Dim counter As Long
For counter = 1 To 10
    Set rs1 = db.OpenRecordset("SELECT * FROM tblUserButtons WHERE User = '" & Environ("username") & "' AND ButtonNum = '" & counter & "'")
    If rs1.RecordCount = 0 Then GoTo setFavorite
Next counter

MsgBox "Please reset a favorite button to set this link as a favorite", vbInformation, "No Open Spot Found"

Exit Sub
setFavorite:

Dim rs2 As Recordset
Set rs2 = db.OpenRecordset("tblUserButtons")
rs2.addNew

rs2!User = Environ("username")
rs2!ButtonNum = counter
rs2!Link = Me.Link
rs2!Caption = Me.LinkCaption

rs2.Update

rs2.CLOSE
Set rs2 = Nothing
Set db = Nothing

Call Form_DASHBOARD.loadUserBtns

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

userType = Nz(DLookup("[Dept]", "[tblPermissions]", "[User] = '" & Environ("username") & "'"))
Call clear_Click

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub openFirst_Click()
On Error GoTo Err_Handler

With Me.RecordsetClone
    .MoveFirst
    Me.Bookmark = .Bookmark
End With

openPath (Me.Link)

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub srchBox_Change()
On Error GoTo Err_Handler
DoCmd.Echo False

TempVars.Add "z", Me.srchBox.Text
Me.refresh

TempVars.Add "w", "obsolete = False"

If IsNull(TempVars!z) Then GoTo skipFilt
    
Select Case TempVars!z
    Case "SLB", "CNL", "LVG", "CUU", "NCM"
        TempVars.Add "w", TempVars!w & " AND [Caption] Like '*" & TempVars!z & "*' OR [Org] = '" & TempVars!z & "'"
    Case Else
        TempVars.Add "w", TempVars!w & " AND [Caption] Like '*" & TempVars!z & "*'"
End Select

If userType <> "Design" Then TempVars.Add "w", TempVars!w & " AND Restricted = False"

Me.Form.filter = TempVars!w
Me.Form.FilterOn = True
If Me.Recordset.RecordCount = 0 Then
    Me.Form.filter = "ID = 61"
    Me.Form.FilterOn = True
End If

skipFilt:
Me.srchBox = TempVars!z
Me.srchBox.SelStart = Me.srchBox.SelLength
DoCmd.Echo True
Me.Repaint

If Me.Form.FilterOn = False Then
    Me.Form.filter = "obsolete = False"
    Me.Form.FilterOn = True
    Me.srchBox.SelStart = Me.srchBox.SelLength
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblName_Click()
On Error GoTo Err_Handler

Me.lblOrg.Caption = "Org -"

Dim label() As String
label = labelCycle(Me.lblName.Caption, "Caption")

Me.lblName.Caption = label(0)
Me.OrderBy = label(1)
Me.OrderByOn = True
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub lblOrg_Click()
On Error GoTo Err_Handler

Me.lblName.Caption = "Caption -"

Dim label() As String
label = labelCycle(Me.lblOrg.Caption, "Org")

Me.lblOrg.Caption = label(0)
Me.OrderBy = label(1)
Me.OrderByOn = True
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
