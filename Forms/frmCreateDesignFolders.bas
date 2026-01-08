Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub assembly_Click()
On Error GoTo Err_Handler

Me.Detail.Visible = Me.assembly
Me.lbl1.Visible = Me.assembly
Me.lbl2.Visible = Me.assembly
Me.lbl3.Visible = Me.assembly
Me.lbl4.Visible = Me.assembly
Me.lbl5.Visible = Me.assembly

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Function createDocHisFolder(partNum As String, addFullSetup As Boolean) As String
On Error GoTo Err_Handler

createDocHisFolder = ""

Dim thousZeros, hundZeros, mainPath, prtFilePath As String, packetFolder As String

If partNum Like "[A-Z][A-Z]##[A-Z]##[A-Z]" Or partNum Like "[A-Z][A-Z]##[A-Z]##" Or partNum Like "##[A-Z]##" Then
    '---NCM PART---
    If Not partNum Like "##[A-Z]##" Then partNum = Mid(partNum, 3, 5)
    mainPath = mainFolder("ncmDrawingMaster")
    prtFilePath = mainPath & Left(partNum, 3) & "00\" & partNum & "\"
    
    'i.e. PN 12A51
    If Not FolderExists(mainPath & Left(partNum, 3) & "00\") Then MkDir (mainPath & Left(partNum, 3) & "00\") 'check for 12A00
    If Not FolderExists(prtFilePath) Then MkDir (prtFilePath) 'create 12A51
    prtFilePath = prtFilePath & "Documents" 'final folder
Else
    '---NAM PART---
    thousZeros = Left(partNum, 2) & "000\"
    hundZeros = Left(partNum, 3) & "00\"
    mainPath = mainFolder("docHisSearch")
    prtFilePath = mainPath & thousZeros & hundZeros
    
    'i.e. PN 12345
    If Not FolderExists(mainPath & thousZeros) Then MkDir (mainPath & thousZeros) 'check for 12000
    If Not FolderExists(prtFilePath) Then MkDir (prtFilePath) 'check for 12300
    prtFilePath = prtFilePath & partNum 'final folder
End If

'first, do a quick check to make sure the folder doesn't already exist
'if it does exist, put that string as the doc his folder and move on

'Make final doc history folder
createDocHisFolder = prtFilePath
If Not FolderExists(prtFilePath) Then
    MkDir (prtFilePath)
Else
    'if for some reason the doc his folder already exists, set string and exit sub
    Exit Function
End If

If addFullSetup Then
    packetFolder = prtFilePath & "\Misc\Drawing Packets\" & partNum & "_00_DWG_PKT_1000_CHECK\"
    MkDir (prtFilePath & "\Misc\")
    MkDir (prtFilePath & "\Misc\Drawing Packets\")
    MkDir (packetFolder)
End If

Exit Function
Err_Handler:
    Call handleError(Me.name, "createDocHisFolder", Err.DESCRIPTION, Err.number)
End Function

Function makeAllDocumentHistory()
On Error GoTo Err_Handler

Dim createdMain As String
createdMain = createDocHisFolder(Me.partNumber, Me.includeFullFolderSetup1)

Dim createdChild As String
If Len(Me.childPartNumber) > 3 Then
    createdChild = createDocHisFolder(Me.childPartNumber, False)
    Call createShortcut(createdMain & "\" & Me.childPartNumber, createdChild, Nz(Me.cPartNumber_ext, "")) 'non master to master
    Call createShortcut(createdChild & "\" & Me.partNumber, createdMain, Nz(Me.partNumber_ext, "")) 'master to non master
End If

If Not Me.assembly Then Exit Function

Dim db As Database
Set db = CurrentDb()
Dim rs1 As Recordset, createdLoc As String, createdCompChild As String
Set rs1 = db.OpenRecordset("SELECT * from tblSessionVariables WHERE componentNumber is not null")

Do While Not rs1.EOF
    createdLoc = createDocHisFolder(rs1!componentNumber, rs1!componentFull)
    
    Call createShortcut(createdMain & "\" & rs1!componentNumber, createdLoc, Nz(rs1!componentCustomExt, "")) 'assembly to component
    Call createShortcut(createdLoc & "\" & Me.partNumber, createdMain, Nz(Me.partNumber_ext, "")) 'component to assembly
    
    If Len(rs1!componentNumberChild) > 3 Then
        createdCompChild = createDocHisFolder(rs1!componentNumberChild, False)
        Call createShortcut(createdLoc & "\" & rs1!componentNumberChild, createdCompChild, Nz(rs1!componentChildCustomExt, "")) 'master component -> non master component
        Call createShortcut(createdCompChild & "\" & rs1!componentNumber, createdLoc, Nz(rs1!componentCustomExt, "")) 'non master component -> master component
        Call createShortcut(createdCompChild & "\" & Me.partNumber, createdMain, Nz(Me.partNumber_ext, "")) 'non-master component -> master assembly
    End If
    
    rs1.MoveNext
Loop

rs1.CLOSE
Set rs1 = Nothing
Set db = Nothing

Exit Function
Err_Handler:
    Call handleError(Me.name, "makeAllDocumentHistory", Err.DESCRIPTION, Err.number)
End Function

Function makeAllModelV5()
On Error GoTo Err_Handler

Dim createdMain As String
createdMain = createModelV5fold(Me.partNumber, Me.includeFullFolderSetup1)

Dim createdChild As String
If Len(Me.childPartNumber) > 3 Then
    createdChild = createModelV5fold(Me.childPartNumber, False)
    Call createShortcut(createdMain & "\" & Me.childPartNumber, createdChild, Nz(Me.cPartNumber_ext, "")) 'non master to master
    Call createShortcut(createdChild & "\" & Me.partNumber, createdMain, Nz(Me.partNumber_ext, "")) 'master to non master
End If

If Not Me.assembly Then Exit Function

Dim db As Database
Set db = CurrentDb()
Dim rs1 As Recordset, createdLoc As String, createdCompChild As String
Set rs1 = db.OpenRecordset("SELECT * from tblSessionVariables WHERE componentNumber is not null")

Do While Not rs1.EOF
    createdLoc = createModelV5fold(rs1!componentNumber, rs1!componentFull)
    
    Call createShortcut(createdMain & "\" & rs1!componentNumber, createdLoc, Nz(rs1!componentCustomExt, "")) 'assembly to component
    Call createShortcut(createdLoc & "\" & Me.partNumber, createdMain, Nz(Me.partNumber_ext, "")) 'component to assembly
    
    If Len(rs1!componentNumberChild) > 3 Then
        createdCompChild = createModelV5fold(rs1!componentNumberChild, False)
        Call createShortcut(createdLoc & "\" & rs1!componentNumberChild, createdCompChild, Nz(rs1!componentChildCustomExt, "")) 'master component -> non master component
        Call createShortcut(createdCompChild & "\" & rs1!componentNumber, createdLoc, Nz(rs1!componentCustomExt, "")) 'non master component -> master component
        Call createShortcut(createdCompChild & "\" & Me.partNumber, createdMain, Nz(Me.partNumber_ext, "")) 'non-master component -> master assembly
    End If
    
    rs1.MoveNext
Loop

rs1.CLOSE
Set rs1 = Nothing
Set db = Nothing

Exit Function
Err_Handler:
    Call handleError(Me.name, "makeAllModelV5", Err.DESCRIPTION, Err.number)
End Function

Private Sub btnMakeAll_Click()
On Error GoTo Err_Handler

If Me.Dirty Then Me.Dirty = False

Select Case Me.txtWhatFolders
    Case "Document History"
        Call makeAllDocumentHistory
    Case "ModelV5"
        Call makeAllModelV5
    Case "Both"
        Call makeAllDocumentHistory
        Call makeAllModelV5
End Select

DoCmd.CLOSE

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Function createModelV5fold(partNum As String, addFullSetup As Boolean) As String
On Error GoTo Err_Handler

createModelV5fold = openModelV5Folder(partNum, False)

If createModelV5fold <> "" Then Exit Function

Dim thousZeros, hundZeros, catRev
Dim mainPath, prtFilePath, OEM, lastName, Path As String

If partNum Like "[A-Z][A-Z]##[A-Z]##[A-Z]" Or partNum Like "[A-Z][A-Z]##[A-Z]##" Or partNum Like "##[A-Z]##" Then
    If Not partNum Like "##[A-Z]##" Then partNum = Mid(partNum, 3, 5)
    mainPath = mainFolder("ncmDrawingMaster")
    prtFilePath = mainPath & Left(partNum, 3) & "00\" & partNum & "\"
    
    If Not FolderExists(mainPath & Left(partNum, 3) & "00\") Then MkDir (mainPath & Left(partNum, 3) & "00\")
    If Not FolderExists(prtFilePath) Then MkDir (prtFilePath)

    Path = prtFilePath & "CATIA"
Else
    thousZeros = Left(partNum, 2) & "000\"
    hundZeros = Left(partNum, 3) & "00\"
    mainPath = mainFolder("modelV5search")
    prtFilePath = mainPath & thousZeros & hundZeros
    
    If Not FolderExists(mainPath & thousZeros) Then MkDir (mainPath & thousZeros)
    If Not FolderExists(prtFilePath) Then MkDir (prtFilePath)
    
    catRev = Me.txtCatRev.Value
    OEM = Me.txtOEM.Value
    lastName = Me.txtName
    Path = prtFilePath & partNum & "_" & catRev & "_" & OEM & "_" & lastName
End If

MkDir (Path)

'Make final movelV5 folder

createModelV5fold = Path

If addFullSetup Then
    MkDir (Path & "\Misc\")
    MkDir (Path & "\Misc\Receive\")
    MkDir (Path & "\Misc\Receive\YYMMDD Customer KO -00\")
    MkDir (Path & "\Misc\Send\")
    MkDir (Path & "\Misc\Send\YYMMDD Tooling TKO -00\")
End If

createModelV5fold = Path

Exit Function
Err_Handler:
    Call handleError(Me.name, "createModelV5fold", Err.DESCRIPTION, Err.number)
End Function

Private Sub createFoldersBasic_Click()
On Error GoTo Err_Handler

Select Case Me.txtWhatFolders
    Case "Document History"
        Call createDocHisFolder(Me.partNumber, Me.includeFullFolderSetup)
    Case "ModelV5"
        Call createModelV5fold(Me.partNumber, Me.includeFullFolderSetup)
    Case "Both"
        Call createDocHisFolder(Me.partNumber, Me.includeFullFolderSetup)
        Call createModelV5fold(Me.partNumber, Me.includeFullFolderSetup)
End Select

DoCmd.CLOSE

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Dim db As Database
Set db = CurrentDb()
db.Execute "DELETE * from tblSessionVariables WHERE componentNumber is not null"
Me.Requery

Me.Detail.Visible = False
Me.lbl1.Visible = False
Me.lbl2.Visible = False
Me.lbl3.Visible = False
Me.btnMakeAll.Visible = False

Dim partNum
partNum = Form_DASHBOARD.partNumberSearch
Me.partNumber = partNum

Dim rs1 As Recordset
Set rs1 = db.OpenRecordset("Select lastName FROM tblPermissions where user = '" & Environ("username") & "'", dbOpenSnapshot)

Me.txtName.Value = StrConv(rs1!lastName, vbProperCase)

rs1.CLOSE
Set rs1 = Nothing
Set db = Nothing

Me.txtWhatFolders = "Document History"
If Form_DASHBOARD.ActiveControl.name <> "docHisSearch" Then Me.txtWhatFolders = "ModelV5"

Call txtWhatFolders_AfterUpdate

If partNum Like "[A-Z][A-Z]##[A-Z]##[A-Z]" Or partNum Like "[A-Z][A-Z]##[A-Z]##" Or partNum Like "##[A-Z]##" Then
    Me.lblCatRev.Visible = False
    Me.lblOEM.Visible = False
    Me.lblName.Visible = False
    Me.txtCatRev.Visible = False
    Me.txtOEM.Visible = False
    Me.txtName.Visible = False
End If

If Me.Dirty Then Me.Dirty = False

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub TabCtl192_Change()
On Error GoTo Err_Handler

Me.btnMakeAll.Visible = Me.TabCtl192.Value = 1

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub txtWhatFolders_AfterUpdate()
On Error GoTo Err_Handler

Dim trigger As Boolean
trigger = False
If Me.txtWhatFolders <> "Document History" Then trigger = True

Me.lblCatRev.Visible = trigger
Me.lblOEM.Visible = trigger
Me.lblName.Visible = trigger
Me.txtCatRev.Visible = trigger
Me.txtOEM.Visible = trigger
Me.txtName.Visible = trigger
Me.btn2.Visible = trigger
Me.bx2.Visible = trigger
Me.lbl2a.Visible = trigger

If trigger Then
    Me.btn3.Caption = "3"
Else
    Me.btn3.Caption = "2"
End If

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
