Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Load()
On Error GoTo Err_Handler

Call setTheme(Me)

Dim db As Database
Dim rs1 As Recordset
Set db = CurrentDb()
Set rs1 = db.OpenRecordset("tblPartShortcuts")

Do Until rs1.EOF
    rs1.Delete
    rs1.MoveNext
Loop

rs1.CLOSE
Set rs1 = Nothing

Dim partNum
partNum = Form_DASHBOARD.partNumberSearch

Dim thousZeros, hundZeros, i, mainPath
Dim direct As String
Dim alist() As String
thousZeros = Left(partNum, 2) & "000\"
hundZeros = Left(partNum, 3) & "00\"
i = 0

If partNum Like "[A-Z][A-Z]##[A-Z]##[A-Z]" Or partNum Like "[A-Z][A-Z]##[A-Z]##" Or partNum Like "##[A-Z]##" Then
    If Not partNum Like "##[A-Z]##" Then partNum = Mid(partNum, 3, 5)
    hundZeros = Left(partNum, 3) & "00\"
    mainPath = mainFolder("ncmDrawingMaster") & hundZeros & partNum & "\Documents\"
Else
    mainPath = mainFolder("docHisSearch") & thousZeros & hundZeros & partNum & "\"
End If

direct = Dir(mainPath)
If (Right(mainPath, 10) = "DOCUMENTS\") And direct = "" Then direct = Dir(Left(mainPath, Len(mainPath) - 10)) 'NCM shortcuts are placed in the main NCM folder, not within DOCUMENTS

Do While direct <> vbNullString
    If direct <> "." And direct <> ".." Then
        If direct Like "*.lnk" Then
            
            db.Execute "insert into tblPartShortcuts(linkTitle,linkAddress) values ('" & direct & "','" & mainPath & direct & "');"
            GoTo nosave
        End If
        ReDim Preserve alist(i)
        alist(i) = direct
        i = i + 1
    End If
nosave:
    direct = Dir
Loop

mainPath = mainPath & "shortcuts\"
direct = Dir(mainPath)

Do While direct <> vbNullString
    If direct <> "." And direct <> ".." Then
        If direct Like "*.lnk" Then
            
            db.Execute "insert into tblPartShortcuts(linkTitle,linkAddress) values ('" & direct & "','" & mainPath & direct & "');"
            GoTo nosave1
        End If
        ReDim Preserve alist(i)
        alist(i) = direct
        i = i + 1
    End If
nosave1:
    direct = Dir
Loop
Set db = Nothing
Me.Requery

Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Private Sub open_Click()
On Error GoTo Err_Handler
openPath (DLookup("[linkAddress]", "tblPartShortcuts", "[ID] = " & Me.ID))
Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub

Private Sub srch_Click()
On Error GoTo Err_Handler

Form_DASHBOARD.partNumberSearch = Left(DLookup("[linkTitle]", "tblPartShortcuts", "[ID] = " & Me.ID), 5)
Call Form_DASHBOARD.filterbyPN_Click

Exit Sub
Err_Handler:
    Call handleError(Me.name, Me.ActiveControl.name, Err.DESCRIPTION, Err.number)
End Sub
