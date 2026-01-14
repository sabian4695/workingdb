Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Load()
On Error GoTo Err_Handler

TempVars.Add "loadAmount", 0
TempVars.Add "loadWd", 8160
TempVars.Add "wdbVersion", grabVersion()
Me.lblFrozen.Visible = False
Call setSplashLoading("Setting up app stuff...")
Me.lblVersion.Caption = TempVars!wdbVersion
Me.lblVersion.Visible = True

SizeAccess 280, 280
Me.Move -2600, -1000

bClone = False
Call logClick("Form_Load", Me.module.name)

Me.Picture = "\\data\mdbdata\WorkingDB\Pictures\Splash\splash" & randomNumber(0, DLookup("splashCount", "tblDBinfoBE", "ID = 1")) & ".png"
On Error Resume Next
Me.imgUser.Picture = "\\data\mdbdata\WorkingDB\Pictures\Avatars\" & Environ("username") & ".png"
On Error GoTo Err_Handler
Call setSplashLoading("Setting up app stuff...")

DoEvents
Form_frmSplash.SetFocus
DoEvents

Call setSplashLoading("Setting up app stuff...")
If CommandBars("Ribbon").Height > 100 Then CommandBars.ExecuteMso "MinimizeRibbon"
DoCmd.ShowToolbar "Ribbon", acToolbarNo

On Error Resume Next 'copy shortcut to databases folder and desktop
    Dim desktopLoc As String, databasesLoc As String, fso, desktopNCM
    desktopLoc = "\\homes\data\" & Environ("username") & "\Desktop\Working DB.lnk"
    desktopNCM = "\\ncm-fs2\homes\" & Environ("username") & "\Desktop\Working DB.lnk"
    databasesLoc = "C:\Users\" & Environ("username") & "\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Working DB.lnk"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Call fso.CopyFile("\\data\mdbdata\WorkingDB\Batch\Working DB.lnk", desktopLoc)
    Call fso.CopyFile("\\data\mdbdata\WorkingDB\Batch\Working DB.lnk", databasesLoc)
    Call fso.CopyFile("\\data\mdbdata\WorkingDB\build\workingdb_ghost\WorkingDB_ghost.accde", "C:\workingdb\WorkingDB_ghost.accde") 'copy newest GHOST DB
    openPath "\\data\mdbdata\WorkingDB\build\workingdb_commands\openGhost.vbs"
On Error GoTo Err_Handler

Call setSplashLoading("Doing some digging on you...")

Dim db As Database
Set db = CurrentDb()
Dim rs1 As Recordset, rsFiltered As Recordset, rsUserSettings As Recordset
Set rsUserSettings = db.OpenRecordset("tblUserSettings")
Set rs1 = db.OpenRecordset("tblPermissions")
rs1.filter = "[User] = '" & Environ("username") & "'"
Set rsFiltered = rs1.OpenRecordset
rsUserSettings.filter = "[username] = '" & Environ("username") & "'"
Set rsUserSettings = rsUserSettings.OpenRecordset

'---SETUP USER SETTINGS---
If rsUserSettings.RecordCount = 0 Then
    With rsUserSettings
        .addNew
            !userName = Environ("username")
        .Update
        .Bookmark = .lastModified
    End With
End If
'---

'---SET UP PERMISSIONS---
If rsFiltered.RecordCount = 0 Then 'Add new user!
    On Error Resume Next
    Dim rsEmployee As Recordset, firstN As String, lastN As String, emailAdd As String
    firstN = ""
    lastN = ""
    Set rsEmployee = db.OpenRecordset("SELECT FIRST_NAME, LAST_NAME, EMAIL_ADDRESS FROM APPS_XXCUS_USER_EMPLOYEES_V WHERE USER_NAME = '" & StrConv(Environ("username"), vbUpperCase) & "'")
    If rsEmployee.RecordCount <> 0 Then
        firstN = StrConv(rsEmployee!First_Name, vbProperCase)
        lastN = StrConv(rsEmployee!Last_Name, vbProperCase)
        emailAdd = rsEmployee!EMAIL_ADDRESS
    End If
    On Error GoTo Err_Handler
    
    'ask user to confirm names
    Dim fN, lN
    
firstNameLabel:
    fN = InputBox("Please confirm your first name", "Enter your First Name", firstN)
    If StrPtr(fN) = 0 Or fN = "" Then GoTo firstNameLabel
    
lastNameLabel:
    lN = InputBox("Please confirm your last name", "Enter your Last Name", lastN)
    If StrPtr(lN) = 0 Or lN = "" Then GoTo lastNameLabel
    
    rs1.filter = adFilterNone
    With rs1
        .addNew
            !User = Environ("username")
            !firstName = fN
            !lastName = lN
            !userEmail = emailAdd
            !designWOpermissions = 3
            !designWOid = 2
        .Update
        .Bookmark = .lastModified
    End With
    
    Dim initials As String
    initials = Left(fN, 1) & Left(lN, 1)
    
    Call getAvatar(Environ("username"), initials)
    'Call convertSVGtoPNG("\\data\mdbdata\WorkingDB\Pictures\Avatars\svg\" & Environ("username") & ".svg", "\\data\mdbdata\WorkingDB\Pictures\Avatars\" & Environ("username") & ".png")
    
    Me.imgUser.Picture = "\\data\mdbdata\WorkingDB\Pictures\Avatars\" & Environ("username") & ".png"
    
    rs1.filter = "[User] = '" & Environ("username") & "'"
    Set rsFiltered = rs1.OpenRecordset
    MsgBox "Thank you for opening the workingDB." & vbNewLine & "This was made to help you navigate through our file system and to help track your projects and workload." & _
        "It is connected to Oracle and has many functions." & vbNewLine & "Click the ""Help"" button in the TOP RIGHT CORNER of the screen to see some details of things you can do.", vbOKOnly, "Welcome"
End If

Call setSplashLoading("Doing some digging on you...")

TempVars.Add "dept", Nz(rsFiltered!dept, "")
TempVars.Add "org", Nz(rsFiltered!Org, 4)
TempVars.Add "smallScreen", Nz(rsUserSettings!smallScreenMode, "False")

'THEME
Dim rsTheme As Recordset

If Nz(rsUserSettings!themeId, 0) <> 0 Then
    Set rsTheme = db.OpenRecordset("SELECT * FROM tblTheme WHERE recordId = " & rsUserSettings!themeId)
    
    If rsTheme!darkMode Then
        TempVars.Add "themeMode", "Dark"
    Else
        TempVars.Add "themeMode", "Light"
    End If
    
    TempVars.Add "themePrimary", CStr(rsTheme!primaryColor)
    TempVars.Add "themeSecondary", CStr(rsTheme!secondaryColor)
    TempVars.Add "themeColorLevels", CStr(rsTheme!colorLevels)
    
    rsTheme.CLOSE
    Set rsTheme = Nothing
End If

Call setSplashLoading("Running daily checks...")
Call checkForFirstTimeRun

Call setSplashLoading("Writing a funny joke...")
Call grabJoke

Call setSplashLoading("Building dashboard...")

DoCmd.OpenForm "DASHBOARD"
Form_DASHBOARD.Visible = False
Call setSplashLoading("Wrapping up...")
DoCmd.CLOSE acForm, "frmSplash"
DoEvents
Call maximizeAccess
Form_DASHBOARD.Visible = True
DoCmd.Maximize

On Error Resume Next
rsUserSettings.CLOSE: Set rsUserSettings = Nothing
rsFiltered.CLOSE: Set rsFiltered = Nothing
rs1.CLOSE: Set rs1 = Nothing
Set db = Nothing
    
Exit Sub
Err_Handler:
    Call handleError(Me.name, "Form_Load", Err.DESCRIPTION, Err.number)
End Sub

Function grabJoke()
On Error GoTo Err_Handler

Dim Joke As String
Joke = Nz(DLookup("[factText]", "tblFacts", "[factDate] = #" & Date & "#"))

If Nz(Joke) = "" Then
    Dim X, Y, z, C
getAnotherOne:
    C = getAPI("https://icanhazdadjoke.com/", "Accept: text/plain", "User-Agent: Jacob Brown  (jbrow4@gmail.com")
    Y = Split(C, "<p class=" & Chr(34) & "subtitle" & Chr(34) & ">")(1)
    Joke = Split(Y, "</p>")(0)
    
    Dim jokeArr, ITEM, finalJoke As String
    jokeArr = Split(Joke, "</br>")
    finalJoke = ""
    
    For Each ITEM In jokeArr
        If Trim(ITEM) = "" Then GoTo nextItem
        If finalJoke = "" Then
            finalJoke = ITEM
        Else
            finalJoke = finalJoke & vbCrLf & Trim(ITEM)
        End If
nextItem:
    Next ITEM
    
    If (UBound(Split(finalJoke, vbCrLf)) + 1) > 2 Then GoTo getAnotherOne
    dbExecute "INSERT INTO tblFacts(factDate,factText,userName,Created) VALUES ('" & Date & "','" & StrQuoteReplace(Joke) & "','" & Environ("username") & "','" & Now() & "')"
End If

TempVars.Add "joke", Joke

Exit Function
Err_Handler:
    Call handleError(Me.name, "grabJoke", Err.DESCRIPTION, Err.number)
End Function
