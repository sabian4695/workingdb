Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private olApp As Object 'Outlook.Application

'-- containers for info on error logging
Private TotalErrors As Long
Private xLogFilePath As String
Private xErrorLogging As Boolean

'-- outlook constants
Public Enum OlItemType
    olAppointmentItem = 1
    olContactItem = 2
    olMailItem = 0
    olNoteItem = 5
    olTaskItem = 3
End Enum

Public Enum OlMailRecipientType
    olCC = 2
    olBCC = 3
    olTo = 1
End Enum

Public Enum OlImportance
    olImportanceHigh = 2
    olImportanceLow = 0
    olImportanceNormal = 1
End Enum

Public Enum OlSensitivity
    olConfidential = 3
    olNormal = 0
    olPersonal = 1
    olPrivate = 2
End Enum

Public Enum OlMeetingStatus
    olMeeting = 1
End Enum

Public Enum OlMeetingRecipientType
    olOptional = 2
    olOrganizer = 0
    olRequired = 1
    olResource = 3
End Enum

Public Enum OlBusyStatus
    olBusy = 2
    olFree = 0
    olOutOfOffice = 3
    olTentative = 1
End Enum

Public Enum OlInspectorClose
    olSave = 0
End Enum

Public Enum OlDefaultFolders
    olFolderCalendar = 9
    olFolderContacts = 10
    olFolderDrafts = 16
    olFolderInbox = 6
    olFolderJournal = 11
    olFolderNotes = 12
    olFolderOutbox = 4
    olFolderSentMail = 5
    olFolderTasks = 13
    olPublicFoldersAllPublicFolders = 18
End Enum

Private Sub Class_Initialize()
    Set olApp = CreateObject("Outlook.Application")
End Sub

Private Sub Class_Terminate()
    Set olApp = Nothing
End Sub

Public Function AddFolder(name As String, Optional ParentFolder As Object, _
    Optional FolderType As OlDefaultFolders = -999) As Object 'Outlook.Folder
    If FolderType = -999 Then
        ParentFolder.Folders.Add name
    Else
        ParentFolder.Folders.Add name, FolderType
    End If
    
End Function

Public Function GetDefaultFolder(DefaultFolderType As OlDefaultFolders) As Object 'Outlook.Folder
    Set GetDefaultFolder = Me.OutlookApplication.GetNamespace("MAPI").GetDefaultFolder(DefaultFolderType)
End Function

Public Function GetFolderFromPath(PathString As String) As Object 'Outlook.Folder
    Dim PathArray As Variant
    Dim xFolder As Object 'Outlook.Folder
    Dim counter As Long
    
    On Error GoTo ErrHandler
    
    If Left(PathString, 2) = "\\" Then PathString = Mid(PathString, 3)
    PathArray = Split(PathString, "\\")
    
    Set xFolder = Me.OutlookApplication.Session.Folders.ITEM(PathArray(0))
    
    For counter = 1 To UBound(PathArray)
        Set xFolder = xFolder.Folders.ITEM(PathArray(counter))
    Next
    
    Set GetFolderFromPath = xFolder
    Exit Function
    
ErrHandler:
    Set GetFolderFromPath = Nothing
End Function

Public Function CreateMailItem(sendTo As Variant, Optional CC As Variant = "", _
    Optional BCC As Variant = "", Optional subject As String = "", _
    Optional body As String = "", Optional htmlBody As String = "", _
    Optional Attachments As Variant, Optional Importance As OlImportance = olImportanceNormal, _
    Optional Categories As String = "", Optional DeferredDeliveryTime As Date = #1/1/1950#, _
    Optional DeleteAfterSubmit As Boolean = False, Optional FlagRequest As String = "", _
    Optional ReadReceiptRequested As Boolean = False, Optional Sensitivity As OlSensitivity = olNormal, _
    Optional SaveSentMessageFolder As Variant = "", Optional CloseRightAway As Boolean = False, _
    Optional EnableReply As Boolean = True, Optional EnableReplyAll As Boolean = True, _
    Optional EnableForward As Boolean = True, Optional ReplyRecipients As Variant = "", _
    Optional OtherProperties As Variant = "", Optional tag As String = "") As Boolean

    Dim olMsg As Object                 'Outlook.MailItem
    Dim olRecip As Object               'Outlook.Recipient
    Dim counter As Long
    Dim CurrentProperty As String
    Dim ItemProp As Object              'Outlook.ItemProperty
    Dim SaveToFolder As Object          'Outlook.Folder
    
    On Error GoTo ErrHandler
    
    CurrentProperty = "CreateItem"
    Set olMsg = olApp.CreateItem(olMailItem)
    
    With olMsg
        
'-- for SendTo, CC, and BCC, if they are arrays, process each element of array through the Recipients
'-- collection.  If not, then if Len > 0 then pass in the string values via To, CC, and BCC
        CurrentProperty = "To"
        If IsArray(sendTo) Then
            For counter = LBound(sendTo) To UBound(sendTo)
                Set olRecip = .Recipients.Add(sendTo(counter))
                olRecip.type = olTo
            Next
        Else
            If sendTo <> "" Then .To = sendTo
        End If
        CurrentProperty = "CC"
        If IsArray(CC) Then
            For counter = LBound(CC) To UBound(CC)
                Set olRecip = .Recipients.Add(CC(counter))
                olRecip.type = olCC
            Next
        Else
            If CC <> "" Then .CC = CC
        End If
        CurrentProperty = "BCC"
        If IsArray(BCC) Then
            For counter = LBound(BCC) To UBound(BCC)
                Set olRecip = .Recipients.Add(BCC(counter))
                olRecip.type = olBCC
            Next
        Else
            If BCC <> "" Then .BCC = BCC
        End If
        
'-- set ReplyRecipients
'-- if a 1-D array, adds each element in array
'-- if string, adds from the string
        CurrentProperty = "ReplyRecipients"
        If IsArray(ReplyRecipients) Then
            For counter = LBound(ReplyRecipients) To UBound(ReplyRecipients)
                Set olRecip = .ReplyRecipients.Add(ReplyRecipients(counter))
            Next
        Else
            If ReplyRecipients <> "" Then .ReplyRecipients.Add ReplyRecipients
        End If
        
'-- standard field
        CurrentProperty = "Subject"
        .subject = subject
        
'-- if both Body and HTMLBody are given, Body wins
        CurrentProperty = "Body"
        If body <> "" Then .body = body
        CurrentProperty = "HTMLBody"
        If htmlBody <> "" And body = "" Then .htmlBody = htmlBody
        
'-- process attachments.  For multiple files, use an array.  For a single file, use a string
        CurrentProperty = "Attachments"
        If IsArray(Attachments) Then
            For counter = LBound(Attachments) To UBound(Attachments)
                .Attachments.Add Attachments(counter)
            Next
        Else
            If Not IsMissing(Attachments) Then
                If Attachments <> "" Then .Attachments.Add Attachments
            End If
        End If
        
'-- standard
        CurrentProperty = "Importance"
        .Importance = Importance
        CurrentProperty = "Categories"
        If Categories <> "" Then .Categories = Categories
        CurrentProperty = "DeferredDeliveryTime"
        If DeferredDeliveryTime >= DateAdd("n", 2, Now) Then .DeferredDeliveryTime = DeferredDeliveryTime
        CurrentProperty = "DeleteAfterSubmit"
        .DeleteAfterSubmit = DeleteAfterSubmit
        
'-- added in Outlook 2007.  By checking the Outlook version we avoid a potential error
        If val(olApp.Version) >= 12 Then
            CurrentProperty = "FlagRequest"
            If FlagRequest <> "" Then .FlagRequest = FlagRequest
        End If
        
'-- standard
        CurrentProperty = "ReadReceiptRequested"
        .ReadReceiptRequested = ReadReceiptRequested
        CurrentProperty = "Sensitivity"
        .Sensitivity = Sensitivity

        CurrentProperty = "SaveSentMessageFolder"
        If IsObject(SaveSentMessageFolder) Then
            If Not SaveSentMessageFolder Is Nothing Then Set .SaveSentMessageFolder = SaveSentMessageFolder
        Else
            If SaveSentMessageFolder <> "" Then
                Set SaveToFolder = Me.GetFolderFromPath(CStr(SaveSentMessageFolder))
                If SaveToFolder Is Nothing Then Set SaveToFolder = Me.AddFolderFromPath(CStr(SaveSentMessageFolder))
                Set .SaveSentMessageFolder = SaveToFolder
            End If
        End If
        
'-- keep in mind that these do not apply outside your organization, and can be reversed!
        .Actions("Reply").Enabled = EnableReply
        .Actions("Reply to All").Enabled = EnableReplyAll
        .Actions("Forward").Enabled = EnableForward
        
' process OtherProperties, if applicable.  If argument value is not an array, or is an array
' with just one dimension, this will throw an error
        CurrentProperty = "OtherProperties"
        If IsArray(OtherProperties) Then
            For counter = LBound(OtherProperties, 1) To UBound(OtherProperties, 1)
                Set ItemProp = .ItemProperties.ITEM(OtherProperties(counter, LBound(OtherProperties, 2)))
                ItemProp.Value = OtherProperties(counter, UBound(OtherProperties, 2))
            Next
        End If
        
'-- determine whether to send or display
        If CloseRightAway Then
            CurrentProperty = "Send"
            .send
        Else
            CurrentProperty = "Display"
            .display
        End If
    End With
    
    CreateMailItem = True
    
    GoTo cleanUp
    
ErrHandler:
    CreateMailItem = False
cleanUp:
    Set olMsg = Nothing
End Function

Public Function GetSubFolder(UsingFolder As Object, index As Variant) As Object 'Outlook.Folder
    On Error Resume Next
    
    Set GetSubFolder = UsingFolder.Folders(index)
    
    If Err <> 0 Then
        Err.clear
        Set GetSubFolder = Nothing
    End If
    
    On Error GoTo 0
End Function

Public Function AddFolderFromPath(PathString As String, _
    Optional FolderType As OlDefaultFolders = -999) As Object 'Outlook.Folder
' Creates a new Outlook folder placed according to the indicated path, and returns that new
' folder
    Dim PathArray As Variant
    Dim FolderName As String
    Dim xFolder As Object   'Outlook.Folder
    Dim yFolder As Object   'Outlook.Folder
    Dim counter As Long
    
    If Left(PathString, 2) = "\\" Then PathString = Mid(PathString, 3)
    PathArray = Split(PathString, "\\")
    
    FolderName = PathArray(counter)
    Set xFolder = Me.GetFolderFromPath(FolderName)
    If xFolder Is Nothing Then
        If FolderType = -999 Then
            Set xFolder = Me.OutlookApplication.Session.Folders.Add(FolderName)
        Else
            Set xFolder = Me.OutlookApplication.Session.Folders.Add(FolderName, FolderType)
        End If
    End If
    
    For counter = 1 To UBound(PathArray)
        FolderName = PathArray(counter)
        Set yFolder = Me.GetSubFolder(xFolder, FolderName)
        If yFolder Is Nothing Then
            If FolderType = -999 Then
                Set yFolder = xFolder.Folders.Add(FolderName)
            Else
                Set xFolder = xFolder.Folders.Add(FolderName, FolderType)
            End If
        End If
        Set xFolder = yFolder
    Next
    
    Set AddFolderFromPath = xFolder
    
End Function
    
Function CreateAppointmentItem(StartAt As Date, Optional duration As Long = 30, _
    Optional EndAt As Date = #1/1/1950#, Optional RequiredAttendees As Variant = "", _
    Optional OptionalAttendees As Variant = "", Optional subject As String = "", _
    Optional body As String = "", Optional location As String = "", _
    Optional AllDayEvent As Boolean = False, Optional Attachments As Variant = "", _
    Optional BusyStatus As OlBusyStatus = olBusy, Optional Categories As String = "", _
    Optional Importance As OlImportance = olImportanceNormal, Optional Organizer As Variant = "", _
    Optional ReminderMinutesBeforeStart As Long = 15, Optional ReminderSet As Boolean = True, _
    Optional Resources As Variant = "", Optional Sensitivity As OlSensitivity = olNormal, _
    Optional tag As String = "", Optional CloseRightAway As Boolean = True, _
    Optional OtherProperties As Variant = "", Optional SaveToFolder As Variant = "") As Boolean

    Dim olAppt As Object            'Outlook.AppointmentItem
    Dim counter As Long
    Dim CurrentProperty As String
    Dim EndFromDuration As Date
    Dim olRecip As Object           'Outlook.Recipient
    Dim ItemProp As Object          'Outlook.ItemProperty
    Dim TestFolder As Object        'Outlook.Folder
    
    On Error GoTo ErrHandler
    
    CurrentProperty = "CreateItem"
    Set olAppt = olApp.CreateItem(olAppointmentItem)
    With olAppt
        
'-- if there are attendees, make this a meeting
        CurrentProperty = "MeetingStatus"
        If IsArray(RequiredAttendees) Then
            .Meetingstatus = olMeeting
        ElseIf RequiredAttendees <> "" Then
            .Meetingstatus = olMeeting
        ElseIf IsArray(OptionalAttendees) Then
            .Meetingstatus = olMeeting
        ElseIf OptionalAttendees <> "" Then
            .Meetingstatus = olMeeting
        ElseIf IsArray(Organizer) Then
            .Meetingstatus = olMeeting
        ElseIf Organizer <> "" Then
            .Meetingstatus = olMeeting
        ElseIf IsArray(Resources) Then
            .Meetingstatus = olMeeting
        ElseIf Resources <> "" Then
            .Meetingstatus = olMeeting
        End If
        
'-- standard field
        CurrentProperty = "Start"
        If StartAt >= Date Then .start = StartAt
        
        CurrentProperty = "End"
        EndFromDuration = DateAdd("n", duration, StartAt)
        If EndFromDuration >= EndAt Then
            .duration = duration
        Else
            .End = EndAt
        End If
        
'-- add RequiredAttendees, OptionalAttendees, Resources, and Organizer.
'-- may come in as arrays or strings
        CurrentProperty = "RequiredAttendees"
        If IsArray(RequiredAttendees) Then
            For counter = LBound(RequiredAttendees) To UBound(RequiredAttendees)
                Set olRecip = .Recipients.Add(RequiredAttendees(counter))
                olRecip.type = olRequired
            Next
        Else
            If RequiredAttendees <> "" Then
                Set olRecip = .Recipients.Add(RequiredAttendees)
                olRecip.type = olRequired
            End If
        End If
        CurrentProperty = "OptionalAttendees"
        If IsArray(OptionalAttendees) Then
            For counter = LBound(OptionalAttendees) To UBound(OptionalAttendees)
                Set olRecip = .Recipients.Add(OptionalAttendees(counter))
                olRecip.type = olOptional
            Next
        Else
            If OptionalAttendees <> "" Then
                Set olRecip = .Recipients.Add(OptionalAttendees)
                olRecip.type = olOptional
            End If
        End If
        CurrentProperty = "Resources"
        If IsArray(Resources) Then
            For counter = LBound(Resources) To UBound(Resources)
                Set olRecip = .Recipients.Add(Resources(counter))
                olRecip.type = olResource
            Next
        Else
            If Resources <> "" Then
                Set olRecip = .Recipients.Add(Resources)
                olRecip.type = olResource
            End If
        End If
        CurrentProperty = "Organizer"
        If IsArray(Organizer) Then
            For counter = LBound(Organizer) To UBound(Organizer)
                Set olRecip = .Recipients.Add(Organizer(counter))
                olRecip.type = olOrganizer
            Next
        Else
            If Organizer <> "" Then
                Set olRecip = .Recipients.Add(Organizer)
                olRecip.type = olOrganizer
            End If
        End If
        
'-- standard fields
        CurrentProperty = "Subject"
        .subject = subject
        CurrentProperty = "Body"
        If body <> "" Then .body = body
        CurrentProperty = "Location"
        If location <> "" Then .location = location
        CurrentProperty = "AllDayEvent"
        .AllDayEvent = AllDayEvent
        
'-- process attachments.  For multiple files, use an array.  For a single file, use a string
        CurrentProperty = "Attachments"
        If IsArray(Attachments) Then
            For counter = LBound(Attachments) To UBound(Attachments)
                .Attachments.Add Attachments(counter)
            Next
        Else
            If Attachments <> "" Then .Attachments.Add Attachments
        End If
        
'-- standard fields
        CurrentProperty = "BusyStatus"
        .BusyStatus = BusyStatus
        CurrentProperty = "Categories"
        If Categories <> "" Then .Categories = Categories
        CurrentProperty = "Importance"
        .Importance = Importance
        CurrentProperty = "ReminderSet"
        If ReminderSet Then
            .ReminderSet = True
            If ReminderMinutesBeforeStart < 0 Then ReminderMinutesBeforeStart = 0
            .ReminderMinutesBeforeStart = ReminderMinutesBeforeStart
        Else
            .ReminderSet = False
            .ReminderMinutesBeforeStart = 0
        End If
        CurrentProperty = "Sensitivity"
        .Sensitivity = Sensitivity
        
'-- process OtherProperties, if applicable.  If argument value is not an array,
'-- or is an array with just one dimension, this will throw an error
        CurrentProperty = "OtherProperties"
        If IsArray(OtherProperties) Then
            For counter = LBound(OtherProperties, 1) To UBound(OtherProperties, 1)
                Set ItemProp = .ItemProperties.ITEM(OtherProperties(counter, LBound(OtherProperties, 2)))
                ItemProp.Value = OtherProperties(counter, UBound(OtherProperties, 2))
            Next
        End If

        CurrentProperty = "SaveToFolder"
        If IsObject(SaveToFolder) Then
            If Not SaveToFolder Is Nothing Then
                .save
                .Move SaveToFolder
            End If
        Else
            If SaveToFolder <> "" Then
                Set TestFolder = Me.GetFolderFromPath(CStr(SaveToFolder))
                If TestFolder Is Nothing Then Set TestFolder = Me.AddFolderFromPath(CStr(SaveToFolder))
                .save
                .Move TestFolder
            End If
        End If
        
'-- if there are no attendees, then it is just an appointment, and the choice is Close/Display.
'-- if there are attendees, then it is a meeting request, and the choice is Send/Display.
        If .Recipients.count > 0 Then
            If CloseRightAway Then
                CurrentProperty = "Send"
                .send
            Else
                CurrentProperty = "Display"
                .display
            End If
        Else
            If CloseRightAway Then
                CurrentProperty = "Close"
                .CLOSE olSave
            Else
                CurrentProperty = "Display"
                .display
            End If
        End If
    End With
    
    CreateAppointmentItem = True
    
    GoTo cleanUp
    
ErrHandler:
    CreateAppointmentItem = False
cleanUp:
    Set olAppt = Nothing
End Function

Property Get OutlookApplication() As Object
    Set OutlookApplication = olApp
End Property