Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private olApp As Object 'Outlook.Application

'-- used for error logging, if applicable
Private fso As Object
Private tsLog As Object

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
    
    Me.LogErrors False
    Set olApp = Nothing
    
End Sub

Public Function AddFolder(name As String, Optional ParentFolder As Object, _
    Optional FolderType As OlDefaultFolders = -999) As Object 'Outlook.Folder
'==============================================================================
'Creates a new Outlook folder, adding it as a subfolder to the designated ParentFolder.
'If a folder with the same name already exists at that node, this will throw an error.
'
' Note that not all OlDefaultFolders values are allowed here.  Per Outlook VBA help,
' only olFolderCalendar, olFolderContacts, olFolderDrafts, olFolderInbox, olFolderJournal,
' olFolderNotes, or olFolderTasks can be used
'
' If FolderType is omitted, subfolder is of same type as parent folder
'==============================================================================
    If FolderType = -999 Then
        ParentFolder.Folders.Add name
    Else
        ParentFolder.Folders.Add name, FolderType
    End If
    
End Function

Public Function AddFolderFromPath(PathString As String, _
    Optional FolderType As OlDefaultFolders = -999) As Object 'Outlook.Folder
'==============================================================================
' Creates a new Outlook folder placed according to the indicated path, and returns that new
' folder
'
' Along the way, if folders in the specified PathString do not exist, those folders are created
'
' If a folder with that path already exists, an error results
'
' Use two backslash characters to delimit nested folders.  Double backslash at the beginning
' for the top level folder is optional.  For example, these both get the same folder:
'
' \\Inbox - Joe Schmoe\\Customers\\Nifco Corp
' Inbox - Joe Schmoe\\Customers\\Nifco Corp
'
' Note that not all OlDefaultFolders values are allowed here.  Per Outlook VBA help,
' only olFolderCalendar, olFolderContacts, olFolderDrafts, olFolderInbox, olFolderJournal,
' olFolderNotes, or olFolderTasks can be used
'
' If FolderType is omitted, subfolder is of same type as parent folder
'==============================================================================
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

Property Get CountOfErrors() As Long

'-- returns number of errors.  Read only
    CountOfErrors = TotalErrors
    
End Property

Property Get ErrorLoggingEnabled() As Boolean

'-- returns whether error logging is currently enabled (True) or not.  Read only
    ErrorLoggingEnabled = xErrorLogging
    
End Property

Public Function GetDefaultFolder(DefaultFolderType As OlDefaultFolders) As Object 'Outlook.Folder
'-- returns one of the default folders
    
    Set GetDefaultFolder = Me.OutlookApplication.GetNamespace("MAPI").GetDefaultFolder(DefaultFolderType)
    
End Function

Public Function GetFolderFromPath(PathString As String) As Object 'Outlook.Folder
'==============================================================================
' Returns an Outlook folder found by traversing the indicated path, or Nothing if no
' such folder exists
'
' Use two backslash characters to delimit nested folders.  Double backslash at the beginning
' for the top level folder is optional.  For example, these both get the same folder:
'
' \\Inbox - Joe Schmoe\\Customers\\Nifco Corp
'   Inbox - Joe Schmoe\\Customers\\Nifco Corp
'==============================================================================
    Dim PathArray As Variant
    Dim xFolder As Object 'Outlook.Folder
    Dim counter As Long
    
    On Error GoTo ErrHandler
    
    If Left(PathString, 2) = "\\" Then PathString = Mid(PathString, 3)
    PathArray = Split(PathString, "\\")
    
    Set xFolder = Me.OutlookApplication.Session.Folders.item(PathArray(0))
    
    For counter = 1 To UBound(PathArray)
        Set xFolder = xFolder.Folders.item(PathArray(counter))
    Next
    
    Set GetFolderFromPath = xFolder
    Exit Function
    
ErrHandler:
    
    Set GetFolderFromPath = Nothing
    
End Function

Public Function GetParentFolder(UsingFolder As Object) As Object 'Outlook.Folder

'-- returns an Outlook folder, the parent of the folder passed in as UsingFolder
    Set GetParentFolder = UsingFolder.Parent
    
End Function

Public Function GetSubFolder(UsingFolder As Object, Index As Variant) As Object 'Outlook.Folder
'==============================================================================
' Returns an Outlook folder, the indicated subfolder of UsingFolder.  The Index may
' be a String (the name of the subfolder) or a Long
'
' If subfolder does not exist, returns Nothing
'==============================================================================
    On Error Resume Next
    
    Set GetSubFolder = UsingFolder.Folders(Index)
    
    If Err <> 0 Then
        Err.clear
        Set GetSubFolder = Nothing
    End If
    
    On Error GoTo 0
    
End Function

Public Sub LogErrors(Optional Enable As Boolean = True, Optional LogPath As String = "", _
    Optional Append As Boolean = True)
'==============================================================================
' Method for enabling/disabling error logging.  If nothing is passed for the LogPath,
' use a default path.  Error logging can either append to an existing file, if applicable,
' or write to a new file.  Using Append = False will force an over-write, if applicable
'==============================================================================
    Dim DefaultPath As String
    Dim SpecialFolders As Object
    
    Const ForAppending As Long = 8
    Const Headers As String = "DateTime,ItemType,Tag,Property,ErrDescr"
    
    If Not Enable Then
    
'-- close the log if it's open, set objects to Nothing, and update private variables
'-- used for Property Gets
        If Not tsLog Is Nothing Then
            tsLog.Close
            Set tsLog = Nothing
            Set fso = Nothing
        End If
        xLogFilePath = ""
        xErrorLogging = False
    Else
        
'-- update private variable
        xErrorLogging = True
        
'-- resolve path name.  If path is not specified, then write to user's "My Documents"
'-- folder
        If Trim(LogPath) = "" Then
            Set SpecialFolders = CreateObject("WScript.Shell").SpecialFolders
            DefaultPath = SpecialFolders("mydocuments") & "\clsOutlookCreateItem Log.txt"
            Set SpecialFolders = Nothing
            xLogFilePath = DefaultPath
        Else
            xLogFilePath = LogPath
        End If
        
'-- close existing log, if applicable
        If Not tsLog Is Nothing Then tsLog.Close
        If fso Is Nothing Then Set fso = CreateObject("Scripting.FileSystemObject")
        If Append And fso.FileExists(xLogFilePath) Then
            
'-- if we want to append, and that file already exists, use OpenTextFile
            Set tsLog = fso.OpenTextFile(xLogFilePath, ForAppending)
        Else
        
'-- otherwise, create new file, overwriting if need be
            Set tsLog = fso.CreateTextFile(xLogFilePath, True)
            tsLog.WriteLine Headers
        End If
    End If
        
End Sub

Property Get LogFilePath() As String

'-- returns current path for log file.  Read only
'-- to set the path for the log file, use LogErrors
    LogFilePath = xLogFilePath
    
End Property

Public Function CreateContactItem(Optional lastName As String = "", Optional firstName As String = "", _
    Optional MiddleName As String = "", Optional Title As String = "", _
    Optional Suffix As String = "", Optional fullName As String = "", _
    Optional FileAs As String = "", Optional Email1Address As String = "", _
    Optional Email1DisplayName As String = "", Optional Email1AddressType As String = "", _
    Optional CompanyName As String = "", Optional BusinessTelephoneNumber As String = "", _
    Optional BusinessFaxNumber As String = "", Optional BusinessAddressCity As String = "", _
    Optional BusinessAddressCountry As String, Optional BusinessAddressPostalCode As String = "", _
    Optional BusinessAddressPostOfficeBox As String = "", Optional BusinessAddressState As String, _
    Optional BusinessAddressStreet As String = "", Optional BusinessAddress As String = "", _
    Optional HomeTelephoneNumber As String = "", Optional HomeFaxNumber As String = "", _
    Optional HomeAddressCity As String = "", Optional HomeAddressCountry As String, _
    Optional HomeAddressPostalCode As String = "", Optional HomeAddressPostOfficeBox As String = "", _
    Optional HomeAddressState As String, Optional HomeAddressStreet As String = "", _
    Optional HomeAddress As String = "", Optional MobileTelephoneNumber As String = "", _
    Optional Categories As String = "", Optional AddPicture As String = "", _
    Optional Importance As OlImportance = olImportanceNormal, Optional Sensitivity As OlSensitivity = olNormal, _
    Optional Journal As Boolean = False, Optional Attachments As Variant = "", _
    Optional OtherProperties As Variant = "", Optional tag As String = "", _
    Optional CloseRightAway As Boolean = True, Optional SaveToFolder As Variant = "") As Boolean
'===============================================================================
' Method for creating a contact item.  All arguments are optional
'
' Returns True for successful completion, and False otherwise.  If unsuccessful, no message is sent,
' and if logging is enabled at the class level, the procedure writes details to the log file
'
' Most arguments correspond to ContactItem properties in Outlook, so refer to Outlook VBA help
' file for documentation on them.  The exceptions:
'
' - AddPicture:      AddPicture is a ContactItem method, not a property.  Adds a picture to the
'                    ContactItem.  Argument value is a string specifying the path to the
'                    picture file
'
' - Attachments:     Can be either a 1-dimensional array of attachment paths, or a string.  If an
'                    array, use each member of the array to list the file path.  If a string with
'                    Len > 0, add the file given the path
'
' - Tag:             Allows passing a user-defined value that is not used in the MailItem, but may
'                    be helpful for troubleshooting errors.  For example, you might pass in a row
'                    number from Excel or a primary key value from Access
'
' - CloseRightAway:  Determines whether to send the item (True), or open the item in an Inspector
'
' - OtherProperties: Allows user to set other item properties not already handled in the basic
'                    class.  If used, pass as a two-dimensional array, with the first dimension
'                    indicating the property name, and the second indicating the value to use
'
' - SaveToFolder:    If you wish to save the new item to a folder other than its default folder, use
'                    this argument.  Argument can be an object representing the Outlook folder, or it
'                    can be a string.  If a string, it identifies the folder by path.  Use two backslash
'                    characters to delimit nested folders.  Double backslash at the beginning for the
'                    top level folder is optional.  For example, these both get the same folder:
'                    -------
'                    \\Inbox - Joe Schmoe\\Customers\\Nifco Corp
'                      Inbox - Joe Schmoe\\Customers\\Nifco Corp
'                    -------
'                    If the folder does not already exist, the folder will be created (as will
'                    any "missing" levels in the indicated path)
'===============================================================================
    Dim olContact As Object     'Outlook.ContactItem
    Dim counter As Long
    Dim CurrentProperty As String
    Dim ItemProp As Object      'Outlook.ItemProperty
    Dim TestFolder As Object    'Outlook.Folder
    
    On Error GoTo ErrHandler
    
    CurrentProperty = "CreateItem"
    Set olContact = olApp.CreateItem(olContactItem)
    
    With olContact
    
'-- standard properties
        CurrentProperty = "LastName"
        If lastName <> "" Then .lastName = lastName
        CurrentProperty = "FirstName"
        If firstName <> "" Then .firstName = firstName
        CurrentProperty = "MiddleName"
        If MiddleName <> "" Then .MiddleName = MiddleName
        CurrentProperty = "Title"
        If Title <> "" Then .Title = Title
        CurrentProperty = "Suffix"
        If Suffix <> "" Then .Suffix = Suffix
        CurrentProperty = "FullName"
        If fullName <> "" Then .fullName = fullName
        CurrentProperty = "FileAs"
        If FileAs <> "" Then .FileAs = FileAs
        CurrentProperty = "Email1Address"
        If Email1Address <> "" Then .Email1Address = Email1Address
        CurrentProperty = "Email1DisplayName"
        If Email1DisplayName <> "" Then .Email1DisplayName = Email1DisplayName
        CurrentProperty = "Email1AddressType"
        If Email1AddressType <> "" Then .Email1AddressType = Email1AddressType
        CurrentProperty = "CompanyName"
        If CompanyName <> "" Then .CompanyName = CompanyName
        CurrentProperty = "BusinessTelephoneNumber"
        If BusinessTelephoneNumber <> "" Then .BusinessTelephoneNumber = BusinessTelephoneNumber
        CurrentProperty = "BusinessFaxNumber"
        If BusinessFaxNumber <> "" Then .BusinessFaxNumber = BusinessFaxNumber
        CurrentProperty = "BusinessAddressCity"
        If BusinessAddressCity <> "" Then .BusinessAddressCity = BusinessAddressCity
        CurrentProperty = "BusinessAddressCountry"
        If BusinessAddressCountry <> "" Then .BusinessAddressCountry = BusinessAddressCountry
        CurrentProperty = "BusinessAddressPostalCode"
        If BusinessAddressPostalCode <> "" Then .BusinessAddressPostalCode = BusinessAddressPostalCode
        CurrentProperty = "BusinessAddressPostOfficeBox"
        If BusinessAddressPostOfficeBox <> "" Then .BusinessAddressPostOfficeBox = BusinessAddressPostOfficeBox
        CurrentProperty = "BusinessAddressState"
        If BusinessAddressState <> "" Then .BusinessAddressState = BusinessAddressState
        CurrentProperty = "BusinessAddressStreet"
        If BusinessAddressStreet <> "" Then .BusinessAddressStreet = BusinessAddressStreet
        CurrentProperty = "BusinessAddress"
        If BusinessAddress <> "" Then .BusinessAddress = BusinessAddress
        CurrentProperty = "HomeTelephoneNumber"
        If HomeTelephoneNumber <> "" Then .HomeTelephoneNumber = HomeTelephoneNumber
        CurrentProperty = "HomeFaxNumber"
        If HomeFaxNumber <> "" Then .HomeFaxNumber = HomeFaxNumber
        CurrentProperty = "HomeAddressCity"
        If HomeAddressCity <> "" Then .HomeAddressCity = HomeAddressCity
        CurrentProperty = "HomeAddressCountry"
        If HomeAddressCountry <> "" Then .HomeAddressCountry = HomeAddressCountry
        CurrentProperty = "HomeAddressPostalCode"
        If HomeAddressPostalCode <> "" Then .HomeAddressPostalCode = HomeAddressPostalCode
        CurrentProperty = "HomeAddressPostOfficeBox"
        If HomeAddressPostOfficeBox <> "" Then .HomeAddressPostOfficeBox = HomeAddressPostOfficeBox
        CurrentProperty = "HomeAddressState"
        If HomeAddressState <> "" Then .HomeAddressState = HomeAddressState
        CurrentProperty = "HomeAddressStreet"
        If HomeAddressStreet <> "" Then .HomeAddressStreet = HomeAddressStreet
        CurrentProperty = "HomeAddress"
        If HomeAddress <> "" Then .HomeAddress = HomeAddress
        CurrentProperty = "MobileTelephoneNumber"
        If MobileTelephoneNumber <> "" Then .MobileTelephoneNumber = MobileTelephoneNumber
        CurrentProperty = "Categories"
        If Categories <> "" Then .Categories = Categories
        CurrentProperty = "Importance"
        .Importance = Importance
        CurrentProperty = "Sensitivity"
        .Sensitivity = Sensitivity
        CurrentProperty = "Journal"
        .Journal = Journal
        
'-- add picture if applicable
        CurrentProperty = "AddPicture"
        If AddPicture <> "" Then .AddPicture AddPicture
        
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
        
'-- process OtherProperties, if applicable.  If argument value is not an array, or is an array
'-- with just one dimension, this will throw an error
        CurrentProperty = "OtherProperties"
        If IsArray(OtherProperties) Then
'===============================================================================
' The code does not specify what the bounds have to be for the OtherProperties array, so
' to get the properties, use LBound/UBound to set the parameters for the loop.  Likewise,
' since we do not know ahead of time what the bounds will be for the second dimension
' (could be 0 To 1, could be 1 To 2), use LBound to get property name and UBound to get
' the property value.  Theoretically, the second dimension could have >= 3 elements, but
' the code will still work if the property name is in the first element and property
' value is in the second
'===============================================================================
            For counter = LBound(OtherProperties, 1) To UBound(OtherProperties, 1)
                Set ItemProp = .ItemProperties.item(OtherProperties(counter, LBound(OtherProperties, 2)))
                ItemProp.Value = OtherProperties(counter, UBound(OtherProperties, 2))
            Next
        End If
'===============================================================================
' Use this to save the item in a folder other than the default folder for its type.
' If SaveToFolder is passed as an Outlook.Folder object, then use that.  If it is
' passed as a path string, then first try to get that folder, or if it does not
' exist, create it
'===============================================================================
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
        
'-- save/close or display the item in an Inspector
        If CloseRightAway Then
            CurrentProperty = "Close"
            .Close olSave
        Else
            CurrentProperty = "Display"
            .display
        End If
        
    End With
    
    CreateContactItem = True
    
    GoTo Cleanup
    
ErrHandler:
    
'-- log error if applicable
    If Me.ErrorLoggingEnabled Then WriteToLog "ContactItem", tag, CurrentProperty, Err.description
    CreateContactItem = False
    
Cleanup:
    
    Set olContact = Nothing
    
End Function

Public Function CreateMailItem(SendTo As Variant, Optional CC As Variant = "", _
    Optional BCC As Variant = "", Optional subject As String = "", _
    Optional body As String = "", Optional HTMLBody As String = "", _
    Optional Attachments As Variant, Optional Importance As OlImportance = olImportanceNormal, _
    Optional Categories As String = "", Optional DeferredDeliveryTime As Date = #1/1/1950#, _
    Optional DeleteAfterSubmit As Boolean = False, Optional FlagRequest As String = "", _
    Optional ReadReceiptRequested As Boolean = False, Optional Sensitivity As OlSensitivity = olNormal, _
    Optional SaveSentMessageFolder As Variant = "", Optional CloseRightAway As Boolean = False, _
    Optional EnableReply As Boolean = True, Optional EnableReplyAll As Boolean = True, _
    Optional EnableForward As Boolean = True, Optional ReplyRecipients As Variant = "", _
    Optional OtherProperties As Variant = "", Optional tag As String = "") As Boolean
'===============================================================================
' Method for creating (and sending, if desired) mail items.  SendTo is required; all other arguments
' are optional.
'
' Returns True for successful completion, and False otherwise.  If unsuccessful, no message is sent,
' and if logging is enabled at the class level, the procedure writes details to the log file
'
' Most arguments correspond to MailItem properties in Outlook, so look in the Outlook VBA help file
' for documentation on them.  The exceptions:
'
' - SendTo, CC, BCC:     Can be either a 1-dimensional array of recipients, or a string.  If an
'                        array, use each member of the array to list the recipients.  If a string
'                        with Len > 0, use the To/CC/BCC properties to add via string.  For
'                        multiple recipients using string, use semicolon as delimiter
'
' - Attachments:         Can be either a 1-dimensional array of attachment paths, or a string.  If an
'                        array, use each member of the array to list the file path.  If a string with
'                        Len > 0, add the file given the path
'
' - Tag:                 Allows passing a user-defined value that is not used in the MailItem,
'                        but may be helpful for troubleshooting.  For example, you might pass in a
'                        row number from Excel or a primary key value from Access
'
' SaveSentMessageFolder: In Outlook this is always an Outlook folder.  Here, it can be an object
'                        representing the Outlook folder, or it can be a string.  If a string, it
'                        identifies the folder by path.  Use two backslash characters to delimit
'                        nested folders.  Double backslash at the beginning for the top level folder
'                        is optional.  For example, these both get the same folder:
'                        -------
'                        \\Inbox - Joe Schmoe\\Customers\\Nifco Corp
'                          Inbox - Joe Schmoe\\Customers\\Nifco Corp
'                        -------
'                        If the folder does not already exist, the folder will be created (as will
'                        any "missing" levels in the indicated path)
'
' - CloseRightAway:      Determines whether to send the item (True), or open the item in an Inspector
'
' - EnableReply:         Allow user to reply to the message
'
' - EnableReplyAll:      Allow user to reply all to the message
'
' - EnableForward:       Allow user to forward the message
'
' - ReplyRecipients:     Can be either a 1-dimensional array of recipients, or a string.  If an
'                        array, use each member of the array to list the recipients.  For multiple
'                        recipients always use an array
'
' - OtherProperties:     Allows user to set other item properties not already handled in the basic
'                        class.  If used, pass as a two-dimensional array, with the first dimension
'                        indicating the property name, and the second indicating the value to use
'
' Note on EnableReply/EnableReplyAll/EnableForward: These will only work on messages sent within
' your organization, and can be easily "undone" by someone who knows how to manipulate Actions
'
' At each step, CurrentProperty variable updates to reflect what we are doing.  This facilitates
' error logging
'===============================================================================
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
        If IsArray(SendTo) Then
            For counter = LBound(SendTo) To UBound(SendTo)
                Set olRecip = .Recipients.Add(SendTo(counter))
                olRecip.type = olTo
            Next
        Else
            If SendTo <> "" Then .To = SendTo
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
        If HTMLBody <> "" And body = "" Then .HTMLBody = HTMLBody
        
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
        If Val(olApp.version) >= 12 Then
            CurrentProperty = "FlagRequest"
            If FlagRequest <> "" Then .FlagRequest = FlagRequest
        End If
        
'-- standard
        CurrentProperty = "ReadReceiptRequested"
        .ReadReceiptRequested = ReadReceiptRequested
        CurrentProperty = "Sensitivity"
        .Sensitivity = Sensitivity
'===============================================================================
' Argument can be passed as an Outlook folder or a string (indicating the folder to use by path).
' If a folder, then set the property directly.  If a string, first try to get the folder using
' GetFolderFromPath.  If folder does not already exist, use AddFolderFromPath to create the
' folder
'===============================================================================
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
'===============================================================================
' The code does not specify what the bounds have to be for the OtherProperties array, so
' to get the properties, use LBound/UBound to set the parameters for the loop.  Likewise,
' since we do not know ahead of time what the bounds will be for the second dimension
' (could be 0 To 1, could be 1 To 2), use LBound to get property name and UBound to get
' the property value.  Theoretically, the second dimension could have >= 3 elements, but
' the code will still work if the property name is in the first element and property
' value is in the second
'===============================================================================
            For counter = LBound(OtherProperties, 1) To UBound(OtherProperties, 1)
                Set ItemProp = .ItemProperties.item(OtherProperties(counter, LBound(OtherProperties, 2)))
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
    
    GoTo Cleanup
    
ErrHandler:
    
'-- log error if applicable
    If Me.ErrorLoggingEnabled Then WriteToLog "MailItem", tag, CurrentProperty, Err.description
    CreateMailItem = False
    
Cleanup:
    
    Set olMsg = Nothing
    
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
'===============================================================================
' Method for creating appointments / sending meeting requests.  StartAt is required; all other
' arguments are optional.
'
' Due to the complexity involved, this class does not set recurrence patterns
'
' If there are no attendees, then it is simply an appointment.  If there are, then a meeting
' request is sent
'
' Returns True for successful completion, and False otherwise.  If unsuccessful, no message is sent,
' and if logging is enabled at the class level, the procedure writes details to the log file
'
' Most arguments correspond to AppointmentItem properties in Outlook, so look in the Outlook VBA
' help file for documentation on them.  The exceptions:
'
' - RequiredAttendees,  Each can be a 1-dimensional array of the various recipient types, or can
' OptionalAttendees,    be a string.  Only use a string when you have one of that recipient type
' Resources,
' Organizer:            NOTE! In Outlook 2007, if you set CloseRightAway = False -- i.e., display
'                       first rather than sending, all attendees get forced to Required!  Not
'                       tested on other versions, so I cannot say whether or not that behavior
'                       happens in other Outlook versions...
'
' - Attachments:        Can be either a 1-dimensional array of attachment paths, or a string.  If
'                       an array, use each member of the array to list the file path.  If a string,
'                       add the file given the path
'
' - Tag:                Allows passing a user-defined value that is not used in the MailItem, but
'                       may be helpful for troubleshooting errors.  For example, you might pass in
'                       a row number from Excel or a primary key value from Access
'
' - CloseRightAway:     Determines whether to send the item (True), or open the item in an Inspector
'
' - OtherProperties:    Allows user to set other item properties not already handled in the basic
'                       class.  If used, pass as a two-dimensional array, with the first dimension
'                       indicating the property name, and the second indicating the value to use
'
' - SaveToFolder:       If you wish to save the new item to a folder other than its default folder, use
'                       this argument.  Argument can be an object representing the Outlook folder, or it
'                       can be a string.  If a string, it identifies the folder by path.  Use two backslash
'                       characters to delimit nested folders.  Double backslash at the beginning for the
'                       top level folder is optional.  For example, these both get the same folder:
'                       -------
'                       \\Inbox - Joe Schmoe\\Customers\\Nifco Corp
'                       Inbox - Joe Schmoe\\Customers\\Nifco Corp
'                       -------
'                       If the folder does not already exist, the folder will be created (as will
'                       any "missing" levels in the indicated path)
'
' At each step, CurrentProperty variable updates to reflect what we are doing.
' This facilitates error logging
'===============================================================================
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
'===============================================================================
' There are two ways to indicate the end of the meeting: End and Duration.  Method uses
' whichever argument would lead to the longer meeting.  Thus, for:
'
'         StartAt = #2010-12-01 08:00
'         EndAt = #2010-12-01 08:30
'         Duration = 60
'
' we would set the ending at #2010-12-01 09:00 and for:
'
'         StartAt = #2010-12-01 08:00
'         EndAt = #2010-12-01 10:00
'         Duration = 30
'
' we would set the ending at #2010-12-01 10:00
'===============================================================================
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
'===============================================================================
' The code does not specify what the bounds have to be for the OtherProperties array, so
' to get the properties, use LBound/UBound to set the parameters for the loop.  Likewise,
' since we do not know ahead of time what the bounds will be for the second dimension
' (could be 0 To 1, could be 1 To 2), use LBound to get property name and UBound to get
' the property value.  Theoretically, the second dimension could have >= 3 elements, but
' the code will still work if the property name is in the first element and property
' value is in the second
'===============================================================================
            For counter = LBound(OtherProperties, 1) To UBound(OtherProperties, 1)
                Set ItemProp = .ItemProperties.item(OtherProperties(counter, LBound(OtherProperties, 2)))
                ItemProp.Value = OtherProperties(counter, UBound(OtherProperties, 2))
            Next
        End If
'===============================================================================
' Use this to save the item in a folder other than the default folder for its type.  If SaveToFolder
' is passed as an Outlook.Folder object, then use that.  If it is passed as a path string, then
' first try to get that folder, or if it does not exist, create it
'===============================================================================
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
                .Close olSave
            Else
                CurrentProperty = "Display"
                .display
            End If
        End If
    End With
    
    CreateAppointmentItem = True
    
    GoTo Cleanup
    
ErrHandler:
    
'-- log error if applicable
    If Me.ErrorLoggingEnabled Then WriteToLog "AppointmentItem", tag, CurrentProperty, Err.description
    CreateAppointmentItem = False
    
Cleanup:
    
    Set olAppt = Nothing
    
End Function

Function CreateNoteItem(body As String, Optional Categories As String = "", _
    Optional OtherProperties As Variant = "", Optional tag As String = "", _
    Optional CloseRightAway As Boolean = True, Optional SaveToFolder As Variant = "") As Boolean
'===============================================================================
' Method for creating Notes.  Body is required; all other arguments are optional.
'
' Most arguments correspond to TaskItem properties in Outlook, so look in the Outlook VBA
' help file for documentation on them.  The exceptions:
'
' - Tag:             Allows passing a user-defined value that is not used in the MailItem, but may
'                    be helpful for troubleshooting errors.  For example, you might pass in a row
'                    number from Excel or a primary key value from Access
'
' - CloseRightAway:  Determines whether to send the item (True), or open the item in an Inspector
'
' - OtherProperties: Allows user to set other item properties not already handled in the basic
'                    class.  If used, pass as a two-dimensional array, with the first dimension
'                    indicating the property name, and the second indicating the value to use
'
' SaveToFolder:      If you wish to save the new item to a folder other than its default folder, use
'                    this argument.  Argument can be an object representing the Outlook folder, or it
'                    can be a string.  If a string, it identifies the folder by path.  Use two backslash
'                    characters to delimit nested folders.  Double backslash at the beginning for the
'                    top level folder is optional.  For example, these both get the same folder:
'                    -------
'                    \\Inbox - Joe Schmoe\\Customers\\Nifco Corp
'                      Inbox - Joe Schmoe\\Customers\\Nifco Corp
'                    -------
'                    If the folder does not already exist, the folder will be created (as will
'                    any "missing" levels in the indicated path)
'===============================================================================
    Dim olNote As Object                'Outlook.NoteItem
    Dim counter As Long
    Dim CurrentProperty As String
    Dim ItemProp As Object              'Outlook.ItemProperty
    Dim TestFolder As Object            'Outlook.Folder
    
    On Error GoTo ErrHandler
    
    CurrentProperty = "CreateItem"
    Set olNote = olApp.CreateItem(olNoteItem)
    With olNote
        
'-- standard
        .body = body
        If Categories <> "" Then .Categories = Categories
        
'-- process OtherProperties
        CurrentProperty = "OtherProperties"
        If IsArray(OtherProperties) Then
'===============================================================================
' The code does not specify what the bounds have to be for the OtherProperties array, so
' to get the properties, use LBound/UBound to set the parameters for the loop.  Likewise,
' since we do not know ahead of time what the bounds will be for the second dimension
' (could be 0 To 1, could be 1 To 2), use LBound to get property name and UBound to get
' the property value.  Theoretically, the second dimension could have >= 3 elements, but
' the code will still work if the property name is in the first element and property
' value is in the second
'===============================================================================
            For counter = LBound(OtherProperties, 1) To UBound(OtherProperties, 1)
                Set ItemProp = .ItemProperties.item(OtherProperties(counter, LBound(OtherProperties, 2)))
                ItemProp.Value = OtherProperties(counter, UBound(OtherProperties, 2))
            Next
        End If
'==============================================================================
' Use this to save the item in a folder other than the default folder for its type.  If SaveToFolder
' is passed as an Outlook.Folder object, then use that.  If it is passed as a path string, then
' first try to get that folder, or if it does not exist, create it
'==============================================================================
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
        
'-- save/close or display the item in an Inspector
        If CloseRightAway Then
            CurrentProperty = "Close"
            .Close olSave
        Else
            CurrentProperty = "Display"
            .display
        End If
    
    End With
    
    CreateNoteItem = True
    
    GoTo Cleanup
    
ErrHandler:
    
'-- log error if applicable
    If Me.ErrorLoggingEnabled Then WriteToLog "NoteItem", tag, CurrentProperty, Err.description
    CreateNoteItem = False
    
Cleanup:
    
    Set olNote = Nothing
    
End Function

Function CreateTaskItem(subject As String, Optional AssignTo As Variant = "", _
    Optional dueDate As Date = 0, Optional body As String = "", _
    Optional Importance As OlImportance = olImportanceNormal, Optional ReminderSet As Boolean = True, _
    Optional ReminderTime As Date = 0, Optional Attachments As Variant, _
    Optional Categories As String = "", Optional Sensitivity As OlSensitivity = olNormal, _
    Optional tag As String = "", Optional CloseRightAway As Boolean = True, _
    Optional OtherProperties As Variant = "", Optional SaveToFolder As Variant = "") As Boolean
'===============================================================================
' Method for creating tasks / sending task requests.  StartAt is required; all other arguments
' are optional.
'
' Due to the complexity involved, this class does not set recurrence patterns
'
' If the task is not assigned to someone, then it is simply a task.  If there are, then a task
' request is sent
'
' Returns True for successful completion, and False otherwise.  If unsuccessful, no message is sent,
' and if logging is enabled at the class level, the procedure writes details to the log file
'
' Most arguments correspond to TaskItem properties in Outlook, so look in the Outlook VBA
' help file for documentation on them.  The exceptions:
'
' - AssignTo:       Can be either a 1-dimensional array of recipients, or a string.  If an
'                   array, use each member of the array to list the recipients.  For multiple
'                   recipients always use an array.  Passing a value here makes the TaskItem a
'                   Task Request
'
' - Attachments:    Can be either a 1-dimensional array of attachment paths, or a string.  If an
'                   array, use each member of the array to list the file path.  If a string with
'                   Len > 0, add the file given the path
' - Tag:            Allows passing a user-defined value that is not used in the MailItem, but may
'                   be helpful for troubleshooting errors.  For example, you might pass in a row
'                   number from Excel or a primary key value from Access
' - CloseRightAway: Determines whether to send the item (True), or open the item in an Inspector
'
' - OtherProperties: Allows user to set other item properties not already handled in the basic
'                    class.  If used, pass as a two-dimensional array, with the first dimension
'                    indicating the property name, and the second indicating the value to use
'
' - SaveToFolder:    If you wish to save the new item to a folder other than its default folder, use
'                    this argument.  Argument can be an object representing the Outlook folder, or it
'                    can be a string.  If a string, it identifies the folder by path.  Use two backslash
'                    characters to delimit nested folders.  Double backslash at the beginning for the
'                    top level folder is optional.  For example, these both get the same folder:
'                    -------
'                    \\Inbox - Joe Schmoe\\Customers\\Nifco Corp
'                      Inbox - Joe Schmoe\\Customers\\Nifco Corp
'                    -------
'                    If the folder does not already exist, the folder will be created (as will
'                    any "missing" levels in the indicated path)
'
' At each step, CurrentProperty variable updates to reflect what we are doing.  This facilitates
' error logging
'===============================================================================
    Dim olTask As Object            'Outlook.TaskItem
    Dim olRecip As Object           'Outlook.Recipient
    Dim counter As Long
    Dim CurrentProperty As String
    Dim ItemProp As Object          'Outlook.ItemProperty
    Dim TestFolder As Object        'Outlook.Folder
    
    On Error GoTo ErrHandler
    
    CurrentProperty = "CreateItem"
    Set olTask = olApp.CreateItem(olTaskItem)
    With olTask
        
'-- standard fields
        CurrentProperty = "Subject"
        .subject = subject
        CurrentProperty = "DueDate"
        If dueDate >= 0 Then .dueDate = dueDate
        CurrentProperty = "Body"
        If body <> "" Then .body = body
        CurrentProperty = "Importance"
        .Importance = Importance
        
'-- set reminder, if applicable.  If ReminderSet = True but no ReminderTime is specified,
'-- then use the DueDate
        CurrentProperty = "ReminderSet"
        If ReminderSet Then
            .ReminderSet = True
            If ReminderTime > 0 Then
                .ReminderTime = ReminderTime
            Else
                .ReminderTime = .dueDate
            End If
        Else
            .ReminderSet = False
        End If
        
'-- process Attachments.  For multiple files, use an array.  For a single file, use a string
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
        
'-- standard fields
        CurrentProperty = "Categories"
        If Categories <> "" Then .Categories = Categories
        CurrentProperty = "Sensitivity"
        .Sensitivity = Sensitivity
        
        CurrentProperty = "AssignTo"
        If IsArray(AssignTo) Then
            .Assign
            For counter = LBound(AssignTo) To UBound(AssignTo)
                .Recipients.Add AssignTo(counter)
            Next
        Else
            If AssignTo <> "" Then
                .Assign
                .Recipients.Add AssignTo(counter)
            End If
        End If
        
'-- process OtherProperties, if applicable.  If argument value is not an array, or is an array
'-- with just one dimension, this will throw an error
        CurrentProperty = "OtherProperties"
        If IsArray(OtherProperties) Then
'===============================================================================
' The code does not specify what the bounds have to be for the OtherProperties array, so
' to get the properties, use LBound/UBound to set the parameters for the loop.  Likewise,
' since we do not know ahead of time what the bounds will be for the second dimension
' (could be 0 To 1, could be 1 To 2), use LBound to get property name and UBound to get
' the property value.  Theoretically, the second dimension could have >= 3 elements, but
' the code will still work if the property name is in the first element and property
' value is in the second
'===============================================================================
            For counter = LBound(OtherProperties, 1) To UBound(OtherProperties, 1)
                Set ItemProp = .ItemProperties.item(OtherProperties(counter, LBound(OtherProperties, 2)))
                ItemProp.Value = OtherProperties(counter, UBound(OtherProperties, 2))
            Next
        End If
'===============================================================================
' Use this to save the item in a folder other than the default folder for its type.  If SaveToFolder
' is passed as an Outlook.Folder object, then use that.  If it is passed as a path string, then
' first try to get that folder, or if it does not exist, create it
'===============================================================================
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
        
'-- if there are no recipients, then it is just a task, and the choice is Close/Display.
'-- if there are recipients, then it is a task request, and the choice is Send/Display.
        CurrentProperty = "Send/Display"
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
                CurrentProperty = "Save"
                .Close olSave
            Else
                CurrentProperty = "Display"
                .display
            End If
        End If
    End With
    
    CreateTaskItem = True
    
    GoTo Cleanup
    
ErrHandler:
    
'-- log error if applicable
    If Me.ErrorLoggingEnabled Then WriteToLog "TaskRequestItem", tag, CurrentProperty, Err.description
    CreateTaskItem = False
    
Cleanup:
    
    Set olTask = Nothing
    
End Function

Property Get OutlookApplication() As Object 'Outlook.Application
    
'-- exposes the Outlook.Application object
    Set OutlookApplication = olApp
    
End Property

Private Sub WriteToLog(ItemType As String, tag As String, CurrentProperty As String, ErrDescr As String)
    
'-- increment error count, and write record to log file
    TotalErrors = TotalErrors + 1
    
    tsLog.WriteLine Join(Array(Format(Now, "yyyy-mm-dd hh:nn:ss"), ItemType, tag, CurrentProperty, Err.description), ",")
    
End Sub