VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ListThreadFolders 
   Caption         =   "Select folder to move emails to"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11190
   OleObjectBlob   =   "ListThreadFolders.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ListThreadFolders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub UserForm_Initialize()

    GetConverstationInformation

End Sub

Public Sub GetConverstationInformation()

    ' Original code obtained from the following site (credit user TimO):
    ' https://stackoverflow.com/questions/29304844/outlook-2010-vba-to-save-selected-email-to-a-folder-other-emails-in-that-convers?rq=1

    ' Get root items in conversation

    Dim host As Outlook.Application
    Set host = ThisOutlookSession.Application

    ' Get the user's currently selected item
    Set selectedItem = host.ActiveExplorer.Selection.item(1)
    Debug.Print ("Selected item: " & selectedItem.ConversationTopic)

    ' Check to see that the item's current folder has conversations enabled
    Dim parentFolder As Outlook.folder
    Dim parentStore As Outlook.store
    Set parentFolder = selectedItem.Parent
    Set parentStore = parentFolder.store
    If parentStore.IsConversationEnabled Then
        ' Try and get the conversation.
        Dim theConversation As Outlook.conversation
        Set theConversation = selectedItem.GetConversation
        If Not IsNull(theConversation) Then
            ' Outlook provides a table object the contains all of the items in the conversation
            Dim itemsTable As Outlook.table
            Set itemsTable = theConversation.GetTable

            ' Get the Root Items
            ' Enumerate the list of items
            ' Then use a helper method and recursion to walk all the items in the conversation
            Dim group As Outlook.SimpleItems
            Set group = theConversation.GetRootItems
            Dim obj As Object ' an email
            Dim fld As Outlook.folder ' full path to the folder the email is in (\\AcountName\Folder)
            Dim sfld As String ' path to the folder the email is in excluding the account name (\Folder)
            Dim IsInListBox As Boolean
            For Each obj In group
                If TypeOf obj Is Outlook.MailItem Or TypeOf obj Is Outlook.AppointmentItem Or TypeOf obj Is Outlook.MeetingItem Then
                    ' If ROOT item is an email, add it to ListBox1

                    Set fld = obj.Parent
                    FolderPathEncoded = Replace(fld.FolderPath, "%2F", "/")
                    Debug.Print ("FolderPathEncoded: " & FolderPathEncoded & " (" & TypeName(obj) & ")")

                    ' Don't include generic folders
                    sfld = Mid(FolderPathEncoded, InStr(3, FolderPathEncoded, "\") + 1)
                    If (sfld <> "Inbox") And _
                        (sfld <> "Drafts") And _
                        (sfld <> "Sent Items") And _
                        (sfld <> "Calendar") And _
                        (sfld <> "Auto Replies") And _
                        (InStr(sfld, "Shared Data") = 0) Then

                        ' Make IsInListBox true if folder has already been added
                        IsInListBox = False
                        For i = 0 To Me.ListBox1.ListCount - 1
                            If Me.ListBox1.Column(0, i) = FolderPathEncoded Then
                                IsInListBox = True
                            End If
                        Next

                        If (IsInListBox = False) Then
                            Me.ListBox1.AddItem FolderPathEncoded
                            Debug.Print ("Added " & FolderPathEncoded & " to ListBox")
                        End If

                    End If

                Else
                    Debug.Print ("Skipping obj of type " & TypeName(obj))
                End If

                ' Repeat the process if this email is also a root item
                GetConversationDetails obj, theConversation

            Next obj
        Else
            MsgBox "The currently selected item is not a part of a conversation."
        End If
    Else
        MsgBox "The currently selected item is not in a folder with conversations enabled."
    End If

    ' Display message box and/or move emails
    If Me.ListBox1.ListCount = 0 Then
        ' Don't open the window
        MsgBox ("No folders found")
        End
    End If
    If Me.ListBox1.ListCount = 1 Then
        ' Move emails and don't open window
        Call MoveMail(Me.ListBox1.Column(0, 0))
        MsgBox ("Moved email(s) to " & Me.ListBox1.Column(0, 0))
        End
    End If

End Sub

Private Sub GetConversationDetails(anItem As Object, theConversation As Outlook.conversation)

    ' Original code obtained from the following site (credit user TimO):
    ' https://stackoverflow.com/questions/29304844/outlook-2010-vba-to-save-selected-email-to-a-folder-other-emails-in-that-convers?rq=1

    ' From the root items, find all the messages and add to ListBox1

    Dim group As Outlook.SimpleItems
    Set group = theConversation.GetChildren(anItem)

    If group.Count > 0 Then
        Debug.Print ("Getting conversation details...")
        Dim obj As Object ' an email
        Dim fld As Outlook.folder ' full path to the folder the email is in (\\AcountName\Folder)
        Dim sfld As String ' path to the folder the email is in excluding the account name (\Folder)        Dim i As Integer
        Dim IsInListBox As Boolean
        For Each obj In group
            If TypeOf obj Is Outlook.MailItem Or TypeOf obj Is Outlook.AppointmentItem Or TypeOf obj Is Outlook.MeetingItem Then
                ' If CHILD item is an email, add it to ListBox1

                Set fld = obj.Parent
                FolderPathEncoded = Replace(fld.FolderPath, "%2F", "/")
                Debug.Print ("  FolderPathEncoded: " & FolderPathEncoded & " (" & TypeName(obj) & ")")

                ' Don't include generic folders
                sfld = Mid(FolderPathEncoded, InStr(3, FolderPathEncoded, "\") + 1)
                If (sfld <> "Inbox") And _
                    (sfld <> "Drafts") And _
                    (sfld <> "Sent Items") And _
                    (sfld <> "Calendar") And _
                    (sfld <> "Auto Replies") And _
                    (InStr(sfld, "Shared Data") = 0) Then

                    ' Make IsInListBox true if folder has already been added
                    IsInListBox = False
                    For i = 0 To Me.ListBox1.ListCount - 1
                        If Me.ListBox1.Column(0, i) = FolderPathEncoded Then
                            IsInListBox = True
                        End If
                    Next

                    ' Add folder to ListBox if IsInListBox is false
                    If IsInListBox = False Then
                        Me.ListBox1.AddItem FolderPathEncoded
                        Debug.Print ("  Added " & FolderPathEncoded & " to ListBox")
                    End If

                End If

            Else
                Debug.Print ("  Skipping obj of type " & TypeName(obj))
            End If

            ' Repeat the process if this email is also a root item
            GetConversationDetails obj, theConversation

        Next obj
    End If

End Sub

Private Sub ListBox1_Click()

    ' Move mail to selected folder
    Call MoveMail(Me.ListBox1.Value)

    ' Close UserForm
    Unload Me

End Sub

Sub MoveMail(inputfolder As String)

    ' Original code obtained from the following site (credit Diane Poremsky):
    ' https://www.slipstick.com/outlook/macro-move-folder/

    Dim objOutlook As Outlook.Application
    Dim objNamespace As Outlook.NameSpace
    Dim objSourceFolder As Outlook.MAPIFolder
    Dim objDestFolder As Outlook.MAPIFolder
    Dim objItems As MailItem

    Set objOutlook = Application
    Set objNamespace = objOutlook.GetNamespace("MAPI")
    Set objSourceFolder = objNamespace.GetDefaultFolder(olFolderDrafts)
    Set objDestFolder = GetFolder(inputfolder)

    For Each objItem In objOutlook.ActiveExplorer.Selection

        ' Move folder if destination is different than current
        If objItem.Parent <> objDestFolder Then
            objItem.Move objDestFolder
            Debug.Print ("Moved '" & objItem.ConversationTopic & "' to '" & objDestFolder.name & "'")
        Else
            Debug.Print ("Skipped moving '" & objItem.ConversationTopic & "' to '" & objDestFolder.name & "' (same folder)")
        End If

    Next

    Set objDestFolder = Nothing

End Sub

Function GetFolder(ByVal FolderPath As String) As Outlook.folder

    ' Original code obtained from the following site (credit users "office 365 dev account", "Office GSX", Kim Brandl - MSFT, JiayueHu):
    ' https://docs.microsoft.com/en-us/office/vba/outlook/how-to/items-folders-and-stores/obtain-a-folder-object-from-a-folder-path

    ' Convert folder path in form of "\\folder1\folder2\folder3" to a folder object

    Dim TestFolder As Outlook.folder
    Dim FoldersArray As Variant
    Dim i As Integer

    On Error GoTo GetFolder_Error
    If Left(FolderPath, 2) = "\\" Then
        FolderPath = Right(FolderPath, Len(FolderPath) - 2)
    End If

    ' Convert folderpath to array
    FoldersArray = Split(FolderPath, "\")
    Set TestFolder = Application.Session.Folders.item(FoldersArray(0))
    If Not TestFolder Is Nothing Then
        For i = 1 To UBound(FoldersArray, 1)
            Dim SubFolders As Outlook.Folders
            Set SubFolders = TestFolder.Folders
            Set TestFolder = SubFolders.item(FoldersArray(i))
            If TestFolder Is Nothing Then
                Set GetFolder = Nothing
            End If
        Next
    End If

    ' Return the TestFolder
    Set GetFolder = TestFolder
    Exit Function

GetFolder_Error:
    Set GetFolder = Nothing
    Exit Function

End Function
