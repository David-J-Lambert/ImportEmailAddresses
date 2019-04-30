VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDestinations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Module clsDestinations.
' Author: David Lambert, David5Lambert7@Gmail.com.
' MICROSOFT OFFICE SETUP: The spreadsheet "INSTRUCTIONS" has macro setup instructions.

Option Explicit

' Collection of Subfolders and Contact Groups to work with.
Private sDestinationList As Collection

' Outlook Objects
Private outlookApp As Object
Private objContactGroup As Outlook.DistListItem
Private objFolder As Outlook.Folder
Private objFolders As Outlook.Folders
Private objNameSpace As Outlook.Namespace
Private objSubFolder As Outlook.Folder

Private Sub Class_Initialize()
    Set sDestinationList = New Collection

    Set outlookApp = CreateObject("Outlook.Application")

    Set objNameSpace = outlookApp.GetNamespace("MAPI")
    Set objFolders = objNameSpace.Folders
End Sub

Private Sub Class_Terminate()
    Set sDestinationList = Nothing

    Set outlookApp = Nothing

    Set objFolders = Nothing
    Set objNameSpace = Nothing

    Set objContactGroup = Nothing
    Set objFolder = Nothing
    Set objSubFolder = Nothing
End Sub

' PUBLIC FUNCTIONS AND SUBROUTINES RELATED TO DESTINATIONS

Public Sub SetDestinationsInUI(ByVal sCmdButtonCaption As String, ByVal sCmdButtonAccelerator As String)

    Dim objItem As Object
    Dim i As Integer, iListRow As Integer
    Dim sName As String

    Dim sDestinationListEntry(0 To 2) As String
    ' Field 0: objFolder.Name
    ' Field 1: objSubFolder.Name
    ' Field 2: for a Contact Folder: vbNullString,
    '          for a Contact Group: objContactGroup.DLName

    ' Initialize Form
    boolSelectionMade = False
    frmSelectDestinations.Caption = sThisProgram & ", " & sThisVersion
    frmSelectDestinations.Show vbModeless

    ' Get list of folders containing subfolders for contacts.
    With frmSelectDestinations
        ' Go through all top-level folders.
        For Each objFolder In objFolders
            ' Go through all subfolders of current top-level folder.
            For Each objSubFolder In objFolder.Folders
                ' This subfolder must be for contacts, and must not be a hidden object.
                sName = objSubFolder.Name
                If (objSubFolder.DefaultItemType = olContactItem And sName <> "PersonMetadata" And sName <> "ExternalContacts") Then
                    ' List box, column 0: folder, column 1: subfolder, column 2: nothing
                    .lstFolders.AddItem objFolder.Name
                    .lstFolders.List(.lstFolders.ListCount - 1, 1) = objSubFolder.Name
                    .lstFolders.List(.lstFolders.ListCount - 1, 2) = vbNullString
                    ' Go through items in current subfolder.
                    For Each objItem In objSubFolder.Items
                        ' This item must be a contact group.
                        If (TypeOf objItem Is DistListItem) Then
                            Set objContactGroup = objItem
                            ' List box, column 0: folder, column 1: subfolder, column 2: contact group (collection)
                            .lstFolders.AddItem objFolder.Name
                            .lstFolders.List(.lstFolders.ListCount - 1, 1) = objSubFolder.Name
                            .lstFolders.List(.lstFolders.ListCount - 1, 2) = objContactGroup.DLName
                        End If
                    Next
                End If
            Next
        Next

        .cmdSelectDestinations.Accelerator = sCmdButtonAccelerator
        .cmdSelectDestinations.Caption = sCmdButtonCaption
        .cmdSelectDestinations.Enabled = True

        .cmdDefaultDestinations.Accelerator = "S"
        .cmdDefaultDestinations.Caption = "Select Default Contact Groups (Unimplemented)"
        .cmdDefaultDestinations.Enabled = True

        ' frmSelectDestinations is Modeless, so force wait here until selection made, to keep rest of program from running.
        Do Until boolSelectionMade = True
            DoEvents
        Loop

        ' Get list of chosen folders and contact groups.
        For iListRow = 0 To .lstFolders.ListCount - 1
            If .lstFolders.Selected(iListRow) Then
                For i = 0 To 2
                    sDestinationListEntry(i) = .lstFolders.List(iListRow, i)
                Next
            sDestinationList.Add sDestinationListEntry
            End If
        Next
    End With

    Set objItem = Nothing

End Sub ' SetDestinationsInUI(ByVal sCmdButtonCaption As String, ByVal sCmdButtonAccelerator As String)

Public Function GetDestinations() As Collection
    Set GetDestinations = sDestinationList
End Function
