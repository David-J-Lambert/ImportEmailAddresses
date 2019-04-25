Attribute VB_Name = "ImportEmails"
' Module ImportEmails.
' Author: David Lambert, David5Lambert7@Gmail.com.
Option Explicit

Public Const sThisProgram As String = "ImportEmails"
Private Const sThisVersion As String = "Version 1.0"
Public Const iImportFolder As Integer = 0
Public Const iImportContactGroup As Integer = 1
Private Const iBatchSize As Integer = 50

Public boolSelectionMade As Boolean
Public sWorksheetName As String
Public sProgressCaption As String

' MICROSOFT OFFICE SETUP: The spreadsheet "How to set up macro" has macro setup instructions.

Public Sub Importer()
    ' MAIN PROGRAM, DRIVES EVERYTHING ELSE.

    ' Dictionary of all email addresses. Key: email address.
    ' Value: operation(s) ("A"=add, "D"=delete, "F"=update from, "T"=update to, "B"=banned because spreadsheet lists multiple operations).
    Dim dicEmailOperation As Object
    Set dicEmailOperation = CreateObject("Scripting.Dictionary")
    dicEmailOperation.CompareMode = vbTextCompare

    ' Dictionary of all email address updates. Key and Value: email addresses.
    ' For "F", key=update from address, value=update to address.  For "T", key=update to address, value=update from address.
    Dim dicEmailUpdate As Object
    Set dicEmailUpdate = CreateObject("Scripting.Dictionary")
    dicEmailUpdate.CompareMode = vbTextCompare

    ' Collection of Subfolders and Contact Groups to work with.
    Dim sDestinationList As Collection
    Set sDestinationList = New Collection

    ' Summary of all operations performed, printed at end.
    Dim sSummary As String
    sSummary = vbNullString

    ' Set public variables.
    sWorksheetName = " (worksheet '" & ActiveSheet.Name & "')"
    sProgressCaption = sThisProgram & sWorksheetName & " is running..."

    ' CHOOSE WHICH SUBFOLDERS AND CONTACT GROUPS TO WORK WITH, PLACE LIST OF THEM INTO sDestinationList.
    GetListsFromUI sDestinationList, "Add, Delete, or Update Emails Now", "A"

    ' GET EMAIL ADDRESSES FROM SPREADSHEET.
    ReadSpreadsheet dicEmailOperation, dicEmailUpdate, sSummary

    ' PROCESS ADDS, DELETES, AND UPDATES IN dicEmailOperation FOR EACH DESTINATION IN sDestinationList.
    AddDeleteUpdate dicEmailOperation, dicEmailUpdate, sSummary, sDestinationList

    sSummary = sSummary & vbCrLf & vbCrLf & "ALL DONE!"

    ' PRINT SUMMARY
    MsgBox sSummary, , sThisProgram & sWorksheetName

    ' Clean up
    Set dicEmailOperation = Nothing
    Set dicEmailUpdate = Nothing
    Set sDestinationList = Nothing

    ' Finish
    End
End Sub ' Importer()

Public Sub GetListsFromUI(ByRef sDestinationList As Collection, ByVal sCmdButtonCaption As String, ByVal sCmdButtonAccelerator As String)

    Dim outlookApp As Object
    Set outlookApp = CreateObject("Outlook.Application")

    Dim objNameSpace As Outlook.Namespace
    Set objNameSpace = outlookApp.GetNamespace("MAPI")

    Dim objFolders As Outlook.Folders
    Set objFolders = objNameSpace.Folders
    Dim objFolder As Outlook.Folder, objSubFolder As Outlook.Folder

    Dim objContactGroup As Outlook.DistListItem

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
    frmSelectFolders.Caption = sThisProgram & ", " & sThisVersion
    frmSelectFolders.Show vbModeless

    ' Get list of folders containing subfolders for contacts.
    With frmSelectFolders
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

        .cmdSelectFolders.Accelerator = sCmdButtonAccelerator
        .cmdSelectFolders.Caption = sCmdButtonCaption
        .cmdSelectFolders.Enabled = True

        .cmdDefaultFolders.Accelerator = "S"
        .cmdDefaultFolders.Caption = "Select Default Contact Groups (Unimplemented)"
        .cmdDefaultFolders.Enabled = True

        ' frmSelectFolders is Modeless, so force wait here until selection made, to keep rest of program from running.
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

    ' Clean up.
    Set objFolders = Nothing
    Set objNameSpace = Nothing
    Set outlookApp = Nothing

    Set objContactGroup = Nothing
    Set objItem = Nothing
    Set objSubFolder = Nothing
    Set objFolder = Nothing

End Sub ' GetListsFromUI

Public Sub ReadSpreadsheet(ByVal dicEmailOperation As Object, ByVal dicEmailUpdate As Object, ByRef sSummary As String)

    Dim numBadEmails As Integer, iFirstRow As Integer, iRowNumber As Integer, numUpdates As Integer
    Dim numUpdateFrom As Integer, numUpdateTo As Integer, numUpdateFromFinal As Integer, numUpdateToFinal As Integer

    Dim rColumn As Range, rColUpdateFrom As Range, rColUpdateTo As Range, rCell As Range, rCellFrom As Range, rCellTo As Range

    Dim vArray As Variant, vItem As Variant, vKey As Variant

    Dim sCol As String, sColTitle As String, sMessage As String, sEmailAddress As String, sUpdateFrom As String, sUpdateTo As String

    Dim ValidEmailAddress As Boolean, ValidEmailAddressTo As Boolean

    ' Dictionary of the Excel cell each email address comes from. Key: email addresses.  Value: cell address (range data type).
    Dim dicEmailSource As Object
    Set dicEmailSource = CreateObject("Scripting.Dictionary")
    dicEmailSource.CompareMode = vbTextCompare

    numUpdateFrom = 0 ' Enforce no more than 1.
    numUpdateTo = 0 ' Enforce no more than 1.
    numUpdateFromFinal = 0 ' Enforce no more than 1.
    numUpdateToFinal = 0 ' Enforce no more than 1.

    ' Collections to store columns to Add and Delete.
    Dim sListAdd As Collection, sListDelete As Collection
    Set sListAdd = New Collection
    Set sListDelete = New Collection

    ' Select active sheet in this Excel workbook.
    ActiveSheet.UsedRange.Select

    ' Extract valid email addresses from active sheet.
    numBadEmails = 0

    ' Go through cells in first row of active part of this Excel worksheet.
    For Each rColumn In Selection.Columns
        ' Get column id
        vArray = Split(rColumn.Cells(1, 1).Address, "$")
        sCol = vArray(1) ' sRow = vArray(2)

        ' USE COLUMN HEADER TEXT, IF PRESENT, TO TENTATIVELY SET RADIO BUTTON IN frmAddDeleteUpdate.
        sColTitle = UCase$(rColumn.Cells(1, 1).Value)
        If (InStr(sColTitle, "ADD") > 0) Then
            frmAddDeleteUpdate.rbAdd.Value = True
        ElseIf (InStr(sColTitle, "DELETE") > 0) Then
            frmAddDeleteUpdate.rbDelete.Value = True
        ElseIf (InStr(sColTitle, "UPDATE FROM") > 0) Then
            numUpdateFrom = numUpdateFrom + 1
            If (numUpdateFrom > 1) Then
                MsgBox "More than one column has 'UPDATE FROM' in row 1, for email addresses to be updated.  Quitting.", , sThisProgram & sWorksheetName
                End
            End If
            frmAddDeleteUpdate.rbUpdateFrom.Value = True
        ElseIf (InStr(sColTitle, "UPDATE TO") > 0) Then
            numUpdateTo = numUpdateTo + 1
            If (numUpdateTo > 1) Then
                MsgBox "More than one column has 'UPDATE TO' in row 1, for new email addresses to update to.  Quitting.", , sThisProgram & sWorksheetName
                End
            End If
            frmAddDeleteUpdate.rbUpdateTo.Value = True
        Else ' "IGNORE" or anything besides the strings above.
            frmAddDeleteUpdate.rbIgnore.Value = True
        End If

        ' GET EMAIL ADDRESS OPERATION DESIRED FOR THIS COLUMN: ADD, DELETE, UPDATE FROM, UPDATE TO.

        frmAddDeleteUpdate.lblColumn.Caption = sCol
        frmAddDeleteUpdate.Caption = sThisProgram & sWorksheetName
        frmAddDeleteUpdate.Show

        ' RECORD CHOICES MADE IN frmAddDeleteUpdate.

        If frmAddDeleteUpdate.rbAdd.Value Then
            sListAdd.Add rColumn
        ElseIf frmAddDeleteUpdate.rbDelete.Value Then
            sListDelete.Add rColumn
        ElseIf frmAddDeleteUpdate.rbUpdateFrom.Value Then
            Set rColUpdateFrom = rColumn
            numUpdateFromFinal = numUpdateFromFinal + 1
            If (numUpdateFromFinal > 1) Then
                MsgBox "More than one column contains email addresses to be updated.  Quitting.", , sThisProgram & sWorksheetName
                End
            End If
        ElseIf frmAddDeleteUpdate.rbUpdateTo.Value Then
            Set rColUpdateTo = rColumn
            numUpdateToFinal = numUpdateToFinal + 1
            If (numUpdateToFinal > 1) Then
                MsgBox "More than one column contains new email addresses to update to.  Quitting.", , sThisProgram & sWorksheetName
                End
            End If
        End If
    Next ' For Each rColumn In Selection.Columns

    ' READ CONTENTS OF CELLS IN "ADD" COLUMNS.
    For Each vItem In sListAdd ' vItem is rColumn in disguise
        For Each rCell In vItem.Cells
            ' Cell text.
            sEmailAddress = Trim$(rCell.Value)

            ValidEmailAddress = CheckEmailAddressText(sEmailAddress, rCell, numBadEmails)
            If ValidEmailAddress Then
                If dicEmailOperation.Exists(sEmailAddress) Then
                    sMessage = GetRejectMessage("A", dicEmailOperation(sEmailAddress))
                    If Not IsEmpty(dicEmailSource(sEmailAddress)) Then _
                        AddCellComment dicEmailSource(sEmailAddress), sMessage
                    AddCellComment rCell, sMessage
                    dicEmailOperation(sEmailAddress) = "B" & dicEmailOperation(sEmailAddress)
                    numBadEmails = numBadEmails + 1
                Else
                    dicEmailOperation(sEmailAddress) = "A"
                    Set dicEmailSource(sEmailAddress) = rCell
                End If
            End If
        Next ' For Each rCell In vItem.Cells
    Next ' For Each vItem In sListAdd

    ' READ CONTENTS OF CELLS IN "DELETE" COLUMNS.
    For Each vItem In sListDelete ' vItem is rColumn in disguise
        For Each rCell In vItem.Cells
            ' Cell text.
            sEmailAddress = Trim$(rCell.Value)

            ValidEmailAddress = CheckEmailAddressText(sEmailAddress, rCell, numBadEmails)
            If ValidEmailAddress Then
                If dicEmailOperation.Exists(sEmailAddress) Then
                    sMessage = GetRejectMessage("D", dicEmailOperation(sEmailAddress))
                    If Not IsEmpty(dicEmailSource(sEmailAddress)) Then _
                        AddCellComment dicEmailSource(sEmailAddress), sMessage
                    AddCellComment rCell, sMessage
                    dicEmailOperation(sEmailAddress) = "B" & dicEmailOperation(sEmailAddress)
                    numBadEmails = numBadEmails + 1
                Else
                    dicEmailOperation(sEmailAddress) = "D"
                    Set dicEmailSource(sEmailAddress) = rCell
                End If
            End If
        Next ' For Each rCell In vItem.Cells
    Next ' For Each vItem In sListAdd

    ' READ CONTENTS OF CELLS IN "UPDATE FROM" AND "UPDATE TO" COLUMNS.
    If (numUpdateFromFinal = 1 And numUpdateToFinal = 1) Then
        If (rColUpdateFrom.Cells.Count <> rColUpdateTo.Cells.Count) Then
            MsgBox "The two update columns are not the same height!  Quitting.", , sThisProgram & sWorksheetName
            End
        End If
        ' Walk down both columns at the same time.
        numUpdates = rColUpdateFrom.Cells.Count
        iFirstRow = rColUpdateFrom.Cells(1, 1).Row
        For iRowNumber = iFirstRow To (iFirstRow + numUpdates)
            ' Treat both cells as a unit, not separately.
            Set rCellFrom = rColUpdateFrom.Cells(iRowNumber, 1)
            sUpdateFrom = rCellFrom.Value
            ValidEmailAddress = CheckEmailAddressText(sUpdateFrom, rCellFrom, numBadEmails)

            Set rCellTo = rColUpdateTo.Cells(iRowNumber, 1)
            sUpdateTo = rCellTo.Value
            ValidEmailAddressTo = CheckEmailAddressText(sUpdateTo, rCellTo, numBadEmails)
            ' Both email addresses must be valid and not part of other email operations.
            If (ValidEmailAddress And ValidEmailAddressTo) Then
                If dicEmailOperation.Exists(sUpdateFrom) Then
                    sMessage = GetRejectMessage("F", dicEmailOperation(sUpdateFrom))
                    If Not IsEmpty(dicEmailSource(sUpdateFrom)) Then _
                        AddCellComment dicEmailSource(sUpdateFrom), sMessage
                    If Not IsEmpty(dicEmailSource(sUpdateTo)) Then _
                        AddCellComment dicEmailSource(sUpdateTo), sMessage
                    AddCellComment rCellFrom, sMessage
                    AddCellComment rCellTo, sMessage
                    dicEmailOperation(sUpdateFrom) = "B" & dicEmailOperation(sUpdateFrom)
                    dicEmailOperation(sUpdateTo) = "B" & dicEmailOperation(sUpdateTo)
                    numBadEmails = numBadEmails + 1 ' sUpdateFrom recorded 2x, once for "T", once for "F".
                End If
                If dicEmailOperation.Exists(sUpdateTo) Then
                    sMessage = GetRejectMessage("T", dicEmailOperation(sUpdateTo))
                    If Not IsEmpty(dicEmailSource(sUpdateFrom)) Then _
                        AddCellComment dicEmailSource(sUpdateFrom), sMessage
                    If Not IsEmpty(dicEmailSource(sUpdateTo)) Then _
                        AddCellComment dicEmailSource(sUpdateTo), sMessage
                    AddCellComment rCellFrom, sMessage
                    AddCellComment rCellTo, sMessage
                    dicEmailOperation(sUpdateFrom) = "B" & dicEmailOperation(sUpdateFrom)
                    dicEmailOperation(sUpdateTo) = "B" & dicEmailOperation(sUpdateTo)
                    numBadEmails = numBadEmails + 1 ' sUpdateFrom recorded 2x, once for "T", once for "F".
                End If
                If (sUpdateFrom = sUpdateTo) Then
                    sMessage = "Email address updated to itself, no change."
                    If Not IsEmpty(dicEmailSource(sUpdateFrom)) Then _
                        AddCellComment dicEmailSource(sUpdateFrom), sMessage
                    If Not IsEmpty(dicEmailSource(sUpdateTo)) Then _
                        AddCellComment dicEmailSource(sUpdateTo), sMessage
                    AddCellComment rCellFrom, sMessage
                    AddCellComment rCellTo, sMessage
                    dicEmailOperation(sUpdateFrom) = "B" & dicEmailOperation(sUpdateFrom)
                    dicEmailOperation(sUpdateTo) = "B" & dicEmailOperation(sUpdateTo)
                    numBadEmails = numBadEmails + 1 ' sUpdateFrom recorded 2x, once for "T", once for "F".
                Else
                    dicEmailUpdate.Add Key:=sUpdateFrom, Item:=sUpdateTo
                    dicEmailOperation(sUpdateFrom) = "F"
                    Set dicEmailSource(sUpdateFrom) = rCellFrom
                    dicEmailUpdate.Add Key:=sUpdateTo, Item:=sUpdateFrom
                    dicEmailOperation(sUpdateTo) = "T"
                    Set dicEmailSource(sUpdateTo) = rCellTo
                End If
            ElseIf ValidEmailAddress Then ' ValidEmailAddressTo = False
                If dicEmailOperation.Exists(sUpdateFrom) Then
                    sMessage = GetRejectMessage("F", dicEmailOperation(sUpdateFrom))
                    If Not IsEmpty(dicEmailSource(sUpdateFrom)) Then _
                        AddCellComment dicEmailSource(sUpdateFrom), sMessage
                    AddCellComment rCellFrom, sMessage
                End If
                sMessage = "Invalid or missing email address to update to."
                If Not IsEmpty(dicEmailSource(sUpdateFrom)) Then _
                    AddCellComment dicEmailSource(sUpdateFrom), sMessage
                AddCellComment rCellFrom, sMessage
                dicEmailOperation(sUpdateFrom) = "B" & dicEmailOperation(sUpdateFrom)
                numBadEmails = numBadEmails + 1
            ElseIf ValidEmailAddressTo Then ' ValidEmailAddressFrom = False
                If dicEmailOperation.Exists(sUpdateTo) Then
                    sMessage = GetRejectMessage("T", dicEmailOperation(sUpdateTo))
                    If Not IsEmpty(dicEmailSource(sUpdateTo)) Then _
                        AddCellComment dicEmailSource(sUpdateTo), sMessage
                    AddCellComment rCellTo, sMessage
                End If
                sMessage = "Invalid or missing email address to update from."
                If Not IsEmpty(dicEmailSource(sUpdateTo)) Then _
                    AddCellComment dicEmailSource(sUpdateTo), sMessage
                AddCellComment rCellTo, sMessage
                dicEmailOperation(sUpdateTo) = "B" & dicEmailOperation(sUpdateTo)
                numBadEmails = numBadEmails + 1
            Else
                ' Do nothing, both cells have comments.
            End If
        Next ' For iRowNumber = iFirstRow To (iFirstRow + numUpdates)
    ElseIf (numUpdateFromFinal + numUpdateToFinal > 0) Then
        MsgBox "Found an 'UPDATE FROM' column without an 'UPDATE TO' column, or vice versa.  Quitting.", , sThisProgram & sWorksheetName
        End
    End If ' If (numUpdateFromFinal = 1 And numUpdateToFinal = 1) Then

    ' No longer need complete list of operations, just "B" if multiple operations, or the one operation if one operation.
    For Each vKey In dicEmailOperation
        If Len(dicEmailOperation(vKey)) Then
            dicEmailOperation(vKey) = Left$(dicEmailOperation(vKey), 1)
        End If
    Next ' For Each vKey In dicEmailOperation

    If (numBadEmails > 0) Then
        sSummary = CStr(numBadEmails) & " email addresses in Excel can not be used." & _
            vbCrLf & "The cell for each of them has a comment explaining why."
        MsgBox sSummary, , sThisProgram & sWorksheetName
    End If

    ' Clean up.
    Set sListAdd = Nothing
    Set sListDelete = Nothing
    Set rColUpdateFrom = Nothing
    Set rColUpdateTo = Nothing
    Set dicEmailSource = Nothing

End Sub ' ReadSpreadsheet

Public Function CheckEmailAddressText(ByVal sEmailAddress As String, ByRef rCell As Range, ByRef numBadEmails As Integer) As Boolean

    Dim vArray As Variant
    Dim sLocalPart As String, sMessage As String
    Dim i As Integer, numDnsFields As Integer

    CheckEmailAddressText = True

    ' Ignore empty cells.
    If (sEmailAddress = vbNullString) Then GoTo EmptyCell

    ' CHECK VALIDITY OF EMAIL ADDRESSES.

    ' Split email address into local part (before the "@") and the rest: array with elements 0, 1, etc.
    vArray = Split(sEmailAddress, "@")

    ' Require exactly one "@".  Otherwise, assume not intended to be an email address.
    If (UBound(vArray) <> 1) Then GoTo NotEmailAddress

    ' Require characters before and after "@".  Otherwise, assume not intended to be an email address.
    sLocalPart = vArray(0)
    If (sLocalPart = vbNullString Or vArray(1) = vbNullString) Then GoTo NotEmailAddress

    ' Split up rest of email address into DNS name fields.
    vArray = Split(vArray(1), ".")
    numDnsFields = UBound(vArray)

    ' Require at least one ".".  Otherwise, assume not intended to be an email address.
    If (numDnsFields = 0) Then GoTo NotEmailAddress

    ' Ensure that each DNS name field is valid.
    For i = 0 To numDnsFields
        sMessage = ValidDnsField(ByVal vArray(i))
        If (sMessage <> vbNullString) Then GoTo BadEmailAddress
    Next

    ' Ensure local part of email address (before the "@") is valid.
    sMessage = ValidLocalPart(ByVal sLocalPart)
    If (sMessage <> vbNullString) Then GoTo BadEmailAddress

    If False Then
EmptyCell:
        CheckEmailAddressText = False
    End If

    If False Then
NotEmailAddress:
        CheckEmailAddressText = False
        AddCellComment rCell, "Not an email address."
    End If

    If False Then
BadEmailAddress:
        CheckEmailAddressText = False
        AddCellComment rCell, sMessage
        numBadEmails = numBadEmails + 1
    End If

    ' Nothing to clean up.

End Function ' CheckEmailAddressText

Public Function GetRejectMessage(ByVal sOperationThis As String, ByVal sOperationPrev As String) As String
    Dim sMessage As String, sOperationPrevLast As String
    Dim numAdds As Integer, numDeletes As Integer, numUpdates As Integer

    ' If there are multiple operations for this email address
    sOperationPrevLast = Right$(sOperationPrev, 1)

    numAdds = 0
    numDeletes = 0
    numUpdates = 0

    If (sOperationThis = "A") Then
        numAdds = 1
    ElseIf (sOperationThis = "D") Then
        numDeletes = 1
    ElseIf (sOperationThis = "F" Or sOperationThis = "T") Then
        numUpdates = 1
    End If

    If (sOperationPrevLast = "A") Then
        numAdds = numAdds + 1
    ElseIf (sOperationPrevLast = "D") Then
        numDeletes = numDeletes + 1
    ElseIf (sOperationPrevLast = "F" Or sOperationPrevLast = "T") Then
        numUpdates = numUpdates + 1
    End If

    If (numAdds = 2) Then
        sMessage = "Cannot add email address twice."
    ElseIf (numDeletes = 2) Then
        sMessage = "Cannot delete email address twice."
    ElseIf (numAdds = 1 And numDeletes = 1) Then
        sMessage = "Cannot add and delete same email address."
    ElseIf (numAdds = 1 And numUpdates = 1) Then
        sMessage = "Cannot add and update same email address."
    ElseIf (numDeletes = 1 And numUpdates = 1) Then
        sMessage = "Cannot delete and update same email address."
    End If

    GetRejectMessage = sMessage
End Function ' GetRejectMessage

Public Sub AddDeleteUpdate(ByVal dicEmailOperation As Object, ByVal dicEmailUpdate As Object, _
                           ByRef sSummary As String, ByRef sDestinationList As Collection)
    Dim outlookApp As Object
    Set outlookApp = CreateObject("Outlook.Application")

    Dim objNameSpace As Outlook.Namespace
    Set objNameSpace = outlookApp.GetNamespace("MAPI")

    Dim objSubFolder As Outlook.Folder

    Dim objContactGroup As Outlook.DistListItem

    Dim vItem As Variant

    Dim sFolder As String, sSubFolder As String, sContactGroup As String, sDestination As String
    Dim iImportType As Integer

    For Each vItem In sDestinationList ' vItem is "sDestinationListEntry(0 To 2) As String" in disguise
        ' SET UP FOR NEXT IMPORT DESTINATION.
        sFolder = vItem(0)
        sSubFolder = vItem(1)
        sContactGroup = vItem(2)
        sDestination = "the folder '\\" & sFolder & "\" & sSubFolder & "'."
        If (sContactGroup = vbNullString) Then
            iImportType = iImportFolder
            Set objSubFolder = objNameSpace.Folders(sFolder).Folders(sSubFolder)
        Else
            iImportType = iImportContactGroup
            sDestination = "the contact group '" & sContactGroup & "' in " & sDestination
            Set objContactGroup = objNameSpace.Folders(sFolder).Folders(sSubFolder).Items(sContactGroup)
        End If

        ' ADD EMAIL ADDRESSES
        EmailAddressAdd iImportType, sDestination, sSummary, objSubFolder, objContactGroup, dicEmailOperation

        ' DELETE EMAIL ADDRESSES
        EmailAddressDelete iImportType, sDestination, sSummary, objSubFolder, objContactGroup, dicEmailOperation

        ' UPDATE EMAIL ADDRESSES
        EmailAddressUpdate iImportType, sDestination, sSummary, objSubFolder, objContactGroup, dicEmailOperation, dicEmailUpdate
    Next

    'Selection.Cells.ClearComments

    ' Clean up.
    Set objContactGroup = Nothing
    Set objNameSpace = Nothing
    Set outlookApp = Nothing

End Sub ' AddDeleteUpdate

Public Sub EmailAddressAdd(ByVal iImportType As Integer, ByVal sDestination As String, ByRef sSummary As String, _
                           ByVal objSubFolder As Outlook.Folder, ByVal objContactGroup As Outlook.DistListItem, ByVal dicEmailOperation As Object)
    Dim outlookApp As Object
    Set outlookApp = CreateObject("Outlook.Application")

    Dim objNameSpace As Outlook.Namespace
    Set objNameSpace = outlookApp.GetNamespace("MAPI")

    Dim objContact As Outlook.contactItem
    Dim objItem As Object

    Dim objMailItem As Outlook.MailItem
    Set objMailItem = Outlook.CreateItem(olMailItem)

    Dim objRecipients As Outlook.Recipients
    Set objRecipients = objMailItem.Recipients

    Dim vEmailAddress As Variant ' only variants can iterate thru dictionary.

    Dim sOperation As String, sEmailAddress As String
    Dim i As Integer, iAdded As Integer, iTotal As Integer

    ' FETCH ALL EMAIL ADDRESSES FROM NEXT IMPORT DESTINATION INTO dicExistingEmail.

    ' Existing: already in the selected folder or contact group.
    Dim dicExistingEmail As Object
    Set dicExistingEmail = CreateObject("Scripting.Dictionary")
    dicExistingEmail.CompareMode = vbTextCompare

    dicExistingEmail.RemoveAll
    If (iImportType = iImportFolder) Then
        For Each objItem In objSubFolder.Items
            If (TypeOf objItem Is contactItem) Then
                Set objContact = objItem
                sEmailAddress = objContact.Email1Address
                ' 1st argument is dictionary key, 2nd is dictionary value (here, key=value).
                dicExistingEmail.Add Key:=sEmailAddress, Item:=sEmailAddress
            End If
        Next
    ElseIf (iImportType = iImportContactGroup) Then
        If (objContactGroup.MemberCount > 0) Then
            For i = 1 To objContactGroup.MemberCount
                sEmailAddress = objContactGroup.GetMember(i).Address
                ' 1st argument is dictionary key, 2nd is dictionary value (here, key=value).
                dicExistingEmail.Add Key:=sEmailAddress, Item:=sEmailAddress
            Next
        End If
    End If

    ' Initialize and show progress indicator.
    iAdded = 0
    iTotal = 0
    sOperation = "Added"
    frmProgress.Caption = sProgressCaption
    frmProgress.Show
    ShowProgress sOperation, iAdded, iTotal, "to " & sDestination

    For Each vEmailAddress In dicEmailOperation.Keys
        If (dicEmailOperation(vEmailAddress) = "A") Then
            iTotal = iTotal + 1
            If Not dicExistingEmail.Exists(vEmailAddress) Then
                iAdded = iAdded + 1
                If (iImportType = iImportFolder) Then
                    Set objContact = objSubFolder.Items.Add
                    objContact.Email1Address = vEmailAddress
                    objContact.Email1AddressType = "SMTP"
                    objContact.Save
                ElseIf (iImportType = iImportContactGroup) Then
                    objRecipients.Add vEmailAddress
                    ' Add emails in batches of iBatchSize.
                    If (iAdded Mod iBatchSize = 0) Then
                        ' Show progress.
                        ShowProgress sOperation, iAdded, iTotal, "to " & sDestination
                        ' Add batch
                        objContactGroup.AddMembers objRecipients
                        objContactGroup.Close (olSave)
                        ' Initialize for next batch.
                        For i = objRecipients.Count To 1 Step -1
                            objRecipients.Remove (i)
                        Next
                    End If
                End If
                If (iAdded Mod iBatchSize = 0) Then ShowProgress sOperation, iAdded, iTotal, "to " & sDestination
            End If
        End If
    Next
    ' Save last batch.
    If (iImportType = iImportContactGroup) Then
        If (objRecipients.Count > 0) Then
            objContactGroup.AddMembers objRecipients
            objContactGroup.Close (olSave)
        End If
    End If

    frmProgress.Hide
    sSummary = sSummary & vbCrLf & vbCrLf & ProgressString(sOperation, iAdded, iTotal, "to " & sDestination)

    ' Clean up.
    Set dicExistingEmail = Nothing
    Set objRecipients = Nothing
    Set objMailItem = Nothing
    Set objContact = Nothing

    Set objNameSpace = Nothing
    Set outlookApp = Nothing

End Sub ' EmailAddressAdd

Public Sub EmailAddressDelete(ByVal iImportType As Integer, ByVal sDestination As String, ByRef sSummary As String, _
                              ByVal objSubFolder As Outlook.Folder, ByVal objContactGroup As Outlook.DistListItem, ByVal dicEmailOperation As Object)
    Dim outlookApp As Object
    Set outlookApp = CreateObject("Outlook.Application")

    Dim objNameSpace As Outlook.Namespace
    Set objNameSpace = outlookApp.GetNamespace("MAPI")

    Dim objContact As Outlook.contactItem

    Dim objMailItem As Outlook.MailItem
    Set objMailItem = Outlook.CreateItem(olMailItem)

    Dim objRecipients As Outlook.Recipients
    Set objRecipients = objMailItem.Recipients
    Dim objRecipient As Outlook.Recipient

    Dim vEmailAddress As Variant ' only variants can iterate thru dictionary.

    Dim sOperation As String
    Dim i As Integer, iDeleted As Integer, iTotal As Integer

    ' Initialize and show progress indicator.
    iDeleted = 0
    iTotal = 0
    sOperation = "Deleted"
    frmProgress.Caption = sProgressCaption
    frmProgress.Show
    ShowProgress sOperation, iDeleted, iTotal, "from " & sDestination

    If (iImportType = iImportFolder) Then
        ' Start from end.  See http://www.vbaexpress.com/forum/showthread.php?38566-Solved-Delete-contacts-help.
        For i = objSubFolder.Items.Count To 1 Step -1
            If (TypeOf objSubFolder.Items(i) Is contactItem) Then
                Set objContact = objSubFolder.Items(i)
                vEmailAddress = objContact.Email1Address
                If dicEmailOperation.Exists(vEmailAddress) Then
                    iTotal = iTotal + 1
                    If (dicEmailOperation(vEmailAddress) = "D") Then
                        iDeleted = iDeleted + 1
                        objContact.Delete
                        If (iDeleted Mod iBatchSize = 0) Then ShowProgress sOperation, iDeleted, iTotal, "from " & sDestination
                    End If
                End If
            End If
        Next
    ElseIf (iImportType = iImportContactGroup) Then
        ' Start from end.  See http://www.vbaexpress.com/forum/showthread.php?38566-Solved-Delete-contacts-help.
        For i = objContactGroup.MemberCount To 1 Step -1
            vEmailAddress = objContactGroup.GetMember(i).Address
            If dicEmailOperation.Exists(vEmailAddress) Then
                iTotal = iTotal + 1
                If (dicEmailOperation(vEmailAddress) = "D") Then
                    iDeleted = iDeleted + 1
                    Set objRecipient = objMailItem.Recipients.Add(Name:=vEmailAddress)
                    objRecipient.Resolve
                    objContactGroup.RemoveMember Recipient:=objRecipient
                    objContactGroup.Save
                    ' Delete email addresses from objRecipients, otherwise they build up and slow us down.
                    objMailItem.Recipients.Remove (1)
                    If (iDeleted Mod iBatchSize = 0) Then ShowProgress sOperation, iDeleted, iTotal, "from " & sDestination
                End If
            End If
        Next
    End If
    frmProgress.Hide
    sSummary = sSummary & vbCrLf & vbCrLf & ProgressString(sOperation, iDeleted, iTotal, "from " & sDestination)

    ' Clean up.
    Set objRecipient = Nothing
    Set objRecipients = Nothing
    Set objMailItem = Nothing
    Set objContact = Nothing

    Set objNameSpace = Nothing
    Set outlookApp = Nothing

End Sub ' EmailAddressDelete

Public Sub EmailAddressUpdate(ByVal iImportType As Integer, ByVal sDestination As String, ByRef sSummary As String, _
                              ByVal objSubFolder As Outlook.Folder, ByVal objContactGroup As Outlook.DistListItem, ByVal dicEmailOperation As Object, _
                              ByVal dicEmailUpdate As Object)
    Dim outlookApp As Object
    Set outlookApp = CreateObject("Outlook.Application")

    Dim objNameSpace As Outlook.Namespace
    Set objNameSpace = outlookApp.GetNamespace("MAPI")

    Dim objContact As Outlook.contactItem

    Dim objMailItem As Outlook.MailItem
    Set objMailItem = Outlook.CreateItem(olMailItem)

    Dim objRecipients As Outlook.Recipients
    Set objRecipients = objMailItem.Recipients
    Dim objRecipient As Outlook.Recipient

    Dim vEmailAddress As Variant ' only variants can iterate thru dictionary.

    Dim sOperation As String
    Dim i As Integer, iUpdated As Integer, iTotal As Integer

    ' Initialize and show progress indicator.
    iUpdated = 0
    iTotal = 0
    sOperation = "Updated"
    frmProgress.Caption = sProgressCaption
    frmProgress.Show
    ShowProgress sOperation, iUpdated, iTotal, "in " & sDestination

    If (iImportType = iImportFolder) Then
        ' Start from end.  See http://www.vbaexpress.com/forum/showthread.php?38566-Solved-Delete-contacts-help.
        For i = objSubFolder.Items.Count To 1 Step -1
            If (TypeOf objSubFolder.Items(i) Is contactItem) Then
                Set objContact = objSubFolder.Items(i)
                vEmailAddress = objContact.Email1Address
                If dicEmailOperation.Exists(vEmailAddress) Then
                    iTotal = iTotal + 1
                    If (dicEmailOperation(vEmailAddress) = "F") Then
                        iUpdated = iUpdated + 1
                        objContact.Email1Address = dicEmailUpdate(vEmailAddress)
                        objContact.Save
                        If (iUpdated Mod iBatchSize = 0) Then ShowProgress sOperation, iUpdated, iTotal, "in " & sDestination
                    End If
                End If
            End If
        Next
    ElseIf (iImportType = iImportContactGroup) Then
        ' Start from end.  See http://www.vbaexpress.com/forum/showthread.php?38566-Solved-Delete-contacts-help.
        For i = objContactGroup.MemberCount To 1 Step -1
            vEmailAddress = objContactGroup.GetMember(i).Address
            If dicEmailOperation.Exists(vEmailAddress) Then
                iTotal = iTotal + 1
                If (dicEmailOperation(vEmailAddress) = "F") Then
                    ' Contact Group emails not updateable.  Must remove old email, add new email.
                    iUpdated = iUpdated + 1
                    If (iUpdated Mod iBatchSize = 0) Then ShowProgress sOperation, iUpdated, iTotal, "in " & sDestination

                    ' Remove old email address.
                    Set objRecipient = objMailItem.Recipients.Add(Name:=vEmailAddress)
                    objRecipient.Resolve
                    objContactGroup.RemoveMember Recipient:=objRecipient
                    objContactGroup.Save
                    ' Delete email addresses from objRecipients, otherwise they build up and slow us down.
                    objMailItem.Recipients.Remove (1)

                    ' Add new email address.
                    Set objRecipient = objMailItem.Recipients.Add(Name:=dicEmailUpdate(vEmailAddress))
                    objRecipient.Resolve
                    objContactGroup.AddMember Recipient:=objRecipient
                    objContactGroup.Save
                    ' Delete email addresses from objRecipients, otherwise they build up and slow us down.
                    objMailItem.Recipients.Remove (1)
                End If
            End If
        Next
    End If

    frmProgress.Hide
    sSummary = sSummary & vbCrLf & vbCrLf & ProgressString(sOperation, iUpdated, iTotal, "in " & sDestination)

    ' Clean up.
    Set objRecipient = Nothing
    Set objRecipients = Nothing
    Set objMailItem = Nothing
    Set objContact = Nothing

    Set objNameSpace = Nothing
    Set outlookApp = Nothing

End Sub ' EmailAddressUpdate

Public Sub ShowProgress(ByVal sOperation As String, ByVal iProcessed As Integer, ByVal iTotal As Integer, ByVal sDestination As String)
    frmProgress.lblProgress.Caption = ProgressString(sOperation, iProcessed, iTotal, sDestination)
    DoEvents
End Sub ' ShowProgress

Public Function ProgressString(ByVal sOperation As String, ByVal iProcessed As Integer, ByVal iTotal As Integer, ByVal sDestination As String)
    ProgressString = sOperation & " " & Format$(iProcessed, "#,##0") & " of " & Format$(iTotal, "#,##0") & _
                                      " email addresses " & sDestination
End Function ' ProgressString

Public Sub AddCellComment(ByRef rCell As Range, ByVal sComment As String)
    Dim sCommentNew As String
    If (rCell.Comment Is Nothing) Then
        rCell.AddComment sComment
    Else
        sCommentNew = rCell.Comment.Text & vbCrLf & sComment
        rCell.ClearComments
        rCell.AddComment sCommentNew
    End If
End Sub ' AddCellComment

Public Function DescriptionChar(ByVal iAsc As Integer) As String
    If (iAsc <= 31) Then
        DescriptionChar = "non-printable character with ascii code " & CStr(iAsc)
    ElseIf (iAsc = 32) Then
        DescriptionChar = "space"
    ElseIf (iAsc = 34) Then
        DescriptionChar = """"
    ElseIf (iAsc = 127) Then
        DescriptionChar = "delete character"
    Else
        DescriptionChar = Chr$(iAsc)
    End If
End Function ' DescriptionChar

Public Function ValidDnsField(ByVal sDnsField As String) As String
    'https://en.wikipedia.org/wiki/Email_address has description of valid email addresses.

    Dim i As Integer, iAsc As Integer

    ' Valid characters:
    ' hyphen (ascii character 45), but not the first or last character.
    ' 0-9 (ascii characters 48-57)
    ' A-Z (ascii characters 65-90)
    ' a-z (ascii characters 97-122)

    ' Require at least one character.
    If (Len(sDnsField) = 0) Then
        ValidDnsField = "Illegal characters: '..' after @."
        Exit Function
    End If

    ' Special rules for hyphens.
    If (Left$(sDnsField, 1) = "-") Then
        ValidDnsField = "Illegal: '-' at start of DNS field."
        Exit Function
    ElseIf (Right$(sDnsField, 1) = "-") Then
        ValidDnsField = "Illegal: '-' at end of DNS field."
        Exit Function
    End If

    ' Look for invalid characters.
    For i = 1 To Len(sDnsField)
        iAsc = Asc(Mid$(sDnsField, i, 1))
        If (iAsc < 45) Then
            ValidDnsField = "Illegal " & DescriptionChar(iAsc) & " after @."
            Exit Function
        ElseIf (iAsc > 45 And iAsc < 48) Then
            ValidDnsField = "Illegal " & DescriptionChar(iAsc) & " after @."
            Exit Function
        ElseIf (iAsc > 57 And iAsc < 65) Then
            ValidDnsField = "Illegal " & DescriptionChar(iAsc) & " after @."
            Exit Function
        ElseIf (iAsc > 90 And iAsc < 97) Then
            ValidDnsField = "Illegal " & DescriptionChar(iAsc) & " after @."
            Exit Function
        ElseIf (iAsc > 122) Then
            ValidDnsField = "Illegal " & DescriptionChar(iAsc) & " after @."
            Exit Function
        End If
    Next

    ' Nothing invalid found
    ValidDnsField = vbNullString

    ' Nothing to clean up.

End Function ' ValidDnsField

Public Function ValidLocalPart(ByVal sLocalPart As String) As String
    'https://en.wikipedia.org/wiki/Email_address has description of valid email addresses.

    ' Valid (True) and invalid (False) characters.
    Dim boolValid(0 To 127) As Boolean
    boolValid(0) = False  'NUL (null)
    boolValid(1) = False  'SOH (start of heading)
    boolValid(2) = False  'STX (start of text)
    boolValid(3) = False  'ETX (end of text)
    boolValid(4) = False  'EOT (end of transmission)
    boolValid(5) = False  'ENQ (enquiry)
    boolValid(6) = False  'ACK (acknowledge)
    boolValid(7) = False  'BEL (bell)
    boolValid(8) = False  'BS  (backspace)
    boolValid(9) = False  'TAB (horizontal tab)
    boolValid(10) = False 'LF  (NL line feed, new line)
    boolValid(11) = False 'VT  (vertical tab)
    boolValid(12) = False 'FF  (NP form feed, new page)
    boolValid(13) = False 'CR  (carriage return)
    boolValid(14) = False 'SO  (shift out)
    boolValid(15) = False 'SI  (shift in)
    boolValid(16) = False 'DLE (data link escape)
    boolValid(17) = False 'DC1 (device control 1)
    boolValid(18) = False 'DC2 (device control 2)
    boolValid(19) = False 'DC3 (device control 3)
    boolValid(20) = False 'DC4 (device control 4)
    boolValid(21) = False 'NAK (negative acknowledge)
    boolValid(22) = False 'SYN (synchronous idle)
    boolValid(23) = False 'ETB (end of trans. block)
    boolValid(24) = False 'CAN (cancel)
    boolValid(25) = False 'EM  (end of medium)
    boolValid(26) = False 'SUB (substitute)
    boolValid(27) = False 'ESC (escape)
    boolValid(28) = False 'FS  (file separator)
    boolValid(29) = False 'GS  (group separator)
    boolValid(30) = False 'RS  (record separator)
    boolValid(31) = False 'US  (unit separator)
    boolValid(32) = False 'SPACE
    boolValid(33) = True  '!
    boolValid(34) = False '"
    boolValid(35) = True  '#
    boolValid(36) = True  '$
    boolValid(37) = True  '%
    boolValid(38) = True  '&
    boolValid(39) = True  '' (single quote)
    boolValid(40) = False '(
    boolValid(41) = False ')
    boolValid(42) = True  '*
    boolValid(43) = True  '+
    boolValid(44) = False ',
    boolValid(45) = True  '-
    boolValid(46) = True  '.
    boolValid(47) = True  '/
    boolValid(48) = True  '0
    boolValid(49) = True  '1
    boolValid(50) = True  '2
    boolValid(51) = True  '3
    boolValid(52) = True  '4
    boolValid(53) = True  '5
    boolValid(54) = True  '6
    boolValid(55) = True  '7
    boolValid(56) = True  '8
    boolValid(57) = True  '9
    boolValid(58) = False ':
    boolValid(59) = False ';
    boolValid(60) = False '<
    boolValid(61) = True  '=
    boolValid(62) = False '>
    boolValid(63) = True  '?
    boolValid(64) = False '@
    boolValid(65) = True  'A
    boolValid(66) = True  'B
    boolValid(67) = True  'C
    boolValid(68) = True  'D
    boolValid(69) = True  'E
    boolValid(70) = True  'F
    boolValid(71) = True  'G
    boolValid(72) = True  'H
    boolValid(73) = True  'I
    boolValid(74) = True  'J
    boolValid(75) = True  'K
    boolValid(76) = True  'L
    boolValid(77) = True  'M
    boolValid(78) = True  'N
    boolValid(79) = True  'O
    boolValid(80) = True  'P
    boolValid(81) = True  'Q
    boolValid(82) = True  'R
    boolValid(83) = True  'S
    boolValid(84) = True  'T
    boolValid(85) = True  'U
    boolValid(86) = True  'V
    boolValid(87) = True  'W
    boolValid(88) = True  'X
    boolValid(89) = True  'Y
    boolValid(90) = True  'Z
    boolValid(91) = False '[
    boolValid(92) = False '\
    boolValid(93) = False ']
    boolValid(94) = True  '^
    boolValid(95) = True  '_
    boolValid(96) = True  '`
    boolValid(97) = True  'a
    boolValid(98) = True  'b
    boolValid(99) = True  'c
    boolValid(100) = True 'd
    boolValid(101) = True 'e
    boolValid(102) = True 'f
    boolValid(103) = True 'g
    boolValid(104) = True 'h
    boolValid(105) = True 'i
    boolValid(106) = True 'j
    boolValid(107) = True 'k
    boolValid(108) = True 'l
    boolValid(109) = True 'm
    boolValid(110) = True 'n
    boolValid(111) = True 'o
    boolValid(112) = True 'p
    boolValid(113) = True 'q
    boolValid(114) = True 'r
    boolValid(115) = True 's
    boolValid(116) = True 't
    boolValid(117) = True 'u
    boolValid(118) = True 'v
    boolValid(119) = True 'w
    boolValid(120) = True 'x
    boolValid(121) = True 'y
    boolValid(122) = True 'z
    boolValid(123) = True '{
    boolValid(124) = True '|
    boolValid(125) = True '}
    boolValid(126) = True '~
    boolValid(127) = False 'DEL

    ' Special rules for periods: not first or character, and never two in a row.
    If (Left$(sLocalPart, 1) = ".") Then
        ValidLocalPart = "Illegal: first character is a period."
        Exit Function
    ElseIf (Right$(sLocalPart, 1) = ".") Then
        ValidLocalPart = "Illegal: '.@'."
        Exit Function
    ElseIf (InStr(sLocalPart, "..") > 0) Then
        ValidLocalPart = "Illegal: '..' before @."
        Exit Function
    End If

    ' Look for invalid characters.
    Dim i As Integer, iAsc As Integer
    For i = 1 To Len(sLocalPart)
        iAsc = Asc(Mid$(sLocalPart, i, 1))
        If (Not boolValid(iAsc)) Then
            ValidLocalPart = "Illegal " & DescriptionChar(iAsc) & " before @."
            Exit Function
        End If
    Next

    ' Nothing invalid found
    ValidLocalPart = vbNullString

    ' Nothing to clean up.

End Function ' ValidLocalPart

