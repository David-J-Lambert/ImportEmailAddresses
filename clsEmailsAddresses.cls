VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEmailsAddresses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Module clsEmailAddresses.
' Author: David Lambert, David5Lambert7@Gmail.com.
' MICROSOFT OFFICE SETUP: The spreadsheet "INSTRUCTIONS" has macro setup instructions.

Option Explicit

' PRIVATE PROPERTIES

' Dictionary of all email addresses. Key: email address.
' Value: operation(s) ("A"=add, "D"=delete, "F"=update from, "T"=update to, "B"=banned because spreadsheet lists multiple operations).
Private dicEmailOperation As Object

' Dictionary of all email address updates. Key and Value: email addresses.
' For "F", key=update from address, value=update to address.  For "T", key=update to address, value=update from address.
Private dicEmailUpdate As Object

' Other Objects.
Private oSummary As New clsSummary
Private oDestinations As New clsDestinations

' Module-level variables.
Private sProgressCaption As String
Private sFinalCaption As String

' Outlook Objects
Private outlookApp As Object
Private objNameSpace As Outlook.Namespace
Private objFolders As Outlook.Folders
Private objSubFolder As Outlook.Folder
Private objContactGroup As Outlook.DistListItem
Private objContact As Outlook.contactItem
Private objMailItem As Outlook.MailItem
Private objRecipients As Outlook.Recipients
Private objRecipient As Outlook.Recipient

' END OF PRIVATE PROPERTIES
' PRIVATE FUNCTIONS AND SUBROUTINES

Private Sub Class_Initialize()
    Set dicEmailOperation = CreateObject("Scripting.Dictionary")
    dicEmailOperation.CompareMode = vbTextCompare

    Set dicEmailUpdate = CreateObject("Scripting.Dictionary")
    dicEmailUpdate.CompareMode = vbTextCompare

    Set objMailItem = Outlook.CreateItem(olMailItem)
    Set objRecipients = objMailItem.Recipients

    Set outlookApp = CreateObject("Outlook.Application")
    Set objNameSpace = outlookApp.GetNamespace("MAPI")
    Set objFolders = objNameSpace.Folders

End Sub ' Class_Initialize()

Public Sub Class_Terminate()
    Set dicEmailOperation = Nothing
    Set dicEmailUpdate = Nothing

    Set objContactGroup = Nothing
    Set objSubFolder = Nothing
    Set objContact = Nothing

    Set objRecipients = Nothing
    Set objMailItem = Nothing

    Set objFolders = Nothing
    Set objNameSpace = Nothing
    Set outlookApp = Nothing

    End
End Sub ' Class_Terminate()

' END OF PRIVATE FUNCTIONS AND SUBROUTINES
' PUBLIC FUNCTIONS AND SUBROUTINES RELATED TO GETTING EMAIL ADDRESSES

Public Sub GetEmailAddressesFromSpreadsheet()
    Dim numBadAddresses As Integer, iFirstRow As Integer, iRowNumber As Integer, numUpdates As Integer
    Dim numUpdateFrom As Integer, numUpdateTo As Integer, numUpdateFromFinal As Integer, numUpdateToFinal As Integer
    Dim vArray As Variant, vItem As Variant, vKey As Variant
    Dim sCol As String, sColTitle As String, sMessage As String, sEmailAddress As String, sUpdateFrom As String, sUpdateTo As String
    Dim ValidEmailAddress As Boolean, ValidEmailAddressTo As Boolean

    Dim rColumn As Range, rColUpdateFrom As Range, rColUpdateTo As Range, rCell As Range, rCellFrom As Range, rCellTo As Range

    ' Dictionary of the Excel cell each email address comes from. Key: email addresses.  Value: cell address (range data type).
    Dim dicEmailSource As Object
    Set dicEmailSource = CreateObject("Scripting.Dictionary")
    dicEmailSource.CompareMode = vbTextCompare

    sFinalCaption = sThisProgram & " (worksheet '" & ActiveSheet.Name & "')"
    sProgressCaption = sFinalCaption & " is running..."

    ' Enforce no more than 1 of each of these 4 variables.
    numUpdateFrom = 0
    numUpdateTo = 0
    numUpdateFromFinal = 0
    numUpdateToFinal = 0

    ' Collections to store columns to Add and Delete.
    Dim sListAdd As Collection, sListDelete As Collection
    Set sListAdd = New Collection
    Set sListDelete = New Collection

    ' Select active sheet in this Excel workbook.
    ActiveSheet.UsedRange.Select

    ' Extract valid email addresses from active sheet.
    numBadAddresses = 0

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
                MsgBox "More than one column has 'UPDATE FROM' in row 1, for email addresses to be updated.  Quitting.", , sFinalCaption
                End
            End If
            frmAddDeleteUpdate.rbUpdateFrom.Value = True
        ElseIf (InStr(sColTitle, "UPDATE TO") > 0) Then
            numUpdateTo = numUpdateTo + 1
            If (numUpdateTo > 1) Then
                MsgBox "More than one column has 'UPDATE TO' in row 1, for new email addresses to update to.  Quitting.", , sFinalCaption
                End
            End If
            frmAddDeleteUpdate.rbUpdateTo.Value = True
        Else ' "IGNORE" or anything besides the strings above.
            frmAddDeleteUpdate.rbIgnore.Value = True
        End If

        ' GET EMAIL ADDRESS OPERATION DESIRED FOR THIS COLUMN: ADD, DELETE, UPDATE FROM, UPDATE TO.

        frmAddDeleteUpdate.lblColumn.Caption = sCol
        frmAddDeleteUpdate.Caption = sFinalCaption
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
                MsgBox "More than one column contains email addresses to be updated.  Quitting.", , sFinalCaption
                End
            End If
        ElseIf frmAddDeleteUpdate.rbUpdateTo.Value Then
            Set rColUpdateTo = rColumn
            numUpdateToFinal = numUpdateToFinal + 1
            If (numUpdateToFinal > 1) Then
                MsgBox "More than one column contains new email addresses to update to.  Quitting.", , sFinalCaption
                End
            End If
        End If
    Next ' For Each rColumn In Selection.Columns

    ' READ CONTENTS OF CELLS IN "ADD" COLUMNS.
    For Each vItem In sListAdd ' vItem is rColumn in disguise
        For Each rCell In vItem.Cells
            ' Cell text.
            sEmailAddress = rCell.Value
            If (sEmailAddress <> Trim$(sEmailAddress)) Then
                sEmailAddress = Trim$(sEmailAddress)
                rCell.Value = sEmailAddress
            End If

            ValidEmailAddress = CheckEmailAddressText(sEmailAddress, rCell, numBadAddresses)
            If ValidEmailAddress Then
                If dicEmailOperation.Exists(sEmailAddress) Then
                    sMessage = GetRejectMessage("A", dicEmailOperation(sEmailAddress))
                    If Not IsEmpty(dicEmailSource(sEmailAddress)) Then _
                        AddCellComment dicEmailSource(sEmailAddress), sMessage
                    AddCellComment rCell, sMessage
                    dicEmailOperation(sEmailAddress) = "B" & dicEmailOperation(sEmailAddress)
                    numBadAddresses = numBadAddresses + 1
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
            sEmailAddress = rCell.Value
            If (sEmailAddress <> Trim$(sEmailAddress)) Then
                sEmailAddress = Trim$(sEmailAddress)
                rCell.Value = sEmailAddress
            End If

            ValidEmailAddress = CheckEmailAddressText(sEmailAddress, rCell, numBadAddresses)
            If ValidEmailAddress Then
                If dicEmailOperation.Exists(sEmailAddress) Then
                    sMessage = GetRejectMessage("D", dicEmailOperation(sEmailAddress))
                    If Not IsEmpty(dicEmailSource(sEmailAddress)) Then _
                        AddCellComment dicEmailSource(sEmailAddress), sMessage
                    AddCellComment rCell, sMessage
                    dicEmailOperation(sEmailAddress) = "B" & dicEmailOperation(sEmailAddress)
                    numBadAddresses = numBadAddresses + 1
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
            MsgBox "The two update columns are not the same height!  Quitting.", , sFinalCaption
            End
        End If
        ' Walk down both columns at the same time.
        numUpdates = rColUpdateFrom.Cells.Count
        iFirstRow = rColUpdateFrom.Cells(1, 1).Row
        For iRowNumber = iFirstRow To (iFirstRow + numUpdates)
            ' Treat both cells as a unit, not separately.
            Set rCellFrom = rColUpdateFrom.Cells(iRowNumber, 1)
            sUpdateFrom = rCellFrom.Value
            If (sUpdateFrom <> Trim$(sUpdateFrom)) Then
                sUpdateFrom = Trim$(sUpdateFrom)
                rCellFrom.Value = sUpdateFrom
            End If
            ValidEmailAddress = CheckEmailAddressText(sUpdateFrom, rCellFrom, numBadAddresses)

            Set rCellTo = rColUpdateTo.Cells(iRowNumber, 1)
            sUpdateTo = rCellTo.Value
            If (sUpdateTo <> Trim$(sUpdateTo)) Then
                sUpdateTo = Trim$(sUpdateTo)
                rCellTo.Value = sUpdateTo
            End If
            ValidEmailAddressTo = CheckEmailAddressText(sUpdateTo, rCellTo, numBadAddresses)
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
                    numBadAddresses = numBadAddresses + 1 ' sUpdateFrom recorded 2x, once for "T", once for "F".
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
                    numBadAddresses = numBadAddresses + 1 ' sUpdateFrom recorded 2x, once for "T", once for "F".
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
                    numBadAddresses = numBadAddresses + 1 ' sUpdateFrom recorded 2x, once for "T", once for "F".
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
                numBadAddresses = numBadAddresses + 1
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
                numBadAddresses = numBadAddresses + 1
            Else
                ' Do nothing, both cells have comments.
            End If
        Next ' For iRowNumber = iFirstRow To (iFirstRow + numUpdates)
    ElseIf (numUpdateFromFinal + numUpdateToFinal > 0) Then
        MsgBox "Found an 'UPDATE FROM' column without an 'UPDATE TO' column, or vice versa.  Quitting.", , sFinalCaption
        End
    End If ' If (numUpdateFromFinal = 1 And numUpdateToFinal = 1) Then

    ' No longer need complete list of operations, just "B" if multiple operations, or the one operation if one operation.
    For Each vKey In dicEmailOperation
        If Len(dicEmailOperation(vKey)) Then
            dicEmailOperation(vKey) = Left$(dicEmailOperation(vKey), 1)
        End If
    Next ' For Each vKey In dicEmailOperation

    If (numBadAddresses > 0) Then
        oSummary.AddToSummary CStr(numBadAddresses) & " email addresses in Excel can not be used." & vbCrLf & "The cell for each of them has a comment explaining why."
        oSummary.PrintSummary sFinalCaption
    End If

    ' Clean up.
    Set sListAdd = Nothing
    Set sListDelete = Nothing
    Set rColUpdateFrom = Nothing
    Set rColUpdateTo = Nothing
    Set dicEmailSource = Nothing
End Sub ' GetEmailAddressesFromSpreadsheet(ByRef oSummary As clsSummary)

' END OF PUBLIC FUNCTIONS AND SUBROUTINES RELATED TO GETTING EMAIL ADDRESSES
' PUBLIC FUNCTIONS AND SUBROUTINES RELATED TO ADDING, DELETING, AND UPDATING EMAIL ADDRESSES

Public Sub AddDeleteUpdate()

    Dim vItem As Variant
    Dim sFolder As String, sSubFolder As String, sContactGroup As String, sDestination As String
    Dim iDestinationType As Integer

    ' ADD, REMOVE, AND UPDATE EMAIL DESTINATIONS.
    For Each vItem In oDestinations.GetDestinations ' vItem is "sDestinationListEntry(0 To 2) As String" in disguise
        ' SET UP FOR NEXT DESTINATION.
        sFolder = vItem(0)
        sSubFolder = vItem(1)
        sContactGroup = vItem(2)
        sDestination = "the folder '\\" & sFolder & "\" & sSubFolder & "'."
        If (sContactGroup = vbNullString) Then
            iDestinationType = iDestinationFolder
            Set objSubFolder = objNameSpace.Folders(sFolder).Folders(sSubFolder)
        Else
            iDestinationType = iDestinationContactGroup
            sDestination = "the contact group '" & sContactGroup & "' in " & sDestination
            Set objContactGroup = objNameSpace.Folders(sFolder).Folders(sSubFolder).Items(sContactGroup)
        End If

        ' ADD EMAIL ADDRESSES
        EmailAddressAdd iDestinationType, sDestination

        ' DELETE EMAIL ADDRESSES
        EmailAddressDelete iDestinationType, sDestination

        ' UPDATE EMAIL ADDRESSES
        EmailAddressUpdate iDestinationType, sDestination
    Next

    'Selection.Cells.ClearComments

End Sub ' AddDeleteUpdate()

Public Sub EmailAddressAdd(ByVal iDestinationType As Integer, ByVal sDestination As String)
    Dim objItem As Object

    Dim vEmailAddress As Variant ' only variants can iterate thru dictionary.

    Dim sOperation As String, sEmailAddress As String
    Dim i As Integer, iAdded As Integer, iTotal As Integer

    ' FETCH ALL EMAIL ADDRESSES FROM NEXT DESTINATION INTO dicExistingEmail.

    ' Existing: already in the selected folder or contact group.
    Dim dicExistingEmail As Object
    Set dicExistingEmail = CreateObject("Scripting.Dictionary")
    dicExistingEmail.CompareMode = vbTextCompare

    dicExistingEmail.RemoveAll
    If (iDestinationType = iDestinationFolder) Then
        For Each objItem In objSubFolder.Items
            If (TypeOf objItem Is contactItem) Then
                Set objContact = objItem
                sEmailAddress = objContact.Email1Address
                ' Folders allow multiple contacts with the same email address.
                If Not dicExistingEmail.Exists(sEmailAddress) Then
                    ' 1st argument is dictionary key, 2nd is dictionary value (here, key=value).
                    dicExistingEmail.Add Key:=sEmailAddress, Item:=sEmailAddress
                End If
            End If
        Next
    ElseIf (iDestinationType = iDestinationContactGroup) Then
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
                If (iDestinationType = iDestinationFolder) Then
                    Set objContact = objSubFolder.Items.Add
                    objContact.Email1Address = vEmailAddress
                    objContact.Email1AddressType = "SMTP"
                    objContact.Save
                ElseIf (iDestinationType = iDestinationContactGroup) Then
                    objRecipients.Add vEmailAddress
                    ' Add email addresses in batches of iBatchSize.
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
    If (iDestinationType = iDestinationContactGroup) Then
        If (objRecipients.Count > 0) Then
            objContactGroup.AddMembers objRecipients
            objContactGroup.Close (olSave)
        End If
    End If

    frmProgress.Hide
    oSummary.AddToSummary ProgressString(sOperation, iAdded, iTotal, "to " & sDestination)

    ' Clean up.
    Set dicExistingEmail = Nothing
    Set objItem = Nothing

End Sub ' EmailAddressAdd

Public Sub EmailAddressDelete(ByVal iDestinationType As Integer, ByVal sDestination As String)
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

    If (iDestinationType = iDestinationFolder) Then
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
    ElseIf (iDestinationType = iDestinationContactGroup) Then
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
    oSummary.AddToSummary ProgressString(sOperation, iDeleted, iTotal, "from " & sDestination)

End Sub ' EmailAddressDelete

Public Sub EmailAddressUpdate(ByVal iDestinationType As Integer, ByVal sDestination As String)

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

    If (iDestinationType = iDestinationFolder) Then
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
    ElseIf (iDestinationType = iDestinationContactGroup) Then
        ' Start from end.  See http://www.vbaexpress.com/forum/showthread.php?38566-Solved-Delete-contacts-help.
        For i = objContactGroup.MemberCount To 1 Step -1
            vEmailAddress = objContactGroup.GetMember(i).Address
            If dicEmailOperation.Exists(vEmailAddress) Then
                iTotal = iTotal + 1
                If (dicEmailOperation(vEmailAddress) = "F") Then
                    ' Contact Group email addresses not updateable.  Must remove old address, then add new one.
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
    oSummary.AddToSummary ProgressString(sOperation, iUpdated, iTotal, "in " & sDestination)

End Sub ' EmailAddressUpdate

' END OF PUBLIC FUNCTIONS AND SUBROUTINES RELATED TO ADDING, DELETING, AND UPDATING EMAIL ADDRESSES
' ACCESS OF OTHER OBJECTS

Public Sub PrintSummary()
    oSummary.PrintSummary sFinalCaption
End Sub

Public Sub SetDestinationsInUI()
    oDestinations.SetDestinationsInUI "Add, Delete, or Update Addresses Now", "A"
End Sub

