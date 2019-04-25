Attribute VB_Name = "ManageEmailAddresses"
' Module ManageEmailAddresses.
' Author: David Lambert, David5Lambert7@Gmail.com.
' MICROSOFT OFFICE SETUP: The spreadsheet "INSTRUCTIONS" has macro setup instructions.

Option Explicit

Public Const sThisProgram As String = "ManageEmailAddresses"
Public Const sThisVersion As String = "Version 1.0"
Public Const iDestinationFolder As Integer = 0
Public Const iDestinationContactGroup As Integer = 1
Public Const iBatchSize As Integer = 50

' Purposely breaking encapsulation of frmSelectDestinations.
Public boolSelectionMade As Boolean

Public Sub ManageEmailAddresses()
    ' MAIN PROGRAM, DRIVES EVERYTHING ELSE.

    ' Create objects.
    Dim oEmailsAddresses As New clsEmailsAddresses

    ' CHOOSE WHICH SUBFOLDERS AND CONTACT GROUPS TO WORK WITH.
    oEmailsAddresses.SetDestinationsInUI

    ' READ EMAIL ADDRESSES FROM ACTIVE SPREADSHEET.
    oEmailsAddresses.GetEmailAddressesFromSpreadsheet

    ' DO THE ADDS, DELETES, AND UPDATES.
    oEmailsAddresses.AddDeleteUpdate

    oEmailsAddresses.PrintSummary

    Set oEmailsAddresses = Nothing
End Sub

