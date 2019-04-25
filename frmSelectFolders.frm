VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectFolders 
   ClientHeight    =   9295.001
   ClientLeft      =   39
   ClientTop       =   325
   ClientWidth     =   10660
   OleObjectBlob   =   "frmSelectFolders.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSelectFolders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Me.Hide
    MsgBox "At your request, quitting.", , sThisProgram
    End
End Sub

Private Sub cmdSelectFolders_Click()
    Dim i As Integer
    With Me
        For i = 0 To .lstFolders.ListCount - 1
            If .lstFolders.Selected(i) Then
                ' Button exits Form only if something selected.
                .Hide
                boolSelectionMade = True
                Exit Sub
            End If
        Next
        If boolSelectionMade = False Then .cmdSelectFolders.Caption = "Select something, Doofus."
    End With
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Make clicking the red X Close button in the top right corner of this form the same as pressing the Cancel button.
    cmdCancel_Click
End Sub

Private Sub cmdDefaultFolders_Click()
    Dim i As Integer, j As Integer
    ' List of default contact groups to put new email addresses into.
    Const numDefaults As Integer = 1
    Dim sDefaults(numDefaults - 1) As String
    sDefaults(0) = "Yadda Yadda"
    ' End of list
    With Me.lstFolders
        ' Check each folder/contact group listed in this form.
        For i = 0 To .ListCount - 1
            .Selected(i) = False
            ' If this contact group is a default contact group...
            For j = 0 To numDefaults - 1
                If LCase$(.List(i, 2)) = LCase$(sDefaults(j)) Then
                    .Selected(i) = True ' ... Select it and ...
                    Exit For ' ... Go to the next folder/contact group.
                End If
            Next
        Next
    End With
End Sub

