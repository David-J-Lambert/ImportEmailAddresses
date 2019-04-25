VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProgress 
   Caption         =   "Running..."
   ClientHeight    =   1105
   ClientLeft      =   39
   ClientTop       =   325
   ClientWidth     =   7722
   OleObjectBlob   =   "frmProgress.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Disables the red X Close button in the top right corner of this form (progress form).
    If CloseMode = 0 Then Cancel = True
End Sub
