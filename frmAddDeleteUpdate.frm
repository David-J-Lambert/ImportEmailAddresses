VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAddDeleteUpdate 
   ClientHeight    =   3432
   ClientLeft      =   39
   ClientTop       =   325
   ClientWidth     =   6097
   OleObjectBlob   =   "frmAddDeleteUpdate.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAddDeleteUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    MsgBox "At your request, quitting.", , sThisProgram
    Me.Hide
    End
End Sub

Private Sub cmdOK_Click()
    Me.Hide
End Sub

