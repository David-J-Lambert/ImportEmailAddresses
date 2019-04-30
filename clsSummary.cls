VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Module clsSummary.
' Author: David Lambert, David5Lambert7@Gmail.com.
' MICROSOFT OFFICE SETUP: The spreadsheet "INSTRUCTIONS" has macro setup instructions.

Option Explicit

' Summary of all operations performed, printed at end.
Private sSummary As String

Private Sub Class_Initialize()
    sSummary = vbNullString
End Sub

Private Sub Class_Terminate()
    sSummary = vbNullString
End Sub

Public Sub AddToSummary(ByVal sAdd As String)
    sSummary = sSummary & vbCrLf & vbCrLf & sAdd
End Sub

Public Sub PrintSummary(ByVal sCaption As String)
    MsgBox sSummary, , sCaption
End Sub

' Not Used
'Public Function GetSummary() As String
'    GetSummary = sSummary
'End Function
