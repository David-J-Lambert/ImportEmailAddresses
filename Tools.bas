Attribute VB_Name = "Tools"
' Module Tools.
' Author: David Lambert, David5Lambert7@Gmail.com.
' MICROSOFT OFFICE SETUP: The spreadsheet "INSTRUCTIONS" has macro setup instructions.

Option Explicit

' MISCELLANEOUS STAND-ALONE FUNCTIONS AND SUBROUTINES.

Public Function CheckEmailAddressText(ByVal sEmailAddress As String, ByRef rCell As Range, ByRef numBadAddresses As Integer) As Boolean

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
        numBadAddresses = numBadAddresses + 1
    End If

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

Public Sub ShowProgress(ByVal sOperation As String, ByVal iProcessed As Integer, ByVal iTotal As Integer, ByVal sDestination As String)
    frmProgress.lblProgress.Caption = ProgressString(sOperation, iProcessed, iTotal, sDestination)
    DoEvents
End Sub ' ShowProgress

Public Function ProgressString(ByVal sOperation As String, ByVal iProcessed As Integer, ByVal iTotal As Integer, ByVal sDestination As String) As String
    ProgressString = sOperation & " " & Format$(iProcessed, "#,##0") & " of " & Format$(iTotal, "#,##0") & " email addresses " & sDestination
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

End Function ' ValidLocalPart

