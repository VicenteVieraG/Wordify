Attribute VB_Name = "modTokenDetection"
Option Explicit

Public Enum WordifyInputMode
    WordifyInputNumeric = 1
    WordifyInputSpelling = 2
    WordifyInputTime = 3
    WordifyInputLand = 4
    WordifyInputDeleteToken = 5
End Enum

' Gets token from selection or adjacency rules.
Public Function Wordify_GetSelectionOrAdjacentToken(ByRef tokenRange As Range, ByRef tokenText As String, ByVal mode As WordifyInputMode, ByRef errMsg As String) As Boolean
    On Error GoTo EH

    If Selection Is Nothing Then
        errMsg = "No hay una selección activa."
        Exit Function
    End If

    If Selection.Range.Start <> Selection.Range.End Then
        Set tokenRange = Selection.Range.Duplicate
        tokenText = tokenRange.Text

        If LenB(Trim$(tokenText)) = 0 Then
            errMsg = "La selección está vacía o contiene solo espacios."
            Exit Function
        End If

        If mode = WordifyInputSpelling Then
            If InStr(1, tokenText, " ", vbBinaryCompare) > 0 Or InStr(1, tokenText, vbTab, vbBinaryCompare) > 0 Or InStr(1, tokenText, vbCr, vbBinaryCompare) > 0 Then
                errMsg = "Para deletrear, la selección no debe contener espacios ni saltos de línea."
                Exit Function
            End If
        End If

        Wordify_GetSelectionOrAdjacentToken = True
        Exit Function
    End If

    Wordify_GetAdjacentTokenAtCursor tokenRange, tokenText, mode, errMsg
    Wordify_GetSelectionOrAdjacentToken = (Len(errMsg) = 0)
    Exit Function
EH:
    errMsg = "No se pudo obtener el texto objetivo."
End Function

' Finds the contiguous non-whitespace token around the insertion point.
Public Sub Wordify_GetAdjacentTokenAtCursor(ByRef tokenRange As Range, ByRef tokenText As String, ByVal mode As WordifyInputMode, ByRef errMsg As String)
    Dim pos As Long
    Dim docRng As Range
    Dim startPos As Long
    Dim endPos As Long
    Dim leftChar As String
    Dim rightChar As String

    errMsg = vbNullString
    tokenText = vbNullString

    pos = Selection.Range.Start
    Set docRng = ActiveDocument.Range(0, ActiveDocument.Content.End)

    leftChar = Wordify_GetCharAt(pos - 1)
    rightChar = Wordify_GetCharAt(pos)

    If Not Wordify_IsTokenChar(leftChar, mode) And Not Wordify_IsTokenChar(rightChar, mode) Then
        errMsg = "No hay un token adyacente al cursor."
        Exit Sub
    End If

    startPos = pos
    Do While startPos > 0
        If Not Wordify_IsTokenChar(Wordify_GetCharAt(startPos - 1), mode) Then Exit Do
        startPos = startPos - 1
    Loop

    endPos = pos
    Do While endPos < docRng.End
        If Not Wordify_IsTokenChar(Wordify_GetCharAt(endPos), mode) Then Exit Do
        endPos = endPos + 1
    Loop

    If endPos <= startPos Then
        errMsg = "No hay un token adyacente válido."
        Exit Sub
    End If

    Set tokenRange = ActiveDocument.Range(startPos, endPos)
    tokenText = tokenRange.Text

    If LenB(Trim$(tokenText)) = 0 Then
        errMsg = "No hay un token válido junto al cursor."
    End If
End Sub

Private Function Wordify_GetCharAt(ByVal pos As Long) As String
    If pos < 0 Or pos >= ActiveDocument.Content.End Then
        Wordify_GetCharAt = vbNullString
    Else
        Wordify_GetCharAt = ActiveDocument.Range(pos, pos + 1).Text
    End If
End Function

Private Function Wordify_IsTokenChar(ByVal ch As String, ByVal mode As WordifyInputMode) As Boolean
    If Len(ch) = 0 Then Exit Function
    If ch = " " Or ch = vbTab Or ch = vbCr Or ch = ChrW(11) Then Exit Function

    Select Case mode
        Case WordifyInputTime
            Wordify_IsTokenChar = (InStr("0123456789:", ch) > 0)
        Case WordifyInputLand
            Wordify_IsTokenChar = (InStr("0123456789-", ch) > 0)
        Case WordifyInputNumeric
            Wordify_IsTokenChar = (InStr("0123456789.", ch) > 0)
        Case WordifyInputSpelling, WordifyInputDeleteToken
            Wordify_IsTokenChar = True
        Case Else
            Wordify_IsTokenChar = True
    End Select
End Function

' Visible line approximation in Word layout.
Public Function Wordify_GetVisibleLineRange(ByRef lineRange As Range, ByRef errMsg As String) As Boolean
    On Error GoTo EH
    Dim r As Range

    Set r = Selection.Range.Duplicate
    r.Collapse wdCollapseStart

    r.MoveStart wdLine, 0
    r.MoveEnd wdLine, 1

    If r.End > ActiveDocument.Content.End Then
        r.End = ActiveDocument.Content.End
    End If

    Set lineRange = r
    Wordify_GetVisibleLineRange = True
    Exit Function
EH:
    errMsg = "No se pudo identificar la línea visible actual."
End Function
