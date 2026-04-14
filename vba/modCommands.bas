Attribute VB_Name = "modCommands"
Option Explicit

Public Sub Wordify_F2_Time()
    Wordify_RunConversion "F2"
End Sub

Public Sub Wordify_F3_Spell()
    Wordify_RunConversion "F3"
End Sub

Public Sub Wordify_ShiftF3_SpellUpper()
    Wordify_RunConversion "SHIFT_F3"
End Sub

Public Sub Wordify_F4_LandMeasure()
    Wordify_RunConversion "F4"
End Sub

Public Sub Wordify_F5_Number()
    Wordify_RunConversion "F5"
End Sub

Public Sub Wordify_F6_CurrencyMXN()
    Wordify_RunConversion "F6"
End Sub

Public Sub Wordify_F7_LinearMeters()
    Wordify_RunConversion "F7"
End Sub

Public Sub Wordify_F8_SquareMeters()
    Wordify_RunConversion "F8"
End Sub

Public Sub Wordify_F9_DeleteToken()
    Wordify_DeleteSelectionOrAdjacentToken
End Sub

Public Sub Wordify_F10_DeleteToVisibleLineEnd()
    Wordify_DeleteToVisibleLineEnd
End Sub

Public Sub Wordify_F11_DeleteVisibleLine()
    Wordify_DeleteCurrentVisibleLine
End Sub

Private Sub Wordify_RunConversion(ByVal commandId As String)
    On Error GoTo EH

    Dim tokenRng As Range
    Dim tokenText As String
    Dim errMsg As String
    Dim outText As String
    Dim h As Long, m As Long, ha As Long, ar As Long, ca As Long
    Dim intPart As String, decPart As String

    Select Case commandId
        Case "F2"
            If Not Wordify_GetSelectionOrAdjacentToken(tokenRng, tokenText, WordifyInputTime, errMsg) Then GoTo ValidationFailed
            If Not Wordify_ValidateTimeToken(tokenText, h, m, errMsg) Then GoTo ValidationFailed
            outText = Wordify_ConvertTimeToWords(h, m)

        Case "F3"
            If Not Wordify_GetSelectionOrAdjacentToken(tokenRng, tokenText, WordifyInputSpelling, errMsg) Then GoTo ValidationFailed
            If Not Wordify_ValidateSpellingToken(tokenText, errMsg) Then GoTo ValidationFailed
            outText = Wordify_ConvertSpelling(tokenText, False, errMsg)
            If Len(errMsg) > 0 Then GoTo ValidationFailed

        Case "SHIFT_F3"
            If Not Wordify_GetSelectionOrAdjacentToken(tokenRng, tokenText, WordifyInputSpelling, errMsg) Then GoTo ValidationFailed
            If Not Wordify_ValidateSpellingToken(tokenText, errMsg) Then GoTo ValidationFailed
            outText = Wordify_ConvertSpelling(tokenText, True, errMsg)
            If Len(errMsg) > 0 Then GoTo ValidationFailed

        Case "F4"
            If Not Wordify_GetSelectionOrAdjacentToken(tokenRng, tokenText, WordifyInputLand, errMsg) Then GoTo ValidationFailed
            If Not Wordify_ValidateLandToken(tokenText, ha, ar, ca, errMsg) Then GoTo ValidationFailed
            outText = Wordify_ConvertLandToWords(ha, ar, ca)

        Case "F5"
            If Not Wordify_GetSelectionOrAdjacentToken(tokenRng, tokenText, WordifyInputNumeric, errMsg) Then GoTo ValidationFailed
            If Not Wordify_ValidateNumericToken(tokenText, intPart, decPart, errMsg) Then GoTo ValidationFailed
            outText = Wordify_NumberToWords(intPart, decPart, False)

        Case "F6"
            If Not Wordify_GetSelectionOrAdjacentToken(tokenRng, tokenText, WordifyInputNumeric, errMsg) Then GoTo ValidationFailed
            If Not Wordify_ValidateNumericToken(tokenText, intPart, decPart, errMsg) Then GoTo ValidationFailed
            outText = Wordify_ConvertCurrencyMXN(intPart, decPart)

        Case "F7"
            If Not Wordify_GetSelectionOrAdjacentToken(tokenRng, tokenText, WordifyInputNumeric, errMsg) Then GoTo ValidationFailed
            If Not Wordify_ValidateNumericToken(tokenText, intPart, decPart, errMsg) Then GoTo ValidationFailed
            outText = Wordify_NumberToWords(intPart, decPart, True) & " " & Wordify_UnitForMeters(intPart, decPart, False)

        Case "F8"
            If Not Wordify_GetSelectionOrAdjacentToken(tokenRng, tokenText, WordifyInputNumeric, errMsg) Then GoTo ValidationFailed
            If Not Wordify_ValidateNumericToken(tokenText, intPart, decPart, errMsg) Then GoTo ValidationFailed
            outText = Wordify_NumberToWords(intPart, decPart, True) & " " & Wordify_UnitForMeters(intPart, decPart, True)

        Case Else
            Wordify_ShowError "Comando no reconocido."
            Exit Sub
    End Select

    Wordify_AppendConvertedText tokenRng, outText
    Exit Sub

ValidationFailed:
    Wordify_ShowError errMsg
    Exit Sub
EH:
    Wordify_ShowUnexpectedError "Wordify_RunConversion", Err.Number, Err.Description
End Sub

Public Sub Wordify_AppendConvertedText(ByVal originalRange As Range, ByVal converted As String)
    Dim insertPos As Long
    Dim outRng As Range

    insertPos = originalRange.End
    Set outRng = ActiveDocument.Range(insertPos, insertPos)
    outRng.InsertAfter " " & converted
End Sub

Private Function Wordify_UnitForMeters(ByVal intPart As String, ByVal decPart As String, ByVal squared As Boolean) As String
    Dim isOne As Boolean
    isOne = (CLng(Val(intPart)) = 1 And Len(decPart) = 0)

    If squared Then
        If isOne Then
            Wordify_UnitForMeters = "metro cuadrado"
        Else
            Wordify_UnitForMeters = "metros cuadrados"
        End If
    Else
        If isOne Then
            Wordify_UnitForMeters = "metro"
        Else
            Wordify_UnitForMeters = "metros"
        End If
    End If
End Function
