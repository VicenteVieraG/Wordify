Attribute VB_Name = "modDeletion"
Option Explicit

Public Sub Wordify_DeleteSelectionOrAdjacentToken()
    On Error GoTo EH

    If Selection.Range.Start <> Selection.Range.End Then
        Selection.Range.Delete
        Exit Sub
    End If

    Dim tokenRng As Range
    Dim tokenText As String
    Dim errMsg As String

    If Not Wordify_GetSelectionOrAdjacentToken(tokenRng, tokenText, WordifyInputDeleteToken, errMsg) Then
        Wordify_ShowError errMsg
        Exit Sub
    End If

    tokenRng.Delete
    Exit Sub
EH:
    Wordify_ShowUnexpectedError "Wordify_DeleteSelectionOrAdjacentToken", Err.Number, Err.Description
End Sub

Public Sub Wordify_DeleteToVisibleLineEnd()
    On Error GoTo EH
    Dim lineRng As Range
    Dim delRng As Range
    Dim startPos As Long
    Dim errMsg As String

    If Not Wordify_GetVisibleLineRange(lineRng, errMsg) Then
        Wordify_ShowError errMsg
        Exit Sub
    End If

    startPos = Selection.Range.Start
    If startPos < lineRng.Start Then startPos = lineRng.Start
    If startPos > lineRng.End Then startPos = lineRng.End

    Set delRng = lineRng.Duplicate
    delRng.SetRange startPos, lineRng.End
    If delRng.Start < delRng.End Then delRng.Delete
    Exit Sub
EH:
    Wordify_ShowUnexpectedError "Wordify_DeleteToVisibleLineEnd", Err.Number, Err.Description
End Sub

Public Sub Wordify_DeleteCurrentVisibleLine()
    On Error GoTo EH
    Dim lineRng As Range
    Dim errMsg As String

    If Not Wordify_GetVisibleLineRange(lineRng, errMsg) Then
        Wordify_ShowError errMsg
        Exit Sub
    End If

    lineRng.Delete
    Exit Sub
EH:
    Wordify_ShowUnexpectedError "Wordify_DeleteCurrentVisibleLine", Err.Number, Err.Description
End Sub
