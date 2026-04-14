Attribute VB_Name = "modSpanishNumbers"
Option Explicit

Public Function Wordify_NumberToWords(ByVal intPart As String, Optional ByVal decPart As String = vbNullString, Optional ByVal apocopateInteger As Boolean = False) As String
    Dim intWords As String

    intWords = Wordify_IntegerToWords(intPart)
    If apocopateInteger Then intWords = Wordify_Apocopate(intWords)

    If Len(decPart) = 0 Then
        Wordify_NumberToWords = intWords
    Else
        Wordify_NumberToWords = intWords & " punto " & Wordify_DecimalDigitsToWords(decPart)
    End If
End Function

Public Function Wordify_IntegerToWords(ByVal numText As String) As String
    Dim cleanNum As String
    Dim groups As Collection
    Dim i As Long
    Dim groupValue As Long
    Dim parts As Collection
    Dim grpText As String
    Dim idxFromRight As Long

    cleanNum = Wordify_StripLeadingZeros(numText)
    If cleanNum = "" Then cleanNum = "0"

    If cleanNum = "0" Then
        Wordify_IntegerToWords = "cero"
        Exit Function
    End If

    Set groups = Wordify_SplitGroups(cleanNum)
    Set parts = New Collection

    For i = 1 To groups.Count
        idxFromRight = groups.Count - i
        groupValue = CLng(groups(i))
        If groupValue <> 0 Then
            grpText = Wordify_ThreeDigitsToWords(groupValue)
            Select Case idxFromRight
                Case 0
                    parts.Add grpText
                Case 1
                    If groupValue = 1 Then
                        parts.Add "mil"
                    Else
                        parts.Add Wordify_Apocopate(grpText) & " mil"
                    End If
                Case 2
                    If groupValue = 1 Then
                        parts.Add "un millón"
                    Else
                        parts.Add Wordify_Apocopate(grpText) & " millones"
                    End If
                Case Else
                    parts.Add Wordify_Apocopate(grpText) & " *10^" & CStr(idxFromRight * 3)
            End Select
        End If
    Next i

    Wordify_IntegerToWords = Wordify_JoinCollection(parts, " ")
End Function

Private Function Wordify_ThreeDigitsToWords(ByVal n As Long) As String
    Dim hundreds As Long
    Dim tensUnits As Long

    hundreds = n \ 100
    tensUnits = n Mod 100

    Select Case hundreds
        Case 0
            Wordify_ThreeDigitsToWords = Wordify_TwoDigitsToWords(tensUnits)
        Case 1
            If tensUnits = 0 Then
                Wordify_ThreeDigitsToWords = "cien"
            Else
                Wordify_ThreeDigitsToWords = "ciento " & Wordify_TwoDigitsToWords(tensUnits)
            End If
        Case 2
            Wordify_ThreeDigitsToWords = "doscientos" & Wordify_SpaceJoin(Wordify_TwoDigitsToWords(tensUnits))
        Case 3
            Wordify_ThreeDigitsToWords = "trescientos" & Wordify_SpaceJoin(Wordify_TwoDigitsToWords(tensUnits))
        Case 4
            Wordify_ThreeDigitsToWords = "cuatrocientos" & Wordify_SpaceJoin(Wordify_TwoDigitsToWords(tensUnits))
        Case 5
            Wordify_ThreeDigitsToWords = "quinientos" & Wordify_SpaceJoin(Wordify_TwoDigitsToWords(tensUnits))
        Case 6
            Wordify_ThreeDigitsToWords = "seiscientos" & Wordify_SpaceJoin(Wordify_TwoDigitsToWords(tensUnits))
        Case 7
            Wordify_ThreeDigitsToWords = "setecientos" & Wordify_SpaceJoin(Wordify_TwoDigitsToWords(tensUnits))
        Case 8
            Wordify_ThreeDigitsToWords = "ochocientos" & Wordify_SpaceJoin(Wordify_TwoDigitsToWords(tensUnits))
        Case 9
            Wordify_ThreeDigitsToWords = "novecientos" & Wordify_SpaceJoin(Wordify_TwoDigitsToWords(tensUnits))
    End Select
End Function

Public Function Wordify_TwoDigitsToWords(ByVal n As Long) As String
    Select Case n
        Case 0: Wordify_TwoDigitsToWords = ""
        Case 1: Wordify_TwoDigitsToWords = "uno"
        Case 2: Wordify_TwoDigitsToWords = "dos"
        Case 3: Wordify_TwoDigitsToWords = "tres"
        Case 4: Wordify_TwoDigitsToWords = "cuatro"
        Case 5: Wordify_TwoDigitsToWords = "cinco"
        Case 6: Wordify_TwoDigitsToWords = "seis"
        Case 7: Wordify_TwoDigitsToWords = "siete"
        Case 8: Wordify_TwoDigitsToWords = "ocho"
        Case 9: Wordify_TwoDigitsToWords = "nueve"
        Case 10: Wordify_TwoDigitsToWords = "diez"
        Case 11: Wordify_TwoDigitsToWords = "once"
        Case 12: Wordify_TwoDigitsToWords = "doce"
        Case 13: Wordify_TwoDigitsToWords = "trece"
        Case 14: Wordify_TwoDigitsToWords = "catorce"
        Case 15: Wordify_TwoDigitsToWords = "quince"
        Case 16: Wordify_TwoDigitsToWords = "dieciséis"
        Case 17: Wordify_TwoDigitsToWords = "diecisiete"
        Case 18: Wordify_TwoDigitsToWords = "dieciocho"
        Case 19: Wordify_TwoDigitsToWords = "diecinueve"
        Case 20: Wordify_TwoDigitsToWords = "veinte"
        Case 21: Wordify_TwoDigitsToWords = "veintiuno"
        Case 22: Wordify_TwoDigitsToWords = "veintidós"
        Case 23: Wordify_TwoDigitsToWords = "veintitrés"
        Case 24: Wordify_TwoDigitsToWords = "veinticuatro"
        Case 25: Wordify_TwoDigitsToWords = "veinticinco"
        Case 26: Wordify_TwoDigitsToWords = "veintiséis"
        Case 27: Wordify_TwoDigitsToWords = "veintisiete"
        Case 28: Wordify_TwoDigitsToWords = "veintiocho"
        Case 29: Wordify_TwoDigitsToWords = "veintinueve"
        Case 30: Wordify_TwoDigitsToWords = "treinta"
        Case 40: Wordify_TwoDigitsToWords = "cuarenta"
        Case 50: Wordify_TwoDigitsToWords = "cincuenta"
        Case 60: Wordify_TwoDigitsToWords = "sesenta"
        Case 70: Wordify_TwoDigitsToWords = "setenta"
        Case 80: Wordify_TwoDigitsToWords = "ochenta"
        Case 90: Wordify_TwoDigitsToWords = "noventa"
        Case Else
            Wordify_TwoDigitsToWords = Wordify_TwoDigitsToWords((n \ 10) * 10) & " y " & Wordify_TwoDigitsToWords(n Mod 10)
    End Select
End Function

Public Function Wordify_DecimalDigitsToWords(ByVal decPart As String) As String
    Dim i As Long
    Dim items As Collection
    Set items = New Collection

    For i = 1 To Len(decPart)
        items.Add Wordify_DigitName(Mid$(decPart, i, 1), False)
    Next i

    Wordify_DecimalDigitsToWords = Wordify_JoinCollection(items, " ")
End Function

Public Function Wordify_DigitName(ByVal ch As String, Optional ByVal forSpelling As Boolean = True) As String
    Select Case ch
        Case "0": Wordify_DigitName = "cero"
        Case "1": Wordify_DigitName = "uno"
        Case "2": Wordify_DigitName = "dos"
        Case "3": Wordify_DigitName = "tres"
        Case "4": Wordify_DigitName = "cuatro"
        Case "5": Wordify_DigitName = "cinco"
        Case "6": Wordify_DigitName = "seis"
        Case "7": Wordify_DigitName = "siete"
        Case "8": Wordify_DigitName = "ocho"
        Case "9": Wordify_DigitName = "nueve"
    End Select

    If forSpelling Then Wordify_DigitName = LCase$(Wordify_DigitName)
End Function

Public Function Wordify_Apocopate(ByVal txt As String) As String
    Dim w As String
    w = txt
    w = Replace(w, " veintiuno", " veintiún")
    If Right$(w, 8) = "veintiuno" Then w = Left$(w, Len(w) - 8) & "veintiún"

    w = Replace(w, " y uno", " y un")
    If Right$(w, 3) = "uno" Then w = Left$(w, Len(w) - 3) & "un"

    Wordify_Apocopate = w
End Function

Private Function Wordify_StripLeadingZeros(ByVal txt As String) As String
    Dim i As Long
    For i = 1 To Len(txt)
        If Mid$(txt, i, 1) <> "0" Then
            Wordify_StripLeadingZeros = Mid$(txt, i)
            Exit Function
        End If
    Next i
    Wordify_StripLeadingZeros = ""
End Function

Private Function Wordify_SplitGroups(ByVal txt As String) As Collection
    Dim c As New Collection
    Dim n As Long
    Dim startPos As Long
    Dim lengthGroup As Long

    n = Len(txt)
    startPos = 1
    lengthGroup = n Mod 3
    If lengthGroup = 0 Then lengthGroup = 3

    Do While startPos <= n
        c.Add Mid$(txt, startPos, lengthGroup)
        startPos = startPos + lengthGroup
        lengthGroup = 3
    Loop

    Set Wordify_SplitGroups = c
End Function

Private Function Wordify_JoinCollection(ByVal c As Collection, ByVal sep As String) As String
    Dim i As Long
    For i = 1 To c.Count
        If i > 1 Then Wordify_JoinCollection = Wordify_JoinCollection & sep
        Wordify_JoinCollection = Wordify_JoinCollection & c(i)
    Next i
End Function

Private Function Wordify_SpaceJoin(ByVal txt As String) As String
    If Len(txt) > 0 Then Wordify_SpaceJoin = " " & txt
End Function
