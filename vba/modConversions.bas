Attribute VB_Name = "modConversions"
Option Explicit

Public Function Wordify_ConvertTimeToWords(ByVal hourVal As Long, ByVal minuteVal As Long) As String
    Dim hourWords As String
    Dim minuteWords As String
    Dim hourUnit As String
    Dim minuteUnit As String

    hourWords = Wordify_NumberToWords(CStr(hourVal), vbNullString, True)
    minuteWords = Wordify_NumberToWords(CStr(minuteVal), vbNullString, True)

    If hourVal = 1 Then
        hourUnit = "hora"
    Else
        hourUnit = "horas"
    End If

    If minuteVal = 1 Then
        minuteUnit = "minuto"
    Else
        minuteUnit = "minutos"
    End If

    Wordify_ConvertTimeToWords = hourWords & " " & hourUnit & " " & minuteWords & " " & minuteUnit
End Function

Public Function Wordify_ConvertLandToWords(ByVal h As Long, ByVal a As Long, ByVal c As Long) As String
    Wordify_ConvertLandToWords = Wordify_NumberForFeminineNoun(h) & " " & Wordify_Pluralize("hectárea", "hectáreas", h) & " " & _
                                 Wordify_NumberForFeminineNoun(a) & " " & Wordify_Pluralize("área", "áreas", a) & " " & _
                                 Wordify_NumberForFeminineNoun(c) & " " & Wordify_Pluralize("centiárea", "centiáreas", c)
End Function

Public Function Wordify_ConvertCurrencyMXN(ByVal intPart As String, ByVal decPart As String) As String
    Dim pesos As Long
    Dim cents As Long
    Dim pesoWords As String
    Dim centWords As String

    pesos = CLng(Val(intPart))
    cents = Wordify_GetRoundedCents(decPart)

    If cents = 100 Then
        pesos = pesos + 1
        cents = 0
    End If

    pesoWords = Wordify_NumberToWords(CStr(pesos), vbNullString, True)
    centWords = Wordify_NumberToWords(CStr(cents), vbNullString, True)

    Wordify_ConvertCurrencyMXN = pesoWords & " " & Wordify_Pluralize("peso", "pesos", pesos)

    If cents > 0 Then
        Wordify_ConvertCurrencyMXN = Wordify_ConvertCurrencyMXN & " " & centWords & " " & Wordify_Pluralize("centavo", "centavos", cents)
    End If

    Wordify_ConvertCurrencyMXN = Wordify_ConvertCurrencyMXN & " moneda nacional"
End Function

Public Function Wordify_ConvertSpelling(ByVal token As String, Optional ByVal upperOutput As Boolean = False, Optional ByRef errMsg As String) As String
    Dim i As Long
    Dim ch As String
    Dim mapped As String
    Dim items As Collection

    Set items = New Collection

    For i = 1 To Len(token)
        ch = Mid$(token, i, 1)
        mapped = Wordify_MapSpellingSymbol(ch)

        If Len(mapped) = 0 Then
            errMsg = "Símbolo no soportado para deletreo: '" & ch & "'."
            Exit Function
        End If

        items.Add mapped
    Next i

    Wordify_ConvertSpelling = Wordify_JoinWithComma(items)

    If upperOutput Then
        Wordify_ConvertSpelling = UCase$(Wordify_ConvertSpelling)
    End If
End Function

Private Function Wordify_MapSpellingSymbol(ByVal ch As String) As String
    Dim u As String
    u = UCase$(ch)

    Select Case u
        Case "A": Wordify_MapSpellingSymbol = "letra ""A"""
        Case "B": Wordify_MapSpellingSymbol = "letra ""BE"""
        Case "C": Wordify_MapSpellingSymbol = "letra ""CE"""
        Case "D": Wordify_MapSpellingSymbol = "letra ""DE"""
        Case "E": Wordify_MapSpellingSymbol = "letra ""E"""
        Case "F": Wordify_MapSpellingSymbol = "letra ""EFE"""
        Case "G": Wordify_MapSpellingSymbol = "letra ""GE"""
        Case "H": Wordify_MapSpellingSymbol = "letra ""HACHE"""
        Case "I": Wordify_MapSpellingSymbol = "letra ""I"""
        Case "J": Wordify_MapSpellingSymbol = "letra ""JOTA"""
        Case "K": Wordify_MapSpellingSymbol = "letra ""KA"""
        Case "L": Wordify_MapSpellingSymbol = "letra ""ELE"""
        Case "M": Wordify_MapSpellingSymbol = "letra ""EME"""
        Case "N": Wordify_MapSpellingSymbol = "letra ""ENE"""
        Case "Ñ": Wordify_MapSpellingSymbol = "letra ""EÑE"""
        Case "O": Wordify_MapSpellingSymbol = "letra ""O"""
        Case "P": Wordify_MapSpellingSymbol = "letra ""PE"""
        Case "Q": Wordify_MapSpellingSymbol = "letra ""CU"""
        Case "R": Wordify_MapSpellingSymbol = "letra ""ERRE"""
        Case "S": Wordify_MapSpellingSymbol = "letra ""ESE"""
        Case "T": Wordify_MapSpellingSymbol = "letra ""TE"""
        Case "U": Wordify_MapSpellingSymbol = "letra ""U"""
        Case "V": Wordify_MapSpellingSymbol = "letra ""UVE"""
        Case "W": Wordify_MapSpellingSymbol = "letra ""DOBLE U"""
        Case "X": Wordify_MapSpellingSymbol = "letra ""EQUIS"""
        Case "Y": Wordify_MapSpellingSymbol = "letra ""I GRIEGA"""
        Case "Z": Wordify_MapSpellingSymbol = "letra ""ZETA"""
        Case "Á": Wordify_MapSpellingSymbol = "letra ""A"""
        Case "É": Wordify_MapSpellingSymbol = "letra ""E"""
        Case "Í": Wordify_MapSpellingSymbol = "letra ""I"""
        Case "Ó": Wordify_MapSpellingSymbol = "letra ""O"""
        Case "Ú", "Ü": Wordify_MapSpellingSymbol = "letra ""U"""
        Case "0" To "9": Wordify_MapSpellingSymbol = Wordify_DigitName(ch)
        Case ".": Wordify_MapSpellingSymbol = "punto"
        Case ",": Wordify_MapSpellingSymbol = "coma"
        Case ":": Wordify_MapSpellingSymbol = "dos puntos"
        Case ";": Wordify_MapSpellingSymbol = "punto y coma"
        Case "-": Wordify_MapSpellingSymbol = "guion"
        Case "_": Wordify_MapSpellingSymbol = "guion bajo"
        Case "/": Wordify_MapSpellingSymbol = "barra"
        Case "\\": Wordify_MapSpellingSymbol = "barra invertida"
        Case "(": Wordify_MapSpellingSymbol = "paréntesis abre"
        Case ")": Wordify_MapSpellingSymbol = "paréntesis cierra"
        Case "[": Wordify_MapSpellingSymbol = "corchete abre"
        Case "]": Wordify_MapSpellingSymbol = "corchete cierra"
        Case "{": Wordify_MapSpellingSymbol = "llave abre"
        Case "}": Wordify_MapSpellingSymbol = "llave cierra"
        Case "'": Wordify_MapSpellingSymbol = "comilla simple"
        Case """": Wordify_MapSpellingSymbol = "comillas"
        Case "!": Wordify_MapSpellingSymbol = "signo de exclamación"
        Case "¡": Wordify_MapSpellingSymbol = "signo de exclamación de apertura"
        Case "?": Wordify_MapSpellingSymbol = "signo de interrogación"
        Case "¿": Wordify_MapSpellingSymbol = "signo de interrogación de apertura"
        Case "#": Wordify_MapSpellingSymbol = "almohadilla"
        Case "$": Wordify_MapSpellingSymbol = "signo de peso"
        Case "%": Wordify_MapSpellingSymbol = "por ciento"
        Case "&": Wordify_MapSpellingSymbol = "ampersand"
        Case "=": Wordify_MapSpellingSymbol = "igual"
        Case "+": Wordify_MapSpellingSymbol = "más"
        Case "*": Wordify_MapSpellingSymbol = "asterisco"
        Case "@": Wordify_MapSpellingSymbol = "arroba"
        Case Else
            Wordify_MapSpellingSymbol = vbNullString
    End Select
End Function

Private Function Wordify_JoinWithComma(ByVal items As Collection) As String
    Dim i As Long
    For i = 1 To items.Count
        If i > 1 Then Wordify_JoinWithComma = Wordify_JoinWithComma & ", "
        Wordify_JoinWithComma = Wordify_JoinWithComma & items(i)
    Next i
End Function

Private Function Wordify_GetRoundedCents(ByVal decPart As String) As Long
    Dim padded As String
    Dim firstTwo As Long

    If Len(decPart) = 0 Then
        Wordify_GetRoundedCents = 0
        Exit Function
    End If

    padded = decPart & String$(8, "0")
    firstTwo = CLng(Left$(padded, 2))

    If Mid$(padded, 3, 1) >= "5" Then
        Wordify_GetRoundedCents = firstTwo + 1
    Else
        Wordify_GetRoundedCents = firstTwo
    End If
End Function

Private Function Wordify_NumberForFeminineNoun(ByVal n As Long) As String
    Dim w As String
    w = Wordify_NumberToWords(CStr(n), vbNullString, True)

    If Right$(w, 6) = "un" Then
        w = Left$(w, Len(w) - 2) & "una"
    ElseIf Right$(w, 8) = "veintiún" Then
        w = Left$(w, Len(w) - 8) & "veintiuna"
    Else
        w = Replace(w, " y un", " y una")
    End If

    Wordify_NumberForFeminineNoun = w
End Function

Private Function Wordify_Pluralize(ByVal singular As String, ByVal plural As String, ByVal n As Long) As String
    If n = 1 Then
        Wordify_Pluralize = singular
    Else
        Wordify_Pluralize = plural
    End If
End Function
