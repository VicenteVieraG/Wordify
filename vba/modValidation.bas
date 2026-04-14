Attribute VB_Name = "modValidation"
Option Explicit

Public Function Wordify_ValidateTimeToken(ByVal txt As String, ByRef hourVal As Long, ByRef minuteVal As Long, ByRef errMsg As String) As Boolean
    Dim rx As Object
    Dim parts() As String

    txt = Trim$(txt)
    Set rx = CreateObject("VBScript.RegExp")
    rx.Pattern = "^(\d{1,2}):(\d{2})$"
    rx.Global = False

    If Not rx.Test(txt) Then
        errMsg = "Formato de hora inválido. Use HH:MM o H:MM."
        Exit Function
    End If

    parts = Split(txt, ":")
    hourVal = CLng(parts(0))
    minuteVal = CLng(parts(1))

    If hourVal < 0 Or hourVal > 23 Then
        errMsg = "La hora debe estar entre 0 y 23."
        Exit Function
    End If

    If minuteVal < 0 Or minuteVal > 59 Then
        errMsg = "Los minutos deben estar entre 0 y 59."
        Exit Function
    End If

    Wordify_ValidateTimeToken = True
End Function

Public Function Wordify_ValidateLandToken(ByVal txt As String, ByRef h As Long, ByRef a As Long, ByRef c As Long, ByRef errMsg As String) As Boolean
    Dim rx As Object
    Dim parts() As String

    txt = Trim$(txt)
    Set rx = CreateObject("VBScript.RegExp")
    rx.Pattern = "^\d{2}-\d{2}-\d{2}$"

    If Not rx.Test(txt) Then
        errMsg = "Formato inválido. Use exactamente NN-NN-NN."
        Exit Function
    End If

    parts = Split(txt, "-")
    h = CLng(parts(0))
    a = CLng(parts(1))
    c = CLng(parts(2))

    Wordify_ValidateLandToken = True
End Function

Public Function Wordify_ValidateNumericToken(ByVal txt As String, ByRef intPart As String, ByRef decPart As String, ByRef errMsg As String) As Boolean
    Dim rx As Object

    txt = Trim$(txt)

    If Len(txt) = 0 Then
        errMsg = "No hay valor numérico para convertir."
        Exit Function
    End If

    If Left$(txt, 1) = "-" Then
        errMsg = "No se permiten números negativos."
        Exit Function
    End If

    If InStr(txt, ",") > 0 Then
        errMsg = "No se permiten separadores de miles ni coma decimal. Use punto."
        Exit Function
    End If

    Set rx = CreateObject("VBScript.RegExp")
    rx.Pattern = "^\d+(\.\d{1,8})?$"
    rx.Global = False

    If Not rx.Test(txt) Then
        errMsg = "Formato numérico inválido. Use entero o decimal con punto y máximo 8 decimales."
        Exit Function
    End If

    If Right$(txt, 1) = "." Then
        errMsg = "Formato numérico inválido: no se permite terminar en punto."
        Exit Function
    End If

    If InStr(txt, ".") > 0 Then
        intPart = Split(txt, ".")(0)
        decPart = Split(txt, ".")(1)
    Else
        intPart = txt
        decPart = vbNullString
    End If

    Wordify_ValidateNumericToken = True
End Function

Public Function Wordify_ValidateSpellingToken(ByVal txt As String, ByRef errMsg As String) As Boolean
    txt = Trim$(txt)

    If Len(txt) = 0 Then
        errMsg = "No hay texto para deletrear."
        Exit Function
    End If

    If InStr(1, txt, " ", vbBinaryCompare) > 0 Or InStr(1, txt, vbTab, vbBinaryCompare) > 0 Or InStr(1, txt, vbCr, vbBinaryCompare) > 0 Then
        errMsg = "Para deletrear, el texto no debe contener espacios ni saltos de línea."
        Exit Function
    End If

    Wordify_ValidateSpellingToken = True
End Function
