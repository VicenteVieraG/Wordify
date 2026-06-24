Attribute VB_Name = "modUI"
Option Explicit

Public Sub Wordify_ShowError(ByVal reasonText As String)
    MsgBox reasonText, vbExclamation, "Wordify"
End Sub

Public Sub Wordify_ShowUnexpectedError(ByVal procName As String, ByVal errNumber As Long, ByVal errDesc As String)
    MsgBox "Ocurrió un error inesperado en " & procName & "." & vbCrLf & _
           "Código: " & CStr(errNumber) & vbCrLf & _
           "Detalle: " & errDesc, _
           vbCritical, "Wordify"
End Sub
