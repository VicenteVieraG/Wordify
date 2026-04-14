Attribute VB_Name = "modSetup"
Option Explicit

' Installs all Wordify keybindings in the current customization context.
Public Sub Wordify_InstallKeybindings()
    Dim ctx As Template
    Set ctx = NormalTemplate

    CustomizationContext = ctx

    Wordify_RemoveKeybindings

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, Command:="Wordify_F2_Time", KeyCode:=BuildKeyCode(wdKeyF2)
    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, Command:="Wordify_F3_Spell", KeyCode:=BuildKeyCode(wdKeyF3)
    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, Command:="Wordify_ShiftF3_SpellUpper", KeyCode:=BuildKeyCode(wdKeyShift, wdKeyF3)
    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, Command:="Wordify_F4_LandMeasure", KeyCode:=BuildKeyCode(wdKeyF4)
    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, Command:="Wordify_F5_Number", KeyCode:=BuildKeyCode(wdKeyF5)
    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, Command:="Wordify_F6_CurrencyMXN", KeyCode:=BuildKeyCode(wdKeyF6)
    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, Command:="Wordify_F7_LinearMeters", KeyCode:=BuildKeyCode(wdKeyF7)
    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, Command:="Wordify_F8_SquareMeters", KeyCode:=BuildKeyCode(wdKeyF8)
    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, Command:="Wordify_F9_DeleteToken", KeyCode:=BuildKeyCode(wdKeyF9)
    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, Command:="Wordify_F10_DeleteToVisibleLineEnd", KeyCode:=BuildKeyCode(wdKeyF10)
    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, Command:="Wordify_F11_DeleteVisibleLine", KeyCode:=BuildKeyCode(wdKeyF11)

    MsgBox "Atajos de Wordify instalados correctamente en la plantilla Normal.dotm.", vbInformation, "Wordify"
End Sub

' Removes Wordify keybindings from the current customization context.
Public Sub Wordify_RemoveKeybindings()
    Dim kb As KeyBinding
    Dim names As Object

    Set names = CreateObject("Scripting.Dictionary")
    names.CompareMode = 1
    names("WORDIFY_F2_TIME") = True
    names("WORDIFY_F3_SPELL") = True
    names("WORDIFY_SHIFTF3_SPELLUPPER") = True
    names("WORDIFY_F4_LANDMEASURE") = True
    names("WORDIFY_F5_NUMBER") = True
    names("WORDIFY_F6_CURRENCYMXN") = True
    names("WORDIFY_F7_LINEARMETERS") = True
    names("WORDIFY_F8_SQUAREMETERS") = True
    names("WORDIFY_F9_DELETETOKEN") = True
    names("WORDIFY_F10_DELETETOVISIBLELINEEND") = True
    names("WORDIFY_F11_DELETEVISIBLELINE") = True

    For Each kb In KeyBindings
        If kb.KeyCategory = wdKeyCategoryMacro Then
            If names.Exists(UCase$(kb.Command)) Then
                kb.Clear
            End If
        End If
    Next kb
End Sub
