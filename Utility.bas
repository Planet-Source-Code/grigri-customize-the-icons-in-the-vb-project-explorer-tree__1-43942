Attribute VB_Name = "Utility"
'---------------------------------------------------------------------------------------
' Module    : Utility
' Author    : JP Mortiboys
' Purpose   : Some utility functions
'---------------------------------------------------------------------------------------
Option Explicit

' Adds the string to the combo if it doesn't already exist
Public Sub AddOnceToCombo(cboBox As ComboBox, sText As String)
    Dim lRet As Long
    If Trim$(sText) = "" Then Exit Sub
    lRet = SendMessage(cboBox.hwnd, CB_FINDSTRINGEXACT, -1, ByVal sText)
    If lRet <> -1 Then
        cboBox.RemoveItem lRet
    End If
    cboBox.AddItem sText, 0
End Sub

' Load the combo strings from the registry
Public Sub LoadComboStrings(cboBox As ComboBox)
    Dim I As Long, s As String
    I = 0
    Do
        s = GetSetting(App.Title, "MRU_" & cboBox.Name, "Item" & CStr(I), "")
        If s = "" Or I >= 10 Then Exit Sub
        cboBox.AddItem s
        I = I + 1
    Loop
End Sub

' Save the combo strings in the registry
Public Sub SaveComboStrings(cboBox As ComboBox)
    Dim I As Long
    On Error Resume Next
    DeleteSetting App.Title, "MRU_" & cboBox.Name
    For I = 0 To cboBox.ListCount - 1
        If I >= 10 Then Exit Sub
        SaveSetting App.Title, "MRU_" & cboBox.Name, "Item" & CStr(I), cboBox.List(I)
    Next
End Sub

' For display purposes: when it means "" it says "(default)"
Public Function BlankToDefault(ByVal sText As String) As String
    If sText = "" Then BlankToDefault = "(default)" Else BlankToDefault = sText
End Function

Public Function DefaultToBlank(ByVal sText As String) As String
    If sText = "(default)" Then DefaultToBlank = "" Else DefaultToBlank = sText
End Function

' A classic handy function...
Public Function IsCompiled() As Boolean
    IsCompiled = True
    On Error GoTo NOT_COMPILED
    Debug.Print 1 / 0
    Exit Function
NOT_COMPILED:
    IsCompiled = False
End Function
