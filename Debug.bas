Attribute VB_Name = "Debug"
'---------------------------------------------------------------------------------------
' Module    : Debug
' Author    : JP Mortiboys
' Purpose   : Handles debug output in a compiled app
'---------------------------------------------------------------------------------------
#If 0 Then

Option Explicit

Private fDebugWin As DebugWindow

Public Sub DebugOut(ByVal sText As String)
    If fDebugWin Is Nothing Then
        Set fDebugWin = New DebugWindow
        fDebugWin.Visible = True
    End If
    fDebugWin.OutputDebugString sText
End Sub

#End If
