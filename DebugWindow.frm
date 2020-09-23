VERSION 5.00
Begin VB.Form DebugWindow 
   Caption         =   "Debug output"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnClear 
      Caption         =   "Clear"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox txtDebugOutput 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   4455
   End
End
Attribute VB_Name = "DebugWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : DebugWindow
' Author    : JP Mortiboys
' Purpose   : Displays debug output in a compiled app
'---------------------------------------------------------------------------------------
#If 0 Then

Option Explicit

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1

Private Sub btnClear_Click()
    txtDebugOutput.Text = ""
End Sub

Public Sub OutputDebugString(ByVal sText As String)
    txtDebugOutput.SelText = sText & vbCrLf
End Sub

Private Sub Form_Load()
    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub Form_Resize()
    txtDebugOutput.Move 0, btnClear.Top + btnClear.Height, ScaleWidth, ScaleHeight - btnClear.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

#End If
