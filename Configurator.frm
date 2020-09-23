VERSION 5.00
Begin VB.Form ConfigPanel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VBIcons: IDE - Configuration"
   ClientHeight    =   3060
   ClientLeft      =   2175
   ClientTop       =   1935
   ClientWidth     =   6435
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   204
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   429
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton opFrame 
      Caption         =   "Component Icons"
      Height          =   375
      Index           =   1
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   480
      Width           =   1455
   End
   Begin VB.OptionButton opFrame 
      Caption         =   "Default Icons"
      Height          =   375
      Index           =   0
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   0
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.PictureBox picFrame 
      HasDC           =   0   'False
      Height          =   3015
      Index           =   0
      Left            =   1560
      ScaleHeight     =   197
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   317
      TabIndex        =   1
      Top             =   0
      Width           =   4815
      Begin VB.PictureBox picDefIconsFrame 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   360
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   269
         TabIndex        =   13
         Top             =   1080
         Width           =   4095
         Begin VB.PictureBox picDefIcons 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   300
            Left            =   0
            ScaleHeight     =   20
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   460
            TabIndex        =   5
            Top             =   0
            Width           =   6900
            Begin VB.Shape shpSelect 
               BorderColor     =   &H000000FF&
               Height          =   300
               Left            =   0
               Top             =   0
               Width           =   300
            End
         End
      End
      Begin VB.CommandButton btnScrollLeft 
         Caption         =   "ï"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   255
      End
      Begin VB.CommandButton btnScrollRight 
         Caption         =   "ð"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   10
         Top             =   1080
         Width           =   255
      End
      Begin VB.ComboBox cboDefIconPath 
         Height          =   315
         ItemData        =   "Configurator.frx":0000
         Left            =   120
         List            =   "Configurator.frx":0007
         TabIndex        =   7
         Top             =   1920
         Width           =   4095
      End
      Begin VB.CommandButton btnResetAllDef 
         Caption         =   "Reset All"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   2400
         Width           =   1455
      End
      Begin VB.CommandButton btnDefIconBrowse 
         Caption         =   "..."
         Height          =   315
         Left            =   4200
         TabIndex        =   9
         ToolTipText     =   "Browse for another icon"
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label lblName3 
         Caption         =   "Here you can change the default VB icons for each component type, as well as the project, folder and state (overlay) icons."
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   4575
      End
      Begin VB.Label lblName2 
         Caption         =   "Path to icon file"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label lblName1 
         Caption         =   "Select a default icon:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   1815
      End
   End
   Begin VB.PictureBox picFrame 
      HasDC           =   0   'False
      Height          =   3015
      Index           =   1
      Left            =   1560
      ScaleHeight     =   197
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   317
      TabIndex        =   2
      Top             =   0
      Width           =   4815
      Begin VB.ComboBox cboCmpIconPath 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Configurator.frx":0016
         Left            =   120
         List            =   "Configurator.frx":001D
         TabIndex        =   16
         Top             =   2520
         Width           =   3975
      End
      Begin VB.CommandButton btnCmpIconBrowse 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   315
         Left            =   4200
         TabIndex        =   18
         ToolTipText     =   "Browse for another icon"
         Top             =   2520
         Width           =   495
      End
      Begin VB.ListBox lstCmps 
         Height          =   1425
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   4575
      End
      Begin VB.ListBox lstCmpsHidden 
         Height          =   1425
         Left            =   360
         Sorted          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   840
         Visible         =   0   'False
         Width           =   4575
      End
      Begin VB.Label Label1 
         Caption         =   "Path to icon file"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label lblName4 
         Caption         =   "Here you can specify the icon for each component in the currently loaded projects."
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   4695
      End
   End
End
Attribute VB_Name = "ConfigPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : ConfigPanel
' Author    : JP Mortiboys
' Purpose   : Handles the UI configuration of the AddIn
'---------------------------------------------------------------------------------------
Option Explicit

Public VBInstance As VBIDE.VBE
Public m_Connect As Connect
Private WithEvents cmpEvents As VBComponentsEvents
Attribute cmpEvents.VB_VarHelpID = -1
Private WithEvents prjEvents As VBProjectsEvents
Attribute prjEvents.VB_VarHelpID = -1

' Currently selected default icon
Private m_iDefIconSelIdx
' Last path in PickIcon Dialog
Private m_sLastIconPath As String

' Initialization
Public Property Set Connect(c As Connect)
    Set m_Connect = c
    Set VBInstance = m_Connect.VBInstance
    Set prjEvents = VBInstance.Events.VBProjectsEvents
    Set cmpEvents = VBInstance.Events.VBComponentsEvents(Nothing)
End Property

' Browse for a component icon
Private Sub btnCmpIconBrowse_Click()
    If SelectedComponent Is Nothing Then Exit Sub
    Dim sIconPath As String
    sIconPath = RemoveIconIdx(DefaultToBlank(cboCmpIconPath.Text))
    If sIconPath = "" Then sIconPath = m_sLastIconPath
    sIconPath = BrowseForIcon(hwnd, sIconPath)
    If sIconPath <> "" Then
        cboCmpIconPath.Text = sIconPath
        m_sLastIconPath = FolderFromPath(sIconPath)
        cboCmpIconPath_Validate False
    End If
End Sub

' Browse for a default icon
Private Sub btnDefIconBrowse_Click()
    Dim sIconPath As String
    sIconPath = RemoveIconIdx(DefaultToBlank(cboDefIconPath.Text))
    If sIconPath = "" Then sIconPath = m_sLastIconPath
    sIconPath = BrowseForIcon(hwnd, sIconPath)
    If sIconPath <> "" Then
        cboDefIconPath.Text = sIconPath
        m_sLastIconPath = FolderFromPath(sIconPath)
        cboDefIconPath_Validate False
    End If
End Sub

' Reset all
Private Sub btnResetAllDef_Click()
    If MsgBox("Are you sure you want to reset all default icons to the standard settings?", vbYesNo Or vbQuestion) = vbNo Then Exit Sub
    On Error Resume Next
    DeleteSetting App.Path, "Replace Default Icons"
    DrawDefIcons
    InvalidateRectAsNull m_Connect.TreeViewManager.hWndProjectTree, 0, -1
End Sub

' Scroll the default icons
Private Sub btnScrollLeft_Click()
    If picDefIcons.Left < 0 Then picDefIcons.Left = picDefIcons.Left + 20
End Sub
Private Sub btnScrollRight_Click()
    If picDefIcons.Left + picDefIcons.Width > picDefIconsFrame.ScaleWidth Then picDefIcons.Left = picDefIcons.Left - 20
End Sub

' Draw the default icons on the picture box
' Note that its Auto-Redraw is True
Private Sub DrawDefIcons()
    Dim i As Long, n As Long
    Dim X As Long, Y As Long
    Dim hImlTree As Long, hWndTree As Long
    With m_Connect.TreeViewManager
        hWndTree = .hWndProjectTree
        hImlTree = .hImlProjectTree
    End With
    n = ImageList_GetImageCount(hImlTree)
    X = 2: Y = 2
    For i = 0 To n - 1
        ImageList_Draw hImlTree, i, picDefIcons.hdc, X, Y, ILD_NORMAL
        X = X + 20
    Next
End Sub

Private Sub PopulateComponentList()
    Dim prj As VBProject
    Dim cmp As VBComponent
    Dim i As Integer, J As Integer
    ' This is a really pathetic, lazy and hacky way of accomplishing something very simple:
    ' Sorting the Projects and Components hierarchally (is that a word?) and alphabetically.
    ' Using a hidden window for this is really lazy, I must re-write this with a sorting algorithm.
    ' To say that, I should really use a treeview or owner-drawn listbox to show the icons here too... lazy me
    With lstCmpsHidden
        .Clear
        For Each prj In VBInstance.VBProjects
            .AddItem prj.Name
            For Each cmp In prj.VBComponents
                If cmp.Type <> vbext_ct_RelatedDocument And cmp.Type <> vbext_ct_ResFile Then
                    ' Set the name for sorting
                    .AddItem prj.Name & "/" & cmp.Name
                    
                    AddOnceToCombo cboCmpIconPath, m_Connect.TreeViewManager.GetComponentIcon(cmp)
                End If
            Next
        Next
    End With
    With lstCmps
        LockWindowUpdate .hwnd
        .Clear
        For i = 0 To lstCmpsHidden.ListCount
            .AddItem lstCmpsHidden.List(i)
            If .List(i) Like "*/*" Then
                ' This is a component node - strip off the project name
                .List(i) = "   " & Mid$(.List(i), InStr(.List(i), "/") + 1)
                .ItemData(i) = J
            Else
                ' This is a project node - remember it
                J = i
                .ItemData(i) = -1
            End If
        Next
        LockWindowUpdate 0
    End With
End Sub

Private Sub cboCmpIconPath_Click()
    'cboCmpIconPath_Validate False
End Sub

Private Sub cboCmpIconPath_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc(vbCr) Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub cboCmpIconPath_Validate(Cancel As Boolean)
    Dim cmp As VBComponent
    Set cmp = SelectedComponent
    
    If cmp Is Nothing Then Exit Sub
    
    Dim sIconPath As String
    sIconPath = DefaultToBlank(cboCmpIconPath.Text)
    Call m_Connect.TreeViewManager.SetComponentIcon(cmp, sIconPath)
    Call m_Connect.TreeViewManager.ApplyComponentIcon(cmp)
    cboCmpIconPath.Text = BlankToDefault(cboCmpIconPath.Text)
    AddOnceToCombo cboCmpIconPath, cboCmpIconPath.Text
End Sub

Private Sub cboDefIconPath_Click()
    'cboDefIconPath_Validate False
End Sub

Private Sub cboDefIconPath_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc(vbCr) Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
End Sub

Private Sub cboDefIconPath_Validate(Cancel As Boolean)
    Dim sIconPath As String
    sIconPath = DefaultToBlank(cboDefIconPath.Text)
    
    Dim iIconIndex As Long, hIcon As Long, sBuf As String
    SaveSetting App.Title, "Replace Default Icons", CStr(DefIconSelIndex), sIconPath
    If sIconPath = "" Then
        hIcon = ImageList_GetIcon(m_Connect.TreeViewManager.hImlOldProjectTree, DefIconSelIndex, 0)
    Else
        iIconIndex = PathParseIconLocation(sIconPath)
        sIconPath = Left$(sIconPath, lstrlen(sIconPath))
        ExtractIconEx sIconPath, iIconIndex, ByVal 0, hIcon, 1
        sBuf = String$(260, 0)
        ExpandEnvironmentStrings sIconPath, sBuf, 260
        sIconPath = Left$(sBuf, lstrlen(sBuf))
    End If
    ImageList_ReplaceIcon m_Connect.TreeViewManager.hImlProjectTree, DefIconSelIndex, hIcon
    DestroyIcon hIcon
    
    cboDefIconPath.Text = BlankToDefault(cboDefIconPath.Text)
    AddOnceToCombo cboDefIconPath, cboDefIconPath.Text
    shpSelect.Visible = False
    DrawDefIcons
    shpSelect.Visible = True
    InvalidateRectAsNull m_Connect.TreeViewManager.hWndProjectTree, 0, -1
End Sub

Private Sub Form_Load()
    LoadComboStrings cboCmpIconPath
    LoadComboStrings cboDefIconPath

    DrawDefIcons
    DefIconSelIndex = 0
    PopulateComponentList
    
    opFrame(0).Value = True
End Sub

Private Property Get SelectedComponent() As VBComponent
    Dim prj As VBProject
    With lstCmps
        ' Check for no selection
        If .ListIndex = -1 Then Exit Property
        ' Check for project node selection
        If .ItemData(.ListIndex) = -1 Then Exit Property
        ' The ItemData of component nodes is the index of the project node
        ' The text (.List) of a project node is its key
        Set prj = VBInstance.VBProjects(Trim$(.List(.ItemData(.ListIndex))))
        ' The text (.List) of a component node is its key
        Set SelectedComponent = prj.VBComponents(Trim$(.List(.ListIndex)))
    End With
End Property

Private Sub Form_Unload(Cancel As Integer)
    SaveComboStrings cboCmpIconPath
    SaveComboStrings cboDefIconPath
End Sub

Private Sub lstCmps_Click()
    UpdateComponentData
End Sub

Private Sub lstCmps_DblClick()
    btnCmpIconBrowse_Click
End Sub

Private Sub opFrame_Click(Index As Integer)
    Dim i As Integer
    For i = opFrame.LBound To opFrame.UBound
        If opFrame(i).Value Then
            ShowFrame i
        End If
    Next
End Sub

Public Sub ShowFrame(ByVal iFrame As Integer)
    picFrame(iFrame).ZOrder
End Sub

Private Sub picDefIcons_DblClick()
    btnDefIconBrowse_Click
End Sub

Private Sub picDefIcons_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyLeft
        DefIconSelIndex = DefIconSelIndex - 1
    Case vbKeyRight
        DefIconSelIndex = DefIconSelIndex + 1
    End Select
End Sub

Public Property Let DefIconSelIndex(ByVal idx As Long)
    If idx < 0 Then idx = 0
    If idx > 22 Then idx = 22
    If m_iDefIconSelIdx = idx Then Exit Property
    
    m_iDefIconSelIdx = idx
    shpSelect.Move m_iDefIconSelIdx * 20, 0
    If picDefIcons.Left + shpSelect.Left < 0 Then
        picDefIcons.Left = -shpSelect.Left
    ElseIf picDefIcons.Left + shpSelect.Left + 20 > picDefIconsFrame.ScaleWidth Then
        picDefIcons.Left = picDefIconsFrame.ScaleWidth - shpSelect.Left - 20
    End If
    
    cboDefIconPath.Text = BlankToDefault(GetSetting(App.Title, "Replace Default Icons", CStr(idx), ""))
End Property

Public Property Get DefIconSelIndex() As Long
    DefIconSelIndex = m_iDefIconSelIdx
End Property

Private Sub picDefIcons_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DefIconSelIndex = CLng(X) \ 20
End Sub

Private Sub UpdateComponentData()
    Dim cmp As VBComponent
    Set cmp = SelectedComponent
    If cmp Is Nothing Then
        cboCmpIconPath.Text = ""
        cboCmpIconPath.Enabled = False
        btnCmpIconBrowse.Enabled = False
        Exit Sub
    End If
    btnCmpIconBrowse.Enabled = True
    cboCmpIconPath.Enabled = True
    cboCmpIconPath.Text = BlankToDefault(m_Connect.TreeViewManager.GetComponentIcon(cmp))
End Sub

Private Sub cmpEvents_ItemActivated(ByVal VBComponent As VBIDE.VBComponent)
'
End Sub
Private Sub cmpEvents_ItemSelected(ByVal VBComponent As VBIDE.VBComponent)
'
End Sub
Private Sub cmpEvents_ItemAdded(ByVal VBComponent As VBIDE.VBComponent)
    PopulateComponentList
End Sub
Private Sub cmpEvents_ItemReloaded(ByVal VBComponent As VBIDE.VBComponent)
    PopulateComponentList
End Sub
Private Sub cmpEvents_ItemRemoved(ByVal VBComponent As VBIDE.VBComponent)
    PopulateComponentList
End Sub
Private Sub cmpEvents_ItemRenamed(ByVal VBComponent As VBIDE.VBComponent, ByVal OldName As String)
    PopulateComponentList
End Sub
Private Sub prjEvents_ItemActivated(ByVal VBProject As VBIDE.VBProject)
    '
End Sub
Private Sub prjEvents_ItemAdded(ByVal VBProject As VBIDE.VBProject)
    Set cmpEvents = VBInstance.Events.VBComponentsEvents(Nothing)
    PopulateComponentList
End Sub
Private Sub prjEvents_ItemRemoved(ByVal VBProject As VBIDE.VBProject)
    Set cmpEvents = VBInstance.Events.VBComponentsEvents(Nothing)
    PopulateComponentList
End Sub
Private Sub prjEvents_ItemRenamed(ByVal VBProject As VBIDE.VBProject, ByVal OldName As String)
    PopulateComponentList
End Sub
