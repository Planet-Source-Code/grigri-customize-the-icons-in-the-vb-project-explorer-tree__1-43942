VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TreeViewManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Module    : TreeViewManager
' Author    : JP Mortiboys
' Purpose   : Manages the nitty-gritty details of the TreeView
'---------------------------------------------------------------------------------------
Option Explicit

Private Const WINDOWCLASS_PROJECT As String = "PROJECT"
Private Const WINDOWCLASS_TREEVIEW As String = "SysTreeView32"

Private VBInstance As VBE
Private hWndTree As Long
Private hImlOld As Long
Private hImlNew As Long
Private WithEvents cmpEvents As VBComponentsEvents
Attribute cmpEvents.VB_VarHelpID = -1
Private WithEvents prjEvents As VBProjectsEvents
Attribute prjEvents.VB_VarHelpID = -1
Private IconCache As IconCache

Public Property Get hWndProjectTree() As Long
    hWndProjectTree = hWndTree
End Property
Public Property Get hImlProjectTree() As Long
    hImlProjectTree = hImlNew
End Property
Public Property Get hImlOldProjectTree() As Long
    hImlOldProjectTree = hImlOld
End Property

Public Function Init(vbInst As VBE) As Boolean
    Set VBInstance = vbInst
    
    ' Find the TreeView window
    hWndTree = VBInstance.MainWindow.hwnd
    hWndTree = FindWindowEx(hWndTree, 0, WINDOWCLASS_PROJECT, vbNullString)
    hWndTree = FindWindowEx(hWndTree, 0, WINDOWCLASS_TREEVIEW, vbNullString)
    If hWndTree = 0 Then
        'Err.Raise vbObjectError, "TreeViewManager.Init()", "Cannot find Window Handle of Project Browser TreeView"
        MsgBox "Cannot find Window Handle of Project Browser TreeView"
        Init = False
        Exit Function
    End If
    
    Dim nImages As Long, iIndex As Long
    Dim hIcon As Long, sIconPath As String, iIconIndex As Long
    Dim sBuf As String
    
    ' Retrieve old ImageList (16-colour)
    hImlOld = SendMessage(hWndTree, TVM_GETIMAGELIST, TVSIL_NORMAL, ByVal 0)
    nImages = ImageList_GetImageCount(hImlOld)
    'MsgBox "hImlOld: " & hImlOld & vbCrLf & "Count: " & nImages
    ' Make a new one (now you're talking! 32-bit colour)
    hImlNew = ImageList_Create(16, 16, ILC_COLOR32 Or ILC_MASK, nImages, 10)
    If hImlNew = 0 Then
        MsgBox "Error creating imagelist"
        Exit Function
    End If
    ImageList_SetBkColor hImlNew, vbWhite
    
    ImageList_SetImageCount hImlNew, nImages
    
    ' Copy the old images into the new list, or the customized icon if it's set in the registry
    ' This maps the colours correctly from 16-colour to 32-bit colour
    For iIndex = 0 To nImages - 1
        ' Try and read the setting for this icon
        sIconPath = GetSetting(App.Title, "Replace Default Icons", CStr(iIndex), "")
        hIcon = 0
        If sIconPath <> "" Then
            ' This icon has been customized
            ' Load the icon
            iIconIndex = PathParseIconLocation(sIconPath)
            sIconPath = Left$(sIconPath, lstrlen(sIconPath))
            ExtractIconEx sIconPath, iIconIndex, ByVal 0, hIcon, 1
            sBuf = String$(260, 0)
            ExpandEnvironmentStrings sIconPath, sBuf, 260
            sIconPath = Left$(sBuf, lstrlen(sBuf))
        End If
        If hIcon = 0 Then
            ' Either the icon hadn't been customized or the registry
            ' path was invalid (couldn't load)
            ' In either case, copy the icon from the old imagelist
            hIcon = ImageList_GetIcon(hImlOld, iIndex, 0)
        End If
        ImageList_ReplaceIcon hImlNew, iIndex, hIcon
        DestroyIcon hIcon
    Next
    ' The original imagelist has 3 overlays set - we need to set them in our imagelist too
    ImageList_SetOverlayImage hImlNew, 1, 1
    ImageList_SetOverlayImage hImlNew, 2, 2
    ImageList_SetOverlayImage hImlNew, 3, 3
    
    ' Tell the Tree to use our ImageList
    If SendMessage(hWndTree, TVM_SETIMAGELIST, TVSIL_NORMAL, ByVal hImlNew) = 0 Then
        MsgBox "Cannot set image list"
    End If
    ' Setup our cache object
    Set IconCache = New IconCache
    IconCache.Init hImlNew, nImages
    ' Initialize all icons
    Dim prj As VBProject
    Dim cmp As VBComponent
    For Each prj In VBInstance.VBProjects
        For Each cmp In prj.VBComponents
            ApplyComponentIcon cmp
        Next
    Next
    ' Sink events
    Set prjEvents = VBInstance.Events.VBProjectsEvents
    Set cmpEvents = VBInstance.Events.VBComponentsEvents(Nothing)
End Function

Public Sub ClearUp()
    If hWndTree = 0 Then Exit Sub
    ' Use default icons for everyone
    Dim prj As VBProject
    Dim cmp As VBComponent
    For Each prj In VBInstance.VBProjects
        For Each cmp In prj.VBComponents
            UseDefaultComponentIcon cmp
        Next
    Next
    ' Use old imagelist
    SendMessage hWndTree, TVM_SETIMAGELIST, TVSIL_NORMAL, ByVal hImlOld
    ' Delete ours
    ImageList_Destroy hImlNew
End Sub

Private Sub Class_Terminate()
    ClearUp
End Sub

Private Sub cmpEvents_ItemActivated(ByVal VBComponent As VBIDE.VBComponent)
    '
End Sub

Private Sub cmpEvents_ItemAdded(ByVal VBComponent As VBIDE.VBComponent)
    ApplyComponentIcon VBComponent
End Sub

Private Sub cmpEvents_ItemReloaded(ByVal VBComponent As VBIDE.VBComponent)
    ApplyComponentIcon VBComponent
End Sub

Private Sub cmpEvents_ItemRemoved(ByVal VBComponent As VBIDE.VBComponent)
    Dim hItem As Long
    
    hItem = FindComponentNodeHandle(VBComponent)
    If hItem = 0 Then
        MsgBox "Error - Cannot find node for item '" & VBComponent.Name & "'"
        Exit Sub
    End If
    
    Dim Item As TVITEM
    
    Item.hItem = hItem
    Item.mask = TVIF_IMAGE Or TVIF_SELECTEDIMAGE
    Call SendMessage(hWndTree, TVM_GETITEM, 0, Item)
    IconCache.ReleaseIconIndex Item.iImage
    
    ' Not strictly necessary, but looks cleaner
    UseDefaultComponentIcon VBComponent
End Sub

Private Sub cmpEvents_ItemRenamed(ByVal VBComponent As VBIDE.VBComponent, ByVal OldName As String)
    '
End Sub

Private Sub prjEvents_ItemActivated(ByVal VBProject As VBIDE.VBProject)
    '
End Sub

Private Sub prjEvents_ItemAdded(ByVal VBProject As VBIDE.VBProject)
    Set cmpEvents = VBInstance.Events.VBComponentsEvents(Nothing)
End Sub

Private Sub prjEvents_ItemRemoved(ByVal VBProject As VBIDE.VBProject)
    Set cmpEvents = VBInstance.Events.VBComponentsEvents(Nothing)
End Sub

Private Sub prjEvents_ItemRenamed(ByVal VBProject As VBIDE.VBProject, ByVal OldName As String)
    '
End Sub

'
Public Sub SetComponentIcon(ByVal cmp As VBComponent, ByVal sIconFile As String)
    If cmp.Type = vbext_ct_RelatedDocument Or cmp.Type = vbext_ct_ResFile Then
        ' Might embed this data in the project file, at a later date...
    Else
        cmp.WriteProperty "VBIconsIDE_IconPath", sIconFile
        cmp.IsDirty = True
    End If
End Sub
Public Function GetComponentIcon(ByVal cmp As VBComponent) As String
    On Error Resume Next
    Err.Clear
    If cmp.Type = vbext_ct_RelatedDocument Or cmp.Type = vbext_ct_ResFile Then
        ' Might embed this data in the project file, at a later date...
        GetComponentIcon = ""
    Else
        GetComponentIcon = cmp.ReadProperty("VBIconsIDE_IconPath")
    End If
    If Err Then
        GetComponentIcon = ""
    End If
End Function
Public Sub ApplyComponentIcon(ByVal cmp As VBComponent)
    Dim hItem As Long, sIconPath As String, iIconIndex As Long
    
    hItem = FindComponentNodeHandle(cmp)
    If hItem = 0 Then
        MsgBox "Error - Cannot find node for item '" & cmp.Name & "'"
        Exit Sub
    End If
    
    Dim Item As TVITEM
    
    Item.hItem = hItem
    Item.mask = TVIF_IMAGE Or TVIF_SELECTEDIMAGE
    Call SendMessage(hWndTree, TVM_GETITEM, 0, Item)
    ' We're not using this one any more... so let the cache free it up
    IconCache.ReleaseIconIndex Item.iImage
    ' Retrieves the path
    sIconPath = GetComponentIcon(cmp)
    If sIconPath = "" Then
        UseDefaultComponentIcon cmp
        Exit Sub
    End If
    
    ' Load it from the cache and increase the refcount
    iIconIndex = IconCache.UseIcon(sIconPath)
    ' Set the icon for the treeview
    Item.hItem = hItem
    Item.mask = TVIF_IMAGE Or TVIF_SELECTEDIMAGE
    Item.iImage = iIconIndex
    Item.iSelectedImage = Item.iImage
    Call SendMessage(hWndTree, TVM_SETITEM, 0, Item)
End Sub

' Quite straighforward...
Public Sub UseDefaultComponentIcon(ByVal cmp As VBComponent)
    Dim hItem As Long, iIconIndex As Long
    
    hItem = FindComponentNodeHandle(cmp)
    If hItem = 0 Then
        MsgBox "Error - Cannot find node for item '" & cmp.Name & "'"
        Exit Sub
    End If
    
    Select Case cmp.Type
    Case vbext_ct_StdModule
        iIconIndex = 8
    Case vbext_ct_ClassModule
        iIconIndex = 9
    Case vbext_ct_MSForm
        iIconIndex = 11
    Case vbext_ct_ResFile
        iIconIndex = 10
    Case vbext_ct_VBForm
        iIconIndex = IIf(CBool(cmp.Properties("MDIChild")), 7, 5)
    Case vbext_ct_VBMDIForm
        iIconIndex = 6
    Case vbext_ct_PropPage
        iIconIndex = 14
    Case vbext_ct_UserControl
        iIconIndex = 13
    Case vbext_ct_DocObject
        iIconIndex = 15
    Case vbext_ct_RelatedDocument
        iIconIndex = 12
    Case vbext_ct_ActiveXDesigner
        iIconIndex = 18
    End Select
    
    Dim Item As TVITEM
    Item.hItem = hItem
    Item.mask = TVIF_IMAGE Or TVIF_SELECTEDIMAGE
    Item.iImage = iIconIndex
    Item.iSelectedImage = iIconIndex
    
    Call SendMessage(hWndTree, TVM_SETITEM, 0, Item)
End Sub

' Returns the text that the project tree node should have for a given component
' Related Docs and ResFiles don't have Names
' if the component is not saved its displayed filename will be the same as its name
Private Function ComponentNodeText(ByVal cmp As VBComponent) As String
    If cmp.Type <> vbext_ct_RelatedDocument And cmp.Type <> vbext_ct_ResFile Then
        ComponentNodeText = cmp.Name
    End If
    If cmp.FileNames(1) = "" Then
        ComponentNodeText = ComponentNodeText & " (" & cmp.Name & ")"
    Else
        ComponentNodeText = ComponentNodeText & " (" & FileNameFromPath(cmp.FileNames(1)) & ")"
    End If
End Function

' Find the TreeView handle to the component node
Private Function FindComponentNodeHandle(ByVal cmp As VBComponent) As Long
    FindComponentNodeHandle = FindItemRecursiveHelper(ComponentNodeText(cmp), SendMessage(hWndTree, TVM_GETNEXTITEM, TVGN_ROOT, ByVal 0))
End Function

' Note that this function can accept wildcards (* and ?)
Public Function FindItemRecursive(sItemText As String) As Long
    FindItemRecursive = FindItemRecursiveHelper(sItemText, SendMessage(hWndTree, TVM_GETNEXTITEM, TVGN_ROOT, ByVal 0))
End Function

' Note that this function can accept wildcards (* and ?)
Private Function FindItemRecursiveHelper(sItemText As String, ByVal hItem As Long) As Long
    If hItem = 0 Then Exit Function
    Dim s As String
    s = GetTVItemText(hWndTree, hItem)
    'MsgBox s & IIf(s Like sItemText, " like ", " not like ") & sItemText
    If s Like sItemText Then
        FindItemRecursiveHelper = hItem
        Exit Function
    End If
    Dim hTempItem As Long
    hTempItem = SendMessage(hWndTree, TVM_GETNEXTITEM, TVGN_CHILD, ByVal hItem)
    If hTempItem <> 0 Then
        FindItemRecursiveHelper = FindItemRecursiveHelper(sItemText, hTempItem)
        If FindItemRecursiveHelper <> 0 Then Exit Function
    End If
    
    hTempItem = SendMessage(hWndTree, TVM_GETNEXTITEM, TVGN_NEXT, ByVal hItem)
    If hTempItem <> 0 Then
        FindItemRecursiveHelper = FindItemRecursiveHelper(sItemText, hTempItem)
    End If
End Function
