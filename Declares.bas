Attribute VB_Name = "Declares"
'---------------------------------------------------------------------------------------
' Module    : Declares
' Author    : JP Mortiboys
' Purpose   : All the API declares used
'---------------------------------------------------------------------------------------
Option Explicit

' ****************************************************************************************************
' * Declares for: General GDI                                                                        *
' ****************************************************************************************************
Public Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, phiconLarge As Long, phiconSmall As Long, ByVal nIcons As Long) As Long
Public Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Public Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
Public Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Public Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Public Const DI_NORMAL = &H3
Public Const DI_MASK = &H1
Public Const DI_IMAGE = &H2
Public Declare Function GetIconInfo Lib "user32" (ByVal hIcon As Long, piconinfo As ICONINFO) As Long
Public Type ICONINFO
    fIcon As Long
    xHotspot As Long
    yHotspot As Long
    hbmMask As Long
    hbmColor As Long
End Type
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function LoadIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As String) As Long
Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Public Const IMAGE_ICON = 1
Public Const IMAGE_BITMAP = 0
Public Const IMAGE_CURSOR = 2
Public Const LR_COLOR = &H2
Public Const LR_COPYDELETEORG = &H8
Public Const LR_COPYFROMRESOURCE = &H4000
Public Const LR_COPYRETURNORG = &H4
Public Const LR_CREATEDIBSECTION = &H2000
Public Const LR_DEFAULTCOLOR = &H0
Public Const LR_DEFAULTSIZE = &H40
Public Const LR_LOADFROMFILE = &H10
Public Const LR_LOADMAP3DCOLORS = &H1000
Public Const LR_LOADTRANSPARENT = &H20
Public Const LR_MONOCHROME = &H1
Public Const LR_SHARED = &H8000
Public Const LR_VGACOLOR = &H80

Public Const CLR_NONE As Long = &HFFFFFFFF
Public Const CLR_DEFAULT As Long = &HFF000000
' ****************************************************************************************************
' * Declares for: General Windowing API                                                              *                                                                       *
' ****************************************************************************************************
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Const GWL_STYLE = (-16)
Public Const LBS_SORT = &H2&
Public Const CB_FINDSTRINGEXACT = &H158
Public Declare Function InvalidateRectAsNull Lib "user32" Alias "InvalidateRect" (ByVal hwnd As Long, ByVal lpRect As Long, ByVal bErase As Long) As Long
' ****************************************************************************************************
' * Declares for: Shell API                                                                          *
' ****************************************************************************************************
Private Declare Function SHChangeIconDialog Lib "Shell32" Alias "#62" (ByVal hOwner As Long, ByVal szFilename As String, ByVal dwMaxFile As Long, lpIconIndex As Long) As Long
Public Declare Function PathParseIconLocation Lib "shlwapi.dll" Alias "PathParseIconLocationA" (ByVal pszIconFile As String) As Long
Public Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long
' ****************************************************************************************************
' * Declares for: ImageList API                                                                      *
' ****************************************************************************************************
Public Declare Function ImageList_Add Lib "comctl32.dll" (ByVal hIml As Long, ByVal hbmImage As Long, ByVal hbmMask As Long) As Long
Public Declare Function ImageList_AddMasked Lib "comctl32.dll" (ByVal hIml As Long, ByVal hbmImage As Long, ByVal crMask As Long) As Long
Public Declare Function ImageList_BeginDrag Lib "comctl32.dll" (ByVal himlTrack As Long, ByVal iTrack As Long, ByVal dxHotspot As Long, ByVal dyHotspot As Long) As Long
Public Declare Function ImageList_Copy Lib "comctl32.dll" (ByVal himlDst As Long, ByVal iDst As Long, ByVal himlSrc As Long, ByVal iSrc As Long, ByVal uFlags As Long) As Long
Public Declare Function ImageList_Create Lib "comctl32.dll" (ByVal cx As Long, ByVal cy As Long, ByVal flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
Public Declare Function ImageList_Destroy Lib "comctl32.dll" (ByVal hIml As Long) As Long
Public Declare Function ImageList_DragEnter Lib "comctl32.dll" (ByVal hwndLock As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function ImageList_DragLeave Lib "comctl32.dll" (ByVal hwndLock As Long) As Long
Public Declare Function ImageList_DragMove Lib "comctl32.dll" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function ImageList_DragShowNolock Lib "comctl32.dll" (ByVal fShow As Long) As Long
Public Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal hIml As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal X As Long, ByVal Y As Long, ByVal fStyle As Long) As Long
Public Declare Function ImageList_DrawEx Lib "comctl32.dll" (ByVal hIml As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal rgbBk As Long, ByVal rgbFg As Long, ByVal fStyle As Long) As Long
'Public Declare Function ImageList_DrawIndirect Lib "comctl32.dll" (ByRef pimldp As IMAGELISTDRAWPARAMS) As Long
Public Declare Function ImageList_Duplicate Lib "comctl32.dll" (ByVal hIml As Long) As Long
Public Declare Sub ImageList_EndDrag Lib "comctl32.dll" ()
Public Declare Function ImageList_GetBkColor Lib "comctl32.dll" (ByVal hIml As Long) As Long
'Public Declare Function ImageList_GetDragImage Lib "comctl32.dll" (ByRef ppt As Point, ByRef pptHotspot As Point) As Long
Public Declare Function ImageList_GetIcon Lib "comctl32.dll" (ByVal hIml As Long, ByVal i As Long, ByVal flags As Long) As Long
Public Declare Function ImageList_ExtractIcon Lib "comctl32.dll" (ByVal hIml As Long, ByVal hInstance As Long, ByVal i As Long) As Long
Public Declare Function ImageList_GetIconSize Lib "comctl32.dll" (ByVal hIml As Long, ByRef cx As Long, ByRef cy As Long) As Long
Public Declare Function ImageList_GetImageCount Lib "comctl32.dll" (ByVal hIml As Long) As Long
'Public Declare Function ImageList_GetImageInfo Lib "comctl32.dll" (ByVal hIml As Long, ByVal i As Long, ByRef pImageInfo As IMAGEINFO) As Long
Public Declare Function ImageList_LoadImage Lib "comctl32.dll" (ByVal hi As Long, ByVal lpbmp As String, ByVal cx As Long, ByVal cGrow As Long, ByVal crMask As Long, ByVal uType As Long, ByVal uFlags As Long) As Long
Public Declare Function ImageList_Merge Lib "comctl32.dll" (ByVal himl1 As Long, ByVal i1 As Long, ByVal hIml2 As Long, ByVal i2 As Long, ByVal dx As Long, ByVal dy As Long) As Long
Public Declare Function ImageList_Read Lib "comctl32.dll" (ByRef pstm As Long) As Long
Public Declare Function ImageList_Remove Lib "comctl32.dll" (ByVal hIml As Long, ByVal i As Long) As Long
Public Declare Function ImageList_Replace Lib "comctl32.dll" (ByVal hIml As Long, ByVal i As Long, ByVal hbmImage As Long, ByVal hbmMask As Long) As Long
Public Declare Function ImageList_ReplaceIcon Lib "comctl32.dll" (ByVal hIml As Long, ByVal i As Long, ByVal hIcon As Long) As Long
Public Declare Function ImageList_SetBkColor Lib "comctl32.dll" (ByVal hIml As Long, ByVal clrBk As Long) As Long
Public Declare Function ImageList_SetDragCursorImage Lib "comctl32.dll" (ByVal himlDrag As Long, ByVal iDrag As Long, ByVal dxHotspot As Long, ByVal dyHotspot As Long) As Long
Public Declare Function ImageList_SetIconSize Lib "comctl32.dll" (ByVal hIml As Long, ByVal cx As Long, ByVal cy As Long) As Long
Public Declare Function ImageList_SetImageCount Lib "comctl32.dll" (ByVal hIml As Long, ByVal uNewCount As Long) As Long
Public Declare Function ImageList_SetOverlayImage Lib "comctl32.dll" (ByVal hIml As Long, ByVal iImage As Long, ByVal iOverlay As Long) As Long
Public Declare Function ImageList_Write Lib "comctl32.dll" (ByVal hIml As Long, ByRef pstm As Long) As Long
Public Const ILC_COLOR As Long = &H0
Public Const ILC_COLOR16 As Long = &H10
Public Const ILC_COLOR24 As Long = &H18
Public Const ILC_COLOR32 As Long = &H20
Public Const ILC_COLOR4 As Long = &H4
Public Const ILC_COLOR8 As Long = &H8
Public Const ILC_COLORDDB As Long = &HFE
Public Const ILC_MASK As Long = &H1
Public Const ILC_PALETTE As Long = &H800
Public Const ILD_BLEND25 As Long = &H2
Public Const ILD_BLEND50 As Long = &H4
Public Const ILD_BLEND As Long = ILD_BLEND50
Public Const ILD_FOCUS As Long = ILD_BLEND25
Public Const ILD_IMAGE As Long = &H20
Public Const ILD_MASK As Long = &H10
Public Const ILD_NORMAL As Long = &H0
Public Const ILD_OVERLAYMASK As Long = &HF00
Public Const ILD_ROP As Long = &H40
Public Const ILD_SELECTED As Long = ILD_BLEND50
Public Const ILD_TRANSPARENT As Long = &H1

' ****************************************************************************************************
' * Declares for: TreeView API                                                                       *
' ****************************************************************************************************
Public Const TV_FIRST As Long = &H1100
Public Const TVM_GETIMAGELIST As Long = (TV_FIRST + 8)
Public Const TVM_SETIMAGELIST As Long = (TV_FIRST + 9)
Public Const TVM_GETITEM = 4364
Public Const TVM_SETITEM = 4365
Public Const TVM_GETNEXTITEM = 4362

Public Const TVSIL_NORMAL As Long = 0
Public Const TVSIL_STATE As Long = 2

Public Const TVIF_IMAGE As Long = 2
Public Const TVIF_SELECTEDIMAGE As Long = 32
Public Const TVIF_TEXT = 1

Public Const TVGN_ROOT = 0
Public Const TVGN_CHILD = 4
Public Const TVGN_NEXT = 1

Public Type TVITEM   ' was TV_ITEM
  mask As Long
  hItem As Long
  State As Long
  stateMask As Long
  pszText As String   ' Long   ' pointer
  cchTextMax As Long
  iImage As Long
  iSelectedImage As Long
  cChildren As Long
  lParam As Long
End Type
'=========Checking OS staff=============
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (LpVersionInformation As OSVERSIONINFO) As Long

'
Public Function INDEXTOOVERLAYMASK(ByVal iState As Long) As Long
' #define INDEXTOOVERLAYMASK(i) ((i) << 8)
  'INDEXTOOVERLAYMASK = iState * (2 ^ 8)
  INDEXTOOVERLAYMASK = iState * 256
End Function

' Returns the text of the specified treeview item if successful,
' returns an empty string otherwise.

'   hwndTV      - treeview's window handle
'   hItem          - item's handle whose text is to be to returned
'   cbItem        - length of the specified item's text.

Public Function GetTVItemText(hwndTV As Long, _
                                                  hItem As Long, _
                                                  Optional cbItem As Long = 256) As String
  Dim tvi As TVITEM
  
  ' Initialize the struct to retrieve the item's text.
  tvi.mask = TVIF_TEXT
  tvi.hItem = hItem
  tvi.pszText = String$(cbItem, 0)
  tvi.cchTextMax = cbItem
  
  If TreeView_GetItem(hwndTV, tvi) Then
    GetTVItemText = GetStrFromBufferA(tvi.pszText)
  End If
End Function

Public Function TreeView_GetItem(hwnd As Long, pitem As TVITEM) As Boolean
  TreeView_GetItem = SendMessage(hwnd, TVM_GETITEM, 0, pitem)
End Function

' Returns the string before first null char encountered (if any) from an ANSII string.
Public Function GetStrFromBufferA(sz As String) As String
  If InStr(sz, vbNullChar) Then
    GetStrFromBufferA = Left$(sz, InStr(sz, vbNullChar) - 1)
  Else
    ' If sz had no null char, the Left$ function
    ' above would return a zero length string ("").
    GetStrFromBufferA = sz
  End If
End Function

' Extract the filename part from a full path
Public Function FileNameFromPath(ByVal sPath As String) As String
    Dim i As Integer
    i = InStrRev(sPath, "\")
    If i < 1 Then
        FileNameFromPath = sPath
    Else
        FileNameFromPath = Mid$(sPath, i + 1)
    End If
End Function

' Extract the folder part from a full path
Public Function FolderFromPath(ByVal sPath As String) As String
    Dim i As Integer
    i = InStrRev(sPath, "\")
    If i < 1 Then
        FolderFromPath = ""
    Else
        FolderFromPath = Left$(sPath, i - 1)
    End If
End Function

' Remove the ",index" from an icon path
Public Function RemoveIconIdx(ByVal sPath As String)
    Dim i As Integer
    i = InStrRev(sPath, ",")
    If i < 1 Then
        RemoveIconIdx = sPath
    Else
        RemoveIconIdx = Left$(sPath, i - 1)
    End If
End Function

' A quick note: SHChangeIconDialog is an undocumented windows function, and as such
' it only accepts ANSI on 9x/Me and UNICODE on NT/2k/XP, so this function handles both
Public Function BrowseForIcon(ByVal hWndOwner As Long, ByVal sDir As String) As String
    Dim sFileName As String, iIndex As Long
    sFileName = sDir & String$(260 - Len(sDir), 0)
    If IsWindowsNT Then sFileName = StrConv(sFileName, vbUnicode)
    If SHChangeIconDialog(hWndOwner, sFileName, 260, iIndex) = 0 Then
        BrowseForIcon = ""
    Else
        If IsWindowsNT Then sFileName = StrConv(sFileName, vbFromUnicode)
        BrowseForIcon = GetStrFromBufferA(sFileName) & "," & iIndex
    End If
End Function

Public Function IsWindowsNT() As Boolean
   Dim verinfo As OSVERSIONINFO
   verinfo.dwOSVersionInfoSize = Len(verinfo)
   If (GetVersionEx(verinfo)) = 0 Then Exit Function
   If verinfo.dwPlatformId = 2 Then IsWindowsNT = True
End Function

