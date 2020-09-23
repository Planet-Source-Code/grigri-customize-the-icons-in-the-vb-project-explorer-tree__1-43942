VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IconCache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------------------
' Module    : IconCache
' Author    : JP Mortiboys
' Purpose   : Handles a cache of icons in an imagelist, refcounting and automatic releasing
'---------------------------------------------------------------------------------------
Option Explicit

Private Const GROW As Long = 10

' Handle to the imagelist
Private hIml As Long
' Our lowest index, below which we will not touch
Private iBaseIdx As Long

Private Cache() As CACHE_ENTRY
Private KeyLUT As Collection    ' LookUp Table by key [ contains indices into Cache() ]
Private iNextFreeIndex As Long

Private Type CACHE_ENTRY
    IconPath As String
    RefCount As Long
End Type

Public Sub Init(ByVal hImageList As Long, Optional ByVal iBaseIndex As Long = 0, Optional ByVal iInitSize As Long = 10)
    hIml = hImageList
    iBaseIdx = iBaseIndex
    ReDim Cache(0 To iInitSize)
    ImageList_SetImageCount hIml, iBaseIdx + iInitSize
    Set KeyLUT = New Collection
    iNextFreeIndex = 0
End Sub

Public Sub ClearUp()
    ImageList_SetImageCount hIml, iBaseIdx
End Sub

Public Function UseIcon(ByVal sIconPath As String) As Long
    Dim iIndex As Long
    On Error Resume Next
    iIndex = KeyLUT(sIconPath)
    If Err.Number <> 0 Then
        ' Icon not loaded, load it now
        iIndex = AddItemToCache(sIconPath)
    End If
    On Error GoTo 0
    ' Offset the index by iBaseIndex - the end user doesn't care about our cache,
    ' he's interested in the REAL ImageList index
    iIndex = iIndex + iBaseIdx
    ' Increase reference count
    UseIconIndex iIndex
    ' Return index
    UseIcon = iIndex
End Function

' These 2 functions use indices referring to the ImageList itself, not our cache
' therefore we need to test and offset by iBaseIdx
Public Sub UseIconIndex(ByVal iIndex As Long)
    If iIndex < iBaseIdx Then Exit Sub
    With Cache(iIndex - iBaseIdx)
        .RefCount = .RefCount + 1
        If iNextFreeIndex = iIndex Then
            iNextFreeIndex = -1
            NextFreeIndex
        End If
        'DebugOut "UseIconIndex(" & iIndex - iBaseIdx & "), count=" & .RefCount
    End With
End Sub
Public Sub ReleaseIconIndex(ByVal iIndex As Long)
    If iIndex < iBaseIdx Then Exit Sub
    With Cache(iIndex - iBaseIdx)
        .RefCount = .RefCount - 1
'        DebugOut "ReleaseIconIndex(" & iIndex - iBaseIdx & "), count=" & .RefCount

        ' If this is now unused and lower than the next free index
        ' then use it as the next free index
        If .RefCount <= 0 Then
            If (iIndex - iBaseIdx) < iNextFreeIndex Then
                iNextFreeIndex = iIndex - iBaseIdx
            End If
        End If
    End With
End Sub

Private Function AddItemToCache(ByVal sIconPath As String) As Long
    Dim iIndex As Long, iIconIndex As Long, hIcon As Long
    Dim sBuf As String
    iIndex = NextFreeIndex
    With Cache(iIndex)
        ' Remove entry from lookup table
        If .IconPath <> "" Then KeyLUT.Remove .IconPath
        ' Init the data
        .IconPath = sIconPath
        ' Although it's cached, it's not necessarily used YET...
        ' if is is to be used, this will be set in UseIconIndex()
        .RefCount = 0
        
        'DebugOut "AddItemToCache(" & iIndex & "), count=" & .RefCount

        ' Add new entry to lookup table
        KeyLUT.Add iIndex, .IconPath
    End With
    
    iIconIndex = PathParseIconLocation(sIconPath)
    sIconPath = Left$(sIconPath, lstrlen(sIconPath))
    sBuf = String$(260, 0)
    ExpandEnvironmentStrings sIconPath, sBuf, 260
    sIconPath = Left$(sBuf, lstrlen(sBuf))
    
    ExtractIconEx sIconPath, iIconIndex, ByVal 0, hIcon, 1
    ImageList_ReplaceIcon hIml, iIndex + iBaseIdx, hIcon
    DestroyIcon hIcon
    
    AddItemToCache = iIndex
End Function

Private Function NextFreeIndex() As Long
    NextFreeIndex = iNextFreeIndex
    Dim I As Long
    For I = iNextFreeIndex + 1 To UBound(Cache)
        If Cache(I).RefCount <= 0 Then
            iNextFreeIndex = I
            Exit Function
        End If
    Next
    iNextFreeIndex = UBound(Cache) + 1
    ReDim Preserve Cache(0 To UBound(Cache) + GROW)
    ImageList_SetImageCount hIml, iBaseIdx + UBound(Cache)
End Function
