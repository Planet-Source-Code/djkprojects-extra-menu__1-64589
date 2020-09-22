VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type MENUINFO
  cbSize As Long
  fMask As Long
  dwStyle As Long
  cyMax As Long
  hbrBack As Long
  dwContextHelpID As Long
  dwMenuData As Long
End Type

Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wid As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
    hbmpItem As Long
End Type

Private Const MIM_BACKGROUND = &H2
Private Const MIM_APPLYTOSUBMENUS = &H80000000

Private Const MIIM_BITMAP = &H80
Private Const MIIM_CHECKMARKS = &H8

Private Const MIM_MAXHEIGHT = &H1


Private mnuHwnd As Long
Private mnuInfo As MENUINFO
Private mnuItemInfo As MENUITEMINFO

Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function SetMenuInfo Lib "user32.dll" (ByVal hMenu As Long, ByRef LPCMENUINFO As MENUINFO) As Long
Private Declare Function SetMenuItemInfo Lib "user32.dll" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, ByRef lpcMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long

Public Property Let MainMenuHwnd(ByVal hwnd As Long)
   mnuHwnd = GetMenu(hwnd)
End Property

Public Property Get MainMenuHwnd() As Long
   MainMenuHwnd = mnuHwnd
End Property

Public Function SubMenu(ByVal Parent As Long, ByVal nPos As Integer) As Long
SubMenu = GetSubMenu(Parent, nPos - 1)
End Function

Public Function SetBkgColor(ByVal SubMenu As Long, ByVal Color As Long, Optional ApplyToSubMenus As Boolean = True) As Long

With mnuInfo
        .cbSize = Len(mnuInfo)
        .fMask = MIM_BACKGROUND
        If ApplyToSubMenus = True Then .fMask = .fMask Or MIM_APPLYTOSUBMENUS
        .hbrBack = CreateSolidBrush(Color)
End With

SetBkgColor = SetMenuInfo(SubMenu, mnuInfo)

End Function

Public Function SetIcon(ByVal SubMenu As Long, ByVal ItemIndex As Long, ByVal hIcon As Long, Optional ByID As Boolean = False) As Long

With mnuItemInfo
    .cbSize = Len(mnuItemInfo)
    .fMask = MIIM_BITMAP
    .hbmpItem = hIcon
End With

If ByID = False Then
    SetIcon = SetMenuItemInfo(SubMenu, ItemIndex - 1, True, mnuItemInfo)
Else
    SetIcon = SetMenuItemInfo(GetSubMenu(mnuHwnd, SubMenu - 1), ItemIndex, False, mnuItemInfo)
End If

End Function

Public Function SetCheckIcons(ByVal SubMenu As Long, ByVal ItemIndex As Long, ByVal hCheckIcon As Long, ByVal hUncheckIcon As Long, Optional ByID As Boolean = False) As Long

With mnuItemInfo
        .cbSize = Len(mnuItemInfo)
        .fMask = MIIM_CHECKMARKS
        .hbmpChecked = hCheckIcon
        .hbmpUnchecked = hUncheckIcon
End With
    
If ByID = False Then
    SetCheckIcons = SetMenuItemInfo(SubMenu, ItemIndex - 1, True, mnuItemInfo)
Else
    SetCheckIcons = SetMenuItemInfo(GetSubMenu(mnuHwnd, SubMenu - 1), ItemIndex, False, mnuItemInfo)
End If

End Function

Public Function SubMaxHeight(ByVal SubMenu As Long, ByVal Height As Integer) As Long

With mnuInfo
        .cbSize = Len(mnuInfo)
        .fMask = MIM_MAXHEIGHT
        .cyMax = Height
End With

SubMaxHeight = SetMenuInfo(SubMenu, mnuInfo)

End Function
