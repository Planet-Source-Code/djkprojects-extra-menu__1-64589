VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Extra Menu"
   ClientHeight    =   3225
   ClientLeft      =   150
   ClientTop       =   420
   ClientWidth     =   6120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ImageList imgLarge 
      Left            =   3480
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   16777215
      UseMaskColor    =   0   'False
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":0000
            Key             =   "cancel"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":0712
            Key             =   "key"
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList imgSmall 
      Left            =   2760
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":0E24
            Key             =   "bag"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":1176
            Key             =   "window"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":14C8
            Key             =   "no"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":1722
            Key             =   "yes"
         EndProperty
      EndProperty
   End
   Begin VB.Menu submnu1 
      Caption         =   "First SubMenu"
      NegotiatePosition=   3  'Right
      Begin VB.Menu mnu2 
         Caption         =   " Menu Item 1"
      End
      Begin VB.Menu mnu3 
         Caption         =   " Menu Item 2"
         Begin VB.Menu mnu4 
            Caption         =   " Menu Item 21"
         End
         Begin VB.Menu mnu5 
            Caption         =   " Menu Item 22"
            Begin VB.Menu mnu6 
               Caption         =   " Menu Item 221"
            End
         End
         Begin VB.Menu mnu7 
            Caption         =   " Menu Item 23"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnu8 
            Caption         =   " Menu Item 24"
         End
      End
      Begin VB.Menu mnu9 
         Caption         =   " Menu Item3"
      End
   End
   Begin VB.Menu submnu2 
      Caption         =   "Second SubMenu"
      Begin VB.Menu mnu10 
         Caption         =   " Menu Item 1"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Copyright (C) 2005 DJK's Projects (DJK)

'Wbesite:    http://djkprojects.prv.pl/
'Contact:       djkprojects@interia.pl

'Autor przyk³adu nie odpowiada za ewentualne szkody wywo³ane dzia³aniem
'poni¿szego kodu.

Private myMenu As New clsMenu

Private Sub Form_Load()

Dim submnu As Long
Dim mnuitem As Long

'main menu hwnd
myMenu.MainMenuHwnd = hwnd
    
'main menu background
myMenu.SetBkgColor myMenu.MainMenuHwnd, RGB(183, 215, 232), False
    
'First method - we use hwnds - step by step to our item
With myMenu
    
    'get first submenu hwnd
    submnu = .SubMenu(.MainMenuHwnd, 1)

    'first subMenu backgroundcolor - BLUE
    myMenu.SetBkgColor submnu, RGB(155, 197, 219), False

    'submenu 1 icons
   .SetIcon submnu, 1, imgSmall.ListImages(1).Picture
   .SetIcon submnu, 2, imgSmall.ListImages(2).Picture
   .SetIcon submnu, 3, imgLarge.ListImages(2).Picture
    
    'get menuitem 2 hwnd
    mnuitem = myMenu.SubMenu(submnu, 2)
    
    'menuitem 2 backgroundcolor - GREEN
    myMenu.SetBkgColor mnuitem, RGB(155, 230, 100), False
    
    'max height of menuitem 2 sub
    .SubMaxHeight mnuitem, 1
    
    'icon and checks for menuitem 21
    myMenu.SetIcon mnuitem, 1, imgLarge.ListImages("key").Picture
    myMenu.SetCheckIcons mnuitem, 1, imgSmall.ListImages("yes").Picture, imgSmall.ListImages("no").Picture

    'icon for menuitem 23
    myMenu.SetIcon mnuitem, 3, imgLarge.ListImages("cancel").Picture

    'get menuitem 22 hwnd
    mnuitem = myMenu.SubMenu(mnuitem, 2)

    'menuitem 22 backgroundcolor - RED
    myMenu.SetBkgColor mnuitem, RGB(214, 18, 64), False
    
    'checks for menu item 3
    myMenu.SetCheckIcons submnu, 3, imgSmall.ListImages("yes").Picture, imgSmall.ListImages("no").Picture
        
    'get second submenu hwnd
    submnu = .SubMenu(.MainMenuHwnd, 2)

    'backgroundcolor for second SubMenu
    myMenu.SetBkgColor submnu, RGB(155, 230, 100)

    'checks for menu item 1
    myMenu.SetCheckIcons submnu, 1, imgSmall.ListImages("yes").Picture, imgSmall.ListImages("no").Picture
End With


'Second method - we use menu ID if we know it ;) - but it doesn't work with items which have subitems
With myMenu
    'menuitem 221
    .SetIcon 1, 6, imgLarge.ListImages("cancel").Picture, True
    .SetCheckIcons 1, 6, imgSmall.ListImages("yes").Picture, imgSmall.ListImages("no").Picture, True
End With

End Sub

Private Sub mnu10_Click()
mnu10.Checked = Not mnu10.Checked
End Sub

Private Sub mnu4_Click()
mnu4.Checked = Not mnu4.Checked
End Sub

Private Sub mnu6_Click()
mnu6.Checked = Not mnu6.Checked
End Sub

Private Sub mnu9_Click()
mnu9.Checked = Not mnu9.Checked
End Sub
