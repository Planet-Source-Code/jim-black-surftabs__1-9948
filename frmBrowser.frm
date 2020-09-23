VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmBrowser 
   Caption         =   "SurfTabs"
   ClientHeight    =   7485
   ClientLeft      =   2340
   ClientTop       =   1515
   ClientWidth     =   9765
   Icon            =   "frmBrowser.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7485
   ScaleWidth      =   9765
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   8880
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgTreeFavorites 
      Left            =   1155
      Top             =   5850
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":0894
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":0CE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":1278
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":16CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":1E1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":2136
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":2888
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":2CDA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   7200
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ImageList imlTrayIcons 
      Left            =   9075
      Top             =   6480
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":3128
            Key             =   "earth"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":357A
            Key             =   "monit"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Height          =   855
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   1508
      BandCount       =   2
      _CBWidth        =   7815
      _CBHeight       =   855
      _Version        =   "6.0.8450"
      Child1          =   "tbToolBar"
      MinHeight1      =   450
      Width1          =   2865
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      Caption2        =   "Address"
      Child2          =   "cboAddress"
      MinHeight2      =   315
      Width2          =   3795
      NewRow2         =   -1  'True
      AllowVertical2  =   0   'False
      Begin VB.ComboBox cboAddress 
         Height          =   315
         ItemData        =   "frmBrowser.frx":39CC
         Left            =   795
         List            =   "frmBrowser.frx":39CE
         TabIndex        =   5
         Text            =   "http://"
         Top             =   510
         Width           =   6930
      End
      Begin MSComctlLib.Toolbar tbToolBar 
         Height          =   450
         Left            =   165
         TabIndex        =   4
         Top             =   30
         Width           =   7560
         _ExtentX        =   13335
         _ExtentY        =   794
         ButtonWidth     =   820
         ButtonHeight    =   794
         Style           =   1
         ImageList       =   "imlIcons"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   14
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "NewTab"
               Object.ToolTipText     =   "New Tab"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Back"
               Object.ToolTipText     =   "Back"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Forward"
               Object.ToolTipText     =   "Forward"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Stop"
               Object.ToolTipText     =   "Stop"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Refresh"
               Object.ToolTipText     =   "Refresh"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Home"
               Object.ToolTipText     =   "Home"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Search"
               Object.ToolTipText     =   "Search"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Favorites"
               Object.ToolTipText     =   "Favorites"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "DeleteTab"
               Object.ToolTipText     =   "Delete Current Tab"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "DeleteAllTabs"
               Object.ToolTipText     =   "Delete All Tabs"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Print"
               Object.ToolTipText     =   "Print"
               ImageIndex      =   10
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   7230
      Width           =   9765
      _ExtentX        =   17224
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser brwWebBrowser 
      Height          =   5280
      Index           =   0
      Left            =   2160
      TabIndex        =   1
      Top             =   1200
      Width           =   6090
      ExtentX         =   10731
      ExtentY         =   9313
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   -1  'True
      NoClientEdge    =   -1  'True
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   5895
      Left            =   2040
      TabIndex        =   0
      Top             =   840
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   10398
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Blank"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Timer timTimer 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   8760
      Top             =   1440
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   8145
      Top             =   -30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":39D0
            Key             =   "Back"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":3CB2
            Key             =   "Forward"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":3F94
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":4276
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":4558
            Key             =   "Home"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":483A
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":4B1C
            Key             =   "New"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":5060
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":56E4
            Key             =   "DeleteAll"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":5D68
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":62AA
            Key             =   "Favorites"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":683C
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":6C8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBrowser.frx":70DC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView treeFavorites 
      Height          =   5412
      Left            =   0
      TabIndex        =   8
      Top             =   1320
      Width           =   2052
      _ExtentX        =   3625
      _ExtentY        =   9551
      _Version        =   393217
      Indentation     =   212
      LabelEdit       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      ImageList       =   "imgTreeFavorites"
      Appearance      =   1
      MouseIcon       =   "frmBrowser.frx":73F6
   End
   Begin ComCtl3.CoolBar FavoritesCoolBar 
      Height          =   360
      Left            =   -15
      TabIndex        =   7
      Top             =   840
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   635
      BandCount       =   1
      _CBWidth        =   2055
      _CBHeight       =   360
      _Version        =   "6.0.8450"
      Caption1        =   "Favorites"
      MinHeight1      =   300
      Width1          =   2880
      NewRow1         =   0   'False
      Begin VB.CommandButton cmdCloseFavorites 
         Appearance      =   0  'Flat
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1740
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   30
         Width           =   300
      End
   End
   Begin VB.Menu mnu_File 
      Caption         =   "&File"
      Begin VB.Menu mnu_NewBrowser 
         Caption         =   "&New Browser Tab"
         Begin VB.Menu mnu_NewBrowserCurrent 
            Caption         =   "&Current"
         End
         Begin VB.Menu mnu_NewBrowserHome 
            Caption         =   "&Home"
            Shortcut        =   ^H
         End
         Begin VB.Menu mnu_NewBrowserBlank 
            Caption         =   "&Blank"
         End
      End
      Begin VB.Menu mnu_Open 
         Caption         =   "&Go To Address..."
         Shortcut        =   ^G
      End
      Begin VB.Menu mnu_SaveAs 
         Caption         =   "&Save As..."
      End
      Begin VB.Menu mnu_FileBreak1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_PageSetup 
         Caption         =   "Page Set&up"
      End
      Begin VB.Menu mnu_Print 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnu_FileBreak2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_SendLinkEmail 
         Caption         =   "Send Link by Email"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_SendPageEmail 
         Caption         =   "Send Page by Email"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_FileBreak3 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_Properties 
         Caption         =   "Properties"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_WorkOffline 
         Caption         =   "&Work Offline"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_FileBreak4 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnu_Edit 
      Caption         =   "&Edit"
      Begin VB.Menu mnu_EditCut 
         Caption         =   "C&ut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnu_EditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnu_EditPast 
         Caption         =   "&Past"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnu_EditBreak2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_EditSelectAll 
         Caption         =   "&Select All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnu_EditBreak1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_EditFind 
         Caption         =   "&Find on this page"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu mnu_View 
      Caption         =   "&View"
      Begin VB.Menu mnu_ViewAddressBar 
         Caption         =   "Address Bar"
         Checked         =   -1  'True
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_ViewStatusBar 
         Caption         =   "Status Bar"
         Checked         =   -1  'True
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_ViewToolBar 
         Caption         =   "Tool Bar"
         Checked         =   -1  'True
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_ViewStop 
         Caption         =   "&Stop"
      End
      Begin VB.Menu mnu_ViewRefresh 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu mnu_ViewBreak1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_ViewStopAllTabs 
         Caption         =   "Stop All Tabs"
      End
      Begin VB.Menu mnu_ViewRefreshAllTabs 
         Caption         =   "Refresh All Tabs"
      End
      Begin VB.Menu mnu_ViewBreak2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_ViewSource 
         Caption         =   "&Source"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnu_Go 
      Caption         =   "&Go"
      Begin VB.Menu mnu_GoBack 
         Caption         =   "&Back"
      End
      Begin VB.Menu mnu_GoForward 
         Caption         =   "&Forward"
      End
      Begin VB.Menu mnu_GoBreak1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_GoHome 
         Caption         =   "&Home"
      End
      Begin VB.Menu mnu_GoSearch 
         Caption         =   "&Search"
      End
      Begin VB.Menu mnu_GoBreak2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_GoPreviousTab 
         Caption         =   "&Previous Tab"
      End
      Begin VB.Menu mnu_GoNextTab 
         Caption         =   "&Next Tab"
      End
   End
   Begin VB.Menu mnu_Favorites 
      Caption         =   "F&avorites"
      Begin VB.Menu mnu_AddToFavorites 
         Caption         =   "&Add to Favorites"
      End
      Begin VB.Menu mnu_ViewFavorites 
         Caption         =   "&View Favorites"
      End
      Begin VB.Menu mnu_OrganizeFavorites 
         Caption         =   "&Organize Favorites"
      End
   End
   Begin VB.Menu mnu_Options 
      Caption         =   "&Options"
      Begin VB.Menu mnu_SurfTabOptions 
         Caption         =   "SurfTab Options"
      End
      Begin VB.Menu mnu_InterNetOptions 
         Caption         =   "Internet Options"
      End
   End
   Begin VB.Menu mnu_Tabs 
      Caption         =   "&Tabs"
      Begin VB.Menu mnu_PrevTab 
         Caption         =   "&Previoius Tab"
      End
      Begin VB.Menu mnu_NextTab 
         Caption         =   "&Next Tab"
      End
      Begin VB.Menu mnu_AddTab 
         Caption         =   "&New Tab"
      End
      Begin VB.Menu mnu_DeleteTab 
         Caption         =   "&Delete Current Tab"
      End
      Begin VB.Menu mnu_DeleteAllTabs 
         Caption         =   "Delete All Tabs"
      End
   End
   Begin VB.Menu mnu_Tools 
      Caption         =   "T&ools"
      Begin VB.Menu mnu_EditViewHistory 
         Caption         =   "Edit/View History"
      End
      Begin VB.Menu mnu_RefreshHistory 
         Caption         =   "Refresh History"
      End
   End
   Begin VB.Menu mnu_Help 
      Caption         =   "&Help"
      Begin VB.Menu mnu_About 
         Caption         =   "About"
      End
   End
   Begin VB.Menu mnu_TrayPopup 
      Caption         =   "mnuTrayPopup"
      Visible         =   0   'False
      Begin VB.Menu mnu_SurfTabs 
         Caption         =   "SurfTabs by Tiger Studios"
      End
      Begin VB.Menu mnu_SurfTabsLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_TrayExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'//////////////////////////////////////////////////////////////////
'// Surf Tabs
'// A product of Tiger Studios
'// http://TigerStudios.BlacksWeb.com

'//////////////////////////////////////////////////////////////////
'// MOST, but not ALL global variables begin with g


Public StartingAddress As String
Private WithEvents gSysTray As clsSysTray
Attribute gSysTray.VB_VarHelpID = -1
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const CB_FINDSTRING = &H14C
Private Const CB_ERR = (-1)
Private Const CB_SETCURSEL = &H14E


Private Sub brwWebBrowser_DocumentComplete(index As Integer, ByVal pDisp As Object, URL As Variant)
On Error GoTo DocumentComplete_Error:
    Dim brwIndex As Integer
    Dim brwLocation As String
    brwIndex = index
    brwLocation = brwWebBrowser(index).LocationName
    SetTabCaption brwLocation, brwIndex
    Exit Sub
DocumentComplete_Error:
    ShowErrorMessageBox ("DocumentComplete")
End Sub

Private Sub brwWebBrowser_DownloadComplete(index As Integer)
    '*** Occures to often, not used ***
    'SetTabCaption (brwWebBrowser(index).LocationName)

End Sub

Private Sub brwWebBrowser_NavigateComplete2(index As Integer, ByVal pDisp As Object, URL As Variant)
On Error GoTo NavigateComplete2_Error:
    'Add URL to History Pull Down
    'If URL exists in the list, remove it and
    'Insert it at index 0 to keep the most recent at the top
    Dim CurAddress As String
    CurAddress = brwWebBrowser(TabStrip1.Tabs(CurTab_Index).Tag).LocationURL
    If AddURL = True Or ClickHistory = True Then
        
        Dim i As Integer
        Dim Found As Boolean
        i = 0
        Found = False
        
        While i <= cboAddress.ListCount And Not Found
            If cboAddress.List(i) = Right(CurAddress, Len(CurAddress) - Len("http://")) Then
                Found = True
            End If
            i = i + 1
        Wend
        
        If Not Found Then
            cboAddress.AddItem Right(CurAddress, Len(CurAddress) - Len("http://")), 0
        Else
            'Delete the item and add item as index 0
            cboAddress.RemoveItem i - 1
            cboAddress.text = Right(CurAddress, Len(CurAddress) - Len("http://"))
            cboAddress.AddItem Right(CurAddress, Len(CurAddress) - Len("http://")), 0
        End If
        
        AddURL = False
        ClickHistory = False
    
    End If
    TabStrip1.SetFocus
    brwWebBrowser(TabStrip1.Tabs(CurTab_Index).Tag).SetFocus
    Exit Sub
    
NavigateComplete2_Error:
    ShowErrorMessageBox ("NavigateComplete2")
End Sub

Private Sub brwWebBrowser_NewWindow2(index As Integer, ppDisp As Object, Cancel As Boolean)
On Error GoTo NewWindow2_Error:
    'Should do a New Tab here...
    Dim URL As String
    URL = ""
    Call NewTab(Me, URL, -99)
    Set ppDisp = brwWebBrowser(TabStrip1.Tabs(CurTab_Index).Tag).Object
    Call SelectBrowserTab(CurTab_Index)
    Exit Sub
    
NewWindow2_Error:
    ShowErrorMessageBox ("NewWindow2")
End Sub

Private Sub brwWebBrowser_ProgressChange(index As Integer, ByVal Progress As Long, ByVal ProgressMax As Long)
On Error GoTo ProgressChange_Error:
    'Display Progress Indicator in the status bar
    
    If index = TabStrip1.Tabs(CurTab_Index).Tag And brwWebBrowser(TabStrip1.Tabs(CurTab_Index).Tag).Busy Then
        
        If Progress = -1 Then
            ProgressBar.Value = 0
            ProgressBar.Visible = False
        End If

        If Progress > 0 And ProgressMax > 0 Then
            
            If ((Progress * 100) / ProgressMax) < 100 Then
                ProgressBar.Value = (Progress * 100) / ProgressMax
                ProgressBar.Visible = True
                ProgressBar.ZOrder (0)
                'make sure progress bar is position correctly
                'some times it doesn't reposition it self properly
                'durring resize
                RepositionProgressBar
            Else
                ProgressBar.Visible = False
            End If
        
        End If
    
    End If
    Exit Sub
    
ProgressChange_Error:
    ShowErrorMessageBox ("ProgressChange_Error")
End Sub

Private Sub brwWebBrowser_StatusTextChange(index As Integer, ByVal text As String)
On Error GoTo StatusTextChange_Error:
    'Display in Statusbar, the status of the browser
    'Downlaod status and hyperlink fly-overs
    If index = TabStrip1.Tabs(CurTab_Index).Tag Then
        StatusBar.Panels(1).text = text
    End If
    Exit Sub
    
StatusTextChange_Error:
    ShowErrorMessageBox ("StatusTextChange")
End Sub

Private Sub brwWebBrowser_TitleChange(index As Integer, ByVal text As String)
    'If the title of the doc has changed, make adjustments
'    Dim brwIndex As Integer
'    Dim brwLocation As String
    
'    brwIndex = index
'    brwLocation = brwWebBrowser(brwIndex).LocationName
'    SetTabCaption brwLocation, brwIndex

End Sub

Private Sub cboAddress_Change()
On Error GoTo cboAddress_Change_Error:
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    
    If Not bBackSpace Then
        With cboAddress
            strPartial = .text
            i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial)
            
            If i <> CB_ERR Then
              strTotal = .List(i)
              j = Len(strTotal) - Len(strPartial)
              
              If j <> 0 Then
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
              Else
              End If
              
            End If
            
        End With
    End If
    Exit Sub
    
cboAddress_Change_Error:
    ShowErrorMessageBox ("cboAddress_Change")
End Sub

Private Sub cmdCloseFavorites_Click()
On Error GoTo cmdCloseFavorites_Click_Error:
    'Call the view favorites code
    Call mnu_ViewFavorites_Click
    Exit Sub
    
cmdCloseFavorites_Click_Error:
    ShowErrorMessageBox ("DocumentComplete")
End Sub

Private Sub CoolBar1_Resize()
    Form_Resize

End Sub

Private Sub Form_Load()
    On Error Resume Next
    
'*** Handle command line parameters NOT implamented at this time ***
    'Check if there is a command line parameter
'    If Left$(Command$, 5) = "http""" Or _
'        Left$(Command$, 5) = "file""" Or _
'        Right$(Command$, 4) = "htm""" Or _
'        Right$(Command$, 5) = "html""" Then
'        StartingAddress = Left$(Command$, Len(Command$) - 1)
'        StartingAddress = Right$(StartingAddress, Len(StartingAddress) - 1)
'    Else
'        If Right$(Command$, 4) = "URL""" Then
'            StartingAddress = Left$(Command$, Len(Command$) - 1)
'            StartingAddress = Right$(StartingAddress, Len(StartingAddress) - 1)
'        Else
'            StartingAddress = ""
'        End If
'    End If
'    MsgBox Command$
'    MsgBox StartingAddress
    
    '//////////////////////////////////////////////////////////////
    'Check if there is an instance of the program already running
    If App.PrevInstance = True Then
        MsgBox "SurfTabs is already running, check your task bar and system tray"
        Unload Me
    End If
    
    '//////////////////////////////////////////////////////////////
    'Setup the System Tray icon
    'NOTE: icons must be 16 colors MAX
    Set gSysTray = New clsSysTray
    Set gSysTray.SourceWindow = Me
    gSysTray.Icon = imlTrayIcons.ListImages("earth").Picture
    gSysTray.ChangeToolTip (PROGRAM_NAME)
        
    '//////////////////////////////////////////////////////////////
    'Initialize program data
    CurTab_Index = 1
    MaxTab_Index = 1
    
    TabStrip1.Tabs.Item(CurTab_Index).Tag = 0
    
    StatusBar.Panels(2).Width = 1200
    StatusBar.Panels(3).Width = 1600
    StatusBar.Panels(1).Width = Me.ScaleWidth - (StatusBar.Panels(2).Width + StatusBar.Panels(3).Width)
    
    bBackSpace = False
    bDeleteKey = False
    
    Call SetProgPath
    Call GetFormInfo
    Call GetOptions
    Call GetTypedURLs
    Call GetFavorites
    ViewingFavorites = False
    
    tbToolBar.Refresh
    
    Call Form_Resize
    
    Me.Show
    Me.Refresh

    If StartingAddress = "" Then
        
        Select Case gStartPage
            Case 0 ' Start at Home Page
                brwWebBrowser(TabStrip1.Tabs(CurTab_Index).Tag).GoHome
            Case 1 ' Start with Saved URLs
                    Call GetSavedTabURLs
            Case 2 ' Start on Blank page
                brwWebBrowser(TabStrip1.Tabs(CurTab_Index).Tag).Navigate BLANK_URL
        End Select
    
    Else
        brwWebBrowser(TabStrip1.Tabs(CurTab_Index).Tag).Navigate StartingAddress
    End If
    
    CurTab_Index = 1
    TabStrip1.Tabs(CurTab_Index).Selected = True
    brwWebBrowser(TabStrip1.Tabs(CurTab_Index).Tag).Visible = True
    brwWebBrowser(TabStrip1.Tabs(CurTab_Index).Tag).ZOrder (0)
    
    ReDim Preserve FreedBrowserArray(0)
    FreedBrowserArrayEmpty = True
   
End Sub

Private Sub cboAddress_Click()
On Error GoTo cboAddress_Click_Error:
    'Handle selection of URL from history
    ClickHistory = True
    
    If Not UnLoading Then
        brwWebBrowser(TabStrip1.Tabs.Item(CurTab_Index).Tag).Navigate cboAddress.text
        brwWebBrowser(TabStrip1.Tabs.Item(CurTab_Index).Tag).SetFocus
    End If
    Exit Sub
cboAddress_Click_Error:
    ShowErrorMessageBox ("cboAddress_Click")
End Sub

Private Sub cboAddress_KeyPress(KeyAscii As Integer)
On Error GoTo cboAddress_KeyPress_Error:
    'Handle the entry of a URL into the address text box
    On Error Resume Next
    
    If KeyAscii = vbKeyReturn Then
        AddURL = True
        
        If gNewTabAddressTyped Then
            Call NewTab(Me, cboAddress.text, -1)
        Else
            brwWebBrowser(TabStrip1.Tabs.Item(CurTab_Index).Tag).Navigate cboAddress.text
        End If
    
    End If
    
    If KeyAscii = 8 Then
        bBackSpace = True
    Else
        bBackSpace = False
    End If
    Exit Sub

cboAddress_KeyPress_Error:
    ShowErrorMessageBox ("cboAddress_KeyPress_Error")
End Sub

Private Sub Form_Resize()
On Error GoTo Form_Resize_Error:
    'Handle resizing MOST controls on the form
    
    If Me.WindowState = vbMinimized Then
        If gMinToSysTray Then gSysTray.MinToSysTray
    Else
        Dim FormLeft As Long
        Dim xRatio
        xRatio = (Me.ScaleWidth * 100) \ Me.ScaleWidth
        
        '/////////////////////////////////////////////////////////////////
        'RESIZE THE BUTTON BAR
        lLeft = CLng((CoolBar1.Left * xRatio) \ 100)
        lTop = CoolBar1.Top
        lWidth = Me.ScaleWidth
        CoolBar1.Move lLeft, lTop, lWidth
                
        '/////////////////////////////////////////////////////////////////
        'RESIZE ADDRESS TEXT BOX
        lLeft = CLng((cboAddress.Left * xRatio) \ 100)
        lTop = cboAddress.Top
        lWidth = Me.ScaleWidth - 800
        cboAddress.Move lLeft, lTop, lWidth
        
        '/////////////////////////////////////////////////////////////////
        'RESIZE THE FAVORITES TREE (wont wory about the coolba\toolbar
        If ViewingFavorites Then
            FavoritesCoolBar.Top = Me.ScaleTop + Me.CoolBar1.Height
            lTop = Me.ScaleTop + Me.CoolBar1.Height + FavoritesCoolBar.Height
            lHeight = Me.ScaleHeight - Me.CoolBar1.Height - FavoritesCoolBar.Height - Me.StatusBar.Height
            
            If lHeight > 0 Then
                treeFavorites.Move treeFavorites.Left, lTop, treeFavorites.Width, lHeight
                treeFavorites.Visible = True
            Else
                treeFavorites.Visible = False
            End If
        
        End If
        
        '/////////////////////////////////////////////////////////////////
        'RESIZE THE BROWSER TABS
        If ViewingFavorites Then
            lLeft = Me.ScaleLeft + treeFavorites.Width
        Else
            lLeft = Me.ScaleLeft
        End If
        
        lTop = Me.ScaleTop + Me.CoolBar1.Height
        
        If ViewingFavorites Then
            lWidth = Me.Width - FavoritesCoolBar.Width
        Else
            lWidth = Me.Width
        End If
        
        lHeight = Me.ScaleHeight - Me.CoolBar1.Height - Me.StatusBar.Height
        
        If lWidth > 0 And lHeight > 0 Then
            TabStrip1.Move lLeft, lTop, lWidth, lHeight
            TabStrip1.Visible = True
        Else
            TabStrip1.Visible = False
        End If
        
        '/////////////////////////////////////////////////////////////////
        'RESIZE THE BROWSER WINDOW
        lLeft = Me.TabStrip1.Left + 60
        lTop = Me.TabStrip1.Top + 340
        lWidth = Me.TabStrip1.Width - 180
        lHeight = Me.TabStrip1.Height - 400
        
        If lWidth > 0 And lHeight > 0 Then
            brwWebBrowser(TabStrip1.Tabs.Item(CurTab_Index).Tag).Move lLeft, lTop, lWidth, lHeight
            brwWebBrowser(TabStrip1.Tabs.Item(CurTab_Index).Tag).Visible = True
        Else
            brwWebBrowser(TabStrip1.Tabs.Item(CurTab_Index).Tag).Visible = False
        End If
        
        '/////////////////////////////////////////////////////////////////
        'RESIZE THE BROWSER STATUS BAR
        lWidth = Me.ScaleWidth - (StatusBar.Panels(2).Width + StatusBar.Panels(3).Width)
        
        If lWidth > 0 Then
            StatusBar.Panels(1).Width = lWidth
            StatusBar.Panels(1).Visible = True
        Else
            StatusBar.Panels(1).Visible = False
        End If
        
        '/////////////////////////////////////////////////////////////////
        'RESIZE and POSITION THE PROGRESS BAR
        RepositionProgressBar
    
    End If
    Exit Sub
Form_Resize_Error:
    ShowErrorMessageBox ("Form_Resize")
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Form_Unload_Error:
    'Save Form info, Current open URL's and History
    UnLoading = True
    
    gSysTray.RemoveFromSysTray
    Call SaveFormInfo
    Call SaveCurrentTabURLs
    Call SaveTypedURLs
    Exit Sub

Form_Unload_Error:
    ShowErrorMessageBox ("Form_Unload")
End Sub

Private Sub mnu_About_Click()
On Error GoTo mnu_About_Click_Error:
    frmAbout.Show 1
    If cboAddress.text = "http://TigerStudios.BlacksWeb.com" Then
        cboAddress_KeyPress (vbKeyReturn)
    End If
    Exit Sub

mnu_About_Click_Error:
    ShowErrorMessageBox ("mnu_About_Click")
End Sub

Private Sub mnu_AddTab_Click()
On Error GoTo mnu_AddTab_Click_Error:
    Call NewTab(Me, "", -1)
    Exit Sub

mnu_AddTab_Click_Error:
    ShowErrorMessageBox ("mnu_AddTab_Click")
End Sub

Private Sub mnu_AddToFavorites_Click()
On Error GoTo mnu_AddToFavorites_Click_Error:
    Call AddToFavorites
    Exit Sub

mnu_AddToFavorites_Click_Error:
    ShowErrorMessageBox ("mnu_AddToFavorites_Click")
End Sub

Private Sub mnu_DeleteAllTabs_Click()
On Error GoTo mnu_DeleteAllTabs_Click_Error:
    Call DeleteAllTabs
    Exit Sub

mnu_DeleteAllTabs_Click_Error:
    ShowErrorMessageBox ("mnu_DeleteAllTabs_Click")
End Sub

Private Sub mnu_DeleteTab_Click()
On Error GoTo mnu_DeleteTab_Click_Error:
    Call DeleteTab
    Exit Sub

mnu_DeleteTab_Click_Error:
    ShowErrorMessageBox ("mnu_DeleteTab_Click")
End Sub

Private Sub mnu_EditCopy_Click()
On Error GoTo mnu_EditCopy_Click_Error:
    brwWebBrowser(TabStrip1.Tabs(CurTab_Index).Tag).ExecWB OLECMDID_COPY, OLECMDEXECOPT_DODEFAULT
    Exit Sub

mnu_EditCopy_Click_Error:
    ShowErrorMessageBox ("mnu_EditCopy_Click")
End Sub

Private Sub mnu_EditCut_Click()
On Error GoTo mnu_EditCut_Click_Error:
    brwWebBrowser(TabStrip1.Tabs(CurTab_Index).Tag).ExecWB OLECMDID_CUT, OLECMDEXECOPT_DODEFAULT
    Exit Sub

mnu_EditCut_Click_Error:
    ShowErrorMessageBox ("mnu_EditCut_Click")
End Sub

Private Sub mnu_EditFind_Click()
On Error GoTo mnu_EditFind_Click_Error:
    SetFocusOnly = True
    TabStrip1.SetFocus
    brwWebBrowser(TabStrip1.Tabs(CurTab_Index).Tag).SetFocus
    SendKeys "^f"
    Exit Sub

mnu_EditFind_Click_Error:
    ShowErrorMessageBox ("mnu_EditFind_Click")
End Sub

Private Sub mnu_EditPast_Click()
On Error GoTo mnu_EditPast_Click_Error:
    brwWebBrowser(TabStrip1.Tabs(CurTab_Index).Tag).ExecWB OLECMDID_PASTE, OLECMDEXECOPT_DODEFAULT
    Exit Sub

mnu_EditPast_Click_Error:
    ShowErrorMessageBox ("mnu_EditPast_Click")
End Sub

Private Sub mnu_EditSelectAll_Click()
On Error GoTo mnu_EditSelectAll_Click_Error:
    brwWebBrowser(TabStrip1.Tabs(CurTab_Index).Tag).ExecWB OLECMDID_SELECTALL, OLECMDEXECOPT_DODEFAULT
    Exit Sub

mnu_EditSelectAll_Click_Error:
    ShowErrorMessageBox ("mnu_EditSelectAll_Click")
End Sub

Private Sub mnu_EditViewHistory_Click()
On Error GoTo NotInPath
    '//////////////////////////////////////////////
    'Save current URL history to file
    'Then load it into Notepad for Edit and Viewing
    Dim X As Integer
    
    UnLoading = True
    Open gProgPath & "History.dat" For Output As #1
    
    For X = 0 To frmBrowser.cboAddress.ListCount - 1
        frmBrowser.cboAddress.ListIndex = X
        Print #1, frmBrowser.cboAddress.text
    Next
    
    Close #1
    
    'Then load it into Notepad for Edit and Viewing
    X = Shell("Notepad.exe " & gProgPath & "History.dat", vbMaximizedFocus)
            
    HistoryFileChanged = True
    Exit Sub
    
NotInPath:
    MsgBox "ERROR, Notepad.exe was not found in your system path."
End Sub

Private Sub mnu_Exit_Click()
On Error GoTo mnu_Exit_Click_Error:
    Unload Me
    Exit Sub
    
mnu_Exit_Click_Error:
    ShowErrorMessageBox ("mnu_Exit_Click")
End Sub

Private Sub mnu_GoBack_Click()
On Error GoTo mnu_GoBack_Click_Error:
    brwWebBrowser(TabStrip1.Tabs(CurTab_Index).Tag).GoBack
    Exit Sub
    
mnu_GoBack_Click_Error:
    ShowErrorMessageBox ("mnu_GoBack_Click")
End Sub

Private Sub mnu_GoForward_Click()
On Error GoTo mnu_GoForward_Click_Error:
    brwWebBrowser(TabStrip1.Tabs(CurTab_Index).Tag).GoForward
    Exit Sub
    
mnu_GoForward_Click_Error:
    ShowErrorMessageBox ("mnu_GoForward_Click")
End Sub

Private Sub mnu_GoHome_Click()
On Error GoTo mnu_GoHome_Click_Error:
    brwWebBrowser(TabStrip1.Tabs(CurTab_Index).Tag).GoHome
    Exit Sub
    
mnu_GoHome_Click_Error:
    ShowErrorMessageBox ("mnu_GoHome_Click")
End Sub

Private Sub mnu_GoNextTab_Click()
On Error GoTo mnu_GoNextTab_Click_Error:
    'Move the current browser off the form\tab
    'Move the next browser onto the form\tab
    Call MoveBrowserOffFormTab(TabStrip1.Tabs(CurTab_Index).Tag)
    
    If CurTab_Index = MaxTab_Index Then
        CurTab_Index = 1
    Else
        CurTab_Index = CurTab_Index + 1
    End If
    
    Call SelectBrowserTab(CurTab_Index)
    Exit Sub
    
mnu_GoNextTab_Click_Error:
    ShowErrorMessageBox ("mnu_GoNextTab_Click")
End Sub

Private Sub mnu_GoPreviousTab_Click()
On Error GoTo mnu_GoPreviousTab_Click_Error:
    'Move the current browser off the form\tab
    'Move the next browser onto the form\tab
    Call MoveBrowserOffFormTab(TabStrip1.Tabs(CurTab_Index).Tag)
    
    If CurTab_Index = 1 Then
        CurTab_Index = MaxTab_Index
    Else
        CurTab_Index = CurTab_Index - 1
    End If
    
    Call SelectBrowserTab(CurTab_Index)
    Exit Sub
    
mnu_GoPreviousTab_Click_Error:
    ShowErrorMessageBox ("mnu_GoPreviousTab_Click")
End Sub

Private Sub mnu_GoSearch_Click()
On Error GoTo mnu_GoSearch_Click_Error:
    brwWebBrowser(TabStrip1.Tabs(CurTab_Index).Tag).GoSearch
    Exit Sub
    
mnu_GoSearch_Click_Error:
    ShowErrorMessageBox ("mnu_GoSearch_Click")
End Sub

Private Sub mnu_InterNetOptions_Click()
    Dim RetVal
    RetVal = Shell("rundll32.exe shell32.dll,Control_RunDLL Inetcpl.cpl", vbNormalFocus)

End Sub

Private Sub mnu_NewBrowserBlank_Click()
    Call NewTab(frmBrowser, "", NEW_TAB_BLANK)

End Sub

Private Sub mnu_NewBrowserCurrent_Click()
    Call NewTab(frmBrowser, brwWebBrowser(TabStrip1.Tabs(CurTab_Index).Tag).LocationURL, NEW_TAB_CUR_URL)

End Sub

Private Sub mnu_NewBrowserHome_Click()
    Call NewTab(frmBrowser, "", NEW_TAB_HOME)

End Sub

Private Sub mnu_NextTab_Click()
    Call mnu_GoNextTab_Click

End Sub

Private Sub mnu_Open_Click()
    frmOpen.Show 1
    
    If OpenOk = True Then
        If OpenNewTab Then
            Call NewTab(Me, OpenURL, -1)
        Else
            brwWebBrowser(TabStrip1.Tabs(CurTab_Index).Tag).Navigate OpenURL
        End If
    End If

End Sub

Private Sub mnu_OrganizeFavorites_Click()
On Error GoTo mnu_OrganizeFavorites_Click_Error:
    Dim lpszRootFolder As String
    Dim success As Long
    Dim CSIDL As Long

    'open the organize folder at the path specified by the CSIDL
    CSIDL = CSIDL_FAVORITES
     
    lpszRootFolder = GetFolderPath(CSIDL)
    success = DoOrganizeFavDlg(hwnd, lpszRootFolder)
            
    Call GetFavorites  'To refresh the favorites tree
    Exit Sub
mnu_OrganizeFavorites_Click_Error:
    ShowErrorMessageBox ("mnu_OrganizeFavorites_Click")
End Sub

Private Sub mnu_PageSetup_Click()
    brwWebBrowser(TabStrip1.Tabs(CurTab_Index).Tag).ExecWB OLECMDID_PAGESETUP, OLECMDEXECOPT_DODEFAULT

End Sub

Private Sub mnu_PrevTab_Click()
    Call mnu_GoPreviousTab_Click

End Sub

Private Sub mnu_Print_Click()
    Call PrintBrowser

End Sub

Private Sub mnu_RefreshHistory_Click()
On Error GoTo mnu_RefreshHistory_Click_Error:
    If HistoryFileChanged Then
        UnLoading = True
        
        'Delete all current 'Typed' URL's in the registry
        Call DeleteKey(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\TypedURLs")
        
        'Load history from history file into Address list
        Get_History
        
        cboAddress.text = brwWebBrowser(TabStrip1.Tabs(CurTab_Index).Tag).LocationURL
        
        UnLoading = False
        
        TabStrip1.SetFocus
        brwWebBrowser(TabStrip1.Tabs(CurTab_Index).Tag).SetFocus
        
        HistoryFileChanged = False
    Else
        MsgBox "History file has not been changed." & vbCrLf _
               & "Use Edit/View History first, then Refresh History."
    End If
    Exit Sub

mnu_RefreshHistory_Click_Error:
    ShowErrorMessageBox ("mnu_RefreshHistory_Click")
End Sub

Private Sub mnu_SaveAs_Click()
    brwWebBrowser(TabStrip1.Tabs(CurTab_Index).Tag).ExecWB OLECMDID_SAVEAS, OLECMDEXECOPT_DODEFAULT

End Sub

Private Sub mnu_SelectAll_Click()
    brwWebBrowser(TabStrip1.Tabs(CurTab_Index).Tag).ExecWB OLECMDID_SELECTALL, OLECMDEXECOPT_DODEFAULT

End Sub

Private Sub mnu_SurfTabOptions_Click()
    frmOptions.Show 1

End Sub

Private Sub mnu_TrayExit_Click()
On Error GoTo mnu_TrayExit_Click_Error:
    Unload Me
    Exit Sub

mnu_TrayExit_Click_Error:
    ShowErrorMessageBox ("mnu_TrayExit_Click")
End Sub
Private Sub mnu_SurfTabs_Click()
    gSysTray.LButtonDown

End Sub

Private Sub mnu_ViewFavorites_Click()
    If ViewingFavorites Then
        FavoritesCoolBar.Visible = False
        FavoritesCoolBar.Enabled = False
        treeFavorites.Visible = False
        treeFavorites.Enabled = False
        mnu_ViewFavorites.Checked = False
        ViewingFavorites = False
    Else
        FavoritesCoolBar.Visible = True
        FavoritesCoolBar.Enabled = True
        treeFavorites.Visible = True
        treeFavorites.Enabled = True
        mnu_ViewFavorites.Checked = True
        ViewingFavorites = True
    End If
    
    Form_Resize
    
End Sub

Private Sub mnu_ViewRefresh_Click()
    brwWebBrowser(TabStrip1.Tabs(CurTab_Index).Tag).Refresh

End Sub

Private Sub mnu_ViewRefreshAllTabs_Click()
    Dim X
    For X = 1 To TabStrip1.Tabs.Count
        brwWebBrowser(TabStrip1.Tabs(X).Tag).Refresh
    Next

End Sub

Private Sub mnu_ViewSource_Click()
On Error GoTo mnu_ViewSource_Click_Error:
    '////////////////////////////////////////////////////////////////
    '// ViewSource is NOT being used
    '// For some reason it doesn't get all the source
    '////////////////////////////////////////////////////////////////
    Inet1.AccessType = icUseDefault
    frmSource.txtSource = Inet1.OpenURL(cboAddress.text)
    
    frmSource.Show
    Exit Sub

mnu_ViewSource_Click_Error:
     ShowErrorMessageBox ("mnu_ViewSource_Click")
End Sub

Private Sub mnu_ViewStop_Click()
    brwWebBrowser(TabStrip1.Tabs(CurTab_Index).Tag).Stop

End Sub

Private Sub mnu_ViewStopAllTabs_Click()
    Dim X
    For X = 1 To TabStrip1.Tabs.Count
        brwWebBrowser(TabStrip1.Tabs(X).Tag).Stop
    Next

End Sub

Private Sub TabStrip1_GotFocus()
On Error GoTo TabStrip1_GotFocus_Error:
    Dim NewTab, X, i As Integer
    Dim Found As Boolean
    
    If Not SetFocusOnly And Not InTabSetFocus Then
        
        InTabSetFocus = True
        X = 0
        i = TabStrip1.Tabs.Count
        
        While X < i
            X = X + 1
            
            If TabStrip1.Tabs.Item(X).Selected = True Then
                NewTab = X
            Else
                'Make all browsers except NewBrowserTab not viewable
                'moving them far off the form\tab
                Call MoveBrowserOffFormTab(TabStrip1.Tabs(X).Tag)
            End If
        
        Wend
        
        'Set current Tab and WebBrowser indexes
        CurTab_Index = NewTab
        
        'Enable the new current browser
        brwWebBrowser(TabStrip1.Tabs(CurTab_Index).Tag).Visible = True
        brwWebBrowser(TabStrip1.Tabs(CurTab_Index).Tag).ZOrder (0)
        
        'Clear Progress Panel so it gets reset.
        StatusBar.Panels(2).text = ""
        ProgressBar.Value = 0
        brwWebBrowser(TabStrip1.Tabs.Item(CurTab_Index).Tag).SetFocus
        
        InTabSetFocus = False
        
        Form_Resize
    
    End If
    
    SetFocusOnly = False
    Call SetTabCaption(brwWebBrowser(TabStrip1.Tabs.Item(CurTab_Index).Tag).LocationName, TabStrip1.Tabs.Item(CurTab_Index).Tag)
    Exit Sub

TabStrip1_GotFocus_Error:
    ShowErrorMessageBox ("TabStrip1_GotFocus")
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As Button)
    On Error Resume Next
    Dim X As Integer
    
    Select Case Button.Key
        Case "Back"
            brwWebBrowser(TabStrip1.Tabs(CurTab_Index).Tag).GoBack
        
        Case "Forward"
            brwWebBrowser(TabStrip1.Tabs(CurTab_Index).Tag).GoForward
        
        Case "Refresh"
            brwWebBrowser(TabStrip1.Tabs(CurTab_Index).Tag).Refresh
        
        Case "Home"
            If gNewTabHome Then
                Call NewTab(Me, "", NEW_TAB_HOME)
            Else
                brwWebBrowser(TabStrip1.Tabs(CurTab_Index).Tag).GoHome
            End If
        
        Case "Search"
            If gNewTabSearch Then
                Call NewTab(Me, "", NEW_TAB_SEARCH)
            Else
                brwWebBrowser(TabStrip1.Tabs(CurTab_Index).Tag).GoSearch
            End If
        
        Case "Stop"
            StatusBar.Panels(2).text = ""
            brwWebBrowser(TabStrip1.Tabs(CurTab_Index).Tag).Stop
            Me.Caption = brwWebBrowser(TabStrip1.Tabs(CurTab_Index).Tag).LocationName & " - " & PROGRAM_NAME
            gSysTray.ChangeToolTip (Me.Caption)
            
        Case "NewTab"
            Call NewTab(Me, "", -1)
        
        Case "DeleteTab"
            Call DeleteTab
        
        Case "DeleteAllTabs"
            Call DeleteAllTabs
        
        Case "Print"
            Call PrintBrowser
        
        Case "Favorites"
            Call mnu_ViewFavorites_Click
    
    End Select

End Sub

Private Sub gSysTray_RButtonUP()
    PopupMenu Me.mnu_TrayPopup

End Sub

Public Sub NewTab(b As Object, URL As String, intOption As Integer)
On Error GoTo NewTab_Error:
    Dim X
    
    'Disable current browser
    Call MoveBrowserOffFormTab(TabStrip1.Tabs(CurTab_Index).Tag)
    
    'Increment Tab indexes
    MaxTab_Index = MaxTab_Index + 1
    CurTab_Index = MaxTab_Index
    
    'Add new tab and set properties
    TabStrip1.Tabs.Add
    TabStrip1.Tabs.Item(CurTab_Index).Caption = "Blank"
    TabStrip1.Tabs.Item(CurTab_Index).Selected = True
    
    'Check FreedBrowserArray for an index number
    'If BrowserArrayEmpty Then NOTE: Never using index 0
    If LBound(FreedBrowserArray) = UBound(FreedBrowserArray) Then
        'Array is empty, Increment WebBrowser indexes
        TabStrip1.Tabs(CurTab_Index).Tag = MaxTab_Index - 1
    Else
        'Array is NOT empty, get first index
        TabStrip1.Tabs(CurTab_Index).Tag = FreedBrowserArray(1)
        
        'Adjust the indexes
        For X = LBound(FreedBrowserArray) + 1 To (UBound(FreedBrowserArray) - 1)
            FreedBrowserArray(X) = FreedBrowserArray(X + 1)
        Next
        
        ReDim Preserve FreedBrowserArray(X - 1)
    
    End If
    
    'Load new browser, and enable it
    Load brwWebBrowser(TabStrip1.Tabs(CurTab_Index).Tag)
    brwWebBrowser(TabStrip1.Tabs(CurTab_Index).Tag).Visible = True
    
    '3 Options (Blank, Home or Current)
    If URL = "" Then
        
        If intOption = -99 Then Exit Sub
        
        If intOption = -1 Then
            
            Select Case Int(frmOptions.optDefaultNewButton(gDefaultNewButton).Tag)
                
                Case 1 'NEW_TAB_HOME
                    brwWebBrowser(TabStrip1.Tabs(CurTab_Index).Tag).GoHome
                
                Case 2 'NEW_TAB_BLANK
                    brwWebBrowser(TabStrip1.Tabs(CurTab_Index).Tag).Navigate BLANK_URL
                
                Case 3 'NEW_TAB_CUR_URL
                    brwWebBrowser(TabStrip1.Tabs(CurTab_Index).Tag).GoHome
            
            End Select
        
        Else
            
            Select Case intOption
                
                Case 1 'NEW_TAB_HOME
                    brwWebBrowser(TabStrip1.Tabs(CurTab_Index).Tag).GoHome
                
                Case 2 'NEW_TAB_BLANK
                    brwWebBrowser(TabStrip1.Tabs(CurTab_Index).Tag).Navigate BLANK_URL
                
                Case 3 'NEW_TAB_CUR_URL
                    brwWebBrowser(TabStrip1.Tabs(CurTab_Index).Tag).Navigate brwWebBrowser(TabStrip1.Tabs(CurTab_Index).Tag).LocationURL
                
                Case 5 'NEW_TAB_SEARCH
                    brwWebBrowser(TabStrip1.Tabs(CurTab_Index).Tag).GoSearch
            
            End Select
        
        End If
    
    Else
        
        brwWebBrowser(TabStrip1.Tabs(CurTab_Index).Tag).Navigate URL
    
    End If
    
    Call SelectBrowserTab(CurTab_Index)
    Exit Sub

NewTab_Error:
    ShowErrorMessageBox ("NewTab")
End Sub

Private Sub DeleteTab()
On Error GoTo DeleteTab_Error:
    Dim X
    
    'Stop the current tab\browser before deleteing it
    brwWebBrowser(TabStrip1.Tabs(CurTab_Index).Tag).Stop
    
    If TabStrip1.Tabs.Count > 1 Then
        
        Deleting = True
        
        'Add the current web index to the FreedBrowserArray
        'NOTE never using index 0
        If LBound(FreedBrowserArray) <> UBound(FreedBrowserArray) Then
            X = UBound(FreedBrowserArray)
        Else
            X = 0
        End If
        
        ReDim Preserve FreedBrowserArray(X + 1)
        
        'Delete the tab
        If TabStrip1.Tabs(CurTab_Index).Tag = 0 Then
            brwWebBrowser(TabStrip1.Tabs(CurTab_Index).Tag).Navigate brwWebBrowser(TabStrip1.Tabs(CurTab_Index + 1).Tag).LocationURL
            Unload brwWebBrowser(TabStrip1.Tabs(CurTab_Index + 1).Tag)
            FreedBrowserArray(X + 1) = TabStrip1.Tabs(CurTab_Index + 1).Tag
            TabStrip1.Tabs.Remove (CurTab_Index + 1)
        Else
            Unload brwWebBrowser(TabStrip1.Tabs(CurTab_Index).Tag)
            FreedBrowserArray(X + 1) = TabStrip1.Tabs(CurTab_Index).Tag
            TabStrip1.Tabs.Remove (CurTab_Index)
        End If
        
        FreedBrowserArrayEmpty = False
        
        'Decrement Tab indexes
        If CurTab_Index = MaxTab_Index Then
            MaxTab_Index = MaxTab_Index - 1
            CurTab_Index = MaxTab_Index
        Else
            MaxTab_Index = MaxTab_Index - 1
        End If
        
        'Set the new current tab
        TabStrip1.Tabs.Item(CurTab_Index).Selected = True
        
        Deleting = False
    
    End If
    
    Call SelectBrowserTab(CurTab_Index)
    Exit Sub

DeleteTab_Error:
    ShowErrorMessageBox ("DeleteTab")
End Sub

Sub SetTabCaption(LocationName As String, brwIndex As Integer)
On Error GoTo SetTabCaption_Error:
    Dim theTab As Integer
    
    theTab = 1
    Found = False
    
    While Not Found
        
        'Find the correct tab with a tag = brwIndex
        If TabStrip1.Tabs(theTab).Tag = brwIndex Then
            Found = True
        Else
            theTab = theTab + 1
        End If
    
    Wend
    
    If Left(brwWebBrowser(TabStrip1.Tabs(theTab).Tag).LocationURL, 6) = "about:" Then
        cboAddress.text = ""
        Me.Caption = PROGRAM_NAME
    Else
        cboAddress.text = brwWebBrowser(TabStrip1.Tabs(theTab).Tag).LocationURL
        
        If gtxtBrowserTitleLength <> "" Then
            
            If gBrowserTitleLength = 1 And Int(gtxtBrowserTitleLength) <> 0 And Len(LocationName) > Int(gtxtBrowserTitleLength) Then
                TabStrip1.Tabs(theTab).Caption = Left(LocationName, Int(gtxtBrowserTitleLength))
                Me.Caption = Left(LocationName, Int(gtxtBrowserTitleLength)) & "... - " & PROGRAM_NAME
            Else
                TabStrip1.Tabs(theTab).Caption = LocationName
                Me.Caption = LocationName & " - " & PROGRAM_NAME
            End If
        
        Else
            
            If Me.Caption = "SurfTabs" Then
                TabStrip1.Tabs(theTab).Caption = "Blank"
            Else
                TabStrip1.Tabs(theTab).Caption = LocationName
            End If
        
        End If
    
    End If
    
    gSysTray.ChangeToolTip (Me.Caption)
    Exit Sub
    
SetTabCaption_Error:
    ShowErrorMessageBox ("SetTabCaption")
End Sub

Sub DeleteAllTabs()
On Error GoTo DeleteAllTabs_Error:
    Dim TabCount, X As Integer
    
    brwWebBrowser(TabStrip1.Tabs(CurTab_Index).Tag).Stop
    
    If TabStrip1.Tabs.Count > 1 Then
        
        'Delete tabs and browsers 2 through N
        X = 1
        TabCount = TabStrip1.Tabs.Count
        
        While TabCount >= X
            TabStrip1.Tabs(TabCount).Selected = True
            Call DeleteTab
            TabCount = TabCount - 1
        Wend
        
        brwWebBrowser(TabStrip1.Tabs(CurTab_Index).Tag).Refresh
    
    End If
    
    TabStrip1.SetFocus
    Exit Sub

DeleteAllTabs_Error:
    ShowErrorMessageBox ("DeleteAllTabs")
End Sub

Private Sub treeFavorites_NodeClick(ByVal Node As MSComctlLib.Node)
On Error GoTo treeFavorites_NodeClick_Error:
    'Navigate current Tab\Browser to the selected URL
    If Right(Node.Key, 4) = "_URL" Then
        Set Itm = Node
        brwWebBrowser(TabStrip1.Tabs(CurTab_Index).Tag).Navigate Itm.Tag
        
    End If
    Exit Sub
    
treeFavorites_NodeClick_Error:
    ShowErrorMessageBox ("treeFavorites_NodeClick")
End Sub
