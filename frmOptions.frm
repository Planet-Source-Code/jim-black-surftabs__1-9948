VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   Caption         =   "Surf Tabs Options"
   ClientHeight    =   4680
   ClientLeft      =   4188
   ClientTop       =   3792
   ClientWidth     =   5700
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   5700
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   492
      Left            =   3576
      TabIndex        =   22
      Top             =   4080
      Width           =   972
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   492
      Left            =   2520
      TabIndex        =   21
      Top             =   4080
      Width           =   972
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   492
      Left            =   1440
      TabIndex        =   20
      Top             =   4080
      Width           =   972
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3372
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   5175
      Begin VB.CheckBox chkMinToSysTray 
         Caption         =   "Minimize to system tray"
         Height          =   252
         Left            =   360
         TabIndex        =   29
         Top             =   1080
         Width           =   3612
      End
      Begin VB.CheckBox chkMultipleInstances 
         Caption         =   "Allow multiple instances of SurfTabs"
         Height          =   252
         Left            =   360
         TabIndex        =   3
         Top             =   600
         Value           =   2  'Grayed
         Width           =   3372
      End
      Begin VB.CheckBox chkDefaultBrowser 
         Caption         =   "Ask to make SurfTabs the default browser"
         Height          =   252
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Value           =   2  'Grayed
         Width           =   3735
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Frame4"
      Height          =   3372
      Left            =   240
      TabIndex        =   17
      Top             =   480
      Width           =   5175
      Begin VB.OptionButton optStartPage 
         Caption         =   "Start SurfTabs on a Blank page"
         Height          =   252
         Index           =   2
         Left            =   240
         TabIndex        =   23
         Top             =   1080
         Width           =   2652
      End
      Begin VB.OptionButton optStartPage 
         Caption         =   "Start SurfTabs with Last Open Sites"
         Height          =   252
         Index           =   1
         Left            =   240
         TabIndex        =   19
         Top             =   720
         Width           =   2892
      End
      Begin VB.OptionButton optStartPage 
         Caption         =   "Start SurfTabs at Home Page"
         Height          =   252
         Index           =   0
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   2532
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3372
      Left            =   240
      TabIndex        =   14
      Top             =   480
      Width           =   5175
      Begin VB.TextBox txtRefreshBrowser 
         Height          =   285
         Left            =   3000
         TabIndex        =   27
         Text            =   "15"
         Top             =   720
         Width           =   375
      End
      Begin VB.CheckBox chkRefreshBrowser 
         Caption         =   "Refresh All Browser Tabs Every"
         Enabled         =   0   'False
         Height          =   255
         Left            =   360
         TabIndex        =   26
         Top             =   720
         Value           =   2  'Grayed
         Width           =   2535
      End
      Begin VB.TextBox txtBrowserTitleLength 
         Height          =   288
         Left            =   2640
         TabIndex        =   16
         Text            =   "35"
         Top             =   360
         Width           =   372
      End
      Begin VB.CheckBox chkBrowserTitleLength 
         Caption         =   "Limit Browser Title Length"
         Height          =   252
         Left            =   360
         TabIndex        =   15
         Top             =   360
         Width           =   2292
      End
      Begin VB.Label Label4 
         Caption         =   "minutes."
         Height          =   255
         Left            =   3480
         TabIndex        =   28
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   3372
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   5175
      Begin VB.CheckBox chkNewTabAddressTyped 
         Caption         =   "Address typed in adress field"
         Height          =   252
         Left            =   240
         TabIndex        =   12
         Top             =   1440
         Width           =   2772
      End
      Begin VB.CheckBox chkNewTabFavorites 
         Caption         =   "Favorites"
         Height          =   252
         Left            =   240
         TabIndex        =   11
         Top             =   480
         Value           =   2  'Grayed
         Width           =   1812
      End
      Begin VB.CheckBox chkNewTabHistory 
         Caption         =   "History"
         Height          =   252
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Value           =   2  'Grayed
         Width           =   3252
      End
      Begin VB.Frame boxNewTabOption 
         Height          =   1332
         Left            =   240
         TabIndex        =   5
         Top             =   1800
         Width           =   3612
         Begin VB.OptionButton optDefaultNewButton 
            Caption         =   "New browser tab at home page"
            Height          =   252
            Index           =   2
            Left            =   360
            TabIndex        =   8
            Tag             =   "1"
            Top             =   960
            Width           =   2652
         End
         Begin VB.OptionButton optDefaultNewButton 
            Caption         =   "New browser tab at current location"
            Height          =   252
            Index           =   1
            Left            =   360
            TabIndex        =   7
            Tag             =   "3"
            Top             =   720
            Width           =   3012
         End
         Begin VB.OptionButton optDefaultNewButton 
            Caption         =   "New blank browser tab"
            Height          =   252
            Index           =   0
            Left            =   360
            TabIndex        =   6
            Tag             =   "2"
            Top             =   480
            Width           =   2172
         End
         Begin VB.Label Label2 
            Caption         =   "Default action for New button"
            Height          =   252
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   2292
         End
      End
      Begin VB.CheckBox chkNewTabSearch 
         Caption         =   "Search"
         Height          =   252
         Left            =   240
         TabIndex        =   13
         Top             =   1200
         Value           =   1  'Checked
         Width           =   3372
      End
      Begin VB.CheckBox chkNewTabHome 
         Caption         =   "Home"
         Height          =   252
         Left            =   240
         TabIndex        =   24
         Top             =   960
         Width           =   2052
      End
      Begin VB.Label Label3 
         Caption         =   "Open a New Tab when going to..."
         Height          =   252
         Left            =   240
         TabIndex        =   25
         Top             =   240
         Width           =   3612
      End
   End
   Begin MSComctlLib.TabStrip tbsOptions 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      _ExtentX        =   9546
      _ExtentY        =   6795
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "   General   "
            Key             =   "General"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "   Browser Tabs   "
            Key             =   "BrowserTabs"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "   New Tabs   "
            Key             =   "NewTabs"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "   Startup   "
            Key             =   "Startup"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'//////////////////////////////////////////////////////////
'//
'// SurfTabs
'// A Tiger Studios product
'// By Jim Black
'// Jim@BlacksWeb.com
'//
'//////////////////////////////////////////////////////////


Private Sub cmdApply_Click()
    'Save all the options and set focus to OK button
    Call SaveOptions
    OptionsSaved = True
    cmdOk.SetFocus

End Sub

Private Sub cmdCancel_Click()
    Unload Me

End Sub

Private Sub cmdOK_Click()
    If Not OptionsSaved Then Call SaveOptions
    
    Unload Me

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'If the ESC key was entered
    If KeyAscii = 27 Then Unload Me

End Sub

Private Sub Form_Load()
    'Initialize form data
    OptionsSaved = False
    
    Frame1.Visible = True
    Frame1.Caption = ""
    Frame2.Visible = False
    Frame2.Caption = ""
    Frame3.Visible = False
    Frame3.Caption = ""
    Frame4.Visible = False
    Frame4.Caption = ""
    tbsOptions.Tabs(1).Selected = True
    
    '///////////////////////////////////////////////////////
    'GENERAL tab
    chkMinToSysTray.Value = gMinToSysTray
    'chkMultipleInstances.Value = gMultipleInstances
    chkMultipleInstances.Enabled = False
    'chkDefaultBrowser.Value = gDefaultBrowser
    chkDefaultBrowser.Enabled = False
    
    '///////////////////////////////////////////////////////
    'BROWSER TABS tab
    chkBrowserTitleLength.Value = gBrowserTitleLength
    txtBrowserTitleLength.text = gtxtBrowserTitleLength
    chkRefreshBrowser.Value = gchkRefreshBrowser
    txtRefreshBrowser.text = gtxtRefreshBrowser
        
    '///////////////////////////////////////////////////////
    'NEW BROWSESR TAB tab
    chkNewTabHome.Value = gNewTabHome
    chkNewTabSearch.Value = gNewTabSearch
    chkNewTabAddressTyped.Value = gNewTabAddressTyped
    'chkNewTabFavorites.Value = gNewTabFavorites
    chkNewTabFavorites.Enabled = False
    'chkNewTabHistory.Value = gNewTabHistory
    chkNewTabHistory.Enabled = False
    optDefaultNewButton(gDefaultNewButton).Value = True
    
    '///////////////////////////////////////////////////////
    'START UP tab
    optStartPage(gStartPage).Value = True
    
End Sub

Private Sub tbsOptions_Click()
    'Show and enable the selected tab's controls
    'And hide and disable all others
    Dim i As Integer
    
    For i = 0 To tbsOptions.Tabs.Count - 1
        
        If i = tbsOptions.SelectedItem.index - 1 Then
            
            Select Case i
                
                Case 0: Frame1.Visible = True
                
                Case 1: Frame2.Visible = True
                
                Case 2: Frame3.Visible = True
                
                Case 3: Frame4.Visible = True
            
            End Select
        
        Else
            
            Select Case i
                
                Case 0: Frame1.Visible = False
                
                Case 1: Frame2.Visible = False
                
                Case 2: Frame3.Visible = False
                
                Case 3: Frame4.Visible = False
            
            End Select
        
        End If
    
    Next

End Sub

Private Sub tbsOptions_KeyPress(KeyAscii As Integer)
    'If the ESC key was entered
    If KeyAscii = 27 Then Unload Me

End Sub
