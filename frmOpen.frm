VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOpen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Open"
   ClientHeight    =   2304
   ClientLeft      =   36
   ClientTop       =   276
   ClientWidth     =   5556
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2304
   ScaleWidth      =   5556
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   5040
      Top             =   240
      _ExtentX        =   699
      _ExtentY        =   699
      _Version        =   393216
      Filter          =   "Web Documents (*.htm, *.html)|*.htm, *.html|Any (*.*)|*.*"
   End
   Begin VB.ComboBox cboAddress 
      Height          =   288
      Left            =   840
      TabIndex        =   0
      Top             =   960
      Width           =   4572
   End
   Begin VB.CheckBox optNewBrowserTab 
      Caption         =   "Open on a new browser tab"
      Height          =   252
      Left            =   840
      TabIndex        =   1
      Top             =   1440
      Width           =   2532
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse"
      Height          =   372
      Left            =   4200
      TabIndex        =   5
      Top             =   1800
      Width           =   1212
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   372
      Left            =   2880
      TabIndex        =   3
      Top             =   1800
      Width           =   1212
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Height          =   372
      Left            =   1560
      TabIndex        =   2
      Top             =   1800
      Width           =   1212
   End
   Begin VB.Label Label2 
      Caption         =   "Open:"
      Height          =   252
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   492
   End
   Begin VB.Label Label1 
      Caption         =   "Type the Internet address of a document or folder,       and SurfTabs will open it for you."
      Height          =   492
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   4092
   End
End
Attribute VB_Name = "frmOpen"
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


Private Sub cboAddress_Change()
    If cboAddress.text <> "" Then cmdOk.Enabled = True

End Sub

Private Sub cboAddress_KeyDown(KeyCode As Integer, Shift As Integer)
    If cboAddress.text <> "" Then cmdOk.Enabled = True

End Sub

Private Sub cboAddress_KeyPress(KeyAscii As Integer)
    If cboAddress.text <> "" Then cmdOk.Enabled = True

End Sub

Private Sub cmdBrowse_Click()
    'Open a common dialog for browsing to a document
    'Then load the path to that document into the address text box
    frmOpen.CommonDialog.Filter = "Web Documents (*.htm, *.html)|*.htm;*.html|Any (*.*)|*.*"
    frmOpen.CommonDialog.Flags = &H1000 Or &H2000000 Or &H800
    frmOpen.CommonDialog.Action = 1
    
    If frmOpen.CommonDialog.FileName <> "" Then
        frmOpen.cboAddress.text = frmOpen.CommonDialog.FileName
    End If
    
    frmOpen.cmdOk.Enabled = True
    frmOpen.cmdOk.SetFocus

End Sub

Private Sub cmdCancel_Click()
    Unload frmOpen

End Sub

Private Sub cmdOK_Click()
    If optNewBrowserTab.Value = 1 Then
        OpenNewTab = True
    End If
    
    OpenURL = cboAddress
    OpenOk = True
    Unload frmOpen

End Sub

Private Sub Form_Load()
    'Initialize form data
    cmdOk.Enabled = False
    optNewBrowserTab.Value = 1
    cboAddress.text = "http://"

End Sub
