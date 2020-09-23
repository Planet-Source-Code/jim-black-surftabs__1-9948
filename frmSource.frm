VERSION 5.00
Begin VB.Form frmSource 
   Caption         =   "Source"
   ClientHeight    =   5760
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   6588
   Icon            =   "frmSource.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5760
   ScaleWidth      =   6588
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtSource 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   0
      MaxLength       =   10000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   6615
   End
End
Attribute VB_Name = "frmSource"
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

Private Sub Form_Load()
    On Error GoTo ErrHndlr
    
    frmSource.Caption = "Source - " & frmBrowser.cboAddress.text
    
    Exit Sub
ErrHndlr:
 Exit Sub

End Sub

Private Sub Form_Resize()
    Dim xRatio
    Dim lLeft, lTop, lWidth, lHeight As Integer
    
    xRatio = (Me.ScaleWidth * 100) \ Me.ScaleWidth
    
    '/////////////////////////////////////////////////////////////////
    'RESIZE TEXT BOX txtSource
    lLeft = CLng((txtSource.Left * xRatio) \ 100)
    lTop = txtSource.Top
    lWidth = Me.ScaleWidth
    lHeight = Me.ScaleHeight
    txtSource.Move lLeft, lTop, lWidth, lHeight

End Sub
