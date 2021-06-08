VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Begin VB.Form frmOpenDrawer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Open Drawer ..."
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   4110
   Begin BiSAButtonProject.BiSAButton BiSAButton1 
      Height          =   465
      Left            =   1005
      TabIndex        =   1
      Top             =   660
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   820
      Caption         =   "Open"
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483633
   End
   Begin BiSATextBoxProject.BiSATextBox cKode 
      Height          =   330
      Left            =   150
      TabIndex        =   0
      Top             =   165
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   582
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Verdana"
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmOpenDrawer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BiSAButton1_Click()
  OpenNewDrawer cKode.Text
End Sub

Private Sub Form_Load()
  CenterForm Me
  SetIcon Me.hWnd
End Sub
