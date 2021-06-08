VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Begin VB.Form frmSetupPrinterKasir 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Setup Printer Kasir"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   5220
   Begin BiSAFramProject.BiSAFrame BiSAFrame2 
      Height          =   615
      Left            =   30
      Top             =   1695
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   1085
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483633
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   3960
         TabIndex        =   1
         Top             =   75
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   767
         Caption         =   "    &Save"
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         BackColor       =   -2147483633
         Picture         =   "frmSetupPrinterKasir.frx":0000
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1605
      Left            =   30
      Top             =   105
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   2831
      Caption         =   "Port Printer"
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483633
      Begin BiSATextBoxProject.BiSATextBox cPortPrinter 
         Height          =   330
         Left            =   255
         TabIndex        =   0
         Top             =   360
         Width           =   3060
         _ExtentX        =   5398
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
         Appearance      =   0
         GetPicture      =   1
         Caption         =   "Port Printer"
         CaptionWidth    =   1500
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
      Begin VB.Label Label2 
         Caption         =   "(Contoh)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3615
         TabIndex        =   3
         Top             =   735
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   "COM1: LPT1: COM2: LPT2:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1905
         TabIndex        =   2
         Top             =   735
         Width           =   1740
      End
   End
End
Attribute VB_Name = "frmSetupPrinterKasir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objData As New CodeSuiteLibrary.Data

Private Sub cmdSimpan_Click()
  UpdCfg msPortPrinter, cPortPrinter.Text, objData
  MsgBox "Data sudah disimpan" & vbCrLf & "Untuk melihat perubahan yg telah dilakukan, silahkan login kembali atau restart Aplikasi ini."
End Sub

Private Sub Form_Load()
  SetIcon Me.hWnd
  cPortPrinter.Text = IIf(aCfg(objData, msPortPrinter) = "", "LPT1:", aCfg(objData, msPortPrinter))
  CenterForm Me
End Sub
