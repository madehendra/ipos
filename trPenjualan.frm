VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form trPenjualan 
   BorderStyle     =   0  'None
   Caption         =   "Penjualan"
   ClientHeight    =   10545
   ClientLeft      =   0
   ClientTop       =   -45
   ClientWidth     =   19230
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10545
   ScaleWidth      =   19230
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   10335
      Left            =   -60
      TabIndex        =   0
      Top             =   45
      Width           =   19215
      Begin SizerOneLibCtl.TabOne TabOne2 
         Height          =   2700
         Left            =   5910
         TabIndex        =   15
         Top             =   1440
         Width           =   13170
         _cx             =   23230
         _cy             =   4762
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   2
         MousePointer    =   0
         _ConvInfo       =   1
         Version         =   700
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FrontTabColor   =   -2147483633
         BackTabColor    =   -2147483633
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   -2147483630
         Caption         =   "TOTAL|SALES && DISKON"
         Align           =   0
         CurrTab         =   1
         FirstTab        =   0
         Style           =   3
         Position        =   0
         AutoSwitch      =   -1  'True
         AutoScroll      =   -1  'True
         TabPreview      =   -1  'True
         ShowFocusRect   =   -1  'True
         TabsPerPage     =   0
         BorderWidth     =   0
         BoldCurrent     =   -1  'True
         DogEars         =   -1  'True
         MultiRow        =   0   'False
         MultiRowOffset  =   200
         CaptionStyle    =   0
         TabHeight       =   0
         TabCaptionPos   =   4
         TabPicturePos   =   0
         CaptionEmpty    =   ""
         Separators      =   0   'False
         Begin VB.Frame Frame3 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2325
            Left            =   45
            TabIndex        =   21
            Top             =   330
            Width           =   13080
            Begin BiSADateProject.BiSADate dJthTmp 
               Height          =   330
               Left            =   90
               TabIndex        =   24
               TabStop         =   0   'False
               Top             =   255
               Width           =   2565
               _ExtentX        =   4524
               _ExtentY        =   582
               Value           =   "03-04-2011"
               Appearance      =   0
               BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   -2147483640
               Caption         =   "Due Date"
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
            Begin BiSANumberBoxProject.BiSANumberBox nPPn 
               Height          =   336
               Left            =   96
               TabIndex        =   28
               TabStop         =   0   'False
               Top             =   996
               Width           =   2556
               _ExtentX        =   4524
               _ExtentY        =   609
               Appearance      =   0
               BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BackColor       =   -2147483634
               Caption         =   "PPn %"
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
            Begin BiSANumberBoxProject.BiSANumberBox nPersDisc 
               Height          =   336
               Left            =   96
               TabIndex        =   26
               TabStop         =   0   'False
               Top             =   624
               Width           =   2556
               _ExtentX        =   4524
               _ExtentY        =   582
               Appearance      =   0
               Enabled         =   0   'False
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
               Caption         =   "Disc %"
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
            Begin BiSANumberBoxProject.BiSANumberBox nKomisi 
               Height          =   336
               Left            =   3420
               TabIndex        =   27
               TabStop         =   0   'False
               Top             =   984
               Width           =   3120
               _ExtentX        =   5503
               _ExtentY        =   582
               Appearance      =   0
               Decimals        =   0
               BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "Komisi Sales"
               CaptionWidth    =   1300
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
            Begin BiSATextBoxProject.BiSATextBox cUp 
               Height          =   330
               Left            =   5730
               TabIndex        =   22
               TabStop         =   0   'False
               Top             =   -2220
               Width           =   6945
               _ExtentX        =   12250
               _ExtentY        =   582
               Text            =   "12345678901234567890"
               BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontBold        =   -1  'True
               FontName        =   "Verdana"
               BackColor       =   16777215
               MaxLength       =   200
               Appearance      =   0
               Caption         =   "UP (Untuk)"
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
            Begin BiSATextBoxProject.BiSABrowse cSalesman 
               Height          =   330
               Left            =   3420
               TabIndex        =   23
               TabStop         =   0   'False
               Top             =   240
               Width           =   3105
               _ExtentX        =   5477
               _ExtentY        =   582
               Text            =   "12345678"
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
               Button          =   -1  'True
               Caption         =   "Sales"
               CaptionWidth    =   1300
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
            Begin BiSATextBoxProject.BiSATextBox cFakturAsli 
               Height          =   330
               Left            =   3420
               TabIndex        =   25
               TabStop         =   0   'False
               Top             =   615
               Width           =   3795
               _ExtentX        =   6694
               _ExtentY        =   582
               Text            =   "12345678901234567890"
               BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontBold        =   -1  'True
               FontName        =   "Verdana"
               BackColor       =   16777215
               MaxLength       =   20
               Appearance      =   0
               Caption         =   "No Faktur Asli"
               CaptionWidth    =   1300
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
            Begin VB.Label Label6 
               Caption         =   "Label6"
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
               Left            =   180
               TabIndex        =   29
               Top             =   1905
               Width           =   12255
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Frame2"
            Height          =   2325
            Left            =   -13725
            TabIndex        =   16
            Top             =   330
            Width           =   13080
            Begin BiSANumberBoxProject.BiSANumberBox BiSANumberBox1 
               Height          =   1590
               Left            =   1950
               TabIndex        =   20
               TabStop         =   0   'False
               Top             =   570
               Width           =   10980
               _ExtentX        =   19368
               _ExtentY        =   2805
               Appearance      =   0
               Decimals        =   0
               Enabled         =   0   'False
               BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   72
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BackColor       =   8454143
               Caption         =   " "
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
            Begin BiSAFramProject.BiSAFrame BiSAFrame10 
               Height          =   1440
               Left            =   105
               Top             =   720
               Visible         =   0   'False
               Width           =   1665
               _ExtentX        =   2937
               _ExtentY        =   2540
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
               Begin BiSANumberBoxProject.BiSANumberBox nInfoStok 
                  Height          =   330
                  Left            =   195
                  TabIndex        =   19
                  Top             =   525
                  Width           =   1320
                  _ExtentX        =   2328
                  _ExtentY        =   582
                  Appearance      =   0
                  Decimals        =   1
                  MinValue        =   0
                  Enabled         =   0   'False
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
                  Caption         =   " "
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
               Begin VB.Shape Shape1 
                  BackColor       =   &H00C0FFC0&
                  BorderColor     =   &H000000FF&
                  BorderStyle     =   3  'Dot
                  FillColor       =   &H00C0FFC0&
                  FillStyle       =   0  'Solid
                  Height          =   870
                  Left            =   195
                  Shape           =   1  'Square
                  Top             =   285
                  Width           =   1290
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "INFO STOK"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   -1260
                  TabIndex        =   18
                  Top             =   630
                  Width           =   870
               End
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "INFO STOK"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   90
               TabIndex        =   17
               Top             =   165
               Width           =   6735
            End
         End
      End
      Begin BiSAFramProject.BiSAFrame BiSAFrame1 
         Height          =   2700
         Left            =   150
         Top             =   1425
         Width           =   5685
         _ExtentX        =   10028
         _ExtentY        =   4763
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   4
         BackColor       =   -2147483633
         Begin MSComDlg.CommonDialog CommonDialog2 
            Left            =   4455
            Top             =   630
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Timer Timer1 
            Left            =   3870
            Top             =   195
         End
         Begin BiSADateProject.BiSADate dTgl 
            Height          =   324
            Left            =   180
            TabIndex        =   6
            Top             =   372
            Width           =   3060
            _ExtentX        =   5398
            _ExtentY        =   582
            Value           =   "03-04-2011"
            Appearance      =   0
            BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
            Caption         =   "Tanggal"
            CaptionWidth    =   1400
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
         Begin BiSATextBoxProject.BiSABrowse cFaktur 
            Height          =   330
            Left            =   180
            TabIndex        =   7
            Top             =   735
            Width           =   3645
            _ExtentX        =   6429
            _ExtentY        =   582
            Text            =   "12345678"
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
            Caption         =   "Nomor"
            CaptionWidth    =   1400
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
         Begin BiSATextBoxProject.BiSABrowse cGudang 
            Height          =   330
            Left            =   180
            TabIndex        =   8
            Top             =   1095
            Width           =   3060
            _ExtentX        =   5398
            _ExtentY        =   582
            Text            =   "12345678"
            BorderStyle     =   0
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
            Button          =   -1  'True
            Caption         =   "Gudang"
            CaptionWidth    =   1400
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
         Begin BiSATextBoxProject.BiSATextBox cNamaGudang 
            Height          =   336
            Left            =   3264
            TabIndex        =   9
            Top             =   1092
            Width           =   2328
            _ExtentX        =   4101
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
            BackColor       =   -2147483633
            Enabled         =   0   'False
            Appearance      =   0
            CaptionWidth    =   1400
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
         Begin BiSATextBoxProject.BiSABrowse cAkunKas 
            Height          =   330
            Left            =   180
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   1455
            Width           =   3060
            _ExtentX        =   5398
            _ExtentY        =   582
            BorderStyle     =   0
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
            Button          =   -1  'True
            Caption         =   "Akun Kas"
            CaptionWidth    =   1400
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
         Begin BiSAFramProject.BiSAFrame BiSAFrame8 
            Height          =   870
            Left            =   120
            Top             =   1800
            Width           =   5505
            _ExtentX        =   9710
            _ExtentY        =   1535
            BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483637
            Begin BiSATextBoxProject.BiSABrowse cNamaCustomer 
               Height          =   330
               Left            =   3150
               TabIndex        =   12
               Top             =   90
               Width           =   2280
               _ExtentX        =   4022
               _ExtentY        =   582
               BorderStyle     =   0
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
               Button          =   -1  'True
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
            Begin BiSATextBoxProject.BiSATextBox cKota 
               Height          =   330
               Left            =   3990
               TabIndex        =   14
               Top             =   465
               Width           =   1425
               _ExtentX        =   2514
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
               BackColor       =   -2147483633
               Enabled         =   0   'False
               Appearance      =   0
               CaptionWidth    =   0
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
            Begin BiSATextBoxProject.BiSATextBox cAlamat 
               Height          =   330
               Left            =   60
               TabIndex        =   13
               Top             =   465
               Width           =   3915
               _ExtentX        =   6906
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
               BackColor       =   -2147483633
               Enabled         =   0   'False
               Appearance      =   0
               Caption         =   "Alamat"
               CaptionWidth    =   1400
               CaptionBackColor=   -2147483637
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
            Begin BiSATextBoxProject.BiSABrowse cCustomer 
               Height          =   330
               Left            =   60
               TabIndex        =   11
               Top             =   90
               Width           =   3060
               _ExtentX        =   5398
               _ExtentY        =   582
               Text            =   "12345678"
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
               BackColor       =   -2147483638
               Enabled         =   0   'False
               Appearance      =   0
               Caption         =   "Pelanggan"
               CaptionWidth    =   1400
               CaptionBackColor=   -2147483637
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
         Begin BiSAButtonProject.BiSAButton cmdAddOrder 
            Height          =   330
            Left            =   4785
            TabIndex        =   5
            Top             =   285
            Visible         =   0   'False
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   582
            Caption         =   "Or"
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
      End
      Begin SizerOneLibCtl.TabOne TabOne1 
         Height          =   5310
         Left            =   135
         TabIndex        =   30
         Top             =   4170
         Width           =   18975
         _cx             =   33470
         _cy             =   9366
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   2
         MousePointer    =   0
         _ConvInfo       =   1
         Version         =   700
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FrontTabColor   =   -2147483633
         BackTabColor    =   -2147483633
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   -2147483630
         Caption         =   "PENJUALAN|&2 KONSINYASI KHUSUS|HADIAH"
         Align           =   0
         CurrTab         =   0
         FirstTab        =   0
         Style           =   3
         Position        =   0
         AutoSwitch      =   -1  'True
         AutoScroll      =   -1  'True
         TabPreview      =   -1  'True
         ShowFocusRect   =   -1  'True
         TabsPerPage     =   0
         BorderWidth     =   0
         BoldCurrent     =   -1  'True
         DogEars         =   -1  'True
         MultiRow        =   0   'False
         MultiRowOffset  =   200
         CaptionStyle    =   0
         TabHeight       =   0
         TabCaptionPos   =   4
         TabPicturePos   =   0
         CaptionEmpty    =   ""
         Separators      =   0   'False
         Flags(1)        =   1
         Flags(2)        =   1
         Begin VB.Frame Frame4 
            Height          =   4935
            Left            =   19920
            TabIndex        =   70
            Top             =   330
            Width           =   18885
            Begin TrueOleDBGrid70.TDBGrid TDBGrid3 
               Height          =   4635
               Left            =   75
               TabIndex        =   85
               Top             =   180
               Width           =   18630
               _ExtentX        =   32861
               _ExtentY        =   8176
               _LayoutType     =   4
               _RowHeight      =   -2147483647
               _WasPersistedAsPixels=   0
               Columns(0)._VlistStyle=   0
               Columns(0)._MaxComboItems=   5
               Columns(0).Caption=   "NO."
               Columns(0).DataField=   ""
               Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(1)._VlistStyle=   0
               Columns(1)._MaxComboItems=   5
               Columns(1).Caption=   "BELI QTY"
               Columns(1).DataField=   ""
               Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(2)._VlistStyle=   0
               Columns(2)._MaxComboItems=   5
               Columns(2).Caption=   "NAMA BARANG"
               Columns(2).DataField=   ""
               Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(3)._VlistStyle=   0
               Columns(3)._MaxComboItems=   5
               Columns(3).Caption=   "GRATIS"
               Columns(3).DataField=   ""
               Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(4)._VlistStyle=   0
               Columns(4)._MaxComboItems=   5
               Columns(4).Caption=   "NAMA BARANG"
               Columns(4).DataField=   ""
               Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns.Count   =   5
               Splits(0)._UserFlags=   0
               Splits(0).PartialRightColumn=   0   'False
               Splits(0).MarqueeStyle=   3
               Splits(0).AllowRowSizing=   0   'False
               Splits(0).RecordSelectors=   0   'False
               Splits(0).RecordSelectorWidth=   503
               Splits(0)._SavedRecordSelectors=   0   'False
               Splits(0).AllowColSelect=   0   'False
               Splits(0).AllowRowSelect=   0   'False
               Splits(0).DividerColor=   12632256
               Splits(0).SpringMode=   0   'False
               Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
               Splits(0)._ColumnProps(0)=   "Columns.Count=5"
               Splits(0)._ColumnProps(1)=   "Column(0).Width=926"
               Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
               Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=847"
               Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
               Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=514"
               Splits(0)._ColumnProps(6)=   "Column(0).WrapText=1"
               Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
               Splits(0)._ColumnProps(8)=   "Column(1).Width=3651"
               Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
               Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=3572"
               Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
               Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=512"
               Splits(0)._ColumnProps(13)=   "Column(1).WrapText=1"
               Splits(0)._ColumnProps(14)=   "Column(1).Order=2"
               Splits(0)._ColumnProps(15)=   "Column(2).Width=9790"
               Splits(0)._ColumnProps(16)=   "Column(2).DividerColor=0"
               Splits(0)._ColumnProps(17)=   "Column(2)._WidthInPix=9710"
               Splits(0)._ColumnProps(18)=   "Column(2)._EditAlways=0"
               Splits(0)._ColumnProps(19)=   "Column(2)._ColStyle=512"
               Splits(0)._ColumnProps(20)=   "Column(2).WrapText=1"
               Splits(0)._ColumnProps(21)=   "Column(2).Order=3"
               Splits(0)._ColumnProps(22)=   "Column(3).Width=2381"
               Splits(0)._ColumnProps(23)=   "Column(3).DividerColor=0"
               Splits(0)._ColumnProps(24)=   "Column(3)._WidthInPix=2302"
               Splits(0)._ColumnProps(25)=   "Column(3)._EditAlways=0"
               Splits(0)._ColumnProps(26)=   "Column(3)._ColStyle=197122"
               Splits(0)._ColumnProps(27)=   "Column(3).WrapText=1"
               Splits(0)._ColumnProps(28)=   "Column(3).Order=4"
               Splits(0)._ColumnProps(29)=   "Column(4).Width=2699"
               Splits(0)._ColumnProps(30)=   "Column(4).DividerColor=0"
               Splits(0)._ColumnProps(31)=   "Column(4)._WidthInPix=2619"
               Splits(0)._ColumnProps(32)=   "Column(4)._EditAlways=0"
               Splits(0)._ColumnProps(33)=   "Column(4)._ColStyle=513"
               Splits(0)._ColumnProps(34)=   "Column(4).WrapText=1"
               Splits(0)._ColumnProps(35)=   "Column(4).Order=5"
               Splits.Count    =   1
               PrintInfos(0)._StateFlags=   3
               PrintInfos(0).Name=   "piInternal 0"
               PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
               PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
               PrintInfos(0).PageHeaderHeight=   0
               PrintInfos(0).PageFooterHeight=   0
               PrintInfos.Count=   1
               AllowUpdate     =   0   'False
               BorderStyle     =   0
               ColumnFooters   =   -1  'True
               DataMode        =   4
               DefColWidth     =   0
               HeadLines       =   1,5
               FootLines       =   1
               MultipleLines   =   0
               CellTipsWidth   =   0
               InsertMode      =   0   'False
               DeadAreaBackColor=   16777215
               RowDividerColor =   12632256
               RowSubDividerColor=   12632256
               DirectionAfterEnter=   1
               MaxRows         =   250000
               ViewColumnCaptionWidth=   0
               ViewColumnWidth =   0
               _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
               _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
               _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
               _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
               _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=2,.bold=0,.fontsize=825,.italic=0"
               _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
               _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000005&,.bold=0,.fontsize=825"
               _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
               _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
               _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.bgcolor=&H8000000F&"
               _StyleDefs(11)  =   ":id=2,.fgcolor=&H80000007&,.bold=0,.fontsize=825,.italic=0,.underline=0"
               _StyleDefs(12)  =   ":id=2,.strikethrough=0,.charset=0"
               _StyleDefs(13)  =   ":id=2,.fontname=Tahoma"
               _StyleDefs(14)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
               _StyleDefs(15)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(16)  =   ":id=3,.fontname=Tahoma"
               _StyleDefs(17)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(18)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
               _StyleDefs(19)  =   "EditorStyle:id=7,.parent=1"
               _StyleDefs(20)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
               _StyleDefs(21)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
               _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
               _StyleDefs(23)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
               _StyleDefs(24)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
               _StyleDefs(25)  =   "Splits(0).Style:id=13,.parent=1,.wraptext=-1"
               _StyleDefs(26)  =   "Splits(0).CaptionStyle:id=22,.parent=4,.bold=0,.fontsize=825,.italic=0"
               _StyleDefs(27)  =   ":id=22,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(28)  =   ":id=22,.fontname=Tahoma"
               _StyleDefs(29)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bold=-1,.fontsize=825,.italic=0"
               _StyleDefs(30)  =   ":id=14,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(31)  =   ":id=14,.fontname=Verdana"
               _StyleDefs(32)  =   "Splits(0).FooterStyle:id=15,.parent=3"
               _StyleDefs(33)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
               _StyleDefs(34)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
               _StyleDefs(35)  =   "Splits(0).EditorStyle:id=17,.parent=7"
               _StyleDefs(36)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
               _StyleDefs(37)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
               _StyleDefs(38)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
               _StyleDefs(39)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
               _StyleDefs(40)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
               _StyleDefs(41)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=1,.bold=0,.fontsize=825"
               _StyleDefs(42)  =   ":id=28,.italic=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(43)  =   ":id=28,.fontname=Tahoma"
               _StyleDefs(44)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14,.bold=-1,.fontsize=825"
               _StyleDefs(45)  =   ":id=25,.italic=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(46)  =   ":id=25,.fontname=Tahoma"
               _StyleDefs(47)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
               _StyleDefs(48)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
               _StyleDefs(49)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=0"
               _StyleDefs(50)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
               _StyleDefs(51)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
               _StyleDefs(52)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
               _StyleDefs(53)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=0"
               _StyleDefs(54)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
               _StyleDefs(55)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
               _StyleDefs(56)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
               _StyleDefs(57)  =   "Splits(0).Columns(3).Style:id=66,.parent=13,.alignment=1"
               _StyleDefs(58)  =   "Splits(0).Columns(3).HeadingStyle:id=63,.parent=14"
               _StyleDefs(59)  =   "Splits(0).Columns(3).FooterStyle:id=64,.parent=15,.alignment=1"
               _StyleDefs(60)  =   "Splits(0).Columns(3).EditorStyle:id=65,.parent=17"
               _StyleDefs(61)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.alignment=2"
               _StyleDefs(62)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
               _StyleDefs(63)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
               _StyleDefs(64)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
               _StyleDefs(65)  =   "Named:id=33:Normal"
               _StyleDefs(66)  =   ":id=33,.parent=0"
               _StyleDefs(67)  =   "Named:id=34:Heading"
               _StyleDefs(68)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(69)  =   ":id=34,.wraptext=-1"
               _StyleDefs(70)  =   "Named:id=35:Footing"
               _StyleDefs(71)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(72)  =   "Named:id=36:Selected"
               _StyleDefs(73)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(74)  =   "Named:id=37:Caption"
               _StyleDefs(75)  =   ":id=37,.parent=34,.alignment=2"
               _StyleDefs(76)  =   "Named:id=38:HighlightRow"
               _StyleDefs(77)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(78)  =   "Named:id=39:EvenRow"
               _StyleDefs(79)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
               _StyleDefs(80)  =   "Named:id=40:OddRow"
               _StyleDefs(81)  =   ":id=40,.parent=33"
               _StyleDefs(82)  =   "Named:id=41:RecordSelector"
               _StyleDefs(83)  =   ":id=41,.parent=34"
               _StyleDefs(84)  =   "Named:id=42:FilterBar"
               _StyleDefs(85)  =   ":id=42,.parent=33"
            End
         End
         Begin BiSAFramProject.BiSAFrame BiSAFrame5 
            Height          =   4935
            Left            =   19620
            Top             =   330
            Width           =   18885
            _ExtentX        =   33311
            _ExtentY        =   8705
            BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderStyle     =   4
            BackColor       =   -2147483633
            Begin TrueOleDBGrid70.TDBGrid TDBGrid2 
               Height          =   3300
               Left            =   60
               TabIndex        =   69
               Top             =   765
               Width           =   18720
               _ExtentX        =   33020
               _ExtentY        =   5821
               _LayoutType     =   4
               _RowHeight      =   -2147483647
               _WasPersistedAsPixels=   0
               Columns(0)._VlistStyle=   0
               Columns(0)._MaxComboItems=   5
               Columns(0).Caption=   "NO."
               Columns(0).DataField=   ""
               Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(1)._VlistStyle=   0
               Columns(1)._MaxComboItems=   5
               Columns(1).Caption=   "KD.SUPPLIER"
               Columns(1).DataField=   ""
               Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(2)._VlistStyle=   0
               Columns(2)._MaxComboItems=   5
               Columns(2).Caption=   "NM.SUPPLIER"
               Columns(2).DataField=   ""
               Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(3)._VlistStyle=   0
               Columns(3)._MaxComboItems=   5
               Columns(3).Caption=   "BARCODE"
               Columns(3).DataField=   ""
               Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(4)._VlistStyle=   0
               Columns(4)._MaxComboItems=   5
               Columns(4).Caption=   "ITEM BRG"
               Columns(4).DataField=   ""
               Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(5)._VlistStyle=   0
               Columns(5)._MaxComboItems=   5
               Columns(5).Caption=   "QTY"
               Columns(5).DataField=   ""
               Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(6)._VlistStyle=   0
               Columns(6)._MaxComboItems=   5
               Columns(6).Caption=   "H.JUAL"
               Columns(6).DataField=   ""
               Columns(6).NumberFormat=   "###,###,###,###"
               Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(7)._VlistStyle=   0
               Columns(7)._MaxComboItems=   5
               Columns(7).Caption=   "H.BELI"
               Columns(7).DataField=   ""
               Columns(7).NumberFormat=   "###,###,###,###,##0.00"
               Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(8)._VlistStyle=   0
               Columns(8)._MaxComboItems=   5
               Columns(8).Caption=   "JUMLAH"
               Columns(8).DataField=   ""
               Columns(8).NumberFormat=   "###,###,###,###"
               Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns.Count   =   9
               Splits(0)._UserFlags=   0
               Splits(0).PartialRightColumn=   0   'False
               Splits(0).MarqueeStyle=   3
               Splits(0).AllowRowSizing=   0   'False
               Splits(0).RecordSelectors=   0   'False
               Splits(0).RecordSelectorWidth=   503
               Splits(0)._SavedRecordSelectors=   0   'False
               Splits(0).AllowColSelect=   0   'False
               Splits(0).AllowRowSelect=   0   'False
               Splits(0).DividerColor=   12632256
               Splits(0).SpringMode=   0   'False
               Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
               Splits(0)._ColumnProps(0)=   "Columns.Count=9"
               Splits(0)._ColumnProps(1)=   "Column(0).Width=926"
               Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
               Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=847"
               Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
               Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=514"
               Splits(0)._ColumnProps(6)=   "Column(0).WrapText=1"
               Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
               Splits(0)._ColumnProps(8)=   "Column(1).Width=2805"
               Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
               Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2725"
               Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
               Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=512"
               Splits(0)._ColumnProps(13)=   "Column(1).WrapText=1"
               Splits(0)._ColumnProps(14)=   "Column(1).Order=2"
               Splits(0)._ColumnProps(15)=   "Column(2).Width=4445"
               Splits(0)._ColumnProps(16)=   "Column(2).DividerColor=0"
               Splits(0)._ColumnProps(17)=   "Column(2)._WidthInPix=4366"
               Splits(0)._ColumnProps(18)=   "Column(2)._EditAlways=0"
               Splits(0)._ColumnProps(19)=   "Column(2)._ColStyle=516"
               Splits(0)._ColumnProps(20)=   "Column(2).WrapText=1"
               Splits(0)._ColumnProps(21)=   "Column(2).Order=3"
               Splits(0)._ColumnProps(22)=   "Column(3).Width=3519"
               Splits(0)._ColumnProps(23)=   "Column(3).DividerColor=0"
               Splits(0)._ColumnProps(24)=   "Column(3)._WidthInPix=3440"
               Splits(0)._ColumnProps(25)=   "Column(3)._EditAlways=0"
               Splits(0)._ColumnProps(26)=   "Column(3)._ColStyle=512"
               Splits(0)._ColumnProps(27)=   "Column(3).WrapText=1"
               Splits(0)._ColumnProps(28)=   "Column(3).Order=4"
               Splits(0)._ColumnProps(29)=   "Column(4).Width=9313"
               Splits(0)._ColumnProps(30)=   "Column(4).DividerColor=0"
               Splits(0)._ColumnProps(31)=   "Column(4)._WidthInPix=9234"
               Splits(0)._ColumnProps(32)=   "Column(4)._EditAlways=0"
               Splits(0)._ColumnProps(33)=   "Column(4)._ColStyle=197122"
               Splits(0)._ColumnProps(34)=   "Column(4).WrapText=1"
               Splits(0)._ColumnProps(35)=   "Column(4).Order=5"
               Splits(0)._ColumnProps(36)=   "Column(5).Width=1746"
               Splits(0)._ColumnProps(37)=   "Column(5).DividerColor=0"
               Splits(0)._ColumnProps(38)=   "Column(5)._WidthInPix=1667"
               Splits(0)._ColumnProps(39)=   "Column(5)._EditAlways=0"
               Splits(0)._ColumnProps(40)=   "Column(5)._ColStyle=513"
               Splits(0)._ColumnProps(41)=   "Column(5).WrapText=1"
               Splits(0)._ColumnProps(42)=   "Column(5).Order=6"
               Splits(0)._ColumnProps(43)=   "Column(6).Width=2672"
               Splits(0)._ColumnProps(44)=   "Column(6).DividerColor=0"
               Splits(0)._ColumnProps(45)=   "Column(6)._WidthInPix=2593"
               Splits(0)._ColumnProps(46)=   "Column(6)._EditAlways=0"
               Splits(0)._ColumnProps(47)=   "Column(6)._ColStyle=514"
               Splits(0)._ColumnProps(48)=   "Column(6).WrapText=1"
               Splits(0)._ColumnProps(49)=   "Column(6).Order=7"
               Splits(0)._ColumnProps(50)=   "Column(7).Width=2884"
               Splits(0)._ColumnProps(51)=   "Column(7).DividerColor=0"
               Splits(0)._ColumnProps(52)=   "Column(7)._WidthInPix=2805"
               Splits(0)._ColumnProps(53)=   "Column(7)._EditAlways=0"
               Splits(0)._ColumnProps(54)=   "Column(7)._ColStyle=514"
               Splits(0)._ColumnProps(55)=   "Column(7).WrapText=1"
               Splits(0)._ColumnProps(56)=   "Column(7).Order=8"
               Splits(0)._ColumnProps(57)=   "Column(8).Width=4471"
               Splits(0)._ColumnProps(58)=   "Column(8).DividerColor=0"
               Splits(0)._ColumnProps(59)=   "Column(8)._WidthInPix=4392"
               Splits(0)._ColumnProps(60)=   "Column(8)._EditAlways=0"
               Splits(0)._ColumnProps(61)=   "Column(8)._ColStyle=514"
               Splits(0)._ColumnProps(62)=   "Column(8).WrapText=1"
               Splits(0)._ColumnProps(63)=   "Column(8).Order=9"
               Splits.Count    =   1
               PrintInfos(0)._StateFlags=   3
               PrintInfos(0).Name=   "piInternal 0"
               PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
               PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
               PrintInfos(0).PageHeaderHeight=   0
               PrintInfos(0).PageFooterHeight=   0
               PrintInfos.Count=   1
               AllowUpdate     =   0   'False
               BorderStyle     =   0
               ColumnFooters   =   -1  'True
               DataMode        =   4
               DefColWidth     =   0
               HeadLines       =   1,5
               FootLines       =   1
               MultipleLines   =   0
               CellTipsWidth   =   0
               InsertMode      =   0   'False
               DeadAreaBackColor=   16777215
               RowDividerColor =   12632256
               RowSubDividerColor=   12632256
               DirectionAfterEnter=   1
               MaxRows         =   250000
               ViewColumnCaptionWidth=   0
               ViewColumnWidth =   0
               _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
               _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
               _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
               _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
               _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=2,.bold=0,.fontsize=825,.italic=0"
               _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
               _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000005&,.bold=0,.fontsize=825"
               _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
               _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
               _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.bgcolor=&H8000000F&"
               _StyleDefs(11)  =   ":id=2,.fgcolor=&H80000007&,.bold=0,.fontsize=825,.italic=0,.underline=0"
               _StyleDefs(12)  =   ":id=2,.strikethrough=0,.charset=0"
               _StyleDefs(13)  =   ":id=2,.fontname=Tahoma"
               _StyleDefs(14)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
               _StyleDefs(15)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(16)  =   ":id=3,.fontname=Tahoma"
               _StyleDefs(17)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(18)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
               _StyleDefs(19)  =   "EditorStyle:id=7,.parent=1"
               _StyleDefs(20)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
               _StyleDefs(21)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
               _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
               _StyleDefs(23)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
               _StyleDefs(24)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
               _StyleDefs(25)  =   "Splits(0).Style:id=13,.parent=1,.wraptext=-1"
               _StyleDefs(26)  =   "Splits(0).CaptionStyle:id=22,.parent=4,.bold=0,.fontsize=825,.italic=0"
               _StyleDefs(27)  =   ":id=22,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(28)  =   ":id=22,.fontname=Tahoma"
               _StyleDefs(29)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bold=-1,.fontsize=825,.italic=0"
               _StyleDefs(30)  =   ":id=14,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(31)  =   ":id=14,.fontname=Verdana"
               _StyleDefs(32)  =   "Splits(0).FooterStyle:id=15,.parent=3"
               _StyleDefs(33)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
               _StyleDefs(34)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
               _StyleDefs(35)  =   "Splits(0).EditorStyle:id=17,.parent=7"
               _StyleDefs(36)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
               _StyleDefs(37)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
               _StyleDefs(38)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
               _StyleDefs(39)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
               _StyleDefs(40)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
               _StyleDefs(41)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=1,.bold=0,.fontsize=825"
               _StyleDefs(42)  =   ":id=28,.italic=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(43)  =   ":id=28,.fontname=Tahoma"
               _StyleDefs(44)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14,.bold=-1,.fontsize=825"
               _StyleDefs(45)  =   ":id=25,.italic=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(46)  =   ":id=25,.fontname=Tahoma"
               _StyleDefs(47)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
               _StyleDefs(48)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
               _StyleDefs(49)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=0"
               _StyleDefs(50)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
               _StyleDefs(51)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
               _StyleDefs(52)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
               _StyleDefs(53)  =   "Splits(0).Columns(2).Style:id=70,.parent=13"
               _StyleDefs(54)  =   "Splits(0).Columns(2).HeadingStyle:id=67,.parent=14"
               _StyleDefs(55)  =   "Splits(0).Columns(2).FooterStyle:id=68,.parent=15"
               _StyleDefs(56)  =   "Splits(0).Columns(2).EditorStyle:id=69,.parent=17"
               _StyleDefs(57)  =   "Splits(0).Columns(3).Style:id=46,.parent=13,.alignment=0"
               _StyleDefs(58)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
               _StyleDefs(59)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
               _StyleDefs(60)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
               _StyleDefs(61)  =   "Splits(0).Columns(4).Style:id=66,.parent=13,.alignment=1"
               _StyleDefs(62)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=14"
               _StyleDefs(63)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=15,.alignment=1"
               _StyleDefs(64)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=17"
               _StyleDefs(65)  =   "Splits(0).Columns(5).Style:id=54,.parent=13,.alignment=2"
               _StyleDefs(66)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
               _StyleDefs(67)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
               _StyleDefs(68)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
               _StyleDefs(69)  =   "Splits(0).Columns(6).Style:id=50,.parent=13,.alignment=1"
               _StyleDefs(70)  =   "Splits(0).Columns(6).HeadingStyle:id=47,.parent=14"
               _StyleDefs(71)  =   "Splits(0).Columns(6).FooterStyle:id=48,.parent=15"
               _StyleDefs(72)  =   "Splits(0).Columns(6).EditorStyle:id=49,.parent=17"
               _StyleDefs(73)  =   "Splits(0).Columns(7).Style:id=58,.parent=13,.alignment=1"
               _StyleDefs(74)  =   "Splits(0).Columns(7).HeadingStyle:id=55,.parent=14"
               _StyleDefs(75)  =   "Splits(0).Columns(7).FooterStyle:id=56,.parent=15"
               _StyleDefs(76)  =   "Splits(0).Columns(7).EditorStyle:id=57,.parent=17"
               _StyleDefs(77)  =   "Splits(0).Columns(8).Style:id=62,.parent=13,.alignment=1"
               _StyleDefs(78)  =   "Splits(0).Columns(8).HeadingStyle:id=59,.parent=14"
               _StyleDefs(79)  =   "Splits(0).Columns(8).FooterStyle:id=60,.parent=15"
               _StyleDefs(80)  =   "Splits(0).Columns(8).EditorStyle:id=61,.parent=17"
               _StyleDefs(81)  =   "Named:id=33:Normal"
               _StyleDefs(82)  =   ":id=33,.parent=0"
               _StyleDefs(83)  =   "Named:id=34:Heading"
               _StyleDefs(84)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(85)  =   ":id=34,.wraptext=-1"
               _StyleDefs(86)  =   "Named:id=35:Footing"
               _StyleDefs(87)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(88)  =   "Named:id=36:Selected"
               _StyleDefs(89)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(90)  =   "Named:id=37:Caption"
               _StyleDefs(91)  =   ":id=37,.parent=34,.alignment=2"
               _StyleDefs(92)  =   "Named:id=38:HighlightRow"
               _StyleDefs(93)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(94)  =   "Named:id=39:EvenRow"
               _StyleDefs(95)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
               _StyleDefs(96)  =   "Named:id=40:OddRow"
               _StyleDefs(97)  =   ":id=40,.parent=33"
               _StyleDefs(98)  =   "Named:id=41:RecordSelector"
               _StyleDefs(99)  =   ":id=41,.parent=34"
               _StyleDefs(100) =   "Named:id=42:FilterBar"
               _StyleDefs(101) =   ":id=42,.parent=33"
            End
            Begin BiSATextBoxProject.BiSABrowse cNamaSupplier 
               Height          =   360
               Left            =   2145
               TabIndex        =   60
               Top             =   360
               Width           =   2535
               _ExtentX        =   4471
               _ExtentY        =   635
               BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontBold        =   -1  'True
               FontName        =   "Verdana"
               BackColor       =   12648384
               Appearance      =   0
               GetPicture      =   1
               Button          =   -1  'True
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
            Begin BiSATextBoxProject.BiSABrowse cNamaBarang2 
               Height          =   360
               Left            =   6660
               TabIndex        =   62
               Top             =   360
               Width           =   5310
               _ExtentX        =   9366
               _ExtentY        =   635
               BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontBold        =   -1  'True
               FontName        =   "Verdana"
               BackColor       =   12648384
               Appearance      =   0
               GetPicture      =   1
               Button          =   -1  'True
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
            Begin BiSANumberBoxProject.BiSANumberBox nQty2 
               Height          =   360
               Left            =   11985
               TabIndex        =   63
               Top             =   360
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   635
               Appearance      =   0
               Decimals        =   0
               MaxValue        =   9999999999999
               MinValue        =   1
               BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   " "
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
            Begin BiSANumberBoxProject.BiSANumberBox nHargaJual2 
               Height          =   360
               Left            =   12960
               TabIndex        =   64
               Top             =   360
               Width           =   1500
               _ExtentX        =   2646
               _ExtentY        =   635
               Appearance      =   0
               Decimals        =   1
               MinValue        =   0
               BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   " "
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
            Begin BiSANumberBoxProject.BiSANumberBox nHargaBeli2 
               Height          =   360
               Left            =   14475
               TabIndex        =   65
               Top             =   360
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   635
               Appearance      =   0
               Decimals        =   1
               MinValue        =   0
               BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   " "
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
            Begin BiSANumberBoxProject.BiSANumberBox nJumlah2 
               Height          =   360
               Left            =   16140
               TabIndex        =   66
               Top             =   360
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   635
               Appearance      =   0
               Decimals        =   1
               MinValue        =   0
               Enabled         =   0   'False
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
               Caption         =   " "
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
            Begin BiSATextBoxProject.BiSATextBox cKodeSupplier 
               Height          =   360
               Left            =   570
               TabIndex        =   59
               Top             =   360
               Width           =   1560
               _ExtentX        =   2752
               _ExtentY        =   635
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
               BackColor       =   -2147483633
               Enabled         =   0   'False
               Appearance      =   0
               CaptionWidth    =   1400
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
            Begin BiSATextBoxProject.BiSATextBox cBarcode2 
               Height          =   360
               Left            =   4695
               TabIndex        =   61
               Top             =   360
               Width           =   1950
               _ExtentX        =   3440
               _ExtentY        =   635
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
               BackColor       =   -2147483633
               Enabled         =   0   'False
               Appearance      =   0
               CaptionWidth    =   1400
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
            Begin BiSANumberBoxProject.BiSANumberBox nNo2 
               Height          =   360
               Left            =   60
               TabIndex        =   58
               Top             =   360
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   635
               Appearance      =   0
               Decimals        =   0
               DecimalPoint    =   ""
               BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   " "
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
            Begin BiSAButtonProject.BiSAButton cmdOK2 
               Height          =   345
               Left            =   17955
               TabIndex        =   67
               Top             =   360
               Width           =   420
               _ExtentX        =   741
               _ExtentY        =   609
               Caption         =   "+"
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
            End
            Begin BiSAButtonProject.BiSAButton cmdDel2 
               Height          =   345
               Left            =   18390
               TabIndex        =   68
               TabStop         =   0   'False
               Top             =   360
               Width           =   390
               _ExtentX        =   688
               _ExtentY        =   609
               Caption         =   "-"
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
            End
         End
         Begin BiSAFramProject.BiSAFrame BisaFrame2 
            Height          =   4935
            Left            =   45
            Top             =   330
            Width           =   18885
            _ExtentX        =   33311
            _ExtentY        =   8705
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
            Begin BiSAButtonProject.BiSAButton Command1 
               Height          =   360
               Left            =   18330
               TabIndex        =   40
               Top             =   120
               Width           =   390
               _ExtentX        =   688
               _ExtentY        =   635
               Caption         =   "-"
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
               Picture         =   "trPenjualan.frx":0000
            End
            Begin BiSAButtonProject.BiSAButton cmdOK 
               Height          =   360
               Left            =   17865
               TabIndex        =   39
               Top             =   120
               Width           =   435
               _ExtentX        =   767
               _ExtentY        =   635
               Caption         =   ""
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
               Picture         =   "trPenjualan.frx":059A
            End
            Begin BiSAButtonProject.BiSAButton BiSAButton2 
               Height          =   375
               Left            =   12780
               TabIndex        =   42
               Top             =   2085
               Visible         =   0   'False
               Width           =   435
               _ExtentX        =   767
               _ExtentY        =   661
               Caption         =   "Clr"
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
            Begin BiSAButtonProject.BiSAButton BiSAButton1 
               Height          =   330
               Left            =   13260
               TabIndex        =   45
               Top             =   2655
               Visible         =   0   'False
               Width           =   435
               _ExtentX        =   767
               _ExtentY        =   582
               Caption         =   "+"
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
            Begin BiSANumberBoxProject.BiSANumberBox nDP 
               Height          =   345
               Left            =   15540
               TabIndex        =   44
               Top             =   2190
               Visible         =   0   'False
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   609
               Appearance      =   0
               Decimals        =   0
               Enabled         =   0   'False
               BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   " "
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
            Begin BiSAFramProject.BiSAFrame BiSAFrame9 
               Height          =   765
               Left            =   15555
               Top             =   4020
               Width           =   3225
               _ExtentX        =   5689
               _ExtentY        =   1349
               BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BorderStyle     =   0
               BackColor       =   12648384
               Begin VB.CheckBox chkTunai 
                  BackColor       =   &H00C0FFC0&
                  Height          =   300
                  Left            =   1110
                  TabIndex        =   47
                  TabStop         =   0   'False
                  Top             =   75
                  Width           =   255
               End
               Begin BiSANumberBoxProject.BiSANumberBox nTunai 
                  Height          =   330
                  Left            =   1410
                  TabIndex        =   48
                  Top             =   60
                  Width           =   1725
                  _ExtentX        =   3043
                  _ExtentY        =   582
                  Appearance      =   0
                  Decimals        =   0
                  Enabled         =   0   'False
                  BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   " "
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
               Begin BiSANumberBoxProject.BiSANumberBox nPiutang 
                  Height          =   330
                  Left            =   1410
                  TabIndex        =   50
                  Top             =   405
                  Width           =   1725
                  _ExtentX        =   3043
                  _ExtentY        =   582
                  Appearance      =   0
                  Decimals        =   0
                  Enabled         =   0   'False
                  BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   " "
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
               Begin VB.Label Label1 
                  BackColor       =   &H00C0FFC0&
                  Caption         =   "T U N A I"
                  BeginProperty Font 
                     Name            =   "Times New Roman"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   225
                  TabIndex        =   49
                  Top             =   435
                  Width           =   1110
               End
            End
            Begin BiSAButtonProject.BiSAButton cmdImport 
               Height          =   420
               Left            =   17550
               TabIndex        =   43
               Top             =   2055
               Visible         =   0   'False
               Width           =   795
               _ExtentX        =   1402
               _ExtentY        =   741
               Caption         =   "I"
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
            Begin BiSANumberBoxProject.BiSANumberBox nQty 
               Height          =   360
               Left            =   8265
               TabIndex        =   34
               Top             =   120
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   635
               Appearance      =   0
               Decimals        =   1
               MinValue        =   0
               BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   " "
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
            Begin BiSANumberBoxProject.BiSANumberBox nHarga 
               Height          =   360
               Left            =   11145
               TabIndex        =   36
               Top             =   120
               Width           =   2655
               _ExtentX        =   4683
               _ExtentY        =   635
               Appearance      =   0
               Decimals        =   0
               Enabled         =   0   'False
               BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   " "
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
            Begin BiSATextBoxProject.BiSATextBox cSatuan 
               Height          =   360
               Left            =   9615
               TabIndex        =   35
               Top             =   120
               Width           =   1485
               _ExtentX        =   2619
               _ExtentY        =   635
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
               BackColor       =   -2147483633
               Enabled         =   0   'False
               Appearance      =   0
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
            Begin BiSATextBoxProject.BiSABrowse cNama 
               Height          =   360
               Left            =   2700
               TabIndex        =   33
               Top             =   120
               Width           =   5535
               _ExtentX        =   9763
               _ExtentY        =   635
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
               Button          =   -1  'True
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
            Begin BiSATextBoxProject.BiSABrowse cBarcode 
               Height          =   360
               Left            =   645
               TabIndex        =   32
               Top             =   120
               Width           =   2025
               _ExtentX        =   3572
               _ExtentY        =   635
               BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontBold        =   -1  'True
               FontName        =   "Verdana"
               BackColor       =   12648384
               Appearance      =   0
               GetPicture      =   1
               Button          =   -1  'True
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
            Begin BiSANumberBoxProject.BiSANumberBox nNomor 
               Height          =   360
               Left            =   105
               TabIndex        =   31
               Top             =   120
               Width           =   540
               _ExtentX        =   953
               _ExtentY        =   635
               Appearance      =   0
               Decimals        =   0
               DecimalPoint    =   ""
               BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   " "
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
            Begin BiSANumberBoxProject.BiSANumberBox nJumlah 
               Height          =   360
               Left            =   15390
               TabIndex        =   38
               Top             =   120
               Width           =   2445
               _ExtentX        =   4313
               _ExtentY        =   635
               Appearance      =   0
               Decimals        =   0
               Enabled         =   0   'False
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
               Caption         =   " "
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
            Begin BiSANumberBoxProject.BiSANumberBox nDisc1 
               Height          =   360
               Left            =   13830
               TabIndex        =   37
               Top             =   120
               Width           =   1545
               _ExtentX        =   2725
               _ExtentY        =   635
               Appearance      =   0
               Enabled         =   0   'False
               BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   " "
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
            Begin BiSAFramProject.BiSAFrame BiSAFrame4 
               Height          =   765
               Left            =   10260
               Top             =   4035
               Width           =   2520
               _ExtentX        =   4445
               _ExtentY        =   1349
               BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BorderStyle     =   0
               BackColor       =   12640511
               Begin BiSANumberBoxProject.BiSANumberBox nSubTotal 
                  Height          =   330
                  Left            =   60
                  TabIndex        =   54
                  Top             =   45
                  Width           =   2415
                  _ExtentX        =   4260
                  _ExtentY        =   582
                  Appearance      =   0
                  Decimals        =   0
                  Enabled         =   0   'False
                  BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   " Sub TTL"
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
               Begin BiSANumberBoxProject.BiSANumberBox nDiscount 
                  Height          =   315
                  Left            =   60
                  TabIndex        =   55
                  Top             =   405
                  Width           =   2415
                  _ExtentX        =   4260
                  _ExtentY        =   556
                  Appearance      =   0
                  Decimals        =   0
                  Enabled         =   0   'False
                  BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   " Disc .00"
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
            Begin BiSAFramProject.BiSAFrame BiSAFrame7 
               Height          =   765
               Left            =   12795
               Top             =   4035
               Width           =   2775
               _ExtentX        =   4895
               _ExtentY        =   1349
               BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BorderStyle     =   0
               BackColor       =   12640511
               Begin BiSANumberBoxProject.BiSANumberBox nPajak 
                  Height          =   330
                  Left            =   90
                  TabIndex        =   56
                  Top             =   45
                  Width           =   2640
                  _ExtentX        =   4657
                  _ExtentY        =   582
                  Appearance      =   0
                  Decimals        =   0
                  Enabled         =   0   'False
                  BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   " PPN .00"
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
               Begin BiSANumberBoxProject.BiSANumberBox nTotal 
                  Height          =   315
                  Left            =   90
                  TabIndex        =   57
                  Top             =   405
                  Width           =   2640
                  _ExtentX        =   4657
                  _ExtentY        =   556
                  Appearance      =   0
                  Decimals        =   0
                  Enabled         =   0   'False
                  BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "Total"
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
            Begin BiSATextBoxProject.BiSATextBox cKeterangan 
               Height          =   330
               Left            =   120
               TabIndex        =   46
               TabStop         =   0   'False
               Top             =   4080
               Width           =   7410
               _ExtentX        =   13070
               _ExtentY        =   582
               Text            =   "12345678901234567890"
               BorderStyle     =   0
               BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontBold        =   -1  'True
               FontName        =   "Verdana"
               BackColor       =   16777215
               MaxLength       =   20
               Appearance      =   0
               Caption         =   "Keterangan"
               CaptionWidth    =   1400
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
            Begin BiSATextBoxProject.BiSABrowse cNamaCOD 
               Height          =   345
               Left            =   105
               TabIndex        =   51
               Top             =   4470
               Width           =   3540
               _ExtentX        =   6244
               _ExtentY        =   609
               Text            =   "12345678"
               BorderStyle     =   0
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
               Button          =   -1  'True
               Caption         =   "Ongkir"
               CaptionWidth    =   1400
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
            Begin BiSATextBoxProject.BiSATextBox cKodeCOD 
               Height          =   330
               Left            =   3660
               TabIndex        =   52
               TabStop         =   0   'False
               Top             =   4470
               Width           =   2130
               _ExtentX        =   3757
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
               BackColor       =   -2147483633
               Enabled         =   0   'False
               Appearance      =   0
               CaptionWidth    =   1400
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
            Begin BiSANumberBoxProject.BiSANumberBox nHargaCOD 
               Height          =   330
               Left            =   5835
               TabIndex        =   53
               TabStop         =   0   'False
               Top             =   4470
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   582
               Appearance      =   0
               Enabled         =   0   'False
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
            Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
               Height          =   3390
               Left            =   105
               TabIndex        =   41
               Top             =   510
               Width           =   18630
               _ExtentX        =   32861
               _ExtentY        =   5980
               _LayoutType     =   4
               _RowHeight      =   -2147483647
               _WasPersistedAsPixels=   0
               Columns(0)._VlistStyle=   0
               Columns(0)._MaxComboItems=   5
               Columns(0).Caption=   "NO."
               Columns(0).DataField=   ""
               Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(1)._VlistStyle=   0
               Columns(1)._MaxComboItems=   5
               Columns(1).Caption=   "BARCODE"
               Columns(1).DataField=   ""
               Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(2)._VlistStyle=   0
               Columns(2)._MaxComboItems=   5
               Columns(2).Caption=   "NAMA BARANG"
               Columns(2).DataField=   ""
               Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(3)._VlistStyle=   0
               Columns(3)._MaxComboItems=   5
               Columns(3).Caption=   "QTY"
               Columns(3).DataField=   ""
               Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(4)._VlistStyle=   0
               Columns(4)._MaxComboItems=   5
               Columns(4).Caption=   "SATUAN"
               Columns(4).DataField=   ""
               Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(5)._VlistStyle=   0
               Columns(5)._MaxComboItems=   5
               Columns(5).Caption=   "HARGA"
               Columns(5).DataField=   ""
               Columns(5).NumberFormat=   "###,###,###,###"
               Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(6)._VlistStyle=   0
               Columns(6)._MaxComboItems=   5
               Columns(6).Caption=   "DISC"
               Columns(6).DataField=   ""
               Columns(6).NumberFormat=   "###,###,###,###,##0.00"
               Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(7)._VlistStyle=   0
               Columns(7)._MaxComboItems=   5
               Columns(7).Caption=   "JUMLAH"
               Columns(7).DataField=   ""
               Columns(7).NumberFormat=   "###,###,###,###"
               Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns.Count   =   8
               Splits(0)._UserFlags=   0
               Splits(0).PartialRightColumn=   0   'False
               Splits(0).MarqueeStyle=   3
               Splits(0).AllowRowSizing=   0   'False
               Splits(0).RecordSelectors=   0   'False
               Splits(0).RecordSelectorWidth=   503
               Splits(0)._SavedRecordSelectors=   0   'False
               Splits(0).AllowColSelect=   0   'False
               Splits(0).AllowRowSelect=   0   'False
               Splits(0).DividerColor=   12632256
               Splits(0).SpringMode=   0   'False
               Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
               Splits(0)._ColumnProps(0)=   "Columns.Count=8"
               Splits(0)._ColumnProps(1)=   "Column(0).Width=926"
               Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
               Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=847"
               Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
               Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=514"
               Splits(0)._ColumnProps(6)=   "Column(0).WrapText=1"
               Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
               Splits(0)._ColumnProps(8)=   "Column(1).Width=3651"
               Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
               Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=3572"
               Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
               Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=512"
               Splits(0)._ColumnProps(13)=   "Column(1).WrapText=1"
               Splits(0)._ColumnProps(14)=   "Column(1).Order=2"
               Splits(0)._ColumnProps(15)=   "Column(2).Width=9790"
               Splits(0)._ColumnProps(16)=   "Column(2).DividerColor=0"
               Splits(0)._ColumnProps(17)=   "Column(2)._WidthInPix=9710"
               Splits(0)._ColumnProps(18)=   "Column(2)._EditAlways=0"
               Splits(0)._ColumnProps(19)=   "Column(2)._ColStyle=512"
               Splits(0)._ColumnProps(20)=   "Column(2).WrapText=1"
               Splits(0)._ColumnProps(21)=   "Column(2).Order=3"
               Splits(0)._ColumnProps(22)=   "Column(3).Width=2381"
               Splits(0)._ColumnProps(23)=   "Column(3).DividerColor=0"
               Splits(0)._ColumnProps(24)=   "Column(3)._WidthInPix=2302"
               Splits(0)._ColumnProps(25)=   "Column(3)._EditAlways=0"
               Splits(0)._ColumnProps(26)=   "Column(3)._ColStyle=197122"
               Splits(0)._ColumnProps(27)=   "Column(3).WrapText=1"
               Splits(0)._ColumnProps(28)=   "Column(3).Order=4"
               Splits(0)._ColumnProps(29)=   "Column(4).Width=2699"
               Splits(0)._ColumnProps(30)=   "Column(4).DividerColor=0"
               Splits(0)._ColumnProps(31)=   "Column(4)._WidthInPix=2619"
               Splits(0)._ColumnProps(32)=   "Column(4)._EditAlways=0"
               Splits(0)._ColumnProps(33)=   "Column(4)._ColStyle=513"
               Splits(0)._ColumnProps(34)=   "Column(4).WrapText=1"
               Splits(0)._ColumnProps(35)=   "Column(4).Order=5"
               Splits(0)._ColumnProps(36)=   "Column(5).Width=4736"
               Splits(0)._ColumnProps(37)=   "Column(5).DividerColor=0"
               Splits(0)._ColumnProps(38)=   "Column(5)._WidthInPix=4657"
               Splits(0)._ColumnProps(39)=   "Column(5)._EditAlways=0"
               Splits(0)._ColumnProps(40)=   "Column(5)._ColStyle=514"
               Splits(0)._ColumnProps(41)=   "Column(5).WrapText=1"
               Splits(0)._ColumnProps(42)=   "Column(5).Order=6"
               Splits(0)._ColumnProps(43)=   "Column(6).Width=2725"
               Splits(0)._ColumnProps(44)=   "Column(6).DividerColor=0"
               Splits(0)._ColumnProps(45)=   "Column(6)._WidthInPix=2646"
               Splits(0)._ColumnProps(46)=   "Column(6)._EditAlways=0"
               Splits(0)._ColumnProps(47)=   "Column(6)._ColStyle=514"
               Splits(0)._ColumnProps(48)=   "Column(6).WrapText=1"
               Splits(0)._ColumnProps(49)=   "Column(6).Order=7"
               Splits(0)._ColumnProps(50)=   "Column(7).Width=3149"
               Splits(0)._ColumnProps(51)=   "Column(7).DividerColor=0"
               Splits(0)._ColumnProps(52)=   "Column(7)._WidthInPix=3069"
               Splits(0)._ColumnProps(53)=   "Column(7)._EditAlways=0"
               Splits(0)._ColumnProps(54)=   "Column(7)._ColStyle=514"
               Splits(0)._ColumnProps(55)=   "Column(7).WrapText=1"
               Splits(0)._ColumnProps(56)=   "Column(7).Order=8"
               Splits.Count    =   1
               PrintInfos(0)._StateFlags=   3
               PrintInfos(0).Name=   "piInternal 0"
               PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
               PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
               PrintInfos(0).PageHeaderHeight=   0
               PrintInfos(0).PageFooterHeight=   0
               PrintInfos.Count=   1
               AllowUpdate     =   0   'False
               BorderStyle     =   0
               ColumnFooters   =   -1  'True
               DataMode        =   4
               DefColWidth     =   0
               HeadLines       =   1,5
               FootLines       =   1
               MultipleLines   =   0
               CellTipsWidth   =   0
               InsertMode      =   0   'False
               DeadAreaBackColor=   16777215
               RowDividerColor =   12632256
               RowSubDividerColor=   12632256
               DirectionAfterEnter=   1
               MaxRows         =   250000
               ViewColumnCaptionWidth=   0
               ViewColumnWidth =   0
               _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
               _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
               _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
               _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
               _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=2,.bold=0,.fontsize=825,.italic=0"
               _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
               _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000005&,.bold=0,.fontsize=825"
               _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
               _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
               _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.bgcolor=&H8000000F&"
               _StyleDefs(11)  =   ":id=2,.fgcolor=&H80000007&,.bold=0,.fontsize=825,.italic=0,.underline=0"
               _StyleDefs(12)  =   ":id=2,.strikethrough=0,.charset=0"
               _StyleDefs(13)  =   ":id=2,.fontname=Tahoma"
               _StyleDefs(14)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
               _StyleDefs(15)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(16)  =   ":id=3,.fontname=Tahoma"
               _StyleDefs(17)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(18)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
               _StyleDefs(19)  =   "EditorStyle:id=7,.parent=1"
               _StyleDefs(20)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
               _StyleDefs(21)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
               _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
               _StyleDefs(23)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
               _StyleDefs(24)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
               _StyleDefs(25)  =   "Splits(0).Style:id=13,.parent=1,.wraptext=-1"
               _StyleDefs(26)  =   "Splits(0).CaptionStyle:id=22,.parent=4,.bold=0,.fontsize=825,.italic=0"
               _StyleDefs(27)  =   ":id=22,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(28)  =   ":id=22,.fontname=Tahoma"
               _StyleDefs(29)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bold=-1,.fontsize=825,.italic=0"
               _StyleDefs(30)  =   ":id=14,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(31)  =   ":id=14,.fontname=Verdana"
               _StyleDefs(32)  =   "Splits(0).FooterStyle:id=15,.parent=3"
               _StyleDefs(33)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
               _StyleDefs(34)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
               _StyleDefs(35)  =   "Splits(0).EditorStyle:id=17,.parent=7"
               _StyleDefs(36)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
               _StyleDefs(37)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
               _StyleDefs(38)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
               _StyleDefs(39)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
               _StyleDefs(40)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
               _StyleDefs(41)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=1,.bold=0,.fontsize=825"
               _StyleDefs(42)  =   ":id=28,.italic=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(43)  =   ":id=28,.fontname=Tahoma"
               _StyleDefs(44)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14,.bold=-1,.fontsize=825"
               _StyleDefs(45)  =   ":id=25,.italic=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(46)  =   ":id=25,.fontname=Tahoma"
               _StyleDefs(47)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
               _StyleDefs(48)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
               _StyleDefs(49)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=0"
               _StyleDefs(50)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
               _StyleDefs(51)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
               _StyleDefs(52)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
               _StyleDefs(53)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=0"
               _StyleDefs(54)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
               _StyleDefs(55)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
               _StyleDefs(56)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
               _StyleDefs(57)  =   "Splits(0).Columns(3).Style:id=66,.parent=13,.alignment=1"
               _StyleDefs(58)  =   "Splits(0).Columns(3).HeadingStyle:id=63,.parent=14"
               _StyleDefs(59)  =   "Splits(0).Columns(3).FooterStyle:id=64,.parent=15,.alignment=1"
               _StyleDefs(60)  =   "Splits(0).Columns(3).EditorStyle:id=65,.parent=17"
               _StyleDefs(61)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.alignment=2"
               _StyleDefs(62)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
               _StyleDefs(63)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
               _StyleDefs(64)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
               _StyleDefs(65)  =   "Splits(0).Columns(5).Style:id=50,.parent=13,.alignment=1"
               _StyleDefs(66)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14"
               _StyleDefs(67)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
               _StyleDefs(68)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
               _StyleDefs(69)  =   "Splits(0).Columns(6).Style:id=58,.parent=13,.alignment=1"
               _StyleDefs(70)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
               _StyleDefs(71)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
               _StyleDefs(72)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
               _StyleDefs(73)  =   "Splits(0).Columns(7).Style:id=62,.parent=13,.alignment=1"
               _StyleDefs(74)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14"
               _StyleDefs(75)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
               _StyleDefs(76)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
               _StyleDefs(77)  =   "Named:id=33:Normal"
               _StyleDefs(78)  =   ":id=33,.parent=0"
               _StyleDefs(79)  =   "Named:id=34:Heading"
               _StyleDefs(80)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(81)  =   ":id=34,.wraptext=-1"
               _StyleDefs(82)  =   "Named:id=35:Footing"
               _StyleDefs(83)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(84)  =   "Named:id=36:Selected"
               _StyleDefs(85)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(86)  =   "Named:id=37:Caption"
               _StyleDefs(87)  =   ":id=37,.parent=34,.alignment=2"
               _StyleDefs(88)  =   "Named:id=38:HighlightRow"
               _StyleDefs(89)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(90)  =   "Named:id=39:EvenRow"
               _StyleDefs(91)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
               _StyleDefs(92)  =   "Named:id=40:OddRow"
               _StyleDefs(93)  =   ":id=40,.parent=33"
               _StyleDefs(94)  =   "Named:id=41:RecordSelector"
               _StyleDefs(95)  =   ":id=41,.parent=34"
               _StyleDefs(96)  =   "Named:id=42:FilterBar"
               _StyleDefs(97)  =   ":id=42,.parent=33"
            End
            Begin VB.Label Label9 
               Caption         =   "F8 - Atur Limit Pencarian"
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
               Left            =   7620
               TabIndex        =   86
               Top             =   4590
               Width           =   1785
            End
         End
      End
      Begin BiSAFramProject.BiSAFrame BiSAFrame3 
         Height          =   825
         Left            =   105
         Top             =   9495
         Width           =   19005
         _ExtentX        =   33523
         _ExtentY        =   1455
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
         Begin VB.Timer tmrSpecialOffer 
            Left            =   10065
            Top             =   0
         End
         Begin BiSAFramProject.BiSAFrame BiSAFrame6 
            Height          =   585
            Left            =   6285
            Top             =   90
            Visible         =   0   'False
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   1032
            BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderStyle     =   4
            BackColor       =   -2147483633
            Begin VB.PictureBox pbTray 
               AutoSize        =   -1  'True
               Height          =   300
               Left            =   225
               Picture         =   "trPenjualan.frx":0B34
               ScaleHeight     =   240
               ScaleWidth      =   240
               TabIndex        =   75
               Top             =   150
               Visible         =   0   'False
               Width           =   300
            End
            Begin VB.OptionButton optPromo 
               Caption         =   "&Reguler"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Index           =   1
               Left            =   1650
               TabIndex        =   76
               Top             =   150
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.OptionButton optPromo 
               Caption         =   "&Promo"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Index           =   0
               Left            =   615
               TabIndex        =   77
               Top             =   165
               Visible         =   0   'False
               Width           =   975
            End
            Begin MSComDlg.CommonDialog CommonDialog1 
               Left            =   3930
               Top             =   75
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin BiSATextBoxProject.BiSATextBox cNamaKatalog 
               Height          =   345
               Left            =   30
               TabIndex        =   79
               TabStop         =   0   'False
               Top             =   615
               Visible         =   0   'False
               Width           =   960
               _ExtentX        =   1693
               _ExtentY        =   609
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
               BackColor       =   -2147483633
               Enabled         =   0   'False
               Appearance      =   0
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
            Begin BiSATextBoxProject.BiSABrowse cMasterKatalog 
               Height          =   330
               Left            =   2400
               TabIndex        =   78
               Top             =   195
               Visible         =   0   'False
               Width           =   975
               _ExtentX        =   1720
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
               Button          =   -1  'True
               CaptionWidth    =   1400
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
         Begin BiSAButtonProject.BiSAButton cmdHapus 
            Height          =   435
            Left            =   4305
            TabIndex        =   73
            Top             =   180
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   767
            Caption         =   "    &Delete"
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
            Picture         =   "trPenjualan.frx":10BE
         End
         Begin BiSAButtonProject.BiSAButton cmdAktivasi 
            Height          =   435
            Left            =   13890
            TabIndex        =   81
            Top             =   195
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   767
            Caption         =   ""
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
            Picture         =   "trPenjualan.frx":1348
         End
         Begin BiSAButtonProject.BiSAButton cmdEdit 
            Height          =   435
            Left            =   3225
            TabIndex        =   72
            Top             =   180
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   767
            Caption         =   "  &Edit"
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
            Picture         =   "trPenjualan.frx":14E7
         End
         Begin BiSAButtonProject.BiSAButton cmdAdd 
            Height          =   435
            Left            =   165
            TabIndex        =   71
            Top             =   180
            Width           =   3030
            _ExtentX        =   5345
            _ExtentY        =   767
            Caption         =   "  &Add [F1 = TUNAI/F2 = BON]"
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
            Picture         =   "trPenjualan.frx":1613
         End
         Begin BiSAButtonProject.BiSAButton cmdKeluar 
            Cancel          =   -1  'True
            Height          =   435
            Left            =   17790
            TabIndex        =   84
            Top             =   195
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   767
            Caption         =   "     &Exit"
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
            Picture         =   "trPenjualan.frx":17BE
         End
         Begin BiSAButtonProject.BiSAButton cmdSimpan 
            Height          =   435
            Left            =   14370
            TabIndex        =   82
            Top             =   195
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   767
            Caption         =   "    &Save [F4]"
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
            Picture         =   "trPenjualan.frx":1864
         End
         Begin BiSANumberBoxProject.BiSANumberBox nPoinReguler 
            Height          =   435
            Left            =   10875
            TabIndex        =   80
            Top             =   195
            Width           =   2985
            _ExtentX        =   5265
            _ExtentY        =   767
            Appearance      =   0
            Decimals        =   0
            Enabled         =   0   'False
            BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   0
            ForeColor       =   16777215
            Caption         =   " POIN HADIAH"
            CaptionWidth    =   1500
            BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin BiSAButtonProject.BiSAButton cmdPending 
            Height          =   435
            Left            =   16065
            TabIndex        =   83
            TabStop         =   0   'False
            Top             =   195
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   767
            Caption         =   "    &Pending [F5]"
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
            Picture         =   "trPenjualan.frx":1AEA
         End
         Begin BiSAButtonProject.BiSAButton cmdExport 
            Height          =   435
            Left            =   5460
            TabIndex        =   74
            Top             =   180
            Visible         =   0   'False
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   767
            Caption         =   ""
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
            Picture         =   "trPenjualan.frx":1D70
         End
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   195
         TabIndex        =   2
         Top             =   405
         Width           =   18945
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Label7"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6765
         TabIndex        =   4
         Top             =   1095
         Width           =   12345
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Label8"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   165
         TabIndex        =   3
         Top             =   1095
         Width           =   6570
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Program Support : 081999962828"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   14760
         TabIndex        =   1
         Top             =   105
         Visible         =   0   'False
         Width           =   4395
      End
   End
End
Attribute VB_Name = "trPenjualan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private blnSpecialOffer As Boolean

Public cNoOrder As String
'Public nKasirTotal As Double
Public nKasirKembalian As Double
Public lSign As Single
Public nKasirVoucher As Double
Public nKasirBayar As Double
Public nKasirTotalKartu As Double
Public nKasirKodeKartu As Integer
Public nKasirFeeKartu As Double
Public nKasirFeeTotalKartu As Double
Public nKasirNoTraceKartu As String
Public nKasirNoKartu As String
Public nKasirNamaDiKartu As String
Public nDPKasir As Double
Public nKasirKeteranganKartu As String
Public vaVoucher As New XArrayDB
Public lModeCompact As Boolean


Dim nPos As SisPos
Dim lEdit As Boolean
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB
Dim vaExport As New XArrayDB
Dim cKode As String
Dim cJenis  As String
Dim nSaldoStock As Double
Dim cTelp As String
Dim nBValue As Double
Dim nStockSelected As Double

'Dim Excel As Excel.Application
'Dim ExcelWBk As Excel.Workbook
'Dim ExcelWS As Excel.Worksheet
Dim objMenu As New CodeSuiteLibrary.Menu


'Private Sub StartExcel()
'  On Error GoTo err:
'  Set Excel = GetObject(, "Excel.Application")
'  Exit Sub
'err:
'  Set Excel = CreateObject("Excel.Application")
'End Sub

'Private Sub CloseWorkSheet()
'  ExcelWBk.Close
'  Excel.Quit
'End Sub

'Private Sub FinishExcel()
'  'Jangan lupa, selalu bersihkan memory saat mengakhiri
'  If Not ExcelWS Is Nothing Then Set ExcelWS = Nothing
'  If Not ExcelWBk Is Nothing Then Set ExcelWBk = Nothing
'  If Not Excel Is Nothing Then Set Excel = Nothing
'End Sub

Private Sub GetFakturBrowse(ByVal lStat As Boolean)
  cFaktur.Button = lStat
End Sub

Private Sub BiSABrowse1_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "supplier", "kodesupplier,nama", "kodesupplier", sisContent, cNamaSupplier.Text, " or nama like '%" & cNamaSupplier.Text & "'")
  If Not dbData.EOF Then
    cNamaSupplier.Text = cNamaSupplier.Browse(dbData, Array("KODE", "NAMA"), , Array(10, 35))
    cKodeSupplier.Text = GetNull(dbData!kodesupplier)
    cNamaSupplier.Text = GetNull(dbData!nama)
  End If
End Sub

Private Sub BiSAButton1_Click()
Dim nTmp As Integer

  SendKeysA vbKeyReturn, True
  TambahkanDariRekapPromo
  SendKeysA vbKeyReturn, True
End Sub

Private Sub BiSAButton2_Click()
  GetClearGrid
  BiSAButton1.Caption = "+"
End Sub

Private Sub getDataPending()
Dim db As New ADODB.Recordset
Dim n, nQtyTmp As Single


    If objMenu.UserLevel <> 0 Then
      Set db = objData.Browse(GetDSN, "pendingtrans p", "s.barcode,p.kodestock,s.Nama,p.qty,p.kodesatuan,p.harga,p.discount,p.jumlah,p.bv", "userid", sisAssign, GetRegistry(reg_Username), , "p.urutfaktur asc", Array("Left join stock s on s.kodestock = p.kodestock"))
    Else
      Set db = objData.Browse(GetDSN, "pendingtrans p", "s.barcode,p.kodestock,s.Nama,p.qty,p.kodesatuan,p.harga,p.discount,p.jumlah,p.bv", , , , , "p.urutfaktur asc", Array("Left join stock s on s.kodestock = p.kodestock"))
    End If
    
    If Not db.EOF Then
      vaArray.ReDim 0, -1, 0, 11
      Do While Not db.EOF
        vaArray.InsertRows vaArray.UpperBound(1) + 1
        n = vaArray.UpperBound(1)
        vaArray(n, 0) = n + 1
        vaArray(n, 1) = GetNull(db!barcode)
        vaArray(n, 2) = GetNull(db!nama)
        vaArray(n, 3) = GetNull(db!qty)
        vaArray(n, 4) = GetNull(db!kodesatuan)
        vaArray(n, 5) = GetNull(db!Harga)
        vaArray(n, 6) = GetNull(db!Discount)
        vaArray(n, 7) = GetNull(db!jumlah)
        vaArray(n, 8) = GetNull(db!KodeStock)
        vaArray(n, 10) = GetNull(db!bv)
        nQtyTmp = nQtyTmp + vaArray(n, 3)
        db.MoveNext
      Loop
      Set TDBGrid1.Array = vaArray
      TDBGrid1.ReBind
      TDBGrid1.Refresh
      TDBGrid1.Columns(3).FooterText = Format(nQtyTmp, "###,###,##0")
      Me.Refresh
      nNomor.value = vaArray.UpperBound(1) + 2
      nPoinReguler.value = GetHitungPoinHadiah(aCfg(objData, msKelipatan))
      GetUpdateTotal
    Else
      MsgBox "Maaf tidak ada data Pending yg bisa dibuka", vbExclamation
    End If
End Sub

Private Sub cAkunKas_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "akun", "kodeakun,keterangan", "jenis", sisAssign, "D", , "kodeakun")
  If Not dbData.EOF Then
    cAkunKas.Text = cAkunKas.Browse(dbData)
  End If
End Sub

Private Sub cBarcode_ButtonClick()
Dim kdestock As String
Dim cSQL As String
Dim cValidate As String
Dim cWhere As String
    
    cKode = ""
    If Trim(GetRegistry(reg_KodeGroupPenjualan)) <> "" Then
      cWhere = " and s.groupsales='" & GetRegistry(reg_KodeGroupPenjualan) & "'"
    End If
    'Set dbdata = objData.Browse(GetDSN, "stock s", "s.barcode,s.nama,format(s.hargajual,2) as hargajual,s.kodestock,s.kodesatuan,s.jenis,s.diskonpenjualan,s.bv", "s.barcode", sisContent, cBarcode.Text, " AND s.statusnonaktif <> 1", , , 0, 15)
    Set dbData = objData.Browse(GetDSN, "stock s", "s.barcode,s.nama,s.hargajual,s.kodesatuan", "s.barcode", sisContent, cBarcode.Text, " AND (s.statusnonaktif <> 1 and s.hargajual >0) " & cWhere, , , 0, 15)
  '  cSQL = "select barcode,format(hargajual,0) as hargajual, kodesatuan from stock where barcode like '%" & cBarcode.Text & "%' and statusnonaktif <> 1 limit 0,10"
  '  Set dbData = objData.SQL(GetDSN, cSQL)
    If Not dbData.EOF Then
      cBarcode.Text = cBarcode.Browse(dbData, Array("BARCODE", "NAMA", "JUAL", "SATUAN"), , Array(13, 35, 10, 8))
    Else
      cBarcode.Default
    End If
End Sub

Private Sub GetNewDataStock(ByVal kdBarcode As String)
Dim dbData As New ADODB.Recordset
Dim db As New ADODB.Recordset
Dim nCekHarga As Double

  Set dbData = objData.Browse(GetDSN, "stock s", "s.barcode,s.nama,format(s.hargajual,2) as hargajual,s.kodestock,s.kodesatuan,s.jenis,s.diskonpenjualan,s.bv", "s.barcode", sisAssign, kdBarcode, " AND s.statusnonaktif <> 1", , , 0, 10)
  If Not dbData.EOF Then
    cKode = GetNull(dbData!KodeStock, "")
    nBValue = GetNull(dbData!bv, "")
    cNama.Text = GetNull(dbData!nama, "")
    cSatuan.Text = GetNull(dbData!kodesatuan, "")
    nDisc1.value = GetNull(dbData!diskonpenjualan)
    
    If aCfg(objData, msCHKdiscountPenjualan) = 1 Then
      nDisc1.value = aCfg(objData, msDiscountPenjualan)
    End If
    
    'tentukan harga jual sesuai dengan konfigurasi yg telah di setup
    If aCfg(objData, msHargaPenjualanNonTunai) = "3" Then
      nHarga.value = GetHargaKontrak(objData, cCustomer.Text, cKode)
    ElseIf aCfg(objData, msHargaPenjualanNonTunai) = "2" Then
      nHarga.value = GetHargaJualLastByCustomer(objData, cKode, cCustomer.Text)
    Else
      nHarga.value = GetNull(dbData!HargaJual)
    End If
    
    'Lakukan markup harga jika non member
    nHarga.value = MarkUpHarga(objData, cCustomer.Text, nHarga.value)
    cJenis = GetNull(dbData!jenis)
    
    'jika di master customer tersetup diskon, maka abaikan semuanya
    Set dbData = objData.Browse(GetDSN, "anggota", "diskon", "kodeanggota", sisAssign, cCustomer.Text)
    If Not dbData.EOF Then
      If GetNull(dbData!diskon) <> 0 Then
        nDisc1.value = GetNull(dbData!diskon)
      End If
    End If
  End If
End Sub

'Private Function GetHargaBarang(ByVal obj As CodeSuiteLibrary.Data, ByVal Barcode As String, ByVal nHargaBatas) As Double
'  GetHargaBarang = 0
'  Set dbData = obj.Browse(GetDSN, "stock", "hargajual", "barcode", sisAssign, Barcode)
'  If Not dbData.EOF Then
'    If GetNull(dbData!hargajual) > nHargaBatas Then
'      MsgBox "HELLOO. BARANG INI HARGA NYA " & Format(GetNull(dbData!hargajual), "###,###,###,##00") & vbCrLf & " BENER MAU DISIMPAN.. SERIUS "
'    End If
'  End If
'End Function

Private Sub SumBayar()
  nPiutang.value = nTotal.value - IIf(nTunai.value > nTotal.value, nTotal.value, nTunai.value)
End Sub

Private Sub cBarcode_LostFocus()
Dim kdestock As String

  If Trim(cBarcode.Text) <> "" Then
    Set dbData = objData.Browse(GetDSN, "stock s", "s.kodestock,s.barcode,s.nama,s.kodesatuan,s.hargajual,s.jenis,s.diskonpenjualan,s.bv,s.stok,s.kategori", "s.barcode", sisAssign, cBarcode.Text, " AND s.statusnonaktif <> 1 and s.hargajual >0 and s.groupsales = '" & GetRegistry(reg_KodeGroupPenjualan) & "'")
    If Not dbData.EOF Then
      'cBarcode.Text = cBarcode.Browse(dbData)
      kdestock = GetNull(dbData!KodeStock)
      GetDataStock
      SumJumlah
      cmdOK_Click
    Else
      MsgBox "Maaf, data tidak ada - atau Harga jual belum di set atau Status barang sudah non aktif", vbCritical
      cBarcode.Default
      cBarcode.SetFocus
    End If
  End If
End Sub

Public Function validateInput(ByRef cInputCek As String, ByVal cChr As String) As Boolean
  validateInput = True
  If InStr(1, cInputCek, cChr) <= 0 Then
    validateInput = False
    cInputCek = ""
  End If
End Function

Private Sub cFaktur_ButtonClick()
Dim db As New ADODB.Recordset
Dim n As Integer
Dim lSave As Boolean
Dim nQtyTmp As Single
Dim cSQL As String
Dim cFilterUsername As String

  If aCfg(objData, msOtorisasiPenuh) = "Y" Then
    If GetRegistry(reg_UserLevel) <> 0 Then
      If objMenu.GetPassword("", Me, GetDSN) Then
        If objMenu.UserLevel <> 0 Then
            MsgBox "Maaf, Anda tidak diberikan wewenang untuk melakukan PENGEDITAN." & vbCrLf & _
                   "Hanya user dengan LEVEL 0 (SUPERVISOR) yg diijinkan", vbInformation, "OTORISASI not ALLOWED"
            Exit Sub
'        Else
'          MsgBox "OTORISASI DIBATALKAN", vbCritical
'          Exit Sub
        End If
      Else
        Exit Sub
      End If
    End If
  End If

  lSave = True
  nQtyTmp = 0

  'jika level 0 boleh edit semua
  'jika tidak maka yg boleh edit adalah sesuai dengan user login nya
  cFilterUsername = ""
  If objMenu.UserLevel <> 0 Then
    cFilterUsername = " and username = '" & GetRegistry(reg_Username) & "'"
  End If
  
  If nPos = Edit Then
    If Trim(cCustomer.Text) = "" Then
      MsgBox "Input dulu bos, kode/nama customer nya hehe", vbExclamation
    End If
    Set db = objData.Browse(GetDSN, "totpenjualan", "nomorpenjualan,tgl,subtotal,total,tunai,piutang,voucher", "nomorpenjualan", sisContent, cFaktur.Text, " and tgl = '" & Format(dTgl.value, "yyyy-MM-dd") & "' and kodeanggota = '" & cCustomer.Text & "'" & cFilterUsername & " and (flaglunas=0 or tunai<>0) and kodegroupsales='" & GetRegistry(reg_KodeGroupPenjualan) & "' and voucher = 0")
  End If
  
  If nPos = Delete Then
    If Trim(cCustomer.Text) = "" Then
      MsgBox "Input dulu bos, kode/nama customer nya hehe", vbExclamation
    End If
    Set db = objData.Browse(GetDSN, "totpenjualan", "nomorpenjualan,tgl,subtotal,total,tunai,piutang,voucher", "nomorpenjualan", sisContent, cFaktur.Text, " and tgl = '" & Format(dTgl.value, "yyyy-MM-dd") & "' and kodeanggota = '" & cCustomer.Text & "'" & cFilterUsername & " and (flaglunas=0 or tunai<>0) and kodegroupsales='" & GetRegistry(reg_KodeGroupPenjualan) & "'")
  End If
  
  If Not db.EOF Then
  
    cFaktur.Text = cFaktur.Browse(db, Array("NOMOR", "TGL", "SUBTOTAL", "TOTAL", "TUNAI", "PIUTANG", "VOUCHER"), , Array(16, 10, 10, 10, 10, 10, 10))
    
    'ambil nilai total
    Set db = objData.Browse(GetDSN, "totpenjualan t", "t.*,g.keterangan as namagudang", "t.nomorpenjualan", sisAssign, cFaktur.Text, , , Array("left join gudang g on g.kodegudang = t.kodegudang"))
    If Not db.EOF Then
      
      If GetNull(db!Tunai) <> 0 Then
        If GetRegistry(reg_UserLevel) <> 0 Then
          If objMenu.GetPassword("", Me, GetDSN) Then
            If objMenu.UserLevel <> 0 Then
              MsgBox "Maaf, Anda tidak diberikan wewenang untuk melakukan PENGEDITAN." & vbCrLf & _
                     "Hanya user dengan LEVEL 0 (SUPERVISOR) yg diijinkan", vbInformation, "OTORISASI not ALLOWED"
              GetFakturBrowse False
              cmdKeluar_Click
              Exit Sub
            End If
          Else
            MsgBox "OTORISASI DIBATALKAN", vbCritical
            GetFakturBrowse False
            cmdKeluar_Click
            Exit Sub
          End If
        End If
      End If
      
      cFakturAsli.Text = GetNull(db!fakturasli, "")
      If GetNull(db!jthtmp) = Format(Date, "yyyy-MM-dd") Then
        dJthTmp.value = Format(DateAdd("D", 7, Format(Date, "yyyy-MM-dd")), "yyyy-MM-dd")
      Else
        dJthTmp.value = GetNull(db!jthtmp)
      End If
      
      nPersDisc.value = GetNull(db!PersDisc, 0)
      nPPn.value = GetNull(db!ppn, 0)
      nSubTotal.value = GetNull(db!Subtotal, 0)
      nDiscount.value = GetNull(db!Discount, 0)
      nPajak.value = GetNull(db!PAJAK, 0)
      nTotal.value = GetNull(db!Total, 0)
      nTunai.value = GetNull(db!Tunai, 0)
      nPiutang.value = GetNull(db!Piutang, "")
      cAkunKas.Text = GetNull(db!kodeakun)
      cSalesman.Text = GetNull(db!kodesalesman, "")
      nKomisi.value = GetNull(db!komisi)
      nDP.value = GetNull(db!dp)
      cGudang.Text = GetNull(db!kodegudang)
      cNamaGudang.Text = GetNull(db!namagudang)
      cUp.Text = GetNull(db!upkepada)
      SetOpt optPromo, GetNull(db!jenis)
      cKeterangan.Text = GetNull(db!keterangan)
      nHargaCOD.value = GetNull(db!ongkir)
      If GetNull(db!Piutang) = 0 Then
        chkTunai.value = 1
      Else
        chkTunai.value = 0
      End If
    End If
    
    'ambil nilai detail
    Set db = objData.Browse(GetDSN, "penjualan p", "s.barcode,p.kodestock,s.Nama,p.qty,p.kodesatuan,p.harga,p.discount,p.jumlah,p.bv,s.asbiaya,s.autobiaya", "nomorpenjualan", sisAssign, cFaktur.Text, , "p.urutfaktur asc", Array("Left join stock s on s.kodestock = p.kodestock"))
    If Not db.EOF Then
      vaArray.ReDim 0, -1, 0, 11
      Do While Not db.EOF
        vaArray.InsertRows vaArray.UpperBound(1) + 1
        n = vaArray.UpperBound(1)
        vaArray(n, 0) = n + 1
        vaArray(n, 1) = GetNull(db!barcode)
        vaArray(n, 2) = GetNull(db!nama)
        vaArray(n, 3) = GetNull(db!qty)
        vaArray(n, 4) = GetNull(db!kodesatuan)
        vaArray(n, 5) = GetNull(db!Harga)
        vaArray(n, 6) = GetNull(db!Discount)
        vaArray(n, 7) = GetNull(db!jumlah)
        vaArray(n, 8) = GetNull(db!KodeStock)
        vaArray(n, 10) = GetNull(db!bv)
        nQtyTmp = nQtyTmp + vaArray(n, 3)
        db.MoveNext
      Loop
      Set TDBGrid1.Array = vaArray
      TDBGrid1.ReBind
      TDBGrid1.Refresh
      TDBGrid1.Columns(3).FooterText = Format(nQtyTmp, "###,###,##0")
      Me.Refresh
      nNomor.value = vaArray.UpperBound(1) + 2
      nPoinReguler.value = GetHitungPoinHadiah(aCfg(objData, msKelipatan))
    End If
    
'    If nPos = Delete Then
'      If MsgBox("Data akan dihapus?", vbYesNo) = vbYes Then
'        objData.Start GetDSN
'
'        'Patch
'        'Dim cSQL As String
'        cSQL = ""
'        cSQL = " select distinct(nomorpelunasanpiutang) as nomorpelunasanpiutang from pelunasanpiutang where nomorpenjualan = '" & cFaktur.Text & "'"
'        Set db = objData.Sql(GetDSN, cSQL)
'        If Not db.EOF Then
'          If MsgBox("Transaksi ini sudah pernah dilunasi sebelumnya!" & vbCrLf & "Dengan menghapus berarti seluruh data pelunasan yg berkenaan dengan transaksi ini akan ikut terhapus juga" & vbCrLf & "Apakah anda yakin akan menghapus?", vbYesNo) = vbYes Then
'            'tambahkan penghapusan di topup
'            'rutin menghapus pada modul pelunasan piutang
'
'            lSave = IIf(lSave, DelKodeTr(objData, msPelunasanPiutang, GetNull(db!nomorpelunasanpiutang)), False)
'            lSave = IIf(lSave, objData.Delete(GetDSN, "pelunasanpiutang", "nomorpelunasanpiutang", sisAssign, GetNull(db!nomorpelunasanpiutang)), False)
'            lSave = IIf(lSave, objData.Delete(GetDSN, "kartupiutang", "nomorkartupiutang", sisAssign, GetNull(db!nomorpelunasanpiutang)), False)
'            lSave = IIf(lSave, objData.Delete(GetDSN, "totpelunasanpiutang", "nomorpelunasanpiutang", sisAssign, GetNull(db!nomorpelunasanpiutang)), False)
'            lSave = IIf(lSave, objData.Delete(GetDSN, "bukubesar", "faktur", sisAssign, GetNull(db!nomorpelunasanpiutang)), False)
'          Else
'            MsgBox "Penghapusan Dibatalkan"
'            GetEdit False
'            initvalue
'            objData.Cancel GetDSN
'            Exit Sub
'          End If
'        End If
'
'        cSQL = ""
'        cSQL = " select * from totrtnpenjualan where nomorpenjualan = '" & cFaktur.Text & "'"
'        Set db = objData.Sql(GetDSN, cSQL)
'        If Not db.EOF Then
'          If MsgBox("Transaksi ini masih dirujuk oleh retur penjualan!" & vbCrLf & "Dengan menghapus berarti seluruh rujukan pada retur penjualan akan ikut dihapus pula" & vbCrLf & "Apakah anda yakin akan menghapus?", vbYesNo) = vbYes Then
'            Do While Not db.EOF
'              lSave = IIf(lSave, objData.Edit(GetDSN, "totrtnpenjualan", "nomorreturpenjualan = '" & GetNull(db!nomorreturpenjualan) & "'", Array("nomorpenjualan"), Array("")), False)
'              db.MoveNext
'            Loop
'          Else
'            MsgBox "Penghapusan dibatalkan"
'            GetEdit False
'            initvalue
'            objData.Cancel GetDSN
'            Exit Sub
'          End If
'        End If
'
'        'end patch
'
'        'Update dulu ke table order
'        lSave = IIf(lSave, objData.Edit(GetDSN, "totmemberorder", "nomorpenjualan = '" & cFaktur.Text & "'", Array("nomorpenjualan", "status"), Array("", 0)), False)
'
'        'Rutin menghapus transaksi penjualan
'        lSave = IIf(lSave, objData.Delete(GetDSN, "orderan", "reffid", sisAssign, cFaktur.Text), False)
'        lSave = IIf(lSave, DelKodeTr(objData, msPenjualan, cFaktur.Text), False)
'        lSave = IIf(lSave, objData.Delete(GetDSN, "penjualan", "nomorpenjualan", sisAssign, cFaktur.Text), False)
'        lSave = IIf(lSave, objData.Delete(GetDSN, "kartupiutang", "nomorkartupiutang", sisAssign, cFaktur.Text), False)
'        lSave = IIf(lSave, objData.Delete(GetDSN, "kartustock", "nomor", sisAssign, cFaktur.Text), False)
'        lSave = IIf(lSave, objData.Delete(GetDSN, "totpenjualan", "nomorpenjualan", sisAssign, cFaktur.Text), False)
'
'        If lSave Then
'          objData.Save GetDSN
'        Else
'          objData.Cancel GetDSN
'        End If
'
'      End If
'      initvalue
'      GetEdit False
'    End If
    
    If nPos = Edit Or Delete Then
      'cek kalau nota ini sudah pernah ditukarkan hadiah (poin)
      
      cSQL = ""
      cSQL = "SELECT * FROM tukarpoin WHERE faktur = '" & cFaktur.Text & "'"
      Set db = objData.SQL(GetDSN, cSQL)
      If Not db.EOF Then
        MsgBox "Maaf Faktur ini sudah ditukarkan dengan hadiah POIN, tidak bisa dikoreksi"
        GetEdit False
        Exit Sub
      End If
      
      'cek kalau nota ini sudah ditukar, maka tidak bisa dihapus/diedit
      cSQL = ""
      cSQL = "select * from poinhadiah where faktur = '" & cFaktur.Text & "'"
      Set db = objData.SQL(GetDSN, cSQL)
      If Not db.EOF Then
        If GetNull(db!tukar) > 0 Then
          MsgBox "Maaf, pembelanjaan ini sudah ditukar poin nya. Tidak bisa dihapus/diedit lagi"
          initvalue
          GetEdit False
          Exit Sub
        End If
      End If
      
      'Patch
      'Cek juga jikalau nota ini sudah selesai di lunasi, maka tidak bisa lagi di edit atau di hapus
      cSQL = ""
      cSQL = " select distinct(nomorpelunasanpiutang) as nomorpelunasanpiutang from pelunasanpiutang where nomorpenjualan = '" & cFaktur.Text & "'"
      Set db = objData.SQL(GetDSN, cSQL)
      If Not db.EOF Then
        MsgBox "MAAF.. " & vbCrLf & "Data sudah pernah dilunasi, TIDAK BISA DIEDIT ATAU DIHAPUS", vbExclamation
        GetEdit False
        initvalue
        objData.Cancel GetDSN
        Exit Sub
      Else
        If nPos = Delete Then
'         Proses rutin penghapusan
          If MsgBox("Data akan dihapus?", vbYesNo) = vbYes Then
            
            objData.Start GetDSN
            cSQL = ""
            cSQL = " select * from totrtnpenjualan where nomorpenjualan = '" & cFaktur.Text & "'"
            Set db = objData.SQL(GetDSN, cSQL)
            If Not db.EOF Then
                MsgBox "Penghapusan dibatalkan, Transaksi ini masih dirujuk oleh retur penjualan!"
                GetEdit False
                initvalue
                objData.Cancel GetDSN
                Exit Sub
            End If
  '         Update dulu ke table order
            lSave = IIf(lSave, objData.Edit(GetDSN, "totmemberorder", "nomorpenjualan = '" & cFaktur.Text & "'", Array("nomorpenjualan", "status"), Array("", 0)), False)
    
  '         Rutin menghapus transaksi penjualan
            lSave = IIf(lSave, objData.Delete(GetDSN, "orderan", "reffid", sisAssign, cFaktur.Text), False)
            lSave = IIf(lSave, DelKodeTr(objData, msPenjualan, cFaktur.Text), False)
            lSave = IIf(lSave, objData.Delete(GetDSN, "penjualan", "nomorpenjualan", sisAssign, cFaktur.Text), False)
            lSave = IIf(lSave, objData.Delete(GetDSN, "kartupiutang", "nomorkartupiutang", sisAssign, cFaktur.Text), False)
            lSave = IIf(lSave, objData.Delete(GetDSN, "kartustock", "nomor", sisAssign, cFaktur.Text), False)
            lSave = IIf(lSave, objData.Delete(GetDSN, "totpenjualan", "nomorpenjualan", sisAssign, cFaktur.Text), False)
            lSave = IIf(lSave, objData.Delete(GetDSN, "poinhadiah", "faktur", sisAssign, cFaktur.Text), False)
            lSave = IIf(lSave, objData.Delete(GetDSN, "membertopup", "nomormembertopup ", sisAssign, cFaktur.Text), False)
            lSave = IIf(lSave, objData.Delete(GetDSN, "membertopup", "nomormembertopup ", sisAssign, cFaktur.Text), False)
            lSave = IIf(lSave, objData.Delete(GetDSN, "penjualankonsinyasi", "nomorpenjualan", sisAssign, cFaktur.Text), False)

            'hapus jika pernah menggunakan kartu
            lSave = IIf(lSave, objData.Delete(GetDSN, "trkartu", "nomorpenjualan", sisAssign, cFaktur.Text), False)
            'update kembali penggunaan voucher,jika ada
            cSQL = ""
            cSQL = "select * from membertopup where lstatusid = '" & cFaktur.Text & "'"
            Set dbData = objData.SQL(GetDSN, cSQL)
            If Not dbData.EOF Then
              Do While Not dbData.EOF
                lSave = IIf(lSave, objData.Edit(GetDSN, "membertopup", "nomormembertopup = '" & GetNull(dbData!nomormembertopup) & "'", Array("lstatus", "lstatusid"), Array(sisFlag.Nul, "")), False)
                dbData.MoveNext
              Loop
            End If
            
            If lSave Then
              objData.Save GetDSN
              'hapus table lock
            Else
              objData.Cancel GetDSN
              MsgBox "Err. Data gagal dihapus", vbCritical
              'hapus table lock
            End If
          End If
    
          initvalue
          GetEdit False
        End If
      End If
    End If
    If nPos = Edit Then
      SendKeysA vbKeyReturn, True
    End If
  End If
End Sub

Private Function isInLock(ByVal obj As CodeSuiteLibrary.Data, ByVal cNo As String) As Boolean
Dim d As Integer
Dim cSQL As String
  
  cSQL = "select idbukubesar from bukubesar where status = 20 ORDER BY idbukubesar desc LIMIT 0,1"
  
End Function

Private Sub cFaktur_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If nPos = Edit Then
      
      cFaktur.Enabled = False
      cFaktur.BackColor = vbButtonFace
    End If
End If
End Sub

Private Sub cFaktur_KeyPress(KeyAscii As Integer)
  If nPos = Edit Then
      KeyAscii = 0
      cFaktur.Enabled = False
      cFaktur.BackColor = vbButtonFace
  End If
End Sub

Private Sub cFaktur_Validate(Cancel As Boolean)
  If nPos = Edit Then
      cFaktur.Enabled = False
      cFaktur.BackColor = vbButtonFace
  End If
End Sub

Private Sub cGudang_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "gudang", "kodegudang,keterangan", "lstatus", sisAssign, "A")
  If Not dbData.EOF Then
    cGudang.Text = cGudang.Browse(dbData)
    cGudang.Text = GetNull(dbData!kodegudang)
    cNamaGudang.Text = GetNull(dbData!keterangan)
  End If
End Sub

Private Sub chkTunai_Click()
  If chkTunai.value = 1 Then

    nPiutang.value = 0
    nTunai.value = nTotal.value - nDP.value
  Else
    nPiutang.value = nTotal.value - nDP.value
    nTunai.value = 0
  End If
  If chkTunai.value = 1 Then
    Label1.Caption = "T U N A I"
  Else
    Label1.Caption = "B O N"
  End If
End Sub

Private Sub cmdAddOrder_Click()
Dim db As New ADODB.Recordset
Dim nJumlah1 As Double
Dim n As Integer

'  vaArray(n, 0) = nNomor.Value
'  vaArray(n, 1) = cBarcode.Text
'  vaArray(n, 2) = cNama.Text
'  vaArray(n, 3) = nQty.Value
'  vaArray(n, 4) = cSatuan.Text
'  vaArray(n, 5) = nHarga.Value
'  vaArray(n, 6) = nDisc1.Value
'  vaArray(n, 7) = nJumlah.Value
'  vaArray(n, 8) = cKode
'  vaArray(n, 9) = cJenis


  frmLoadOrder.cKodeMember = cCustomer.Text
  frmLoadOrder.Show vbModal
  
  If cNoOrder <> "" Then
    vaArray.ReDim 0, -1, 0, 11
    Set db = objData.Browse(GetDSN, "memberorder m", "m.*,s.nama as namastock,s.barcode,s.kodesatuan,s.jenis", "t.nomormemberorder", sisAssign, cNoOrder, , "m.nourut asc", Array("LEFT JOIN totmemberorder t on t.nomormemberorder = m.nomormemberorder", "LEFT JOIN stock s on s.kodestock = m.kodestock"))
    If Not db.EOF Then
      Do While Not db.EOF
        vaArray.InsertRows vaArray.UpperBound(1) + 1
        n = vaArray.UpperBound(1)
        vaArray(n, 0) = n + 1
        vaArray(n, 1) = GetNull(db!barcode)
        vaArray(n, 2) = GetNull(db!namastock)
        vaArray(n, 3) = GetNull(db!qty)
        vaArray(n, 4) = GetNull(db!kodesatuan)
        vaArray(n, 5) = GetNull(db!Harga)
        vaArray(n, 6) = GetNull(db!Discount)
        vaArray(n, 7) = GetNull(db!jumlah)
        vaArray(n, 8) = GetNull(db!KodeStock)
        vaArray(n, 9) = GetNull(db!jenis)
        db.MoveNext
      Loop
    End If
    
    TDBGrid1.Array = vaArray
    TDBGrid1.ReBind
    
    nJumlah1 = 0
    For n = 0 To vaArray.UpperBound(1)
      nJumlah1 = nJumlah1 + vaArray(n, 7)
    Next
    nSubTotal.value = nJumlah1
    
    SumTotal
    
    InitValue1
    
    nNomor.value = vaArray.UpperBound(1) + 2
    cBarcode.SetFocus
  End If
End Sub

Private Sub cmdAktivasi_Click()
  frmAktivasi.Action Me
End Sub

Private Sub GetDataStock()
Dim db As New ADODB.Recordset
Dim nCekHarga As Double
    
  If GetRegistry(reg_TampilNotifikasi) = 1 Then
    GetNotifikasiAdd "STOK " & GetNull(dbData!stok), GetNull(dbData!nama, "") & " ", IIf(GetNull(dbData!stok) = 0, 3, 0)
    GetNotifikasiRemove
  End If
  nStockSelected = GetNull(dbData!stok)
  Label5.Caption = "STOK : " & GetNull(dbData!nama) & " : " & GetNull(dbData!stok)
  nInfoStok.value = GetNull(dbData!stok)
  cBarcode.Text = GetNull(dbData!barcode, "")
  cKode = GetNull(dbData!KodeStock, "")
  nBValue = GetNull(dbData!bv, "")
  cNama.Text = GetNull(dbData!nama, "")
  cSatuan.Text = GetNull(dbData!kodesatuan, "")
  
  'set kelompok harga/markup dan diskon
  Dim nDiskonKategori As Double
  nHarga.value = GetNull(dbData!HargaJual)
  nDisc1.value = GetNull(dbData!diskonpenjualan)
  
  nDiskonKategori = GetDiskonPenjualanByKategori(objData, dbData!kategori)
  If nDiskonKategori <> 0 Then
    nDisc1.value = nDiskonKategori
  End If
  
  nInfoStok_Change
  
  If aCfg(objData, msCHKdiscountPenjualan) = 1 Then
    nDisc1.value = aCfg(objData, msDiscountPenjualan)
  End If
  
  'tentukan harga jual sesuai dengan konfigurasi yg telah di setup
  If aCfg(objData, msHargaPenjualanNonTunai) = "3" Then
    nHarga.value = GetHargaKontrak(objData, cCustomer.Text, cKode)
  ElseIf aCfg(objData, msHargaPenjualanNonTunai) = "2" Then
    nHarga.value = GetHargaJualLastByCustomer(objData, cKode, cCustomer.Text)
'  Else
    'cek di kelompok harga
    'nHarga.Value = GetNull(dbData!HargaJual)
  End If
  
  'Lakukan markup harga jika non member
  nHarga.value = MarkUpHarga(objData, cCustomer.Text, nHarga.value)
  cJenis = GetNull(dbData!jenis)
  
'  'jika di master customer tersetup diskon, maka abaikan semuanya
'  Set dbData = objData.Browse(GetDSN, "anggota")
'  If Not dbData.EOF Then
'    If GetNull(dbData!diskon) <> 0 Then
'      nDisc1.Value = GetNull(dbData!diskon)
'    End If
'  End If
  
  'jika di master customer tersetup diskon, maka abaikan semuanya
  Set dbData = objData.Browse(GetDSN, "anggota", "diskon", "kodeanggota", sisAssign, cCustomer.Text)
  If Not dbData.EOF Then
    If GetNull(dbData!diskon) <> 0 Then
      nDisc1.value = GetNull(dbData!diskon)
    End If
  End If
End Sub

Private Function GetDiskonPenjualanByKategori(ByVal obj As CodeSuiteLibrary.Data, ByVal cKategori As String) As Double
Dim db As New ADODB.Recordset

  Set db = obj.Browse(GetDSN, "kelompokharga", , "kategori", sisAssign, cKategori, " and kodegudang = '" & GetGudangUser(obj, GetRegistry(reg_Username)) & "'")
  If Not db.EOF Then
    GetDiskonPenjualanByKategori = GetNull(db!diskon)
  End If
End Function

Private Function GetHargaJualLastByCustomer(ByVal obj As CodeSuiteLibrary.Data, ByVal cStock As String, ByVal cCust As String) As Double
Dim db As New ADODB.Recordset

  Set db = obj.Browse(GetDSN, "penjualan p", "p.tgl,p.kodestock,p.harga", "p.kodestock", sisAssign, cStock, " and t.kodeanggota = '" & cCust & "'", "p.tgl desc", Array("left join totpenjualan t on t.nomorpenjualan = p.nomorpenjualan"), 0, 1)
  If Not db.EOF Then
    GetHargaJualLastByCustomer = GetNull(db!Harga)
  Else
    Set db = obj.Browse(GetDSN, "stock", , "kodestock", sisAssign, cStock)
    If Not db.EOF Then
      GetHargaJualLastByCustomer = GetNull(db!HargaJual)
    End If
  End If
End Function

Private Function GetHargaKontrak(ByVal obj As CodeSuiteLibrary.Data, ByVal cCustomer As String, ByVal cStock As String) As Double
Dim db As New ADODB.Recordset
  
  GetHargaKontrak = 0
  Set db = obj.Browse(GetDSN, "kontrakstock", , "kodeanggota", sisAssign, cCustomer, " and kodestock = '" & cStock & "'")
  If Not db.EOF Then
    GetHargaKontrak = GetNull(db!hargakontrak)
  Else
    Set db = obj.Browse(GetDSN, "stock", , "kodestock", sisAssign, cStock)
    If Not db.EOF Then
      GetHargaKontrak = GetNull(db!HargaJual)
    End If
  End If
End Function

Private Function MarkUpHarga(ByVal obj As CodeSuiteLibrary.Data, ByVal anggota As String, ByVal Harga As Double) As Double
Dim db As New ADODB.Recordset
  MarkUpHarga = Harga
  Set db = obj.Browse(GetDSN, "anggota", , "kodeanggota", sisAssign, anggota)
  If Not db.EOF Then
    If GetNull(db!Status) <> "A" Then
      MarkUpHarga = Harga + (aCfg(objData, msMarkUpHargaJual) * Harga / 100)
    End If
  End If
End Function

Private Sub cmdAdd_Click()
Dim i As Integer
Dim db As New ADODB.Recordset

  nPos = Add
  GetEdit True
  GetFakturBrowse False
  cFaktur.Text = CreateNomorFaktur(objData, sisModulTransaksi.Penjualan, "totpenjualan", "nomorpenjualan")
  If Trim(cFaktur.Text) <> "" Then
    cmdAddOrder.Enabled = True
    cAkunKas.Text = GetAkunKas(objData, GetRegistry(reg_Username))
    
    GetRegCustomer 'History
    
    If GetModePenjualanUser(objData, GetRegistry(reg_Username)) <> 0 Then
      ModeCompact True
      cBarcode.SetFocus
     'pastikan ada pelanggan dengan kode - dalam database. jika tidak ada batalkan mode compact
     Set db = objData.Browse(GetDSN, "anggota", "kodeanggota", "kodeanggota", sisAssign, cCustomer.Text)
     If db.EOF Then
      MsgBox "Maaf Mode Penjualan Compact tidak bisa di buka" & vbCrLf & "Setting terlebih dahulu satu orang Pelanggan dengan Kode Pelanggan - (Tanda Strip)", vbCritical
      Unload trPenjualan
     End If
    Else
        Select Case aCfg(objData, msDefaultModelPenjualan)
         Case "T"
           chkTunai.value = 1
           cBarcode.SetFocus
         Case "B"
           chkTunai.value = 0
           cNamaCustomer.SetFocus
       End Select
       cSalesman.Text = GetNull(GetRegistry(reg_KodeSalesman))
    End If
  Else
    GetEdit False
  End If
End Sub

Private Function GetLewatJatuhTempo(ByVal obj As CodeSuiteLibrary.Data, ByVal kodeanggota As String) As XArrayDB
Dim db As New ADODB.Recordset
Dim n As Integer
Dim nSisaPiutang As Double
Dim vaJatuhTmp As New XArrayDB

  vaJatuhTmp.ReDim 0, -1, 0, 3
  Set db = obj.Browse(GetDSN, "totpenjualan", "nomorpenjualan,tgl,piutang,jthtmp", "kodeanggota", sisAssign, kodeanggota, , "tgl desc")
  If Not db.EOF Then
    Do While Not db.EOF
      If Not isLunas(obj, GetNull(db!nomorpenjualan), nSisaPiutang) Then
        If Format(Date, "yyyy-MM-dd") > Format(GetNull(db!jthtmp), "yyyy-MM-dd") Then
          vaJatuhTmp.InsertRows vaJatuhTmp.UpperBound(1) + 1
          n = vaJatuhTmp.UpperBound(1)
          vaJatuhTmp(n, 0) = n + 1
          vaJatuhTmp(n, 1) = Format(GetNull(db!tgl), "dd-MM-yyyy")
          vaJatuhTmp(n, 2) = GetNull(db!nomorpenjualan)
          vaJatuhTmp(n, 3) = nSisaPiutang
        End If
      End If
      db.MoveNext
    Loop
  End If
  Set GetLewatJatuhTempo = vaJatuhTmp
End Function

Private Sub cmdEdit_Click()
If GetModePenjualanUser(objData, GetRegistry(reg_Username)) = 0 Then
  nPos = Edit
  GetEdit True
  GetFakturBrowse True
  cmdAddOrder.Enabled = False
  cNamaCustomer.SetFocus
Else
  MsgBox "Aktifkan dulu fungsi Full, User ini berada di Fungsi Compact", vbInformation, "Tidak bisa di edit"
End If
End Sub

Private Sub cmdExport_Click()

  If vaArray.UpperBound(1) > -1 Then 'Jika ada datanya maka silahkan di export
    If MsgBox("Export ke Excel?", vbYesNo + vbInformation) = vbYes Then
      Dim a As New exportExcel
      Dim na As Integer
      
  '        vaExport.ReDim 0, 0, 0, 1
          vaExport.ReDim 0, -1, 0, 3
  '        vaExport(0, 0) = "Balasan Order member " & cNamaCustomer.Text & " No: " & cFaktur.Text & " Tg. " & dTgl.Value
          
          For na = vaArray.LowerBound(1) To vaArray.UpperBound(1)
            vaExport.InsertRows na
            vaExport(na, 0) = vaArray(na, 1)
            vaExport(na, 1) = vaArray(na, 3)
            vaExport(na, 2) = vaArray(na, 5) 'export harga jual
            vaExport(na, 3) = vaArray(na, 6) 'export diskon jual
          Next na
          
          CommonDialog2.Filter = "Excel File (*.xls)|*.xls"
          CommonDialog2.ShowSave
          If Trim(CommonDialog2.FileName) <> "" Then
            a.RecordSource = vaExport
            a.ExportToExcel , , , , CommonDialog2.FileName
          End If
    End If
  End If
End Sub

Private Sub cmdHapus_Click()
  If GetModePenjualanUser(objData, GetRegistry(reg_Username)) = 0 Then
    nPos = Delete
    GetEdit True
    GetFakturBrowse True
    cmdAddOrder.Enabled = False
  Else
    MsgBox "Aktifkan dulu fungsi Full, User ini berada di Fungsi Compact", vbInformation, "Tidak bisa di edit"
  End If
End Sub

Private Sub cmdImport_Click()
  CommonDialog1.Filter = "Excel File (*.xls)|*.xls"
  CommonDialog1.ShowOpen
'  GetLoadExcel
End Sub

'Private Sub GetLoadExcel()
'Dim lSave As Boolean
'Dim vaField, vaValue
'Dim i, j, n As Integer
'Dim dbData As New ADODB.Recordset
'
'  On Error GoTo err:
''  StartExcel
'  lSave = True
'
''  Excel.Workbooks.Close
''  Set ExcelWBk = Excel.Workbooks.Open(CommonDialog1.FileName)
''  Set ExcelWS = ExcelWBk.Worksheets(1)
''
''  FrmPB.InitPB ExcelWS.Cells.SpecialCells(xlCellTypeLastCell).Row
'  Dim cBarcode
'  Dim cQty
'
'  For i = 1 To ExcelWS.Cells.SpecialCells(xlCellTypeLastCell).Row
'    FrmPB.RunPB
'    With ExcelWS
'      Set dbData = objData.Browse(GetDSN, "stock", "kodestock,nama,hargabeli,diskonpenjualan,kodesatuan,barcode,jenis", "barcode", sisAssign, .Cells(i, 1).Value)
'      If Not dbData.EOF Then
'        vaArray.InsertRows vaArray.UpperBound(1) + 1
'        n = vaArray.UpperBound(1)
'        vaArray(n, 0) = n + 1
'        vaArray(n, 1) = .Cells(i, 1).Value
'        vaArray(n, 2) = GetNull(dbData!nama)
'        vaArray(n, 3) = .Cells(i, 2).Value
'        vaArray(n, 4) = GetNull(dbData!kodesatuan)
'
'        vaArray(n, 5) = IIf(Trim(.Cells(i, 3)) = "", GetNull(dbData!hargabeli), GetNull(.Cells(i, 3).Value))
'        vaArray(n, 6) = IIf(Trim(.Cells(i, 4)) = "", IIf(GetNull(dbData!diskonpenjualan) = 0, 0, GetNull(dbData!diskonpenjualan)), GetNull(.Cells(i, 4).Value))
'
'
'        vaArray(n, 7) = (vaArray(n, 5) - (vaArray(n, 5) * vaArray(n, 6) / 100)) * vaArray(n, 3)
'
'        vaArray(n, 8) = GetNull(dbData!KodeStock)
'        vaArray(n, 9) = GetNull(dbData!jenis)
'      Else
'        'jika data yg di import tidak ada dalam database simpan
'      End If
'    End With
'  Next i
'  nNomor.Value = vaArray.UpperBound(1) + 2
'  Set TDBGrid1.Array = vaArray
'  TDBGrid1.ReBind
'  TDBGrid1.Refresh
'  SumTotal
'  FrmPB.EndPB
'  CloseWorkSheet
'  FinishExcel
'
'err:
'End Sub

Private Sub cmdKeluar_Click()
  If lEdit Then
    GetEdit False
    initvalue
  Else
    Unload Me
  End If
End Sub

Private Function validOK() As Boolean
Dim nKe As Integer

  validOK = True
  If Not GetValidDataBrowse(objData, "stock", "kodestock", cKode) Then
    MsgBox "Barang tersebut tidak ada dalam database "
    cBarcode.SetFocus
    validOK = False
    Exit Function
  End If
  
  If Not GetValidDataBrowse(objData, "stock", "barcode", cBarcode.Text) Then
    MsgBox "Barang tersebut tidak ada dalam database "
    cBarcode.SetFocus
    validOK = False
    Exit Function
  End If
  
  
'  If isInGrid(vaArray, 8, cKode) And nNomor.Value > vaArray.UpperBound(1) + 1 Then
'    MsgBox "Data sudah pernah dimasukkan sebelumnya ..", vbExclamation
'    cBarcode.SetFocus
'    validOK = False
'    Exit Function
'  End If
  
  Dim nCekHarga As Double
  
  If aCfg(objData, msIjinkanHargaBeliDibawahHargajual) <> 3 Then
    Set dbData = objData.Browse(GetDSN, "stock", , "kodestock", sisAssign, cKode)
    If Not dbData.EOF Then

      Select Case aCfg(objData, msIjinkanHargaBeliDibawahHargajual)
        Case 1
          nCekHarga = GetNull(dbData!hargabeli)
        Case 2
          nCekHarga = GetNull(dbData!cogs)
      End Select
      
      nCekHarga = GetNull(dbData!hargabeli)
      If nCekHarga > (nHarga.value - (nHarga.value * nDisc1.value) / 100) Then
        MsgBox "Stop" & vbCrLf & "Maaf. tidak bisa dilanjutkan." & vbCrLf & "Harga jual tidak sesuai, silahkan hubungi supervisor untuk penjelasan lebih lanjut." & vbCrLf & "Terimaksih", vbInformation
        If MsgBox("Apakah transaksi tetap akan dilanjutkan", vbYesNo + vbExclamation) = vbYes Then
          'jika level 0 kasi otoritas
          If objMenu.UserLevel <> 0 Then
            validOK = False
            MsgBox "Maaf Tidak Bisa dilanjutkan, Gunakan Akun Level yg lebih tinggi", vbCritical
            Exit Function
          End If
        Else
          validOK = False
          Exit Function
        End If
      End If
    End If
  End If

  If isInGrid(vaArray, 8, cKode, , nKe) And nNomor.value > vaArray.UpperBound(1) Then
    'MsgBox "Data sudah pernah dimasukkan sebelumnya dan akan dijumlahkan dengan data sebelumnya", vbExclamation
    cBarcode.SetFocus
    validOK = False
    
    'jika barang yg sama diinput 2x dalam waktu bersamaan, maka akan qty akan
    'dijumlahkan dengan yg sebelumnya, baik harga dan diskon akan sesuai dengan data
    'yg diinput terakhir kali
    
    If nNomor.value > nKe + 1 Then
      vaArray(nKe, 3) = vaArray(nKe, 3) + nQty.value
    Else
      vaArray(nKe, 3) = nQty.value
    End If
    vaArray(nKe, 5) = GetHargaGrosir(cKode, nHarga.value, vaArray(nKe, 3))  'nHarga.Value
    vaArray(nKe, 6) = nDisc1.value
    vaArray(nKe, 7) = vaArray(nKe, 3) * (vaArray(nKe, 5) - vaArray(nKe, 5) * vaArray(nKe, 6) / 100)
    
    TDBGrid1.Update
    TDBGrid1.Refresh
    InitValue1
    SumTotal
    nNomor.value = vaArray.UpperBound(1) + 2
    cBarcode.SetFocus
  End If
End Function

Private Function GetHargaGrosir(ByVal cKodeStock As String, ByVal nHargaJualInput As Double, ByVal nQtyTotal As Double) As Double
Dim db As New ADODB.Recordset
Dim nHrgJual As Double

  nHrgJual = nHargaJualInput
  GetHargaGrosir = nHrgJual
  Set db = objData.Browse(GetDSN, "hargagrosir", , , , , , "minqty desc")
  If Not db.EOF Then
    Do While Not db.EOF
      If nQtyTotal >= GetNull(db!minqty) Then
        GetHargaGrosir = nHrgJual - (nHrgJual * GetNull(db!Discount) / 100)
        Exit Do
      End If
      db.MoveNext
    Loop
  End If
End Function

Private Function GetHargaJ(ByVal obj As CodeSuiteLibrary.Data, ByVal cKodeStock As String) As Double
Dim db As New ADODB.Recordset

  Set db = obj.Browse(GetDSN, "stock", , "kodestock", sisAssign, cKodeStock)
  If Not db.EOF Then
    GetHargaJ = GetNull(db!HargaJual)
  End If
End Function

Private Sub cmdOK_Click()
Dim n As Integer

  cNamaCustomer.Enabled = True
  cNamaCustomer.Button = True
  
  If validOK() Then
    cNamaCustomer.Enabled = False
    cNamaCustomer.Button = False
    'jika baris <= Nomor
    If vaArray.UpperBound(1) + 2 <= nNomor.value Then
      vaArray.ReDim 0, nNomor.value - 1, 0, 11
      n = vaArray.UpperBound(1)
    'jika baris=0
    ElseIf vaArray.UpperBound(1) = -1 Then
      nNomor.value = 1
      vaArray.ReDim 0, nNomor.value - 1, 0, 11
      n = vaArray.UpperBound(1)
    Else
      n = nNomor.value - 1
    End If
        
    vaArray(n, 0) = nNomor.value
    vaArray(n, 1) = cBarcode.Text
    vaArray(n, 2) = GetNamaBarang(cKode)
    vaArray(n, 3) = nQty.value
    vaArray(n, 4) = cSatuan.Text
    vaArray(n, 5) = GetHargaGrosir(cKode, nHarga.value, nQty.value)
    vaArray(n, 6) = nDisc1.value
    vaArray(n, 7) = (vaArray(n, 5) - (vaArray(n, 5) * vaArray(n, 6) / 100)) * nQty.value 'nJumlah.Value
    vaArray(n, 8) = cKode
    vaArray(n, 9) = cJenis
    vaArray(n, 10) = nBValue
    vaArray(n, 11) = nInfoStok.value
  
    TDBGrid1.Array = vaArray
    TDBGrid1.ReBind
    TDBGrid1.MoveNext
    GetUpdateTotal
  End If
End Sub

Private Function GetNamaBarang(cKode As String) As String
Dim db As New ADODB.Recordset
  
  Set db = objData.Browse(GetDSN, "stock", "nama", "kodestock", sisAssign, cKode)
  If Not db.EOF Then
    GetNamaBarang = GetNull(db!nama)
  End If
End Function

Private Sub GetUpdateTotal()
Dim nJumlah1 As Double
Dim nQtyTmp As Single
Dim n As Single

    nJumlah1 = 0
    nQtyTmp = 0
    For n = 0 To vaArray.UpperBound(1)
      nJumlah1 = nJumlah1 + vaArray(n, 7)
      nQtyTmp = nQtyTmp + vaArray(n, 3)
    Next
    
    nSubTotal.value = nJumlah1
    TDBGrid1.Columns(3).FooterText = Format(nQtyTmp, "###,###,##0")
    SumTotal
    InitValue1
    nNomor.value = vaArray.UpperBound(1) + 2
    cBarcode.SetFocus
    nPoinReguler.value = GetHitungPoinHadiah(aCfg(objData, msKelipatan))
    
End Sub

'Private Function GetHargaGrosir(ByVal nHargaJualEceran As Double, ByVal nQtySales As Double) As Double
'  Select Case nQtySales
'    Case 1 To 5
'      GetHargaGrosir = nHargaJualEceran - (nHargaJualEceran * (1 / 100))
'    Case 6 To 10
'      GetHargaGrosir = nHargaJualEceran - (nHargaJualEceran * (2 / 100))
'    Case 11 To 1000
'      GetHargaGrosir = nHargaJualEceran - (nHargaJualEceran * (3 / 100))
'  End Select
'End Function

Private Sub GetClearGrid()
  vaArray.ReDim 0, -1, 0, 11
  Set TDBGrid1.Array = vaArray
  TDBGrid1.ReBind
  TDBGrid1.Refresh
  nNomor.value = vaArray.UpperBound(1) + 2
  GetUpdateTotal
End Sub

Private Sub TambahkanDariRekapPromo()
Dim db As New ADODB.Recordset
Dim n As Single
Dim nJumlah1
Dim nQtyTmp
  
  nJumlah1 = 0
  nQtyTmp = 0
  
  Set db = objData.Browse(GetDSN, "promo p", "p.barcode,p.kodestock,s.nama,p.qty,s.kodesatuan,s.hargajual,s.diskonpenjualan,s.jenis", "p.kodeanggota", sisAssign, cCustomer.Text, " and k.status = 'A'", , Array("left join stock s on s.kodestock = p.kodestock", "left join katalog k on k.kodekatalog = p.kodekatalog"))
  If Not db.EOF Then
    Do While Not db.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = n + 1
      vaArray(n, 1) = GetNull(db!barcode) 'cBarcode.Text
      vaArray(n, 2) = GetNull(db!nama) 'cNama.Text
      vaArray(n, 3) = GetNull(db!qty) 'nQty.Value
      vaArray(n, 4) = GetNull(db!kodesatuan) 'cSatuan.Text
      vaArray(n, 5) = GetNull(db!HargaJual) 'nHarga.Value
      vaArray(n, 6) = GetNull(db!diskonpenjualan) 'nDisc1.Value
      vaArray(n, 7) = vaArray(n, 3) * (vaArray(n, 5) - (vaArray(n, 5) * vaArray(n, 6) / 100)) 'nJumlah.Value
      vaArray(n, 8) = GetNull(db!KodeStock)
      vaArray(n, 9) = GetNull(db!jenis)
      'cek apakah barang ini sudah dibuatkan nota promo
      db.MoveNext
    Loop
  End If
  vaArray.InsertRows vaArray.UpperBound(1) + 1
  
  Set TDBGrid1.Array = vaArray
  TDBGrid1.ReBind
  TDBGrid1.ReBind
  TDBGrid1.Refresh
  InitValue1
  
  nJumlah1 = 0
  nQtyTmp = 0

  For n = 0 To vaArray.UpperBound(1)
    nJumlah1 = nJumlah1 + vaArray(n, 7)
    nQtyTmp = nQtyTmp + vaArray(n, 3)
  Next

  nSubTotal.value = nJumlah1
  TDBGrid1.Columns(3).FooterText = Format(nQtyTmp, "###,###,##0")

  SumTotal
  InitValue1

  nNomor.value = vaArray.UpperBound(1) + 2
  cBarcode.SetFocus
  TDBGrid1.ReBind

End Sub

Private Sub SumTotal()
Dim n As Double
Dim nT As Double
  
  nSubTotal.value = 0
  For n = 0 To vaArray.UpperBound(1)
    nSubTotal.value = nSubTotal.value + vaArray(n, 7)
  Next
  
  If nPersDisc.Enabled = True Then
    nDiscount.value = nPersDisc.value / 100 * (nSubTotal.value)
  End If
  
  nPajak.value = (nPPn.value / 100) * (nSubTotal.value - (nDiscount.value + nDiscount.value))
  nTotal.value = nSubTotal.value + nPajak.value - nDiscount.value
  If aCfg(objData, msPerhitunganKomisi) = 2 Then 'auto
    nKomisi.value = nTotal.value * Devide(aCfg(objData, msPersenKomisi), 100)
  End If
  
  nT = nTotal.value - nDP.value + nHargaCOD.value
  nTotal.value = nT
  If chkTunai.value = 1 Then
    nTunai.value = nT
    nPiutang.value = 0
  Else
    nPiutang.value = nT
    nTunai.value = 0
  End If
End Sub

Private Sub GetHapusPending()
Dim n As Single
Dim lSave As Boolean
  
  lSave = True
  objData.Start GetDSN
  lSave = IIf(lSave, objData.Delete(GetDSN, "pendingtrans", "userid", sisAssign, GetRegistry(reg_Username)), False)
  If lSave Then
    objData.Save GetDSN
    MsgBox "Data pendingan sudah dikosongkan", vbInformation
  Else
    objData.Cancel GetDSN
    MsgBox "Data Pendingan gagal di hapus", vbInformation
  End If
End Sub

Private Sub cmdSaveOK_Click()
'  Simpan
End Sub

Private Function ValidSaving() As Boolean
  ValidSaving = True
End Function

Private Sub cCustomer_ButtonClick()
Dim vaTmp As New XArrayDB

  Set dbData = objData.Browse(GetDSN, "anggota a", "a.kodeanggota,a.nama,a.kodedep,a.alamat,a.telp,a.dd,d.keterangan", "a.kodeanggota", sisContent, cCustomer.Text, , "a.kodeanggota,a.nama", Array("Left join dep d on d.kodedep = a.kodedep"))
  If Not dbData.EOF Then
    cCustomer.Text = cCustomer.Browse(dbData, Array("KODE", "NAMA", "DEP", "ALAMAT"), , Array(10, 20, 6, 10))
    cCustomer.Text = GetNull(dbData!kodeanggota)
    
    If nPos = Add Then
    Set vaTmp = GetLewatJatuhTempo(objData, cCustomer.Text)
    If vaTmp.UpperBound(1) >= 0 Then
        MsgBox "Maaf, customer ini tidak diperkenankan membuka nota baru. Masih " & vaTmp.UpperBound(1) + 1 & " ada nota jatuh tempo yg belum dilunasi. Terimakasih"
      End If
    End If

    cNamaCustomer.Text = GetNull(dbData!nama, "")
    cAlamat.Text = GetNull(dbData!alamat, "")
    cKota.Text = GetNull(dbData!kodedep, "")
    cTelp = GetNull(dbData!telp)

    dJthTmp.value = Format(DateAdd("d", GetNull(dbData!dd), Date), "yyyy-MM-dd")
'    If nPos = Add Then
'      GetOderan cCustomer.Text
'    End If
  End If
End Sub

Private Sub GetRegCustomer()
Dim vaTmp As New XArrayDB

  Set dbData = objData.Browse(GetDSN, "anggota a", "a.kodeanggota,a.nama,a.kodedep,a.alamat,a.telp,a.dd,d.keterangan", "a.kodeanggota", sisContent, GetNull(GetRegistry(reg_KodeAnggota)), , "a.kodeanggota,a.nama", Array("Left join dep d on d.kodedep = a.kodedep"))
  If Not dbData.EOF Then
    'cCustomer.Text = cCustomer.Browse(dbData, Array("KODE", "NAMA", "DEP", "ALAMAT"), , Array(10, 20, 6, 10))
    cCustomer.Text = GetNull(dbData!kodeanggota)
    
'    If nPos = Add Then
'    Set vaTmp = GetLewatJatuhTempo(objData, cCustomer.Text)
'    If vaTmp.UpperBound(1) >= 0 Then
'        MsgBox "Maaf, customer ini tidak diperkenankan membuka nota baru. Masih " & vaTmp.UpperBound(1) + 1 & " ada nota jatuh tempo yg belum dilunasi. Terimakasih"
'      End If
'    End If

    cNamaCustomer.Text = GetNull(dbData!nama, "")
    cAlamat.Text = GetNull(dbData!alamat, "")
    cKota.Text = GetNull(dbData!kodedep, "")
    cTelp = GetNull(dbData!telp)

    dJthTmp.value = Format(DateAdd("d", GetNull(dbData!dd), Date), "yyyy-MM-dd")
  End If
    
    'chkTunai.Value = GetNull(GetRegistry(reg_chkTunaiPenjualan))
End Sub

Private Sub cmdPending_Click()
Dim n As Single
Dim lSave As Boolean
  
  lSave = True
  
  If vaArray.UpperBound(1) >= 0 Then
    'hapus transaksi pending sebelum nya
    objData.Start GetDSN
    lSave = IIf(lSave, objData.Delete(GetDSN, "pendingtrans", "userid", sisAssign, GetRegistry(reg_Username)), False)
    For n = vaArray.LowerBound(1) To vaArray.UpperBound(1)
      lSave = IIf(lSave, objData.Add(GetDSN, "pendingtrans", Array("nomorpenjualan", "kodegudang", "tgl", "kodestock", "qty", "harga", "kodesatuan", "discount", "jumlah", "hb", "tunai", "piutang", "urutfaktur", "bv", "userid"), Array(cFaktur.Text, cGudang.Text, Format(dTgl.value, "yyyy-MM-dd"), vaArray(n, 8), vaArray(n, 3), vaArray(n, 5), vaArray(n, 4), vaArray(n, 6), vaArray(n, 7), GetHargaBeli(objData, vaArray(n, 8)), "1", "1", vaArray(n, 0), vaArray(n, 10), GetRegistry(reg_Username))), False)
    Next n
    If lSave Then
      objData.Save GetDSN
      InitValue1
      MsgBox "Transaksi sudah dipending, silahkan klik F6 untuk memanggil kembali", vbInformation
    Else
      objData.Cancel GetDSN
    End If
  Else
    MsgBox "Tidak ada data untuk disimpan", vbInformation
  End If
End Sub

Private Sub cmdSimpan_Click()
Dim lSave As Boolean
Dim Faktur As String
Dim n As Integer
Dim i As Integer
Dim nValueTunai As Double
Dim nValueKredit As Double
Dim cPreKett As String

  lSave = True

  'simpan pada tabel totpenjualan
  'simpan pada tabel penjualan
  'simpan pada tabel kartustock
  'simpan pada tabel kartupiutang
  'simpan pada tabel bukubesar
  
  'objData.Cancel GetDSN
  
  If isValidSaving Then
    GetNotifikasiAdd "Menyimpan Penjualan"
    objData.Start GetDSN
    Faktur = cFaktur.Text
    
    lSave = IIf(lSave, objData.Delete(GetDSN, "orderan", "reffid", sisAssign, Faktur), False)

    'null kan semua variable public yg dikirim ke form kasir
    nKasirBayar = 0
    nKasirVoucher = 0
    vaVoucher.ReDim 0, -1, 0, 6
    If chkTunai.value = 1 Then
      cPreKett = "Penjualan Tunai"
    Else
      cPreKett = "Penjualan Bon"
    End If
    If lSave = True Then
      If nPos = Add Or Edit Then
        If chkTunai.value = 1 Then
          trFormKasir.nJumlahYgHarusDibayar.value = nTunai.value
          
'          If GetSaldoTopUpMember(objData, cCustomer.Text) < nTunai.Value Then
'            trFormKasir.nVoucher.Value = GetSaldoTopUpMember(objData, cCustomer.Text)
'          Else
'            MsgBox "Member ini memiliki saldo TOP UP/Voucher senilai Rp. " & Format(GetSaldoTopUpMember(objData, cCustomer.Text), "###,###,##0.00") & " Namun tidak bisa digunakan karena total belanja masih kurang. Maaf ya hehe.."
'            trFormKasir.nVoucher.Value = 0
'          End If
          
          trFormKasir.nTunai.value = trFormKasir.nJumlahYgHarusDibayar.value - trFormKasir.nVoucher.value 'nTunai.Value
          trFormKasir.Show vbModal
          If lSign = 1 Then
            lSave = False
          Else
            lSave = True
          End If
        End If
      End If
    End If
  
    'update status voucher
    'set lstatus dan lstatusid
    For n = 0 To vaVoucher.UpperBound(1)
      If vaVoucher(n, 0) = -1 Then
        lSave = IIf(lSave, objData.Edit(GetDSN, "membertopup", "nomormembertopup = '" & vaVoucher(n, 1) & "'", Array("lstatus", "lstatusid"), Array(sisFlag.Posting, cFaktur.Text)), False)
      End If
    Next n
    
    'simpan di table totpenjualan
    lSave = IIf(lSave, objData.Update(GetDSN, "totpenjualan", "nomorpenjualan = '" & Faktur & "'", _
    Array( _
      "nomorpenjualan", "fakturasli", "tgl", "jthtmp", "kodeanggota", "ppn", _
      "persdisc", "persdisc2", "subtotal", "pajak", "discount", "discount2", "bayar", _
      "total", "tunai", "voucher", "piutang", "datetime", "username", _
      "kodeakun", "kodecostcenter", "flaglunas", "kodesalesman", "komisi", "dp", _
      "kodegudang", "upkepada", "jenis", "keterangan", "ongkir", "kodegroupsales"), _
    Array( _
      Faktur, cFakturAsli.Text, Format(dTgl.value, "yyyy-MM-dd"), Format(dJthTmp.value, "yyyy-MM-dd"), cCustomer.Text, nPPn.value, _
      nPersDisc.value, 0, nSubTotal.value, nPajak.value, nDiscount.value, 0, nKasirBayar, _
      nTotal.value, nTunai.value - nKasirVoucher, nKasirVoucher, nPiutang.value, SNow, GetRegistry(reg_Username), _
      cAkunKas.Text, GetCostCenterUser(objData, GetRegistry(reg_Username)), 0, cSalesman.Text, nKomisi.value, nDP.value, _
      cGudang.Text, cUp.Text, GetOpt(optPromo), cKeterangan.Text, nHargaCOD.value, GetRegistry(reg_KodeGroupPenjualan))), False)
    
    lSave = IIf(lSave, objData.Delete(GetDSN, "penjualan", "nomorpenjualan", sisAssign, Faktur), False)
    lSave = IIf(lSave, objData.Delete(GetDSN, "kartustock", "nomor", sisAssign, Faktur), False)
    lSave = IIf(lSave, objData.Delete(GetDSN, "penjualankonsinyasi", "nomorpenjualan", sisAssign, Faktur), False)
    
    'cek dulu apakah nota ini memenuhi kuota hutang?
    
    If aCfg(objData, msMinimumDeposit) > 0 Then
      If (nTotal.value * Devide(aCfg(objData, msMinimumDeposit), 100)) + GetSaldoPiutang(objData, cCustomer.Text) > GetSaldoTopUpMember(objData, cCustomer.Text) Then
        MsgBox "Maaf data tidak bisa disimpan" & vbCrLf _
        & "Piutang di Nota ini + Outstanding Piutang lebih besar dari nilai deposit" & vbCrLf _
        & Format((nTotal.value * 80 / 100), "###,###, ##00") & " + " & Format(GetSaldoPiutang(objData, cCustomer.Text), "###,###,##00") & " > " & Format(GetSaldoTopUpMember(objData, cCustomer.Text), "###,###,##00")
        lSave = False
      End If
    End If
    
    'Update status order menjadi 1
    lSave = IIf(lSave, objData.Edit(GetDSN, "totmemberorder", "nomormemberorder = '" & cNoOrder & "'", Array("status", "nomorpenjualan"), Array(1, Faktur)), False)
    
    For n = vaArray.LowerBound(1) To vaArray.UpperBound(1)
    
      'cek dulu saldo stok apakah ada yg mines, jika ada stop
      If aCfg(objData, msSaldoMinus) = 2 Then
        If GetSaldoStock(objData, cGudang.Text, vaArray(n, 8)) - vaArray(n, 3) < 0 And vaArray(n, 9) <> 9 Then
          MsgBox vaArray(n, 1) & " " & vaArray(n, 2) & " : " & GetSaldoStock(objData, cGudang.Text, vaArray(n, 8)), vbExclamation, "Stok Mines"
          lSave = False
        End If
      End If
      
      
      nValueTunai = 0
      nValueKredit = 0
      
      If chkTunai.value = 1 Then
        nValueTunai = vaArray(n, 5) - (vaArray(n, 5) * vaArray(n, 6) / 100)
        nValueKredit = 0
      Else
        nValueKredit = vaArray(n, 5) - (vaArray(n, 5) * vaArray(n, 6) / 100)
        nValueTunai = 0
      End If
      
      '*******************
      'KARTU STOCK UDPATE
      '*******************
      ' sebelum di proses update harga pokok di table stock dengan yg terbaru
      '
      'MsgBox GetSaldoStock(objData, "", vaArray(n, 8))
      'MsgBox "HPP Lama : " & GetHargaPokok(objData, vaArray(n, 8))
      lSave = IIf(lSave, objData.Edit(GetDSN, "stock", "kodestock = '" & vaArray(n, 8) & "'", Array("cogs"), Array(NewUpdHargaPokok(objData, vaArray(n, 8)))), False)
      'MsgBox "HPP Baru : " & GetHargaPokok(objData, vaArray(n, 8))
      '*******************
      
      lSave = IIf(lSave, objData.Add(GetDSN, "orderan", Array("tgl", "reffid", "kodeanggota", "kodestock", "barcode", "kredit", "username", "datetime"), Array(Format(dTgl.value, "yyyy-MM-dd"), Faktur, cCustomer.Text, vaArray(n, 8), vaArray(n, 1), vaArray(n, 3), GetRegistry(reg_Username), SNow)), False)
      lSave = IIf(lSave, objData.Add(GetDSN, "penjualan", Array("nomorpenjualan", "kodegudang", "tgl", "kodestock", "qty", "harga", "kodesatuan", "discount", "jumlah", "hb", "tunai", "piutang", "urutfaktur", "bv"), Array(Faktur, cGudang.Text, Format(dTgl.value, "yyyy-MM-dd"), vaArray(n, 8), vaArray(n, 3), vaArray(n, 5), vaArray(n, 4), vaArray(n, 6), vaArray(n, 7), GetHargaBeli(objData, vaArray(n, 8)), nValueTunai, nValueKredit, vaArray(n, 0), vaArray(n, 10))), False)
      lSave = IIf(lSave, UpdKartuStock(objData, SisKartuStock.Penjualan, Faktur, Format(dTgl.value, "yyyy-MM-dd"), vaArray(n, 8), vaArray(n, 3), vaArray(n, 5), vaArray(n, 6), cPreKett & " an. " & cNamaCustomer.Text, cGudang.Text, GetHargaPokok(objData, vaArray(n, 8))), False)

      'Update status lunas
      If lCekStatusLunas(objData, Faktur) = True Then
        lSave = IIf(lSave, objData.Edit(GetDSN, "penjualan", "nomorpenjualan = '" & Faktur & "'", Array("statuslunas"), Array(1)), False)
      Else
        lSave = IIf(lSave, objData.Edit(GetDSN, "penjualan", "nomorpenjualan = '" & Faktur & "'", Array("statuslunas"), Array(0)), False)
      End If
      
    Next n
    
    'isi field flaglunas
    'cek apakah dp yg dibayarkan lebih dari/sama dengan yg diminta
    
    If nDP.value >= nTotal.value Then
      lSave = IIf(lSave, objData.Edit(GetDSN, "totpenjualan", "nomorpenjualan = '" & Faktur & "'", Array("flaglunas"), Array(1)), False)
    End If
    
    Dim vaField
    Dim vaValue
    
    If chkTunai.value = 1 Then
      lSave = IIf(lSave, objData.Edit(GetDSN, "totpenjualan", "nomorpenjualan = '" & Faktur & "'", Array("flaglunas"), Array(1)), False)
      
      'Dapat Poin belanja
      If nPoinReguler.value > 0 Then
        vaField = Array("faktur", "tgl", "kodeanggota", "poinhadiah", "exdate", "status")
        vaValue = Array(Faktur, Format(Date, "yyyy-MM-dd"), cCustomer.Text, GetHitungPoinHadiah(aCfg(objData, msKelipatan)), Format(DateAdd("D", aCfg(objData, msTerm), Date), "yyyy-MM-dd"), "1")
        lSave = IIf(lSave, objData.Add(GetDSN, "poinhadiah", vaField, vaValue), False)
      End If
      
    Else
      If lCekStatusLunas(objData, Faktur) = True Then
        lSave = IIf(lSave, objData.Edit(GetDSN, "totpenjualan", "nomorpenjualan = '" & Faktur & "'", Array("flaglunas"), Array(1)), False)
      End If
    End If
    
    lSave = IIf(lSave, UpdKartuHutang(objData, SisPenjualan, Faktur, Format(dTgl.value, "yyyy-MM-dd"), cCustomer.Text, cPreKett & " an. " & cNamaCustomer.Text, nPiutang.value, SNow, GetRegistry(reg_Username), , GetRegistry(reg_KodeGroupPenjualan)), False)
    
    'jika dibayar tunai dan ada dp maka posting ke kartupiutang
    
    If chkTunai.value = 1 And nDP.value <> 0 Then
      lSave = IIf(lSave, UpdKartuHutang(objData, SisPenjualan, Faktur, Format(dTgl.value, "yyyy-MM-dd"), cCustomer.Text, "Pengembalian DP dengan barang an. " & cNamaCustomer.Text, nDP.value, SNow, GetRegistry(reg_Username), , GetRegistry(reg_KodeGroupPenjualan)), False)
    End If
    
    If chkTunai.value <> 1 And nDP.value <> 0 Then
      lSave = IIf(lSave, UpdKartuHutang(objData, SisPenjualan, Faktur, Format(dTgl.value, "yyyy-MM-dd"), cCustomer.Text, "Pengembalian DP dengan barang an. " & cNamaCustomer.Text, nTotal.value, SNow, GetRegistry(reg_Username), , GetRegistry(reg_KodeGroupPenjualan)), False)
    End If
    
    lSave = IIf(lSave, DelKodeTr(objData, vbTrigger.msPenjualan, Faktur), False)
    
    'Piutang, Kas
    'Kas, piutang
    '   Penjualan
    
    'Diskon Penjualan
    '   Penjualan
    
    'PPn Penjualan
    '   Penjualan
    
    'Inventory
    
    
    lSave = IIf(lSave, UpdKodeTr(objData, msPenjualan, Faktur, Format(dTgl.value, "yyyy-MM-dd"), GetAkunMember(objData, cCustomer.Text), GetCostCenterUser(objData, GetRegistry(reg_Username)), cPreKett & " an " & cNamaCustomer.Text, nPiutang.value + nDP.value), False)
    If nKasirTotalKartu > 0 And chkTunai.value = 1 Then
      
      ' jika dibayar menggunakan kartu
      lSave = IIf(lSave, UpdKodeTr(objData, msPenjualan, Faktur, Format(dTgl.value, "yyyy-MM-dd"), GetAkunKartu(nKasirKodeKartu), GetCostCenterUser(objData, GetRegistry(reg_Username)), "Pembayaran Dengan " & nKasirKeteranganKartu & " an " & nKasirNamaDiKartu, nKasirTotalKartu), False)
      ' kredit pendapatan fee
      lSave = IIf(lSave, UpdKodeTr(objData, msPenjualan, Faktur, Format(dTgl.value, "yyyy-MM-dd"), aCfg(objData, msRekeningFeeKartu), GetCostCenterUser(objData, GetRegistry(reg_Username)), "Pembayaran Dengan " & nKasirKeteranganKartu & " an " & nKasirNamaDiKartu, 0, nKasirFeeTotalKartu), False)
      ' simpan dp di kas kasir yg menginput
      lSave = IIf(lSave, UpdKodeTr(objData, msPenjualan, Faktur, Format(dTgl.value, "yyyy-MM-dd"), cAkunKas.Text, GetCostCenterUser(objData, GetRegistry(reg_Username)), "Pembayaran Dengan " & nKasirKeteranganKartu & " an " & nKasirNamaDiKartu, nDPKasir), False)

      'lakukan penyimpanan di tabel kartu
      lSave = IIf(lSave, objData.Update(GetDSN, "trkartu", "nomorpenjualan='" & Faktur & "'", Array( _
        "kodekartu", "nomorpenjualan", "subtotal", "fee", "total", "nomortrace", "nama"), Array( _
        nKasirKodeKartu, Faktur, nKasirTotalKartu - nKasirFeeTotalKartu, nKasirFeeTotalKartu, nKasirTotalKartu, nKasirNoTraceKartu, nKasirNamaDiKartu)), False)
    Else
      'Jika dibayar tunai
      lSave = IIf(lSave, UpdKodeTr(objData, msPenjualan, Faktur, Format(dTgl.value, "yyyy-MM-dd"), cAkunKas.Text, GetCostCenterUser(objData, GetRegistry(reg_Username)), cPreKett & " an " & cNamaCustomer.Text, nTunai.value - nKasirVoucher), False)
    End If
    lSave = IIf(lSave, UpdKodeTr(objData, msPenjualan, Faktur, Format(dTgl.value, "yyyy-MM-dd"), aCfg(objData, msRekeningPenjualan), GetCostCenterUser(objData, GetRegistry(reg_Username)), cPreKett & " an " & cNamaCustomer.Text, 0, nTotal.value), False)
    'lSave = IIf(lSave, UpdKodeTr(objData, msPenjualan, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), aCfg(objData, msRekeningTopUp), GetCostCenterUser(objData, GetRegistry(reg_username)), "Penjualan an " & cNamaCustomer.Text, 0, nKasirVoucher), False)

    'Posting balik dp yg sudah pernah dilakukan:
    
    'Debet
    Dim nTmp As Double
    Dim nSaldoTmp As Double
    Dim nTmpCOGS As Double
    Dim nTmpSaldoCOGS As Double
    Dim db As New ADODB.Recordset
    
    For n = vaArray.LowerBound(1) To vaArray.UpperBound(1)
      'Discount Pembelian per item
      nTmp = vaArray(n, 3) * vaArray(n, 5) - vaArray(n, 7)
      nSaldoTmp = nSaldoTmp + nTmp
      lSave = IIf(lSave, UpdKodeTr(objData, msPenjualan, Faktur, Format(dTgl.value, "yyyy-MM-dd"), aCfg(objData, msRekeningDiscountPenjualan), GetCostCenterUser(objData, GetRegistry(reg_Username)), "Dsc Item Penjualan an " & cNamaCustomer.Text, nTmp, 0, "", SNow), False)
      
      Set db = objData.Browse(GetDSN, "stock", "kodestock,asbiaya,jenis,autobiaya,konsi", "kodestock", sisAssign, vaArray(n, 8))
        If Not db.EOF Then
                  
          If GetNull(db!jenis) = 1 Then
            'dan jika bukan barang konsinyasi
            If GetNull(db!konsi) = "0" Then
              'Wajib posting cogs
              nTmpCOGS = vaArray(n, 3) * GetHargaPokok(objData, vaArray(n, 8))
              lSave = IIf(lSave, UpdKodeTr(objData, msPenjualan, Faktur, Format(dTgl.value, "yyyy-MM-dd"), aCfg(objData, msRekeningCOGS), GetCostCenterUser(objData, GetRegistry(reg_Username)), "COGS Penjualan an " & vaArray(n, 2), nTmpCOGS, 0, "", SNow, vaArray(n, 8)), False)
              lSave = IIf(lSave, UpdKodeTr(objData, msPenjualan, Faktur, Format(dTgl.value, "yyyy-MM-dd"), GetAkunInventory(objData, vaArray(n, 8)), GetCostCenterUser(objData, GetRegistry(reg_Username)), "COGS Penjualan an " & vaArray(n, 2), 0, nTmpCOGS, "", SNow, vaArray(n, 8)), False)
            Else
              'jika barang konsi wajib disimpan di tabel penjualankonsinyasi
              lSave = IIf(lSave, objData.Add(GetDSN, "penjualankonsinyasi", _
                      Array("nomorpenjualan", "tgl", "kodegudang", "kodestock", "qty", "hargajual", "hargabeli", "discount", "margin"), _
                      Array(Faktur, Format(dTgl.value, "yyyy-MM-dd"), cGudang.Text, vaArray(n, 8), vaArray(n, 3), vaArray(n, 5), GetHargaBeli(objData, vaArray(n, 8)), 0, 0)), False)
            End If
          End If
          
          'Jika Non Inventory
          If GetNull(db!jenis) = 9 Then
            'dan auto posting biaya di set 1 maka
            If GetNull(db!autobiaya) = 1 Then
               nTmpCOGS = vaArray(n, 3) * GetHargaBeli(objData, vaArray(n, 8))
               lSave = IIf(lSave, UpdKodeTr(objData, msPenjualan, Faktur, Format(dTgl.value, "yyyy-MM-dd"), aCfg(objData, msRekeningHutangBiaya), GetCostCenterUser(objData, GetRegistry(reg_Username)), "BHP Penjualan an " & vaArray(n, 2), , nTmpCOGS, "", SNow, vaArray(n, 8)), False)
               lSave = IIf(lSave, UpdKodeTr(objData, msPenjualan, Faktur, Format(dTgl.value, "yyyy-MM-dd"), aCfg(objData, msRekeningBiayaBarang), GetCostCenterUser(objData, GetRegistry(reg_Username)), "BHP Penjualan an " & vaArray(n, 2), nTmpCOGS, , "", SNow, vaArray(n, 8)), False)
            End If
          End If
          
          
'          If GetNull(db!asbiaya) <> "1" And (GetNull(db!jenis) = 1) Then
'            'posting cogs
'            nTmpCOGS = vaArray(n, 3) * GetHargaPokok(objData, vaArray(n, 8))
'            lSave = IIf(lSave, UpdKodeTr(objData, msPenjualan, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), aCfg(objData, msRekeningCOGS), aCfg(objData, msCostCenterJualBeli), "COGS Penjualan an " & vaArray(n, 2), nTmpCOGS, 0, "", SNow), False)
'            lSave = IIf(lSave, UpdKodeTr(objData, msPenjualan, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), GetAkunInventory(objData, vaArray(n, 8)), aCfg(objData, msCostCenterJualBeli), "COGS Penjualan an " & vaArray(n, 2), 0, nTmpCOGS, "", SNow), False)
'          Else
'            'Posting sebagai biaya, sebesar harga beli'
'            'Kas (D)
'            '     Penjualan (K)
'            'COGS (D)
'            '     Hutang Titipan Penjualan (K)
'
'            nTmpCOGS = vaArray(n, 3) * GetHargaPokok(objData, vaArray(n, 8))
'
'            lSave = IIf(lSave, UpdKodeTr(objData, msPenjualan, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), aCfg(objData, msRekeningCOGS), aCfg(objData, msCostCenterJualBeli), "COGS Penjualan an " & vaArray(n, 2), nTmpCOGS, 0, "", SNow, vaArray(n, 8)), False)
'            lSave = IIf(lSave, UpdKodeTr(objData, msPenjualan, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), aCfg(objData, msRekeningHutangBiaya), aCfg(objData, msCostCenterJualBeli), "COGS Penjualan an " & vaArray(n, 2), , nTmpCOGS, "", SNow, vaArray(n, 8)), False)
'            'lSave = IIf(lSave, UpdKodeTr(objData, msPenjualan, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), GetAkunInventory(objData, vaArray(n, 8)), aCfg(objData, msCostCenterJualBeli), "COGS Penjualan an " & vaArray(n, 2), 0, nTmpCOGS, "", SNow), False)
'          End If
          
        End If
    Next n
    
    'Kredit
    lSave = IIf(lSave, UpdKodeTr(objData, msPenjualan, Faktur, Format(dTgl.value, "yyyy-MM-dd"), aCfg(objData, msRekeningPenjualan), GetCostCenterUser(objData, GetRegistry(reg_Username)), "Dsc Item Penjualan an  " & cNamaCustomer.Text, 0, nSaldoTmp), False)
    
    'PPn
    lSave = IIf(lSave, UpdKodeTr(objData, msPenjualan, Faktur, Format(dTgl.value, "yyyy-MM-dd"), aCfg(objData, SisCfg.msRekeningPPnPenjualan), GetCostCenterUser(objData, GetRegistry(reg_Username)), "PPn Penjualan an " & cNamaCustomer.Text, 0, nPajak.value, "", SNow), False)
        lSave = IIf(lSave, UpdKodeTr(objData, msPenjualan, Faktur, Format(dTgl.value, "yyyy-MM-dd"), aCfg(objData, msRekeningPenjualan), GetCostCenterUser(objData, GetRegistry(reg_Username)), "PPn Penjualan an " & cNamaCustomer.Text, nPajak.value, 0), False)
        
    'Discount seluruhnya
    lSave = IIf(lSave, UpdKodeTr(objData, msPenjualan, Faktur, Format(dTgl.value, "yyyy-MM-dd"), aCfg(objData, msRekeningDiscountPenjualan), GetCostCenterUser(objData, GetRegistry(reg_Username)), "Dsc Total Penjualan an " & cNamaCustomer.Text, nDiscount.value, 0, "", SNow), False)
        lSave = IIf(lSave, UpdKodeTr(objData, msPenjualan, Faktur, Format(dTgl.value, "yyyy-MM-dd"), aCfg(objData, msRekeningPenjualan), GetCostCenterUser(objData, GetRegistry(reg_Username)), "Dsc Total Penjualan an " & cNamaCustomer.Text, 0, nDiscount.value, "", SNow), False)
    
    'Komisi salesman
    ' Hutang komisi
    lSave = IIf(lSave, UpdKodeTr(objData, msPenjualan, Faktur, Format(dTgl.value, "yyyy-MM-dd"), aCfg(objData, msRekeningBiayaKomisi), GetCostCenterUser(objData, GetRegistry(reg_Username)), "Komisi Penjualan Sales " & cSalesman.Text, nKomisi.value, 0, "", SNow), False)
        lSave = IIf(lSave, UpdKodeTr(objData, msPenjualan, Faktur, Format(dTgl.value, "yyyy-MM-dd"), aCfg(objData, msRekeningHutangSalesman), GetCostCenterUser(objData, GetRegistry(reg_Username)), "Komisi Penjualan Sales " & cSalesman.Text, 0, nKomisi.value, "", SNow), False)
    
    
    'update lastactivity
    objData.Update GetDSN, "anggota", "kodeanggota = '" & cCustomer.Text & "'", Array("lastactivity"), Array(Format(dTgl.value, "yyyy-MM-dd"))

    
'LINE BEKAS KASIR DISINI
    
    
    If nKasirVoucher > 0 Then
    
      '***************
      'Modul Voucher
      '---------------
      lSave = IIf(lSave, objData.Delete(GetDSN, "membertopup", "nomormembertopup ", sisAssign, Faktur), False)
      
      vaField = Array("nomormembertopup", "tgl", "kodeanggota", "keterangan", "kredit", "lstatus")
      vaValue = Array(Faktur, dTgl.value, cCustomer.Text, "Redeem Penjualan di Kasir", nKasirVoucher, 1)
      lSave = IIf(lSave, objData.Add(GetDSN, "membertopup", vaField, vaValue), False)
      'Akunting Voucher
      'lSave = IIf(lSave, UpdKodeTr(objData, msPenjualan, Faktur, Format(dTgl.Value, "yyyy-MM-dd"), GetAkunKas(objData, GetRegistry(reg_username)), "", "Redeem Penjualan di Kasir", nKasirVoucher, 0), False)
      lSave = IIf(lSave, UpdKodeTr(objData, msPenjualan, Faktur, Format(dTgl.value, "yyyy-MM-dd"), aCfg(objData, msRekeningTopUp), "", "Redeem Penjualan di Kasir", nKasirVoucher, 0), False)
      '***************
    
    End If
    
    If lSave Then
      lSave = IIf(lSave, objData.Delete(GetDSN, "tgledit", "tgl", sisAssign, Format(dTgl.value, "yyyy-MM-dd")), False)
      objData.Save GetDSN
      SaveRegistry reg_KodeAnggota, cCustomer.Text
      SaveRegistry reg_KodeSalesman, cSalesman.Text
      'TrayRemove aMainmenu.pbTray
      GetNotifikasiRemove
    Else
      objData.Cancel GetDSN
      MsgBox "Maaf data tidak berhasil disimpan", vbExclamation
    End If
    
    If lSave = True Then
      
      If MsgBox("Apakah akan mencetak transaksi ke printer?", vbYesNo + vbInformation) = vbYes Then
        If GetRegistry(reg_CetakanPenjualanNonTunai) = 1 Then
          'Cetak Penjualan Nota NCR
          GetCetakFakturpenjualan objData, Faktur, False
          Unload frmFaktur
        ElseIf GetRegistry(reg_CetakanPenjualanNonTunai) = 2 Then
          'Cetak Penjualan Nota Wartel
          trPrint2.noOrder = Faktur
          Set dbData = objData.Browse(GetDSN, "totpenjualan t", "t.*,a.nama,a.telp", "t.nomorpenjualan", sisAssign, Faktur, , , Array("left join anggota a on a.kodeanggota = t.kodeanggota"))
          If Not dbData.EOF Then
            trPrint2.nSubTotal = GetNull(dbData!Subtotal)
            trPrint2.nDiscount = GetNull(dbData!dp)
            trPrint2.nCash = GetNull(dbData!Tunai)
            trPrint2.nChange = GetNull(dbData!Piutang)
            trPrint2.cKodeMember = GetNull(dbData!kodeanggota)
            trPrint2.cMember = GetNull(dbData!nama)
            trPrint2.cTeleponMember = GetNull(dbData!telp)
            trPrint2.Ups = GetNull(dbData!upkepada)
            trPrint2.dTgNota = Format(GetNull(dbData!tgl), "dd/MM/yyyy")
            trPrint2.dJthTempoNota = Format(GetNull(dbData!jthtmp), "dd/MM/yyyy")
            trPrint2.nTmpPoinHadiah = nPoinReguler.value
            
            Load trPrint2
            trPrint2.Show vbModal
          End If
        ElseIf GetRegistry(reg_CetakanPenjualanNonTunai) = 3 Then
          'Cetak Penjualan Struk
          PrintThermalNew Faktur
          If GetRegistry(reg_CetakBerulang) = "Y" Then
            Do While (MsgBox("Cetak Lagi", vbYesNo + vbInformation) = vbYes)
              PrintThermalNew Faktur
            Loop
          End If
          'OpenDrawer GetRegistry(reg_PortStruk)
          If GetRegistry(reg_OpenCashDrawer) = "Y" Then
            ShellExecuteCapture "open.bat", False
          End If
        Else
          MsgBox "Error Printing, setting printer terlebih dahulu pada menu Options", vbExclamation
        End If
        
      End If
    Else
      'MsgBox "Maaf, terjadi masalah dalam proses penyimpanan" & vbCrLf & "Data tidak bisa disimpan"
    End If
    'save default
    If lSave = True And chkTunai.value = 1 Then
      If nKasirKembalian > 0 Then
        trKembalian.Label1 = Format(nKasirKembalian, "###,###,###")
        trKembalian.Show vbModal
        nKasirKembalian = 0
      End If
    End If
  End If
  
  If lSave Then
    initvalue
    GetEdit False
  End If
End Sub

Private Function GetAkunKartu(ByVal idKartu As Integer) As String
Dim dba As New ADODB.Recordset

  Set dba = objData.Browse(GetDSN, "kartu", , "kodekartu", sisAssign, idKartu)
  If Not dba.EOF Then
    GetAkunKartu = GetNull(dba!kodeakun)
  End If
End Function

Private Sub PrintThermal(ByVal Faktur As String)
Dim n As Double
Dim nBruto As Double
Dim nTotQty As Double
Dim nHargaArray As Double


    Open "lpt1" For Output As #1
    Print #1, Chr(27); Chr(33); Chr(4);
    Print #1, Chr(27) & Chr(97) & Chr(1)
    Print #1, "STRUK PENJUALAN"
    Print #1, aCfg(objData, msNamaPerusahaan)
    Print #1, aCfg(objData, msAlamatPerusahaan)
    Print #1, ""
    Select Case GetRegistry(reg_AlignmentThermal)
      Case 1 ' rata kiri
                Print #1, Chr(27) & Chr(97) & Chr(0)
      Case 2 ' rata kanan
          Print #1, Chr(27) & Chr(97) & Chr(2)
    End Select
    Print #1, "No. " & Faktur
    Print #1, Format(Now, "dd-MM-yyyy HH:MM:SS")
    Print #1, ""

    Print #1, "A/N "; cCustomer.Text; ""
    Print #1, cNamaCustomer.Text
    Print #1, "Telp. "; cTelp
    Print #1, ""

    Print #1, Replicate("-", 27)
    Print #1, Padl("Qty", 6); Padl("Hrg Net", 11); Padl("Jml", 10)
    Print #1, Replicate("-", 27)
    For n = 0 To vaArray.UpperBound(1)
      If vaArray(n, 3) <> 0 Then
        nHargaArray = vaArray(n, 5) - (vaArray(n, 5) * vaArray(n, 6) / 100)
        nBruto = nBruto + (vaArray(n, 3) * nHargaArray)
        nTotQty = nTotQty + vaArray(n, 3)
        Print #1, vaArray(n, 2)  ' vaArray(n, 1) kolom REF barang ditiadakan karena terlalu panjang
        If vaArray(n, 6) <> 0 Then
          Print #1, vaArray(n, 1) & " Rp." & Format(vaArray(n, 5), "#,##0") & " -" & vaArray(n, 6) & "%"
        End If
        Print #1, Padl(Format(vaArray(n, 3), "#,##0"), 3) & " x " & Padl(Format(nHargaArray, "#,###,##0"), 8) & " = " & Padl(Format(vaArray(n, 3) * nHargaArray, "#,###,##0"), 10)
      End If
    Next
    
    Print #1, Replicate("-", 27)
    Print #1, Format(nTotQty, "###,###,##0") & " Items"
    

'   Print #1, Padl("Sub   : ", 9); Padl(Format(nBruto, "###,###,##0"), 10)

'   Print #1, Padl("Disc   : ", 9); Padl(Format(nDiscount.Value, "###,###,##0"), 10)
    Print #1, Padl("Total  : ", 9); Padl(Format(nTotal.value, "###,###,##0"), 10)
    
    If chkTunai.value = 0 Then
      Print #1, Padl("DP     : ", 9); Padl(Format(nDP.value, "###,###,##0"), 10)
      Print #1, Padl("Tunai  : ", 9); Padl(Format(nTunai.value, "###,###,##0"), 10)
      Print #1, Padl("Hutang : ", 9); Padl(Format(nPiutang.value, "###,###,##0"), 10)
    End If
    
    Print #1, Chr(10) ' feed kertas
    Print #1, ""
    Close #1
End Sub


Private Sub PrintThermal2(ByVal Faktur As String)
Dim n As Double
Dim nBruto As Double
Dim nTotQty As Double
Dim nHargaArray As Double
Dim nSpasi As Single

    nSpasi = 3

    Open "lpt1" For Output As #1
    Print #1, Chr(27); Chr(33); Chr(4);
    Print #1, Chr(27) & Chr(97) & Chr(1)
    'Print #1, "STRUK PENJUALAN"
    Print #1, aCfg(objData, msNamaPerusahaan)
    Print #1, aCfg(objData, msAlamatPerusahaan)
    Print #1, ""
    Select Case GetRegistry(reg_AlignmentThermal)
      Case 1 ' rata kiri
                Print #1, Chr(27) & Chr(97) & Chr(0)
      Case 2 ' rata kanan
          Print #1, Chr(27) & Chr(97) & Chr(2)
    End Select
    Print #1, lMarginStruk(nSpasi) & "No. " & Faktur
    Print #1, lMarginStruk(nSpasi) & Format(Now, "dd-MM-yyyy HH:MM:SS")
'    Print #1, ""

'    Print #1, "A/N "; cCustomer.Text; ""
'    Print #1, cNamaCustomer.Text
'    Print #1, "Telp. "; cTelp
'    Print #1, ""


    Print #1, lMarginStruk(nSpasi) & Replicate("-", 30)
    Print #1, lMarginStruk(nSpasi) & Padl("Qty", 5); Padl("Hrg Net", 10); Padl("Jumlah", 12)
    Print #1, lMarginStruk(nSpasi) & Replicate("-", 30)
    For n = 0 To vaArray.UpperBound(1)
      If vaArray(n, 3) <> 0 Then
        nHargaArray = vaArray(n, 5) - (vaArray(n, 5) * vaArray(n, 6) / 100)
        nBruto = nBruto + (vaArray(n, 3) * nHargaArray)
        nTotQty = nTotQty + vaArray(n, 3)
        Print #1, lMarginStruk(nSpasi) & Left(vaArray(n, 2), 28) ' vaArray(n, 1) kolom REF barang ditiadakan karena terlalu panjang
        If vaArray(n, 6) <> 0 Then
          Print #1, lMarginStruk(nSpasi) & vaArray(n, 1) & " Rp." & Format(vaArray(n, 5), "#,##0") & " -" & vaArray(n, 6) & "%"
        End If
        Print #1, lMarginStruk(nSpasi) & Padl(Format(vaArray(n, 3), "#,##0"), 4) & " x " & Padl(Format(nHargaArray, "#,###,##0"), 7) & " = " & Padl(Format(vaArray(n, 3) * nHargaArray, "#,###,##0"), 11)
      End If
    Next
    
    Print #1, lMarginStruk(nSpasi) & Replicate("-", 30)
    Print #1, lMarginStruk(nSpasi) & Format(nTotQty, "###,###,##0") & " Items"
    

'    Print #1, Padl("Sub   : ", 9); Padl(Format(nBruto, "###,###,##0"), 10)

    'Print #1, lMarginStruk(nSpasi) & Padl("Disc   : ", 9); Padl(Format(nDiscount.Value, "###,###,##0"), 10)
    Print #1, lMarginStruk(nSpasi) & Padl("Total  : ", 9); Padl(Format(nTotal.value, "###,###,##0"), 10)
    Print #1, lMarginStruk(nSpasi) & Padl("Bayar  : ", 9); Padl(Format(nKasirBayar, "###,###,##0"), 10)
    Print #1, lMarginStruk(nSpasi) & Padl("Kembali: ", 9); Padl(Format(nKasirKembalian, "###,###,##0"), 10)

    If chkTunai.value = 0 Then
      Print #1, Padl("DP     : ", 9); Padl(Format(nDP.value, "###,###,##0"), 10)
      Print #1, Padl("Tunai  : ", 9); Padl(Format(nTunai.value, "###,###,##0"), 10)
      Print #1, Padl("Hutang : ", 9); Padl(Format(nPiutang.value, "###,###,##0"), 10)
    End If
    
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, Chr(29) ' feed kertas
    
    Close #1
End Sub

Private Sub PrintThermalNew(ByVal Faktur As String)
Dim n As Double
Dim nBruto As Double
Dim nTotQty As Double
Dim nHargaArray As Double
Dim nSpasi As Single

Dim nLebarKertas As Single
Dim nMarginKiri As Single
Dim nLebarEfektif As Single
Dim nKol1 As Byte
Dim nKol2 As Byte
Dim nKol3 As Byte
Dim nKol1_2 As Byte
Dim nKol2_2 As Byte
Dim nKol3_2 As Byte
Dim nTotKol As Byte
Dim nTotKol2 As Byte

Dim cFoot1, cFoot2, cFoot3, cFoot4, cFoot5
cFoot1 = aCfg(objData, msKasir1)
cFoot2 = aCfg(objData, msKasir2)
cFoot3 = aCfg(objData, msKasir3)
cFoot4 = aCfg(objData, msKasir4)
cFoot5 = aCfg(objData, msKasir5)

On Error Resume Next


    nLebarKertas = GetRegistry(reg_LebarKertas)
    nMarginKiri = GetRegistry(reg_MarginKiri)
    nLebarEfektif = GetRegistry(reg_LebarEfektif)
    nKol1 = GetRegistry(reg_LebarKolom1)
    nKol2 = GetRegistry(reg_LebarKolom2)
    nKol3 = GetRegistry(reg_LebarKolom3)
    nKol1_2 = GetRegistry(reg_LebarKolom1_2)
    nKol2_2 = GetRegistry(reg_LebarKolom2_2)
    nKol3_2 = GetRegistry(reg_LebarKolom3_2)
    nTotKol = nKol1 + nKol2 + nKol3
    nTotKol2 = nKol1_2 + nKol2_2 + nKol3_2
    
    nSpasi = nMarginKiri

    If GetRegistry(reg_PortStruk) = "" Then
      Open "lpt1" For Output As #1
    Else
      Open Trim(GetRegistry(reg_PortStruk)) For Output As #1
    End If
    
    'Print #1, Chr(27); Chr(33); Chr(4);
    'Print #1, Chr(27) & Chr(97) & Chr(1)
    
    Print #1, lMarginStruk(nSpasi + ((nLebarKertas - Len(aCfg(objData, msNamaPerusahaan))) / 2)) & ""; aCfg(objData, msNamaPerusahaan); ""
    Print #1, lMarginStruk(nSpasi + ((nLebarKertas - Len(aCfg(objData, msAlamatPerusahaan))) / 2)) & ""; aCfg(objData, msAlamatPerusahaan); ""
    If Trim(aCfg(objData, msTelepon)) <> "" Then
      Print #1, lMarginStruk(nSpasi + ((nLebarKertas - Len(aCfg(objData, msTelepon))) / 2)) & ""; aCfg(objData, msTelepon); ""
    End If
    
    'Print #1, aCfg(objData, msNamaPerusahaan)
    'Print #1, aCfg(objData, msAlamatPerusahaan)
    
    'cetak header perusahaan
    If Trim(cFoot1) <> "" Then
      Print #1, Left(cFoot1, nLebarEfektif)
    End If
    Select Case GetRegistry(reg_AlignmentThermal)
      Case 1 ' rata kiri
                Print #1, Chr(27) & Chr(97) & Chr(0)
      Case 2 ' rata kanan
          Print #1, Chr(27) & Chr(97) & Chr(2)
    End Select
    Print #1, lMarginStruk(nSpasi) & "No:" & Faktur & " #" & Format(dTgl.value, "dd/MM/yyyy")
    Print #1, lMarginStruk(nSpasi) & Format(Now, "dd-MM-yyyy HH:MM:SS")
    Print #1, lMarginStruk(nSpasi) & Left("Kasir. " & GetRegistry(reg_FullName), nTotKol)
    
    If GetRegistry(reg_CetakLabelCustomer) = "Y" Then
      Print #1, ""
      Print #1, lMarginStruk(nSpasi) & "A/N "; cCustomer.Text; ""
      Print #1, lMarginStruk(nSpasi) & Left(cNamaCustomer.Text, nTotKol)
      Print #1, lMarginStruk(nSpasi) & "Telp. "; cTelp
    End If

    Print #1, lMarginStruk(nSpasi) & Replicate("-", nKol1 + nKol2 + nKol3)
    Print #1, lMarginStruk(nSpasi) & Padl("Qty", nKol1); Padl("Hrg Net", nKol2); Padl("Jumlah", nKol3)
    Print #1, lMarginStruk(nSpasi) & Replicate("-", nKol1 + nKol2 + nKol3)
    For n = 0 To vaArray.UpperBound(1)
      If vaArray(n, 3) <> 0 Then
        nHargaArray = vaArray(n, 5) - (vaArray(n, 5) * vaArray(n, 6) / 100)
        nBruto = nBruto + (vaArray(n, 3) * nHargaArray)
        nTotQty = nTotQty + vaArray(n, 3)
        Print #1, lMarginStruk(nSpasi) & Left(IIf(GetRegistry(reg_TampilkanBarcode) = "Y", (vaArray(n, 1) & " "), "") & vaArray(n, 2), nLebarEfektif)  ' vaArray(n, 1) kolom REF barang ditiadakan karena terlalu panjang
        If vaArray(n, 6) <> 0 Then
          Print #1, lMarginStruk(nSpasi) & Left(IIf(GetRegistry(reg_TampilkanBarcode) = "Y", "", vaArray(n, 1)) & " Rp." & Format(vaArray(n, 5), "#,##0") & " -" & vaArray(n, 6) & "%", nKol1 + nKol2 + nKol3)
        End If
        'Print #1, lMarginStruk(nSpasi) & Padl(Format(vaArray(n, 3), "#,##0"), nKol1_2) & " x " & Padl(Format(nHargaArray, "#,###,##0"), nKol2_2) & " = " & Padl(Format(vaArray(n, 3) * nHargaArray, "#,###,##0"), nKol3_2)
        Print #1, lMarginStruk(nSpasi) & Padl(Format(vaArray(n, 3), "#,##0") & " x", nKol1) & Padl(Format(nHargaArray, "#,###,##0") & " =", nKol2) & Padl(Format(vaArray(n, 3) * nHargaArray, "#,###,##0"), nKol3)
      End If
    Next
    
    Print #1, lMarginStruk(nSpasi) & Replicate("-", nKol1 + nKol2 + nKol3)
    Print #1, lMarginStruk(nSpasi) & Format(nTotQty, "###,###,##0") & " Items"
    

'    Print #1, Padl("Sub   : ", 9); Padl(Format(nBruto, "###,###,##0"), 10)

    'Print #1, lMarginStruk(nSpasi) & Padl("Disc   : ", 9); Padl(Format(nDiscount.Value, "###,###,##0"), 10)
    If nHargaCOD.value > 0 Then
      Print #1, lMarginStruk(nSpasi + nTotKol - 9 - 10) & Padl("SubTTL: ", 9); Padl(Format(nSubTotal.value, "###,###,##0"), 10)
      Print #1, lMarginStruk(nSpasi + nTotKol - 9 - 10) & Padl("Ongkir: ", 9); Padl(Format(nHargaCOD.value, "###,###,##0"), 10)
    End If
    Print #1, lMarginStruk(nSpasi + nTotKol - 9 - 10) & Padl("Total : ", 9); Padl(Format(nTotal.value, "###,###,##0"), 10)
    
    If chkTunai.value = 1 Then
      'munculkan form kembalian
      If nKasirVoucher > 0 Then
        Print #1, lMarginStruk(nSpasi + nTotKol - 9 - 10) & Padl("Voucher: ", 9); Padl(Format(nKasirVoucher, "###,###,##0"), 10)
      End If
      Print #1, lMarginStruk(nSpasi + nTotKol - 9 - 10) & Padl("Bayar : ", 9); Padl(Format(nKasirBayar, "###,###,##0"), 10)
      Print #1, lMarginStruk(nSpasi + nTotKol - 9 - 10) & Padr("Kembali: ", 9); Padl(Format(nKasirKembalian, "###,###,##0"), 10)
    End If
    
    If chkTunai.value = 0 Then
'      Print #1, Padl("DP     : ", 9); Padl(Format(nDP.Value, "###,###,##0"), 10)
'      Print #1, Padl("Tunai  : ", 9); Padl(Format(nTunai.Value, "###,###,##0"), 10)
      'Print #1, Padl("Bon : ", 9); Padl(Format(nPiutang.Value, "###,###,##0"), 10)
      Print #1, lMarginStruk(nSpasi + nTotKol - 9 - 10) & Padl("Bon : ", 9); Padl(Format(nPiutang.value, "###,###,##0"), 10)

    End If
    'rata tengah
'    Print #1, Chr(27); Chr(33); Chr(4);
'    Print #1, Chr(27) & Chr(97) & Chr(1)
    'munculkan voucher/jika ada
    If vaVoucher.UpperBound(1) >= 0 Then
      Print #1, lMarginStruk(nSpasi) & "Voucher : "; ""
    End If
    
    For n = 0 To vaVoucher.UpperBound(1)
      If vaVoucher(n, 0) = -1 Then
        Print #1, lMarginStruk(nSpasi) & "*"; Padl(Format(vaVoucher(n, 4), "###,###,##0"), 7); " *"; vaVoucher(n, 3); ""
      End If
    Next n

    If Trim(cFoot2) <> "" Then
      Print #1, lMarginStruk(nSpasi) & (cFoot2)
    End If
    
    If Trim(cFoot3) <> "" Then
      Print #1, lMarginStruk(nSpasi) & Left(cFoot3, nLebarEfektif)
    End If
    
    If Trim(cFoot4) <> "" Then
      Print #1, lMarginStruk(nSpasi) & Left(cFoot4, nLebarEfektif)
    End If
    
    If Trim(cFoot5) <> "" Then
      Print #1, lMarginStruk(nSpasi) & Left(cFoot5, nLebarEfektif)
    End If
    
    
    For n = 1 To GetRegistry(reg_MarginBawah)
      Print #1, ""
    Next n
    
'    Print #1, Chr(29) ' feed kertas
'    Print #1, Chr$(&H1D); "V"; Chr$(66); Chr$(0);
'    Print #1, Chr$(27); Chr$(100); ; Chr$(51)
'    Print #1, Chr(29); Chr(86); Chr(66); Chr(20);
    Close #1
    
    
End Sub

Private Function lMarginStruk(ByVal n As Integer) As String
Dim i As Single
lMarginStruk = ""
  For i = 1 To n
    lMarginStruk = lMarginStruk & " "
  Next i
End Function


Private Function isValidSaving() As Boolean
Dim dba As New ADODB.Recordset
Dim nPernahBayar As Double
Dim n As Integer

isValidSaving = True
  
  'pastikan kolom anggota sudah diisi lengkap
  
  If vaArray.UpperBound(1) < 0 Then
    MsgBox "Nota kosong, data tidak disimpan", vbCritical, "Error"
    isValidSaving = False
    Exit Function
  End If
  
  If Trim(cFaktur.Text) = "" Then
     MsgBox "Maaf Nomor Faktur Kosong/Tidak Valid" & vbCrLf & "Data tidak bisa disimpan", vbCritical, "Error"
     isValidSaving = False
     Exit Function
  End If
  
  Set dba = objData.Browse(GetDSN, "anggota", "kodeanggota", "kodeanggota", sisAssign, cCustomer.Text)
  If dba.EOF Then
    If dba.RecordCount = 0 Then
      MsgBox "Maaf, Kode Anggota yang dimasukkan tidak ada dalam database komputer" & vbCrLf & "Data tidak bisa disimpan", vbCritical, "Error"
      isValidSaving = False
      Exit Function
    End If
  End If
  
  Set dba = objData.Browse(GetDSN, "kartupiutang k", "sum(k.debet-k.kredit) as saldopiutang,c.plafond", "k.kodeanggota", sisAssign, cCustomer.Text, " AND k.nomorkartupiutang <> '" & cFaktur.Text & "' GROUP BY k.kodeanggota", , Array("LEFT JOIN anggota c on c.kodeanggota = k.kodeanggota"))
  If Not dba.EOF Then
    If GetNull(dba!plafond) > 0 Then
      If GetNull(dba!saldopiutang) + nPiutang.value > GetNull(dba!plafond) Then
        isValidSaving = False
        MsgBox "Plafond melebihih quota" & vbCrLf & "Maaf, saldo piutang yg dimiliki telah melebihi yang ditetapkan." & vbCrLf & "Transaksi tidak bisa dilanjutkan/disimpan" & vbCrLf & "Saldo piutang maksimal yang diijinkan adalah Rp " & Format(GetNull(dba!plafond), "###,###,###,##0.00"), vbCritical, "Error"
        Exit Function
      End If
    End If
  Else
    Set dba = objData.Browse(GetDSN, "anggota", , "kodeanggota", sisAssign, cCustomer.Text)
    If Not dba.EOF Then
      If GetNull(dba!plafond) > 0 Then
        If nPiutang.value > GetNull(dba!plafond) Then
          MsgBox "Plafond melebihih quota" & vbCrLf & "Maaf, saldo piutang yg dimiliki telah melebihi yang ditetapkan." & vbCrLf & "Transaksi tidak bisa dilanjutkan/disimpan" & vbCrLf & "Saldo piutang maksimal yang diijinkan adalah Rp " & Format(GetNull(dba!plafond), "###,###,###,##0.00")
          isValidSaving = False
          Exit Function
        End If
      End If
    End If
  End If

  'cek validitas
  If Not GetValidDataBrowse(objData, "anggota", "kodeanggota", cCustomer.Text) Then
    MsgBox "Kode member tidak tertera dengan benar" & vbCrLf & "Data tidak bisa disimpan", vbCritical, "Error"
    cCustomer.SetFocus
    isValidSaving = False
    Exit Function
  End If
  
  If Not GetValidDataBrowse(objData, "salesman", "kodesalesman", cSalesman.Text) Then
'    MsgBox "Kode salesman tidak tertera dengan benar" & vbCrLf & "Data tidak bisa disimpan"
'    cSalesman.SetFocus
'    isValidSaving = False
'    Exit Function
    cSalesman.Text = ""
  End If
  
  If Not GetValidDataBrowse(objData, "gudang", "kodegudang", cGudang.Text) Then
    MsgBox "Kode gudang tidak tertera dengan benar" & vbCrLf & "Data tidak bisa disimpan"
    cGudang.SetFocus
    isValidSaving = False
    Exit Function
'  Else
'    Dim db As New ADODB.Recordset
'    Set db = objData.Browse(GetDSN, "gudang", "kodegudang,keterangan", "kodegudang", sisAssign, cGudang.Text)
'    If Not db.EOF Then
'      If MsgBox("Anda mengambil barang di " & GetNull(db!keterangan), vbYesNo) = vbNo Then
'        cGudang.SetFocus
'        isValidSaving = False
'        Exit Function
'      End If
'    End If
  End If
  
  If Not GetValidDataBrowse(objData, "akun", "kodeakun", cAkunKas.Text) Then
    MsgBox "Kode akun tidak tertera dengan benar" & vbCrLf & "Data tidak bisa disimpan"
    isValidSaving = False
    Exit Function
  End If
  
  If nPos = Edit Then
    If isPernahBayar(objData, cFaktur.Text, nPernahBayar) = True Then
      MsgBox "Transaksi ini sudah pernah dilunasi" & vbCrLf & "Data tidak bisa disimpan" & vbCrLf & "Hapus data pelunasan terlebih dahulu"
      isValidSaving = False
      Exit Function
    End If
    
  End If
  
  'cek apakah saldo minus diijinkan
  Dim cKodeMinus As String
  
  
'  cKodeMinus = ""
'  If aCfg(objData, msSaldoMinus) = 2 Then
'    If nPos = Edit Then
'      For n = vaArray.LowerBound(1) To vaArray.UpperBound(1)
'        If GetSaldoStock(objData, cGudang.Text, vaArray(n, 8)) + vaArray(n, 3) < vaArray(n, 3) Then
'          cKodeMinus = cKodeMinus & "  " & vaArray(n, 1) & "(" & GetSaldoStock(objData, cGudang.Text, vaArray(n, 8) + vaArray(n, 3)) - vaArray(n, 3) & ")"
'          isValidSaving = False
'        End If
'      Next n
'    End If
'
'    If nPos = Add Then
'      For n = vaArray.LowerBound(1) To vaArray.UpperBound(1)
'        If GetSaldoStock(objData, cGudang.Text, vaArray(n, 8)) < vaArray(n, 3) Then
'          cKodeMinus = cKodeMinus & "  " & vaArray(n, 1) & " (" & GetSaldoStock(objData, cGudang.Text, vaArray(n, 8)) - vaArray(n, 3) & ")"
'          isValidSaving = False
'        End If
'      Next n
'    End If
'
'    If Trim(cKodeMinus) <> "" Then
'      MsgBox "Ada barang dengan saldo stock minus" & vbCrLf & _
'             cKodeMinus
'    End If
'  End If
  
  'Jika kode gudang tidak valid, maka penyimpanan data tidak diijinkan
  Set dbData = objData.Browse(GetDSN, "gudang", "lstatus", "kodegudang", sisAssign, cGudang.Text)
  If Not dbData.EOF Then
    If GetNull(dbData!lStatus) <> "A" Then
      MsgBox "Kode Gudang tidak valid, atau tidak aktif, Data tidak bisa disimpan", vbExclamation
      isValidSaving = False
      Exit Function
    End If
  End If

  If nPos = Add Then
    Set dbData = objData.Browse(GetDSN, "totpenjualan", "nomorpenjualan", "nomorpenjualan", sisAssign, cFaktur.Text)
    If Not dbData.EOF Then
      MsgBox "No Faktur already exist", vbCritical, "Fatal Error"
      isValidSaving = False
    End If
  End If
  
  'pastikan semua faktur berisi kode
  For n = vaArray.LowerBound(1) To vaArray.UpperBound(1)
    If vaArray(n, 8) = "" Then
      MsgBox "Critical Error", vbCritical
      isValidSaving = False
    End If
  Next n
End Function

Private Sub cNama_ButtonClick()
Dim kdestock As String
Dim cWhere As String
Dim nLimitPencarian As Integer

  nLimitPencarian = IIf(GetRegistry(reg_LimitPencarian) <= 0, 10, GetRegistry(reg_LimitPencarian))
  SaveRegistry reg_LimitPencarian, nLimitPencarian
  If Trim(GetRegistry(reg_KodeGroupPenjualan)) <> "" Then
    cWhere = " and s.groupsales='" & GetRegistry(reg_KodeGroupPenjualan) & "'"
  End If
  
  If Len(cNama.Text) >= 3 Then
    'Set dbData = objData.Browse(GetDSN, "stock s", "s.Barcode,s.nama,s.hargajual,s.kodesatuan,s.kodestock,s.jenis,s.diskonpenjualan,s.bv,s.stok,s.kategori", "(s.nama", sisContent, cNama.Text, " or s.barcode like '%" & cNama.Text & "%') AND (s.statusnonaktif <> 1 and s.hargajual >0) " & cWhere, , , 0, GetRegistry(reg_LimitPencarian))
    Set dbData = objData.Browse(GetDSN, "stock s", "s.Barcode,s.nama,s.hargajual,s.kodesatuan,s.kodestock,s.jenis,s.diskonpenjualan,s.bv,s.stok,s.kategori", "(s.nama", sisContent, cNama.Text, " or s.barcode like '%" & cNama.Text & "%') AND (s.statusnonaktif <> 1 and s.hargajual >0) " & cWhere, , , 0, nLimitPencarian)
    If Not dbData.EOF Then
      cNama.Text = cNama.Browse(dbData, Array("BARCODE", "NAMA", "JUAL", "SATUAN"), , Array(13, 35, 10, 8))
      kdestock = GetNull(dbData!KodeStock)
      GetDataStock
      SumJumlah
    Else
      MsgBox "1. Masukkan 3 atau lebih karakter pencarian" & vbCrLf & "1. Harga jual harus lebih dari 0 Rupiah" & vbCrLf & "3. Status barang Masih Aktif", vbInformation, "Maaf Data Tidak Ketemu"
      cNama.Default
    End If
  End If
End Sub

Private Sub cNamaBarang2_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "stock", "barcode,nama,hargajual,hargabeli,kodestock", "konsi", sisAssign, 1, " and kodesupplier='" & cKodeSupplier.Text & "'")
  If Not dbData.EOF Then
    cNamaBarang2.Text = cNamaBarang2.Browse(dbData, Array("BARCODE", "NAMA", "JUAL", "BELI", "KODE"), , Array(15, 35, 15, 15, 15))
    cBarcode2.Text = GetNull(dbData!barcode)
    nHargaJual2.value = GetNull(dbData!HargaJual)
    nHargaBeli2.value = GetNull(dbData!hargabeli)
  End If
End Sub

Private Sub cNamaCOD_ButtonClick()
Set dbData = objData.Browse(GetDSN, "tarifcod", "kode,nama,harga", "kode", sisContent, cKodeCOD.Text, " or nama like '%" & cNamaCOD.Text & "%'")
  cNamaCOD.Text = cNamaCOD.Browse(dbData, Array("KODE", "NAMA", "HARGA"), "ONGKIR", Array(8, 35, 8))
  If Not dbData.EOF Then
    cKodeCOD.Text = GetNull(dbData!Kode, "")
    cNamaCOD.Text = GetNull(dbData!nama, "")
    nHargaCOD.value = GetNull(dbData!Harga, "")
  End If
  GetUpdateTotal
End Sub

Private Sub cNamaCustomer_ButtonClick()
Dim vaTmp As New XArrayDB

  If Len(cNamaCustomer.Text) >= 3 Then
    Set dbData = objData.Browse(GetDSN, "anggota a", "a.kodeanggota,a.nama,a.telp,a.alamat,a.kodedep,d.keterangan,a.whitelist", "a.nama", sisContent, cNamaCustomer.Text, " or a.kodeanggota like '%" & cNamaCustomer.Text & "%'", "a.kodeanggota,a.nama", Array("Left join dep d on d.kodedep = a.kodedep"), 0, 10)
    If Not dbData.EOF Then
      cNamaCustomer.Text = cNamaCustomer.Browse(dbData, Array("KODE", "NAMA", "TELP", "ALAMAT"), , Array(15, 30, 20, 15))
      cCustomer.Text = GetNull(dbData!kodeanggota)
      cNamaCustomer.Text = GetNull(dbData!nama, "")
      cAlamat.Text = GetNull(dbData!alamat, "")
      cKota.Text = GetNull(dbData!kodedep, "")
      cTelp = GetNull(dbData!telp)

      If nPos = Add Then
        Select Case GetNull(dbData!whitelist)
          Case 0
          'reguler
            If aCfg(objData, msBulanBlokir) > 0 Then
              If GetLewatBulan(objData, cCustomer.Text, 2) = True Then
                MsgBox "Maaf, customer ini tidak diperkenankan membuka nota baru. Masih ada nota jatuh tempo yg belum dilunasi. Terimakasih"
                initvalue
                GetEdit False
              End If
            End If
         'Case 1 = Whitelist
          Case 2
          'blokir
            MsgBox "Member ini di blokir/Non Aktif", vbCritical
            GetEdit False
            initvalue
            Exit Sub
        End Select
      
    '    Label7.Caption = "SALDO POIN HADIAH : " & GetPoinHadiahMember(objData, cCustomer.Text, dTgl.Value)
    '    If nPos = Add Then
    '      GetOderan cCustomer.Text
      End If
    End If
  Else
       MsgBox "Masukkan 3 atau lebih karakter pencarian", vbInformation
  End If
End Sub

Private Function GetLewatBulan(ByVal obj As CodeSuiteLibrary.Data, ByVal cKustomer As String, ByVal nBulan As Single) As Boolean
Dim cSQL As String
Dim db As New ADODB.Recordset

  GetLewatBulan = False
  
  cSQL = ""
  cSQL = " select * from totpenjualan t"
  cSQL = cSQL & " where t.flaglunas <> 1 and t.kodeanggota = '" & cKustomer & "' and t.tgl <= '" & Format(EOM(DateAdd("M", -nBulan, Date)), "yyyy-MM-dd") & "'"
  
  Set db = obj.SQL(GetDSN, cSQL)
  If Not db.EOF Then
    GetLewatBulan = True
  End If
  
End Function

Private Sub GetOderan(ByVal cKodeMember)
  Set dbData = objData.Browse(GetDSN, "totmemberorder t", "t.nomormemberorder,t.tgl,t.subtotal,t.dp,t.kodesalesman", "t.kodeanggota", sisAssign, cKodeMember, " AND t.status = 0")
  If Not dbData.EOF Then
    cmdAddOrder_Click
  End If
End Sub

Private Sub cNamaSupplier_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "supplier", "kodesupplier,nama", "kodesupplier", sisContent, cNamaSupplier.Text, " or nama like '%" & cNamaSupplier.Text & "'")
  If Not dbData.EOF Then
    cNamaSupplier.Text = cNamaSupplier.Browse(dbData, Array("KODE", "NAMA"), , Array(10, 35))
    cKodeSupplier.Text = GetNull(dbData!kodesupplier)
    cNamaSupplier.Text = GetNull(dbData!nama)
  End If
End Sub

Private Sub Command1_Click()
  If vaArray.UpperBound(1) >= 0 Then
    TDBGrid1_KeyDown vbKeyDelete, 1
  End If
End Sub

Private Sub cSalesman_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "salesman", "kodesalesman,nama")
  If Not dbData.EOF Then
    cSalesman.Text = cSalesman.Browse(dbData)
  End If
End Sub

Private Sub cSalesman_Validate(Cancel As Boolean)
  cSalesman.Enabled = False
End Sub

Private Sub dTgl_Validate(Cancel As Boolean)
  'If Not IsInPeriod(dTgl.Value) Or (dTgl.Value > DateAdd("d", 7, Date)) Or (dTgl.Value < DateAdd("d", -7, Date)) Then
  If Not IsInPeriod(dTgl.value) Then
'    Cancel = True
'    dTgl.SetFocus
'    GetEdit False
    'cek apakah tgl ini bisa di edit
    If Not GetTglEdit(dTgl.value) Then
      MsgBox "Maaf, transaksi untuk tgl tersebut tidak bisa dikoreksi karena sudah melewati proses audit", vbExclamation
      GetEdit False
      initvalue
    End If
  End If
End Sub

Private Function GetTglEdit(dTg As Date) As Boolean
Dim db As New ADODB.Recordset

  GetTglEdit = False
  Set db = objData.Browse(GetDSN, "tgledit", "tgl", "tgl", sisAssign, Format(dTg, "yyyy-MM-dd"))
  If Not db.EOF Then
    'jika tanggal sama maka boleh di edit
    GetTglEdit = True
  End If
End Function

Private Sub Form_Activate()
' Frame1.Left = (trPenjualan.ScaleWidth - Frame1.Width) * 0.5

Frame1.Left = (trPenjualan.Width / 2) - (Frame1.Width / 2)
Frame1.Top = (trPenjualan.Height / 2) - (Frame1.Height / 2)

' If GetRegistry(reg_OptGroupSales) <> 1 Then
'    SaveRegistry reg_KodeGroupPenjualan, ""
'    Label7.Caption = ""
' Else
'  If Trim(GetRegistry(reg_KodeGroupPenjualan)) <> "" Then
'     Label7.Caption = "GROUP SALES : " & GetRegistry(reg_KodeGroupPenjualan)
'  Else
'    Label7.Caption = ""
'  End If
' End If
  'jika modus compact diaktifkan
 Label7.Caption = "GROUP SALES : " & GetRegistry(reg_KodeGroupPenjualan)
 If GetModePenjualanUser(objData, GetRegistry(reg_Username)) <> 0 Then
  ModeCompact True
 End If
 Me.WindowState = vbMaximized
End Sub

Private Sub ModeCompact(ByVal lStatus As Boolean)
 BiSAFrame8.Visible = Not lStatus
 cCustomer.Text = "-"
 TabOne2.TabVisible(1) = Not lStatus
 TabOne1.TabVisible(1) = Not lStatus
 chkTunai.Enabled = Not lStatus
 chkTunai.value = 1
 nHarga.Enabled = Not lStatus
 nDisc1.Enabled = Not lStatus
 BiSAFrame7.Visible = Not lStatus
 BiSAFrame4.Visible = Not lStatus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  
  Select Case KeyCode
    Case vbKeyF1
      If nPos <> Add Or None <> Edit Or nPos <> Delete Then
        SaveRegistry reg_F1Key, Me.Name & "F1"
        cmdAdd_Click
        cBarcode.SetFocus
        chkTunai.value = 1
        'add tunai
      End If
    Case vbKeyF2
      If nPos <> Add Or None <> Edit Or nPos <> Delete Then
        SaveRegistry reg_F1Key, Me.Name & "F2"
        cmdAdd_Click
        cNamaCustomer.SetFocus
        chkTunai.value = 0
        'add bon
      End If
    Case vbKeyF4
      cmdSimpan_Click
      'simpan
    Case vbKeyF5
      cmdPending_Click
      'pending
    Case vbKeyF6
      If nPos = Add Then
        getDataPending
      End If
      'panggil pending
    Case vbKeyF7
      If nPos = Add Or nPos = Delete Then
        GetHapusPending
      End If
      'hapus pending
    Case vbKeyF3
      If nPos = Add Or nPos = Delete Then
        TDBGrid1.SetFocus
      End If
    Case vbKeyF8
      If nPos = Add Or nPos = Delete Then
        Load cfgLimitPencarian
        cfgLimitPencarian.Show vbModal
      End If
      'limit pencarian
  End Select
End Sub

Private Sub initvalue()
Dim dbgudang As New ADODB.Recordset
    
  Unload trFormKasir
  Label2.Caption = aCfg(objData, msNamaPerusahaan)
  lSign = 0
  Label5.Caption = "INFO STOK"
  Label6.Caption = "PERHATIAN, NOTA INI AKAN DIPROSES DENGAN MINIMUM DEPOSIT " & aCfg(objData, msMinimumDeposit) & "%"
  cNamaGudang.Default
  cFaktur.Default
  dTgl.value = Date
  dJthTmp.value = Date
  nPersDisc.value = 0
  cSalesman.Default
  nPPn.value = 0
  cFakturAsli.Default
  cCustomer.Default
  cNamaCustomer.Default
  cAlamat.Default
  cKota.Default
  nSubTotal.value = 0
  nPajak.value = 0
  nDiscount.value = 0
  nTotal.value = 0
  nTunai.value = 0
  nPiutang.value = 0
  chkTunai.value = 0
  cAkunKas.Text = cKasTeller
  nDP.Default
  nKomisi.Default
  cKodeCOD.Enabled = True
  cSalesman.Enabled = True
  cUp.Default
  cmdAddOrder.Enabled = True
  cNoOrder = ""
  optPromo(1).value = True
  
  nPoinReguler.value = 0
  cKeterangan.Text = ""
  BiSAButton1.Caption = "+"
  cGudang.Text = aCfg(objData, msGudangPenjualan)
  Set dbgudang = objData.Browse(GetDSN, "gudang", "kodegudang,keterangan", "kodegudang", sisAssign, cGudang.Text)
  If Not dbgudang.EOF Then
    cNamaGudang.Text = GetNull(dbgudang!keterangan)
  End If
  
  'cod
  cKodeCOD.Default
  cNamaCOD.Default
  nHargaCOD.Default
  
  If chkTunai.value = 1 Then
    Label1.Caption = "TUNAI"
  Else
    Label1.Caption = "BON"
  End If
  
  
  vaArray.ReDim 0, -1, 0, 11
  Set TDBGrid1.Array = vaArray
  TDBGrid1.ReBind
  InitValue1
  
  If aCfg(objData, msKolomHargaPenjualanNonTunai) = 1 Then
    nHarga.Enabled = True
  Else
    nHarga.Enabled = False
  End If
  
  
  TDBGrid1.Columns(3).FooterText = ""
  
  If GetKunciAkunKas(objData) Then
    cAkunKas.Enabled = False
  End If
  
  nQty.Decimals = aCfg(objData, msNilaiDecimals)

  nDisc1.Enabled = True
  nDisc1.BackColor = vbWhite
  cGudang.Enabled = True
  cGudang.BackColor = vbWhite
  If GetRegistry(reg_UserLevel) <> 0 Then
    'apakah diijinkan untuk kasi diskon per item nya?
    
    nDisc1.Enabled = False
    nDisc1.BackColor = vbButtonFace
    cGudang.Enabled = False
    cGudang.BackColor = vbButtonFace
  
  End If
  
  
  cGudang.Text = GetGudangUser(objData, GetRegistry(reg_Username))
  Set dbData = objData.Browse(GetDSN, "gudang", "kodegudang,keterangan", "kodegudang", sisAssign, cGudang.Text)
  If Not dbData.EOF Then
    cNamaGudang.Text = GetNull(dbData!keterangan)
  Else
    cNamaGudang.Default
  End If
  
  Label8.Caption = "Printer Not Configured"
  Select Case GetRegistry(reg_CetakanPenjualanNonTunai)
    Case 1:
      Label8.Caption = "Model Cetakan Nota NCR"
    Case 2:
      Label8.Caption = "Model Cetakan Nota Kertas Wartel"
    Case 3
      Label8.Caption = "Model Cetakan Struk/Thermal Printer"
  End Select

   BiSAFrame9.Enabled = True
  If GetRegistry(reg_UserLevel) <> 0 Then
    If aCfg(objData, msEditTransaksiPenjualan) = 2 Then
    'jika tidak bisa diedit
      BiSAFrame9.Enabled = False
    End If
  End If

  'kosongkan variable global
  nKasirTotalKartu = 0
  nKasirKodeKartu = 0
  nKasirFeeKartu = 0
  nKasirFeeTotalKartu = 0
  nKasirNoKartu = 0
  nKasirNoTraceKartu = 0
  nKasirNamaDiKartu = 0
  nDPKasir = 0
End Sub

Private Sub Form_Load()
Dim n As Single

'  If CheckTrial(nRecordsTrial, TrialPenjualan) = True Then
'    End
'  End If
  
  SetIcon Me.hWnd, "SIKD"
  TabOne1.TabVisible(1) = False
  TabOne1.TabEnabled(1) = False
  TabOne1 = 0
  
  Timer1.Interval = 400
'
  Me.Width = 1357 * Screen.TwipsPerPixelX
'  Me.Height = 634 * Screen.TwipsPerPixelY
'  Me.Height = Screen.Height
'
'  CenterForm Me
  'CenterChild aMainmenu, Me
  GetEdit False
 

  initvalue
  
  If objMenu.UserLevel <> 0 Then
    If aCfg(objData, msHapusTransaksiPenjualan) = 2 Then
      cmdHapus.Visible = False
    End If
  End If
  
  'Label1.Caption = ""
  'Label1.Caption = aCfg(objData, msNamaPerusahaan)
  'Label3.Caption = aCfg(objData, msAlamatPerusahaan)
'  Set dbData = objData.Browse(GetDSN, "costcenter", , "kodecostcenter", sisAssign, aCfg(objData, msCostCenterJualBeli))
'  If Not dbData.EOF Then
'    Frame2.Caption = "Cost Centre : " & GetNull(dbData!keterangan)
'  End If
  Frame2.Caption = GetCostCenterUser(objData, GetRegistry(reg_Username))
  
  'LOAD FUNGSI SYSTRY
  Dim tmp
    
  'tmp = RegRead(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\", "EnableBalloonTips")
  tmp = RegRead(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\", "EnableBalloonTips")
  
  If tmp = 0 Then
      If MsgBox("Balloon tips are currently disabled on your computer. Would you like to enable them?", vbQuestion + vbYesNo, "Enable Balloon Tips?") = vbYes Then
          WriteDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\", "EnableBalloonTips", 1
          If MsgBox("Balloon tips are now enabled, but you must first logoff your computer" & vbNewLine & "and then log back on before the changes will take effect." & vbNewLine & vbNewLine & "Would you like to be logged off now?", vbQuestion + vbYesNo, "Logoff Now?") = vbYes Then
              LogOffNT True
              End
          End If
      Else
          MsgBox "Without balloon tips enabled on your computer, this program will not function properly.", vbExclamation, "Balloon Tips Disabled"
      End If
  End If
  
  
  TabIndex dTgl, n
  TabIndex cGudang, n
  TabIndex cNamaGudang, n
  TabIndex cAkunKas, n
  TabIndex cCustomer, n
  TabIndex cNamaCustomer, n
  TabIndex cmdAddOrder, n
  TabIndex cAlamat, n
  TabIndex cFaktur, n
  
  TabIndex optPromo(0), n
  TabIndex optPromo(1), n
  
  'TabIndex dJthTmp, n
  'TabIndex nPersDisc, n
  'TabIndex nPPn, n
  
  'TabIndex cSalesman, n
  'TabIndex nKomisi, n
  'TabIndex cUp, n
  'TabIndex cFakturAsli, n
  
  TabIndex nNomor, n
  TabIndex cBarcode, n
  TabIndex cNama, n
  TabIndex nQty, n
  TabIndex cSatuan, n
  TabIndex nHarga, n
  TabIndex nDisc1, n
  TabIndex nJumlah, n
  TabIndex cmdOK, n
  TabIndex nTunai, n
  
  TabIndex nNo2, n
  TabIndex cNamaSupplier, n
  TabIndex cNamaBarang2, n
  TabIndex nQty2, n
  TabIndex nHargaJual2, n
  TabIndex nHargaBeli2, n
  TabIndex cmdOK2, n

  
  TabIndex cmdAdd, n
  TabIndex cmdEdit, n
  TabIndex cmdHapus, n
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
  TabIndex cmdAktivasi, n
  
  If GetRegistry(reg_TampilNotifikasi) <> 1 Then
    Label5.Visible = False
    BiSAFrame10.Visible = False
  End If

'  aMainmenu.EnableMaxButton aMainmenu.hWnd, False
'  aMainmenu.mnuWindowFullScreen.Caption = "Normal Screen"
'  aMainmenu.WindowState = vbMaximized
End Sub

Private Sub Form_Unload(Cancel As Integer)
  aMainmenu.EnableMaxButton aMainmenu.hWnd, True
  aMainmenu.mnuWindowFullScreen.Caption = "Full Screen"
  aMainmenu.WindowState = vbMaximized
End Sub

Private Sub nDisc1_Change()
  SumJumlah
End Sub

Private Sub nDisc1_Validate(Cancel As Boolean)
  SumJumlah
End Sub

Private Sub nDisc2_Validate(Cancel As Boolean)
  SumJumlah
End Sub

Private Sub nDiscount_Validate(Cancel As Boolean)
  SumTotal
End Sub

Private Sub nDiscount2_Validate(Cancel As Boolean)
  SumTotal
End Sub

Private Sub InitValue1()
  nNomor.value = 1
  nNo2.value = 1
  cBarcode.Default
  cNama.Default
  nQty.value = 1
  nQty2.value = 1
  cSatuan.Default
  nHarga.value = 0
  nJumlah.value = 0
  cKode = ""
  nInfoStok.Default
  Shape1.FillColor = vbWhite
  Label3.BackColor = vbWhite
  Label3.ForeColor = vbBlack
End Sub

Private Sub GetEdit(lPar As Boolean)

  cNamaCustomer.Enabled = True
  cNamaCustomer.Button = True
  'Frame1.Enabled = lPar
  
  Frame2.Enabled = lPar
  BiSAFrame1.Enabled = lPar
  BisaFrame2.Enabled = lPar
  
  
  lEdit = lPar
  initvalue
  SetButton cmdSimpan, cmdKeluar, cmdAdd, cmdEdit, cmdHapus, nPos, lPar, cmdAktivasi
  
  If lPar Then
    dTgl.SetFocus
    If nPos = Add Then
      cFaktur.Enabled = False
      cFaktur.BackColor = vbButtonFace
    Else
      cFaktur.Enabled = True
      cFaktur.BackColor = vbWindowBackground
      cFaktur.CaptionBackColor = vbButtonFace
    End If
  End If
  TabOne1 = 0
  TabOne2 = 0
  TabOne1.Enabled = lPar
  TabOne2.Enabled = lPar
End Sub

Private Sub nBiaya_Validate(Cancel As Boolean)
  SumTotal
End Sub

Private Sub SumJumlah()
Dim nSubJumlah As Double

  nSubJumlah = nHarga.value * nQty.value
  nSubJumlah = nSubJumlah - (nSubJumlah * (nDisc1.value / 100))
  nJumlah.value = nSubJumlah
End Sub

Private Sub nHarga_Change()
  SumJumlah
End Sub

Private Sub nHarga_Validate(Cancel As Boolean)
  SumJumlah
End Sub

Private Sub nHargaJual2_Validate(Cancel As Boolean)
  nJumlah2.value = nQty2.value * nHargaJual2.value
End Sub

Private Sub nInfoStok_Change()
  If nInfoStok.value <= 0 Then
    Shape1.FillColor = vbRed
    Label3.BackColor = vbRed
    Label3.ForeColor = vbWhite
  Else
    Shape1.FillColor = vbGreen
    Label3.BackColor = vbGreen
    Label3.ForeColor = vbBlack
  End If
End Sub

Private Sub nNomor_Validate(Cancel As Boolean)
Dim n As Single
  
  If GetValidNomorUrut(nNomor, vaArray) Then
    n = nNomor.value - 1
    If n <= vaArray.UpperBound(1) Then
      cBarcode.Text = vaArray(n, 1)
      cNama.Text = vaArray(n, 2)
      nQty.value = vaArray(n, 3)
      cSatuan.Text = vaArray(n, 4)
      nHarga.value = GetHargaJ(objData, vaArray(n, 8))
      nDisc1.value = vaArray(n, 6)
      nJumlah.value = vaArray(n, 7)
      cKode = vaArray(n, 8)
      cJenis = vaArray(n, 9)
      nBValue = vaArray(n, 10)
      'Label5.Caption = "STOK : " & cNama.Text & " : " & vaArray(n, 11)
    End If
  End If
End Sub

Private Sub nPersDisc_Validate(Cancel As Boolean)
  SumTotal
End Sub

Private Sub nPersDisc2_Validate(Cancel As Boolean)
  SumTotal
End Sub

Private Sub nPPn_Validate(Cancel As Boolean)
  SumTotal
End Sub

Private Sub nQty_Change()
  SumJumlah
End Sub

Private Sub nQty_Validate(Cancel As Boolean)
Dim nNewHarga As Double

  SumJumlah
  If nQty.value <= 0 Then
    MsgBox "Qty Salah, qty harus lebih dari atau sama dengan 1", vbCritical
    Cancel = True
    nQty.SetFocus
  End If
End Sub

Private Sub nQty2_Validate(Cancel As Boolean)
  nJumlah2.value = nQty2.value * nHargaJual2.value
End Sub

Private Sub nTotal_Change()
  TabOne2 = 0
  BiSANumberBox1.value = nTotal.value
End Sub

Private Sub nTunai_Validate(Cancel As Boolean)
  SumBayar
End Sub

Private Sub optPromo_Click(Index As Integer)
'  MsgBox Index
  Select Case Index
    Case 0
      cMasterKatalog.Visible = True
      cNamaKatalog.Visible = True
    Case 1
      cMasterKatalog.Visible = False
      cNamaKatalog.Visible = False
  End Select
End Sub

Private Sub optPromo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    SendKeysA vbKeyTab, True
  End If
End Sub

Private Sub TDBGrid1_Click()
Dim nCo As Integer

 'GetNull(TDBGrid1.Columns(0).value)
  If vaArray.UpperBound(1) >= 0 Then
    nCo = GetNull(TDBGrid1.Columns(0).value)
'    nStockSelected = vaArray(TDBGrid1.Columns(0).Value - 1, 11)
'    Label5.Caption = "STOK : " & TDBGrid1.Columns(2).Text & " : " & vaArray(TDBGrid1.Columns(0).Value - 1, 11)
   
    If nCo < 1 Then
      nCo = 1
    End If
    
    'If GetNull(TDBGrid1.Columns(0).value) > 0 Then
      nStockSelected = GetSaldoStock_DitTableStock(vaArray(nCo - 1, 8))
      'Label5.Caption = "STOK : " & TDBGrid1.Columns(2).Text & " : " & nStockSelected
      Label5.Caption = "STOK : " & GetNamaBarang(vaArray(nCo - 1, 8)) & " : " & nStockSelected
    'End If
  Else
    Label5.Caption = "INFO STOK"
  End If
  On Error Resume Next
End Sub

Private Sub TDBGrid1_DblClick()
  nNomor.value = TDBGrid1.Columns(0).Text
  nNomor_Validate True
  nQty.SetFocus
End Sub

Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim n As Integer
Dim nQtyTmp As Integer

On Error Resume Next

  If KeyCode = vbKeyDelete Then
  
'    If aCfg(objData, msOtorisasiPenuh) = "Y" Then
      
'      If objMenu.GetPassword("", Me, GetDSN) Then
'        If objMenu.UserLevel <> 0 Then
'            MsgBox "Maaf, Anda tidak diberikan wewenang untuk melakukan PENGEDITAN." & vbCrLf & _
'                   "Hanya user dengan LEVEL 0 (SUPERVISOR) yg diijinkan", vbInformation, "OTORISASI not ALLOWED"
'            Exit Sub
'        End If
'
'        If vaArray.UpperBound(1) >= 0 Then
'          TDBGrid1.Delete
'          TDBGrid1.Update
'          SumTotal
'          For n = 0 To vaArray.UpperBound(1)
'            vaArray(n, 0) = n + 1
'            nQtyTmp = nQtyTmp + vaArray(n, 3)
'          Next
'          TDBGrid1.Columns(3).FooterText = Format(nQtyTmp, "###,###,##0")
'          nNomor.value = vaArray.UpperBound(1) + 2
'          nPoinReguler.value = GetHitungPoinHadiah(aCfg(objData, msKelipatan))
'          TDBGrid1.ReBind
'        End If
'
'        If vaArray.UpperBound(1) < 0 Then
'          cNamaCustomer.Enabled = True
'          cNamaCustomer.Button = True
'        End If
'      End If
'
'    Else 'otorisasi
'      Exit Sub
'    End If
  
      If aCfg(objData, msKunciKasirDelete) = "Y" Then
        If vaArray.UpperBound(1) >= 0 Then
          TDBGrid1.Delete
          TDBGrid1.Update
          SumTotal
          For n = 0 To vaArray.UpperBound(1)
            vaArray(n, 0) = n + 1
            nQtyTmp = nQtyTmp + vaArray(n, 3)
          Next
          TDBGrid1.Columns(3).FooterText = Format(nQtyTmp, "###,###,##0")
          nNomor.value = vaArray.UpperBound(1) + 2
          nPoinReguler.value = GetHitungPoinHadiah(aCfg(objData, msKelipatan))
          TDBGrid1.ReBind
        End If
        
        If vaArray.UpperBound(1) < 0 Then
          cNamaCustomer.Enabled = True
          cNamaCustomer.Button = True
        End If
      Else
        MsgBox "Maaf. Penghapusan Tidak diijinkan", vbExclamation
      End If

  
    
  End If 'vbkeydelete
  
  If lEdit = True Then
    If KeyCode = vbKeyF3 Then
        If vaArray.UpperBound(1) >= 0 Then
          nNomor.value = TDBGrid1.Columns(0).Text
          nNomor_Validate True
          nQty.SetFocus
        End If
    End If
    If KeyCode = vbKeyReturn Then
        If vaArray.UpperBound(1) >= 0 Then
          nNomor.value = TDBGrid1.Columns(0).Text
          nNomor_Validate True
          nQty.SetFocus
        End If
    End If
  End If
  
  If KeyCode = vbKeyEscape Then
    InitValue1
    nNomor.value = vaArray.UpperBound(1) + 2
    cBarcode.SetFocus
    nPoinReguler.value = GetHitungPoinHadiah(aCfg(objData, msKelipatan))
  End If
End Sub

Private Function GetUpdateKartuStokPaket(ByVal obj As CodeSuiteLibrary.Data, ByVal cStokPaket As String, ByVal nQtyPaket As Integer, ByVal lDebet As Boolean, ByVal lKredit As Boolean) As Boolean
Dim db As New ADODB.Recordset

  'cek dulu apakah kode sKontrak merupakan kode kontrak
  Set db = obj.Browse(GetDSN, "stokpaket", "kodestock,qty", "kodeisipaket", sisAssign, cStokPaket)
  If Not db.EOF Then
    If lDebet = True Then
    End If
  End If

End Function

Private Function GetHitungPoinHadiah(ByVal nPembagi As Double) As Double
Dim n As Single
Dim nPoinHadiahLangsung As Double

  nPoinHadiahLangsung = 0
  For n = 0 To vaArray.UpperBound(1)
    'ambil yg ada diskon reguler nya
    nPoinHadiahLangsung = nPoinHadiahLangsung + (vaArray(n, 3) * vaArray(n, 5))
    GetHitungPoinHadiah = nPoinHadiahLangsung \ nPembagi
  Next n
End Function


Private Sub Timer1_Timer()
  'Label5.Visible = True
  If nStockSelected <= 0 And Label5.Caption <> "INFO STOK" Then
    Label5.Visible = Not Label5.Visible
    Label5.ForeColor = vbRed
  Else
    Label5.Visible = True
    Label5.ForeColor = vbBlack
  End If
End Sub

Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim nCo As Double

  'GetNull(TDBGrid1.Columns(0).value)
  If vaArray.UpperBound(1) >= 0 Then
    nCo = GetNull(TDBGrid1.Columns(0).value)
'    nStockSelected = vaArray(TDBGrid1.Columns(0).Value - 1, 11)
'    Label5.Caption = "STOK : " & TDBGrid1.Columns(2).Text & " : " & vaArray(TDBGrid1.Columns(0).Value - 1, 11)
    If nCo < 1 Then
      nCo = 1
    End If
    'If GetNull(TDBGrid1.Columns(0).value) > 0 Then
      nStockSelected = GetSaldoStock_DitTableStock(vaArray(nCo - 1, 8))
     ' Label5.Caption = "STOK : " & TDBGrid1.Columns(2).Text & " : " & nStockSelected
      Label5.Caption = "STOK : " & GetNamaBarang(vaArray(nCo - 1, 8)) & " : " & nStockSelected
'      TDBGrid1.Update
'      MsgBox TDBGrid1.Columns(2).Text
    'End If
  Else
    Label5.Caption = "INFO STOK"
  End If
  On Error Resume Next
End Sub

Private Function GetSaldoStock_DitTableStock(ByVal cKodeStock As String) As Double
Dim db As New ADODB.Recordset
  
  GetSaldoStock_DitTableStock = 0
  Set db = objData.Browse(GetDSN, "stock", "stok", "kodestock", sisAssign, cKodeStock)
  If Not db.EOF Then
    GetSaldoStock_DitTableStock = GetNull(db!stok)
  End If
End Function
