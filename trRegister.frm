VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form trRegister 
   BorderStyle     =   0  'None
   Caption         =   "Masukkan Serial Number"
   ClientHeight    =   5532
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5532
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   4770
      Left            =   75
      Top             =   60
      Width           =   4605
      _ExtentX        =   8128
      _ExtentY        =   8424
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   4
      BackColor       =   -2147483633
      Begin BiSADateProject.BiSADate dValid 
         Height          =   405
         Left            =   1620
         TabIndex        =   6
         Top             =   1245
         Width           =   1365
         _ExtentX        =   2413
         _ExtentY        =   720
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
         ForeColor       =   -2147483640
         Enabled         =   0   'False
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin BiSATextBoxProject.BiSATextBox cIDKey 
         Height          =   360
         Left            =   1620
         TabIndex        =   4
         Top             =   465
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   635
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Verdana"
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin BiSAButtonProject.BiSAButton BiSAButton1 
         Height          =   510
         Left            =   255
         TabIndex        =   0
         Top             =   2730
         Width           =   4005
         _ExtentX        =   7070
         _ExtentY        =   910
         Caption         =   "Register"
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
      End
      Begin BiSATextBoxProject.BiSATextBox cSerialKey 
         Height          =   390
         Left            =   255
         TabIndex        =   1
         Top             =   2265
         Width           =   4005
         _ExtentX        =   7070
         _ExtentY        =   677
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Verdana"
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin BiSATextBoxProject.BiSATextBox cTokenAktif 
         Height          =   360
         Left            =   1650
         TabIndex        =   5
         Top             =   855
         Width           =   2610
         _ExtentX        =   4593
         _ExtentY        =   635
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Verdana"
         BackColor       =   -2147483633
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label9 
         Caption         =   "Wa /Tele : 081999962828 - 081338414828"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   255
         TabIndex        =   13
         Top             =   4470
         Width           =   3795
      End
      Begin VB.Label Label8 
         Caption         =   "Hendra Suparyawan"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   255
         TabIndex        =   12
         Top             =   4245
         Width           =   1905
      End
      Begin VB.Label Label7 
         Caption         =   "Untuk Pembelian Token Silahkan Hub : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   270
         TabIndex        =   11
         Top             =   4005
         Width           =   3135
      End
      Begin VB.Label Label6 
         Caption         =   "Terimkasih sudah menggunakan Aplikasi ini. Mohon dukungan nya untuk tidak menggunakan program Bajakan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   270
         TabIndex        =   10
         Top             =   3285
         Width           =   3135
      End
      Begin VB.Label Label5 
         Caption         =   "Valid Until"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   315
         TabIndex        =   9
         Top             =   1320
         Width           =   930
      End
      Begin VB.Label Label4 
         Caption         =   "Active Token"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   315
         TabIndex        =   8
         Top             =   945
         Width           =   1260
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0FF&
         Caption         =   "ACTIVE SUBSCRIPTION :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   7
         Top             =   120
         Width           =   4032
      End
      Begin VB.Label Label2 
         Caption         =   "ID Key"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   315
         TabIndex        =   3
         Top             =   540
         Width           =   930
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080FF80&
         Caption         =   "MASUKKAN NEW SERIAL KEY :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   2
         Top             =   1872
         Width           =   3960
      End
   End
   Begin BiSAButtonProject.BiSAButton BiSAButton2 
      Height          =   510
      Left            =   2700
      TabIndex        =   14
      Top             =   4890
      Width           =   1935
      _ExtentX        =   3408
      _ExtentY        =   910
      Caption         =   "Exit"
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483633
   End
End
Attribute VB_Name = "trRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim nPos As SisPos
Dim lEdit As Boolean
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB
Dim objMenu As New CodeSuiteLibrary.Menu

Private Sub BiSAButton1_Click()
Dim sSecret     As String
Dim cKey As String
Dim db As New ADODB.Recordset
Dim dTgl As Date

On Error GoTo err
    
    sSecret = cSerialKey.Text
    cKey = GetRegistry(reg_KeySecret)
    dTgl = CryptRC4(FromHexDump(sSecret), cKey)
    If IsDate(CryptRC4(FromHexDump(sSecret), cKey)) Then
      'masukkan ke dalam tabel keyapp
      objData.Delete GetDSN, "keyapp", "tokenapp", sisAssign, sSecret
      objData.Add GetDSN, "keyapp", Array("tokenapp", "keyapp", "tgl"), Array(sSecret, cKey, Format(dTgl, "yyyy-MM-dd"))
      MsgBox "Trimakasih Token Berhasil Dimasukkan" & vbCrLf & "Program akan expired lagi tgl " & Format(dTgl, "dd-MM-yyyy"), vbInformation
      Unload aMainmenu
      aMainmenu.Show
    Else
      MsgBox "err", vbCritical, "err"
    End If
    
Exit Sub
err:
MsgBox "ERROR, TOKEN GAGAL DIMASUKKAN", vbCritical, "ERR"
End Sub

Private Sub BiSAButton2_Click()
  End
End Sub

Private Sub Form_Load()
Dim n As Single
Dim cKey As String

  SetIcon Me.hWnd, "SIKD"
  CenterForm Me
  cKey = GetRegistry(reg_KeySecret)
  Label2.Caption = "ID KEY : "
  cIDKey.Text = cKey
  If GetToken(objData) <> "NULL" Then
    cTokenAktif.Text = GetToken(objData)
    If IsDate(CryptRC4(FromHexDump(GetToken(objData)), GetRegistry(reg_KeySecret))) = True Then
      dValid.Value = CryptRC4(FromHexDump(GetToken(objData)), GetRegistry(reg_KeySecret))
    Else
      dValid.Value = "0000-00-00)"
    End If
  End If
  'MsgBox "Token anda akan expired lagi " & GetNewExpiredApp & " Hari Tgl. " & Format(DateAdd("d", cTokenAktif.Text, Date), "dd-MMM-yyyy"), vbCritical
End Sub
