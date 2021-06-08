VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Begin VB.Form trSelectGroupSales 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   1785
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1680
      Left            =   45
      Top             =   30
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   2963
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
      Begin BiSATextBoxProject.BiSABrowse cGroupSales 
         Height          =   330
         Left            =   435
         TabIndex        =   0
         Top             =   420
         Width           =   3135
         _ExtentX        =   5530
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
         Caption         =   "Group Sales"
         CaptionWidth    =   1200
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
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   1455
         TabIndex        =   1
         Top             =   1035
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   767
         Caption         =   "    &Save/OK"
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
         Picture         =   "trSelectGroupSales.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   2925
         TabIndex        =   2
         Top             =   1035
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
         Picture         =   "trSelectGroupSales.frx":0286
      End
   End
End
Attribute VB_Name = "trSelectGroupSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objData As New CodeSuiteLibrary.Data
Dim dbData As New ADODB.Recordset
Public sisModul As enum_ModOpen

Public Enum enum_ModOpen
  enum_OpenPenjualan = 1
  enum_OpenPelunasanPiutang = 2
  enum_OpenPembelian = 3
End Enum

Private Sub cGroupSales_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "groupsales", "kode,keterangan", "status", sisAssign, 1)
  If Not dbData.EOF Then
    If dbData.RecordCount > 1 Then
      cGroupSales.Text = cGroupSales.Browse(dbData)
    End If
    cGroupSales.Text = GetNull(dbData!Kode)
  Else
    cGroupSales.Default
  End If
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdSimpan_Click()
  If lExist(objData, "groupsales", "kode", cGroupSales.Text, " and status=1") = True Then
    
    If sisModul = enum_OpenPenjualan Then
      If GetRegistry(reg_LimitPencarian) < 1 Then
        Load cfgLimitPencarian
        cfgLimitPencarian.Show vbModal
      End If
      
      SaveRegistry reg_KodeGroupPenjualan, cGroupSales.Text
      GetGroupSalesPenjualan = cGroupSales.Text
      Unload Me
      aMainmenu.MembukaModulPenjualan
    End If
    If sisModul = enum_OpenPelunasanPiutang Then
      Unload Me
      aMainmenu.MembukaModulPelunasan
    End If
    If sisModul = enum_OpenPembelian Then
      SaveRegistry reg_KodeGroupSalesPembelian, cGroupSales.Text
      Unload Me
      aMainmenu.MembukaModulPembelian
    End If
  Else
    MsgBox "Masukkan Kode Group Sales yg benar", vbExclamation
  End If
  
End Sub

Private Sub Form_Load()
Dim n As Single
  
  SetIcon Me.hWnd, "SIKD"
  'initvalue
  
  CenterForm Me
'  If GetGroupSales(GetRegistry(reg_KodeGroupSales)) = True Then
'    cGroupSales.Text = GetRegistry(reg_KodeGroupSales)
'  Else
'    cGroupSales.Text = ""
'  End If
'
 ' cGroupSales.Text = IIf(GetGroupSales(GetRegistry(reg_KodeGroupSales)) = True, GetRegistry(reg_KodeGroupSales), "")
  TabIndex cGroupSales, n
  TabIndex cmdSimpan, n
End Sub



 
