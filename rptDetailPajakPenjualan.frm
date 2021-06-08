VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form rptDetailPajakPenjualan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Laporan Detail Pajak Penjualan"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7545
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   7545
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1245
      Left            =   15
      Top             =   0
      Width           =   7530
      _ExtentX        =   13282
      _ExtentY        =   2196
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
      Begin BiSADateProject.BiSADate dTgl 
         Height          =   330
         Index           =   0
         Left            =   420
         TabIndex        =   0
         Top             =   360
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   582
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
         CaptionWidth    =   2000
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
      Begin BiSADateProject.BiSADate dTgl 
         Height          =   330
         Index           =   1
         Left            =   4035
         TabIndex        =   1
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
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
         Caption         =   "s.d"
         CaptionWidth    =   500
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
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   630
      Left            =   15
      Top             =   1230
      Width           =   7530
      _ExtentX        =   13282
      _ExtentY        =   1111
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
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   6360
         TabIndex        =   2
         Top             =   120
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
         Picture         =   "rptDetailPajakPenjualan.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   5925
         TabIndex        =   3
         Top             =   120
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
         Picture         =   "rptDetailPajakPenjualan.frx":00A6
      End
   End
End
Attribute VB_Name = "rptDetailPajakPenjualan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdPreview_Click()
Dim cSQL As String
Dim n As Integer

  vaArray.ReDim 0, -1, 0, 5
  
  cSQL = " select p.tgl,p.kodestock,s.nama,p.qty,p.hb,p.harga,p.harga-p.hb as saldo from penjualan p "
  cSQL = cSQL & " left join stock s on s.kodestock = p.kodestock"
  cSQL = cSQL & " Where p.tgl >= '" & Format(dTgl(0).Value, "yyyy-MM-dd") & "' And p.tgl <= '" & Format(dTgl(1).Value, "yyyy-MM-dd") & "'"
  cSQL = cSQL & " order by p.tgl"
  
  Set dbData = objData.Sql(GetDSN, cSQL)
  If Not dbData.EOF Then
    
    Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = GetNull(dbData!nama)
      vaArray(n, 1) = GetNull(dbData!qty)
      vaArray(n, 2) = GetNull(dbData!hb)
      vaArray(n, 3) = GetNull(dbData!Harga)
      vaArray(n, 4) = GetNull(dbData!saldo)
      vaArray(n, 5) = (GetNull(dbData!saldo) * 10 / 100) * vaArray(n, 1)
      dbData.MoveNext
    Loop
    
    With FrmRPT
      .AddPageHeader UCase("Daftar Detail Barang Kena Pajak"), tdbHalignCenter, , , True, dbArial, 12, True, , , False
      .AddPageHeader aCfg(objData, msNamaPerusahaan), tdbHalignCenter, , , True, , 14
      .AddPageHeader "", , , , True
      .AddPageHeader "", , , , True
            
      .AddTableHeader "Nama"
      .AddTableHeader "Qty", , , , 7
      .AddTableHeader "H.Beli", , , , 13
      .AddTableHeader "H.Jual", , , , 13
      .AddTableHeader "Margin", , , , 13
      .AddTableHeader "Pajak", , , , 10
      
      .AddTableBody
      .AddTableBody Sis_Rpt_Number2, tdbHalignRight
      .AddTableBody Sis_Rpt_Number2, tdbHalignRight
      .AddTableBody Sis_Rpt_Number2, tdbHalignRight
      .AddTableBody Sis_Rpt_Number2, tdbHalignRight
      .AddTableBody Sis_Rpt_Number2, tdbHalignRight
         
      .AddTableFooter "Total", , tdbHalignRight, , , , , , , , , , , , 4
      .AddTableFooter
      .AddTableFooter
      .AddTableFooter
      .AddTableFooter "&Sum", Sis_Rpt_Number2
      .AddTableFooter "&Sum", Sis_Rpt_Number2
      
      .Preview vaArray, True
    End With
    
  End If
End Sub


Private Sub Form_Load()
Dim n As Single

  SetIcon Me.hWnd, "SIKD"
  CenterForm Me
  TabIndex cmdPreview, n
  TabIndex cmdKeluar, n
End Sub
