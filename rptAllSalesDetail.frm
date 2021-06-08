VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form rptAllSalesDetail 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "All Item Sales - Detail"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7560
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   7560
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1140
      Left            =   15
      Top             =   0
      Width           =   7530
      _ExtentX        =   13282
      _ExtentY        =   2011
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
         Left            =   330
         TabIndex        =   0
         Top             =   345
         Width           =   3465
         _ExtentX        =   6112
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
         Left            =   4185
         TabIndex        =   1
         Top             =   345
         Width           =   2010
         _ExtentX        =   3545
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
      Left            =   0
      Top             =   1125
      Width           =   7545
      _ExtentX        =   13309
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
         Picture         =   "rptAllSalesDetail.frx":0000
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
         Picture         =   "rptAllSalesDetail.frx":00A6
      End
   End
End
Attribute VB_Name = "rptAllSalesDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vaArray As New XArrayDB
Dim objData As New CodeSuiteLibrary.Data
Dim dbData As New ADODB.Recordset

Private Sub GetDataAll()
Dim cSQL As String
Dim n As Integer
  
  vaArray.ReDim 0, -1, 0, 5
  
  cSQL = cSQL & " select p.tgl,p.kodestock,s.nama,sum(p.qty) as qty,sum(p.qty*p.tunai) as tunai,sum(p.qty*p.piutang) as piutang from penjualan p"
  cSQL = cSQL & " left join stock s on s.kodestock = p.kodestock"
  cSQL = cSQL & " WHERE p.tgl >= '" & Format(dTgl(0).Value, "yyyy-MM-dd") & "' AND p.tgl <= '" & Format(dTgl(1).Value, "yyyy-MM-dd") & "'"
  cSQL = cSQL & " group by p.tgl,p.kodestock"
  
  Set dbData = objData.Sql(GetDSN, cSQL)
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = Format(GetNull(dbData!tgl), "dd-MM-yyyy")
      vaArray(n, 1) = GetNull(dbData!KodeStock)
      vaArray(n, 2) = GetNull(dbData!nama)
      vaArray(n, 3) = GetNull(dbData!qty)
      vaArray(n, 4) = GetNull(dbData!Tunai)
      vaArray(n, 5) = GetNull(dbData!Piutang)
      dbData.MoveNext
    Loop
  
    With FrmRPT
      .AddPageHeader "Detail Item penjualan", tdbHalignCenter, , , , , 10, True
      .AddPageHeader "Tgl " & Format(dTgl(0).Value, "dd-MM-yyyy") & " - " & Format(dTgl(1).Value, "dd-MM-yyyy"), tdbHalignCenter, , , True, , , True
      .AddPageHeader aCfg(objData, msNamaPerusahaan), tdbHalignCenter, , , True, , 14, True
      .AddPageHeader "", tdbHalignCenter, , , True
      .AddPageHeader "", tdbHalignCenter, , , True
      
      .AddTableHeader "Tgl ", , , , 10
      .AddTableHeader "Kode", , , , 6
      .AddTableHeader "Nama"
      .AddTableHeader "qty", , , , 8
      .AddTableHeader "Tunai", , , , 13
      .AddTableHeader "Kredit", , , , 13
            
      .AddTableBody
      .AddTableBody
      .AddTableBody
      .AddTableBody Sis_Rpt_Number2
      .AddTableBody Sis_Rpt_Number2
      .AddTableBody Sis_Rpt_Number2
            
      .AddTableFooter "Total", , tdbHalignRight, , , , , , , , , , , , 3
      .AddTableFooter
      .AddTableFooter
      .AddTableFooter "&Sum", Sis_Rpt_Number2
      .AddTableFooter "&Sum", Sis_Rpt_Number2
      .AddTableFooter "&Sum", Sis_Rpt_Number2
      
      .Preview vaArray, , False
    End With
    
  End If
End Sub

Private Sub cmdPreview_Click()
  GetDataAll
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub Form_Load()
Dim n As Single

    SetIcon Me.hWnd, "SIKD"
    CenterForm Me
    InitValue
    TabIndex dTgl(0), n
    TabIndex dTgl(1), n
    TabIndex cmdPreview, n
    TabIndex cmdKeluar, n
End Sub

Sub InitValue()
  dTgl(0).Value = BOM(Date)
  dTgl(1).Value = Date
End Sub

