VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form rptSensusHarian 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SENSUS HARIAN SALES"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8820
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   8820
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   2100
      Left            =   15
      Top             =   0
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   3704
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
      Begin BiSADateProject.BiSADate dDate 
         Height          =   330
         Index           =   0
         Left            =   1335
         TabIndex        =   0
         Top             =   555
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   582
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
         Caption         =   "ANTARA TANGGAL"
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
      Begin BiSADateProject.BiSADate dDate 
         Height          =   330
         Index           =   1
         Left            =   4890
         TabIndex        =   1
         Top             =   570
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   582
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
         Caption         =   "S.D"
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
      Begin BiSATextBoxProject.BiSABrowse cGroupSales 
         Height          =   330
         Left            =   1335
         TabIndex        =   2
         Top             =   930
         Width           =   3450
         _ExtentX        =   6085
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
         Caption         =   "Group Sales"
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
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   630
      Left            =   15
      Top             =   2100
      Width           =   8790
      _ExtentX        =   15505
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
         Left            =   7620
         TabIndex        =   3
         Top             =   105
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
         Picture         =   "rptSensusHarian.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   7170
         TabIndex        =   4
         Top             =   105
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
         Picture         =   "rptSensusHarian.frx":00A6
      End
   End
End
Attribute VB_Name = "rptSensusHarian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB

Private Sub cGroupSales_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "groupsales", "kode,keterangan")
  If Not dbData.EOF Then
    cGroupSales.Text = cGroupSales.Browse(dbData)
    cGroupSales.Text = GetNull(dbData!Kode)
  End If
End Sub

Private Sub GetLoadRows()
Dim n As Integer
Dim cSQL As String

  cSQL = cSQL & " select  TGL,sum(piutang) as piutang ,sum(tunai) as tunai from totpenjualan"
  cSQL = cSQL & " where kodegroupsales = '" & cGroupSales.Text & "' and  tgl >= '" & Format(dDate(0).Value, "yyyy-MM-dd") & "' and tgl <= '" & Format(dDate(1).Value, "yyyy-MM-dd") & "'"
  cSQL = cSQL & " GROUP BY tgl"
  
  vaArray.ReDim 0, -1, 0, 4
  Set dbData = objData.SQL(GetDSN, cSQL)
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = Format(GetNull(dbData!tgl), "dd-MM-yyyy")
      vaArray(n, 1) = GetNull(dbData!Piutang)
      vaArray(n, 2) = GetNull(dbData!Tunai)
      vaArray(n, 3) = 0
      vaArray(n, 4) = 0
      dbData.MoveNext
    Loop
    GetPreview
  Else
    MsgBox "Data tidak ada"
  End If
End Sub

Private Sub GetPreview()
   With FrmRPT
    .AddPageHeader UCase("Sensus Harian - " & cGroupSales.Text), tdbHalignCenter, , , True, dbArial, 12, True, , , False
    .AddPageHeader aCfg(objData, msNamaPerusahaan), tdbHalignCenter, , , True, , 14
    .AddPageHeader "Dari Tanggal " & Format(dDate(0).Value, "dd-MM-yyyy") & "-" & Format(dDate(1).Value, "dd-MM-yyyy"), tdbHalignCenter, , , True, , 13
    .AddPageHeader "", , , , True
    .AddPageHeader "", , , , True


    .AddTableHeader "TANGGAL", , , , , , , , , , , , , tdbMergeOnText
    .AddTableHeader "PEMBAYARAN", , , , 15, , , , , , , , , , 3
    .AddTableHeader , , , , 15
    .AddTableHeader , , , , 15
    .AddTableHeader "SALDO", , , , , , , , , , , , , tdbMergeOnText
    
    .AddTableHeader "TANGGAL", , , , 10, , , , , , True, , , tdbMergeOnText
    .AddTableHeader "BON"
    .AddTableHeader "TUNAI"
    .AddTableHeader "PIUTANG"
    .AddTableHeader "SALDO", , , , , , , , , , , , , tdbMergeOnText
    
    .AddTableBody
    .AddTableBody Sis_Rpt_Number2, tdbHalignRight
    .AddTableBody Sis_Rpt_Number2, tdbHalignRight
    .AddTableBody Sis_Rpt_Number2, tdbHalignRight
    .AddTableBody Sis_Rpt_Number2, tdbHalignRight
    
    .Preview vaArray, True
   End With
End Sub

Private Sub cmdPreview_Click()
   GetLoadRows
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub Form_Load()
Dim n As Single

  SetIcon Me.hWnd, "SIKD"
  CenterForm Me
  dDate(0).Value = BOM(Date)
  cGroupSales.Default
  TabIndex dDate(0), n
  TabIndex dDate(1), n
  TabIndex cGroupSales, n
  TabIndex cmdPreview, n
  TabIndex cmdKeluar, n
End Sub


