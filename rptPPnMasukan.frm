VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form rptPPnMasukan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PPn Masukan"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   7605
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1035
      Left            =   15
      Top             =   0
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   1826
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
         Left            =   600
         TabIndex        =   0
         Top             =   285
         Width           =   3480
         _ExtentX        =   6138
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
         Caption         =   "TANGGAL PEMBELIAN "
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
         Left            =   4155
         TabIndex        =   1
         Top             =   285
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
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   645
      Left            =   15
      Top             =   1005
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   1138
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
         Left            =   6435
         TabIndex        =   2
         Top             =   90
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
         Picture         =   "rptPPnMasukan.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   6000
         TabIndex        =   3
         Top             =   90
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
         Picture         =   "rptPPnMasukan.frx":00A6
      End
   End
End
Attribute VB_Name = "rptPPnMasukan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vaArray As New XArrayDB
Dim objData As New CodeSuiteLibrary.Data
Dim dbData As New ADODB.Recordset

Private Sub GetData()
Dim n As Double
Dim cField As String
Dim cWhere As String
Dim vaJoin

  vaArray.ReDim 0, -1, 0, 3
  
  cField = "s.alamat,s.telepon,s.nama,s.kota,sum(t.pajak) as pajak"
  cWhere = " AND tgl <= '" & Format(dTgl(1).Value, "yyyy-MM-dd") & "'"
  cWhere = cWhere & " AND t.pajak <> 0 GROUP BY s.kodesupplier"
  vaJoin = Array("LEFT JOIN supplier s on t.kodesupplier = s.kodesupplier")
  Set dbData = objData.Browse(GetDSN, "totpembelian t", cField, "t.Tgl", sisGTEqual, Format(dTgl(0).Value, "yyyy-MM-dd"), cWhere, "t.tgl,t.nomorpembelian", vaJoin)
  If Not dbData.EOF Then
    FrmPB.InitPB dbData.RecordCount
    dbData.MoveFirst
    Do While Not dbData.EOF
      FrmPB.RunPB
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      
      vaArray(n, 0) = GetNull((dbData!nama), "")
      vaArray(n, 1) = GetNull((dbData!alamat), "")
      vaArray(n, 2) = GetNull((dbData!TELEPON), "")
      vaArray(n, 3) = GetNull(dbData!PAJAK)
      dbData.MoveNext
    Loop
    FrmPB.EndPB
    
    If vaArray.UpperBound(1) >= 0 Then
      GetRpt
    Else
      MsgBox "Data tidak ada...", vbInformation
      Exit Sub
    End If
  Else
    MsgBox "Data tidak ada...", vbInformation
    Exit Sub
  End If
End Sub

Private Sub GetRpt()
  With FrmRPT
    .AddPageHeader "LAPORAN PPn MASUKAN", tdbHalignCenter, , , , dbArial, 12, True
    .AddPageHeader aCfg(objData, msNamaPerusahaan, ""), tdbHalignCenter, , , True, dbArial, 12, True
    .AddPageHeader aCfg(objData, msAlamatPerusahaan, "") & " " & aCfg(objData, msKota, ""), tdbHalignCenter, , , True, dbArial, 12, True
    .AddPageHeader "TELP : " & aCfg(objData, msTelepon, ""), tdbHalignCenter, , , True, dbArial, 12, True
    .AddPageHeader "Antara Tanggal : " & Format(dTgl(0).Value, "dd-MM-yyyy") & " S/D " & Format(dTgl(1).Value, "dd-MM-yyyy"), tdbHalignCenter, , , True, , 10
    .AddPageHeader " ", , , , True
    .AddPageHeader " ", , , , True
    
    .AddTableHeader "NAMA SUPPLIER", , , , , , , 10, , , , , , , , False, 8
    .AddTableHeader "ALAMAT", , , , , , , 10
    .AddTableHeader "TELEPON", , , , 13, , , 10
    .AddTableHeader "PPn (Rp)", , , , 17, , , 10
    
    'isi Laporan
    .AddTableBody , , , , , , 9
    .AddTableBody , , , , , , 9
    .AddTableBody Sis_Rpt_dd_MM_yyyy, tdbHalignCenter, , , , , 9
    .AddTableBody Sis_Rpt_Number2, tdbHalignRight, , , , , 9
    
    .AddTableFooter "TOTAL", , tdbHalignRight, , , , , 10, , , , , , , 3, False, 8
    .AddTableFooter
    .AddTableFooter
    .AddTableFooter "&Sum", Sis_Rpt_Number2, , , , , , 10
    
    .Preview vaArray, True
  End With
End Sub

Private Sub cmdPreview_Click()
  GetData
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

