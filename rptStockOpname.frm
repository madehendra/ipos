VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form rptStockOpname 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "STOCK OPNAME"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5865
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   5865
   Begin SizerOneLibCtl.ElasticOne ElasticOne2 
      Height          =   1830
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   5865
      _cx             =   10345
      _cy             =   3228
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
      Appearance      =   4
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   700
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   5
      AutoSizeChildren=   0
      BorderWidth     =   6
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   1
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   2
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   0
      GridCols        =   0
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   ""
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   195
         Left            =   1575
         TabIndex        =   5
         Top             =   885
         Width           =   210
      End
      Begin BiSADateProject.BiSADate dTgl 
         Height          =   330
         Index           =   0
         Left            =   450
         TabIndex        =   2
         Top             =   330
         Width           =   2550
         _ExtentX        =   4498
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
         Caption         =   "Tgl"
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
         Left            =   3015
         TabIndex        =   3
         Top             =   330
         Width           =   1890
         _ExtentX        =   3334
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
         Caption         =   "sd"
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
      Begin BiSATextBoxProject.BiSABrowse cNomor 
         Height          =   360
         Left            =   1845
         TabIndex        =   4
         Top             =   870
         Width           =   1995
         _ExtentX        =   3519
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
      Begin VB.Label Label1 
         Caption         =   "Nomor"
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
         Left            =   510
         TabIndex        =   6
         Top             =   810
         Width           =   750
      End
   End
   Begin SizerOneLibCtl.ElasticOne ElasticOne1 
      Height          =   630
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1830
      Width           =   5865
      _cx             =   10345
      _cy             =   1111
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
      Appearance      =   4
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   700
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   2
      AutoSizeChildren=   0
      BorderWidth     =   6
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   1
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   2
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   0
      GridCols        =   0
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   ""
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   4665
         TabIndex        =   7
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
         Picture         =   "rptStockOpname.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   4230
         TabIndex        =   8
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
         Picture         =   "rptStockOpname.frx":00A6
      End
   End
End
Attribute VB_Name = "rptStockOpname"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB

Private Sub Check1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    keybd_event VK_TAB, 0, 0, 0
  End If
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdPreview_Click()
  GetRpt
End Sub

Private Sub GetRpt()
Dim cWhere As String
Dim n As Integer
  
   vaArray.ReDim 0, -1, 0, 7
   cWhere = ""
   If Check1.Value = 1 Then
    cWhere = " 1=1 AND t.nomorstockopname = '" & cNomor.Text & "'"
   End If
   
   Set dbData = objData.Browse(GetDSN, "stockopname p", "p.nomorstockopname,  t.tgl,t.keterangan as reason,p.kodestock,s.nama,g.keterangan as namagolongan,p.adjust,s.kodesatuan", , , , cWhere, "t.nomorstockopname", Array("LEFT JOIN totstockopname t on t.nomorstockopname = p.nomorstockopname", "LEFT JOIN stock s on s.kodestock = p.kodestock", "LEFT JOIN golongan g on g.kodegolongan = s.kodegolongan"))
   If Not dbData.EOF Then
     Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = "Nomor - " & GetNull(dbData!nomorstockopname)
      vaArray(n, 1) = "Tgl   : " & Format(GetNull(dbData!tgl), "dd-MM-yyyy")
      vaArray(n, 2) = "Ket.  : " & GetNull(dbData!reason)
      vaArray(n, 3) = GetNull(dbData!KodeStock)
      vaArray(n, 4) = GetNull(dbData!nama)
      vaArray(n, 5) = GetNull(dbData!namagolongan)
      vaArray(n, 6) = GetNull(dbData!Adjust)
      vaArray(n, 7) = GetNull(dbData!kodesatuan)
      dbData.MoveNext
     Loop
     
     With FrmRPT
      .AddPageHeader "STOCK OPNAME", tdbHalignCenter, , , True, , 12, True, True, , False, tdbPageHeaderSect
      .AddPageHeader "", , , , True
      .AddPageHeader "", , , , True
           
      .AddTableGroupHeader True, , , , , , , , , , , , , , , , , , , , True
      .AddTableGroupHeader , , , , , , True, , , , , , , , , , , , , , True
      .AddTableGroupHeader , , , , , , True, , , , , , , , , , , , , , True
      .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
      .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
      .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
      .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
      .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
      
      .AddTableHeader , , , , , , , , , , , , , , , , , , , False
      .AddTableHeader , , , , , , , , , , , , , , , , , , , False
      .AddTableHeader , , , , , , , , , , , , , , , , , , , False
      .AddTableHeader "Kode", , , , 12
      .AddTableHeader "Nama"
      .AddTableHeader "Golongan", , , , 20
      .AddTableHeader "Adjust", , , , 11
      .AddTableHeader "Satuan", , , , 8
      
      .AddTableBody , , , , , , , , , , , , , False
      .AddTableBody , , , , , , , , , , , , , False
      .AddTableBody , , , , , , , , , , , , , False
      .AddTableBody
      .AddTableBody
      .AddTableBody
      .AddTableBody Sis_Rpt_Number2
      .AddTableBody
      
'      .AddTableGroupFooter , , , , , , , , , , , , , , , , , , , False
'      .AddTableGroupFooter , , , , , , , , , , , , , , , , , , , False
'      .AddTableGroupFooter "SubTotal", , tdbHalignRight, , , , , , , , , , , , 5
'      .AddTableGroupFooter
'      .AddTableGroupFooter
'      .AddTableGroupFooter
'      .AddTableGroupFooter
'      .AddTableGroupFooter "&Sum", Sis_Rpt_Number2
'
'
'      .AddTableFooter , , , , , , , , , , , , , , , , , , , False
'      .AddTableFooter , , , , , , , , , , , , , , , , , , , False
'      .AddTableFooter "Grand Total", , tdbHalignRight, , , , , , , , , , , , 5
'      .AddTableFooter
'      .AddTableFooter
'      .AddTableFooter
'      .AddTableFooter
'      .AddTableFooter "&Sum", Sis_Rpt_Number2
      
      .Refresh
      .Preview vaArray, True
    End With
  End If
End Sub

Private Sub cNomor_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "totstockopname", "nomorstockopname,tgl,keterangan", "tgl", sisGTEqual, Format(dTgl(0).Value, "yyyy-MM-dd"), " AND tgl <= '" & Format(dTgl(1).Value, "yyyy-MM-dd") & "'", "tgl")
  If Not dbData.EOF Then
    cNomor.Text = cNomor.Browse(dbData)
  End If
End Sub

Private Sub Form_Load()
Dim n As Single

  SetIcon Me.hWnd, "SIKD"
  CenterForm Me
  dTgl(0).Value = Date
  dTgl(1).Value = Date
  TabIndex dTgl(0), n
  TabIndex dTgl(1), n
  TabIndex Check1, n
  TabIndex cNomor, n
  TabIndex cmdPreview, n
  TabIndex cmdKeluar, n
End Sub

