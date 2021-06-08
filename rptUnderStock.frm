VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form rptUnderStock 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Laporan Under Stock/ Stock Minimum"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   6540
   Begin BiSANumberBoxProject.BiSANumberBox nMin 
      Height          =   330
      Left            =   360
      TabIndex        =   9
      Top             =   1665
      Width           =   2160
      _ExtentX        =   3810
      _ExtentY        =   582
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
      Caption         =   "Minimum Stock"
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
   Begin VB.CheckBox Check1 
      Caption         =   "Tampilkan Seluruh Stock"
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
      Left            =   1500
      TabIndex        =   2
      Top             =   135
      Width           =   2190
   End
   Begin VB.OptionButton optKodeStock 
      Caption         =   "Kode Index"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   1605
      TabIndex        =   1
      Top             =   1230
      Width           =   1365
   End
   Begin VB.OptionButton optKodeStock 
      Caption         =   "Barcode"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   3030
      TabIndex        =   0
      Top             =   1245
      Width           =   1395
   End
   Begin BiSADateProject.BiSADate dTgl 
      Height          =   330
      Left            =   360
      TabIndex        =   3
      Top             =   795
      Width           =   2505
      _ExtentX        =   4419
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
      Caption         =   "sd Tgl"
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
   Begin BiSATextBoxProject.BiSABrowse cGolongan 
      Height          =   330
      Index           =   0
      Left            =   360
      TabIndex        =   4
      Top             =   420
      Width           =   3105
      _ExtentX        =   5477
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
      Button          =   -1  'True
      Caption         =   "Golongan"
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
   Begin BiSATextBoxProject.BiSABrowse cGolongan 
      Height          =   330
      Index           =   1
      Left            =   3465
      TabIndex        =   5
      Top             =   420
      Width           =   2685
      _ExtentX        =   4736
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
      Button          =   -1  'True
      Caption         =   "s.d"
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
   Begin BiSAButtonProject.BiSAButton cmdKeluar 
      Cancel          =   -1  'True
      Height          =   435
      Left            =   5085
      TabIndex        =   6
      Top             =   1815
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
      Picture         =   "rptUnderStock.frx":0000
   End
   Begin BiSAButtonProject.BiSAButton cmdPreview 
      Height          =   435
      Left            =   4650
      TabIndex        =   7
      Top             =   1815
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
      Picture         =   "rptUnderStock.frx":00A6
   End
   Begin VB.Label Label2 
      Caption         =   "Tampilkan"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   390
      TabIndex        =   8
      Top             =   1260
      Width           =   945
   End
End
Attribute VB_Name = "rptUnderStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB

Private Sub cGolongan_ButtonClick(Index As Integer)
  Set dbData = objData.Browse(GetDSN, "golongan", "kodegolongan,keterangan", "kodegolongan", sisContent, cGolongan(Index).Text, , "kodegolongan")
  If Not dbData.EOF Then
    cGolongan(Index).Text = cGolongan(Index).Browse(dbData)
  End If
End Sub

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
   Set vaArray = GetNilaiPersediaan2(cGolongan(0).Text, cGolongan(1).Text, True, dTgl.Value, IIf(Check1.Value = 1, True, False), IIf(optKodeStock(0).Value = True, True, False), nMin.Value)
   With FrmRPT

    .AddPageHeader "Nilai Persediaan Stock", tdbHalignCenter, , , True, , 12, True, True, , False, tdbPageHeaderSect
    If Check1.Value = 1 Then
      .AddPageHeader "All Category", , , 10, True, , , True
    Else
      .AddPageHeader "Golongan ", , , 10, True, , , True
      .AddPageHeader " : [ " & cGolongan(0).Text & " ] s/d [ " & cGolongan(1).Text & " ]", tdbHalignLeft, , , , , , True, , , False, , , , , , 15
    End If
  
    .AddTableGroupHeader True, "[]", , , , 10, , , , , , , , , , , , , , , True
    .AddTableGroupHeader
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    
    .AddTableHeader , , , , , , , , , , , , , , , , , , , False
    .AddTableHeader , , , , , , , , , , , , , , , , , , , False
    .AddTableHeader "Kode", , , , 11
    .AddTableHeader "Nama", , , , 0
    .AddTableHeader "Satuan", , , , 6
    .AddTableHeader "Stock", , , , 10
    .AddTableHeader "Hrg Beli", , , , 14
    .AddTableHeader "Hrg Jual", , , , 14
    
    .AddTableBody , , , , , , , , , , , , , False
    .AddTableBody , , , , , , , , , , , , , False
    .AddTableBody
    .AddTableBody
    .AddTableBody
    .AddTableBody Sis_Rpt_Number2, , , 8
    .AddTableBody Sis_Rpt_Number2, , , 15
    .AddTableBody Sis_Rpt_Number2, , , 15
    
    .AddTableGroupFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupFooter "SubTotal", , tdbHalignRight, , , , , , , , , , , , 5
    .AddTableGroupFooter
    .AddTableGroupFooter
    .AddTableGroupFooter
    .AddTableGroupFooter
    '.AddTableGroupFooter "&Sum", Sis_Rpt_Number2
    .AddTableGroupFooter
    
    .AddTableFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableFooter "Grand Total", , tdbHalignRight, , , , , , , , , , , , 5
    .AddTableFooter
    .AddTableFooter
    .AddTableFooter
    .AddTableFooter
    '.AddTableFooter "&Sum", Sis_Rpt_Number2
    .AddTableFooter
    
    .Refresh
    .Preview vaArray, True
  End With
End Sub

Private Sub Form_Load()
Dim n As Single

  SetIcon Me.hwnd, "SIKD"
  CenterForm Me
  Check1.Value = 1
  dTgl.Value = Date
  optKodeStock(1).Value = True
  nMin.Value = 0
  
  TabIndex Check1, n
  TabIndex cGolongan(0), n
  TabIndex cGolongan(1), n
  TabIndex dTgl, n
  TabIndex optKodeStock(0), n
  TabIndex optKodeStock(1), n
  TabIndex nMin, n
  
  TabIndex cmdPreview, n
  TabIndex cmdKeluar, n
End Sub

Private Sub optKodeStock_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    keybd_event VK_TAB, 0, 0, 0
  End If
End Sub

