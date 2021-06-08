VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form trPostingKasir 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Posting Kasir"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   5295
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1140
      Left            =   15
      Top             =   0
      Width           =   5250
      _ExtentX        =   9260
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
         Width           =   2235
         _ExtentX        =   3942
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
      Begin BiSADateProject.BiSADate dTgl 
         Height          =   330
         Index           =   1
         Left            =   2640
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
      Left            =   15
      Top             =   1125
      Width           =   5250
      _ExtentX        =   9260
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
         Left            =   4110
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
         Picture         =   "trPostingKasir.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   3675
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
         Picture         =   "trPostingKasir.frx":00A6
      End
   End
End
Attribute VB_Name = "trPostingKasir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB

Private Sub cmdPreview_Click()
  GetProses
End Sub

Private Sub Form_Load()
Dim n As Single

  CenterForm Me
  SetIcon Me.hWnd

  dTgl(0).Value = Date
  dTgl(1).Value = Date
  
  TabIndex dTgl(0), n
  TabIndex dTgl(1), n
  TabIndex cmdPreview, n
  TabIndex cmdKeluar, n
  
End Sub

Private Sub GetProses()
Dim lSave As Boolean

lSave = True

  objData.Start GetDSN
  lSave = IIf(lSave, objData.Delete(GetDSN, "bukubesar", "status", sisAssign, msPenjualanKasir, " and tgl >= '" & Format(dTgl(0).Value, "yyyy-MM-dd") & "'  and tgl <= '" & Format(dTgl(0).Value, "yyyy-MM-dd") & "'"), False)
  Set dbData = objData.Browse(GetDSN, "kasir k", "k.nomorkasir,t.tgl,k.kodestock", "t.tgl", sisGTEqual, Format(dTgl(0).Value, "yyyy-MM-dd"), " and t.tgl <= '" & Format(dTgl(1).Value, "yyyy-MM-dd") & "'", , Array("left join totkasir t on t.nomorkasir = k.nomorkasir"))
  If Not dbData.EOF Then
    FrmPB.InitPB dbData.RecordCount
    Do While Not dbData.EOF
      FrmPB.RunPB
      'HP (5)
      'persediaan (1)
        
        'Kas
          'Penjualan
        lSave = IIf(lSave, UpdKodeTr(objData, msPenjualanKasir, GetNull(dbData!nomorkasir), Format(GetNull(dbData!tgl), "yyyy-MM-dd"), cAkunKas.Text, aCfg(objData, msCostCenterJualBeli), "Penjualan Kasir no " & GetNull(dbData!nomorkasir), nTotal.Value, 0, "K", SNow), False)
            lSave = IIf(lSave, UpdKodeTr(objData, msPenjualanKasir, GetNull(dbData!nomorkasir), Format(GetNull(dbData!tgl), "yyyy-MM-dd"), aCfg(objData, msRekeningPenjualan), aCfg(objData, msCostCenterJualBeli), "Penjualan Kasir No " & GetNull(dbData!nomorkasir), 0, nTotal.Value, "N"), False)
        
        
        lSave = IIf(lSave, UpdKodeTr(objData, msPenjualanKasir, GetNull(dbData!nomorkasir), Format(GetNull(dbData!tgl), "yyyy-MM-dd"), aCfg(objData, msRekeningCOGS), aCfg(objData, msCostCenterJualBeli), "Harga Pokok Penjualan Kasir No " & GetNull(dbData!nomorkasir), GetHargaBeli(objData, GetNull(dbData!KodeStock)), 0, "N", SNow), False)
            lSave = IIf(lSave, UpdKodeTr(objData, msPenjualanKasir, GetNull(dbData!nomorkasir), Format(GetNull(dbData!tgl), "yyyy-MM-dd"), GetAkunInventory(objData, GetNull(dbData!KodeStock)), aCfg(objData, msCostCenterJualBeli), "Harga Pokok Penjualan Kasir No " & GetNull(dbData!nomorkasir), 0, GetHargaBeli(objData, GetNull(dbData!KodeStock)), "N"), False)

      dbData.MoveNext
    Loop
    FrmPB.EndPB
  End If
  If lSave Then
    objData.Save GetDSN
  Else
    objData.Cancel GetDSN
  End If

End Sub
