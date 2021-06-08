VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Begin VB.Form rptDaftarStock 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DAFTAR INVENTORY"
   ClientHeight    =   1740
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   6675
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1125
      Left            =   0
      Top             =   15
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   1984
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
      Begin BiSATextBoxProject.BiSABrowse cGolongan 
         Height          =   330
         Index           =   0
         Left            =   180
         TabIndex        =   0
         Top             =   120
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
         Left            =   3285
         TabIndex        =   1
         Top             =   120
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
         Caption         =   "sd."
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
      Begin BiSATextBoxProject.BiSABrowse cKode 
         Height          =   330
         Index           =   0
         Left            =   180
         TabIndex        =   2
         Top             =   465
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
         Caption         =   "Kode"
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
      Begin BiSATextBoxProject.BiSABrowse cKode 
         Height          =   330
         Index           =   1
         Left            =   3285
         TabIndex        =   3
         Top             =   465
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
         Caption         =   "sd."
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
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   630
      Left            =   0
      Top             =   1125
      Width           =   6645
      _ExtentX        =   11721
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
         Left            =   5505
         TabIndex        =   4
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
         Picture         =   "rptDaftarStock.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   5070
         TabIndex        =   5
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
         Picture         =   "rptDaftarStock.frx":00A6
      End
   End
End
Attribute VB_Name = "rptDaftarStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB

Private Sub cGolongan_ButtonClick(Index As Integer)
  Set dbData = objData.Browse(GetDSN, "golongan", "kodegolongan,keterangan", "kodegolongan", sisContent, cGolongan(Index).Text, , "kodegolongan,keterangan")
  If Not dbData.EOF Then
    cGolongan(Index).Text = cGolongan(Index).Browse(dbData)
  End If
End Sub

Private Sub cKode_ButtonClick(Index As Integer)
  Set dbData = objData.Browse(GetDSN, "stock s", "s.kodestock,s.barcode,s.nama,s.hargajual,s.kodesatuan", "s.kodestock", sisContent, cKode(Index).Text, " AND s.kodegolongan = '" & cGolongan(Index).Text & "'", "s.kodestock")
  If Not dbData.EOF Then
    cKode(Index).Text = cKode(Index).Browse(dbData)
  End If
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdPreview_Click()
  GetData
End Sub

Private Sub GetData()
Dim n As Double
Dim nSaldo As Double
Dim cWhere As String
Dim cFields As String
  
  If cGolongan(0).Text > cGolongan(1).Text Or cKode(0).Value > cKode(1).Value Then
    MsgBox "Kode Golongan Tidak Valid, atau Kode Inventory Tidak Valid", vbExclamation
  Else
    cWhere = cWhere & "s.kodegolongan >= '" & cGolongan(0).Text & "' AND s.kodegolongan <= '" & cGolongan(1).Text & "' "
    cWhere = cWhere & " AND s.kodestock >= '" & cKode(0).Text & "' AND s.kodestock <= '" & cKode(1).Text & "'"
    cWhere = cWhere & " GROUP BY kodestock"
    cFields = cFields & "s.kodegolongan,g.keterangan as namagolongan,s.kodestock,s.barcode,s.nama,s.kodesatuan,s.hargajual,sum(k.debet-k.kredit) as saldostock"
    Set dbData = objData.Browse(GetDSN, "stock s", cFields, _
                                , , , cWhere, "s.nama,s.kodegolongan, s.kodestock", _
                                Array("Left Join golongan g on g.kodegolongan = s.kodegolongan", "LEFT JOIN kartustock k ON k.kodestock = s.kodestock"))
                                
    If dbData.RecordCount > 0 Then
      vaArray.LoadRows dbData.GetRows(dbData.RecordCount)
      GetRpt
    End If
  End If
End Sub

Private Sub GetRpt()
  With FrmRPT
    
    .AddPageHeader "Daftar Stock", tdbHalignCenter, , , True, , 14, True, True, True, False, tdbPageHeaderSect, , , , , 3
    
    .AddPageHeader "Golongan", , , 10, True, , , True
    .AddPageHeader " : " & cGolongan(0).Text & " s/d " & cGolongan(1).Text, tdbHalignLeft, , , , , , True, , , False
    .AddPageHeader "Kode ", , , 10, True, , , True
    .AddPageHeader " : " & cKode(0).Text & " s/d " & cKode(1).Text, tdbHalignLeft, , , , , , True, , , False
    
    .AddTableGroupHeader True, "[]", , , , 10
    .AddTableGroupHeader
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    
    .AddTableHeader "", , , , , , , , , , , , , , , , , , , False
    .AddTableHeader "", , , , , , , , , , , , , , , , , , , False
    .AddTableHeader "Kode ", , , , 12
    .AddTableHeader "Barcode", , , , 12
    .AddTableHeader "Nama"
    .AddTableHeader "Satuan", , , , 8
    .AddTableHeader "Sales", , , , 12
    .AddTableHeader "Stock", , , , 12
    
     
    .AddTableBody , , , , , , , , , , , , , False
    .AddTableBody , , , , , , , , , , , , , False
    .AddTableBody
    .AddTableBody
    .AddTableBody
    .AddTableBody
    .AddTableBody Sis_Rpt_Number2, tdbHalignRight
    .AddTableBody Sis_Rpt_Number2, tdbHalignRight
    
    .Refresh
    .Preview vaArray, True
  End With
End Sub

Private Sub Form_Load()
Dim n As Single

  SetIcon Me.hwnd, "SIKD"
  CenterForm Me
  TabIndex cGolongan(0), n
  TabIndex cGolongan(1), n
  TabIndex cKode(0), n
  TabIndex cKode(1), n
  TabIndex cmdPreview, n
End Sub
