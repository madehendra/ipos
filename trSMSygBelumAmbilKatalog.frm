VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Begin VB.Form trSMSygBelumAmbilKatalog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SMS MEMBER YG BELUM AMBIL KATALOG"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   4755
   Begin BiSAButtonProject.BiSAButton BiSAButton1 
      Height          =   465
      Left            =   1530
      TabIndex        =   3
      Top             =   1200
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   820
      Caption         =   "OK"
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483633
   End
   Begin VB.OptionButton optBeli 
      Caption         =   "Belum beli"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   2670
      TabIndex        =   2
      Top             =   465
      Width           =   1320
   End
   Begin VB.OptionButton optBeli 
      Caption         =   "Sudah beli"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   1320
      TabIndex        =   1
      Top             =   465
      Width           =   1215
   End
   Begin BiSATextBoxProject.BiSABrowse cBarcode 
      Height          =   330
      Left            =   435
      TabIndex        =   0
      Top             =   780
      Width           =   3150
      _ExtentX        =   5556
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
      GetPicture      =   1
      Button          =   -1  'True
      Caption         =   "Barcode"
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
Attribute VB_Name = "trSMSygBelumAmbilKatalog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB


Private Sub BiSAButton1_Click()
Dim cSQL As String
Dim n As Integer
Dim a As New exportExcel
  
  vaArray.ReDim 0, -1, 0, 2
  cSQL = "SELECT * FROM anggota WHERE telp <> '' AND telp IS NOT NULL"
  Set dbData = objData.Sql(GetDSN, cSQL)
  If Not dbData.EOF Then
    FrmPB.InitPB dbData.RecordCount
    Do While Not dbData.EOF
       FrmPB.RunPB
       vaArray.InsertRows vaArray.UpperBound(1) + 1
       n = vaArray.UpperBound(1)
       vaArray(n, 0) = "'" & GetNull(dbData!telp)
       vaArray(n, 1) = GetNull(dbData!nama)
       vaArray(n, 2) = GetLastAktif(objData, GetNull(dbData!kodeanggota))
       If optBeli(0).Value = True Then
        'jika beli
        If isBeliBarang(GetNull(dbData!kodeanggota), cBarcode.Text) = False Then
          vaArray.DeleteRows n
        End If
       ElseIf optBeli(1).Value = True Then
        If isBeliBarang(GetNull(dbData!kodeanggota), cBarcode.Text) = True Then
          vaArray.DeleteRows n
        End If
       End If
      dbData.MoveNext
    Loop
    FrmPB.EndPB
  End If
  
  vaArray.QuickSort 0, vaArray.UpperBound(1), 2, XORDER_DESCEND, XTYPE_DATE, 1, XORDER_DESCEND, XTYPE_STRING
  a.RecordSource = vaArray
  a.ExportToExcel
End Sub

Private Function GetLastAktif(ByVal obj As CodeSuiteLibrary.Data, ByVal cKodeAnggota As String) As Date
Dim db As New ADODB.Recordset

  
  Set db = obj.Browse(GetDSN, "totpenjualan", , "kodeanggota", sisAssign, cKodeAnggota, , "tgl desc", , 0, 1)
  If Not db.EOF Then
    GetLastAktif = GetNull(db!Tgl)
  End If
End Function

Private Function isBeliBarang(ByVal cKodeAnggota As String, ByVal cBarcode As String) As Boolean
Dim cSQL As String
Dim db As New ADODB.Recordset


  isBeliBarang = False
  cSQL = cSQL & " select * from penjualan p"
  cSQL = cSQL & " LEFT JOIN totpenjualan t on t.nomorpenjualan = p.nomorpenjualan"
  cSQL = cSQL & " left join stock s on s.kodestock = p.kodestock"
  cSQL = cSQL & " Where t.kodeanggota = '" & cKodeAnggota & "' And s.Barcode = '" & cBarcode & "'"
    Set db = objData.Sql(GetDSN, cSQL)
    If Not db.EOF Then
      isBeliBarang = True
    End If
End Function


Private Sub cBarcode_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "stock", "barcode,kodestock,nama,kodesatuan,hargajual", "barcode", sisContent, cBarcode.Text, " AND jenis < 9", "kodestock")
  If Not dbData.EOF Then
    cBarcode.Text = cBarcode.Browse(dbData)
  End If
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub Form_Load()
Dim n As Single
  
  SetIcon Me.hWnd, "SIKD"
  CenterForm Me
  cBarcode.Default
  optBeli(1).Value = True
  
  TabIndex cBarcode, n
  
End Sub

