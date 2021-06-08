VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Begin VB.Form rptConsignmentStock 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consignment Stock"
   ClientHeight    =   885
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6000
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   885
   ScaleWidth      =   6000
   Begin BiSAButtonProject.BiSAButton cmdKeluar 
      Cancel          =   -1  'True
      Height          =   435
      Left            =   4845
      TabIndex        =   0
      Top             =   195
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
      Picture         =   "rptConsignmentStock.frx":0000
   End
   Begin BiSAButtonProject.BiSAButton cmdPreview 
      Height          =   435
      Left            =   4410
      TabIndex        =   1
      Top             =   195
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
      Picture         =   "rptConsignmentStock.frx":00A6
   End
End
Attribute VB_Name = "rptConsignmentStock"
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
  GetData
End Sub

Private Sub Form_Load()
Dim n As Single

  SetIcon Me.hWnd, "SIKD"
  CenterForm Me
  TabIndex cmdPreview, n
  TabIndex cmdKeluar, n
End Sub

Private Sub GetData()
Dim n As Integer
  
  vaArray.ReDim 0, -1, 0, 4
  Set dbData = objData.Browse(GetDSN, "stock s", "s.kodesupplier,su.nama as namasupplier,s.kodestock,s.nama,s.hargabeli", "s.jenis", sisAssign, "9", , "s.kodesupplier", Array("LEFT JOIN supplier su ON su.kodesupplier = s.kodesupplier"))
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = GetNull(dbData!kodesupplier)
      vaArray(n, 1) = GetNull(dbData!namasupplier)
      vaArray(n, 2) = GetNull(dbData!KodeStock)
      vaArray(n, 3) = GetNull(dbData!nama)
      vaArray(n, 4) = GetNull(dbData!hargabeli)
      dbData.MoveNext
    Loop
    GetRpt
  End If
End Sub

Private Sub GetRpt()
  With FrmRPT
    
    .AddPageHeader "Stock Konsinyasi", tdbHalignCenter, , , True, , 10, True, False, True, False, tdbPageHeaderSect
    .AddPageHeader aCfg(objData, msNamaPerusahaan), tdbHalignCenter, , , True, , 14, True, False, True, False, tdbPageHeaderSect
        
    .AddTableGroupHeader True, "[]", , , , 10
    .AddTableGroupHeader
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    
    .AddTableHeader "", , , , , , , , , , , , , , , , , , , False
    .AddTableHeader "", , , , , , , , , , , , , , , , , , , False
    .AddTableHeader "Kode Barang", , , , 12
    .AddTableHeader "Nama Barang", , , , 52
    .AddTableHeader "Harga Beli", , , , 15
    
     
    .AddTableBody , , , , , , , , , , , , , False
    .AddTableBody , , , , , , , , , , , , , False
    .AddTableBody
    .AddTableBody
    .AddTableBody Sis_Rpt_Number2
    
    .Refresh
    .Preview vaArray, True
  End With
End Sub

