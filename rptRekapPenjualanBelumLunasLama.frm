VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Begin VB.Form rptRekapPenjualanBelumLunasLama 
   Caption         =   "Form3"
   ClientHeight    =   2205
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7155
   LinkTopic       =   "Form3"
   ScaleHeight     =   2205
   ScaleWidth      =   7155
   StartUpPosition =   3  'Windows Default
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1110
      Left            =   0
      Top             =   0
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   1958
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
      Begin VB.OptionButton Option1 
         Caption         =   "Laporan 1"
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
         Index           =   0
         Left            =   555
         TabIndex        =   1
         Top             =   435
         Width           =   1350
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Laporan 2"
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
         Index           =   1
         Left            =   1830
         TabIndex        =   0
         Top             =   435
         Width           =   1350
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   630
      Left            =   15
      Top             =   1095
      Width           =   6960
      _ExtentX        =   12277
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
         Left            =   5805
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
         Picture         =   "rptRekapPenjualanBelumLunasLama.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   5355
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
         Picture         =   "rptRekapPenjualanBelumLunasLama.frx":00A6
      End
   End
End
Attribute VB_Name = "rptRekapPenjualanBelumLunasLama"
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
  If Option1(0).Value = True Then
    GetData
  ElseIf Option1(1).Value = True Then
    GetData2
  End If
End Sub

Private Sub Form_Load()
Dim n As Single

  Option1(1).Value = True
  SetIcon Me.hWnd, "SIKD"
  CenterForm Me
  TabIndex cmdPreview, n
  TabIndex cmdKeluar, n
End Sub

Private Sub GetData()
Dim n As Integer
Dim cSQL As String

  cSQL = ""
  cSQL = cSQL & " select t.kodeanggota,a.nama,t.username,t.dp,t.tgl,t.nomorpenjualan,t.total from totpenjualan t"
  cSQL = cSQL & " left join anggota a on a.kodeanggota = t.kodeanggota"
  cSQL = cSQL & " Where t.flaglunas = 0 Or t.flaglunas Is Null"
  cSQL = cSQL & " order by t.kodeanggota,t.tgl"
  
  vaArray.ReDim 0, -1, 0, 5
  Set dbData = objData.SQL(GetDSN, cSQL)
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = GetNull(dbData!kodeanggota)
      vaArray(n, 1) = GetNull(dbData!nama)
      vaArray(n, 2) = GetNull(dbData!UserName)
      vaArray(n, 3) = Format(GetNull(dbData!tgl), "dd-MM-yyyy")
      vaArray(n, 4) = GetNull(dbData!nomorpenjualan)
      vaArray(n, 5) = GetNull(dbData!Total) - GetNull(dbData!dp)
      dbData.MoveNext
    Loop
    GetRpt
  End If
End Sub

Private Sub GetData2()
Dim n As Single
Dim cSQL As String
Dim a As New exportExcel
Dim na As Integer


  cSQL = ""
  cSQL = cSQL & " select d.kodedep,d.keterangan,a.nama,t.kodeanggota,a.telp,sum(t.total)-sum(t.dp) as totalnya from totpenjualan t"
  cSQL = cSQL & " left join anggota a on a.kodeanggota = t.kodeanggota"
  cSQL = cSQL & " left join dep d on d.kodedep = a.kodedep"
  cSQL = cSQL & " Where t.flaglunas = 0 Or t.flaglunas Is Null"
  cSQL = cSQL & " group by t.kodeanggota"
  cSQL = cSQL & " order by d.kodedep,d.keterangan,a.nama"
  
  vaArray.ReDim 0, -1, 0, 5
  Set dbData = objData.SQL(GetDSN, cSQL)
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = GetNull(dbData!kodedep)
      vaArray(n, 1) = GetNull(dbData!keterangan)
      vaArray(n, 2) = GetNull(dbData!nama)
      vaArray(n, 3) = GetNull(dbData!kodeanggota)
      vaArray(n, 4) = GetNull(dbData!telp)
      vaArray(n, 5) = GetNull(dbData!totalnya)
      dbData.MoveNext
    Loop
    GetRpt2

    a.RecordSource = vaArray
    a.ExportToExcel
    
  End If
End Sub

Private Sub GetRpt()
  With FrmRPT
    
    .AddPageHeader "Penjualan Belum Lunas", tdbHalignCenter, , , True, , 10, True, False, True, False, tdbPageHeaderSect
    .AddPageHeader aCfg(objData, msNamaPerusahaan), tdbHalignCenter, , , True, , 14, True, False, True, False, tdbPageHeaderSect
        
    .AddTableGroupHeader True, "[]", , , , 15
    .AddTableGroupHeader
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    
    
    .AddTableHeader "", , , , , , , , , , , , , , , , , , , False
    .AddTableHeader "", , , , , , , , , , , , , , , , , , , False
    .AddTableHeader "Kasir", , , , 15
    .AddTableHeader "Tgl", , , , 12
    .AddTableHeader "Nomor", , , , 15
    .AddTableHeader "Total", , , , 15
    
     
    .AddTableBody , , , , , , , , , , , , , False
    .AddTableBody , , , , , , , , , , , , , False
    .AddTableBody
    .AddTableBody
    .AddTableBody
    .AddTableBody Sis_Rpt_Number2
    
    .AddTableGroupFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupFooter "SubTotal", , tdbHalignRight, , , , , , , , , , , , 3
    .AddTableGroupFooter
    .AddTableGroupFooter
    .AddTableGroupFooter "&Sum", Sis_Rpt_Number2
    
    .Refresh
    .Preview vaArray, True
  End With
End Sub

Private Sub GetRpt2()

  With FrmRPT
    .AddPageHeader "Penjualan Belum Lunas", tdbHalignCenter, , , True, , 10, True, False, True, False, tdbPageHeaderSect
    .AddPageHeader aCfg(objData, msNamaPerusahaan), tdbHalignCenter, , , True, , 14, True, False, True, False, tdbPageHeaderSect
            
    .AddTableGroupHeader True, "[]", , , , 15
    .AddTableGroupHeader
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False

    
    .AddTableHeader "Dep", , , , 15, , , , , , , , , , , , , , , False
    .AddTableHeader "Ket", , , , 15, , , , , , , , , , , , , , , False
    .AddTableHeader "Member", , , , 25
    .AddTableHeader "Kode", , , , 12
    .AddTableHeader "Telp", , , , 15
    .AddTableHeader "Total", , , , 15
    
    .AddTableBody , , , , , , , , , , , , , False
    .AddTableBody , , , , , , , , , , , , , False
    .AddTableBody
    .AddTableBody
    .AddTableBody
    .AddTableBody Sis_Rpt_Number2
        
    .AddTableFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableFooter "Total", , tdbHalignRight, , , , , , , , , , , , 3
    .AddTableFooter
    .AddTableFooter
    .AddTableFooter "&Sum", Sis_Rpt_Number2
   
    .Refresh
    .Preview vaArray, True
  End With

End Sub



