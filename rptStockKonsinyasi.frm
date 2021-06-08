VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Begin VB.Form rptStockKonsinyasi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stok Konsinyasi"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6255
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   6255
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   630
      Left            =   105
      Top             =   1455
      Width           =   6060
      _ExtentX        =   10689
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
         Left            =   4740
         TabIndex        =   1
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
         Picture         =   "rptStockKonsinyasi.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   4305
         TabIndex        =   0
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
         Picture         =   "rptStockKonsinyasi.frx":00A6
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1410
      Left            =   120
      Top             =   30
      Width           =   6045
      _ExtentX        =   10663
      _ExtentY        =   2487
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
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   195
         Left            =   2325
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   570
         Width           =   240
      End
      Begin BiSATextBoxProject.BiSABrowse cSupplier 
         Height          =   330
         Left            =   2565
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   570
         Width           =   1725
         _ExtentX        =   3043
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
         Caption         =   "Supplier"
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
         Left            =   270
         TabIndex        =   2
         Top             =   555
         Width           =   1575
      End
   End
End
Attribute VB_Name = "rptStockKonsinyasi"
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
  GetRpt
End Sub

Private Sub cSupplier_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "supplier", "kodesupplier,nama,alamat")
  If Not dbData.EOF Then
    cSupplier.Text = cSupplier.Browse(dbData)
  End If
End Sub

Private Sub Form_Load()
Dim n As Single
  
  CenterForm Me
  SetIcon Me.hWnd

  TabIndex cmdPreview, n
  TabIndex cmdKeluar, n
End Sub

Private Sub GetRpt()
Dim n As Integer
Dim cSQL As String
Dim cWhere As String

  cWhere = ""
  If Check1.value = 1 Then
    cWhere = cWhere & " and s.kodesupplier = '" & cSupplier.Text & "'"
  End If
  
  cSQL = ""
  cSQL = "select s.kodesupplier,s.kodestock,s.nama,s.hargabeli,s.hargajual,s.kodesatuan,s.stok,sp.nama as supplier "
  cSQL = cSQL & " from stock s"
  cSQL = cSQL & " left join supplier sp on sp.kodesupplier = s.kodesupplier"
  cSQL = cSQL & " where s.konsi = '1'" & cWhere
  

  
  vaArray.ReDim 0, -1, 0, 7
  Set dbData = objData.SQL(GetDSN, cSQL)
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      
      vaArray(n, 0) = GetNull(dbData!kodesupplier)
      vaArray(n, 1) = GetNull(dbData!supplier)
      vaArray(n, 2) = GetNull(dbData!KodeStock)
      vaArray(n, 3) = GetNull(dbData!nama)
      vaArray(n, 4) = GetNull(dbData!hargabeli)
      vaArray(n, 5) = GetNull(dbData!HargaJual)
      vaArray(n, 6) = GetNull(dbData!kodesatuan)
      vaArray(n, 7) = GetNull(dbData!stok)
      dbData.MoveNext
    Loop
    
    With FrmRPT
      
      .AddPageHeader "Laporan Stock Konsinyasi", tdbHalignCenter, , , True, , 14, True, True, True, False, tdbPageHeaderSect, , , , , 3
  '    .AddPageHeader "Periode : " & Format(dDate(0).value, "dd-mm-yyyy") & " s.d " & Format(dDate(1).value, "dd-mm-yyyy"), tdbHalignCenter, , , True, , 10, True, True
  
      
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
      .AddTableHeader "Kode", , , , 8
      .AddTableHeader "Nama"
      .AddTableHeader "Beli", , , , 10
      .AddTableHeader "Jual", , , , 10
      .AddTableHeader "Satuan", , , , 6
      .AddTableHeader "Stok", , , , 6
  
      
       
      .AddTableBody , , , , , , , , , , , , , False
      .AddTableBody , , , , , , , , , , , , , False
      .AddTableBody
      .AddTableBody
      .AddTableBody Sis_Rpt_Number2, tdbHalignRight
      .AddTableBody Sis_Rpt_Number2, tdbHalignRight
      .AddTableBody
      .AddTableBody Sis_Rpt_Number, tdbHalignRight
  
      .Refresh
      .Preview vaArray, True
    End With
  Else
    MsgBox "Maat tidak ada data", vbExclamation
  End If
  

  
End Sub


