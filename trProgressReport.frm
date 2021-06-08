VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form trProgressReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PROGRESS REPORT"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   4455
   Begin BiSAButtonProject.BiSAButton cmdOK 
      Height          =   330
      Left            =   870
      TabIndex        =   1
      Top             =   735
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   582
      Caption         =   "Ok, Export To Excel"
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
   Begin BiSADateProject.BiSADate dTgl2 
      Height          =   330
      Left            =   2265
      TabIndex        =   0
      Top             =   315
      Width           =   1335
      _ExtentX        =   2355
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
   Begin BiSADateProject.BiSADate dTgl1 
      Height          =   345
      Left            =   720
      TabIndex        =   2
      Top             =   315
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
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
Attribute VB_Name = "trProgressReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB

Private Sub cmdOK_Click()
Dim cSQL As String
Dim n As Single
Dim a As New exportExcel
Dim nTemp As Double

  cSQL = ""
'  cSQL = "select tgl,sum(jumlah) as jumlah from pembelian p"
'  cSQL = cSQL & " Where Tgl >= '" & Format(dTgl1.Value, "yyyy-MM-dd") & "' And tgl <= '" & Format(dTgl2.Value, "yyyy-MM-dd") & "' and (Discount = 33 Or Discount = 13) AND t.kodesupplier = 'SOPHIE'"
'  cSQL = cSQL & " GROUP BY tgl ASC"
  
  cSQL = cSQL & " select p.tgl,sum(p.jumlah) as jumlah from pembelian p"
  cSQL = cSQL & " LEFT JOIN totpembelian t on t.nomorpembelian = p.nomorpembelian"
'  cSQL = cSQL & " Where p.Tgl >= '" & Format(dTgl1.Value, "yyyy-MM-dd") & "' And p.tgl <= '" & Format(dTgl2.Value, "yyyy-MM-dd") & "' and (p.Discount = 33 Or p.Discount = 13) AND t.kodesupplier = 'SOPHIE' GROUP BY p.tgl ASC"
  cSQL = cSQL & " Where p.Tgl >= '" & Format(dTgl1.Value, "yyyy-MM-dd") & "' And p.tgl <= '" & Format(dTgl2.Value, "yyyy-MM-dd") & "' and (p.Discount = 28 or p.discount = 23) AND t.kodesupplier = 'SOPHIE' GROUP BY p.tgl ASC"
    
  vaArray.ReDim 0, -1, 0, 2
  nTemp = 0
  Set dbData = objData.Sql(GetDSN, cSQL)
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = GetNull(dbData!Tgl)
      vaArray(n, 1) = GetNull(dbData!jumlah)
      nTemp = nTemp + vaArray(n, 1)
      vaArray(n, 2) = nTemp
      dbData.MoveNext
    Loop
    a.RecordSource = vaArray
    a.ExportToExcel
  End If
End Sub

Private Sub Form_Load()
Dim n As Single

  CenterForm Me
  TabIndex dTgl1, n
  TabIndex dTgl2, n
  TabIndex cmdOK, n
  
  'initvalue
  dTgl1.Value = BOM(Date)
  dTgl2.Value = EOM(Date)
  
End Sub
