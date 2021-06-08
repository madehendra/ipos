VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form rptNeracaPercobaan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Neraca Percobaan..."
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   5865
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1035
      Left            =   0
      Top             =   0
      Width           =   5865
      _ExtentX        =   10345
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
      Begin BiSADateProject.BiSADate dDate 
         Height          =   330
         Index           =   0
         Left            =   195
         TabIndex        =   0
         Top             =   375
         Width           =   3180
         _ExtentX        =   5609
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
         Caption         =   "ANTARA TANGGAL"
         CaptionWidth    =   1700
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
      Begin BiSADateProject.BiSADate dDate 
         Height          =   330
         Index           =   1
         Left            =   3480
         TabIndex        =   1
         Top             =   390
         Width           =   1995
         _ExtentX        =   3519
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
   Begin BiSAFramProject.BiSAFrame BiSAFrame2 
      Height          =   630
      Left            =   0
      Top             =   1020
      Width           =   5865
      _ExtentX        =   10345
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
         Left            =   4620
         TabIndex        =   2
         Top             =   90
         Width           =   1140
         _ExtentX        =   2011
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
         Picture         =   "rptNeracaPercobaan.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   3450
         TabIndex        =   3
         Top             =   90
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   767
         Caption         =   "     &Preview"
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
         Picture         =   "rptNeracaPercobaan.frx":00A6
      End
   End
End
Attribute VB_Name = "rptNeracaPercobaan"
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
  getSQL
End Sub

Private Sub Form_Load()
Dim n As Single
  
  CenterForm Me
  SetIcon Me.hWnd
  dDate(0).Value = BOM(Date)
  dDate(1).Value = EOM(Date)
  
  TabIndex dDate(0), n
  TabIndex dDate(1), n
  TabIndex cmdPreview, n
  TabIndex cmdKeluar, n
End Sub

Private Sub getSQL()
Dim n As Integer
Dim vaField As String
Dim cWhere As String

  vaArray.Clear
  vaArray.ReDim 0, -1, 0, 5
  vaField = "b.kodeakun,r.Keterangan,sum(b.Debet) as Debet,sum(b.Kredit) as Kredit"
  Set dbData = objData.Browse(GetDSN, "bukubesar b", vaField, "b.Tgl", sisGTEqual, Format(dDate(0).Value, "yyyy-mm-dd"), "And b.Tgl <='" & Format(dDate(1).Value, "yyyy-MM-dd") & "' Group by b.kodeakun", "b.kodeakun", _
               Array("Left Join akun r on b.kodeakun = r.kodeakun"))
  If Not dbData.EOF Then
    FrmPB.InitPB dbData.RecordCount
    dbData.MoveFirst
    Do While Not dbData.EOF
      FrmPB.RunPB
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = GetNull(dbData!kodeakun, "")
      vaArray(n, 1) = GetNull(dbData!keterangan, "")
      vaArray(n, 2) = GetMutasi(vaArray(n, 0))
      vaArray(n, 3) = GetNull(dbData!debet)
      vaArray(n, 4) = GetNull(dbData!kredit)
      vaArray(n, 5) = SumRekening(vaArray(n, 0), vaArray(n, 2), vaArray(n, 3), vaArray(n, 4))
      dbData.MoveNext
     Loop
     FrmPB.EndPB
     Rpt
    Else
      MsgBox "Maaf, mutasi tgl " & Format(dDate(0).Value, "dd/MM/yy") & " - " & Format(dDate(1).Value, "dd/MM/yy") & " tidak ada", vbInformation
      Exit Sub
    End If
End Sub

Private Function SumRekening(ByVal cRekening As String, ByVal nAwal As Double, ByVal nDebet As Double, ByVal nKredit As Double) As Double
  If Left(cRekening, 1) = "1" Or Left(cRekening, 1) = "5" Then
    SumRekening = nAwal + nDebet - nKredit
  Else
    SumRekening = nAwal - nDebet + nKredit
  End If
End Function

Private Sub Rpt()
  With FrmRPT
    .AddPageHeader UCase("Neraca Percobaan Konsolidasi"), tdbHalignCenter, , , , , 10, True
    .AddPageHeader aCfg(objData, msNamaPerusahaan), tdbHalignCenter, , , True, , 12, True
    .AddPageHeader "Antara Tanggal : " & Format(dDate(0).Value, "dd-MM-yyyy") & " s.d " & Format(dDate(1).Value, "dd-MM-yyyy"), tdbHalignCenter, , , True, , , , , , , , , , , , 10
    .AddPageHeader " ", , , , True
    .AddPageHeader " ", , , , True
    
    .AddTableHeader "REKENING", , , , 11, , , , , , True, tdbTableHeaderSect, , tdbMergeOnText, , , , 5
    .AddTableHeader "KETERANGAN", , , , , , , , , , , , , tdbMergeOnText
    .AddTableHeader "OPEN BALANCE", , , , 15, , , , , , , , , tdbMergeOnText
    .AddTableHeader "MUTASI", , , , 15, , , , , , , , , , 2
    .AddTableHeader "", , , , 15
    .AddTableHeader "CLOSE BALANCE", , , , 15, , , , , , , , , tdbMergeOnText
    
    .AddTableHeader "REKENING", , , , 10, , , , , , True, tdbTableHeaderSect, , tdbMergeOnText
    .AddTableHeader "KETERANGAN", , , , , , , , , , , , , tdbMergeOnText
    .AddTableHeader "OPEN BALANCE", , , , 15, , , , , , , , , tdbMergeOnText
    .AddTableHeader "DEBET", , , , 15
    .AddTableHeader "KREDIT", , , , 15
    .AddTableHeader "CLOSE BALANCE", , , , 15, , , , , , , , , tdbMergeOnText
    
    .AddTableBody
    .AddTableBody
    .AddTableBody Sis_Rpt_Number2, tdbHalignRight
    .AddTableBody Sis_Rpt_Number2, tdbHalignRight
    .AddTableBody Sis_Rpt_Number2, tdbHalignRight
    .AddTableBody Sis_Rpt_Number2, tdbHalignRight
    
    .AddTableFooter "TOTAL", , tdbHalignCenter, , , , , , , , , , , , 2
    .AddTableFooter
    .AddTableFooter
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    .AddTableFooter
    
    .Preview vaArray, True
  End With
End Sub

Private Function GetMutasi(ByVal cRekening As String) As Double
Dim dbData1 As New ADODB.Recordset
Dim vaField As String

  vaField = "sum(b.Debet) as debet,sum(b.Kredit)as kredit"
  Set dbData1 = objData.Browse(GetDSN, "bukubesar b", vaField, "b.kodeakun", sisAssign, cRekening, "And b.tgl < '" & Format(dDate(0).Value, "yyyy-MM-dd") & "' Group by b.kodeakun", "b.kodeakun")
  If Not dbData1.EOF Then
    GetMutasi = (dbData1!debet) - (dbData1!kredit)
  Else
    GetMutasi = 0
  End If
  
  If Not (Left(cRekening, 1) = "1" Or Left(cRekening, 1) = "5") Then
    GetMutasi = -GetMutasi
  End If
End Function
