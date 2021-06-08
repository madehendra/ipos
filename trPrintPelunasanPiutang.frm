VERSION 5.00
Object = "{A8561640-E93C-11D3-AC3B-CE6078F7B616}#1.0#0"; "VSPRINT7.ocx"
Begin VB.Form trPrintPelunasanPiutang 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Invoice"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9795
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   9795
   StartUpPosition =   3  'Windows Default
   Begin VSPrinter7LibCtl.VSPrinter vp 
      Height          =   6630
      Left            =   -15
      TabIndex        =   0
      Top             =   270
      Width           =   9780
      _cx             =   17251
      _cy             =   11695
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      MousePointer    =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty HdrFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _ConvInfo       =   1
      AutoRTF         =   -1  'True
      Preview         =   -1  'True
      DefaultDevice   =   0   'False
      PhysicalPage    =   -1  'True
      AbortWindow     =   -1  'True
      AbortWindowPos  =   0
      AbortCaption    =   "Printing..."
      AbortTextButton =   "Cancel"
      AbortTextDevice =   "on the %s on %s"
      AbortTextPage   =   "Now printing Page %d of"
      FileName        =   ""
      MarginLeft      =   1440
      MarginTop       =   1440
      MarginRight     =   1440
      MarginBottom    =   1440
      MarginHeader    =   0
      MarginFooter    =   0
      IndentLeft      =   0
      IndentRight     =   0
      IndentFirst     =   0
      IndentTab       =   720
      SpaceBefore     =   0
      SpaceAfter      =   0
      LineSpacing     =   100
      Columns         =   1
      ColumnSpacing   =   180
      ShowGuides      =   2
      LargeChangeHorz =   300
      LargeChangeVert =   300
      SmallChangeHorz =   30
      SmallChangeVert =   30
      Track           =   0   'False
      ProportionalBars=   -1  'True
      Zoom            =   100
      ZoomMode        =   0
      ZoomMax         =   400
      ZoomMin         =   10
      ZoomStep        =   25
      EmptyColor      =   -2147483636
      TextColor       =   0
      HdrColor        =   0
      BrushColor      =   0
      BrushStyle      =   0
      PenColor        =   0
      PenStyle        =   0
      PenWidth        =   0
      PageBorder      =   0
      Header          =   ""
      Footer          =   ""
      TableSep        =   "|;"
      TableBorder     =   7
      TablePen        =   0
      TablePenLR      =   0
      TablePenTB      =   0
      NavBar          =   3
      NavBarColor     =   -2147483633
      ExportFormat    =   0
      URL             =   ""
      Navigation      =   3
      NavBarMenuText  =   "Whole &Page|Page &Width|&Two Pages|Thumb&nail"
   End
End
Attribute VB_Name = "trPrintPelunasanPiutang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Public noOrder As String
Public nTotal As Double
Public nSubTotal As Double
Public nCash As Double
Public nChange As Double
Public nTax As String
Public nDiscount As String
Public lStatus As Double
Public cMember As String
Public cKodeMember As String
Public cTeleponMember As String
Public Ups As String
Public nSaldoTopUp As Double
Public nWithDraw As Double
Public nTarikTunai As Double
Public nKembali1 As Double
Public nKembali2 As Double
Public nSisa As Double
Public nTunai As Double
Public lKembali As Boolean
Public nMetodePembayaran As Integer
Public nPoinHadiah As Double

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then
    Unload Me
  End If
End Sub

Private Sub Form_Load()
  CenterForm Me
  SetIcon Me.hWnd
  DoText
End Sub

Private Sub DoText()
Dim i%
Dim nTahun As Double
Dim nBulan As Double
Dim nHari As Double
Dim nItemKatalog As Double

    MousePointer = 11
    SetOriginalSettings
    vp.ZoomMode = zmPageWidth
    vp.StartDoc
      With vp
        Dim n As Single
        Dim nQtyTemp As Single
        n = 0
        
        Dim cSQL As String
        cSQL = "select p.nomorpelunasanpiutang,p.piutang,p.nomorpenjualan,p.pelunasan,t.kodeanggota,a.nama,t.total,tt.tgl,tt.total as totalnota from pelunasanpiutang p"
        cSQL = cSQL & " left join totpelunasanpiutang t on t.nomorpelunasanpiutang = p.nomorpelunasanpiutang"
        cSQL = cSQL & " left join anggota a on a.kodeanggota = t.kodeanggota"
        cSQL = cSQL & " left join totpenjualan tt on tt.nomorpenjualan = p.nomorpenjualan"
        cSQL = cSQL & " where p.nomorpelunasanpiutang = '" & noOrder & "'"
        
        Set dbData = objData.Sql(GetDSN, cSQL)
        If Not dbData.EOF Then
           vp = " "
          .FontSize = 12
          .TextColor = vbBlack
          .Text = aCfg(objData, msNamaPerusahaan) & vbCrLf
          .FontSize = 10
          .Text = aCfg(objData, msAlamatPerusahaan) & vbCrLf
          .Text = aCfg(objData, msTelepon) & vbCrLf
          .Text = "" & vbCrLf
                  
          .FontSize = 9
          .FontName = "Tahoma"
          .Text = "Nota Pelunasan No. " & noOrder & vbCrLf
'          .Text = "Tgl Nota. " & Format(GetNull(dbData!tgl), "dd-MM-yyyy") & vbCrLf & vbCrLf
          
          .Text = "Print By. " & GetRegistry(reg_FullName) & vbCrLf
          .Text = "Jam/Tgl " & Format(SNow, "hh:mm:ss dd-MM-yyyy") & vbCrLf
'          .Text = "" & vbCrLf
          
          .Text = "Member: [" & cKodeMember & "]" & vbCrLf
          .Text = cMember & vbCrLf
          Garis
          Do While Not dbData.EOF
            .TextAlign = taJustTop
            .Text = n + 1 & " - Nota. [" & GetNull(dbData!nomorpenjualan) & "] " & Format(GetNull(dbData!Tgl), "dd/MM/yyyy") & " = " & Format(GetNull(dbData!totalnota), "###,###,##0.00") & vbCrLf
'            .Text = "Sisa : " & Format(GetNull(dbData!Piutang), "###,###,##0.00") & " Lunas : " & Format(GetNull(dbData!Pelunasan), "###,###,##0.00") & vbCrLf
            n = n + 1
            dbData.MoveNext
          Loop
        End If

'        Garis
        
        .Text = "Total: " & vbTab & vbTab & Format(nSubTotal, "###,###,###,##0") & vbCrLf
        
        If nMetodePembayaran = 0 Then
          'jika pembayaran tunai
                    
        End If
        
        If nMetodePembayaran = 2 Then
          .Text = "**Metode Pembayaran : Top Up**" & vbCrLf
          .Text = "Saldo Top Up " & vbTab & Format(nSaldoTopUp, "###,###,##0") & vbCrLf
          'jika nilai top up lebih dari yg dibayarkan
          If nKembali1 > 0 Then
            If lKembali = True Then
              .Text = "(+)Kembali " & vbTab & vbTab & Format(nKembali1, "###,###,##0") & vbCrLf
            Else
              .Text = "(+)Sisa Top Up " & vbTab & vbTab & Format(nKembali1, "###,###,##0") & vbCrLf
            End If
          Else
            'jika nilai top up kurang
            .Text = "(-)Kurang " & vbTab & vbTab & Format(nSisa, "###,###,##0") & vbCrLf
            .Text = "(+)Tunai " & vbTab & vbTab & Format(nTunai, "###,###,##0") & vbCrLf
            .Text = "(=)Kembali " & vbTab & vbTab & Format(nKembali2, "###,###,##0") & vbCrLf
          End If
        End If
        .Text = vbCrLf
        
        If aCfg(objData, msPoin) = 1 Then
          .Text = "POIN HADIAH " & vbTab & nPoinHadiah & vbCrLf
          .Text = "BERLAKU S/D TGL " & Format(DateAdd("D", aCfg(objData, msTerm), Date), "dd-MM-yyyy") & vbCrLf & vbCrLf
          .Text = "**** KUMPULKAN TERUS POIN HADIAH INI ****" & vbCrLf
          .Text = "***DAN TUKARKAN DENGAN HADIAH LANGSUNG***" & vbCrLf
          .Text = "---------TUKARKAN SEBELUM EXPIRED--------"
          '.Text = "Sisa Piutang " & vbTab & Format(GetSaldoPiutang(objData, cKodeMember), "###,###,##0") & vbCrLf
        End If
      End With
      
    vp.EndDoc
    MousePointer = 0
    'simpan poin belanja ke dalam tabel
    
    If nPoinHadiah > 0 Then
      Dim vaField
      Dim vaValue
      Dim lSave As Boolean
      
      lSave = True
      
      vaField = Array("faktur", "tgl", "kodeanggota", "poinhadiah", "exdate", "status")
      vaValue = Array(noOrder, Date, cKodeMember, nPoinHadiah, DateAdd("D", aCfg(objData, msTerm), Date), "1")
      lSave = IIf(lSave, objData.Add(GetDSN, "poinhadiah", vaField, vaValue), False)
      
      If lSave Then
        objData.Save GetDSN
      Else
        objData.Cancel GetDSN
      End If
    End If
End Sub

Private Sub Garis()
Dim a As Integer
  
  vp.Text = vbCrLf
  For a = 1 To 33
    vp.Text = "="
  Next a
  vp.Text = vbCrLf
End Sub

Private Sub SetOriginalSettings()
    With vp
        .PaperSize = pprUser
        
        .ToolTipText = ""
        ' font
        .FontName = "Arial"
        .FontBold = False
        .FontItalic = False
        .FontUnderline = False
        .FontSize = 11
        
        ' text
        .TextColor = 0
        .TextAngle = 0
        .TextAlign = taLeftMiddle
        
        'spacing
        .IndentLeft = 0
        .IndentFirst = 0
        .IndentRight = 0

        .MarginBottom = 0
        .MarginFooter = 0
        .MarginLeft = 500
        .MarginHeader = 0
        .MarginRight = 1440
        .MarginTop = 0
        

        'drawing
        .PenColor = vbBlack
        .PenStyle = psSolid
        .PenWidth = 1
        .BrushColor = &H8080FF
        .BrushStyle = bsSolid
    
        ' table
        .TableBorder = tbAll
    
        .X1 = 0
        .Y1 = 0
        .X2 = 0
        .Y2 = 0
    
    End With

End Sub


