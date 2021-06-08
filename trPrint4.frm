VERSION 5.00
Object = "{A8561640-E93C-11D3-AC3B-CE6078F7B616}#1.0#0"; "VSPRINT7.ocx"
Begin VB.Form trPrint4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   9810
   StartUpPosition =   3  'Windows Default
   Begin VSPrinter7LibCtl.VSPrinter vp 
      Height          =   6720
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9780
      _cx             =   17251
      _cy             =   11853
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
Attribute VB_Name = "trPrint4"
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
        .Text = "BUKTI PENARIKAN NO. " & noOrder & vbCrLf
        .Text = "Jam/Tgl " & Format(SNow, "hh:mm:ss dd-MM-yyyy") & vbCrLf
        .Text = "Print By. " & GetRegistry(reg_FullName) & vbCrLf
        .Text = "" & vbCrLf
        
        If aCfg(objData, msOptUp) = "Y" Then
          If Trim(Ups) <> "" Then
            .Text = "UP (Kepada)--->. " & Ups
          End If
        End If
        Garis
        
        Dim n As Single
        Dim nQtyTemp As Single
        Dim nTotalTopUp As Double
        
        'inisialisasi
        n = 0
        nTotalTopUp = 0
        
        Set dbData = objData.Browse(GetDSN, "membertopup m", "m.kodeanggota,a.nama,m.keterangan,m.kredit", "m.nomormembertopup", sisAssign, noOrder, , , Array("left join anggota a on a.kodeanggota = m.kodeanggota"))
        If Not dbData.EOF Then
          Do While Not dbData.EOF
            .TextAlign = taJustTop
            .Text = n + 1 & " - [" & GetNull(dbData!kodeanggota) & "] " & GetNull(dbData!nama) & vbCrLf
            .Text = Padl(GetNull(dbData!keterangan), 14) & vbTab & vbTab & " = " & Padr(Format(GetNull(dbData!kredit), "###,###,##0"), 15) & vbCrLf
            n = n + 1
            nTotalTopUp = nTotalTopUp + GetNull(dbData!kredit)
            dbData.MoveNext
          Loop
        End If
        
        Garis
          .Text = Padl("Total", 14) & vbTab & vbTab & " = " & Format(nTotalTopUp, "###,###,##0") & vbCrLf
        Garis
          .Text = "Terimakasih, tolong disimpan bukti ini" & vbCrLf
          .Text = "Sebagai tanda penarikan uang."
      End With
    vp.EndDoc
    MousePointer = 0
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




