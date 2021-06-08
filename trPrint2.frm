VERSION 5.00
Object = "{A8561640-E93C-11D3-AC3B-CE6078F7B616}#1.0#0"; "VSPRINT7.ocx"
Begin VB.Form trPrint2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Invoice"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9795
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   9795
   StartUpPosition =   3  'Windows Default
   Begin VSPrinter7LibCtl.VSPrinter vp 
      Height          =   6630
      Left            =   0
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
Attribute VB_Name = "trPrint2"
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
Public dTgNota As String
Public dJthTempoNota As String
Public nTmpPoinHadiah As Double


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then
    Unload Me
  End If
End Sub

Private Sub Form_Load()
  CenterForm Me
  SetIcon Me.hwnd
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
        .Text = "INVOICE NO. " & noOrder & vbCrLf
        .Text = "Tgl." & dTgNota & " Due Date " & dJthTempoNota & vbCrLf
        .Text = "Print By. " & GetRegistry(reg_FullName) & vbCrLf
        .Text = "Jam/Tgl " & Format(SNow, "hh:mm:ss dd-MM-yyyy") & vbCrLf
        .Text = "" & vbCrLf
        .Text = "Member: [" & cKodeMember & "]" & vbCrLf
        .Text = cMember & vbCrLf
        .Text = "Telp." & cTeleponMember & vbCrLf
        
        If aCfg(objData, msOptUp) = "Y" Then
          If Trim(Ups) <> "" Then
            .Text = "UP (Kepada)--->. " & Ups
          End If
        End If
        Garis
        Dim n As Single
        Dim nQtyTemp As Single
        n = 0
        
        Set dbData = objData.Browse(GetDSN, "penjualan m", "s.barcode,m.kodestock,s.nama,m.qty,m.harga,m.discount,m.jumlah,s.kodesatuan", "m.nomorpenjualan", sisAssign, noOrder, , "m.urutfaktur", Array("left join stock s on s.kodestock = m.kodestock"))
        If Not dbData.EOF Then
          Do While Not dbData.EOF
            .TextAlign = taJustTop
            .Text = n + 1 & " - [" & GetNull(dbData!barcode) & "] " & GetNull(dbData!nama) & vbCrLf
            .Text = Padl(GetNull(dbData!qty) & " " & GetNull(dbData!kodesatuan), 5) & vbTab & " x " & Padl(Format(GetNull(dbData!Harga), "###,###,##0"), 14) & vbTab & "-" & Padl(GetNull(dbData!Discount), 2) & "%" & vbTab & " = " & Format(GetNull(dbData!jumlah), "###,###,##0") & vbCrLf
            nItemKatalog = nItemKatalog + (GetNull(dbData!qty) * GetNull(dbData!Harga))
            n = n + 1
            nQtyTemp = nQtyTemp + GetNull(dbData!qty)
            dbData.MoveNext
          Loop
        End If
        Garis
        .Text = "Total Yg Harus Dibayar: " & vbTab & vbTab & Format(nSubTotal, "###,###,###,##0") & vbCrLf
        .Text = "HK" & vbTab & Format(nItemKatalog, "###,###,##0") & vbCrLf
        .Text = "DP " & vbTab & Format(nDiscount, "###,###,##0") & vbTab & vbTab & "Tot. Qty " & vbTab & Format(nQtyTemp, "###,###,##0") & vbCrLf
        .Text = "Tunai " & vbTab & Format(nCash, "###,###,##0") & vbTab & vbTab & "Piutang " & vbTab & Format(nChange, "###,###,##0") & vbCrLf
        
        If aCfg(objData, msPoin) = 1 Then
          .Text = "POIN HADIAH " & vbTab & nTmpPoinHadiah & vbCrLf
          .Text = "POIN BERLAKU JIKA INVOICE DIBAYAR TUNAI" & vbCrLf
          .Text = "ATAU DILUNASI SEBELUM " & aCfg(objData, msTerm) & " HARI"
        End If
        
'        Set dbData = objData.Browse(GetDSN, "kartupiutang", "sum(debet-kredit) as saldopiutang", "kodeanggota", sisAssign, cKodeMember)
'        If Not dbData.EOF Then
'          If GetNull(dbData!saldopiutang) > 0 Then
'            .Text = vbCrLf
'            .Text = "INFO..!! " & vbCrLf
'            .Text = "Saldo Piutang Sampai Tgl " & vbCrLf
'            .Text = Format(Date, "dd-MM-yyyy") & " Jam. " & SNow & vbCrLf
'            .Text = "Sebesar : " & Format(GetNull(dbData!saldopiutang), "###,###,###,##0.00") & vbCrLf
'            .Text = "Silahkan dilunasi sebelum waktu jatuh tempo, Terimakasih" & vbCrLf
'          End If
'        End If
'
'        Garis
        .Text = vbCrLf
        .Text = "Note : " & vbCrLf
        If Trim(aCfg(objData, msFooterPenjualanNonTunai)) <> "" Then
          .Text = aCfg(objData, msFooterPenjualanNonTunai) & vbCrLf
        End If
        If Trim(aCfg(objData, msFooterPenjualanNonTunai2)) <> "" Then
          .Text = aCfg(objData, msFooterPenjualanNonTunai2) & vbCrLf
        End If
        
'        If Trim(aCfg(objData, msKasir3)) <> "" Then
'          .Text = aCfg(objData, msKasir3) & vbCrLf
'        End If
'        If Trim(aCfg(objData, msKasir4)) <> "" Then
'          .Text = aCfg(objData, msKasir4)
'        End If
'        If Trim(aCfg(objData, msKasir5)) <> "" Then
'          .Text = aCfg(objData, msKasir5)
'        End If
'        If Trim(aCfg(objData, msKasir6)) <> "" Then
'          .Text = aCfg(objData, msKasir6)
'        End If
        
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


