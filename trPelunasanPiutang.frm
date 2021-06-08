VERSION 5.00
Object = "{0D6235E7-DBA2-11D1-B5DF-0060976089D0}#1.0#0"; "tdbr6.ocx"
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form trPelunasanPiutang 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PELUNASAN PIUTANG"
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   14400
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   14400
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   2640
      Left            =   7155
      Top             =   60
      Width           =   6945
      _ExtentX        =   12250
      _ExtentY        =   4657
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
      Begin BiSANumberBoxProject.BiSANumberBox nSaldoPiutang 
         Height          =   330
         Left            =   4455
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   885
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   582
         Appearance      =   0
         Enabled         =   0   'False
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
         Caption         =   " "
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
      Begin BiSANumberBoxProject.BiSANumberBox nSaldoTopUpMember 
         Height          =   330
         Left            =   4440
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   525
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   582
         Appearance      =   0
         Enabled         =   0   'False
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
         Caption         =   " "
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
      Begin TrueDBReports60Ctl.TDBReports rptKuitansiLunas 
         Height          =   570
         Left            =   135
         TabIndex        =   21
         Top             =   1770
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   1005
         Caption         =   "Kuitansi Lunas"
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         ErrorMsgCaption =   ""
         Filtered        =   0   'False
         DataMode        =   1
         DataMember      =   ""
         LinkSequence    =   1
         LinkOrder       =   0
         NameSubstitute  =   ""
         ConnectionString=   "DSN=MySalemba"
         ConnectStringType=   3
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   "MySalemba"
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         CursorLocation  =   3
         ConnectionTimeout=   15
         CommandTimeout  =   30
         RecordSource    =   ""
         CursorType      =   1
         CommandType     =   8
         MaxRecords      =   0
         LinkType        =   0
         Master          =   ""
         CallDataRead    =   0   'False
         ConvertNullToEmpty=   -1  'True
         DesignConnection=   -1  'True
         DesignTimeout   =   5
         UnitsOfMeasurement=   4
         Vedit_ShowGrid  =   -1  'True
         Vedit_SnapToGrid=   0   'False
         Vedit_GridUnitWidth=   2.822
         Vedit_GridUnitHeight=   2.822
         Vedit_ShowCellExpressions=   -1  'True
         Norm_rect_left  =   0
         Norm_rect_top   =   0
         Norm_rect_right =   0
         Norm_rect_bottom=   0
         Virgin          =   0   'False
         Parameters.Count=   29
         Parameters(0).Name=   "cSE"
         Parameters(0).ValueExpression=   """"""
         Parameters(1).Name=   "cNama"
         Parameters(1).ValueExpression=   """"""
         Parameters(2).Name=   "cAlamat"
         Parameters(2).ValueExpression=   """"""
         Parameters(3).Name=   "cKota"
         Parameters(3).ValueExpression=   """"""
         Parameters(4).Name=   "cTerbilang"
         Parameters(4).ValueExpression=   """"""
         Parameters(5).Name=   "dTgl"
         Parameters(6).Name=   "dJTHTMP"
         Parameters(6).Type=   7
         Parameters(7).Name=   "cTTD"
         Parameters(8).Name=   "nSubTotal"
         Parameters(8).Type=   5
         Parameters(8).ValueExpression=   "0"
         Parameters(9).Name=   "nTotal"
         Parameters(9).Type=   5
         Parameters(9).ValueExpression=   "0"
         Parameters(10).Name=   "nPPn"
         Parameters(10).ValueExpression=   "0"
         Parameters(11).Name=   "nPajak"
         Parameters(11).Type=   5
         Parameters(11).ValueExpression=   "0"
         Parameters(12).Name=   "cNamaPerusahaan"
         Parameters(13).Name=   "cAlamatPerusahaan"
         Parameters(14).Name=   "cTeleponPerusahaan"
         Parameters(15).Name=   "cReceived"
         Parameters(16).Name=   "cKetReceived"
         Parameters(17).Name=   "cRef"
         Parameters(18).Name=   "cPerusahaanLine"
         Parameters(19).Name=   "cPayment"
         Parameters(20).Name=   "cUserName"
         Parameters(21).Name=   "nDiscount"
         Parameters(21).Type=   5
         Parameters(22).Name=   "cJudul"
         Parameters(23).Name=   "cSales"
         Parameters(24).Name=   "cFooter"
         Parameters(25).Name=   "nDp"
         Parameters(26).Name=   "cFooter2"
         Parameters(27).Name=   "cKodeAnggota"
         Parameters(28).Name=   "keAkun"
         Fields.Count    =   6
         Fields(0).Name  =   "Nomor"
         Fields(0).DisplayName=   "Nomor"
         Fields(0).Type  =   2
         Fields(1).Name  =   "NoInvoice"
         Fields(1).DisplayName=   "NoInvoice"
         Fields(2).Name  =   "Total"
         Fields(2).DisplayName=   "Total"
         Fields(2).Type  =   5
         Fields(3).Name  =   "JatuhTempo"
         Fields(3).DisplayName=   "JatuhTempo"
         Fields(3).Type  =   7
         Fields(4).Name  =   "Denda"
         Fields(4).DisplayName=   "Denda"
         Fields(4).Type  =   5
         Fields(5).Name  =   "Jumlah"
         Fields(5).DisplayName=   "Jumlah"
         Fields(5).Type  =   5
         Sections.Count  =   6
         Sections(0).Name=   "SECTION_2"
         Sections(0).Type=   1
         Sections(0).StyleExp=   "'Tdb_Base'"
         Sections(0).Cells.Count=   16
         Sections(0).Cells(0).Name=   "CELL_22"
         Sections(0).Cells(0).Exp=   "cNamaPerusahaan"
         Sections(0).Cells(0).NewLine=   -1  'True
         Sections(0).Cells(0).PrivateStyle=   -1  'True
         Sections(0).Cells(0).Style.Name=   "<private>"
         Sections(0).Cells(0).Style.ParentName=   "Tdb_Base"
         Sections(0).Cells(0).Style.Font_Name=   "Courier"
         Sections(0).Cells(0).Style.Font_Size=   12
         Sections(0).Cells(0).Style.Font_Bold=   -1  'True
         Sections(0).Cells(0).Style.Font_Italic=   0   'False
         Sections(0).Cells(0).Style.Font_Underline=   0   'False
         Sections(0).Cells(0).Style.Font_Strikeout=   0   'False
         Sections(0).Cells(0).Style.Font_Charset=   0
         Sections(0).Cells(0).Style.TextAlign=   0
         Sections(0).Cells(0).Style.TextVAlign=   1
         Sections(0).Cells(0).Style.TextWrap=   -1  'True
         Sections(0).Cells(0).Style.ForeColor=   0
         Sections(0).Cells(0).Style.BackColor=   16777215
         Sections(0).Cells(0).Style.NoFill=   -1  'True
         Sections(0).Cells(0).Style.BackPicFile=   ""
         Sections(0).Cells(0).Style.ForePicFile=   ""
         Sections(0).Cells(0).Style.BackPicVertPlacement=   0
         Sections(0).Cells(0).Style.BackPicHorzPlacement=   0
         Sections(0).Cells(0).Style.ForePicPlacement=   0
         Sections(0).Cells(0).Style.ForePicDrawMode=   0
         Sections(0).Cells(0).Style.MarginLeft=   6
         Sections(0).Cells(0).Style.MarginTop=   1
         Sections(0).Cells(0).Style.MarginRight=   6
         Sections(0).Cells(0).Style.MarginBottom=   1
         Sections(0).Cells(0).Style.HasBorders=   -1  'True
         Sections(0).Cells(0).Style.BorderHT=   ""
         Sections(0).Cells(0).Style.BorderHI=   ""
         Sections(0).Cells(0).Style.BorderHB=   ""
         Sections(0).Cells(0).Style.BorderVL=   ""
         Sections(0).Cells(0).Style.BorderVI=   ""
         Sections(0).Cells(0).Style.BorderVR=   ""
         Sections(0).Cells(0).Style.NoClipping=   0   'False
         Sections(0).Cells(0).Style.RTF=   0   'False
         Sections(0).Cells(0).Style.fprops=   89391105
         Sections(0).Cells(1).Name=   "CELL_25"
         Sections(0).Cells(1).Exp=   "cAlamatPerusahaan"
         Sections(0).Cells(1).NewLine=   -1  'True
         Sections(0).Cells(1).Height=   5
         Sections(0).Cells(1).AutoHeight=   0   'False
         Sections(0).Cells(1).PrivateStyle=   -1  'True
         Sections(0).Cells(1).Style.Name=   "<private>"
         Sections(0).Cells(1).Style.ParentName=   "Tdb_Base"
         Sections(0).Cells(1).Style.Font_Name=   "Courier"
         Sections(0).Cells(1).Style.Font_Size=   9.75
         Sections(0).Cells(1).Style.Font_Bold=   0   'False
         Sections(0).Cells(1).Style.Font_Italic=   0   'False
         Sections(0).Cells(1).Style.Font_Underline=   0   'False
         Sections(0).Cells(1).Style.Font_Strikeout=   0   'False
         Sections(0).Cells(1).Style.Font_Charset=   0
         Sections(0).Cells(1).Style.TextAlign=   0
         Sections(0).Cells(1).Style.TextVAlign=   1
         Sections(0).Cells(1).Style.TextWrap=   -1  'True
         Sections(0).Cells(1).Style.ForeColor=   0
         Sections(0).Cells(1).Style.BackColor=   16777215
         Sections(0).Cells(1).Style.NoFill=   -1  'True
         Sections(0).Cells(1).Style.BackPicFile=   ""
         Sections(0).Cells(1).Style.ForePicFile=   ""
         Sections(0).Cells(1).Style.BackPicVertPlacement=   0
         Sections(0).Cells(1).Style.BackPicHorzPlacement=   0
         Sections(0).Cells(1).Style.ForePicPlacement=   0
         Sections(0).Cells(1).Style.ForePicDrawMode=   0
         Sections(0).Cells(1).Style.MarginLeft=   6
         Sections(0).Cells(1).Style.MarginTop=   1
         Sections(0).Cells(1).Style.MarginRight=   6
         Sections(0).Cells(1).Style.MarginBottom=   1
         Sections(0).Cells(1).Style.HasBorders=   0   'False
         Sections(0).Cells(1).Style.BorderHT=   ""
         Sections(0).Cells(1).Style.BorderHI=   ""
         Sections(0).Cells(1).Style.BorderHB=   ""
         Sections(0).Cells(1).Style.BorderVL=   ""
         Sections(0).Cells(1).Style.BorderVI=   ""
         Sections(0).Cells(1).Style.BorderVR=   ""
         Sections(0).Cells(1).Style.NoClipping=   0   'False
         Sections(0).Cells(1).Style.RTF=   0   'False
         Sections(0).Cells(1).Style.fprops=   22413313
         Sections(0).Cells(2).Name=   "CELL_2"
         Sections(0).Cells(2).Exp=   """"""
         Sections(0).Cells(2).NewLine=   -1  'True
         Sections(0).Cells(2).PrivateStyle=   -1  'True
         Sections(0).Cells(2).Style.Name=   "<private>"
         Sections(0).Cells(2).Style.ParentName=   "Tdb_Base"
         Sections(0).Cells(2).Style.Font_Name=   "Courier"
         Sections(0).Cells(2).Style.Font_Size=   9.75
         Sections(0).Cells(2).Style.Font_Bold=   -1  'True
         Sections(0).Cells(2).Style.Font_Italic=   0   'False
         Sections(0).Cells(2).Style.Font_Underline=   0   'False
         Sections(0).Cells(2).Style.Font_Strikeout=   0   'False
         Sections(0).Cells(2).Style.Font_Charset=   0
         Sections(0).Cells(2).Style.TextAlign=   3
         Sections(0).Cells(2).Style.TextVAlign=   1
         Sections(0).Cells(2).Style.TextWrap=   -1  'True
         Sections(0).Cells(2).Style.ForeColor=   0
         Sections(0).Cells(2).Style.BackColor=   16777215
         Sections(0).Cells(2).Style.NoFill=   -1  'True
         Sections(0).Cells(2).Style.BackPicFile=   ""
         Sections(0).Cells(2).Style.ForePicFile=   ""
         Sections(0).Cells(2).Style.BackPicVertPlacement=   0
         Sections(0).Cells(2).Style.BackPicHorzPlacement=   0
         Sections(0).Cells(2).Style.ForePicPlacement=   0
         Sections(0).Cells(2).Style.ForePicDrawMode=   0
         Sections(0).Cells(2).Style.MarginLeft=   6
         Sections(0).Cells(2).Style.MarginTop=   1
         Sections(0).Cells(2).Style.MarginRight=   6
         Sections(0).Cells(2).Style.MarginBottom=   1
         Sections(0).Cells(2).Style.HasBorders=   -1  'True
         Sections(0).Cells(2).Style.BorderHT=   ""
         Sections(0).Cells(2).Style.BorderHI=   ""
         Sections(0).Cells(2).Style.BorderHB=   ""
         Sections(0).Cells(2).Style.BorderVL=   ""
         Sections(0).Cells(2).Style.BorderVI=   ""
         Sections(0).Cells(2).Style.BorderVR=   ""
         Sections(0).Cells(2).Style.NoClipping=   0   'False
         Sections(0).Cells(2).Style.RTF=   0   'False
         Sections(0).Cells(2).Style.fprops=   131072
         Sections(0).Cells(3).Name=   "CELL_26"
         Sections(0).Cells(3).Exp=   """ """
         Sections(0).Cells(3).NewLine=   -1  'True
         Sections(0).Cells(3).PrivateStyle=   -1  'True
         Sections(0).Cells(3).Style.Name=   "<private>"
         Sections(0).Cells(3).Style.ParentName=   "Tdb_Base"
         Sections(0).Cells(3).Style.Font_Name=   "Courier"
         Sections(0).Cells(3).Style.Font_Size=   9.75
         Sections(0).Cells(3).Style.Font_Bold=   -1  'True
         Sections(0).Cells(3).Style.Font_Italic=   0   'False
         Sections(0).Cells(3).Style.Font_Underline=   0   'False
         Sections(0).Cells(3).Style.Font_Strikeout=   0   'False
         Sections(0).Cells(3).Style.Font_Charset=   0
         Sections(0).Cells(3).Style.TextAlign=   1
         Sections(0).Cells(3).Style.TextVAlign=   1
         Sections(0).Cells(3).Style.TextWrap=   -1  'True
         Sections(0).Cells(3).Style.ForeColor=   0
         Sections(0).Cells(3).Style.BackColor=   16777215
         Sections(0).Cells(3).Style.NoFill=   -1  'True
         Sections(0).Cells(3).Style.BackPicFile=   ""
         Sections(0).Cells(3).Style.ForePicFile=   ""
         Sections(0).Cells(3).Style.BackPicVertPlacement=   0
         Sections(0).Cells(3).Style.BackPicHorzPlacement=   0
         Sections(0).Cells(3).Style.ForePicPlacement=   0
         Sections(0).Cells(3).Style.ForePicDrawMode=   0
         Sections(0).Cells(3).Style.MarginLeft=   6
         Sections(0).Cells(3).Style.MarginTop=   1
         Sections(0).Cells(3).Style.MarginRight=   6
         Sections(0).Cells(3).Style.MarginBottom=   1
         Sections(0).Cells(3).Style.HasBorders=   -1  'True
         Sections(0).Cells(3).Style.BorderHT=   ""
         Sections(0).Cells(3).Style.BorderHI=   ""
         Sections(0).Cells(3).Style.BorderHB=   ""
         Sections(0).Cells(3).Style.BorderVL=   ""
         Sections(0).Cells(3).Style.BorderVI=   ""
         Sections(0).Cells(3).Style.BorderVR=   ""
         Sections(0).Cells(3).Style.NoClipping=   0   'False
         Sections(0).Cells(3).Style.RTF=   0   'False
         Sections(0).Cells(3).Style.fprops=   68419585
         Sections(0).Cells(4).Name=   "CELL_3"
         Sections(0).Cells(4).Exp=   """ """
         Sections(0).Cells(4).NewLine=   -1  'True
         Sections(0).Cells(4).Width=   30
         Sections(0).Cells(4).PrivateStyle=   -1  'True
         Sections(0).Cells(4).Style.Name=   "<private>"
         Sections(0).Cells(4).Style.ParentName=   "Tdb_Base"
         Sections(0).Cells(4).Style.Font_Name=   "Courier"
         Sections(0).Cells(4).Style.Font_Size=   9.75
         Sections(0).Cells(4).Style.Font_Bold=   0   'False
         Sections(0).Cells(4).Style.Font_Italic=   0   'False
         Sections(0).Cells(4).Style.Font_Underline=   0   'False
         Sections(0).Cells(4).Style.Font_Strikeout=   0   'False
         Sections(0).Cells(4).Style.Font_Charset=   0
         Sections(0).Cells(4).Style.TextAlign=   3
         Sections(0).Cells(4).Style.TextVAlign=   1
         Sections(0).Cells(4).Style.TextWrap=   -1  'True
         Sections(0).Cells(4).Style.ForeColor=   0
         Sections(0).Cells(4).Style.BackColor=   16777215
         Sections(0).Cells(4).Style.NoFill=   -1  'True
         Sections(0).Cells(4).Style.BackPicFile=   ""
         Sections(0).Cells(4).Style.ForePicFile=   ""
         Sections(0).Cells(4).Style.BackPicVertPlacement=   0
         Sections(0).Cells(4).Style.BackPicHorzPlacement=   0
         Sections(0).Cells(4).Style.ForePicPlacement=   0
         Sections(0).Cells(4).Style.ForePicDrawMode=   0
         Sections(0).Cells(4).Style.MarginLeft=   6
         Sections(0).Cells(4).Style.MarginTop=   1
         Sections(0).Cells(4).Style.MarginRight=   6
         Sections(0).Cells(4).Style.MarginBottom=   1
         Sections(0).Cells(4).Style.HasBorders=   -1  'True
         Sections(0).Cells(4).Style.BorderHT=   ""
         Sections(0).Cells(4).Style.BorderHI=   ""
         Sections(0).Cells(4).Style.BorderHB=   ""
         Sections(0).Cells(4).Style.BorderVL=   ""
         Sections(0).Cells(4).Style.BorderVI=   ""
         Sections(0).Cells(4).Style.BorderVR=   ""
         Sections(0).Cells(4).Style.NoClipping=   0   'False
         Sections(0).Cells(4).Style.RTF=   0   'False
         Sections(0).Cells(4).Style.fprops=   22282240
         Sections(0).Cells(5).Name=   "CELL_27"
         Sections(0).Cells(5).Exp=   """Kuitansi"""
         Sections(0).Cells(5).PrivateStyle=   -1  'True
         Sections(0).Cells(5).Style.Name=   "<private>"
         Sections(0).Cells(5).Style.ParentName=   "Tdb_Base"
         Sections(0).Cells(5).Style.Font_Name=   "Courier"
         Sections(0).Cells(5).Style.Font_Size=   9.75
         Sections(0).Cells(5).Style.Font_Bold=   -1  'True
         Sections(0).Cells(5).Style.Font_Italic=   0   'False
         Sections(0).Cells(5).Style.Font_Underline=   0   'False
         Sections(0).Cells(5).Style.Font_Strikeout=   0   'False
         Sections(0).Cells(5).Style.Font_Charset=   0
         Sections(0).Cells(5).Style.TextAlign=   1
         Sections(0).Cells(5).Style.TextVAlign=   1
         Sections(0).Cells(5).Style.TextWrap=   -1  'True
         Sections(0).Cells(5).Style.ForeColor=   0
         Sections(0).Cells(5).Style.BackColor=   16777215
         Sections(0).Cells(5).Style.NoFill=   -1  'True
         Sections(0).Cells(5).Style.BackPicFile=   ""
         Sections(0).Cells(5).Style.ForePicFile=   ""
         Sections(0).Cells(5).Style.BackPicVertPlacement=   0
         Sections(0).Cells(5).Style.BackPicHorzPlacement=   0
         Sections(0).Cells(5).Style.ForePicPlacement=   0
         Sections(0).Cells(5).Style.ForePicDrawMode=   0
         Sections(0).Cells(5).Style.MarginLeft=   6
         Sections(0).Cells(5).Style.MarginTop=   1
         Sections(0).Cells(5).Style.MarginRight=   6
         Sections(0).Cells(5).Style.MarginBottom=   1
         Sections(0).Cells(5).Style.HasBorders=   -1  'True
         Sections(0).Cells(5).Style.BorderHT=   ""
         Sections(0).Cells(5).Style.BorderHI=   ""
         Sections(0).Cells(5).Style.BorderHB=   "None"
         Sections(0).Cells(5).Style.BorderVL=   ""
         Sections(0).Cells(5).Style.BorderVI=   ""
         Sections(0).Cells(5).Style.BorderVR=   ""
         Sections(0).Cells(5).Style.NoClipping=   0   'False
         Sections(0).Cells(5).Style.RTF=   0   'False
         Sections(0).Cells(5).Style.fprops=   84017153
         Sections(0).Cells(6).Name=   "CELL_12"
         Sections(0).Cells(6).Exp=   """Tgl : "" & dTgl"
         Sections(0).Cells(6).Width=   30
         Sections(0).Cells(6).PrivateStyle=   -1  'True
         Sections(0).Cells(6).Style.Name=   "<private>"
         Sections(0).Cells(6).Style.ParentName=   "Tdb_Base"
         Sections(0).Cells(6).Style.Font_Name=   "Courier"
         Sections(0).Cells(6).Style.Font_Size=   9.75
         Sections(0).Cells(6).Style.Font_Bold=   0   'False
         Sections(0).Cells(6).Style.Font_Italic=   0   'False
         Sections(0).Cells(6).Style.Font_Underline=   0   'False
         Sections(0).Cells(6).Style.Font_Strikeout=   0   'False
         Sections(0).Cells(6).Style.Font_Charset=   0
         Sections(0).Cells(6).Style.TextAlign=   2
         Sections(0).Cells(6).Style.TextVAlign=   1
         Sections(0).Cells(6).Style.TextWrap=   -1  'True
         Sections(0).Cells(6).Style.ForeColor=   0
         Sections(0).Cells(6).Style.BackColor=   16777215
         Sections(0).Cells(6).Style.NoFill=   -1  'True
         Sections(0).Cells(6).Style.BackPicFile=   ""
         Sections(0).Cells(6).Style.ForePicFile=   ""
         Sections(0).Cells(6).Style.BackPicVertPlacement=   0
         Sections(0).Cells(6).Style.BackPicHorzPlacement=   0
         Sections(0).Cells(6).Style.ForePicPlacement=   0
         Sections(0).Cells(6).Style.ForePicDrawMode=   0
         Sections(0).Cells(6).Style.MarginLeft=   6
         Sections(0).Cells(6).Style.MarginTop=   1
         Sections(0).Cells(6).Style.MarginRight=   6
         Sections(0).Cells(6).Style.MarginBottom=   1
         Sections(0).Cells(6).Style.HasBorders=   -1  'True
         Sections(0).Cells(6).Style.BorderHT=   ""
         Sections(0).Cells(6).Style.BorderHI=   ""
         Sections(0).Cells(6).Style.BorderHB=   ""
         Sections(0).Cells(6).Style.BorderVL=   ""
         Sections(0).Cells(6).Style.BorderVI=   ""
         Sections(0).Cells(6).Style.BorderVR=   ""
         Sections(0).Cells(6).Style.NoClipping=   0   'False
         Sections(0).Cells(6).Style.RTF=   0   'False
         Sections(0).Cells(6).Style.fprops=   18087937
         Sections(0).Cells(7).Name=   "CELL_13"
         Sections(0).Cells(7).NewLine=   -1  'True
         Sections(0).Cells(7).Width=   30
         Sections(0).Cells(7).PrivateStyle=   -1  'True
         Sections(0).Cells(7).Style.Name=   "<private>"
         Sections(0).Cells(7).Style.ParentName=   "Tdb_Base"
         Sections(0).Cells(7).Style.Font_Name=   "Courier"
         Sections(0).Cells(7).Style.Font_Size=   9.75
         Sections(0).Cells(7).Style.Font_Bold=   0   'False
         Sections(0).Cells(7).Style.Font_Italic=   0   'False
         Sections(0).Cells(7).Style.Font_Underline=   0   'False
         Sections(0).Cells(7).Style.Font_Strikeout=   0   'False
         Sections(0).Cells(7).Style.Font_Charset=   0
         Sections(0).Cells(7).Style.TextAlign=   3
         Sections(0).Cells(7).Style.TextVAlign=   1
         Sections(0).Cells(7).Style.TextWrap=   -1  'True
         Sections(0).Cells(7).Style.ForeColor=   0
         Sections(0).Cells(7).Style.BackColor=   16777215
         Sections(0).Cells(7).Style.NoFill=   -1  'True
         Sections(0).Cells(7).Style.BackPicFile=   ""
         Sections(0).Cells(7).Style.ForePicFile=   ""
         Sections(0).Cells(7).Style.BackPicVertPlacement=   0
         Sections(0).Cells(7).Style.BackPicHorzPlacement=   0
         Sections(0).Cells(7).Style.ForePicPlacement=   0
         Sections(0).Cells(7).Style.ForePicDrawMode=   0
         Sections(0).Cells(7).Style.MarginLeft=   6
         Sections(0).Cells(7).Style.MarginTop=   1
         Sections(0).Cells(7).Style.MarginRight=   6
         Sections(0).Cells(7).Style.MarginBottom=   1
         Sections(0).Cells(7).Style.HasBorders=   -1  'True
         Sections(0).Cells(7).Style.BorderHT=   ""
         Sections(0).Cells(7).Style.BorderHI=   ""
         Sections(0).Cells(7).Style.BorderHB=   ""
         Sections(0).Cells(7).Style.BorderVL=   ""
         Sections(0).Cells(7).Style.BorderVI=   ""
         Sections(0).Cells(7).Style.BorderVR=   ""
         Sections(0).Cells(7).Style.NoClipping=   0   'False
         Sections(0).Cells(7).Style.RTF=   0   'False
         Sections(0).Cells(7).Style.fprops=   22413312
         Sections(0).Cells(8).Name=   "CELL_14"
         Sections(0).Cells(8).Exp=   """No. "" & cSE"
         Sections(0).Cells(8).PrivateStyle=   -1  'True
         Sections(0).Cells(8).Style.Name=   "<private>"
         Sections(0).Cells(8).Style.ParentName=   "Tdb_Base"
         Sections(0).Cells(8).Style.Font_Name=   "Courier"
         Sections(0).Cells(8).Style.Font_Size=   9.75
         Sections(0).Cells(8).Style.Font_Bold=   0   'False
         Sections(0).Cells(8).Style.Font_Italic=   0   'False
         Sections(0).Cells(8).Style.Font_Underline=   0   'False
         Sections(0).Cells(8).Style.Font_Strikeout=   0   'False
         Sections(0).Cells(8).Style.Font_Charset=   0
         Sections(0).Cells(8).Style.TextAlign=   1
         Sections(0).Cells(8).Style.TextVAlign=   1
         Sections(0).Cells(8).Style.TextWrap=   -1  'True
         Sections(0).Cells(8).Style.ForeColor=   0
         Sections(0).Cells(8).Style.BackColor=   16777215
         Sections(0).Cells(8).Style.NoFill=   -1  'True
         Sections(0).Cells(8).Style.BackPicFile=   ""
         Sections(0).Cells(8).Style.ForePicFile=   ""
         Sections(0).Cells(8).Style.BackPicVertPlacement=   0
         Sections(0).Cells(8).Style.BackPicHorzPlacement=   0
         Sections(0).Cells(8).Style.ForePicPlacement=   0
         Sections(0).Cells(8).Style.ForePicDrawMode=   0
         Sections(0).Cells(8).Style.MarginLeft=   6
         Sections(0).Cells(8).Style.MarginTop=   1
         Sections(0).Cells(8).Style.MarginRight=   6
         Sections(0).Cells(8).Style.MarginBottom=   1
         Sections(0).Cells(8).Style.HasBorders=   -1  'True
         Sections(0).Cells(8).Style.BorderHT=   ""
         Sections(0).Cells(8).Style.BorderHI=   ""
         Sections(0).Cells(8).Style.BorderHB=   ""
         Sections(0).Cells(8).Style.BorderVL=   ""
         Sections(0).Cells(8).Style.BorderVI=   ""
         Sections(0).Cells(8).Style.BorderVR=   ""
         Sections(0).Cells(8).Style.NoClipping=   0   'False
         Sections(0).Cells(8).Style.RTF=   0   'False
         Sections(0).Cells(8).Style.fprops=   16908289
         Sections(0).Cells(9).Name=   "CELL_15"
         Sections(0).Cells(9).Exp=   """Page "" & PageNo()"
         Sections(0).Cells(9).Width=   30
         Sections(0).Cells(9).PrivateStyle=   -1  'True
         Sections(0).Cells(9).Style.Name=   "<private>"
         Sections(0).Cells(9).Style.ParentName=   "Tdb_Base"
         Sections(0).Cells(9).Style.Font_Name=   "Courier"
         Sections(0).Cells(9).Style.Font_Size=   9.75
         Sections(0).Cells(9).Style.Font_Bold=   0   'False
         Sections(0).Cells(9).Style.Font_Italic=   0   'False
         Sections(0).Cells(9).Style.Font_Underline=   0   'False
         Sections(0).Cells(9).Style.Font_Strikeout=   0   'False
         Sections(0).Cells(9).Style.Font_Charset=   0
         Sections(0).Cells(9).Style.TextAlign=   2
         Sections(0).Cells(9).Style.TextVAlign=   1
         Sections(0).Cells(9).Style.TextWrap=   -1  'True
         Sections(0).Cells(9).Style.ForeColor=   0
         Sections(0).Cells(9).Style.BackColor=   16777215
         Sections(0).Cells(9).Style.NoFill=   -1  'True
         Sections(0).Cells(9).Style.BackPicFile=   ""
         Sections(0).Cells(9).Style.ForePicFile=   ""
         Sections(0).Cells(9).Style.BackPicVertPlacement=   0
         Sections(0).Cells(9).Style.BackPicHorzPlacement=   0
         Sections(0).Cells(9).Style.ForePicPlacement=   0
         Sections(0).Cells(9).Style.ForePicDrawMode=   0
         Sections(0).Cells(9).Style.MarginLeft=   6
         Sections(0).Cells(9).Style.MarginTop=   1
         Sections(0).Cells(9).Style.MarginRight=   6
         Sections(0).Cells(9).Style.MarginBottom=   1
         Sections(0).Cells(9).Style.HasBorders=   -1  'True
         Sections(0).Cells(9).Style.BorderHT=   ""
         Sections(0).Cells(9).Style.BorderHI=   ""
         Sections(0).Cells(9).Style.BorderHB=   ""
         Sections(0).Cells(9).Style.BorderVL=   ""
         Sections(0).Cells(9).Style.BorderVI=   ""
         Sections(0).Cells(9).Style.BorderVR=   ""
         Sections(0).Cells(9).Style.NoClipping=   0   'False
         Sections(0).Cells(9).Style.RTF=   0   'False
         Sections(0).Cells(9).Style.fprops=   17956865
         Sections(0).Cells(10).Name=   "CELL_4"
         Sections(0).Cells(10).Exp=   """Cust ID : "" & cKodeAnggota"
         Sections(0).Cells(10).NewLine=   -1  'True
         Sections(0).Cells(10).Height=   6
         Sections(0).Cells(10).AutoHeight=   0   'False
         Sections(0).Cells(10).PrivateStyle=   -1  'True
         Sections(0).Cells(10).Style.Name=   "<private>"
         Sections(0).Cells(10).Style.ParentName=   "Tdb_Base"
         Sections(0).Cells(10).Style.Font_Name=   "Courier"
         Sections(0).Cells(10).Style.Font_Size=   9.75
         Sections(0).Cells(10).Style.Font_Bold=   0   'False
         Sections(0).Cells(10).Style.Font_Italic=   0   'False
         Sections(0).Cells(10).Style.Font_Underline=   0   'False
         Sections(0).Cells(10).Style.Font_Strikeout=   0   'False
         Sections(0).Cells(10).Style.Font_Charset=   0
         Sections(0).Cells(10).Style.TextAlign=   3
         Sections(0).Cells(10).Style.TextVAlign=   1
         Sections(0).Cells(10).Style.TextWrap=   -1  'True
         Sections(0).Cells(10).Style.ForeColor=   0
         Sections(0).Cells(10).Style.BackColor=   16777215
         Sections(0).Cells(10).Style.NoFill=   -1  'True
         Sections(0).Cells(10).Style.BackPicFile=   ""
         Sections(0).Cells(10).Style.ForePicFile=   ""
         Sections(0).Cells(10).Style.BackPicVertPlacement=   0
         Sections(0).Cells(10).Style.BackPicHorzPlacement=   0
         Sections(0).Cells(10).Style.ForePicPlacement=   0
         Sections(0).Cells(10).Style.ForePicDrawMode=   0
         Sections(0).Cells(10).Style.MarginLeft=   6
         Sections(0).Cells(10).Style.MarginTop=   1
         Sections(0).Cells(10).Style.MarginRight=   6
         Sections(0).Cells(10).Style.MarginBottom=   1
         Sections(0).Cells(10).Style.HasBorders=   -1  'True
         Sections(0).Cells(10).Style.BorderHT=   ""
         Sections(0).Cells(10).Style.BorderHI=   ""
         Sections(0).Cells(10).Style.BorderHB=   ""
         Sections(0).Cells(10).Style.BorderVL=   ""
         Sections(0).Cells(10).Style.BorderVI=   ""
         Sections(0).Cells(10).Style.BorderVR=   ""
         Sections(0).Cells(10).Style.NoClipping=   0   'False
         Sections(0).Cells(10).Style.RTF=   0   'False
         Sections(0).Cells(10).Style.fprops=   18087936
         Sections(0).Cells(11).Name=   "CELL_17"
         Sections(0).Cells(11).Exp=   """Cust Name : "" & cNama"
         Sections(0).Cells(11).NewLine=   -1  'True
         Sections(0).Cells(11).PrivateStyle=   -1  'True
         Sections(0).Cells(11).Style.Name=   "<private>"
         Sections(0).Cells(11).Style.ParentName=   "Tdb_Base"
         Sections(0).Cells(11).Style.Font_Name=   "Courier"
         Sections(0).Cells(11).Style.Font_Size=   9.75
         Sections(0).Cells(11).Style.Font_Bold=   0   'False
         Sections(0).Cells(11).Style.Font_Italic=   0   'False
         Sections(0).Cells(11).Style.Font_Underline=   0   'False
         Sections(0).Cells(11).Style.Font_Strikeout=   0   'False
         Sections(0).Cells(11).Style.Font_Charset=   0
         Sections(0).Cells(11).Style.TextAlign=   3
         Sections(0).Cells(11).Style.TextVAlign=   1
         Sections(0).Cells(11).Style.TextWrap=   -1  'True
         Sections(0).Cells(11).Style.ForeColor=   0
         Sections(0).Cells(11).Style.BackColor=   16777215
         Sections(0).Cells(11).Style.NoFill=   -1  'True
         Sections(0).Cells(11).Style.BackPicFile=   ""
         Sections(0).Cells(11).Style.ForePicFile=   ""
         Sections(0).Cells(11).Style.BackPicVertPlacement=   0
         Sections(0).Cells(11).Style.BackPicHorzPlacement=   0
         Sections(0).Cells(11).Style.ForePicPlacement=   0
         Sections(0).Cells(11).Style.ForePicDrawMode=   0
         Sections(0).Cells(11).Style.MarginLeft=   6
         Sections(0).Cells(11).Style.MarginTop=   1
         Sections(0).Cells(11).Style.MarginRight=   6
         Sections(0).Cells(11).Style.MarginBottom=   1
         Sections(0).Cells(11).Style.HasBorders=   -1  'True
         Sections(0).Cells(11).Style.BorderHT=   ""
         Sections(0).Cells(11).Style.BorderHI=   ""
         Sections(0).Cells(11).Style.BorderHB=   ""
         Sections(0).Cells(11).Style.BorderVL=   ""
         Sections(0).Cells(11).Style.BorderVI=   ""
         Sections(0).Cells(11).Style.BorderVR=   ""
         Sections(0).Cells(11).Style.NoClipping=   0   'False
         Sections(0).Cells(11).Style.RTF=   0   'False
         Sections(0).Cells(11).Style.fprops=   18087936
         Sections(0).Cells(12).Name=   "CELL_20"
         Sections(0).Cells(12).Exp=   """Print By : ""& cUserName"
         Sections(0).Cells(12).PrivateStyle=   -1  'True
         Sections(0).Cells(12).Style.Name=   "<private>"
         Sections(0).Cells(12).Style.ParentName=   "Tdb_Base"
         Sections(0).Cells(12).Style.Font_Name=   "Courier"
         Sections(0).Cells(12).Style.Font_Size=   9.75
         Sections(0).Cells(12).Style.Font_Bold=   0   'False
         Sections(0).Cells(12).Style.Font_Italic=   0   'False
         Sections(0).Cells(12).Style.Font_Underline=   0   'False
         Sections(0).Cells(12).Style.Font_Strikeout=   0   'False
         Sections(0).Cells(12).Style.Font_Charset=   0
         Sections(0).Cells(12).Style.TextAlign=   2
         Sections(0).Cells(12).Style.TextVAlign=   1
         Sections(0).Cells(12).Style.TextWrap=   -1  'True
         Sections(0).Cells(12).Style.ForeColor=   0
         Sections(0).Cells(12).Style.BackColor=   16777215
         Sections(0).Cells(12).Style.NoFill=   -1  'True
         Sections(0).Cells(12).Style.BackPicFile=   ""
         Sections(0).Cells(12).Style.ForePicFile=   ""
         Sections(0).Cells(12).Style.BackPicVertPlacement=   0
         Sections(0).Cells(12).Style.BackPicHorzPlacement=   0
         Sections(0).Cells(12).Style.ForePicPlacement=   0
         Sections(0).Cells(12).Style.ForePicDrawMode=   0
         Sections(0).Cells(12).Style.MarginLeft=   6
         Sections(0).Cells(12).Style.MarginTop=   1
         Sections(0).Cells(12).Style.MarginRight=   6
         Sections(0).Cells(12).Style.MarginBottom=   1
         Sections(0).Cells(12).Style.HasBorders=   -1  'True
         Sections(0).Cells(12).Style.BorderHT=   ""
         Sections(0).Cells(12).Style.BorderHI=   ""
         Sections(0).Cells(12).Style.BorderHB=   ""
         Sections(0).Cells(12).Style.BorderVL=   ""
         Sections(0).Cells(12).Style.BorderVI=   ""
         Sections(0).Cells(12).Style.BorderVR=   ""
         Sections(0).Cells(12).Style.NoClipping=   0   'False
         Sections(0).Cells(12).Style.RTF=   0   'False
         Sections(0).Cells(12).Style.fprops=   16777217
         Sections(0).Cells(13).Name=   "CELL_16"
         Sections(0).Cells(13).Exp=   "cKota"
         Sections(0).Cells(13).NewLine=   -1  'True
         Sections(0).Cells(13).PrivateStyle=   -1  'True
         Sections(0).Cells(13).Style.Name=   "<private>"
         Sections(0).Cells(13).Style.ParentName=   "Tdb_Base"
         Sections(0).Cells(13).Style.Font_Name=   "Courier"
         Sections(0).Cells(13).Style.Font_Size=   9.75
         Sections(0).Cells(13).Style.Font_Bold=   0   'False
         Sections(0).Cells(13).Style.Font_Italic=   0   'False
         Sections(0).Cells(13).Style.Font_Underline=   0   'False
         Sections(0).Cells(13).Style.Font_Strikeout=   0   'False
         Sections(0).Cells(13).Style.Font_Charset=   0
         Sections(0).Cells(13).Style.TextAlign=   3
         Sections(0).Cells(13).Style.TextVAlign=   1
         Sections(0).Cells(13).Style.TextWrap=   -1  'True
         Sections(0).Cells(13).Style.ForeColor=   0
         Sections(0).Cells(13).Style.BackColor=   16777215
         Sections(0).Cells(13).Style.NoFill=   -1  'True
         Sections(0).Cells(13).Style.BackPicFile=   ""
         Sections(0).Cells(13).Style.ForePicFile=   ""
         Sections(0).Cells(13).Style.BackPicVertPlacement=   0
         Sections(0).Cells(13).Style.BackPicHorzPlacement=   0
         Sections(0).Cells(13).Style.ForePicPlacement=   0
         Sections(0).Cells(13).Style.ForePicDrawMode=   0
         Sections(0).Cells(13).Style.MarginLeft=   6
         Sections(0).Cells(13).Style.MarginTop=   1
         Sections(0).Cells(13).Style.MarginRight=   6
         Sections(0).Cells(13).Style.MarginBottom=   1
         Sections(0).Cells(13).Style.HasBorders=   -1  'True
         Sections(0).Cells(13).Style.BorderHT=   ""
         Sections(0).Cells(13).Style.BorderHI=   ""
         Sections(0).Cells(13).Style.BorderHB=   ""
         Sections(0).Cells(13).Style.BorderVL=   ""
         Sections(0).Cells(13).Style.BorderVI=   ""
         Sections(0).Cells(13).Style.BorderVR=   ""
         Sections(0).Cells(13).Style.NoClipping=   0   'False
         Sections(0).Cells(13).Style.RTF=   0   'False
         Sections(0).Cells(13).Style.fprops=   18087936
         Sections(0).Cells(14).Name=   "CELL_18"
         Sections(0).Cells(14).PrivateStyle=   -1  'True
         Sections(0).Cells(14).Style.Name=   "<private>"
         Sections(0).Cells(14).Style.ParentName=   "Tdb_Base"
         Sections(0).Cells(14).Style.Font_Name=   "Courier"
         Sections(0).Cells(14).Style.Font_Size=   9.75
         Sections(0).Cells(14).Style.Font_Bold=   0   'False
         Sections(0).Cells(14).Style.Font_Italic=   0   'False
         Sections(0).Cells(14).Style.Font_Underline=   0   'False
         Sections(0).Cells(14).Style.Font_Strikeout=   0   'False
         Sections(0).Cells(14).Style.Font_Charset=   0
         Sections(0).Cells(14).Style.TextAlign=   2
         Sections(0).Cells(14).Style.TextVAlign=   1
         Sections(0).Cells(14).Style.TextWrap=   -1  'True
         Sections(0).Cells(14).Style.ForeColor=   0
         Sections(0).Cells(14).Style.BackColor=   16777215
         Sections(0).Cells(14).Style.NoFill=   -1  'True
         Sections(0).Cells(14).Style.BackPicFile=   ""
         Sections(0).Cells(14).Style.ForePicFile=   ""
         Sections(0).Cells(14).Style.BackPicVertPlacement=   0
         Sections(0).Cells(14).Style.BackPicHorzPlacement=   0
         Sections(0).Cells(14).Style.ForePicPlacement=   0
         Sections(0).Cells(14).Style.ForePicDrawMode=   0
         Sections(0).Cells(14).Style.MarginLeft=   6
         Sections(0).Cells(14).Style.MarginTop=   1
         Sections(0).Cells(14).Style.MarginRight=   6
         Sections(0).Cells(14).Style.MarginBottom=   1
         Sections(0).Cells(14).Style.HasBorders=   -1  'True
         Sections(0).Cells(14).Style.BorderHT=   ""
         Sections(0).Cells(14).Style.BorderHI=   ""
         Sections(0).Cells(14).Style.BorderHB=   ""
         Sections(0).Cells(14).Style.BorderVL=   ""
         Sections(0).Cells(14).Style.BorderVI=   ""
         Sections(0).Cells(14).Style.BorderVR=   ""
         Sections(0).Cells(14).Style.NoClipping=   0   'False
         Sections(0).Cells(14).Style.RTF=   0   'False
         Sections(0).Cells(14).Style.fprops=   17825793
         Sections(0).Cells(15).Name=   "CELL_19"
         Sections(0).Cells(15).Exp=   "Now"
         Sections(0).Cells(15).PrivateStyle=   -1  'True
         Sections(0).Cells(15).Format=   "dd-MM-yyyy HH:MM:SS"
         Sections(0).Cells(15).Style.Name=   "<private>"
         Sections(0).Cells(15).Style.ParentName=   "Tdb_Base"
         Sections(0).Cells(15).Style.Font_Name=   "Courier"
         Sections(0).Cells(15).Style.Font_Size=   9.75
         Sections(0).Cells(15).Style.Font_Bold=   0   'False
         Sections(0).Cells(15).Style.Font_Italic=   0   'False
         Sections(0).Cells(15).Style.Font_Underline=   0   'False
         Sections(0).Cells(15).Style.Font_Strikeout=   0   'False
         Sections(0).Cells(15).Style.Font_Charset=   0
         Sections(0).Cells(15).Style.TextAlign=   2
         Sections(0).Cells(15).Style.TextVAlign=   1
         Sections(0).Cells(15).Style.TextWrap=   -1  'True
         Sections(0).Cells(15).Style.ForeColor=   0
         Sections(0).Cells(15).Style.BackColor=   16777215
         Sections(0).Cells(15).Style.NoFill=   -1  'True
         Sections(0).Cells(15).Style.BackPicFile=   ""
         Sections(0).Cells(15).Style.ForePicFile=   ""
         Sections(0).Cells(15).Style.BackPicVertPlacement=   0
         Sections(0).Cells(15).Style.BackPicHorzPlacement=   0
         Sections(0).Cells(15).Style.ForePicPlacement=   0
         Sections(0).Cells(15).Style.ForePicDrawMode=   0
         Sections(0).Cells(15).Style.MarginLeft=   6
         Sections(0).Cells(15).Style.MarginTop=   1
         Sections(0).Cells(15).Style.MarginRight=   6
         Sections(0).Cells(15).Style.MarginBottom=   1
         Sections(0).Cells(15).Style.HasBorders=   -1  'True
         Sections(0).Cells(15).Style.BorderHT=   ""
         Sections(0).Cells(15).Style.BorderHI=   ""
         Sections(0).Cells(15).Style.BorderHB=   ""
         Sections(0).Cells(15).Style.BorderVL=   ""
         Sections(0).Cells(15).Style.BorderVI=   ""
         Sections(0).Cells(15).Style.BorderVR=   ""
         Sections(0).Cells(15).Style.NoClipping=   0   'False
         Sections(0).Cells(15).Style.RTF=   0   'False
         Sections(0).Cells(15).Style.fprops=   16777217
         Sections(1).Name=   "DetailHeader"
         Sections(1).Type=   3
         Sections(1).StyleExp=   "'Tdb_Header'"
         Sections(1).Tabulator=   "Detail"
         Sections(1).Cells.Count=   6
         Sections(1).Cells(0).Name=   "CELL_0"
         Sections(1).Cells(0).Exp=   """No."""
         Sections(1).Cells(0).PrivateStyle=   -1  'True
         Sections(1).Cells(0).Style.Name=   "<private>"
         Sections(1).Cells(0).Style.ParentName=   "Tdb_Header"
         Sections(1).Cells(0).Style.Font_Name=   "Courier"
         Sections(1).Cells(0).Style.Font_Size=   9.75
         Sections(1).Cells(0).Style.Font_Bold=   -1  'True
         Sections(1).Cells(0).Style.Font_Italic=   0   'False
         Sections(1).Cells(0).Style.Font_Underline=   0   'False
         Sections(1).Cells(0).Style.Font_Strikeout=   0   'False
         Sections(1).Cells(0).Style.Font_Charset=   0
         Sections(1).Cells(0).Style.TextAlign=   1
         Sections(1).Cells(0).Style.TextVAlign=   1
         Sections(1).Cells(0).Style.TextWrap=   -1  'True
         Sections(1).Cells(0).Style.ForeColor=   0
         Sections(1).Cells(0).Style.BackColor=   16777215
         Sections(1).Cells(0).Style.NoFill=   -1  'True
         Sections(1).Cells(0).Style.BackPicFile=   ""
         Sections(1).Cells(0).Style.ForePicFile=   ""
         Sections(1).Cells(0).Style.BackPicVertPlacement=   0
         Sections(1).Cells(0).Style.BackPicHorzPlacement=   0
         Sections(1).Cells(0).Style.ForePicPlacement=   0
         Sections(1).Cells(0).Style.ForePicDrawMode=   0
         Sections(1).Cells(0).Style.MarginLeft=   6
         Sections(1).Cells(0).Style.MarginTop=   1
         Sections(1).Cells(0).Style.MarginRight=   6
         Sections(1).Cells(0).Style.MarginBottom=   1
         Sections(1).Cells(0).Style.HasBorders=   -1  'True
         Sections(1).Cells(0).Style.BorderHT=   "Single"
         Sections(1).Cells(0).Style.BorderHI=   "Single"
         Sections(1).Cells(0).Style.BorderHB=   "Single"
         Sections(1).Cells(0).Style.BorderVL=   "Single"
         Sections(1).Cells(0).Style.BorderVI=   "Single"
         Sections(1).Cells(0).Style.BorderVR=   "Single"
         Sections(1).Cells(0).Style.NoClipping=   0   'False
         Sections(1).Cells(0).Style.RTF=   0   'False
         Sections(1).Cells(0).Style.fprops=   1835009
         Sections(1).Cells(1).Name=   "CELL_2"
         Sections(1).Cells(1).Exp=   """No Invoice Penjualan"""
         Sections(1).Cells(1).PrivateStyle=   -1  'True
         Sections(1).Cells(1).Style.Name=   "<private>"
         Sections(1).Cells(1).Style.ParentName=   "Tdb_Header"
         Sections(1).Cells(1).Style.Font_Name=   "Courier"
         Sections(1).Cells(1).Style.Font_Size=   9.75
         Sections(1).Cells(1).Style.Font_Bold=   -1  'True
         Sections(1).Cells(1).Style.Font_Italic=   0   'False
         Sections(1).Cells(1).Style.Font_Underline=   0   'False
         Sections(1).Cells(1).Style.Font_Strikeout=   0   'False
         Sections(1).Cells(1).Style.Font_Charset=   0
         Sections(1).Cells(1).Style.TextAlign=   1
         Sections(1).Cells(1).Style.TextVAlign=   1
         Sections(1).Cells(1).Style.TextWrap=   -1  'True
         Sections(1).Cells(1).Style.ForeColor=   0
         Sections(1).Cells(1).Style.BackColor=   16777215
         Sections(1).Cells(1).Style.NoFill=   -1  'True
         Sections(1).Cells(1).Style.BackPicFile=   ""
         Sections(1).Cells(1).Style.ForePicFile=   ""
         Sections(1).Cells(1).Style.BackPicVertPlacement=   0
         Sections(1).Cells(1).Style.BackPicHorzPlacement=   0
         Sections(1).Cells(1).Style.ForePicPlacement=   0
         Sections(1).Cells(1).Style.ForePicDrawMode=   0
         Sections(1).Cells(1).Style.MarginLeft=   6
         Sections(1).Cells(1).Style.MarginTop=   1
         Sections(1).Cells(1).Style.MarginRight=   6
         Sections(1).Cells(1).Style.MarginBottom=   1
         Sections(1).Cells(1).Style.HasBorders=   -1  'True
         Sections(1).Cells(1).Style.BorderHT=   "Single"
         Sections(1).Cells(1).Style.BorderHI=   "Single"
         Sections(1).Cells(1).Style.BorderHB=   "Single"
         Sections(1).Cells(1).Style.BorderVL=   "Single"
         Sections(1).Cells(1).Style.BorderVI=   "Single"
         Sections(1).Cells(1).Style.BorderVR=   "Single"
         Sections(1).Cells(1).Style.NoClipping=   0   'False
         Sections(1).Cells(1).Style.RTF=   0   'False
         Sections(1).Cells(1).Style.fprops=   1835009
         Sections(1).Cells(2).Name=   "CELL_4"
         Sections(1).Cells(2).Exp=   """Total"""
         Sections(1).Cells(2).PrivateStyle=   -1  'True
         Sections(1).Cells(2).Style.Name=   "<private>"
         Sections(1).Cells(2).Style.ParentName=   "Tdb_Header"
         Sections(1).Cells(2).Style.Font_Name=   "Courier"
         Sections(1).Cells(2).Style.Font_Size=   9.75
         Sections(1).Cells(2).Style.Font_Bold=   -1  'True
         Sections(1).Cells(2).Style.Font_Italic=   0   'False
         Sections(1).Cells(2).Style.Font_Underline=   0   'False
         Sections(1).Cells(2).Style.Font_Strikeout=   0   'False
         Sections(1).Cells(2).Style.Font_Charset=   0
         Sections(1).Cells(2).Style.TextAlign=   1
         Sections(1).Cells(2).Style.TextVAlign=   1
         Sections(1).Cells(2).Style.TextWrap=   -1  'True
         Sections(1).Cells(2).Style.ForeColor=   0
         Sections(1).Cells(2).Style.BackColor=   16777215
         Sections(1).Cells(2).Style.NoFill=   -1  'True
         Sections(1).Cells(2).Style.BackPicFile=   ""
         Sections(1).Cells(2).Style.ForePicFile=   ""
         Sections(1).Cells(2).Style.BackPicVertPlacement=   0
         Sections(1).Cells(2).Style.BackPicHorzPlacement=   0
         Sections(1).Cells(2).Style.ForePicPlacement=   0
         Sections(1).Cells(2).Style.ForePicDrawMode=   0
         Sections(1).Cells(2).Style.MarginLeft=   6
         Sections(1).Cells(2).Style.MarginTop=   1
         Sections(1).Cells(2).Style.MarginRight=   6
         Sections(1).Cells(2).Style.MarginBottom=   1
         Sections(1).Cells(2).Style.HasBorders=   -1  'True
         Sections(1).Cells(2).Style.BorderHT=   "Single"
         Sections(1).Cells(2).Style.BorderHI=   "Single"
         Sections(1).Cells(2).Style.BorderHB=   "Single"
         Sections(1).Cells(2).Style.BorderVL=   "Single"
         Sections(1).Cells(2).Style.BorderVI=   "Single"
         Sections(1).Cells(2).Style.BorderVR=   "Single"
         Sections(1).Cells(2).Style.NoClipping=   0   'False
         Sections(1).Cells(2).Style.RTF=   0   'False
         Sections(1).Cells(2).Style.fprops=   1835009
         Sections(1).Cells(3).Name=   "CELL_1"
         Sections(1).Cells(3).Exp=   """Jatuh Tempo"""
         Sections(1).Cells(3).PrivateStyle=   -1  'True
         Sections(1).Cells(3).Style.Name=   "<private>"
         Sections(1).Cells(3).Style.ParentName=   "Tdb_Header"
         Sections(1).Cells(3).Style.Font_Name=   "Courier"
         Sections(1).Cells(3).Style.Font_Size=   9.75
         Sections(1).Cells(3).Style.Font_Bold=   -1  'True
         Sections(1).Cells(3).Style.Font_Italic=   0   'False
         Sections(1).Cells(3).Style.Font_Underline=   0   'False
         Sections(1).Cells(3).Style.Font_Strikeout=   0   'False
         Sections(1).Cells(3).Style.Font_Charset=   0
         Sections(1).Cells(3).Style.TextAlign=   1
         Sections(1).Cells(3).Style.TextVAlign=   1
         Sections(1).Cells(3).Style.TextWrap=   -1  'True
         Sections(1).Cells(3).Style.ForeColor=   0
         Sections(1).Cells(3).Style.BackColor=   16777215
         Sections(1).Cells(3).Style.NoFill=   -1  'True
         Sections(1).Cells(3).Style.BackPicFile=   ""
         Sections(1).Cells(3).Style.ForePicFile=   ""
         Sections(1).Cells(3).Style.BackPicVertPlacement=   0
         Sections(1).Cells(3).Style.BackPicHorzPlacement=   0
         Sections(1).Cells(3).Style.ForePicPlacement=   0
         Sections(1).Cells(3).Style.ForePicDrawMode=   0
         Sections(1).Cells(3).Style.MarginLeft=   6
         Sections(1).Cells(3).Style.MarginTop=   1
         Sections(1).Cells(3).Style.MarginRight=   6
         Sections(1).Cells(3).Style.MarginBottom=   1
         Sections(1).Cells(3).Style.HasBorders=   -1  'True
         Sections(1).Cells(3).Style.BorderHT=   "Single"
         Sections(1).Cells(3).Style.BorderHI=   "Single"
         Sections(1).Cells(3).Style.BorderHB=   "Single"
         Sections(1).Cells(3).Style.BorderVL=   "Single"
         Sections(1).Cells(3).Style.BorderVI=   "Single"
         Sections(1).Cells(3).Style.BorderVR=   "Single"
         Sections(1).Cells(3).Style.NoClipping=   0   'False
         Sections(1).Cells(3).Style.RTF=   0   'False
         Sections(1).Cells(3).Style.fprops=   1835009
         Sections(1).Cells(4).Name=   "CELL_7"
         Sections(1).Cells(4).Exp=   """Denda"""
         Sections(1).Cells(4).PrivateStyle=   -1  'True
         Sections(1).Cells(4).Style.Name=   "<private>"
         Sections(1).Cells(4).Style.ParentName=   "Tdb_Header"
         Sections(1).Cells(4).Style.Font_Name=   "Courier"
         Sections(1).Cells(4).Style.Font_Size=   9.75
         Sections(1).Cells(4).Style.Font_Bold=   -1  'True
         Sections(1).Cells(4).Style.Font_Italic=   0   'False
         Sections(1).Cells(4).Style.Font_Underline=   0   'False
         Sections(1).Cells(4).Style.Font_Strikeout=   0   'False
         Sections(1).Cells(4).Style.Font_Charset=   0
         Sections(1).Cells(4).Style.TextAlign=   1
         Sections(1).Cells(4).Style.TextVAlign=   1
         Sections(1).Cells(4).Style.TextWrap=   -1  'True
         Sections(1).Cells(4).Style.ForeColor=   0
         Sections(1).Cells(4).Style.BackColor=   16777215
         Sections(1).Cells(4).Style.NoFill=   -1  'True
         Sections(1).Cells(4).Style.BackPicFile=   ""
         Sections(1).Cells(4).Style.ForePicFile=   ""
         Sections(1).Cells(4).Style.BackPicVertPlacement=   0
         Sections(1).Cells(4).Style.BackPicHorzPlacement=   0
         Sections(1).Cells(4).Style.ForePicPlacement=   0
         Sections(1).Cells(4).Style.ForePicDrawMode=   0
         Sections(1).Cells(4).Style.MarginLeft=   6
         Sections(1).Cells(4).Style.MarginTop=   1
         Sections(1).Cells(4).Style.MarginRight=   6
         Sections(1).Cells(4).Style.MarginBottom=   1
         Sections(1).Cells(4).Style.HasBorders=   -1  'True
         Sections(1).Cells(4).Style.BorderHT=   "Single"
         Sections(1).Cells(4).Style.BorderHI=   "Single"
         Sections(1).Cells(4).Style.BorderHB=   "Single"
         Sections(1).Cells(4).Style.BorderVL=   "Single"
         Sections(1).Cells(4).Style.BorderVI=   "Single"
         Sections(1).Cells(4).Style.BorderVR=   "Single"
         Sections(1).Cells(4).Style.NoClipping=   0   'False
         Sections(1).Cells(4).Style.RTF=   0   'False
         Sections(1).Cells(4).Style.fprops=   1835009
         Sections(1).Cells(5).Name=   "CELL_6"
         Sections(1).Cells(5).Exp=   """Total"""
         Sections(1).Cells(5).PrivateStyle=   -1  'True
         Sections(1).Cells(5).Style.Name=   "<private>"
         Sections(1).Cells(5).Style.ParentName=   "Tdb_Header"
         Sections(1).Cells(5).Style.Font_Name=   "Courier"
         Sections(1).Cells(5).Style.Font_Size=   9.75
         Sections(1).Cells(5).Style.Font_Bold=   -1  'True
         Sections(1).Cells(5).Style.Font_Italic=   0   'False
         Sections(1).Cells(5).Style.Font_Underline=   0   'False
         Sections(1).Cells(5).Style.Font_Strikeout=   0   'False
         Sections(1).Cells(5).Style.Font_Charset=   0
         Sections(1).Cells(5).Style.TextAlign=   1
         Sections(1).Cells(5).Style.TextVAlign=   1
         Sections(1).Cells(5).Style.TextWrap=   -1  'True
         Sections(1).Cells(5).Style.ForeColor=   0
         Sections(1).Cells(5).Style.BackColor=   16777215
         Sections(1).Cells(5).Style.NoFill=   -1  'True
         Sections(1).Cells(5).Style.BackPicFile=   ""
         Sections(1).Cells(5).Style.ForePicFile=   ""
         Sections(1).Cells(5).Style.BackPicVertPlacement=   0
         Sections(1).Cells(5).Style.BackPicHorzPlacement=   0
         Sections(1).Cells(5).Style.ForePicPlacement=   0
         Sections(1).Cells(5).Style.ForePicDrawMode=   0
         Sections(1).Cells(5).Style.MarginLeft=   6
         Sections(1).Cells(5).Style.MarginTop=   1
         Sections(1).Cells(5).Style.MarginRight=   6
         Sections(1).Cells(5).Style.MarginBottom=   1
         Sections(1).Cells(5).Style.HasBorders=   -1  'True
         Sections(1).Cells(5).Style.BorderHT=   "Single"
         Sections(1).Cells(5).Style.BorderHI=   "Single"
         Sections(1).Cells(5).Style.BorderHB=   "Single"
         Sections(1).Cells(5).Style.BorderVL=   "Single"
         Sections(1).Cells(5).Style.BorderVI=   "Single"
         Sections(1).Cells(5).Style.BorderVR=   "Single"
         Sections(1).Cells(5).Style.NoClipping=   0   'False
         Sections(1).Cells(5).Style.RTF=   0   'False
         Sections(1).Cells(5).Style.fprops=   1835009
         Sections(2).Name=   "Detail"
         Sections(2).Type=   4
         Sections(2).StyleExp=   "'Tdb_Body'"
         Sections(2).AutoHeight=   0   'False
         Sections(2).Height=   5
         Sections(2).Cells.Count=   6
         Sections(2).Cells(0).Name=   "CELL_0"
         Sections(2).Cells(0).Exp=   "Nomor"
         Sections(2).Cells(0).Width=   4
         Sections(2).Cells(0).PrivateStyle=   -1  'True
         Sections(2).Cells(0).Style.Name=   "<private>"
         Sections(2).Cells(0).Style.ParentName=   "Tdb_Body"
         Sections(2).Cells(0).Style.Font_Name=   "Courier"
         Sections(2).Cells(0).Style.Font_Size=   9.75
         Sections(2).Cells(0).Style.Font_Bold=   0   'False
         Sections(2).Cells(0).Style.Font_Italic=   0   'False
         Sections(2).Cells(0).Style.Font_Underline=   0   'False
         Sections(2).Cells(0).Style.Font_Strikeout=   0   'False
         Sections(2).Cells(0).Style.Font_Charset=   0
         Sections(2).Cells(0).Style.TextAlign=   0
         Sections(2).Cells(0).Style.TextVAlign=   1
         Sections(2).Cells(0).Style.TextWrap=   0   'False
         Sections(2).Cells(0).Style.ForeColor=   0
         Sections(2).Cells(0).Style.BackColor=   16777215
         Sections(2).Cells(0).Style.NoFill=   -1  'True
         Sections(2).Cells(0).Style.BackPicFile=   ""
         Sections(2).Cells(0).Style.ForePicFile=   ""
         Sections(2).Cells(0).Style.BackPicVertPlacement=   0
         Sections(2).Cells(0).Style.BackPicHorzPlacement=   0
         Sections(2).Cells(0).Style.ForePicPlacement=   0
         Sections(2).Cells(0).Style.ForePicDrawMode=   0
         Sections(2).Cells(0).Style.MarginLeft=   6
         Sections(2).Cells(0).Style.MarginTop=   0
         Sections(2).Cells(0).Style.MarginRight=   6
         Sections(2).Cells(0).Style.MarginBottom=   0
         Sections(2).Cells(0).Style.HasBorders=   -1  'True
         Sections(2).Cells(0).Style.BorderHT=   ""
         Sections(2).Cells(0).Style.BorderHI=   ""
         Sections(2).Cells(0).Style.BorderHB=   ""
         Sections(2).Cells(0).Style.BorderVL=   "Single"
         Sections(2).Cells(0).Style.BorderVI=   "Single"
         Sections(2).Cells(0).Style.BorderVR=   "Single"
         Sections(2).Cells(0).Style.NoClipping=   0   'False
         Sections(2).Cells(0).Style.RTF=   0   'False
         Sections(2).Cells(0).Style.fprops=   1855493
         Sections(2).Cells(1).Name=   "CELL_2"
         Sections(2).Cells(1).Exp=   "NoInvoice"
         Sections(2).Cells(1).PrivateStyle=   -1  'True
         Sections(2).Cells(1).Style.Name=   "<private>"
         Sections(2).Cells(1).Style.ParentName=   "Tdb_Body"
         Sections(2).Cells(1).Style.Font_Name=   "Courier"
         Sections(2).Cells(1).Style.Font_Size=   9.75
         Sections(2).Cells(1).Style.Font_Bold=   0   'False
         Sections(2).Cells(1).Style.Font_Italic=   0   'False
         Sections(2).Cells(1).Style.Font_Underline=   0   'False
         Sections(2).Cells(1).Style.Font_Strikeout=   0   'False
         Sections(2).Cells(1).Style.Font_Charset=   0
         Sections(2).Cells(1).Style.TextAlign=   3
         Sections(2).Cells(1).Style.TextVAlign=   1
         Sections(2).Cells(1).Style.TextWrap=   0   'False
         Sections(2).Cells(1).Style.ForeColor=   0
         Sections(2).Cells(1).Style.BackColor=   16777215
         Sections(2).Cells(1).Style.NoFill=   -1  'True
         Sections(2).Cells(1).Style.BackPicFile=   ""
         Sections(2).Cells(1).Style.ForePicFile=   ""
         Sections(2).Cells(1).Style.BackPicVertPlacement=   0
         Sections(2).Cells(1).Style.BackPicHorzPlacement=   0
         Sections(2).Cells(1).Style.ForePicPlacement=   0
         Sections(2).Cells(1).Style.ForePicDrawMode=   0
         Sections(2).Cells(1).Style.MarginLeft=   6
         Sections(2).Cells(1).Style.MarginTop=   0
         Sections(2).Cells(1).Style.MarginRight=   6
         Sections(2).Cells(1).Style.MarginBottom=   0
         Sections(2).Cells(1).Style.HasBorders=   -1  'True
         Sections(2).Cells(1).Style.BorderHT=   ""
         Sections(2).Cells(1).Style.BorderHI=   ""
         Sections(2).Cells(1).Style.BorderHB=   ""
         Sections(2).Cells(1).Style.BorderVL=   "Single"
         Sections(2).Cells(1).Style.BorderVI=   "Single"
         Sections(2).Cells(1).Style.BorderVR=   "Single"
         Sections(2).Cells(1).Style.NoClipping=   0   'False
         Sections(2).Cells(1).Style.RTF=   0   'False
         Sections(2).Cells(1).Style.fprops=   1835012
         Sections(2).Cells(2).Name=   "CELL_3"
         Sections(2).Cells(2).Exp=   "Total"
         Sections(2).Cells(2).Width=   13
         Sections(2).Cells(2).PrivateStyle=   -1  'True
         Sections(2).Cells(2).Format=   "###,###,###,###,###,##0.00"
         Sections(2).Cells(2).Style.Name=   "<private>"
         Sections(2).Cells(2).Style.ParentName=   "Tdb_Body"
         Sections(2).Cells(2).Style.Font_Name=   "Courier"
         Sections(2).Cells(2).Style.Font_Size=   9.75
         Sections(2).Cells(2).Style.Font_Bold=   0   'False
         Sections(2).Cells(2).Style.Font_Italic=   0   'False
         Sections(2).Cells(2).Style.Font_Underline=   0   'False
         Sections(2).Cells(2).Style.Font_Strikeout=   0   'False
         Sections(2).Cells(2).Style.Font_Charset=   0
         Sections(2).Cells(2).Style.TextAlign=   2
         Sections(2).Cells(2).Style.TextVAlign=   1
         Sections(2).Cells(2).Style.TextWrap=   0   'False
         Sections(2).Cells(2).Style.ForeColor=   0
         Sections(2).Cells(2).Style.BackColor=   16777215
         Sections(2).Cells(2).Style.NoFill=   -1  'True
         Sections(2).Cells(2).Style.BackPicFile=   ""
         Sections(2).Cells(2).Style.ForePicFile=   ""
         Sections(2).Cells(2).Style.BackPicVertPlacement=   0
         Sections(2).Cells(2).Style.BackPicHorzPlacement=   0
         Sections(2).Cells(2).Style.ForePicPlacement=   0
         Sections(2).Cells(2).Style.ForePicDrawMode=   0
         Sections(2).Cells(2).Style.MarginLeft=   6
         Sections(2).Cells(2).Style.MarginTop=   0
         Sections(2).Cells(2).Style.MarginRight=   6
         Sections(2).Cells(2).Style.MarginBottom=   0
         Sections(2).Cells(2).Style.HasBorders=   -1  'True
         Sections(2).Cells(2).Style.BorderHT=   ""
         Sections(2).Cells(2).Style.BorderHI=   ""
         Sections(2).Cells(2).Style.BorderHB=   ""
         Sections(2).Cells(2).Style.BorderVL=   "Single"
         Sections(2).Cells(2).Style.BorderVI=   "Single"
         Sections(2).Cells(2).Style.BorderVR=   "Single"
         Sections(2).Cells(2).Style.NoClipping=   0   'False
         Sections(2).Cells(2).Style.RTF=   0   'False
         Sections(2).Cells(2).Style.fprops=   1835013
         Sections(2).Cells(3).Name=   "CELL_4"
         Sections(2).Cells(3).Exp=   "JatuhTempo"
         Sections(2).Cells(3).Width=   10
         Sections(2).Cells(3).PrivateStyle=   -1  'True
         Sections(2).Cells(3).Style.Name=   "<private>"
         Sections(2).Cells(3).Style.ParentName=   "Tdb_Body"
         Sections(2).Cells(3).Style.Font_Name=   "Courier"
         Sections(2).Cells(3).Style.Font_Size=   9.75
         Sections(2).Cells(3).Style.Font_Bold=   0   'False
         Sections(2).Cells(3).Style.Font_Italic=   0   'False
         Sections(2).Cells(3).Style.Font_Underline=   0   'False
         Sections(2).Cells(3).Style.Font_Strikeout=   0   'False
         Sections(2).Cells(3).Style.Font_Charset=   0
         Sections(2).Cells(3).Style.TextAlign=   2
         Sections(2).Cells(3).Style.TextVAlign=   1
         Sections(2).Cells(3).Style.TextWrap=   0   'False
         Sections(2).Cells(3).Style.ForeColor=   0
         Sections(2).Cells(3).Style.BackColor=   16777215
         Sections(2).Cells(3).Style.NoFill=   -1  'True
         Sections(2).Cells(3).Style.BackPicFile=   ""
         Sections(2).Cells(3).Style.ForePicFile=   ""
         Sections(2).Cells(3).Style.BackPicVertPlacement=   0
         Sections(2).Cells(3).Style.BackPicHorzPlacement=   0
         Sections(2).Cells(3).Style.ForePicPlacement=   0
         Sections(2).Cells(3).Style.ForePicDrawMode=   0
         Sections(2).Cells(3).Style.MarginLeft=   6
         Sections(2).Cells(3).Style.MarginTop=   0
         Sections(2).Cells(3).Style.MarginRight=   6
         Sections(2).Cells(3).Style.MarginBottom=   0
         Sections(2).Cells(3).Style.HasBorders=   -1  'True
         Sections(2).Cells(3).Style.BorderHT=   ""
         Sections(2).Cells(3).Style.BorderHI=   ""
         Sections(2).Cells(3).Style.BorderHB=   ""
         Sections(2).Cells(3).Style.BorderVL=   "Single"
         Sections(2).Cells(3).Style.BorderVI=   "Single"
         Sections(2).Cells(3).Style.BorderVR=   "Single"
         Sections(2).Cells(3).Style.NoClipping=   0   'False
         Sections(2).Cells(3).Style.RTF=   0   'False
         Sections(2).Cells(3).Style.fprops=   1835013
         Sections(2).Cells(4).Name=   "CELL_1"
         Sections(2).Cells(4).Exp=   "Denda"
         Sections(2).Cells(4).Width=   11
         Sections(2).Cells(4).PrivateStyle=   -1  'True
         Sections(2).Cells(4).Format=   "###,###,###,###,###,##0.00"
         Sections(2).Cells(4).Style.Name=   "<private>"
         Sections(2).Cells(4).Style.ParentName=   "Tdb_Body"
         Sections(2).Cells(4).Style.Font_Name=   "Courier"
         Sections(2).Cells(4).Style.Font_Size=   9.75
         Sections(2).Cells(4).Style.Font_Bold=   0   'False
         Sections(2).Cells(4).Style.Font_Italic=   0   'False
         Sections(2).Cells(4).Style.Font_Underline=   0   'False
         Sections(2).Cells(4).Style.Font_Strikeout=   0   'False
         Sections(2).Cells(4).Style.Font_Charset=   0
         Sections(2).Cells(4).Style.TextAlign=   1
         Sections(2).Cells(4).Style.TextVAlign=   1
         Sections(2).Cells(4).Style.TextWrap=   0   'False
         Sections(2).Cells(4).Style.ForeColor=   0
         Sections(2).Cells(4).Style.BackColor=   16777215
         Sections(2).Cells(4).Style.NoFill=   -1  'True
         Sections(2).Cells(4).Style.BackPicFile=   ""
         Sections(2).Cells(4).Style.ForePicFile=   ""
         Sections(2).Cells(4).Style.BackPicVertPlacement=   0
         Sections(2).Cells(4).Style.BackPicHorzPlacement=   0
         Sections(2).Cells(4).Style.ForePicPlacement=   0
         Sections(2).Cells(4).Style.ForePicDrawMode=   0
         Sections(2).Cells(4).Style.MarginLeft=   6
         Sections(2).Cells(4).Style.MarginTop=   0
         Sections(2).Cells(4).Style.MarginRight=   6
         Sections(2).Cells(4).Style.MarginBottom=   0
         Sections(2).Cells(4).Style.HasBorders=   -1  'True
         Sections(2).Cells(4).Style.BorderHT=   ""
         Sections(2).Cells(4).Style.BorderHI=   ""
         Sections(2).Cells(4).Style.BorderHB=   ""
         Sections(2).Cells(4).Style.BorderVL=   "Single"
         Sections(2).Cells(4).Style.BorderVI=   "Single"
         Sections(2).Cells(4).Style.BorderVR=   "Single"
         Sections(2).Cells(4).Style.NoClipping=   0   'False
         Sections(2).Cells(4).Style.RTF=   0   'False
         Sections(2).Cells(4).Style.fprops=   1835013
         Sections(2).Cells(5).Name=   "CELL_6"
         Sections(2).Cells(5).Exp=   "Jumlah"
         Sections(2).Cells(5).Width=   13
         Sections(2).Cells(5).PrivateStyle=   -1  'True
         Sections(2).Cells(5).Format=   "###,###,###,###,###,##0.00"
         Sections(2).Cells(5).Style.Name=   "<private>"
         Sections(2).Cells(5).Style.ParentName=   "Tdb_Body"
         Sections(2).Cells(5).Style.Font_Name=   "Courier"
         Sections(2).Cells(5).Style.Font_Size=   9.75
         Sections(2).Cells(5).Style.Font_Bold=   0   'False
         Sections(2).Cells(5).Style.Font_Italic=   0   'False
         Sections(2).Cells(5).Style.Font_Underline=   0   'False
         Sections(2).Cells(5).Style.Font_Strikeout=   0   'False
         Sections(2).Cells(5).Style.Font_Charset=   0
         Sections(2).Cells(5).Style.TextAlign=   2
         Sections(2).Cells(5).Style.TextVAlign=   1
         Sections(2).Cells(5).Style.TextWrap=   0   'False
         Sections(2).Cells(5).Style.ForeColor=   0
         Sections(2).Cells(5).Style.BackColor=   16777215
         Sections(2).Cells(5).Style.NoFill=   -1  'True
         Sections(2).Cells(5).Style.BackPicFile=   ""
         Sections(2).Cells(5).Style.ForePicFile=   ""
         Sections(2).Cells(5).Style.BackPicVertPlacement=   0
         Sections(2).Cells(5).Style.BackPicHorzPlacement=   0
         Sections(2).Cells(5).Style.ForePicPlacement=   0
         Sections(2).Cells(5).Style.ForePicDrawMode=   0
         Sections(2).Cells(5).Style.MarginLeft=   6
         Sections(2).Cells(5).Style.MarginTop=   0
         Sections(2).Cells(5).Style.MarginRight=   6
         Sections(2).Cells(5).Style.MarginBottom=   0
         Sections(2).Cells(5).Style.HasBorders=   -1  'True
         Sections(2).Cells(5).Style.BorderHT=   ""
         Sections(2).Cells(5).Style.BorderHI=   ""
         Sections(2).Cells(5).Style.BorderHB=   ""
         Sections(2).Cells(5).Style.BorderVL=   "Single"
         Sections(2).Cells(5).Style.BorderVI=   "Single"
         Sections(2).Cells(5).Style.BorderVR=   "Single"
         Sections(2).Cells(5).Style.NoClipping=   0   'False
         Sections(2).Cells(5).Style.RTF=   0   'False
         Sections(2).Cells(5).Style.fprops=   1835013
         Sections(3).Name=   "SECTION_7"
         Sections(3).Type=   5
         Sections(3).Condition=   "IsLastRec()"
         Sections(3).Cells.Count=   4
         Sections(3).Cells(0).Name=   "CELL_0"
         Sections(3).Cells(0).PrivateStyle=   -1  'True
         Sections(3).Cells(0).Style.Name=   "<private>"
         Sections(3).Cells(0).Style.ParentName=   "<null>"
         Sections(3).Cells(0).Style.Font_Name=   "Times New Roman"
         Sections(3).Cells(0).Style.Font_Size=   10
         Sections(3).Cells(0).Style.Font_Bold=   0   'False
         Sections(3).Cells(0).Style.Font_Italic=   0   'False
         Sections(3).Cells(0).Style.Font_Underline=   0   'False
         Sections(3).Cells(0).Style.Font_Strikeout=   0   'False
         Sections(3).Cells(0).Style.Font_Charset=   1
         Sections(3).Cells(0).Style.TextAlign=   3
         Sections(3).Cells(0).Style.TextVAlign=   0
         Sections(3).Cells(0).Style.TextWrap=   -1  'True
         Sections(3).Cells(0).Style.ForeColor=   0
         Sections(3).Cells(0).Style.BackColor=   16777215
         Sections(3).Cells(0).Style.NoFill=   -1  'True
         Sections(3).Cells(0).Style.BackPicFile=   ""
         Sections(3).Cells(0).Style.ForePicFile=   ""
         Sections(3).Cells(0).Style.BackPicVertPlacement=   0
         Sections(3).Cells(0).Style.BackPicHorzPlacement=   0
         Sections(3).Cells(0).Style.ForePicPlacement=   0
         Sections(3).Cells(0).Style.ForePicDrawMode=   0
         Sections(3).Cells(0).Style.MarginLeft=   6
         Sections(3).Cells(0).Style.MarginTop=   6
         Sections(3).Cells(0).Style.MarginRight=   6
         Sections(3).Cells(0).Style.MarginBottom=   6
         Sections(3).Cells(0).Style.HasBorders=   -1  'True
         Sections(3).Cells(0).Style.BorderHT=   ""
         Sections(3).Cells(0).Style.BorderHI=   ""
         Sections(3).Cells(0).Style.BorderHB=   "Single"
         Sections(3).Cells(0).Style.BorderVL=   "Single"
         Sections(3).Cells(0).Style.BorderVI=   ""
         Sections(3).Cells(0).Style.BorderVR=   ""
         Sections(3).Cells(0).Style.NoClipping=   0   'False
         Sections(3).Cells(0).Style.RTF=   0   'False
         Sections(3).Cells(0).Style.fprops=   1441792
         Sections(3).Cells(1).Name=   "CELL_1"
         Sections(3).Cells(1).PrivateStyle=   -1  'True
         Sections(3).Cells(1).Style.Name=   "<private>"
         Sections(3).Cells(1).Style.ParentName=   "<null>"
         Sections(3).Cells(1).Style.Font_Name=   "Times New Roman"
         Sections(3).Cells(1).Style.Font_Size=   10
         Sections(3).Cells(1).Style.Font_Bold=   0   'False
         Sections(3).Cells(1).Style.Font_Italic=   0   'False
         Sections(3).Cells(1).Style.Font_Underline=   0   'False
         Sections(3).Cells(1).Style.Font_Strikeout=   0   'False
         Sections(3).Cells(1).Style.Font_Charset=   1
         Sections(3).Cells(1).Style.TextAlign=   3
         Sections(3).Cells(1).Style.TextVAlign=   0
         Sections(3).Cells(1).Style.TextWrap=   -1  'True
         Sections(3).Cells(1).Style.ForeColor=   0
         Sections(3).Cells(1).Style.BackColor=   16777215
         Sections(3).Cells(1).Style.NoFill=   -1  'True
         Sections(3).Cells(1).Style.BackPicFile=   ""
         Sections(3).Cells(1).Style.ForePicFile=   ""
         Sections(3).Cells(1).Style.BackPicVertPlacement=   0
         Sections(3).Cells(1).Style.BackPicHorzPlacement=   0
         Sections(3).Cells(1).Style.ForePicPlacement=   0
         Sections(3).Cells(1).Style.ForePicDrawMode=   0
         Sections(3).Cells(1).Style.MarginLeft=   6
         Sections(3).Cells(1).Style.MarginTop=   6
         Sections(3).Cells(1).Style.MarginRight=   6
         Sections(3).Cells(1).Style.MarginBottom=   6
         Sections(3).Cells(1).Style.HasBorders=   -1  'True
         Sections(3).Cells(1).Style.BorderHT=   ""
         Sections(3).Cells(1).Style.BorderHI=   ""
         Sections(3).Cells(1).Style.BorderHB=   "Single"
         Sections(3).Cells(1).Style.BorderVL=   ""
         Sections(3).Cells(1).Style.BorderVI=   ""
         Sections(3).Cells(1).Style.BorderVR=   ""
         Sections(3).Cells(1).Style.NoClipping=   0   'False
         Sections(3).Cells(1).Style.RTF=   0   'False
         Sections(3).Cells(1).Style.fprops=   1441792
         Sections(3).Cells(2).Name=   "CELL_2"
         Sections(3).Cells(2).PrivateStyle=   -1  'True
         Sections(3).Cells(2).Style.Name=   "<private>"
         Sections(3).Cells(2).Style.ParentName=   "<null>"
         Sections(3).Cells(2).Style.Font_Name=   "Times New Roman"
         Sections(3).Cells(2).Style.Font_Size=   10
         Sections(3).Cells(2).Style.Font_Bold=   0   'False
         Sections(3).Cells(2).Style.Font_Italic=   0   'False
         Sections(3).Cells(2).Style.Font_Underline=   0   'False
         Sections(3).Cells(2).Style.Font_Strikeout=   0   'False
         Sections(3).Cells(2).Style.Font_Charset=   1
         Sections(3).Cells(2).Style.TextAlign=   3
         Sections(3).Cells(2).Style.TextVAlign=   0
         Sections(3).Cells(2).Style.TextWrap=   -1  'True
         Sections(3).Cells(2).Style.ForeColor=   0
         Sections(3).Cells(2).Style.BackColor=   16777215
         Sections(3).Cells(2).Style.NoFill=   -1  'True
         Sections(3).Cells(2).Style.BackPicFile=   ""
         Sections(3).Cells(2).Style.ForePicFile=   ""
         Sections(3).Cells(2).Style.BackPicVertPlacement=   0
         Sections(3).Cells(2).Style.BackPicHorzPlacement=   0
         Sections(3).Cells(2).Style.ForePicPlacement=   0
         Sections(3).Cells(2).Style.ForePicDrawMode=   0
         Sections(3).Cells(2).Style.MarginLeft=   6
         Sections(3).Cells(2).Style.MarginTop=   6
         Sections(3).Cells(2).Style.MarginRight=   6
         Sections(3).Cells(2).Style.MarginBottom=   6
         Sections(3).Cells(2).Style.HasBorders=   -1  'True
         Sections(3).Cells(2).Style.BorderHT=   ""
         Sections(3).Cells(2).Style.BorderHI=   ""
         Sections(3).Cells(2).Style.BorderHB=   "Single"
         Sections(3).Cells(2).Style.BorderVL=   ""
         Sections(3).Cells(2).Style.BorderVI=   ""
         Sections(3).Cells(2).Style.BorderVR=   ""
         Sections(3).Cells(2).Style.NoClipping=   0   'False
         Sections(3).Cells(2).Style.RTF=   0   'False
         Sections(3).Cells(2).Style.fprops=   1441792
         Sections(3).Cells(3).Name=   "CELL_3"
         Sections(3).Cells(3).PrivateStyle=   -1  'True
         Sections(3).Cells(3).Style.Name=   "<private>"
         Sections(3).Cells(3).Style.ParentName=   "<null>"
         Sections(3).Cells(3).Style.Font_Name=   "Times New Roman"
         Sections(3).Cells(3).Style.Font_Size=   10
         Sections(3).Cells(3).Style.Font_Bold=   0   'False
         Sections(3).Cells(3).Style.Font_Italic=   0   'False
         Sections(3).Cells(3).Style.Font_Underline=   0   'False
         Sections(3).Cells(3).Style.Font_Strikeout=   0   'False
         Sections(3).Cells(3).Style.Font_Charset=   1
         Sections(3).Cells(3).Style.TextAlign=   3
         Sections(3).Cells(3).Style.TextVAlign=   0
         Sections(3).Cells(3).Style.TextWrap=   -1  'True
         Sections(3).Cells(3).Style.ForeColor=   0
         Sections(3).Cells(3).Style.BackColor=   16777215
         Sections(3).Cells(3).Style.NoFill=   -1  'True
         Sections(3).Cells(3).Style.BackPicFile=   ""
         Sections(3).Cells(3).Style.ForePicFile=   ""
         Sections(3).Cells(3).Style.BackPicVertPlacement=   0
         Sections(3).Cells(3).Style.BackPicHorzPlacement=   0
         Sections(3).Cells(3).Style.ForePicPlacement=   0
         Sections(3).Cells(3).Style.ForePicDrawMode=   0
         Sections(3).Cells(3).Style.MarginLeft=   6
         Sections(3).Cells(3).Style.MarginTop=   6
         Sections(3).Cells(3).Style.MarginRight=   6
         Sections(3).Cells(3).Style.MarginBottom=   6
         Sections(3).Cells(3).Style.HasBorders=   -1  'True
         Sections(3).Cells(3).Style.BorderHT=   ""
         Sections(3).Cells(3).Style.BorderHI=   ""
         Sections(3).Cells(3).Style.BorderHB=   "Single"
         Sections(3).Cells(3).Style.BorderVL=   ""
         Sections(3).Cells(3).Style.BorderVI=   ""
         Sections(3).Cells(3).Style.BorderVR=   "Single"
         Sections(3).Cells(3).Style.NoClipping=   0   'False
         Sections(3).Cells(3).Style.RTF=   0   'False
         Sections(3).Cells(3).Style.fprops=   1441792
         Sections(4).Name=   "SECTION_6"
         Sections(4).Type=   5
         Sections(4).Condition=   "IsLastRec()=false"
         Sections(4).StyleExp=   "'STYLE_1'"
         Sections(4).AutoHeight=   0   'False
         Sections(4).Height=   5
         Sections(4).Cells.Count=   1
         Sections(4).Cells(0).Name=   "CELL_1"
         Sections(4).Cells(0).Exp=   "IIF(IsLastRec(),"""",""Continued to page.."" & PageNo()+1)"
         Sections(4).Cells(0).NewLine=   -1  'True
         Sections(4).Cells(0).PrivateStyle=   -1  'True
         Sections(4).Cells(0).Style.Name=   "<private>"
         Sections(4).Cells(0).Style.ParentName=   "STYLE_1"
         Sections(4).Cells(0).Style.Font_Name=   "Courier"
         Sections(4).Cells(0).Style.Font_Size=   9.75
         Sections(4).Cells(0).Style.Font_Bold=   0   'False
         Sections(4).Cells(0).Style.Font_Italic=   0   'False
         Sections(4).Cells(0).Style.Font_Underline=   0   'False
         Sections(4).Cells(0).Style.Font_Strikeout=   0   'False
         Sections(4).Cells(0).Style.Font_Charset=   0
         Sections(4).Cells(0).Style.TextAlign=   3
         Sections(4).Cells(0).Style.TextVAlign=   1
         Sections(4).Cells(0).Style.TextWrap=   -1  'True
         Sections(4).Cells(0).Style.ForeColor=   0
         Sections(4).Cells(0).Style.BackColor=   16777215
         Sections(4).Cells(0).Style.NoFill=   -1  'True
         Sections(4).Cells(0).Style.BackPicFile=   ""
         Sections(4).Cells(0).Style.ForePicFile=   ""
         Sections(4).Cells(0).Style.BackPicVertPlacement=   0
         Sections(4).Cells(0).Style.BackPicHorzPlacement=   0
         Sections(4).Cells(0).Style.ForePicPlacement=   0
         Sections(4).Cells(0).Style.ForePicDrawMode=   0
         Sections(4).Cells(0).Style.MarginLeft=   6
         Sections(4).Cells(0).Style.MarginTop=   1
         Sections(4).Cells(0).Style.MarginRight=   6
         Sections(4).Cells(0).Style.MarginBottom=   1
         Sections(4).Cells(0).Style.HasBorders=   -1  'True
         Sections(4).Cells(0).Style.BorderHT=   "Single"
         Sections(4).Cells(0).Style.BorderHI=   ""
         Sections(4).Cells(0).Style.BorderHB=   "Single"
         Sections(4).Cells(0).Style.BorderVL=   "Single"
         Sections(4).Cells(0).Style.BorderVI=   ""
         Sections(4).Cells(0).Style.BorderVR=   "Single"
         Sections(4).Cells(0).Style.NoClipping=   0   'False
         Sections(4).Cells(0).Style.RTF=   0   'False
         Sections(4).Cells(0).Style.fprops=   1474560
         Sections(5).Name=   "SECTION_3"
         Sections(5).Condition=   "IsLastRec()"
         Sections(5).StyleExp=   "'STYLE_1'"
         Sections(5).AutoHeight=   0   'False
         Sections(5).Height=   5
         Sections(5).Cells.Count=   8
         Sections(5).Cells(0).Name=   "CELL_0"
         Sections(5).Cells(0).Exp=   """                 Kasir"""
         Sections(5).Cells(0).NewLine=   -1  'True
         Sections(5).Cells(0).PrivateStyle=   -1  'True
         Sections(5).Cells(0).Style.Name=   "<private>"
         Sections(5).Cells(0).Style.ParentName=   "STYLE_1"
         Sections(5).Cells(0).Style.Font_Name=   "Courier"
         Sections(5).Cells(0).Style.Font_Size=   9.75
         Sections(5).Cells(0).Style.Font_Bold=   0   'False
         Sections(5).Cells(0).Style.Font_Italic=   0   'False
         Sections(5).Cells(0).Style.Font_Underline=   0   'False
         Sections(5).Cells(0).Style.Font_Strikeout=   0   'False
         Sections(5).Cells(0).Style.Font_Charset=   0
         Sections(5).Cells(0).Style.TextAlign=   3
         Sections(5).Cells(0).Style.TextVAlign=   1
         Sections(5).Cells(0).Style.TextWrap=   -1  'True
         Sections(5).Cells(0).Style.ForeColor=   0
         Sections(5).Cells(0).Style.BackColor=   16777215
         Sections(5).Cells(0).Style.NoFill=   -1  'True
         Sections(5).Cells(0).Style.BackPicFile=   ""
         Sections(5).Cells(0).Style.ForePicFile=   ""
         Sections(5).Cells(0).Style.BackPicVertPlacement=   0
         Sections(5).Cells(0).Style.BackPicHorzPlacement=   0
         Sections(5).Cells(0).Style.ForePicPlacement=   0
         Sections(5).Cells(0).Style.ForePicDrawMode=   0
         Sections(5).Cells(0).Style.MarginLeft=   6
         Sections(5).Cells(0).Style.MarginTop=   1
         Sections(5).Cells(0).Style.MarginRight=   6
         Sections(5).Cells(0).Style.MarginBottom=   1
         Sections(5).Cells(0).Style.HasBorders=   -1  'True
         Sections(5).Cells(0).Style.BorderHT=   "Single"
         Sections(5).Cells(0).Style.BorderHI=   ""
         Sections(5).Cells(0).Style.BorderHB=   ""
         Sections(5).Cells(0).Style.BorderVL=   ""
         Sections(5).Cells(0).Style.BorderVI=   ""
         Sections(5).Cells(0).Style.BorderVR=   ""
         Sections(5).Cells(0).Style.NoClipping=   0   'False
         Sections(5).Cells(0).Style.RTF=   0   'False
         Sections(5).Cells(0).Style.fprops=   294912
         Sections(5).Cells(1).Name=   "CELL_15"
         Sections(5).Cells(1).Exp=   """                           """
         Sections(5).Cells(1).PrivateStyle=   -1  'True
         Sections(5).Cells(1).Style.Name=   "<private>"
         Sections(5).Cells(1).Style.ParentName=   "STYLE_1"
         Sections(5).Cells(1).Style.Font_Name=   "Courier"
         Sections(5).Cells(1).Style.Font_Size=   9.75
         Sections(5).Cells(1).Style.Font_Bold=   0   'False
         Sections(5).Cells(1).Style.Font_Italic=   0   'False
         Sections(5).Cells(1).Style.Font_Underline=   0   'False
         Sections(5).Cells(1).Style.Font_Strikeout=   0   'False
         Sections(5).Cells(1).Style.Font_Charset=   0
         Sections(5).Cells(1).Style.TextAlign=   3
         Sections(5).Cells(1).Style.TextVAlign=   1
         Sections(5).Cells(1).Style.TextWrap=   -1  'True
         Sections(5).Cells(1).Style.ForeColor=   0
         Sections(5).Cells(1).Style.BackColor=   16777215
         Sections(5).Cells(1).Style.NoFill=   -1  'True
         Sections(5).Cells(1).Style.BackPicFile=   ""
         Sections(5).Cells(1).Style.ForePicFile=   ""
         Sections(5).Cells(1).Style.BackPicVertPlacement=   0
         Sections(5).Cells(1).Style.BackPicHorzPlacement=   0
         Sections(5).Cells(1).Style.ForePicPlacement=   0
         Sections(5).Cells(1).Style.ForePicDrawMode=   0
         Sections(5).Cells(1).Style.MarginLeft=   6
         Sections(5).Cells(1).Style.MarginTop=   1
         Sections(5).Cells(1).Style.MarginRight=   6
         Sections(5).Cells(1).Style.MarginBottom=   1
         Sections(5).Cells(1).Style.HasBorders=   -1  'True
         Sections(5).Cells(1).Style.BorderHT=   "Single"
         Sections(5).Cells(1).Style.BorderHI=   ""
         Sections(5).Cells(1).Style.BorderHB=   ""
         Sections(5).Cells(1).Style.BorderVL=   ""
         Sections(5).Cells(1).Style.BorderVI=   ""
         Sections(5).Cells(1).Style.BorderVR=   ""
         Sections(5).Cells(1).Style.NoClipping=   0   'False
         Sections(5).Cells(1).Style.RTF=   0   'False
         Sections(5).Cells(1).Style.fprops=   294912
         Sections(5).Cells(2).Name=   "CELL_1"
         Sections(5).Cells(2).Exp=   """Total : """
         Sections(5).Cells(2).Width=   14
         Sections(5).Cells(2).PrivateStyle=   -1  'True
         Sections(5).Cells(2).Style.Name=   "<private>"
         Sections(5).Cells(2).Style.ParentName=   "STYLE_1"
         Sections(5).Cells(2).Style.Font_Name=   "Courier"
         Sections(5).Cells(2).Style.Font_Size=   9.75
         Sections(5).Cells(2).Style.Font_Bold=   0   'False
         Sections(5).Cells(2).Style.Font_Italic=   0   'False
         Sections(5).Cells(2).Style.Font_Underline=   0   'False
         Sections(5).Cells(2).Style.Font_Strikeout=   0   'False
         Sections(5).Cells(2).Style.Font_Charset=   0
         Sections(5).Cells(2).Style.TextAlign=   2
         Sections(5).Cells(2).Style.TextVAlign=   1
         Sections(5).Cells(2).Style.TextWrap=   -1  'True
         Sections(5).Cells(2).Style.ForeColor=   0
         Sections(5).Cells(2).Style.BackColor=   16777215
         Sections(5).Cells(2).Style.NoFill=   -1  'True
         Sections(5).Cells(2).Style.BackPicFile=   ""
         Sections(5).Cells(2).Style.ForePicFile=   ""
         Sections(5).Cells(2).Style.BackPicVertPlacement=   0
         Sections(5).Cells(2).Style.BackPicHorzPlacement=   0
         Sections(5).Cells(2).Style.ForePicPlacement=   0
         Sections(5).Cells(2).Style.ForePicDrawMode=   0
         Sections(5).Cells(2).Style.MarginLeft=   6
         Sections(5).Cells(2).Style.MarginTop=   1
         Sections(5).Cells(2).Style.MarginRight=   6
         Sections(5).Cells(2).Style.MarginBottom=   1
         Sections(5).Cells(2).Style.HasBorders=   -1  'True
         Sections(5).Cells(2).Style.BorderHT=   "Single"
         Sections(5).Cells(2).Style.BorderHI=   ""
         Sections(5).Cells(2).Style.BorderHB=   ""
         Sections(5).Cells(2).Style.BorderVL=   ""
         Sections(5).Cells(2).Style.BorderVI=   ""
         Sections(5).Cells(2).Style.BorderVR=   ""
         Sections(5).Cells(2).Style.NoClipping=   0   'False
         Sections(5).Cells(2).Style.RTF=   0   'False
         Sections(5).Cells(2).Style.fprops=   32769
         Sections(5).Cells(3).Name=   "CELL_2"
         Sections(5).Cells(3).Exp=   "nSubTotal"
         Sections(5).Cells(3).Width=   15
         Sections(5).Cells(3).PrivateStyle=   -1  'True
         Sections(5).Cells(3).Format=   "###,###,###,###,###,##0.00"
         Sections(5).Cells(3).Style.Name=   "<private>"
         Sections(5).Cells(3).Style.ParentName=   "STYLE_1"
         Sections(5).Cells(3).Style.Font_Name=   "Courier"
         Sections(5).Cells(3).Style.Font_Size=   9.75
         Sections(5).Cells(3).Style.Font_Bold=   0   'False
         Sections(5).Cells(3).Style.Font_Italic=   0   'False
         Sections(5).Cells(3).Style.Font_Underline=   0   'False
         Sections(5).Cells(3).Style.Font_Strikeout=   0   'False
         Sections(5).Cells(3).Style.Font_Charset=   0
         Sections(5).Cells(3).Style.TextAlign=   3
         Sections(5).Cells(3).Style.TextVAlign=   1
         Sections(5).Cells(3).Style.TextWrap=   -1  'True
         Sections(5).Cells(3).Style.ForeColor=   0
         Sections(5).Cells(3).Style.BackColor=   16777215
         Sections(5).Cells(3).Style.NoFill=   -1  'True
         Sections(5).Cells(3).Style.BackPicFile=   ""
         Sections(5).Cells(3).Style.ForePicFile=   ""
         Sections(5).Cells(3).Style.BackPicVertPlacement=   0
         Sections(5).Cells(3).Style.BackPicHorzPlacement=   0
         Sections(5).Cells(3).Style.ForePicPlacement=   0
         Sections(5).Cells(3).Style.ForePicDrawMode=   0
         Sections(5).Cells(3).Style.MarginLeft=   6
         Sections(5).Cells(3).Style.MarginTop=   1
         Sections(5).Cells(3).Style.MarginRight=   6
         Sections(5).Cells(3).Style.MarginBottom=   1
         Sections(5).Cells(3).Style.HasBorders=   -1  'True
         Sections(5).Cells(3).Style.BorderHT=   "Single"
         Sections(5).Cells(3).Style.BorderHI=   ""
         Sections(5).Cells(3).Style.BorderHB=   ""
         Sections(5).Cells(3).Style.BorderVL=   ""
         Sections(5).Cells(3).Style.BorderVI=   ""
         Sections(5).Cells(3).Style.BorderVR=   ""
         Sections(5).Cells(3).Style.NoClipping=   0   'False
         Sections(5).Cells(3).Style.RTF=   0   'False
         Sections(5).Cells(3).Style.fprops=   1081344
         Sections(5).Cells(4).Name=   "CELL_4"
         Sections(5).Cells(4).Exp=   "keAkun"
         Sections(5).Cells(4).NewLine=   -1  'True
         Sections(5).Cells(4).PrivateStyle=   -1  'True
         Sections(5).Cells(4).Style.Name=   "<private>"
         Sections(5).Cells(4).Style.ParentName=   "STYLE_1"
         Sections(5).Cells(4).Style.Font_Name=   "Courier"
         Sections(5).Cells(4).Style.Font_Size=   9.75
         Sections(5).Cells(4).Style.Font_Bold=   0   'False
         Sections(5).Cells(4).Style.Font_Italic=   0   'False
         Sections(5).Cells(4).Style.Font_Underline=   0   'False
         Sections(5).Cells(4).Style.Font_Strikeout=   0   'False
         Sections(5).Cells(4).Style.Font_Charset=   0
         Sections(5).Cells(4).Style.TextAlign=   2
         Sections(5).Cells(4).Style.TextVAlign=   1
         Sections(5).Cells(4).Style.TextWrap=   -1  'True
         Sections(5).Cells(4).Style.ForeColor=   0
         Sections(5).Cells(4).Style.BackColor=   16777215
         Sections(5).Cells(4).Style.NoFill=   -1  'True
         Sections(5).Cells(4).Style.BackPicFile=   ""
         Sections(5).Cells(4).Style.ForePicFile=   ""
         Sections(5).Cells(4).Style.BackPicVertPlacement=   0
         Sections(5).Cells(4).Style.BackPicHorzPlacement=   0
         Sections(5).Cells(4).Style.ForePicPlacement=   0
         Sections(5).Cells(4).Style.ForePicDrawMode=   0
         Sections(5).Cells(4).Style.MarginLeft=   6
         Sections(5).Cells(4).Style.MarginTop=   1
         Sections(5).Cells(4).Style.MarginRight=   6
         Sections(5).Cells(4).Style.MarginBottom=   1
         Sections(5).Cells(4).Style.HasBorders=   -1  'True
         Sections(5).Cells(4).Style.BorderHT=   ""
         Sections(5).Cells(4).Style.BorderHI=   ""
         Sections(5).Cells(4).Style.BorderHB=   ""
         Sections(5).Cells(4).Style.BorderVL=   ""
         Sections(5).Cells(4).Style.BorderVI=   ""
         Sections(5).Cells(4).Style.BorderVR=   ""
         Sections(5).Cells(4).Style.NoClipping=   0   'False
         Sections(5).Cells(4).Style.RTF=   0   'False
         Sections(5).Cells(4).Style.fprops=   1
         Sections(5).Cells(5).Name=   "CELL_5"
         Sections(5).Cells(5).NewLine=   -1  'True
         Sections(5).Cells(6).Name=   "CELL_16"
         Sections(5).Cells(6).NewLine=   -1  'True
         Sections(5).Cells(6).Height=   5
         Sections(5).Cells(6).PrivateStyle=   -1  'True
         Sections(5).Cells(6).Style.Name=   "<private>"
         Sections(5).Cells(6).Style.ParentName=   "STYLE_1"
         Sections(5).Cells(6).Style.Font_Name=   "Courier"
         Sections(5).Cells(6).Style.Font_Size=   9.75
         Sections(5).Cells(6).Style.Font_Bold=   0   'False
         Sections(5).Cells(6).Style.Font_Italic=   0   'False
         Sections(5).Cells(6).Style.Font_Underline=   0   'False
         Sections(5).Cells(6).Style.Font_Strikeout=   0   'False
         Sections(5).Cells(6).Style.Font_Charset=   0
         Sections(5).Cells(6).Style.TextAlign=   2
         Sections(5).Cells(6).Style.TextVAlign=   1
         Sections(5).Cells(6).Style.TextWrap=   -1  'True
         Sections(5).Cells(6).Style.ForeColor=   0
         Sections(5).Cells(6).Style.BackColor=   16777215
         Sections(5).Cells(6).Style.NoFill=   -1  'True
         Sections(5).Cells(6).Style.BackPicFile=   ""
         Sections(5).Cells(6).Style.ForePicFile=   ""
         Sections(5).Cells(6).Style.BackPicVertPlacement=   0
         Sections(5).Cells(6).Style.BackPicHorzPlacement=   0
         Sections(5).Cells(6).Style.ForePicPlacement=   0
         Sections(5).Cells(6).Style.ForePicDrawMode=   0
         Sections(5).Cells(6).Style.MarginLeft=   6
         Sections(5).Cells(6).Style.MarginTop=   1
         Sections(5).Cells(6).Style.MarginRight=   6
         Sections(5).Cells(6).Style.MarginBottom=   1
         Sections(5).Cells(6).Style.HasBorders=   -1  'True
         Sections(5).Cells(6).Style.BorderHT=   ""
         Sections(5).Cells(6).Style.BorderHI=   ""
         Sections(5).Cells(6).Style.BorderHB=   ""
         Sections(5).Cells(6).Style.BorderVL=   ""
         Sections(5).Cells(6).Style.BorderVI=   ""
         Sections(5).Cells(6).Style.BorderVR=   ""
         Sections(5).Cells(6).Style.NoClipping=   0   'False
         Sections(5).Cells(6).Style.RTF=   0   'False
         Sections(5).Cells(6).Style.fprops=   2064389
         Sections(5).Cells(7).Name=   "CELL_20"
         Sections(5).Cells(7).Exp=   "cFooter2"
         Sections(5).Cells(7).NewLine=   -1  'True
         Sections(5).Cells(7).PrivateStyle=   -1  'True
         Sections(5).Cells(7).Style.Name=   "<private>"
         Sections(5).Cells(7).Style.ParentName=   "STYLE_1"
         Sections(5).Cells(7).Style.Font_Name=   "Courier"
         Sections(5).Cells(7).Style.Font_Size=   9.75
         Sections(5).Cells(7).Style.Font_Bold=   0   'False
         Sections(5).Cells(7).Style.Font_Italic=   0   'False
         Sections(5).Cells(7).Style.Font_Underline=   0   'False
         Sections(5).Cells(7).Style.Font_Strikeout=   0   'False
         Sections(5).Cells(7).Style.Font_Charset=   0
         Sections(5).Cells(7).Style.TextAlign=   2
         Sections(5).Cells(7).Style.TextVAlign=   1
         Sections(5).Cells(7).Style.TextWrap=   -1  'True
         Sections(5).Cells(7).Style.ForeColor=   0
         Sections(5).Cells(7).Style.BackColor=   16777215
         Sections(5).Cells(7).Style.NoFill=   -1  'True
         Sections(5).Cells(7).Style.BackPicFile=   ""
         Sections(5).Cells(7).Style.ForePicFile=   ""
         Sections(5).Cells(7).Style.BackPicVertPlacement=   0
         Sections(5).Cells(7).Style.BackPicHorzPlacement=   0
         Sections(5).Cells(7).Style.ForePicPlacement=   0
         Sections(5).Cells(7).Style.ForePicDrawMode=   0
         Sections(5).Cells(7).Style.MarginLeft=   6
         Sections(5).Cells(7).Style.MarginTop=   1
         Sections(5).Cells(7).Style.MarginRight=   6
         Sections(5).Cells(7).Style.MarginBottom=   1
         Sections(5).Cells(7).Style.HasBorders=   -1  'True
         Sections(5).Cells(7).Style.BorderHT=   ""
         Sections(5).Cells(7).Style.BorderHI=   ""
         Sections(5).Cells(7).Style.BorderHB=   ""
         Sections(5).Cells(7).Style.BorderVL=   ""
         Sections(5).Cells(7).Style.BorderVI=   ""
         Sections(5).Cells(7).Style.BorderVR=   ""
         Sections(5).Cells(7).Style.NoClipping=   0   'False
         Sections(5).Cells(7).Style.RTF=   0   'False
         Sections(5).Cells(7).Style.fprops=   2064385
         Styles.Count    =   6
         Styles(0).Name  =   "Tdb_Base"
         Styles(0).ParentName=   ""
         Styles(0).Font_Name=   "Courier"
         Styles(0).Font_Size=   9.75
         Styles(0).Font_Bold=   -1  'True
         Styles(0).Font_Charset=   0
         Styles(0).TextVAlign=   1
         Styles(0).MarginTop=   1
         Styles(0).MarginBottom=   1
         Styles(1).Name  =   "STYLE_1"
         Styles(1).ParentName=   "Tdb_Base"
         Styles(1).Font_Name=   "Courier"
         Styles(1).Font_Size=   9.75
         Styles(1).Font_Charset=   0
         Styles(1).TextVAlign=   1
         Styles(1).MarginTop=   1
         Styles(1).MarginBottom=   1
         Styles(1).fprops=   18087936
         Styles(2).Name  =   "Tdb_Body"
         Styles(2).ParentName=   "Tdb_Base"
         Styles(2).Font_Name=   "Courier"
         Styles(2).Font_Size=   9.75
         Styles(2).Font_Charset=   0
         Styles(2).TextVAlign=   1
         Styles(2).MarginTop=   0
         Styles(2).MarginBottom=   0
         Styles(2).fprops=   18862080
         Styles(3).Name  =   "Tdb_Header"
         Styles(3).ParentName=   "Tdb_Base"
         Styles(3).Font_Name=   "Courier"
         Styles(3).Font_Size=   9.75
         Styles(3).Font_Bold=   -1  'True
         Styles(3).Font_Charset=   0
         Styles(3).TextAlign=   0
         Styles(3).TextVAlign=   1
         Styles(3).MarginTop=   1
         Styles(3).MarginBottom=   1
         Styles(3).BorderHT=   "Single"
         Styles(3).BorderHI=   "Single"
         Styles(3).BorderHB=   "Single"
         Styles(3).fprops=   2064385
         Styles(4).Name  =   "Tdb_PageFooter"
         Styles(4).ParentName=   "Tdb_Base"
         Styles(4).Font_Name=   "Courier"
         Styles(4).Font_Size=   9.75
         Styles(4).Font_Bold=   -1  'True
         Styles(4).Font_Charset=   0
         Styles(4).TextVAlign=   1
         Styles(4).MarginTop=   1
         Styles(4).MarginBottom=   1
         Styles(4).BorderHT=   "Single"
         Styles(4).fprops=   163840
         Styles(5).Name  =   "Garis"
         Styles(5).ParentName=   "Tdb_Base"
         Styles(5).Font_Name=   "Courier"
         Styles(5).Font_Size=   9.75
         Styles(5).Font_Bold=   -1  'True
         Styles(5).Font_Charset=   0
         Styles(5).TextAlign=   2
         Styles(5).TextVAlign=   1
         Styles(5).MarginTop=   1
         Styles(5).MarginBottom=   1
         Styles(5).BorderHT=   "Single"
         Styles(5).fprops=   32769
         Lines.Count     =   4
         Lines(0).Name   =   "Single"
         Lines(0).Thickness=   4
         Lines(1).Name   =   "Double"
         Lines(1).Thickness=   5
         Lines(2).Name   =   "Quarter"
         Lines(2).Thickness=   1
         Lines(2).Color  =   8421504
         Lines(3).Name   =   "None"
         Profiles.Count  =   1
         Profiles(0).Name=   "PROFILE_0"
         Profiles(0).Active=   -1  'True
         Profiles(0).PreviewNoMinimize=   -1  'True
         Profiles(0).PreviewNoMaximize=   -1  'True
         Profiles(0).PreviewNoResize=   -1  'True
         Profiles(0).PreviewMaximized=   -1  'True
         Profiles(0).PreviewNoSaveLoad=   -1  'True
         Profiles(0).PrinterMarginLeft=   10
         Profiles(0).PrinterMarginTop=   5
         Profiles(0).PrinterMarginRight=   10
         Profiles(0).PrinterMarginBottom=   5
         Profiles(0).PrinterPaperSize=   256
         Profiles(0).PrinterPaperHeight=   139
         Profiles(0).PrinterPaperWidth=   215
         Profiles(0).PrinterMargins_set=   -1  'True
         Profiles(0).PrinterPaperSize_set=   -1  'True
         Profiles(0).PrinterPaperUserSize_set=   -1  'True
      End
      Begin BiSANumberBoxProject.BiSANumberBox nDiscount 
         Height          =   330
         Left            =   3345
         TabIndex        =   23
         Top             =   1245
         Width           =   3420
         _ExtentX        =   6033
         _ExtentY        =   582
         Appearance      =   0
         Enabled         =   0   'False
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Discount"
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
      Begin BiSANumberBoxProject.BiSANumberBox nTotal 
         Height          =   330
         Left            =   3345
         TabIndex        =   24
         Top             =   1605
         Width           =   3420
         _ExtentX        =   6033
         _ExtentY        =   582
         Appearance      =   0
         Enabled         =   0   'False
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Total"
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
      Begin BiSANumberBoxProject.BiSANumberBox nLunas 
         Height          =   330
         Left            =   3345
         TabIndex        =   25
         Top             =   2115
         Width           =   3420
         _ExtentX        =   6033
         _ExtentY        =   582
         Appearance      =   0
         Enabled         =   0   'False
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Lunas"
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
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   3360
         X2              =   6735
         Y1              =   1995
         Y2              =   1995
      End
      Begin VB.Label Label1 
         Caption         =   "Saldo Piutang Terikini (update)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   390
         TabIndex        =   22
         Top             =   960
         Width           =   4005
      End
      Begin VB.Label Label2 
         Caption         =   "Saldo Top Up : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3180
         TabIndex        =   18
         Top             =   570
         Width           =   1170
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame2 
      Height          =   4215
      Left            =   180
      Top             =   2715
      Width           =   13950
      _ExtentX        =   24606
      _ExtentY        =   7435
      Caption         =   "DATA PIUTANG"
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      BackColor       =   -2147483633
      Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
         Height          =   3915
         Left            =   60
         TabIndex        =   0
         Top             =   270
         Width           =   13875
         _ExtentX        =   24474
         _ExtentY        =   6906
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   4
         Columns(0)._MaxComboItems=   5
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "No"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "FAKTUR"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "TGL"
         Columns(3).DataField=   ""
         Columns(3).NumberFormat=   "dd-MM-yyyy"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "PIUTANG"
         Columns(4).DataField=   ""
         Columns(4).NumberFormat=   "###,###,###,###,##0.00"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "JATUH TEMPO"
         Columns(5).DataField=   ""
         Columns(5).NumberFormat=   "dd-MM-yyyy"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "DISC RP"
         Columns(6).DataField=   ""
         Columns(6).NumberFormat=   "###,###,###,###,##0.00"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "PELUNASAN"
         Columns(7).DataField=   ""
         Columns(7).NumberFormat=   "###,###,###,###,##0.00"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   8
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).ScrollBars=   2
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=8"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=609"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=529"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=197124"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=1005"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=926"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=197124"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=4551"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=4471"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=197124"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=3016"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2937"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=197121"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=4022"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=3942"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=197122"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(26)=   "Column(5).Width=3096"
         Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=3016"
         Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=197121"
         Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(31)=   "Column(6).Width=2302"
         Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=2223"
         Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=197122"
         Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(36)=   "Column(7).Width=4604"
         Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=4524"
         Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=197122"
         Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         ColumnFooters   =   -1  'True
         DataMode        =   4
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   16777215
         RowDividerColor =   12632256
         RowSubDividerColor=   12632256
         DirectionAfterEnter=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000005&,.bold=0,.fontsize=825"
         _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.bgcolor=&H8000000F&"
         _StyleDefs(11)  =   ":id=2,.fgcolor=&H80000007&,.bold=0,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(12)  =   ":id=2,.strikethrough=0,.charset=0"
         _StyleDefs(13)  =   ":id=2,.fontname=Tahoma"
         _StyleDefs(14)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.alignment=1,.bold=0,.fontsize=825"
         _StyleDefs(15)  =   ":id=3,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(16)  =   ":id=3,.fontname=Tahoma"
         _StyleDefs(17)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(18)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(19)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(20)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(21)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(23)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(24)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(25)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(26)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(27)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(28)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(29)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(30)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(31)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(32)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(33)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(34)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(35)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(36)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
         _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
         _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
         _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
         _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
         _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
         _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
         _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=70,.parent=13,.alignment=2"
         _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=67,.parent=14"
         _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=68,.parent=15"
         _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=69,.parent=17"
         _StyleDefs(53)  =   "Splits(0).Columns(4).Style:id=58,.parent=13,.alignment=1"
         _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=55,.parent=14"
         _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=56,.parent=15"
         _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=57,.parent=17"
         _StyleDefs(57)  =   "Splits(0).Columns(5).Style:id=74,.parent=13,.alignment=2"
         _StyleDefs(58)  =   "Splits(0).Columns(5).HeadingStyle:id=71,.parent=14"
         _StyleDefs(59)  =   "Splits(0).Columns(5).FooterStyle:id=72,.parent=15"
         _StyleDefs(60)  =   "Splits(0).Columns(5).EditorStyle:id=73,.parent=17"
         _StyleDefs(61)  =   "Splits(0).Columns(6).Style:id=50,.parent=13,.alignment=1"
         _StyleDefs(62)  =   "Splits(0).Columns(6).HeadingStyle:id=47,.parent=14"
         _StyleDefs(63)  =   "Splits(0).Columns(6).FooterStyle:id=48,.parent=15"
         _StyleDefs(64)  =   "Splits(0).Columns(6).EditorStyle:id=49,.parent=17"
         _StyleDefs(65)  =   "Splits(0).Columns(7).Style:id=54,.parent=13,.alignment=1,.bgcolor=&HFFFF80&"
         _StyleDefs(66)  =   ":id=54,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(67)  =   ":id=54,.fontname=Tahoma"
         _StyleDefs(68)  =   "Splits(0).Columns(7).HeadingStyle:id=51,.parent=14"
         _StyleDefs(69)  =   "Splits(0).Columns(7).FooterStyle:id=52,.parent=15"
         _StyleDefs(70)  =   "Splits(0).Columns(7).EditorStyle:id=53,.parent=17"
         _StyleDefs(71)  =   "Named:id=33:Normal"
         _StyleDefs(72)  =   ":id=33,.parent=0"
         _StyleDefs(73)  =   "Named:id=34:Heading"
         _StyleDefs(74)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(75)  =   ":id=34,.wraptext=-1"
         _StyleDefs(76)  =   "Named:id=35:Footing"
         _StyleDefs(77)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(78)  =   "Named:id=36:Selected"
         _StyleDefs(79)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(80)  =   "Named:id=37:Caption"
         _StyleDefs(81)  =   ":id=37,.parent=34,.alignment=2,.bold=0,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(82)  =   ":id=37,.strikethrough=0,.charset=0"
         _StyleDefs(83)  =   ":id=37,.fontname=Tahoma"
         _StyleDefs(84)  =   "Named:id=38:HighlightRow"
         _StyleDefs(85)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(86)  =   "Named:id=39:EvenRow"
         _StyleDefs(87)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(88)  =   "Named:id=40:OddRow"
         _StyleDefs(89)  =   ":id=40,.parent=33"
         _StyleDefs(90)  =   "Named:id=41:RecordSelector"
         _StyleDefs(91)  =   ":id=41,.parent=34"
         _StyleDefs(92)  =   "Named:id=42:FilterBar"
         _StyleDefs(93)  =   ":id=42,.parent=33"
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   2625
      Left            =   240
      Top             =   75
      Width           =   6900
      _ExtentX        =   12171
      _ExtentY        =   4630
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      BackColor       =   -2147483633
      Begin BiSATextBoxProject.BiSABrowse cAkunKas 
         Height          =   330
         Left            =   75
         TabIndex        =   12
         Top             =   2175
         Visible         =   0   'False
         Width           =   3480
         _ExtentX        =   6138
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
         Button          =   -1  'True
         Caption         =   "Akun Kas"
         CaptionWidth    =   1300
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
      Begin BiSATextBoxProject.BiSABrowse cCustomer 
         Height          =   330
         Left            =   75
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   735
         Width           =   3225
         _ExtentX        =   5689
         _ExtentY        =   582
         Text            =   "12345678"
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
         Enabled         =   0   'False
         Appearance      =   0
         Caption         =   "Member"
         CaptionWidth    =   1300
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
      Begin BiSADateProject.BiSADate dTanggal 
         Height          =   330
         Left            =   75
         TabIndex        =   2
         Top             =   375
         Width           =   2790
         _ExtentX        =   4921
         _ExtentY        =   582
         Value           =   "19-11-2003"
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
         Caption         =   "Tanggal"
         CaptionWidth    =   1300
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
      Begin BiSATextBoxProject.BiSABrowse cNama 
         Height          =   330
         Left            =   3330
         TabIndex        =   3
         Top             =   735
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   582
         Text            =   "123456"
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
         MaxLength       =   50
         Appearance      =   0
         Button          =   -1  'True
         CaptionWidth    =   1500
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
      Begin BiSATextBoxProject.BiSABrowse cAlamat 
         Height          =   330
         Left            =   75
         TabIndex        =   4
         Top             =   1095
         Width           =   6570
         _ExtentX        =   11589
         _ExtentY        =   582
         Text            =   "123456"
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
         BackColor       =   -2147483633
         Enabled         =   0   'False
         MaxLength       =   50
         Appearance      =   0
         Caption         =   "Alamat"
         CaptionWidth    =   1300
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
      Begin BiSATextBoxProject.BiSABrowse cFaktur 
         Height          =   330
         Left            =   75
         TabIndex        =   11
         Top             =   1800
         Width           =   3480
         _ExtentX        =   6138
         _ExtentY        =   582
         Text            =   "12345678"
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
         Button          =   -1  'True
         Caption         =   "Faktur"
         CaptionWidth    =   1300
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
      Begin BiSATextBoxProject.BiSABrowse cNamaDepartmen 
         Height          =   330
         Left            =   1500
         TabIndex        =   14
         Top             =   1455
         Width           =   3240
         _ExtentX        =   5715
         _ExtentY        =   582
         Text            =   "123456"
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
         BackColor       =   -2147483633
         Enabled         =   0   'False
         MaxLength       =   50
         Appearance      =   0
         CaptionWidth    =   1300
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
      Begin VB.Label lbCostCenter 
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   135
         TabIndex        =   13
         Top             =   90
         Width           =   6030
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame4 
      Height          =   630
      Left            =   165
      Top             =   7035
      Width           =   13935
      _ExtentX        =   24580
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
      Begin BiSANumberBoxProject.BiSANumberBox nPoinReguler 
         Height          =   435
         Left            =   8610
         TabIndex        =   17
         Top             =   120
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   767
         BorderStyle     =   0
         Decimals        =   0
         Enabled         =   0   'False
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   12632319
         Caption         =   " POIN"
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
      Begin BiSAButtonProject.BiSAButton cmdPerbaikan 
         Height          =   435
         Left            =   2835
         TabIndex        =   16
         Top             =   105
         Visible         =   0   'False
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   767
         Caption         =   "X"
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
      Begin BiSAButtonProject.BiSAButton cmdPrint 
         Height          =   435
         Left            =   2340
         TabIndex        =   15
         Top             =   105
         Width           =   465
         _ExtentX        =   820
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
         BackColor       =   -2147483633
         Picture         =   "trPelunasanPiutang.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdHapus 
         Height          =   435
         Left            =   1170
         TabIndex        =   5
         Top             =   105
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   767
         Caption         =   "    &Delete"
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
         Picture         =   "trPelunasanPiutang.frx":059A
      End
      Begin BiSAButtonProject.BiSAButton cmdAktivasi 
         Height          =   435
         Left            =   11190
         TabIndex        =   6
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
         Picture         =   "trPelunasanPiutang.frx":0824
      End
      Begin BiSAButtonProject.BiSAButton cmdEdit 
         Height          =   435
         Left            =   4785
         TabIndex        =   7
         Top             =   105
         Visible         =   0   'False
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   767
         Caption         =   "  &Edit"
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
         Picture         =   "trPelunasanPiutang.frx":09C3
      End
      Begin BiSAButtonProject.BiSAButton cmdAdd 
         Height          =   435
         Left            =   90
         TabIndex        =   8
         Top             =   105
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   767
         Caption         =   "  &Add"
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
         Picture         =   "trPelunasanPiutang.frx":0AEF
      End
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   12735
         TabIndex        =   9
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
         Picture         =   "trPelunasanPiutang.frx":0C9A
      End
      Begin BiSAButtonProject.BiSAButton cmdSimpan 
         Height          =   435
         Left            =   11640
         TabIndex        =   10
         Top             =   120
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   767
         Caption         =   "    &Save"
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
         Picture         =   "trPelunasanPiutang.frx":0D40
      End
   End
End
Attribute VB_Name = "trPelunasanPiutang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim nPos As SisPos
Dim lClick As Boolean
Dim lStart As Boolean
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim objMenu As New CodeSuiteLibrary.Menu
Dim vaArray As New XArrayDB
Dim lEdit As Boolean
Dim cPosKas As String
Public lPubStatus As Boolean
Public vaPubReff As New XArrayDB
Public nPubTotal As Double
Public cPubAkun As String
Public lClose As Double
Public nTarikTunai As Double
Public nWithDraw As Double
Public nSisaKurangTopUp As Double
Public lTarikTunai As Boolean
Public nSaldoTopUp As Double
Public nTunai As Double
Public nKembalian As Double
Public nJaminan As Double
Public nTotYgHarusDibayar As Double
Public nMetodePembayaran As Integer

Sub GetEdit(lPar As Boolean)
  lEdit = lPar
  BiSAFrame1.Enabled = lPar
  BiSAFrame2.Enabled = lPar
  BiSAFrame3.Enabled = lPar
  SetButton cmdSimpan, cmdKeluar, cmdAdd, cmdEdit, cmdHapus, nPos, lPar, cmdAktivasi
End Sub

Private Sub cAkunKas_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "akun", "kodeakun,keterangan", "jenis", sisAssign, "D", " and left(kodeakun,1)='1'")
  If Not dbData.EOF Then
    cAkunKas.Text = cAkunKas.Browse(dbData, Array("Kode Akun", "Keterangan"), , Array(25, 30))
  End If
End Sub

Private Sub cCustomer_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "anggota a", "a.kodeanggota,a.nama,a.alamat,d.keterangan as namadep", "a.kodeanggota", sisContent, cCustomer.Text, , , Array("left join dep d on d.kodedep = a.kodedep"))
  If Not dbData.EOF Then
    cCustomer.Text = cCustomer.Browse(dbData)
    cCustomer.Text = GetNull(dbData!kodeanggota)
    cNama.Text = GetNull(dbData!nama)
    cNamaDepartmen.Text = GetNull(dbData!namadep)
    cAlamat.Text = GetNull(dbData!alamat)
    If nPos = Add Then
      GetData
    End If
  End If
End Sub

Private Sub GetMark()
Dim n As Double
  
  n = TDBGrid1.Bookmark
  If n >= 0 Then
    vaArray(n, 0) = Not vaArray(n, 0)
    TDBGrid1.Columns(0) = vaArray(n, 0)
  End If
End Sub

Private Sub GetData()
Dim n As Integer
Dim nSisaPiutang As Double
Dim nTmpSisaPiutang As Double

  vaArray.ReDim 0, -1, 0, 7
  nTmpSisaPiutang = 0
  nSaldoPiutang.value = 0
  Set dbData = objData.Browse(GetDSN, "totpenjualan", "nomorpenjualan,tgl,piutang,jthtmp", "kodeanggota", sisAssign, cCustomer.Text, " and flaglunas = 0", "tgl desc")
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      If Not isLunas(objData, GetNull(dbData!nomorpenjualan), nSisaPiutang) Then
        vaArray.InsertRows vaArray.UpperBound(1) + 1
        n = vaArray.UpperBound(1)
        vaArray(n, 0) = 0
        vaArray(n, 1) = n + 1
        vaArray(n, 2) = GetNull(dbData!nomorpenjualan)
        vaArray(n, 3) = GetNull(dbData!tgl)
        isLunas objData, vaArray(n, 2), nSisaPiutang
        vaArray(n, 4) = nSisaPiutang 'GetNull(dbData!Piutang)
        nTmpSisaPiutang = nTmpSisaPiutang + vaArray(n, 4)
        vaArray(n, 5) = GetNull(dbData!jthtmp)
        vaArray(n, 6) = 0
        'awalny nilai ini diset 0
        'kebijakan baru, tidak ada lagi yg boleh melunasi hutang separo separo, vaArray(n,7) diset = vaArray(n,4)
        vaArray(n, 7) = vaArray(n, 4)
      End If
      dbData.MoveNext
    Loop
  End If
  Set TDBGrid1.Array = vaArray
  TDBGrid1.Columns(4).FooterText = Format(nTmpSisaPiutang, "###,###,###,##0.00")
  TDBGrid1.ReBind
  TDBGrid1.Refresh
  
  'cari jumlah/saldo piutang terakhir setelah dipotong retur
  Set dbData = objData.Browse(GetDSN, "kartupiutang", "sum(debet-kredit) as saldopiutang", "kodeanggota", sisAssign, cCustomer.Text, " and tgl <='" & Format(Date, "yyyy-MM-dd") & "'")
  If Not dbData.EOF Then
    nSaldoPiutang.value = GetNull(dbData!saldopiutang)
  End If

  'cari saldo top up member
  Set dbData = objData.Browse(GetDSN, "membertopup m", "m.kodeanggota,a.nama,a.alamat,sum(debet) as debet,sum(kredit) as kredit,sum(m.debet-m.kredit) as saldo", "m.kodeanggota", sisAssign, cCustomer.Text, " GROUP BY m.kodeanggota", , Array("left join anggota a on a.kodeanggota = m.kodeanggota"))
  If Not dbData.EOF Then
    nSaldoTopUpMember.value = GetNull(dbData!saldo)
  Else
    nSaldoTopUpMember.value = 0
  End If
  
End Sub

Private Sub cFaktur_ButtonClick()
Dim n As Integer
Dim lSave As Boolean


  If aCfg(objData, msOtorisasiPenuh) = "Y" Then
    If GetRegistry(reg_UserLevel) <> 0 Then
      If objMenu.GetPassword("", Me, GetDSN) Then
        If objMenu.UserLevel <> 0 Then
            MsgBox "Maaf, Anda tidak diberikan wewenang untuk melakukan PENGEDITAN." & vbCrLf & _
                   "Hanya user dengan LEVEL 0 (SUPERVISOR) yg diijinkan", vbInformation, "OTORISASI not ALLOWED"
            Exit Sub
'        Else
'          MsgBox "OTORISASI DIBATALKAN", vbCritical
'          Exit Sub
        End If
      Else
        Exit Sub
      End If
    End If
  End If


  lSave = True
  vaArray.ReDim 0, -1, 0, 7
  Set dbData = objData.Browse(GetDSN, "totpelunasanpiutang", "nomorpelunasanpiutang,pelunasan", "tgl", sisAssign, Format(dTanggal.value, "yyyy-MM-dd"), " and kodeanggota = '" & cCustomer.Text & "'", "nomorpelunasanpiutang")
  If Not dbData.EOF Then
    cFaktur.Text = cFaktur.Browse(dbData)
    
    Set dbData = objData.Browse(GetDSN, "totpelunasanpiutang", , "nomorpelunasanpiutang", sisAssign, cFaktur.Text)
    If Not dbData.EOF Then
      nTotal.value = GetNull(dbData!Total)
      nDiscount.value = GetNull(dbData!Discount)
      nLunas.value = GetNull(dbData!Pelunasan)
      cAkunKas.Text = GetNull(dbData!kodeakun)
    End If
    
    Set dbData = objData.Browse(GetDSN, "pelunasanpiutang p", "p.nomorpenjualan,t.tgl,t.jthtmp,p.piutang,p.discount,p.pelunasan", "nomorpelunasanpiutang", sisAssign, cFaktur.Text, , , Array("LEFT JOIN totpenjualan t ON t.nomorpenjualan = p.nomorpenjualan"))
    If Not dbData.EOF Then
      Do While Not dbData.EOF
        vaArray.InsertRows vaArray.UpperBound(1) + 1
        n = vaArray.UpperBound(1)
        vaArray(n, 0) = -1
        vaArray(n, 1) = n + 1
        vaArray(n, 2) = GetNull(dbData!nomorpenjualan)
        vaArray(n, 3) = GetNull(dbData!tgl)
        vaArray(n, 4) = GetNull(dbData!Piutang)
        vaArray(n, 5) = GetNull(dbData!jthtmp)
        vaArray(n, 6) = GetNull(dbData!Discount)
        vaArray(n, 7) = GetNull(dbData!Pelunasan)
        dbData.MoveNext
      Loop
      SumTDB
    End If
    Set TDBGrid1.Array = vaArray
    TDBGrid1.ReBind
    TDBGrid1.Refresh
    Me.Refresh
    If nPos = Delete Then
      Me.Refresh
      If MsgBox("Yakin data akan dihapus?", vbYesNo) = vbYes Then
        objData.Start GetDSN
        lStart = True
        'CEK
        'tidak boleh dilakukan penghapusan atau pengkoreksian apabila faktur bg yg bersangkutang sudah dilunasi
        Set dbData = objData.Browse(GetDSN, "pencairanbg", , "nomorpelunasanpiutang", sisAssign, cFaktur.Text)
        If Not dbData.EOF Then
          MsgBox "Maaf transaksi ini tidak bisa dikoreksi kembali. BG/Cek sudah dicairkan"
          Exit Sub
        End If
        
        lSave = IIf(lSave, DelKodeTr(objData, msPelunasanPiutang, cFaktur.Text), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "pelunasanpiutang", "nomorpelunasanpiutang", sisAssign, cFaktur.Text), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "kartupiutang", "nomorkartupiutang", sisAssign, cFaktur.Text), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "totpelunasanpiutang", "nomorpelunasanpiutang", sisAssign, cFaktur.Text), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "bukubesar", "faktur", sisAssign, cFaktur.Text), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "bg", "nomorpelunasanpiutang", sisAssign, cFaktur.Text), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "membertopup", "nomormembertopup ", sisAssign, cFaktur.Text), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "totmembertopup", "nomormembertopup ", sisAssign, cFaktur.Text), False)
        lSave = IIf(lSave, objData.Delete(GetDSN, "poinhadiah", "faktur", sisAssign, cFaktur.Text), False)
        
        For n = vaArray.LowerBound(1) To vaArray.UpperBound(1)
          If lCekStatusLunas(objData, vaArray(n, 2)) = True Then
            lSave = IIf(lSave, objData.Edit(GetDSN, "penjualan", "nomorpenjualan = '" & vaArray(n, 2) & "'", Array("statuslunas"), Array(1)), False)
            lSave = IIf(lSave, objData.Edit(GetDSN, "totpenjualan", "nomorpenjualan = '" & vaArray(n, 2) & "'", Array("flaglunas"), Array(1)), False)
          Else
            lSave = IIf(lSave, objData.Edit(GetDSN, "penjualan", "nomorpenjualan = '" & vaArray(n, 2) & "'", Array("statuslunas"), Array(0)), False)
            lSave = IIf(lSave, objData.Edit(GetDSN, "totpenjualan", "nomorpenjualan = '" & vaArray(n, 2) & "'", Array("flaglunas"), Array(0)), False)
          End If
        Next n
        
        If lSave Then
          objData.Save GetDSN
        Else
          objData.Cancel GetDSN
        End If
      End If
      GetEdit False
      initvalue
    End If
    If nPos = Edit Then
      SendKeysA vbKeyReturn, True
    End If
  End If
End Sub

Private Sub Check1_Click()

End Sub

Private Sub cmdAdd_Click()
  nPos = Add
  GetEdit True
  initvalue
  dTanggal.SetFocus
  GetBrowseFaktur False
  cFaktur.Text = CreateNomorFaktur(objData, sisModulTransaksi.pelunasanpiutang, "totpelunasanpiutang", "nomorpelunasanpiutang")
End Sub

Private Sub cmdAktivasi_Click()
  frmAktivasi.Action Me
End Sub

Private Sub cmdEdit_Click()
  If GetRegistry(reg_UserLevel) <> 0 Then
    If objMenu.GetPassword("", Me, GetDSN) Then
      If objMenu.UserLevel <> 0 Then
        MsgBox "Maaf, Anda tidak diberikan wewenang untuk melakukan PENGEDITAN." & vbCrLf & _
               "Hanya user dengan LEVEL 0 (SUPERVISOR) yg diijinkan", vbInformation, "OTORISASI not ALLOWED"
        Exit Sub
      End If
    Else
      Unload Me
      GetEdit False
      Exit Sub
    End If
  End If
  
  nPos = Edit
  GetEdit True
  initvalue
  dTanggal.SetFocus
  GetBrowseFaktur True
End Sub

Private Sub cmdHapus_Click()
  If GetRegistry(reg_UserLevel) <> 0 Then
    If objMenu.GetPassword("", Me, GetDSN) Then
      If objMenu.UserLevel <> 0 Then
        MsgBox "Maaf, Anda tidak diberikan wewenang untuk melakukan PENGEDITAN." & vbCrLf & _
               "Hanya user dengan LEVEL 0 (SUPERVISOR) yg diijinkan", vbInformation, "OTORISASI not ALLOWED"
        Exit Sub
      End If
    Else
      Unload Me
      GetEdit False
      Exit Sub
    End If
  End If
  nPos = Delete
  GetEdit True
  initvalue
  dTanggal.SetFocus
  GetBrowseFaktur True
End Sub

Private Sub GetBrowseFaktur(ByVal lStat As Boolean)
  cFaktur.Button = lStat
  cFaktur.Enabled = lStat
End Sub

Private Sub cmdKeluar_Click()
  Unload trLunasPiutang
  If lEdit Then
    initvalue
    GetEdit False
  Else
    Unload Me
  End If
End Sub

Private Sub initvalue()
  cFaktur.Default
  dTanggal.value = Date
  cCustomer.Default
  cNama.Default
  cAlamat.Default
  cAkunKas.Text = cKasTeller
  nDiscount.Default
  nTotal.Default
  nLunas.Default
  nSaldoPiutang.Default
  cNamaDepartmen.Default
  nSaldoTopUpMember.Default
  nPoinReguler.Default
  cNama.Enabled = True
  ClearTdbgrid
  TDBGrid1.Columns(7).FooterText = ""
  TDBGrid1.Columns(4).FooterText = ""
  vaPubReff.ReDim 0, 2, 0, 2
  Label1.Caption = "Saldo Piutang Terikini (update) per tgl " & Format(Date, "dd-MM-yyyy")
  lClose = True
End Sub

Private Sub initvalue2()
  cFaktur.Default
  dTanggal.value = Date
  cCustomer.Default
  cNama.Default
  cAlamat.Default
  cAkunKas.Text = cKasTeller
  nDiscount.Default
  nTotal.Default
  nLunas.Default
  nSaldoPiutang.Default
  cNamaDepartmen.Default
  nSaldoTopUpMember.Default
  
  ClearTdbgrid
  TDBGrid1.Columns(7).FooterText = ""
  TDBGrid1.Columns(4).FooterText = ""
  vaPubReff.ReDim 0, 2, 0, 2
  Label1.Caption = "Saldo Piutang Terikini (update) per tgl " & Format(Date, "dd-MM-yyyy")
End Sub

Private Sub ClearTdbgrid()
  vaArray.ReDim 0, -1, 0, 7
  TDBGrid1.Array = vaArray
  TDBGrid1.ReBind
End Sub

Private Function ValidSaving() As Boolean
  ValidSaving = True
  If Not CheckData(cFaktur.Text, "Nomor Faktur harus diisi...?") Then
    ValidSaving = False
    cFaktur.SetFocus
    Exit Function
  End If
  
  If Not CheckData(cCustomer.Text, "Kode Customer harus diisi...?") Then
    ValidSaving = False
    cCustomer.SetFocus
    Exit Function
  End If
End Function

Private Function validOK() As Boolean
  validOK = True
End Function

Private Sub cmdPerbaikan_Click()
Dim a As New exportExcel
Dim cSQL As String
Dim n As Single
Dim lSave As Boolean

cSQL = " select t.nomorpelunasanpiutang,tt.nomorpenjualan,t.kodeanggota,a.nama,t.tgl,tt.total as totalan from totpelunasanpiutang t"
cSQL = cSQL & " LEFT JOIN anggota a on a.kodeanggota = t.kodeanggota"
cSQL = cSQL & " LEFT JOIN pelunasanpiutang p on p.nomorpelunasanpiutang = t.nomorpelunasanpiutang"
cSQL = cSQL & " LEFT JOIN totpenjualan tt on tt.kodeanggota = t.kodeanggota"
cSQL = cSQL & " Where t.tgl >= '2012-11-01' And p.Pelunasan Is Null And tt.flaglunas <> 1"
cSQL = cSQL & " ORDER BY t.nomorpelunasanpiutang"

vaArray.ReDim 0, -1, 0, 4
Set dbData = objData.SQL(GetDSN, cSQL)
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = GetNull(dbData!nomorpelunasanpiutang)
      vaArray(n, 1) = GetNull(dbData!nomorpenjualan)
      vaArray(n, 2) = GetNull(dbData!kodeanggota)
      vaArray(n, 3) = GetNull(dbData!nama)
      vaArray(n, 4) = GetNull(dbData!totalan)
      dbData.MoveNext
    Loop
    
    If MsgBox("UPS, ada yg salah dalam proses pelunasan, apakah akan dilihat?", vbYesNo + vbCritical) = vbYes Then
      a.RecordSource = vaArray
      a.ExportToExcel
    End If
    
    lSave = True
    objData.Start GetDSN

    If MsgBox("Apakah akan diperbaki?", vbYesNo + vbInformation) = vbYes Then
      For n = vaArray.LowerBound(1) To vaArray.UpperBound(1)
        lSave = IIf(lSave, objData.Add(GetDSN, "pelunasanpiutang", _
        Array("nomorpelunasanpiutang", "nomorpenjualan", "piutang", "discount", "pelunasan"), _
        Array(vaArray(n, 0), vaArray(n, 1), vaArray(n, 4), 0, vaArray(n, 4))), False)
        
        lSave = IIf(lSave, objData.Edit(GetDSN, "penjualan", "nomorpenjualan = '" & vaArray(n, 1) & "'", Array("statuslunas"), Array(1)), False)
        lSave = IIf(lSave, objData.Edit(GetDSN, "totpenjualan", "nomorpenjualan = '" & vaArray(n, 1) & "'", Array("flaglunas"), Array(1)), False)
      Next n
      
      If lSave Then
        objData.Save GetDSN
        MsgBox "Ok, data sudah selesai diperbaiki", vbInformation
      Else
        objData.Cancel GetDSN
        MsgBox "Maaf terjadi kesalahan dalam proses perbaikan, data tidak jadi diperbaiki", vbInformation
      End If
    
    End If
    
  Else
    MsgBox "Hore.. tidak ditemukan satupun yg salah dalam proses pelunasan" & vbCrLf & "SISTEM OK", vbInformation
  End If
  
End Sub

Private Sub cmdPrint_Click()
  trPrint2.noOrder = TDBGrid1.Columns(2).Text
  Set dbData = objData.Browse(GetDSN, "totpenjualan t", "t.*,a.nama,a.telp", "t.nomorpenjualan", sisAssign, TDBGrid1.Columns(2).Text, , , Array("left join anggota a on a.kodeanggota = t.kodeanggota"))
  If Not dbData.EOF Then
    trPrint2.nSubTotal = GetNull(dbData!Subtotal)
    trPrint2.nDiscount = GetNull(dbData!dp)
    trPrint2.nCash = GetNull(dbData!Tunai)
    trPrint2.nChange = GetNull(dbData!Piutang)
    trPrint2.cKodeMember = GetNull(dbData!kodeanggota)
    trPrint2.cMember = GetNull(dbData!nama)
    trPrint2.cTeleponMember = GetNull(dbData!telp)
    trPrint2.Ups = GetNull(dbData!upkepada)
    trPrint2.dTgNota = Format(GetNull(dbData!tgl), "dd/MM/yyyy")
    trPrint2.dJthTempoNota = Format(GetNull(dbData!jthtmp), "dd/MM/yyyy")
    Load trPrint2
    trPrint2.Show vbModal
  End If
End Sub

Private Sub cmdSimpan_Click()
Dim n As Single
Dim lSave As Boolean
Dim Faktur As String


  lSave = True
  objData.Start GetDSN
  lStart = True
  Faktur = cFaktur.Text
  If nPos = Add Then
    If Not GetAvailable(cFaktur.Text, "totpelunasanpiutang", "nomorpelunasanpiutang") Then
      Faktur = GetNomor("totpelunasanpiutang", "nomorpelunasanpiutang", GetID, sisModulTransaksi.pelunasanpiutang)
    End If
  End If
  
  'cek apakah ada data yg akan dilunasi
  If Trim(cFaktur.Text) = "" Then
    MsgBox "Maaf Nomor Faktur Kosong/Tidak Valid" & vbCrLf & "Data tidak bisa disimpan", vbCritical
    Exit Sub
  End If
  
  If GetCekCentang = False Then
    MsgBox "Maaf tidak ada data untuk di proses", vbCritical
    Exit Sub
  End If
  
  If nTotal.value > 0 Then 'jika total > 0 maka proses
  
' please load form pelunasan

  trLunasPiutang.nTotalYangHarusDibayar.value = nLunas.value
  trLunasPiutang.nTunai.value = nLunas.value
  trLunasPiutang.Label3.Caption = cFaktur.Text
  trLunasPiutang.cKodeAnggota.Text = cCustomer.Text
  trLunasPiutang.cNamaAnggota.Text = cNama.Text
  
  If nSaldoTopUpMember.value > 0 Then
    trLunasPiutang.opt(2).value = True
  Else
    trLunasPiutang.opt(0).value = True
  End If
    
  Load trLunasPiutang
  trLunasPiutang.Show vbModal
  cPubAkun = trLunasPiutang.cAkunKas.Text
  If lClose = True Then
    objData.Cancel GetDSN
    Exit Sub
  End If
  
  'simpan di tabel totpenjualan dan penjualan
  
  lSave = IIf(lSave, objData.Delete(GetDSN, "totpelunasanpiutang", "nomorpelunasanpiutang", sisAssign, Faktur), False)
  lSave = IIf(lSave, objData.Delete(GetDSN, "pelunasanpiutang", "nomorpelunasanpiutang", sisAssign, Faktur), False)
  lSave = IIf(lSave, objData.Delete(GetDSN, "kartupiutang", "nomorkartupiutang", sisAssign, Faktur), False)
  lSave = IIf(lSave, objData.Delete(GetDSN, "membertopup", "nomormembertopup ", sisAssign, Faktur), False)
  
  lSave = IIf(lSave, objData.Update(GetDSN, "totpelunasanpiutang", "nomorpelunasanpiutang = '" & Faktur & "'", Array("nomorpelunasanpiutang", "kodeanggota", "tgl", "discount", "total", "pelunasan", "datetime", "username", "kodeakun", "kodecostcenter"), Array(Faktur, cCustomer.Text, Format(dTanggal.value, "yyyy-MM-dd"), nDiscount.value, nTotal.value, nLunas.value, SNow, GetRegistry(reg_Username), cPubAkun, GetCostCenterUser(objData, GetRegistry(reg_Username)))), False)
  For n = 0 To vaArray.UpperBound(1)
    If vaArray(n, 0) = -1 Then
      lSave = IIf(lSave, objData.Add(GetDSN, "pelunasanpiutang", Array("nomorpelunasanpiutang", "nomorpenjualan", "piutang", "discount", "pelunasan"), Array(Faktur, vaArray(n, 2), vaArray(n, 4), vaArray(n, 6), vaArray(n, 7))), False)
      'Set status lunas pada table penjualan
      'Cek dulu apakah faktur penjualan ini sudah dilunasi apa belum
      If lCekStatusLunas(objData, vaArray(n, 2)) = True Then
        lSave = IIf(lSave, objData.Edit(GetDSN, "penjualan", "nomorpenjualan = '" & vaArray(n, 2) & "'", Array("statuslunas"), Array(1)), False)
        lSave = IIf(lSave, objData.Edit(GetDSN, "totpenjualan", "nomorpenjualan = '" & vaArray(n, 2) & "'", Array("flaglunas"), Array(1)), False)
      Else
        lSave = IIf(lSave, objData.Edit(GetDSN, "penjualan", "nomorpenjualan = '" & vaArray(n, 2) & "'", Array("statuslunas"), Array(0)), False)
        lSave = IIf(lSave, objData.Edit(GetDSN, "totpenjualan", "nomorpenjualan = '" & vaArray(n, 2) & "'", Array("flaglunas"), Array(0)), False)
      End If
    End If
    
    If nPos <> Add Then
      If vaArray(n, 0) <> -1 Then
        lSave = IIf(lSave, objData.Edit(GetDSN, "penjualan", "nomorpenjualan = '" & vaArray(n, 2) & "'", Array("statuslunas"), Array(0)), False)
      End If
    End If
  Next n
  
  lSave = IIf(lSave, UpdKartuHutang(objData, SisKartuHutang.SisPelunasanPiutang, Faktur, dTanggal.value, cCustomer.Text, "Pelunasan Piutang an " & cNama.Text, nTotal.value - nDiscount.value), False)
  lSave = IIf(lSave, UpdKartuHutang(objData, SisKartuHutang.SisDiscountPelunasanPiutang, Faktur, dTanggal.value, cCustomer.Text, "Discount Pelunasan Piutang an " & cNama.Text, nDiscount.value), False)
  
  'CEK
  'tidak boleh dilakukan penghapusan atau pengkoreksian apabila faktur bg yg bersangkutang sudah dilunasi
  Set dbData = objData.Browse(GetDSN, "pencairanbg", , "nomorpelunasanpiutang", sisAssign, Faktur)
  If Not dbData.EOF Then
    MsgBox "Maaf transaksi ini tidak bisa dikoreksi kembali. BG/Cek sudah dicairkan"
    objData.Cancel GetDSN
    Exit Sub
  End If
  
  'Simpan di table BG
  'hapus dulu record yg pernah ada
      
  Dim i As Integer
  lSave = IIf(lSave, objData.Delete(GetDSN, "bg", "nomorpelunasanpiutang", sisAssign, Faktur), False)
  For i = vaPubReff.LowerBound(1) To vaPubReff.UpperBound(1)
    If vaPubReff(i, 1) <> 0 Then
      lSave = IIf(lSave, objData.Add(GetDSN, "bg", Array("nomorpelunasanpiutang", "reff", "jumlah", "jatuhtempo"), Array(Faktur, vaPubReff(i, 0), vaPubReff(i, 1), vaPubReff(i, 2))), False)
    End If
  Next i
  
  'Jurnal
  'Kas
  'diskon
  '   Piutang
  
  lSave = IIf(lSave, DelKodeTr(objData, msPelunasanPiutang, Faktur), False)
  
  'akun kas'
  If trLunasPiutang.opt(0).value = True Then
    'Jika pembayaran tunai maka lawannya Kas
    lSave = IIf(lSave, UpdKodeTr(objData, msPelunasanPiutang, Faktur, Format(dTanggal.value, "yyyy-MM-dd"), cPubAkun, GetCostCenterUser(objData, GetRegistry(reg_Username)), "Pelunasan piutang an " & cNama.Text, nLunas.value, 0), False)
  ElseIf trLunasPiutang.opt(1).value = True Then
    'Jika pembayaran BG maka lawannya akun BG
    lSave = IIf(lSave, UpdKodeTr(objData, msPelunasanPiutang, Faktur, Format(dTanggal.value, "yyyy-MM-dd"), aCfg(objData, msRekeningBG), GetCostCenterUser(objData, GetRegistry(reg_Username)), "Pelunasan piutang an " & cNama.Text, nLunas.value, 0), False)
  End If
  
  lSave = IIf(lSave, UpdKodeTr(objData, msPelunasanPiutang, Faktur, Format(dTanggal.value, "yyyy-MM-dd"), aCfg(objData, msRekeningPotonganPiutang), GetCostCenterUser(objData, GetRegistry(reg_Username)), "Pelunasan piutang an " & cNama.Text, nDiscount.value, 0), False)
      'Akun Piutang'
      lSave = IIf(lSave, UpdKodeTr(objData, msPelunasanPiutang, Faktur, Format(dTanggal.value, "yyyy-MM-dd"), GetAkunMember(objData, cCustomer.Text), GetCostCenterUser(objData, GetRegistry(reg_Username)), "Pelunasan piutang an " & cNama.Text, 0, nTotal.value), False)
  
  If trLunasPiutang.opt(2).value = True Then
    'pelunasan dengan menggunakan topup
    'simpan di buku besar
    'simpan di tabel membertopup
    
    Dim vaField
    Dim vaValue
        
    vaField = Array("nomormembertopup", "tgl", "kodeanggota", "keterangan", "kredit")
    
    If lTarikTunai = True Then
    
      'bagi dua pembyaran
      
      vaValue = Array(Faktur, dTanggal.value, cCustomer.Text, "Pelunasan Piutang via Top Up an " & cNama.Text, nSaldoTopUp - nJaminan - nTarikTunai)
      lSave = IIf(lSave, objData.Add(GetDSN, "membertopup", vaField, vaValue), False)

      If nTarikTunai <> 0 Then
        vaValue = Array(Faktur, dTanggal.value, cCustomer.Text, "Tarik Tunai Dana Top Up an " & cNama.Text, nTarikTunai)
        lSave = IIf(lSave, objData.Add(GetDSN, "membertopup", vaField, vaValue), False)
      End If
      
      lSave = IIf(lSave, UpdKodeTr(objData, msMemberTopUp, Faktur, Format(dTanggal.value, "yyyy-MM-dd"), GetAkunKas(objData, GetRegistry(reg_Username)), "", "Tarik tunai Top Up an " & cNama.Text, 0, nTarikTunai), False)
      lSave = IIf(lSave, UpdKodeTr(objData, msMemberTopUp, Faktur, Format(dTanggal.value, "yyyy-MM-dd"), aCfg(objData, msRekeningTopUp), "", "Tarik tunai Top Up an " & cNama.Text, nTarikTunai, 0), False)
      
      lSave = IIf(lSave, UpdKodeTr(objData, msMemberTopUp, Faktur, Format(dTanggal.value, "yyyy-MM-dd"), aCfg(objData, msRekeningTopUp), "", "Pelunasan via Top Up an " & cNama.Text, nWithDraw - nTarikTunai - nJaminan, 0), False)
      
    ElseIf lTarikTunai = False Then
      
'      If nJaminan <> 0 Then
'        vaField = Array("nomormembertopup", "tgl", "kodeanggota", "keterangan", "kredit")
'        vaValue = Array(Faktur, dTanggal.Value, cCustomer.Text, "Pelunasan Piutang Lewat Top Up an " & cNama.Text, nSaldoTopUp - nJaminan + nSisaKurangTopUp)
'      Else
'        vaValue = Array(Faktur, dTanggal.Value, cCustomer.Text, "Pelunasan Piutang Lewat Top Up an " & cNama.Text, nSaldoTopUp - nJaminan - nTarikTunai)
'      End If
      
'      If nSaldoTopUp - nTotYgHarusDibayar > 0 Then
'        vaField = Array("nomormembertopup", "tgl", "kodeanggota", "keterangan", "kredit")
'        vaValue = Array(Faktur, dTanggal.Value, cCustomer.Text, "Pelunasan via top up an " & cNama.Text, nTotYgHarusDibayar)
'        lSave = IIf(lSave, objData.Add(GetDSN, "membertopup", vaField, vaValue), False)
'      End If
      
      
      vaField = Array("nomormembertopup", "tgl", "kodeanggota", "keterangan", "kredit")
      'jika yg tagihan lebih dari atau = top up
      
      If nSaldoTopUp < nTotYgHarusDibayar Then
        vaValue = Array(Faktur, dTanggal.value, cCustomer.Text, "Pembayaran via top up an " & cNama.Text, nSaldoTopUp - nJaminan)
        lSave = IIf(lSave, objData.Add(GetDSN, "membertopup", vaField, vaValue), False)
      Else
        vaValue = Array(Faktur, dTanggal.value, cCustomer.Text, "Pembayaran via top up an " & cNama.Text, nTotYgHarusDibayar)
        lSave = IIf(lSave, objData.Add(GetDSN, "membertopup", vaField, vaValue), False)
      End If
      
      'jika tagihan kurang dari
      
'      If nSisaKurangTopUp <> 0 Then
'        vaField = Array("nomormembertopup", "tgl", "kodeanggota", "keterangan", "kredit")
'        vaValue = Array(Faktur, dTanggal.Value, cCustomer.Text, "Jaminan an " & cNama.Text, nSisaKurangTopUp)
'        lSave = IIf(lSave, objData.Add(GetDSN, "membertopup", vaField, vaValue), False)
'      End If
            
      If nTarikTunai <> 0 Then
        vaField = Array("nomormembertopup", "tgl", "kodeanggota", "keterangan", "kredit")
        vaValue = Array(Faktur, dTanggal.value, cCustomer.Text, "Jaminan an " & cNama.Text, nTarikTunai)
        lSave = IIf(lSave, objData.Add(GetDSN, "membertopup", vaField, vaValue), False)
      End If
      
      'debet = jaminan
      'kredit = sisa kurang bayar
      
      
      If nSaldoTopUp < nTotYgHarusDibayar Then
        lSave = IIf(lSave, UpdKodeTr(objData, msMemberTopUp, Faktur, Format(dTanggal.value, "yyyy-MM-dd"), aCfg(objData, msRekeningTopUp), "", "Pelunasan via Top Up an " & cNama.Text, nSaldoTopUp - nJaminan, 0), False)
      Else
        lSave = IIf(lSave, UpdKodeTr(objData, msMemberTopUp, Faktur, Format(dTanggal.value, "yyyy-MM-dd"), aCfg(objData, msRekeningTopUp), "", "Pelunasan via Top Up an " & cNama.Text, nTotYgHarusDibayar, 0), False)
      End If
      
'      If nSisaKurangTopUp <> 0 Then
'        lSave = IIf(lSave, UpdKodeTr(objData, msMemberTopUp, Faktur, Format(dTanggal.Value, "yyyy-MM-dd"), aCfg(objData, msRekeningTopUp), "", "Pelunasan via Top Up an " & cNama.Text, nSaldoTopUp - nJaminan, 0), False)
'      Else
'        lSave = IIf(lSave, UpdKodeTr(objData, msMemberTopUp, Faktur, Format(dTanggal.Value, "yyyy-MM-dd"), aCfg(objData, msRekeningTopUp), "", "Pelunasan via Top Up an " & cNama.Text, nWithDraw - nTarikTunai - nSisaKurangTopUp, 0), False)
'      End If
      
    End If
    
    lSave = IIf(lSave, UpdKodeTr(objData, msMemberTopUp, Faktur, Format(dTanggal.value, "yyyy-MM-dd"), GetAkunKas(objData, GetRegistry(reg_Username)), "", "Bayar Sisa Pelunasan " & cNama.Text, nSisaKurangTopUp, 0), False)
    
  End If
  
  
  If lSave Then
    objData.Save GetDSN
  Else
    MsgBox "Maaf data tidak berhasil disimpan", vbExclamation
    objData.Cancel GetDSN
  End If
    
  
  'Lakukan pengecekan apakah semua data disimpan sesuai dengan tabel yg dituju
  'langkah awal:
  
'  Set dbData = objData.Browse(GetDSN, "totpelunasanpiutang", , "nomorpelunasanpiutang", sisAssign, Faktur)
'  If Not dbData.EOF Then
'    'cek apakah jumlah nya sudah sesuai dengan yg ada di tabel pelunasan piutang
'
'  End If
'
'  For i = 1 To aCfg(objData, msJumlahCetakanPenjualanNonTunai)
'    trPrintPelunasanPiutang.noOrder = Faktur
'    Set dbData = objData.Browse(GetDSN, "totpelunasanpiutang t", "t.*,a.*", "t.nomorpelunasanpiutang", sisAssign, Faktur, , , Array("left join anggota a on a.kodeanggota = t.kodeanggota"))
'    If Not dbData.EOF Then
'      trPrintPelunasanPiutang.nSubTotal = GetNull(dbData!Total)
'      trPrintPelunasanPiutang.nDiscount = 0
'      trPrintPelunasanPiutang.nCash = 0
'      trPrintPelunasanPiutang.nChange = 0
'      trPrintPelunasanPiutang.cKodeMember = GetNull(dbData!kodeanggota)
'      trPrintPelunasanPiutang.cMember = GetNull(dbData!nama)
'      trPrintPelunasanPiutang.cTeleponMember = GetNull(dbData!telp)
'      trPrintPelunasanPiutang.Ups = 0
'
'      trPrintPelunasanPiutang.nKembali1 = nTarikTunai 'berapa uang yg ditarik
'      trPrintPelunasanPiutang.nSaldoTopUp = nSaldoTopUp 'saldo top up
'      trPrintPelunasanPiutang.nSisa = nSisaKurangTopUp 'kurang
'      trPrintPelunasanPiutang.nTunai = nTunai
'      trPrintPelunasanPiutang.nKembali2 = nKembalian
'      trPrintPelunasanPiutang.lKembali = lTarikTunai
'      trPrintPelunasanPiutang.nMetodePembayaran = nMetodePembayaran
'      trPrintPelunasanPiutang.nPoinHadiah = nPoinReguler.Value
'
'      Load trPrintPelunasanPiutang
'      trPrintPelunasanPiutang.Show vbModal
'
'    End If
'  Next i
  
  
  Unload trLunasPiutang
  
  GetCetakPelunasan objData, Faktur
  GetEdit False
  initvalue
  Else
    MsgBox "Maaf tidak ada data untuk di proses", vbExclamation
  End If 'end if dari ntotal.value
  
End Sub

Private Function GetCekCentang() As Boolean
Dim n As Integer

  GetCekCentang = False
  For n = 0 To vaArray.UpperBound(1)
    If vaArray(n, 0) = -1 Then
      GetCekCentang = True
      Exit For
    End If
  Next n
End Function

'Private Sub PrintStruk(ByVal Faktur As String)
'Dim n As Double
'Dim nBruto As Double
'Dim nTotQty As Double
'
'  With aMainmenu.IO1
'    .Open GetRegistry(reg_PortPrinterKasir), ""
'    .WriteString Chr(27) & Chr(15) & vbCrLf
'    .WriteString Padc(Trim("STRUK PELUNASAN"), 40) & vbCrLf
'    .WriteString Padc(Trim(aCfg(objData, msNamaPerusahaan)), 40) & vbCrLf
'    .WriteString Padc(Trim(aCfg(objData, msAlamatPerusahaan)), 40) & vbCrLf
'    .WriteString Padc(Trim(aCfg(objData, msTelepon) & " " & aCfg(objData, msFax)), 40) & vbCrLf
'    .WriteString Padc(aCfg(objData, msKota), 40) & vbCrLf
'    .WriteString Replicate("-", 40) & vbCrLf
'    .WriteString "Member. " & cCustomer.Text & " " & cNama.Text & vbCrLf
'    .WriteString Replicate("-", 40) & vbCrLf
'    .WriteString "Print By: " & aCfg(objData, msNama) & vbCrLf
'
''    trPrintPelunasanPiutang.nSubTotal = GetNull(dbData!Total)
''    trPrintPelunasanPiutang.nDiscount = 0
''    trPrintPelunasanPiutang.nCash = 0
''    trPrintPelunasanPiutang.nChange = 0
''    trPrintPelunasanPiutang.cKodeMember = GetNull(dbData!kodeanggota)
''    trPrintPelunasanPiutang.cMember = GetNull(dbData!nama)
''    trPrintPelunasanPiutang.cTeleponMember = GetNull(dbData!telp)
''    trPrintPelunasanPiutang.Ups = 0
''
''    trPrintPelunasanPiutang.nKembali1 = nTarikTunai 'berapa uang yg ditarik
''    trPrintPelunasanPiutang.nSaldoTopUp = nSaldoTopUp 'saldo top up
''    trPrintPelunasanPiutang.nSisa = nSisaKurangTopUp 'kurang
''    trPrintPelunasanPiutang.nTunai = nTunai
''    trPrintPelunasanPiutang.nKembali2 = nKembalian
''    trPrintPelunasanPiutang.lKembali = lTarikTunai
''    trPrintPelunasanPiutang.nMetodePembayaran = nMetodePembayaran
'
'    .WriteString Padl("Tarik        : " & Padl(Format(nTarikTunai, "###,###,##0"), 11), 40) & vbCrLf
'    .WriteString Padl("S.Top Up     : " & Padl(Format(nSaldoTopUp, "###,###,##0"), 11), 40) & vbCrLf
'    .WriteString Padl("Sisa Kurang  : " & Padl(Format(nSisaKurangTopUp, "###,###,##0"), 11), 40) & vbCrLf
'    .WriteString Padl("Tunai        : " & Padl(Format(nTunai, "###,###,##0"), 11), 40) & vbCrLf
'    .WriteString Padl("Kembalian    : " & Padl(Format(nKembalian, "###,###,##0"), 11), 40) & vbCrLf
'
''    For n = 0 To vaArray.UpperBound(1)
''      If vaArray(n, 3) <> 0 Then
''        vaArray(n, 5) = vaArray(n, 5) - (vaArray(n, 5) * vaArray(n, 6) / 100)
''        nBruto = nBruto + (vaArray(n, 3) * vaArray(n, 5))
''        nTotQty = nTotQty + vaArray(n, 3)
''        .WriteString Left(vaArray(n, 2), 40) & vbCrLf
''        .WriteString Padl(Left(vaArray(n, 1), 10) & Padl(Format(vaArray(n, 3), "#,##0.00"), 8) & " x" & Padl(Format(vaArray(n, 5), "#,###,##0"), 9) & " =" & Padl(Format(vaArray(n, 3) * vaArray(n, 5), "#,###,##0"), 9), 40) & vbCrLf
''      End If
''    Next
''    .WriteString Replicate("-", 40) & vbCrLf
''    .WriteString Padr("=> " & Format(nTotQty, "###,###,##0.00") & " Items", 20) & Padl("Sub   : " & Padl(Format(nBruto, "###,###,##0"), 11), 20) & vbCrLf
''
''    .WriteString Padl("Disc  : " & Padl(Format(nDiscount.Value, "###,###,##0"), 11), 40) & vbCrLf
''    .WriteString Padl("Total : " & Padl(Format(nTotal.Value, "###,###,##0"), 11), 40) & vbCrLf
''    .WriteString Padl("DP    : " & Padl(Format(nDP.Value, "###,###,##0"), 11), 40) & vbCrLf
''    .WriteString Padl("Bayar Tunai .... " & Padl(Format(nTunai.Value, "###,###,##0"), 11), 40) & vbCrLf
''    .WriteString Padl("Hutang .... " & Padl(Format(nPiutang.Value, "###,###,##0"), 11), 40) & vbCrLf
''    .WriteString Replicate("-", 40) & vbCrLf
''
''    .WriteString "No. " & Faktur & Padl(Format(Now, "dd-MM-yyyy HH:MM:SS"), 22) & vbCrLf
''
''    .WriteString "Print by " & Padl(GetRegistry(reg_UserName), 26) & vbCrLf
''    .WriteString Replicate("-", 40) & vbCrLf
'
'
'    .Close
'    OpenDrawer GetRegistry(reg_PortPrinterKasir)
'  End With
'End Sub


Private Sub cNama_ButtonClick()
  
  nTotal.Default
  nDiscount.Default
  nLunas.Default
  
  Set dbData = objData.Browse(GetDSN, "anggota a", "a.kodeanggota,a.nama,a.alamat,d.keterangan as namadep", "a.nama", sisContent, cNama.Text, , , Array("left join dep d on d.kodedep = a.kodedep"))
  If Not dbData.EOF Then
    cNama.Text = cNama.Browse(dbData)
    cCustomer.Text = GetNull(dbData!kodeanggota)
    cNama.Text = GetNull(dbData!nama)
    cNamaDepartmen.Text = GetNull(dbData!namadep)
    cAlamat.Text = GetNull(dbData!alamat)
    If nPos = Add Then
      GetData
    End If
  End If
End Sub

Private Sub cNama_Validate(Cancel As Boolean)
  cNama.Enabled = False
End Sub

Private Sub dTanggal_Validate(Cancel As Boolean)
  If Not IsInPeriod(dTanggal.value) Or (dTanggal.value > Date) Then
    Cancel = True
    dTanggal.SetFocus
    GetEdit False
  End If
End Sub

Private Sub Form_Activate()
  If nPos = Add Then
    GetData
  End If
End Sub

Private Sub Form_Load()
Dim n As Single

  SetIcon Me.hWnd, "SIKD"
  initvalue
  GetEdit False
  CenterForm Me
  
  Set dbData = objData.Browse(GetDSN, "costcenter", , "kodecostcenter", sisAssign, GetCostCenterUser(objData, GetRegistry(reg_Username)))
  If Not dbData.EOF Then
    lbCostCenter.Caption = "Cost Centre : " & GetNull(dbData!keterangan)
  End If
  
  TabIndex dTanggal, n
  TabIndex cCustomer, n
  TabIndex cNama, n
  TabIndex cFaktur, n
  TabIndex cAkunKas, n
  
  TabIndex cmdAdd, n
  TabIndex cmdEdit, n
  TabIndex cmdHapus, n
  TabIndex cmdSimpan, n
  TabIndex cmdKeluar, n
End Sub

Private Sub SumTDB()
Dim n As Integer
Dim Discount As Double
Dim Pelunasan As Double
Dim nPoinReg As Double
  
  
  For n = 0 To vaArray.UpperBound(1)
    If vaArray(n, 0) = -1 Then
      Discount = Discount + vaArray(n, 6)
      Pelunasan = Pelunasan + vaArray(n, 7)
      If CDbl(vaArray(n, 4)) = CDbl(vaArray(n, 7)) Then
        nPoinReg = getPoinReguler(vaArray(n, 2)) + nPoinReg
      End If
    End If
  Next n
  
  nTotal.value = Pelunasan + Discount
  nDiscount.value = Discount
  nLunas.value = nTotal.value - nDiscount.value
  nPoinReguler.value = nPoinReg \ aCfg(objData, msKelipatan)
End Sub

Private Sub TDBGrid1_AfterColUpdate(ByVal ColIndex As Integer)
  TDBGrid1.Update
  SumTDB
'  MsgBox TDBGrid1.Columns(0).Value & " - " & TDBGrid1.Columns(1).Value
End Sub

Private Sub tdbgrid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
Dim nSisaPiutang As Double
  
  isLunas objData, TDBGrid1.Columns(2).value, nSisaPiutang
  
  If Not IsNumeric(TDBGrid1.Columns(6).value) Or Not IsNumeric(TDBGrid1.Columns(7).value) Or TDBGrid1.Columns(6).value < 0 Then
    Cancel = True
    Exit Sub
  End If

  If ColIndex = 0 Or ColIndex = 6 Or ColIndex = 7 Then
    If TDBGrid1.Columns(0).value = -1 Then
      If ColIndex <> 7 Then
        TDBGrid1.Columns(7).value = TDBGrid1.Columns(4).value - TDBGrid1.Columns(6).value
      Else
        TDBGrid1.Columns(6).value = 0
      End If
'    Else
'      TDBGrid1.Columns(7).Value = 0
'      TDBGrid1.Columns(6).Value = 0
    End If
  Else
    Cancel = True
  End If
  
  TDBGrid1.Refresh
  'pelunasan piutang tidak boleh lebih dari sisa piutang
  If CDbl(TDBGrid1.Columns(7).value) > CDbl(TDBGrid1.Columns(4).value) Then
    MsgBox "Maaf, nilai pelunasan tidak boleh melebihi dari sisa piutang faktur" & vbCrLf & "Silahkan ulangi pengisian. Terimakasih"
    TDBGrid1.Refresh
    Cancel = True
  End If
  
  'lakukan perhitungan poin reguler
  nPoinReguler.value = getPoinReguler(TDBGrid1.Columns(2).value)
End Sub

'Private Sub TDBGrid1_DblClick()
'  GetCetakFakturpenjualan objData, TDBGrid1.Columns(2).Text, False
'End Sub

Private Function getPoinReguler(ByVal cFakturBelanja As String) As Double
Dim cSQL As String
Dim dbD As New ADODB.Recordset
'Fungsi ini akan mengambil poin reguler dari setiap nota
'karena tidak setiap nota berisi poin reguler (bercampur promo)

  cSQL = "select sum(harga*qty) as poinBelanja from penjualan p " & _
  "where nomorpenjualan = '" & cFakturBelanja & "' AND p.tgl >= '" & Format(DateAdd("D", -aCfg(objData, msTerm), Format(Date, "yyyy-MM-dd")), "yyyy-MM-dd") & "'"
 
  Set dbD = objData.SQL(GetDSN, cSQL)
  If Not dbD.EOF Then
    Do While Not dbD.EOF
      getPoinReguler = GetNull(dbD!poinBelanja)
      dbD.MoveNext
    Loop
  End If
End Function


Sub GetCetakPelunasan(ByVal obj As CodeSuiteLibrary.Data, ByVal Faktur As String)
Dim n As Integer
Dim cTerbilang As String
Dim cField As String
Dim vaJoin
Dim vaGrid As New XArrayDB
Dim cHead As String
Dim cSQL As String

  cSQL = ""
  cSQL = cSQL & " select p.nomorpelunasanpiutang,p.nomorpenjualan,p.piutang,p.pelunasan,t.jthtmp from pelunasanpiutang p"
  cSQL = cSQL & " LEFT JOIN totpenjualan t on t.nomorpenjualan = p.nomorpenjualan"
  cSQL = cSQL & " where p.nomorpelunasanpiutang = '" & Faktur & "'"

  Set dbData = obj.SQL(GetDSN, cSQL)
  If Not dbData.EOF Then
    n = 0
    vaGrid.ReDim 0, dbData.RecordCount - 1, 0, 5
    Do While Not dbData.EOF
       vaGrid(n, 0) = n + 1
       vaGrid(n, 1) = (dbData!nomorpenjualan)
       vaGrid(n, 2) = (dbData!Pelunasan)
       vaGrid(n, 3) = (dbData!jthtmp)
       vaGrid(n, 4) = 0
       vaGrid(n, 5) = vaGrid(n, 2) - vaGrid(n, 4)
       dbData.MoveNext
      n = n + 1
    Loop
    
    'AMBIL INFORMASI customer
    cSQL = ""
    cSQL = " select a.kodeanggota,a.nama,a.alamat,a.telp,t.nomorpelunasanpiutang,t.tgl,t.total from totpelunasanpiutang t"
    cSQL = cSQL & " LEFT JOIN anggota a on a.kodeanggota = t.kodeanggota where t.nomorpelunasanpiutang = '" & Faktur & "'"
    
    Set dbData = obj.SQL(GetDSN, cSQL)
    cTerbilang = "# " & Dec2Text(GetNull(dbData!Total)) & "Rupiah #"
    cHead = "Kuitansi Lunas"
    With rptKuitansiLunas
      .Parameters("dTgl").ValueExpression = "'" & Format(GetNull(dbData!tgl), "dd-MM-yyyy") & "'"
      .Parameters("cSE").ValueExpression = "'" & Faktur & "'"
      
      .Parameters("cNama").ValueExpression = "'" & GetNull(dbData!nama, "") & "'"
      .Parameters("cAlamat").ValueExpression = "'" & GetNull(dbData!alamat, "") & "'"
      .Parameters("cKota").ValueExpression = "'" & GetNull(dbData!telp) & "'"
      .Parameters("cKodeAnggota").ValueExpression = "'" & GetNull(dbData!kodeanggota, "") & "'"
      
      .Parameters("cTerbilang").ValueExpression = "'" & cTerbilang & "'"
      .Parameters("cTTD").ValueExpression = "'" & Padc(GetRegistry(reg_FullName), 45) & "'"
      .Parameters("cReceived").ValueExpression = "'" & Padc("", 45) & "'"
      
      .Parameters("nSubtotal").ValueExpression = GetNull(dbData!Total)
      .Parameters("nTotal").ValueExpression = GetNull(dbData!Total)
      .Parameters("cNamaPerusahaan").ValueExpression = "'" & aCfg(obj, msNamaPerusahaan) & "'"
      .Parameters("cAlamatPerusahaan").ValueExpression = "'Alamat : " & aCfg(obj, msAlamatPerusahaan) & " Telp/Fax " & aCfg(objData, msTelepon) & "/" & aCfg(objData, msFax) & "'"
      .Parameters("cUserName").ValueExpression = "'" & GetRegistry(reg_FullName) & "'"
      .Parameters("cJudul").ValueExpression = "'" & cHead & "'"
      .Parameters("keAkun").ValueExpression = "'" & GetNamaAkun(objData, cPubAkun) & "'"
      
      Set .Array = vaGrid
      .Refresh
      If MsgBox("Apakah cetakan mau dalam bentuk kertas A4?!!" & vbCrLf & "Jika tidak maka cetakan akan dalam bentuk 1/2 kertas kuarto", vbYesNo) = vbYes Then
        .Profiles(0).PrinterPaperSize = tdbPPS_A4
      End If
      .PrintPreview
    End With
  End If
End Sub

Private Function GetNamaAkun(ByVal obj As CodeSuiteLibrary.Data, ByVal kodeakun As String) As String
Dim db As New ADODB.Recordset
  GetNamaAkun = ""
  Set db = obj.Browse(GetDSN, "akun", , "kodeakun", sisAssign, kodeakun)
  If Not db.EOF Then
     GetNamaAkun = "Rekening Kas : " & kodeakun & " " & GetNull(db!keterangan)
  End If
End Function
