VERSION 5.00
Object = "{0D6235E7-DBA2-11D1-B5DF-0060976089D0}#1.0#0"; "tdbr6.ocx"
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form rptRekapPenjualanBelumLunas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rekap Tagihan Pembelian/Penjualan"
   ClientHeight    =   5460
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   6984
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   6984
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   4800
      Left            =   15
      Top             =   15
      Width           =   6930
      _ExtentX        =   12234
      _ExtentY        =   8467
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483633
      Begin VB.Frame Frame2 
         Caption         =   "Cetakan"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1470
         Left            =   885
         TabIndex        =   15
         Top             =   3255
         Width           =   4530
         Begin VB.OptionButton optLunasCetak 
            Caption         =   "&1 Belum Lunas"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   330
            TabIndex        =   18
            Top             =   1110
            Visible         =   0   'False
            Width           =   1470
         End
         Begin VB.OptionButton optLunasCetak 
            Caption         =   "&2 Lunas"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   1830
            TabIndex        =   17
            Top             =   1110
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.OptionButton optLunasCetak 
            Caption         =   "&3 Semuanya"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   2865
            TabIndex        =   16
            Top             =   1110
            Visible         =   0   'False
            Width           =   1470
         End
         Begin BiSATextBoxProject.BiSABrowse cNamaAnggota 
            Height          =   330
            Left            =   405
            TabIndex        =   19
            Top             =   555
            Width           =   1935
            _ExtentX        =   3408
            _ExtentY        =   593
            Text            =   "12345678"
            BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "Verdana"
            Appearance      =   0
            Button          =   -1  'True
            CaptionWidth    =   0
            BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin BiSATextBoxProject.BiSATextBox cKodeAnggota 
            Height          =   330
            Left            =   2370
            TabIndex        =   20
            Top             =   555
            Width           =   1905
            _ExtentX        =   3344
            _ExtentY        =   593
            BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "Verdana"
            BackColor       =   -2147483633
            Enabled         =   0   'False
            Appearance      =   0
            CaptionWidth    =   1400
            BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label3 
            Caption         =   "Nama"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.2
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   228
            Left            =   408
            TabIndex        =   22
            Top             =   288
            Width           =   528
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Rekapan"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   870
         TabIndex        =   8
         Top             =   1860
         Width           =   4530
         Begin VB.OptionButton optLunas 
            Caption         =   "&3 Semuanya"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   2865
            TabIndex        =   13
            Top             =   1005
            Width           =   1470
         End
         Begin VB.OptionButton optLunas 
            Caption         =   "&2 Lunas"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   1830
            TabIndex        =   12
            Top             =   1005
            Width           =   1035
         End
         Begin VB.OptionButton optLunas 
            Caption         =   "&1 Belum Lunas"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   312
            TabIndex        =   11
            Top             =   1005
            Width           =   1470
         End
         Begin BiSATextBoxProject.BiSABrowse cNamaGroupSales 
            Height          =   330
            Left            =   405
            TabIndex        =   9
            Top             =   555
            Width           =   1935
            _ExtentX        =   3408
            _ExtentY        =   593
            Text            =   "12345678"
            BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "Verdana"
            Appearance      =   0
            Button          =   -1  'True
            CaptionWidth    =   0
            BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin BiSATextBoxProject.BiSATextBox cKodeGroupSales 
            Height          =   330
            Left            =   2370
            TabIndex        =   10
            Top             =   555
            Width           =   1905
            _ExtentX        =   3344
            _ExtentY        =   593
            BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "Verdana"
            BackColor       =   -2147483633
            Enabled         =   0   'False
            Appearance      =   0
            CaptionWidth    =   1400
            BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label1 
            Caption         =   "Group Sales"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.2
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   420
            TabIndex        =   23
            Top             =   315
            Width           =   1155
         End
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&2 Rekap Pembelian"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   1095
         TabIndex        =   5
         Top             =   1065
         Width           =   2688
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&1 Rekap Penjualan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   1095
         TabIndex        =   4
         Top             =   210
         Width           =   2904
      End
      Begin BiSADateProject.BiSADate dTglJual 
         Height          =   324
         Index           =   1
         Left            =   3036
         TabIndex        =   2
         Top             =   600
         Width           =   1536
         _ExtentX        =   2709
         _ExtentY        =   572
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483640
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin BiSADateProject.BiSADate dTglJual 
         Height          =   324
         Index           =   0
         Left            =   1452
         TabIndex        =   3
         Top             =   600
         Width           =   1548
         _ExtentX        =   2731
         _ExtentY        =   572
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483640
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin BiSADateProject.BiSADate dTglBeli 
         Height          =   345
         Index           =   0
         Left            =   1440
         TabIndex        =   6
         Top             =   1365
         Width           =   1545
         _ExtentX        =   2731
         _ExtentY        =   614
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483640
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin BiSADateProject.BiSADate dTglBeli 
         Height          =   345
         Index           =   1
         Left            =   3000
         TabIndex        =   7
         Top             =   1365
         Width           =   1545
         _ExtentX        =   2731
         _ExtentY        =   614
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483640
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin TrueDBReports60Ctl.TDBReports rptKuitansiLunas 
         Height          =   570
         Left            =   4995
         TabIndex        =   14
         Top             =   420
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   1005
         Caption         =   "Kuitansi Lunas"
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
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
         Fields(1).Name  =   "JatuhTempo"
         Fields(1).DisplayName=   "JatuhTempo"
         Fields(1).Type  =   7
         Fields(2).Name  =   "NoInvoice"
         Fields(2).DisplayName=   "NoInvoice"
         Fields(3).Name  =   "Item"
         Fields(3).DisplayName=   "Item"
         Fields(4).Name  =   "Qty"
         Fields(4).DisplayName=   "Qty"
         Fields(5).Name  =   "Total"
         Fields(5).DisplayName=   "Total"
         Fields(5).Type  =   5
         Sections.Count  =   6
         Sections(0).Name=   "SECTION_2"
         Sections(0).Type=   1
         Sections(0).StyleExp=   "'Tdb_Base'"
         Sections(0).Cells.Count=   12
         Sections(0).Cells(0).Name=   "CELL_22"
         Sections(0).Cells(0).Exp=   "cNamaPerusahaan"
         Sections(0).Cells(0).NewLine=   -1  'True
         Sections(0).Cells(0).Width=   30
         Sections(0).Cells(0).PrivateStyle=   -1  'True
         Sections(0).Cells(0).Style.Name=   "<private>"
         Sections(0).Cells(0).Style.ParentName=   "Tdb_Base"
         Sections(0).Cells(0).Style.Font_Name=   "Courier"
         Sections(0).Cells(0).Style.Font_Size=   9.75
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
         Sections(0).Cells(1).Name=   "CELL_1"
         Sections(0).Cells(1).Exp=   """Tagihan"""
         Sections(0).Cells(1).Width=   30
         Sections(0).Cells(1).PrivateStyle=   -1  'True
         Sections(0).Cells(1).Style.Name=   "<private>"
         Sections(0).Cells(1).Style.ParentName=   "Tdb_Base"
         Sections(0).Cells(1).Style.Font_Name=   "Courier"
         Sections(0).Cells(1).Style.Font_Size=   9.75
         Sections(0).Cells(1).Style.Font_Bold=   -1  'True
         Sections(0).Cells(1).Style.Font_Italic=   0   'False
         Sections(0).Cells(1).Style.Font_Underline=   0   'False
         Sections(0).Cells(1).Style.Font_Strikeout=   0   'False
         Sections(0).Cells(1).Style.Font_Charset=   0
         Sections(0).Cells(1).Style.TextAlign=   1
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
         Sections(0).Cells(1).Style.HasBorders=   -1  'True
         Sections(0).Cells(1).Style.BorderHT=   ""
         Sections(0).Cells(1).Style.BorderHI=   ""
         Sections(0).Cells(1).Style.BorderHB=   ""
         Sections(0).Cells(1).Style.BorderVL=   ""
         Sections(0).Cells(1).Style.BorderVI=   ""
         Sections(0).Cells(1).Style.BorderVR=   ""
         Sections(0).Cells(1).Style.NoClipping=   0   'False
         Sections(0).Cells(1).Style.RTF=   0   'False
         Sections(0).Cells(1).Style.fprops=   1
         Sections(0).Cells(2).Name=   "CELL_21"
         Sections(0).Cells(2).Exp=   "cKodeAnggota"
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
         Sections(0).Cells(2).Style.fprops=   0
         Sections(0).Cells(3).Name=   "CELL_25"
         Sections(0).Cells(3).Exp=   "cAlamatPerusahaan"
         Sections(0).Cells(3).NewLine=   -1  'True
         Sections(0).Cells(3).Width=   30
         Sections(0).Cells(3).PrivateStyle=   -1  'True
         Sections(0).Cells(3).Style.Name=   "<private>"
         Sections(0).Cells(3).Style.ParentName=   "Tdb_Base"
         Sections(0).Cells(3).Style.Font_Name=   "Courier"
         Sections(0).Cells(3).Style.Font_Size=   9.75
         Sections(0).Cells(3).Style.Font_Bold=   0   'False
         Sections(0).Cells(3).Style.Font_Italic=   0   'False
         Sections(0).Cells(3).Style.Font_Underline=   0   'False
         Sections(0).Cells(3).Style.Font_Strikeout=   0   'False
         Sections(0).Cells(3).Style.Font_Charset=   0
         Sections(0).Cells(3).Style.TextAlign=   0
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
         Sections(0).Cells(3).Style.HasBorders=   0   'False
         Sections(0).Cells(3).Style.BorderHT=   ""
         Sections(0).Cells(3).Style.BorderHI=   ""
         Sections(0).Cells(3).Style.BorderHB=   ""
         Sections(0).Cells(3).Style.BorderVL=   ""
         Sections(0).Cells(3).Style.BorderVI=   ""
         Sections(0).Cells(3).Style.BorderVR=   ""
         Sections(0).Cells(3).Style.NoClipping=   0   'False
         Sections(0).Cells(3).Style.RTF=   0   'False
         Sections(0).Cells(3).Style.fprops=   22413317
         Sections(0).Cells(4).Name=   "CELL_2"
         Sections(0).Cells(4).Exp=   "cSE"
         Sections(0).Cells(4).Width=   30
         Sections(0).Cells(4).PrivateStyle=   -1  'True
         Sections(0).Cells(4).Style.Name=   "<private>"
         Sections(0).Cells(4).Style.ParentName=   "Tdb_Base"
         Sections(0).Cells(4).Style.Font_Name=   "Courier"
         Sections(0).Cells(4).Style.Font_Size=   9.75
         Sections(0).Cells(4).Style.Font_Bold=   -1  'True
         Sections(0).Cells(4).Style.Font_Italic=   0   'False
         Sections(0).Cells(4).Style.Font_Underline=   0   'False
         Sections(0).Cells(4).Style.Font_Strikeout=   0   'False
         Sections(0).Cells(4).Style.Font_Charset=   0
         Sections(0).Cells(4).Style.TextAlign=   1
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
         Sections(0).Cells(4).Style.fprops=   131073
         Sections(0).Cells(5).Name=   "CELL_26"
         Sections(0).Cells(5).Exp=   "cNama & "" ("" & cKota & "")"""
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
         Sections(0).Cells(5).Style.TextAlign=   0
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
         Sections(0).Cells(5).Style.BorderHB=   ""
         Sections(0).Cells(5).Style.BorderVL=   ""
         Sections(0).Cells(5).Style.BorderVI=   ""
         Sections(0).Cells(5).Style.BorderVR=   ""
         Sections(0).Cells(5).Style.NoClipping=   0   'False
         Sections(0).Cells(5).Style.RTF=   0   'False
         Sections(0).Cells(5).Style.fprops=   68419585
         Sections(0).Cells(6).Name=   "CELL_6"
         Sections(0).Cells(6).NewLine=   -1  'True
         Sections(0).Cells(6).Width=   30
         Sections(0).Cells(6).PrivateStyle=   -1  'True
         Sections(0).Cells(6).Style.Name=   "<private>"
         Sections(0).Cells(6).Style.ParentName=   "Tdb_Base"
         Sections(0).Cells(6).Style.Font_Name=   "Courier"
         Sections(0).Cells(6).Style.Font_Size=   9.75
         Sections(0).Cells(6).Style.Font_Bold=   -1  'True
         Sections(0).Cells(6).Style.Font_Italic=   0   'False
         Sections(0).Cells(6).Style.Font_Underline=   0   'False
         Sections(0).Cells(6).Style.Font_Strikeout=   0   'False
         Sections(0).Cells(6).Style.Font_Charset=   0
         Sections(0).Cells(6).Style.TextAlign=   3
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
         Sections(0).Cells(6).Style.fprops=   0
         Sections(0).Cells(7).Name=   "CELL_11"
         Sections(0).Cells(7).Width=   30
         Sections(0).Cells(7).PrivateStyle=   -1  'True
         Sections(0).Cells(7).Style.Name=   "<private>"
         Sections(0).Cells(7).Style.ParentName=   "Tdb_Base"
         Sections(0).Cells(7).Style.Font_Name=   "Courier"
         Sections(0).Cells(7).Style.Font_Size=   9.75
         Sections(0).Cells(7).Style.Font_Bold=   -1  'True
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
         Sections(0).Cells(7).Style.fprops=   0
         Sections(0).Cells(8).Name=   "CELL_8"
         Sections(0).Cells(9).Name=   "CELL_3"
         Sections(0).Cells(9).Exp=   "dTgl"
         Sections(0).Cells(9).NewLine=   -1  'True
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
         Sections(0).Cells(9).Style.TextAlign=   3
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
         Sections(0).Cells(9).Style.fprops=   22282240
         Sections(0).Cells(10).Name=   "CELL_7"
         Sections(0).Cells(10).Exp=   """Print by : "" & cUserName & "" "" & Now & "" Page : "" & PageNo()"
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
         Sections(0).Cells(10).Style.TextAlign=   2
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
         Sections(0).Cells(10).Style.fprops=   20971521
         Sections(0).Cells(11).Name=   "CELL_9"
         Sections(0).Cells(11).NewLine=   -1  'True
         Sections(0).Cells(11).Height=   4
         Sections(0).Cells(11).AutoHeight=   0   'False
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
         Sections(1).Cells(1).Exp=   """No Invoice"""
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
         Sections(1).Cells(2).Name=   "CELL_1"
         Sections(1).Cells(2).Exp=   """Jatuh Tempo"""
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
         Sections(1).Cells(3).Name=   "CELL_7"
         Sections(1).Cells(3).Exp=   """Item"""
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
         Sections(1).Cells(3).Style.BorderVL=   ""
         Sections(1).Cells(3).Style.BorderVI=   ""
         Sections(1).Cells(3).Style.BorderVR=   ""
         Sections(1).Cells(3).Style.NoClipping=   0   'False
         Sections(1).Cells(3).Style.RTF=   0   'False
         Sections(1).Cells(3).Style.fprops=   1
         Sections(1).Cells(4).Name=   "CELL_4"
         Sections(1).Cells(4).Exp=   """Qty"""
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
         Sections(2).Cells(1).Width=   20
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
         Sections(2).Cells(2).Name=   "CELL_4"
         Sections(2).Cells(2).Exp=   "JatuhTempo"
         Sections(2).Cells(2).Width=   14
         Sections(2).Cells(2).PrivateStyle=   -1  'True
         Sections(2).Cells(2).Style.Name=   "<private>"
         Sections(2).Cells(2).Style.ParentName=   "Tdb_Body"
         Sections(2).Cells(2).Style.Font_Name=   "Courier"
         Sections(2).Cells(2).Style.Font_Size=   9.75
         Sections(2).Cells(2).Style.Font_Bold=   0   'False
         Sections(2).Cells(2).Style.Font_Italic=   0   'False
         Sections(2).Cells(2).Style.Font_Underline=   0   'False
         Sections(2).Cells(2).Style.Font_Strikeout=   0   'False
         Sections(2).Cells(2).Style.Font_Charset=   0
         Sections(2).Cells(2).Style.TextAlign=   1
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
         Sections(2).Cells(3).Name=   "CELL_3"
         Sections(2).Cells(3).Exp=   "Item"
         Sections(2).Cells(3).Width=   12
         Sections(2).Cells(3).PrivateStyle=   -1  'True
         Sections(2).Cells(3).Format=   "###,###,###,###,###,##0.00"
         Sections(2).Cells(3).Style.Name=   "<private>"
         Sections(2).Cells(3).Style.ParentName=   "Tdb_Body"
         Sections(2).Cells(3).Style.Font_Name=   "Courier"
         Sections(2).Cells(3).Style.Font_Size=   9.75
         Sections(2).Cells(3).Style.Font_Bold=   0   'False
         Sections(2).Cells(3).Style.Font_Italic=   0   'False
         Sections(2).Cells(3).Style.Font_Underline=   0   'False
         Sections(2).Cells(3).Style.Font_Strikeout=   0   'False
         Sections(2).Cells(3).Style.Font_Charset=   0
         Sections(2).Cells(3).Style.TextAlign=   0
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
         Sections(2).Cells(4).Name=   "CELL_6"
         Sections(2).Cells(4).Exp=   "Qty"
         Sections(2).Cells(4).Width=   14
         Sections(2).Cells(4).PrivateStyle=   -1  'True
         Sections(2).Cells(4).Style.Name=   "<private>"
         Sections(2).Cells(4).Style.ParentName=   "Tdb_Body"
         Sections(2).Cells(4).Style.Font_Name=   "Courier"
         Sections(2).Cells(4).Style.Font_Size=   9.75
         Sections(2).Cells(4).Style.Font_Bold=   0   'False
         Sections(2).Cells(4).Style.Font_Italic=   0   'False
         Sections(2).Cells(4).Style.Font_Underline=   0   'False
         Sections(2).Cells(4).Style.Font_Strikeout=   0   'False
         Sections(2).Cells(4).Style.Font_Charset=   0
         Sections(2).Cells(4).Style.TextAlign=   3
         Sections(2).Cells(4).Style.TextVAlign=   1
         Sections(2).Cells(4).Style.TextWrap=   -1  'True
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
         Sections(2).Cells(4).Style.fprops=   1835008
         Sections(2).Cells(5).Name=   "CELL_1"
         Sections(2).Cells(5).Exp=   "Total"
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
         Sections(5).Cells.Count=   4
         Sections(5).Cells(0).Name=   "CELL_0"
         Sections(5).Cells(0).Exp=   """Yang Menerima"""
         Sections(5).Cells(0).NewLine=   -1  'True
         Sections(5).Cells(0).Width=   15
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
         Sections(5).Cells(0).Style.TextAlign=   1
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
         Sections(5).Cells(0).Style.fprops=   294913
         Sections(5).Cells(1).Name=   "CELL_15"
         Sections(5).Cells(1).Exp=   """Yang Menyerahkan"""
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
         Sections(5).Cells(1).Style.TextAlign=   1
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
         Sections(5).Cells(1).Style.fprops=   294913
         Sections(5).Cells(2).Name=   "CELL_4"
         Sections(5).Cells(2).Exp=   """Total : """
         Sections(5).Cells(2).Width=   15
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
         Sections(5).Cells(2).Style.BorderHT=   ""
         Sections(5).Cells(2).Style.BorderHI=   ""
         Sections(5).Cells(2).Style.BorderHB=   ""
         Sections(5).Cells(2).Style.BorderVL=   ""
         Sections(5).Cells(2).Style.BorderVI=   ""
         Sections(5).Cells(2).Style.BorderVR=   ""
         Sections(5).Cells(2).Style.NoClipping=   0   'False
         Sections(5).Cells(2).Style.RTF=   0   'False
         Sections(5).Cells(2).Style.fprops=   1
         Sections(5).Cells(3).Name=   "CELL_2"
         Sections(5).Cells(3).Exp=   "nSubTotal"
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
         Sections(5).Cells(3).Style.TextAlign=   2
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
         Sections(5).Cells(3).Style.fprops=   1081345
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
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   630
      Left            =   30
      Top             =   4815
      Width           =   6915
      _ExtentX        =   12192
      _ExtentY        =   1101
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
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
         Left            =   5670
         TabIndex        =   0
         Top             =   120
         Width           =   1110
         _ExtentX        =   1969
         _ExtentY        =   762
         Caption         =   "     &Exit"
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         BackColor       =   -2147483633
         Picture         =   "rptRekapPenjualanBelumLunas.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   5220
         TabIndex        =   1
         Top             =   120
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   762
         Caption         =   ""
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         BackColor       =   -2147483633
         Picture         =   "rptRekapPenjualanBelumLunas.frx":00A6
      End
      Begin BiSAButtonProject.BiSAButton cmdPrint 
         Height          =   435
         Left            =   4770
         TabIndex        =   21
         Top             =   120
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   762
         Caption         =   ""
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         BackColor       =   -2147483633
         Picture         =   "rptRekapPenjualanBelumLunas.frx":032C
      End
   End
End
Attribute VB_Name = "rptRekapPenjualanBelumLunas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB

Private Sub BiSAButton1_Click()

End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdPreview_Click()
Dim cStatus As String

  cStatus = ""
  If optLunas(0).Value = True Then
    cStatus = 0
  End If
  If optLunas(1).Value = True Then
    cStatus = 1
  End If
  
  If Option1(0).Value = True Then
    GetData cStatus
  End If
  If Option1(1).Value = True Then
    GetDataBeli cStatus
  End If
End Sub

Private Sub cmdPrint_Click()
Dim n As Integer
Dim cTerbilang As String
Dim cField As String
Dim vaJoin
Dim vaGrid As New XArrayDB
Dim cHead As String
Dim cSQL As String
Dim nTotalRekapan As Double
Dim nQT As Double

  'Mengambil data penjualan
  nTotalRekapan = 0
  cSQL = ""
  cSQL = cSQL & " select t.kodeanggota,a.nama,t.username,t.dp,t.tgl,t.kodegroupsales,t.nomorpenjualan,t.total from totpenjualan t"
  cSQL = cSQL & " left join anggota a on a.kodeanggota = t.kodeanggota"
  cSQL = cSQL & " WHERE t.tgl >='" & Format(dTglJual(0).Value, "yyyy-MM-dd") & "' and t.tgl <= '" & Format(dTglJual(1).Value, "yyyy-MM-dd") & "'"
  cSQL = cSQL & " and t.kodeanggota = '" & cKodeAnggota.Text & "'"
  
  Set dbData = objData.SQL(GetDSN, cSQL)
  If Not dbData.EOF Then
    n = 0
    vaGrid.ReDim 0, dbData.RecordCount - 1, 0, 5
    Do While Not dbData.EOF
       nQT = GetPcsJual(GetNull(dbData!nomorpenjualan))
       vaGrid(n, 0) = n + 1
       vaGrid(n, 1) = GetNull(dbData!tgl) 'tgl
       vaGrid(n, 2) = GetNull(dbData!nomorpenjualan) 'invoice
       vaGrid(n, 3) = GetNull(dbData!KodeGroupSales) 'item
       vaGrid(n, 4) = IIf(nQT <> 0, nQT & " Pcs", "") 'qty
       vaGrid(n, 5) = GetNull(dbData!Total) 'total
       nTotalRekapan = nTotalRekapan + vaGrid(n, 5)
       dbData.MoveNext
       n = n + 1
    Loop
    
    'AMBIL INFORMASI customer
    cSQL = ""
    cSQL = cSQL & "select a.kodeanggota,a.nama,a.alamat,a.telp from anggota a"
    cSQL = cSQL & " Where a.kodeanggota = '" & cKodeAnggota.Text & "'"
    
    Set dbData = objData.SQL(GetDSN, cSQL)
    cTerbilang = "# " & Dec2Text(nTotalRekapan) & "Rupiah #"
    cHead = "Kuitansi Lunas"
    With rptKuitansiLunas
      .Parameters("dTgl").ValueExpression = "'" & Format(GetNull(Date), "dd-MM-yyyy") & "'"
      .Parameters("cSE").ValueExpression = "'" & "000-000-000" & "'"
      
      .Parameters("cNama").ValueExpression = "'" & GetNull(dbData!nama, "") & "'"
      .Parameters("cAlamat").ValueExpression = "'" & GetNull(dbData!alamat, "") & "'"
      .Parameters("cKota").ValueExpression = "'" & GetNull(dbData!telp) & "'"
      .Parameters("cKodeAnggota").ValueExpression = "'" & GetNull(dbData!kodeanggota, "") & "'"
      
      .Parameters("cTerbilang").ValueExpression = "'" & cTerbilang & "'"
      .Parameters("cTTD").ValueExpression = "'" & Padc(GetRegistry(reg_FullName), 45) & "'"
      .Parameters("cReceived").ValueExpression = "'" & Padc("", 45) & "'"
      
      .Parameters("nSubtotal").ValueExpression = GetNull(nTotalRekapan)
      .Parameters("nTotal").ValueExpression = GetNull(nTotalRekapan)
      .Parameters("cNamaPerusahaan").ValueExpression = "'" & aCfg(objData, msNamaPerusahaan) & "'"
      .Parameters("cAlamatPerusahaan").ValueExpression = "'" & aCfg(objData, msAlamatPerusahaan) & " " & aCfg(objData, msTelepon) & "'"
      .Parameters("cUserName").ValueExpression = "'" & GetRegistry(reg_FullName) & "'"
      .Parameters("cJudul").ValueExpression = "'" & "Tagihan" & "'"
      .Parameters("keAkun").ValueExpression = "'" & "Akun Palsu" & "'"
      
      Set .Array = vaGrid
      .Refresh
      If MsgBox("Apakah cetakan mau dalam bentuk kertas A4?!!" & vbCrLf & "Jika tidak maka cetakan akan dalam bentuk 1/2 kertas kuarto", vbYesNo) = vbYes Then
        .Profiles(0).PrinterPaperSize = tdbPPS_A4
      End If
      .PrintPreview
    End With
  End If
End Sub

Private Sub cNamaanggota_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "anggota", "kodeanggota,nama", "kodeanggota", sisContent, cNamaAnggota.Text, " OR kodeanggota like '%" & cNamaAnggota.Text & "%'")
  If Not dbData.EOF Then
    cNamaAnggota.Text = cNamaAnggota.Browse(dbData)
    cNamaAnggota.Text = GetNull(dbData!nama)
    cKodeAnggota.Text = GetNull(dbData!kodeanggota)
  End If
End Sub

Private Sub cNamaGroupSales_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "groupsales", "kode,keterangan", "kode", sisContent, cNamaGroupSales.Text, " OR keterangan like '%" & cNamaGroupSales.Text & "%'")
  If Not dbData.EOF Then
    cNamaGroupSales.Text = cNamaGroupSales.Browse(dbData)
    cKodeGroupSales.Text = GetNull(dbData!Kode)
    cNamaGroupSales.Text = GetNull(dbData!keterangan)
  End If
End Sub

Private Sub cNamaGroupSales_Validate(Cancel As Boolean)
  If Trim(cNamaGroupSales.Text) = "" Then
    cKodeGroupSales.Text = ""
  End If
End Sub

Private Sub Form_Load()
Dim n As Single

  SetIcon Me.hWnd, "SIKD"
  CenterForm Me
  cNamaGroupSales.Text = ""
  optLunas(0).Value = True
  Option1(0).Value = True
  cNamaAnggota.Text = ""
  TabIndex Option1(0), n
  TabIndex dTglJual(0), n
  TabIndex dTglJual(1), n
  TabIndex Option1(1), n
  TabIndex dTglBeli(0), n
  TabIndex dTglBeli(1), n
  TabIndex cNamaGroupSales, n
  TabIndex optLunas(0), n
  TabIndex optLunas(1), n
  TabIndex optLunas(2), n
  TabIndex cmdPreview, n
  TabIndex cmdKeluar, n
End Sub

Private Sub GetData(ByVal cStatus As String)
Dim n As Integer
Dim cSQL As String
Dim cSQLStatus As String
Dim nQT As Double

  cSQL = ""
  cSQLStatus = ""
  If cStatus = "0" Or cStatus = "1" Then
    cSQLStatus = " and flaglunas = '" & cStatus & "'"
  End If
  cSQL = cSQL & " select t.kodeanggota,a.nama,t.username,t.dp,t.tgl,t.kodegroupsales,t.nomorpenjualan,t.total from totpenjualan t"
  cSQL = cSQL & " left join anggota a on a.kodeanggota = t.kodeanggota"
  cSQL = cSQL & " WHERE t.tgl >='" & Format(dTglJual(0).Value, "yyyy-MM-dd") & "' and t.tgl <= '" & Format(dTglJual(1).Value, "yyyy-MM-dd") & "'"
  cSQL = cSQL & cSQLStatus
  
  If Trim(cKodeGroupSales.Text) <> "" Then
    cSQL = cSQL & " and t.kodegroupsales = '" & cKodeGroupSales.Text & "'"
  End If
  'cSQL = cSQL & " Where t.flaglunas = 0 Or t.flaglunas Is Null"
  cSQL = cSQL & " order by t.kodeanggota,t.tgl"
  
  vaArray.ReDim 0, -1, 0, 8
  Set dbData = objData.SQL(GetDSN, cSQL)
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      nQT = GetPcsJual(GetNull(dbData!nomorpenjualan))
      vaArray(n, 0) = GetNull(dbData!kodeanggota)
      vaArray(n, 1) = GetNull(dbData!nama)
      vaArray(n, 2) = GetNull(dbData!UserName)
      vaArray(n, 3) = Format(GetNull(dbData!tgl), "dd-MM-yyyy")
      vaArray(n, 4) = GetNull(dbData!nomorpenjualan)
      vaArray(n, 5) = GetNull(dbData!KodeGroupSales) & " " & IIf(nQT <> 0, nQT & " pcs", "")
      vaArray(n, 6) = GetNull(dbData!Total) - GetNull(dbData!dp)
      vaArray(n, 7) = ""
      vaArray(n, 8) = ""
      dbData.MoveNext
    Loop
    GetRpt cStatus
  End If
End Sub

Private Sub GetDataBeli(ByVal cStatus As String)
Dim n As Integer
Dim cSQL As String
Dim nQT As Double

  cSQL = ""
  cSQL = cSQL & " select t.kodesupplier,a.nama,t.username,t.tgl,t.kodegroupsales,t.nomorpembelian,t.total from totpembelian t"
  cSQL = cSQL & " left join supplier a on a.kodesupplier = t.kodesupplier"
  cSQL = cSQL & " WHERE t.tgl >='" & Format(dTglBeli(0).Value, "yyyy-MM-dd") & "' and t.tgl <= '" & Format(dTglBeli(1).Value, "yyyy-MM-dd") & "'"
  If Trim(cKodeGroupSales.Text) <> "" Then
    cSQL = cSQL & " and t.kodegroupsales = '" & cKodeGroupSales.Text & "'"
  End If
  'cSQL = cSQL & " Where t.flaglunas = 0 Or t.flaglunas Is Null"
  cSQL = cSQL & " order by t.kodesupplier,t.tgl"
  
  vaArray.ReDim 0, -1, 0, 8
  Set dbData = objData.SQL(GetDSN, cSQL)
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      nQT = GetPcsBeli(GetNull(dbData!nomorpembelian))
      
      vaArray(n, 0) = GetNull(dbData!kodesupplier)
      vaArray(n, 1) = GetNull(dbData!nama)
      vaArray(n, 2) = GetNull(dbData!UserName)
      vaArray(n, 3) = Format(GetNull(dbData!tgl), "dd-MM-yyyy")
      vaArray(n, 4) = GetNull(dbData!nomorpembelian)
      vaArray(n, 5) = GetNull(dbData!KodeGroupSales) & " " & IIf(nQT <> 0, nQT & " pcs", "")
      vaArray(n, 6) = GetNull(dbData!Total)
      vaArray(n, 7) = ""
      vaArray(n, 8) = ""
      dbData.MoveNext
    Loop
    GetRptBeli cStatus
  End If
End Sub

Private Function GetPcsBeli(ByVal cNomorBeli As String) As Double
Dim dbx As New ADODB.Recordset

  GetPcsBeli = 0
  Set dbx = objData.Browse(GetDSN, "pembelian", "sum(qty) as totalpcs", "nomorpembelian", sisAssign, cNomorBeli)
  If Not dbx.EOF Then
    GetPcsBeli = GetNull(dbx!totalpcs)
  End If
End Function

Private Function GetPcsJual(ByVal cNomorJual As String) As Double
Dim dbx As New ADODB.Recordset

  GetPcsJual = 0
  Set dbx = objData.Browse(GetDSN, "penjualan", "sum(qty) as totalpcs", "nomorpenjualan", sisAssign, cNomorJual)
  If Not dbx.EOF Then
    GetPcsJual = GetNull(dbx!totalpcs)
  End If
End Function

Private Sub GetRptBeli(ByVal cStatus As String)
Dim cHeadReport As String
  
  cHeadReport = ""
  If cStatus = "0" Then
    cHeadReport = " Belum Lunas"
  End If
  If cStatus = "1" Then
    cHeadReport = " Sudah Lunas"
  End If
  
  With FrmRPT
    
    .AddPageHeader "Data Semua Tagihan Pembelian" & cHeadReport, tdbHalignCenter, , , True, , 10, True, False, True, False, tdbPageHeaderSect
    .AddPageHeader aCfg(objData, msNamaPerusahaan), tdbHalignCenter, , , True, , 14, True, False, True, False, tdbPageHeaderSect
        
    .AddTableGroupHeader True, "[]", , , , 15
    .AddTableGroupHeader
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    
    .AddTableHeader "", , , , , , , , , , , , , , , , , , , False
    .AddTableHeader "", , , , , , , , , , , , , , , , , , , False
    .AddTableHeader "Kasir", , , , 15
    .AddTableHeader "Tgl", , , , 12
    .AddTableHeader "Nomor", , , , 17
    .AddTableHeader "Keterangan", , , , 15
    .AddTableHeader "Nominal", , , , 15
    .AddTableHeader "Tgl Trf", , , , 15
    .AddTableHeader "TTD", , , , 15
     
    .AddTableBody , , , , , , , , , , , , , False
    .AddTableBody , , , , , , , , , , , , , False
    .AddTableBody
    .AddTableBody
    .AddTableBody
    .AddTableBody
    .AddTableBody Sis_Rpt_Number2
    .AddTableBody , , , , , , , , , , , tdbMergeAlways
    .AddTableBody , , , , , , , , , , , tdbMergeAlways
    
    .AddTableGroupFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupFooter "Sub Total", , tdbHalignRight, , , , , , , , , , , , 4
    .AddTableGroupFooter
    .AddTableGroupFooter
    .AddTableGroupFooter
    .AddTableGroupFooter "&Sum", Sis_Rpt_Number2
    .AddTableGroupFooter
    .AddTableGroupFooter
    
    .AddTableFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableFooter "Total", , tdbHalignRight, , , , , , , , , , , , 4
    .AddTableFooter
    .AddTableFooter
    .AddTableFooter
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    .AddTableFooter
    .AddTableFooter
    
    .Refresh
    .Preview vaArray, True
  End With
End Sub

Private Sub GetRpt(ByVal cStatus As String)
Dim cHeadReport As String
  
  cHeadReport = ""
  If cStatus = "0" Then
    cHeadReport = " Belum Lunas"
  End If
  If cStatus = "1" Then
    cHeadReport = " Sudah Lunas"
  End If
  
  With FrmRPT
    
    .AddPageHeader "Data Semua Tagihan Penjualan" & cHeadReport, tdbHalignCenter, , , True, , 10, True, False, True, False, tdbPageHeaderSect
    .AddPageHeader aCfg(objData, msNamaPerusahaan), tdbHalignCenter, , , True, , 14, True, False, True, False, tdbPageHeaderSect
        
    .AddTableGroupHeader True, "[]", , , , 15
    .AddTableGroupHeader
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupHeader , , , , , , , , , , , , , , , , , , , , False
    
    .AddTableHeader "", , , , , , , , , , , , , , , , , , , False
    .AddTableHeader "", , , , , , , , , , , , , , , , , , , False
    .AddTableHeader "Kasir", , , , 15
    .AddTableHeader "Tgl", , , , 12
    .AddTableHeader "Nomor", , , , 17
    .AddTableHeader "Keterangan", , , , 15
    .AddTableHeader "Nominal", , , , 15
    .AddTableHeader "Tgl Trf", , , , 15
    .AddTableHeader "TTD", , , , 15
     
    .AddTableBody , , , , , , , , , , , , , False
    .AddTableBody , , , , , , , , , , , , , False
    .AddTableBody
    .AddTableBody
    .AddTableBody
    .AddTableBody
    .AddTableBody Sis_Rpt_Number2
    .AddTableBody , , , , , , , , , , , tdbMergeAlways
    .AddTableBody , , , , , , , , , , , tdbMergeAlways
    
    .AddTableGroupFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableGroupFooter "Sub Total", , tdbHalignRight, , , , , , , , , , , , 4
    .AddTableGroupFooter
    .AddTableGroupFooter
    .AddTableGroupFooter
    .AddTableGroupFooter "&Sum", Sis_Rpt_Number2
    .AddTableGroupFooter
    .AddTableGroupFooter
    
    .AddTableFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableFooter , , , , , , , , , , , , , , , , , , , False
    .AddTableFooter "Total", , tdbHalignRight, , , , , , , , , , , , 4
    .AddTableFooter
    .AddTableFooter
    .AddTableFooter
    .AddTableFooter "&Sum", Sis_Rpt_Number2
    .AddTableFooter
    .AddTableFooter
    
    .Refresh
    .Preview vaArray, True
  End With
End Sub

Private Sub Option1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    SendKeysA vbKeyTab, True
  End If
End Sub

Private Sub optLunas_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    SendKeysA vbKeyTab, True
  End If
End Sub
