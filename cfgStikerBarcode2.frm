VERSION 5.00
Object = "{0D6235E7-DBA2-11D1-B5DF-0060976089D0}#1.0#0"; "tdbr6.ocx"
Object = "{32A289A9-C7B2-11D4-8714-444553540000}#4.1#0"; "SisTFrame.ocx"
Object = "{9A5A31AC-C750-11D4-8714-444553540000}#5.2#0"; "SisTrueNumberBox.ocx"
Object = "{8164BC59-C899-11D4-8714-444553540000}#5.0#0"; "SisTLabel.ocx"
Object = "{8300D29B-6BA1-4A90-A806-DF2ECAC5A300}#2.0#0"; "SisTButton.ocx"
Begin VB.Form cfgStikerBarcode2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stiker Barcode"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   5805
   Begin SisTFrame.sisFrame sisFrame3 
      Height          =   1065
      Left            =   0
      Top             =   0
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   1879
      Caption         =   " Start Pencetakan "
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   4
      BackColor       =   -2147483633
      Begin SisTLabel.SisLabel SisLabel1 
         Height          =   330
         Index           =   2
         Left            =   2010
         TabIndex        =   0
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "1 s/d 5"
         SemiColon       =   0   'False
      End
      Begin vb6projectSisNumber.SisNumber nBaris 
         Height          =   330
         Left            =   210
         TabIndex        =   1
         Top             =   600
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   582
         Decimals        =   0
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Baris"
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin vb6projectSisNumber.SisNumber nKolom 
         Height          =   330
         Left            =   210
         TabIndex        =   2
         Top             =   240
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   582
         Decimals        =   0
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Kolom"
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SisTLabel.SisLabel SisLabel2 
         Height          =   330
         Index           =   2
         Left            =   2010
         TabIndex        =   3
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "1 s/d 8"
         SemiColon       =   0   'False
      End
   End
   Begin SisTFrame.sisFrame sisFrame1 
      Height          =   1545
      Left            =   0
      Top             =   1110
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   2725
      Caption         =   " Batas Halaman "
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   4
      BackColor       =   -2147483633
      Begin VB.OptionButton opt 
         Caption         =   "&1. Portrait"
         Height          =   330
         Index           =   0
         Left            =   1350
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1050
         Width           =   1275
      End
      Begin VB.OptionButton opt 
         Caption         =   "&2. Lanscape"
         Height          =   330
         Index           =   1
         Left            =   2685
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1050
         Width           =   1275
      End
      Begin SisTLabel.SisLabel SisLabel3 
         Height          =   330
         Left            =   210
         TabIndex        =   6
         Top             =   1050
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   582
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Orientasi"
      End
      Begin SisTLabel.SisLabel SisLabel1 
         Height          =   330
         Index           =   0
         Left            =   2250
         TabIndex        =   7
         Top             =   300
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   582
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "mm"
         SemiColon       =   0   'False
      End
      Begin vb6projectSisNumber.SisNumber nBottom 
         Height          =   330
         Left            =   2880
         TabIndex        =   8
         Top             =   660
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   582
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Batas Bawah"
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin vb6projectSisNumber.SisNumber nRight 
         Height          =   330
         Left            =   2880
         TabIndex        =   9
         Top             =   300
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   582
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Batas Kanan"
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin vb6projectSisNumber.SisNumber nTop 
         Height          =   330
         Left            =   210
         TabIndex        =   10
         Top             =   660
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   582
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Batas Atas"
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin vb6projectSisNumber.SisNumber nLeft 
         Height          =   330
         Left            =   210
         TabIndex        =   11
         Top             =   300
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   582
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Batas Kiri"
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SisTLabel.SisLabel SisLabel2 
         Height          =   330
         Index           =   0
         Left            =   2250
         TabIndex        =   12
         Top             =   660
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   582
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "mm"
         SemiColon       =   0   'False
      End
      Begin SisTLabel.SisLabel SisLabel1 
         Height          =   330
         Index           =   1
         Left            =   4920
         TabIndex        =   13
         Top             =   300
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   582
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "mm"
         SemiColon       =   0   'False
      End
      Begin SisTLabel.SisLabel SisLabel2 
         Height          =   330
         Index           =   1
         Left            =   4920
         TabIndex        =   14
         Top             =   660
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   582
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "mm"
         SemiColon       =   0   'False
      End
   End
   Begin SisTFrame.sisFrame sisFrame2 
      Height          =   555
      Left            =   0
      Top             =   2670
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   979
      BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   4
      BackColor       =   -2147483633
      Begin SISTButton.SISButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   4680
         TabIndex        =   15
         Top             =   90
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   661
         Caption         =   "      K&eluar"
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
         Picture         =   "cfgStikerBarcode2.frx":0000
      End
      Begin SISTButton.SISButton cmdPreview 
         Height          =   375
         Left            =   3660
         TabIndex        =   16
         Top             =   90
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   661
         Caption         =   "      &Preview"
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
         Picture         =   "cfgStikerBarcode2.frx":059A
      End
      Begin TrueDBReports60Ctl.TDBReports TDBReports1 
         Height          =   570
         Left            =   60
         TabIndex        =   17
         Top             =   0
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   1005
         Caption         =   "TDBReports1"
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
         ConnectionString=   "DSN=AssistPro"
         ConnectStringType=   3
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   "AssistPro"
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         CursorLocation  =   3
         ConnectionTimeout=   15
         CommandTimeout  =   30
         RecordSource    =   $"cfgStikerBarcode2.frx":0EBC
         CursorType      =   3
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
         Fields.Count    =   15
         Fields(0).Name  =   "Nama1"
         Fields(0).DisplayName=   "Nama1"
         Fields(1).Name  =   "Barcode1"
         Fields(1).DisplayName=   "Barcode1"
         Fields(2).Name  =   "Kode1"
         Fields(2).DisplayName=   "Kode1"
         Fields(3).Name  =   "Nama2"
         Fields(3).DisplayName=   "Nama2"
         Fields(4).Name  =   "Barcode2"
         Fields(4).DisplayName=   "Barcode2"
         Fields(5).Name  =   "Kode2"
         Fields(5).DisplayName=   "Kode2"
         Fields(6).Name  =   "Nama3"
         Fields(6).DisplayName=   "Nama3"
         Fields(7).Name  =   "Barcode3"
         Fields(7).DisplayName=   "Barcode3"
         Fields(8).Name  =   "Kode3"
         Fields(8).DisplayName=   "Kode3"
         Fields(9).Name  =   "Nama4"
         Fields(9).DisplayName=   "Nama4"
         Fields(10).Name =   "Barcode4"
         Fields(10).DisplayName=   "Barcode4"
         Fields(11).Name =   "Kode4"
         Fields(11).DisplayName=   "Kode4"
         Fields(12).Name =   "Nama5"
         Fields(12).DisplayName=   "Nama5"
         Fields(13).Name =   "Barcode5"
         Fields(13).DisplayName=   "Barcode5"
         Fields(14).Name =   "Kode5"
         Fields(14).DisplayName=   "Kode5"
         Sections.Count  =   3
         Sections(0).Name=   "Detail"
         Sections(0).Type=   4
         Sections(0).StyleExp=   "'tdb_Nama'"
         Sections(0).AutoHeight=   0   'False
         Sections(0).Height=   4
         Sections(0).CallBeforePrint=   -1  'True
         Sections(0).Cells.Count=   9
         Sections(0).Cells(0).Name=   "CELL_0"
         Sections(0).Cells(0).Exp=   """Sweety"""
         Sections(0).Cells(0).StyleExp=   "'tdb_label'"
         Sections(0).Cells(0).Width=   38
         Sections(0).Cells(0).WidthInPercent=   0   'False
         Sections(0).Cells(1).Name=   "CELL_7"
         Sections(0).Cells(1).Exp=   """"""
         Sections(0).Cells(1).Width=   3
         Sections(0).Cells(1).WidthInPercent=   0   'False
         Sections(0).Cells(2).Name=   "CELL_1"
         Sections(0).Cells(2).Exp=   """Sweety"""
         Sections(0).Cells(2).StyleExp=   "'tdb_label'"
         Sections(0).Cells(2).Width=   38
         Sections(0).Cells(2).WidthInPercent=   0   'False
         Sections(0).Cells(3).Name=   "CELL_6"
         Sections(0).Cells(3).Exp=   """"""
         Sections(0).Cells(3).Width=   3
         Sections(0).Cells(3).WidthInPercent=   0   'False
         Sections(0).Cells(4).Name=   "CELL_2"
         Sections(0).Cells(4).Exp=   """Sweety"""
         Sections(0).Cells(4).StyleExp=   "'tdb_label'"
         Sections(0).Cells(4).Width=   38
         Sections(0).Cells(4).WidthInPercent=   0   'False
         Sections(0).Cells(5).Name=   "CELL_5"
         Sections(0).Cells(5).Exp=   """"""
         Sections(0).Cells(5).Width=   3
         Sections(0).Cells(5).WidthInPercent=   0   'False
         Sections(0).Cells(6).Name=   "CELL_3"
         Sections(0).Cells(6).Exp=   """Sweety"""
         Sections(0).Cells(6).StyleExp=   "'tdb_label'"
         Sections(0).Cells(6).Width=   38
         Sections(0).Cells(6).WidthInPercent=   0   'False
         Sections(0).Cells(7).Name=   "CELL_9"
         Sections(0).Cells(7).Exp=   """"""
         Sections(0).Cells(7).Width=   3
         Sections(0).Cells(7).WidthInPercent=   0   'False
         Sections(0).Cells(8).Name=   "CELL_4"
         Sections(0).Cells(8).Exp=   """Sweety"""
         Sections(0).Cells(8).StyleExp=   "'tdb_label'"
         Sections(0).Cells(8).Width=   38
         Sections(0).Cells(8).WidthInPercent=   0   'False
         Sections(1).Name=   "SECTION_1"
         Sections(1).StyleExp=   "'tdb_Barode'"
         Sections(1).Tabulator=   "Detail"
         Sections(1).AutoHeight=   0   'False
         Sections(1).Height=   8
         Sections(1).dtopts=   2
         Sections(1).Cells.Count=   9
         Sections(1).Cells(0).Name=   "CELL_0"
         Sections(1).Cells(0).Exp=   "Barcode1"
         Sections(1).Cells(1).Name=   "CELL_1"
         Sections(1).Cells(2).Name=   "CELL_2"
         Sections(1).Cells(2).Exp=   "Barcode2"
         Sections(1).Cells(3).Name=   "CELL_3"
         Sections(1).Cells(4).Name=   "CELL_4"
         Sections(1).Cells(4).Exp=   "Barcode3"
         Sections(1).Cells(5).Name=   "CELL_5"
         Sections(1).Cells(6).Name=   "CELL_6"
         Sections(1).Cells(6).Exp=   "Barcode4"
         Sections(1).Cells(7).Name=   "CELL_7"
         Sections(1).Cells(8).Name=   "CELL_8"
         Sections(1).Cells(8).Exp=   "Barcode5"
         Sections(2).Name=   "SECTION_2"
         Sections(2).StyleExp=   "'tdb_Kode'"
         Sections(2).Tabulator=   "Detail"
         Sections(2).AutoHeight=   0   'False
         Sections(2).Height=   6
         Sections(2).SpacingAfter=   2
         Sections(2).Cells.Count=   9
         Sections(2).Cells(0).Name=   "CELL_0"
         Sections(2).Cells(0).Exp=   "Kode1"
         Sections(2).Cells(1).Name=   "CELL_1"
         Sections(2).Cells(2).Name=   "CELL_2"
         Sections(2).Cells(2).Exp=   "Kode2"
         Sections(2).Cells(3).Name=   "CELL_3"
         Sections(2).Cells(4).Name=   "CELL_4"
         Sections(2).Cells(4).Exp=   "Kode3"
         Sections(2).Cells(5).Name=   "CELL_5"
         Sections(2).Cells(6).Name=   "CELL_6"
         Sections(2).Cells(6).Exp=   "Kode4"
         Sections(2).Cells(7).Name=   "CELL_7"
         Sections(2).Cells(8).Name=   "CELL_8"
         Sections(2).Cells(8).Exp=   "Kode5"
         Styles.Count    =   5
         Styles(0).Name  =   "tdb_Base"
         Styles(0).Font_Name=   "Arial"
         Styles(0).Font_Size=   6
         Styles(0).Font_Charset=   0
         Styles(0).TextAlign=   0
         Styles(0).TextVAlign=   1
         Styles(0).TextWrap=   0   'False
         Styles(0).BackColor=   12632256
         Styles(0).MarginTop=   0
         Styles(0).MarginBottom=   1
         Styles(1).Name  =   "tdb_label"
         Styles(1).ParentName=   "tdb_Base"
         Styles(1).Font_Name=   "Arial"
         Styles(1).Font_Size=   6
         Styles(1).Font_Charset=   0
         Styles(1).TextAlign=   1
         Styles(1).TextVAlign=   2
         Styles(1).TextWrap=   0   'False
         Styles(1).BackColor=   12632256
         Styles(1).MarginTop=   0
         Styles(1).MarginBottom=   1
         Styles(1).fprops=   3
         Styles(2).Name  =   "tdb_Barode"
         Styles(2).ParentName=   "tdb_Base"
         Styles(2).Font_Name=   "CIA EAN Truncated"
         Styles(2).Font_Size=   24
         Styles(2).Font_Charset=   0
         Styles(2).TextAlign=   1
         Styles(2).TextVAlign=   1
         Styles(2).TextWrap=   0   'False
         Styles(2).BackColor=   12632256
         Styles(2).MarginTop=   0
         Styles(2).MarginBottom=   1
         Styles(2).fprops=   6291457
         Styles(3).Name  =   "tdb_Kode"
         Styles(3).ParentName=   "tdb_Base"
         Styles(3).Font_Name=   "Arial"
         Styles(3).Font_Size=   8.25
         Styles(3).Font_Charset=   0
         Styles(3).TextAlign=   1
         Styles(3).TextWrap=   0   'False
         Styles(3).BackColor=   12632256
         Styles(3).MarginTop=   0
         Styles(3).MarginBottom=   1
         Styles(3).fprops=   4194307
         Styles(4).Name  =   "tdb_Nama"
         Styles(4).ParentName=   "tdb_Base"
         Styles(4).Font_Name=   "Arial"
         Styles(4).Font_Size=   8.25
         Styles(4).Font_Charset=   0
         Styles(4).TextAlign=   1
         Styles(4).TextVAlign=   1
         Styles(4).TextWrap=   0   'False
         Styles(4).BackColor=   12632256
         Styles(4).MarginTop=   0
         Styles(4).MarginBottom=   1
         Styles(4).fprops=   4194305
         Profiles.Count  =   1
         Profiles(0).Name=   "PROFILE_0"
         Profiles(0).Active=   -1  'True
         Profiles(0).PreviewNoMinimize=   -1  'True
         Profiles(0).PreviewNoMaximize=   -1  'True
         Profiles(0).PreviewNoResize=   -1  'True
         Profiles(0).PreviewMaximized=   -1  'True
         Profiles(0).PreviewNoSaveLoad=   -1  'True
         Profiles(0).PrinterMarginLeft=   0
         Profiles(0).PrinterMarginTop=   4
         Profiles(0).PrinterMarginRight=   6
         Profiles(0).PrinterMarginBottom=   10
         Profiles(0).PrinterMargins_set=   -1  'True
      End
   End
End
Attribute VB_Name = "cfgStikerBarcode2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim nPage As Double
Dim nMaxPage As Double
Dim dbData As New ADODB.Recordset
Dim objData As New SISMyDLL.Data
Dim vaArray As New XArrayDB
Dim va As New XArrayDB

Sub PrintBarcode(xArray As XArrayDB)
  Set va = xArray
  Me.Show vbModal
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdPreview_Click()
Dim nRow As Double
Dim nCol As Double
Dim nr As Double
Dim nc As Double
Dim n As Double

  UpdCfg msStikerLeftMargin, nLeft.Value
  UpdCfg msStikerRightMargin, nRight.Value
  UpdCfg msStikerTopMargin, nTop.Value
  UpdCfg msStikerBottomMargin, nBottom.Value
  UpdCfg msStikerOrientation, GetOpt(opt)
  nPage = 0
  nMaxPage = 8
  nc = 1000
  
  vaArray.ReDim 0, -1, 0, 14
  ' Tambahkan baris kosong untuk pencetakan jika tidak dimulai dari kolom pertama
  ' Rumusnya adalah
  n = ((5 * (nBaris.Value - 1))) + nKolom.Value - 1
  For nRow = 1 To n Step 1
    va.InsertRows 0
    va(0, 0) = ""
    va(0, 1) = ""
    va(0, 2) = ""
  Next
  
  ' Penambahan Field Barcode
  For nRow = 0 To va.UpperBound(1)
    For nCol = 0 To va.UpperBound(2)
      If nc > vaArray.UpperBound(2) Then
        vaArray.InsertRows vaArray.UpperBound(1) + 1
        nr = vaArray.UpperBound(1)
        nc = 0
      End If
      
      vaArray(nr, nc) = va(nRow, nCol)
      nc = nc + 1
    Next
  Next
  
  With TDBReports1
    Set .Array = vaArray
    .Refresh
    
    .PrintPreview
  End With
End Sub

Private Sub Form_Load()
Dim n As Single
  CenterForm Me
  
  nLeft.Value = aCfg(msStikerLeftMargin, 6)
  nRight.Value = aCfg(msStikerRightMargin, 6)
  nTop.Value = aCfg(msStikerTopMargin, 5)
  nBottom.Value = aCfg(msStikerBottomMargin, 10)
  SetOpt opt, aCfg(msStikerOrientation)
  nKolom.Value = 1
  nBaris.Value = 1
  
  TabIndex nKolom, n
  TabIndex nBaris, n
  TabIndex nLeft, n
  TabIndex nRight, n
  TabIndex nTop, n
  TabIndex nBottom, n
  TabIndex opt(0), n
  TabIndex opt(1), n
  TabIndex cmdPreview, n
  TabIndex cmdKeluar, n
End Sub

Private Sub opt_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 Then
    SendKeys "{Tab}"
  End If
End Sub

Private Sub TDBReports1_SectionBeforePrint(ByVal Section As Integer, ByVal Style As TrueDBReports60Ctl.Style, ByVal Params As TrueDBReports60Ctl.SectionParams)
  If nPage = nMaxPage Then
    nPage = 0
    Params.NewPage = True
  Else
    Params.NewPage = False
    nPage = nPage + 1
  End If
End Sub


