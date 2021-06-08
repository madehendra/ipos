VERSION 5.00
Object = "{0D6235E7-DBA2-11D1-B5DF-0060976089D0}#1.0#0"; "tdbr6.ocx"
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form rptPenjualanKasir 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LAPORAN PENJUALAN KASIR"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   7005
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1740
      Left            =   0
      Top             =   0
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   3069
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
      Begin VB.OptionButton optKodeStock 
         Caption         =   "Barcode"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   3540
         TabIndex        =   9
         Top             =   1155
         Width           =   1035
      End
      Begin VB.OptionButton optKodeStock 
         Caption         =   "Kode Index"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   2280
         TabIndex        =   8
         Top             =   1155
         Width           =   1230
      End
      Begin BiSATextBoxProject.BiSABrowse cKasir 
         Height          =   360
         Left            =   2625
         TabIndex        =   5
         Top             =   690
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   635
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
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   195
         Left            =   2355
         TabIndex        =   4
         Top             =   705
         Width           =   210
      End
      Begin BiSADateProject.BiSADate dTgl 
         Height          =   330
         Index           =   0
         Left            =   225
         TabIndex        =   0
         Top             =   285
         Width           =   3465
         _ExtentX        =   6112
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
         Caption         =   "ANTARA TANGGAL"
         CaptionWidth    =   2000
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
      Begin BiSADateProject.BiSADate dTgl 
         Height          =   330
         Index           =   1
         Left            =   3735
         TabIndex        =   1
         Top             =   285
         Width           =   1965
         _ExtentX        =   3466
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
      Begin TrueDBReports60Ctl.TDBReports tdb 
         Height          =   570
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   1005
         Caption         =   "Penjualan"
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
         ConnectionString=   ""
         ConnectStringType=   1
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
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
         Parameters.Count=   8
         Parameters(0).Name=   "TGL1"
         Parameters(1).Name=   "TGL2"
         Parameters(2).Name=   "TOTALJUMLAH"
         Parameters(2).Type=   5
         Parameters(3).Name=   "TOTALQTY"
         Parameters(3).Type=   5
         Parameters(4).Name=   "TOTALDISCOUNT1"
         Parameters(4).Type=   5
         Parameters(5).Name=   "TOTALDISCOUNT2"
         Parameters(5).Type=   5
         Parameters(6).Name=   "TOTALPAJAK"
         Parameters(6).Type=   5
         Parameters(7).Name=   "GRANDTOTAL"
         Parameters(7).Type=   5
         Fields.Count    =   12
         Fields(0).Name  =   "faktur"
         Fields(0).DisplayName=   "faktur"
         Fields(1).Name  =   "tgl"
         Fields(1).DisplayName=   "tgl"
         Fields(1).Type  =   7
         Fields(2).Name  =   "kode"
         Fields(2).DisplayName=   "kode"
         Fields(3).Name  =   "namabarang"
         Fields(3).DisplayName=   "namabarang"
         Fields(4).Name  =   "qty"
         Fields(4).DisplayName=   "qty"
         Fields(4).Type  =   5
         Fields(5).Name  =   "satuan"
         Fields(5).DisplayName=   "satuan"
         Fields(6).Name  =   "harga"
         Fields(6).DisplayName=   "harga"
         Fields(6).Type  =   5
         Fields(7).Name  =   "jumlah"
         Fields(7).DisplayName=   "jumlah"
         Fields(7).Type  =   5
         Fields(8).Name  =   "subtotal"
         Fields(8).DisplayName=   "subtotal"
         Fields(8).Type  =   5
         Fields(9).Name  =   "discount"
         Fields(9).DisplayName=   "discount"
         Fields(9).Type  =   5
         Fields(10).Name =   "total"
         Fields(10).DisplayName=   "total"
         Fields(10).Type =   5
         Fields(11).Name =   "kasir"
         Fields(11).DisplayName=   "kasir"
         Sections.Count  =   6
         Sections(0).Name=   "SECTION_1"
         Sections(0).Type=   1
         Sections(0).Cells.Count=   6
         Sections(0).Cells(0).Name=   "CELL_0"
         Sections(0).Cells(0).Exp=   """Hal : "" & PageNo()"
         Sections(0).Cells(0).PrivateStyle=   -1  'True
         Sections(0).Cells(0).Style.Name=   "<private>"
         Sections(0).Cells(0).Style.ParentName=   "<null>"
         Sections(0).Cells(0).Style.Font_Name=   "Times New Roman"
         Sections(0).Cells(0).Style.Font_Size=   10
         Sections(0).Cells(0).Style.Font_Bold=   0   'False
         Sections(0).Cells(0).Style.Font_Italic=   0   'False
         Sections(0).Cells(0).Style.Font_Underline=   0   'False
         Sections(0).Cells(0).Style.Font_Strikeout=   0   'False
         Sections(0).Cells(0).Style.Font_Charset=   1
         Sections(0).Cells(0).Style.TextAlign=   2
         Sections(0).Cells(0).Style.TextVAlign=   0
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
         Sections(0).Cells(0).Style.MarginTop=   6
         Sections(0).Cells(0).Style.MarginRight=   6
         Sections(0).Cells(0).Style.MarginBottom=   6
         Sections(0).Cells(0).Style.HasBorders=   -1  'True
         Sections(0).Cells(0).Style.BorderHT=   ""
         Sections(0).Cells(0).Style.BorderHI=   ""
         Sections(0).Cells(0).Style.BorderHB=   ""
         Sections(0).Cells(0).Style.BorderVL=   ""
         Sections(0).Cells(0).Style.BorderVI=   ""
         Sections(0).Cells(0).Style.BorderVR=   ""
         Sections(0).Cells(0).Style.NoClipping=   0   'False
         Sections(0).Cells(0).Style.RTF=   0   'False
         Sections(0).Cells(0).Style.fprops=   1
         Sections(0).Cells(1).Name=   "CELL_1"
         Sections(0).Cells(1).Exp=   """ """
         Sections(0).Cells(1).NewLine=   -1  'True
         Sections(0).Cells(2).Name=   "CELL_2"
         Sections(0).Cells(2).Exp=   """LAPORAN DETAIL PENJUALAN KASIR"""
         Sections(0).Cells(2).NewLine=   -1  'True
         Sections(0).Cells(2).PrivateStyle=   -1  'True
         Sections(0).Cells(2).Style.Name=   "<private>"
         Sections(0).Cells(2).Style.ParentName=   "<null>"
         Sections(0).Cells(2).Style.Font_Name=   "Verdana"
         Sections(0).Cells(2).Style.Font_Size=   12
         Sections(0).Cells(2).Style.Font_Bold=   -1  'True
         Sections(0).Cells(2).Style.Font_Italic=   0   'False
         Sections(0).Cells(2).Style.Font_Underline=   0   'False
         Sections(0).Cells(2).Style.Font_Strikeout=   0   'False
         Sections(0).Cells(2).Style.Font_Charset=   1
         Sections(0).Cells(2).Style.TextAlign=   1
         Sections(0).Cells(2).Style.TextVAlign=   0
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
         Sections(0).Cells(2).Style.MarginTop=   6
         Sections(0).Cells(2).Style.MarginRight=   6
         Sections(0).Cells(2).Style.MarginBottom=   6
         Sections(0).Cells(2).Style.HasBorders=   -1  'True
         Sections(0).Cells(2).Style.BorderHT=   ""
         Sections(0).Cells(2).Style.BorderHI=   ""
         Sections(0).Cells(2).Style.BorderHB=   ""
         Sections(0).Cells(2).Style.BorderVL=   ""
         Sections(0).Cells(2).Style.BorderVI=   ""
         Sections(0).Cells(2).Style.BorderVR=   ""
         Sections(0).Cells(2).Style.NoClipping=   0   'False
         Sections(0).Cells(2).Style.RTF=   0   'False
         Sections(0).Cells(2).Style.fprops=   23068673
         Sections(0).Cells(3).Name=   "CELL_3"
         Sections(0).Cells(3).Exp=   """Antara Tanggal : "" & TGL1 & "" S.D "" & TGL2"
         Sections(0).Cells(3).NewLine=   -1  'True
         Sections(0).Cells(3).PrivateStyle=   -1  'True
         Sections(0).Cells(3).Format=   "dd-Mm-yyyy"
         Sections(0).Cells(3).Style.Name=   "<private>"
         Sections(0).Cells(3).Style.ParentName=   "<null>"
         Sections(0).Cells(3).Style.Font_Name=   "Verdana"
         Sections(0).Cells(3).Style.Font_Size=   9.75
         Sections(0).Cells(3).Style.Font_Bold=   -1  'True
         Sections(0).Cells(3).Style.Font_Italic=   0   'False
         Sections(0).Cells(3).Style.Font_Underline=   0   'False
         Sections(0).Cells(3).Style.Font_Strikeout=   0   'False
         Sections(0).Cells(3).Style.Font_Charset=   1
         Sections(0).Cells(3).Style.TextAlign=   1
         Sections(0).Cells(3).Style.TextVAlign=   0
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
         Sections(0).Cells(3).Style.MarginTop=   6
         Sections(0).Cells(3).Style.MarginRight=   6
         Sections(0).Cells(3).Style.MarginBottom=   6
         Sections(0).Cells(3).Style.HasBorders=   -1  'True
         Sections(0).Cells(3).Style.BorderHT=   ""
         Sections(0).Cells(3).Style.BorderHI=   ""
         Sections(0).Cells(3).Style.BorderHB=   ""
         Sections(0).Cells(3).Style.BorderVL=   ""
         Sections(0).Cells(3).Style.BorderVI=   ""
         Sections(0).Cells(3).Style.BorderVR=   ""
         Sections(0).Cells(3).Style.NoClipping=   0   'False
         Sections(0).Cells(3).Style.RTF=   0   'False
         Sections(0).Cells(3).Style.fprops=   23068673
         Sections(0).Cells(4).Name=   "CELL_4"
         Sections(0).Cells(4).Exp=   """ """
         Sections(0).Cells(4).NewLine=   -1  'True
         Sections(0).Cells(5).Name=   "CELL_5"
         Sections(0).Cells(5).Exp=   """ """
         Sections(0).Cells(5).NewLine=   -1  'True
         Sections(1).Name=   "SECTION_5"
         Sections(1).Condition=   "HasChanged(faktur)"
         Sections(1).StyleExp=   "'tdb_Base'"
         Sections(1).KeepWithNext=   2
         Sections(1).Cells.Count=   6
         Sections(1).Cells(0).Name=   "CELL_0"
         Sections(1).Cells(0).Exp=   """NO. FAKTUR"""
         Sections(1).Cells(0).Width=   20
         Sections(1).Cells(0).Height=   4
         Sections(1).Cells(0).AutoHeight=   0   'False
         Sections(1).Cells(0).PrivateStyle=   -1  'True
         Sections(1).Cells(0).Style.Name=   "<private>"
         Sections(1).Cells(0).Style.ParentName=   "tdb_Base"
         Sections(1).Cells(0).Style.Font_Name=   "Arial"
         Sections(1).Cells(0).Style.Font_Size=   8.25
         Sections(1).Cells(0).Style.Font_Bold=   -1  'True
         Sections(1).Cells(0).Style.Font_Italic=   0   'False
         Sections(1).Cells(0).Style.Font_Underline=   0   'False
         Sections(1).Cells(0).Style.Font_Strikeout=   0   'False
         Sections(1).Cells(0).Style.Font_Charset=   0
         Sections(1).Cells(0).Style.TextAlign=   0
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
         Sections(1).Cells(0).Style.MarginTop=   6
         Sections(1).Cells(0).Style.MarginRight=   6
         Sections(1).Cells(0).Style.MarginBottom=   6
         Sections(1).Cells(0).Style.HasBorders=   -1  'True
         Sections(1).Cells(0).Style.BorderHT=   ""
         Sections(1).Cells(0).Style.BorderHI=   ""
         Sections(1).Cells(0).Style.BorderHB=   ""
         Sections(1).Cells(0).Style.BorderVL=   ""
         Sections(1).Cells(0).Style.BorderVI=   ""
         Sections(1).Cells(0).Style.BorderVR=   ""
         Sections(1).Cells(0).Style.NoClipping=   -1  'True
         Sections(1).Cells(0).Style.RTF=   0   'False
         Sections(1).Cells(0).Style.fprops=   16777216
         Sections(1).Cells(1).Name=   "CELL_1"
         Sections(1).Cells(1).Exp=   """: "" & faktur"
         Sections(1).Cells(1).Height=   4
         Sections(1).Cells(1).AutoHeight=   0   'False
         Sections(1).Cells(1).PrivateStyle=   -1  'True
         Sections(1).Cells(1).Style.Name=   "<private>"
         Sections(1).Cells(1).Style.ParentName=   "tdb_Base"
         Sections(1).Cells(1).Style.Font_Name=   "Arial"
         Sections(1).Cells(1).Style.Font_Size=   8.25
         Sections(1).Cells(1).Style.Font_Bold=   -1  'True
         Sections(1).Cells(1).Style.Font_Italic=   0   'False
         Sections(1).Cells(1).Style.Font_Underline=   0   'False
         Sections(1).Cells(1).Style.Font_Strikeout=   0   'False
         Sections(1).Cells(1).Style.Font_Charset=   0
         Sections(1).Cells(1).Style.TextAlign=   0
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
         Sections(1).Cells(1).Style.MarginTop=   6
         Sections(1).Cells(1).Style.MarginRight=   6
         Sections(1).Cells(1).Style.MarginBottom=   6
         Sections(1).Cells(1).Style.HasBorders=   -1  'True
         Sections(1).Cells(1).Style.BorderHT=   ""
         Sections(1).Cells(1).Style.BorderHI=   ""
         Sections(1).Cells(1).Style.BorderHB=   ""
         Sections(1).Cells(1).Style.BorderVL=   ""
         Sections(1).Cells(1).Style.BorderVI=   ""
         Sections(1).Cells(1).Style.BorderVR=   ""
         Sections(1).Cells(1).Style.NoClipping=   -1  'True
         Sections(1).Cells(1).Style.RTF=   0   'False
         Sections(1).Cells(1).Style.fprops=   16777216
         Sections(1).Cells(2).Name=   "CELL_3"
         Sections(1).Cells(2).Exp=   """KASIR"""
         Sections(1).Cells(2).NewLine=   -1  'True
         Sections(1).Cells(2).Width=   20
         Sections(1).Cells(2).Height=   4
         Sections(1).Cells(2).AutoHeight=   0   'False
         Sections(1).Cells(2).PrivateStyle=   -1  'True
         Sections(1).Cells(2).Style.Name=   "<private>"
         Sections(1).Cells(2).Style.ParentName=   "tdb_Base"
         Sections(1).Cells(2).Style.Font_Name=   "Arial"
         Sections(1).Cells(2).Style.Font_Size=   8.25
         Sections(1).Cells(2).Style.Font_Bold=   -1  'True
         Sections(1).Cells(2).Style.Font_Italic=   0   'False
         Sections(1).Cells(2).Style.Font_Underline=   0   'False
         Sections(1).Cells(2).Style.Font_Strikeout=   0   'False
         Sections(1).Cells(2).Style.Font_Charset=   0
         Sections(1).Cells(2).Style.TextAlign=   0
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
         Sections(1).Cells(2).Style.MarginTop=   6
         Sections(1).Cells(2).Style.MarginRight=   6
         Sections(1).Cells(2).Style.MarginBottom=   6
         Sections(1).Cells(2).Style.HasBorders=   -1  'True
         Sections(1).Cells(2).Style.BorderHT=   ""
         Sections(1).Cells(2).Style.BorderHI=   ""
         Sections(1).Cells(2).Style.BorderHB=   ""
         Sections(1).Cells(2).Style.BorderVL=   ""
         Sections(1).Cells(2).Style.BorderVI=   ""
         Sections(1).Cells(2).Style.BorderVR=   ""
         Sections(1).Cells(2).Style.NoClipping=   -1  'True
         Sections(1).Cells(2).Style.RTF=   0   'False
         Sections(1).Cells(2).Style.fprops=   16777216
         Sections(1).Cells(3).Name=   "CELL_4"
         Sections(1).Cells(3).Exp=   """: "" & kasir"
         Sections(1).Cells(3).Height=   4
         Sections(1).Cells(3).AutoHeight=   0   'False
         Sections(1).Cells(3).PrivateStyle=   -1  'True
         Sections(1).Cells(3).Style.Name=   "<private>"
         Sections(1).Cells(3).Style.ParentName=   "tdb_Base"
         Sections(1).Cells(3).Style.Font_Name=   "Arial"
         Sections(1).Cells(3).Style.Font_Size=   8.25
         Sections(1).Cells(3).Style.Font_Bold=   -1  'True
         Sections(1).Cells(3).Style.Font_Italic=   0   'False
         Sections(1).Cells(3).Style.Font_Underline=   0   'False
         Sections(1).Cells(3).Style.Font_Strikeout=   0   'False
         Sections(1).Cells(3).Style.Font_Charset=   0
         Sections(1).Cells(3).Style.TextAlign=   0
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
         Sections(1).Cells(3).Style.MarginTop=   6
         Sections(1).Cells(3).Style.MarginRight=   6
         Sections(1).Cells(3).Style.MarginBottom=   6
         Sections(1).Cells(3).Style.HasBorders=   -1  'True
         Sections(1).Cells(3).Style.BorderHT=   ""
         Sections(1).Cells(3).Style.BorderHI=   ""
         Sections(1).Cells(3).Style.BorderHB=   ""
         Sections(1).Cells(3).Style.BorderVL=   ""
         Sections(1).Cells(3).Style.BorderVI=   ""
         Sections(1).Cells(3).Style.BorderVR=   ""
         Sections(1).Cells(3).Style.NoClipping=   -1  'True
         Sections(1).Cells(3).Style.RTF=   0   'False
         Sections(1).Cells(3).Style.fprops=   16777216
         Sections(1).Cells(4).Name=   "CELL_6"
         Sections(1).Cells(4).Exp=   """TANGGAL TRANSAKSI"""
         Sections(1).Cells(4).NewLine=   -1  'True
         Sections(1).Cells(4).Width=   20
         Sections(1).Cells(4).Height=   4
         Sections(1).Cells(4).AutoHeight=   0   'False
         Sections(1).Cells(4).PrivateStyle=   -1  'True
         Sections(1).Cells(4).Style.Name=   "<private>"
         Sections(1).Cells(4).Style.ParentName=   "tdb_Base"
         Sections(1).Cells(4).Style.Font_Name=   "Arial"
         Sections(1).Cells(4).Style.Font_Size=   8.25
         Sections(1).Cells(4).Style.Font_Bold=   -1  'True
         Sections(1).Cells(4).Style.Font_Italic=   0   'False
         Sections(1).Cells(4).Style.Font_Underline=   0   'False
         Sections(1).Cells(4).Style.Font_Strikeout=   0   'False
         Sections(1).Cells(4).Style.Font_Charset=   0
         Sections(1).Cells(4).Style.TextAlign=   0
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
         Sections(1).Cells(4).Style.MarginTop=   6
         Sections(1).Cells(4).Style.MarginRight=   6
         Sections(1).Cells(4).Style.MarginBottom=   6
         Sections(1).Cells(4).Style.HasBorders=   -1  'True
         Sections(1).Cells(4).Style.BorderHT=   ""
         Sections(1).Cells(4).Style.BorderHI=   ""
         Sections(1).Cells(4).Style.BorderHB=   ""
         Sections(1).Cells(4).Style.BorderVL=   ""
         Sections(1).Cells(4).Style.BorderVI=   ""
         Sections(1).Cells(4).Style.BorderVR=   ""
         Sections(1).Cells(4).Style.NoClipping=   -1  'True
         Sections(1).Cells(4).Style.RTF=   0   'False
         Sections(1).Cells(4).Style.fprops=   16777216
         Sections(1).Cells(5).Name=   "CELL_7"
         Sections(1).Cells(5).Exp=   """: "" & tgl"
         Sections(1).Cells(5).Height=   4
         Sections(1).Cells(5).AutoHeight=   0   'False
         Sections(1).Cells(5).PrivateStyle=   -1  'True
         Sections(1).Cells(5).Style.Name=   "<private>"
         Sections(1).Cells(5).Style.ParentName=   "tdb_Base"
         Sections(1).Cells(5).Style.Font_Name=   "Arial"
         Sections(1).Cells(5).Style.Font_Size=   8.25
         Sections(1).Cells(5).Style.Font_Bold=   -1  'True
         Sections(1).Cells(5).Style.Font_Italic=   0   'False
         Sections(1).Cells(5).Style.Font_Underline=   0   'False
         Sections(1).Cells(5).Style.Font_Strikeout=   0   'False
         Sections(1).Cells(5).Style.Font_Charset=   0
         Sections(1).Cells(5).Style.TextAlign=   0
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
         Sections(1).Cells(5).Style.MarginTop=   6
         Sections(1).Cells(5).Style.MarginRight=   6
         Sections(1).Cells(5).Style.MarginBottom=   6
         Sections(1).Cells(5).Style.HasBorders=   -1  'True
         Sections(1).Cells(5).Style.BorderHT=   ""
         Sections(1).Cells(5).Style.BorderHI=   ""
         Sections(1).Cells(5).Style.BorderHB=   ""
         Sections(1).Cells(5).Style.BorderVL=   ""
         Sections(1).Cells(5).Style.BorderVI=   ""
         Sections(1).Cells(5).Style.BorderVR=   ""
         Sections(1).Cells(5).Style.NoClipping=   -1  'True
         Sections(1).Cells(5).Style.RTF=   0   'False
         Sections(1).Cells(5).Style.fprops=   16777216
         Sections(2).Name=   "DetailHeader"
         Sections(2).Type=   3
         Sections(2).StyleExp=   "tdb_TableHeader"
         Sections(2).Tabulator=   "Detail"
         Sections(2).Cells.Count=   7
         Sections(2).Cells(0).Name=   "Nomor"
         Sections(2).Cells(0).Exp=   """No."""
         Sections(2).Cells(1).Name=   "Kode"
         Sections(2).Cells(1).Exp=   """KODE """
         Sections(2).Cells(2).Name=   "JudulBuku"
         Sections(2).Cells(2).Exp=   """NAMA STOCK"""
         Sections(2).Cells(3).Name=   "Jumlah"
         Sections(2).Cells(3).Exp=   """QTY"""
         Sections(2).Cells(4).Name=   "Harga"
         Sections(2).Cells(4).Exp=   """HARGA"""
         Sections(2).Cells(5).Name=   "Discount"
         Sections(2).Cells(5).Exp=   """DISC%"""
         Sections(2).Cells(6).Name=   "Qty"
         Sections(2).Cells(6).Exp=   """JUMLAH"""
         Sections(3).Name=   "Detail"
         Sections(3).Type=   4
         Sections(3).StyleExp=   "'tdb_TableOddRow'"
         Sections(3).Cells.Count=   7
         Sections(3).Cells(0).Name=   "No"
         Sections(3).Cells(0).Exp=   "Sum(1,WillChange(faktur))"
         Sections(3).Cells(0).Width=   4
         Sections(3).Cells(1).Name=   "Kode"
         Sections(3).Cells(1).Exp=   "kode"
         Sections(3).Cells(1).Width=   13
         Sections(3).Cells(1).PrivateStyle=   -1  'True
         Sections(3).Cells(1).Style.Name=   "<private>"
         Sections(3).Cells(1).Style.ParentName=   "tdb_TableOddRow"
         Sections(3).Cells(1).Style.Font_Name=   "Arial"
         Sections(3).Cells(1).Style.Font_Size=   8.25
         Sections(3).Cells(1).Style.Font_Bold=   0   'False
         Sections(3).Cells(1).Style.Font_Italic=   0   'False
         Sections(3).Cells(1).Style.Font_Underline=   0   'False
         Sections(3).Cells(1).Style.Font_Strikeout=   0   'False
         Sections(3).Cells(1).Style.Font_Charset=   0
         Sections(3).Cells(1).Style.TextAlign=   1
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
         Sections(3).Cells(1).Style.BorderHT=   "Quarter"
         Sections(3).Cells(1).Style.BorderHI=   "Quarter"
         Sections(3).Cells(1).Style.BorderHB=   "Double"
         Sections(3).Cells(1).Style.BorderVL=   "Single"
         Sections(3).Cells(1).Style.BorderVI=   "Single"
         Sections(3).Cells(1).Style.BorderVR=   "Single"
         Sections(3).Cells(1).Style.NoClipping=   -1  'True
         Sections(3).Cells(1).Style.RTF=   0   'False
         Sections(3).Cells(1).Style.fprops=   1
         Sections(3).Cells(2).Name=   "JudulBuku"
         Sections(3).Cells(2).Exp=   "namabarang"
         Sections(3).Cells(2).Width=   44
         Sections(3).Cells(3).Name=   "QTy"
         Sections(3).Cells(3).Exp=   "qty & "" "" & satuan"
         Sections(3).Cells(3).Width=   13
         Sections(3).Cells(4).Name=   "Harga"
         Sections(3).Cells(4).Exp=   "harga"
         Sections(3).Cells(4).Width=   15
         Sections(3).Cells(4).Format=   "###,###,##0.00"
         Sections(3).Cells(5).Name=   "Discount"
         Sections(3).Cells(5).Width=   6
         Sections(3).Cells(6).Name=   "Total"
         Sections(3).Cells(6).Exp=   "jumlah"
         Sections(3).Cells(6).Width=   15
         Sections(3).Cells(6).Format=   "###,###,##0.00"
         Sections(4).Name=   "SECTION_2"
         Sections(4).Type=   5
         Sections(4).StyleExp=   "'Tdb_TableFooter'"
         Sections(4).Cells.Count=   12
         Sections(4).Cells(0).Name=   "SubTotal"
         Sections(4).Cells(0).Exp=   """Sub Total"""
         Sections(4).Cells(0).NewLine=   -1  'True
         Sections(4).Cells(0).Width=   80
         Sections(4).Cells(0).Height=   4
         Sections(4).Cells(0).AutoHeight=   0   'False
         Sections(4).Cells(1).Name=   "CELL_1"
         Sections(4).Cells(1).Exp=   """: Rp"""
         Sections(4).Cells(1).Width=   5
         Sections(4).Cells(1).Height=   4
         Sections(4).Cells(1).AutoHeight=   0   'False
         Sections(4).Cells(2).Name=   "CELL_2"
         Sections(4).Cells(2).Exp=   "subtotal"
         Sections(4).Cells(2).Height=   4
         Sections(4).Cells(2).AutoHeight=   0   'False
         Sections(4).Cells(2).Format=   "###,###,##0.00"
         Sections(4).Cells(3).Name=   "Discount1"
         Sections(4).Cells(3).Exp=   """Discount"""
         Sections(4).Cells(3).NewLine=   -1  'True
         Sections(4).Cells(3).Width=   80
         Sections(4).Cells(3).Height=   4
         Sections(4).Cells(3).AutoHeight=   0   'False
         Sections(4).Cells(4).Name=   "CELL_4"
         Sections(4).Cells(4).Exp=   """: Rp"""
         Sections(4).Cells(4).Width=   5
         Sections(4).Cells(4).Height=   4
         Sections(4).Cells(4).AutoHeight=   0   'False
         Sections(4).Cells(5).Name=   "CELL_5"
         Sections(4).Cells(5).Exp=   "discount"
         Sections(4).Cells(5).Height=   4
         Sections(4).Cells(5).AutoHeight=   0   'False
         Sections(4).Cells(5).Format=   "###,###,##0.00"
         Sections(4).Cells(6).Name=   "Pajak"
         Sections(4).Cells(6).Exp=   """Pajak"""
         Sections(4).Cells(6).NewLine=   -1  'True
         Sections(4).Cells(6).Width=   80
         Sections(4).Cells(6).Height=   4
         Sections(4).Cells(6).AutoHeight=   0   'False
         Sections(4).Cells(7).Name=   "CELL_10"
         Sections(4).Cells(7).Exp=   """: Rp"""
         Sections(4).Cells(7).Width=   5
         Sections(4).Cells(8).Name=   "CELL_11"
         Sections(4).Cells(8).Width=   15
         Sections(4).Cells(8).Height=   4
         Sections(4).Cells(8).AutoHeight=   0   'False
         Sections(4).Cells(8).Format=   "###,###,##0.00"
         Sections(4).Cells(9).Name=   "GrandTotal"
         Sections(4).Cells(9).Exp=   """Grand Total"""
         Sections(4).Cells(9).NewLine=   -1  'True
         Sections(4).Cells(9).Width=   80
         Sections(4).Cells(9).Height=   4
         Sections(4).Cells(9).AutoHeight=   0   'False
         Sections(4).Cells(10).Name=   "CELL_13"
         Sections(4).Cells(10).Exp=   """: Rp"""
         Sections(4).Cells(10).StyleExp=   "'Tdb_FooterGarisBawah'"
         Sections(4).Cells(10).Width=   5
         Sections(4).Cells(10).Height=   4
         Sections(4).Cells(10).AutoHeight=   0   'False
         Sections(4).Cells(11).Name=   "CELL_14"
         Sections(4).Cells(11).Exp=   "total"
         Sections(4).Cells(11).StyleExp=   "'Tdb_FooterGarisBawah'"
         Sections(4).Cells(11).Height=   4
         Sections(4).Cells(11).AutoHeight=   0   'False
         Sections(4).Cells(11).Format=   "###,###,##0.00"
         Sections(5).Name=   "SECTION_6"
         Sections(5).Condition=   "IsLastRec()"
         Sections(5).StyleExp=   "'total'"
         Sections(5).AutoHeight=   0   'False
         Sections(5).Height=   5
         Sections(5).Cells.Count=   17
         Sections(5).Cells(0).Name=   "CELL_0"
         Sections(5).Cells(0).Exp=   """ """
         Sections(5).Cells(0).NewLine=   -1  'True
         Sections(5).Cells(1).Name=   "CELL_1"
         Sections(5).Cells(1).Exp=   """JUMLAH KESELURUHAN :"""
         Sections(5).Cells(1).NewLine=   -1  'True
         Sections(5).Cells(2).Name=   "CELL_2"
         Sections(5).Cells(2).Exp=   """JUMLAH BARANG (QTY)"""
         Sections(5).Cells(2).NewLine=   -1  'True
         Sections(5).Cells(2).Width=   35
         Sections(5).Cells(2).PrivateStyle=   -1  'True
         Sections(5).Cells(2).Style.Name=   "<private>"
         Sections(5).Cells(2).Style.ParentName=   "total"
         Sections(5).Cells(2).Style.Font_Name=   "MS Sans Serif"
         Sections(5).Cells(2).Style.Font_Size=   8.25
         Sections(5).Cells(2).Style.Font_Bold=   -1  'True
         Sections(5).Cells(2).Style.Font_Italic=   0   'False
         Sections(5).Cells(2).Style.Font_Underline=   0   'False
         Sections(5).Cells(2).Style.Font_Strikeout=   0   'False
         Sections(5).Cells(2).Style.Font_Charset=   0
         Sections(5).Cells(2).Style.TextAlign=   0
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
         Sections(5).Cells(2).Style.MarginTop=   6
         Sections(5).Cells(2).Style.MarginRight=   6
         Sections(5).Cells(2).Style.MarginBottom=   6
         Sections(5).Cells(2).Style.HasBorders=   -1  'True
         Sections(5).Cells(2).Style.BorderHT=   ""
         Sections(5).Cells(2).Style.BorderHI=   ""
         Sections(5).Cells(2).Style.BorderHB=   ""
         Sections(5).Cells(2).Style.BorderVL=   ""
         Sections(5).Cells(2).Style.BorderVI=   ""
         Sections(5).Cells(2).Style.BorderVR=   ""
         Sections(5).Cells(2).Style.NoClipping=   -1  'True
         Sections(5).Cells(2).Style.RTF=   0   'False
         Sections(5).Cells(2).Style.fprops=   0
         Sections(5).Cells(3).Name=   "CELL_3"
         Sections(5).Cells(3).Exp=   """  : """
         Sections(5).Cells(3).Width=   3
         Sections(5).Cells(4).Name=   "CELL_4"
         Sections(5).Cells(4).Exp=   "TOTALQTY"
         Sections(5).Cells(4).Width=   20
         Sections(5).Cells(4).PrivateStyle=   -1  'True
         Sections(5).Cells(4).Format=   "###,###,##0"
         Sections(5).Cells(4).Style.Name=   "<private>"
         Sections(5).Cells(4).Style.ParentName=   "total"
         Sections(5).Cells(4).Style.Font_Name=   "MS Sans Serif"
         Sections(5).Cells(4).Style.Font_Size=   8.25
         Sections(5).Cells(4).Style.Font_Bold=   -1  'True
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
         Sections(5).Cells(4).Style.MarginTop=   6
         Sections(5).Cells(4).Style.MarginRight=   6
         Sections(5).Cells(4).Style.MarginBottom=   6
         Sections(5).Cells(4).Style.HasBorders=   -1  'True
         Sections(5).Cells(4).Style.BorderHT=   ""
         Sections(5).Cells(4).Style.BorderHI=   ""
         Sections(5).Cells(4).Style.BorderHB=   ""
         Sections(5).Cells(4).Style.BorderVL=   ""
         Sections(5).Cells(4).Style.BorderVI=   ""
         Sections(5).Cells(4).Style.BorderVR=   ""
         Sections(5).Cells(4).Style.NoClipping=   -1  'True
         Sections(5).Cells(4).Style.RTF=   0   'False
         Sections(5).Cells(4).Style.fprops=   1
         Sections(5).Cells(5).Name=   "CELL_5"
         Sections(5).Cells(5).Exp=   """SUB TOTAL PENJUALAN KASIR (Rp)"""
         Sections(5).Cells(5).NewLine=   -1  'True
         Sections(5).Cells(5).Width=   35
         Sections(5).Cells(6).Name=   "CELL_6"
         Sections(5).Cells(6).Exp=   """ : """
         Sections(5).Cells(6).Width=   3
         Sections(5).Cells(7).Name=   "CELL_7"
         Sections(5).Cells(7).Exp=   "TOTALJUMLAH"
         Sections(5).Cells(7).Width=   20
         Sections(5).Cells(7).PrivateStyle=   -1  'True
         Sections(5).Cells(7).Format=   "###,###,###,###,##0.00"
         Sections(5).Cells(7).Style.Name=   "<private>"
         Sections(5).Cells(7).Style.ParentName=   "total"
         Sections(5).Cells(7).Style.Font_Name=   "MS Sans Serif"
         Sections(5).Cells(7).Style.Font_Size=   8.25
         Sections(5).Cells(7).Style.Font_Bold=   -1  'True
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
         Sections(5).Cells(7).Style.MarginTop=   6
         Sections(5).Cells(7).Style.MarginRight=   6
         Sections(5).Cells(7).Style.MarginBottom=   6
         Sections(5).Cells(7).Style.HasBorders=   -1  'True
         Sections(5).Cells(7).Style.BorderHT=   ""
         Sections(5).Cells(7).Style.BorderHI=   ""
         Sections(5).Cells(7).Style.BorderHB=   ""
         Sections(5).Cells(7).Style.BorderVL=   ""
         Sections(5).Cells(7).Style.BorderVI=   ""
         Sections(5).Cells(7).Style.BorderVR=   ""
         Sections(5).Cells(7).Style.NoClipping=   -1  'True
         Sections(5).Cells(7).Style.RTF=   0   'False
         Sections(5).Cells(7).Style.fprops=   1
         Sections(5).Cells(8).Name=   "CELL_8"
         Sections(5).Cells(8).Exp=   """TOTAL DISCOUNT"""
         Sections(5).Cells(8).NewLine=   -1  'True
         Sections(5).Cells(8).Width=   35
         Sections(5).Cells(9).Name=   "CELL_9"
         Sections(5).Cells(9).Exp=   """ : """
         Sections(5).Cells(9).Width=   3
         Sections(5).Cells(10).Name=   "CELL_10"
         Sections(5).Cells(10).Exp=   "TOTALDISCOUNT1"
         Sections(5).Cells(10).Width=   20
         Sections(5).Cells(10).PrivateStyle=   -1  'True
         Sections(5).Cells(10).Format=   "###,###,###,###,##0.00"
         Sections(5).Cells(10).Style.Name=   "<private>"
         Sections(5).Cells(10).Style.ParentName=   "total"
         Sections(5).Cells(10).Style.Font_Name=   "MS Sans Serif"
         Sections(5).Cells(10).Style.Font_Size=   8.25
         Sections(5).Cells(10).Style.Font_Bold=   -1  'True
         Sections(5).Cells(10).Style.Font_Italic=   0   'False
         Sections(5).Cells(10).Style.Font_Underline=   0   'False
         Sections(5).Cells(10).Style.Font_Strikeout=   0   'False
         Sections(5).Cells(10).Style.Font_Charset=   0
         Sections(5).Cells(10).Style.TextAlign=   2
         Sections(5).Cells(10).Style.TextVAlign=   1
         Sections(5).Cells(10).Style.TextWrap=   -1  'True
         Sections(5).Cells(10).Style.ForeColor=   0
         Sections(5).Cells(10).Style.BackColor=   16777215
         Sections(5).Cells(10).Style.NoFill=   -1  'True
         Sections(5).Cells(10).Style.BackPicFile=   ""
         Sections(5).Cells(10).Style.ForePicFile=   ""
         Sections(5).Cells(10).Style.BackPicVertPlacement=   0
         Sections(5).Cells(10).Style.BackPicHorzPlacement=   0
         Sections(5).Cells(10).Style.ForePicPlacement=   0
         Sections(5).Cells(10).Style.ForePicDrawMode=   0
         Sections(5).Cells(10).Style.MarginLeft=   6
         Sections(5).Cells(10).Style.MarginTop=   6
         Sections(5).Cells(10).Style.MarginRight=   6
         Sections(5).Cells(10).Style.MarginBottom=   6
         Sections(5).Cells(10).Style.HasBorders=   -1  'True
         Sections(5).Cells(10).Style.BorderHT=   ""
         Sections(5).Cells(10).Style.BorderHI=   ""
         Sections(5).Cells(10).Style.BorderHB=   ""
         Sections(5).Cells(10).Style.BorderVL=   ""
         Sections(5).Cells(10).Style.BorderVI=   ""
         Sections(5).Cells(10).Style.BorderVR=   ""
         Sections(5).Cells(10).Style.NoClipping=   -1  'True
         Sections(5).Cells(10).Style.RTF=   0   'False
         Sections(5).Cells(10).Style.fprops=   1
         Sections(5).Cells(11).Name=   "CELL_14"
         Sections(5).Cells(11).Exp=   """TOTAL PAJAK"""
         Sections(5).Cells(11).NewLine=   -1  'True
         Sections(5).Cells(11).Width=   35
         Sections(5).Cells(12).Name=   "CELL_15"
         Sections(5).Cells(12).Exp=   """ : """
         Sections(5).Cells(12).Width=   3
         Sections(5).Cells(13).Name=   "CELL_16"
         Sections(5).Cells(13).Exp=   "TOTALPAJAK"
         Sections(5).Cells(13).Width=   20
         Sections(5).Cells(13).PrivateStyle=   -1  'True
         Sections(5).Cells(13).Format=   "###,###,###,###,##0.00"
         Sections(5).Cells(13).Style.Name=   "<private>"
         Sections(5).Cells(13).Style.ParentName=   "total"
         Sections(5).Cells(13).Style.Font_Name=   "MS Sans Serif"
         Sections(5).Cells(13).Style.Font_Size=   8.25
         Sections(5).Cells(13).Style.Font_Bold=   -1  'True
         Sections(5).Cells(13).Style.Font_Italic=   0   'False
         Sections(5).Cells(13).Style.Font_Underline=   0   'False
         Sections(5).Cells(13).Style.Font_Strikeout=   0   'False
         Sections(5).Cells(13).Style.Font_Charset=   0
         Sections(5).Cells(13).Style.TextAlign=   2
         Sections(5).Cells(13).Style.TextVAlign=   1
         Sections(5).Cells(13).Style.TextWrap=   -1  'True
         Sections(5).Cells(13).Style.ForeColor=   0
         Sections(5).Cells(13).Style.BackColor=   16777215
         Sections(5).Cells(13).Style.NoFill=   -1  'True
         Sections(5).Cells(13).Style.BackPicFile=   ""
         Sections(5).Cells(13).Style.ForePicFile=   ""
         Sections(5).Cells(13).Style.BackPicVertPlacement=   0
         Sections(5).Cells(13).Style.BackPicHorzPlacement=   0
         Sections(5).Cells(13).Style.ForePicPlacement=   0
         Sections(5).Cells(13).Style.ForePicDrawMode=   0
         Sections(5).Cells(13).Style.MarginLeft=   6
         Sections(5).Cells(13).Style.MarginTop=   6
         Sections(5).Cells(13).Style.MarginRight=   6
         Sections(5).Cells(13).Style.MarginBottom=   6
         Sections(5).Cells(13).Style.HasBorders=   -1  'True
         Sections(5).Cells(13).Style.BorderHT=   ""
         Sections(5).Cells(13).Style.BorderHI=   ""
         Sections(5).Cells(13).Style.BorderHB=   ""
         Sections(5).Cells(13).Style.BorderVL=   ""
         Sections(5).Cells(13).Style.BorderVI=   ""
         Sections(5).Cells(13).Style.BorderVR=   ""
         Sections(5).Cells(13).Style.NoClipping=   -1  'True
         Sections(5).Cells(13).Style.RTF=   0   'False
         Sections(5).Cells(13).Style.fprops=   1
         Sections(5).Cells(14).Name=   "CELL_17"
         Sections(5).Cells(14).Exp=   """TOTAL PENJUALAN KASIR"""
         Sections(5).Cells(14).NewLine=   -1  'True
         Sections(5).Cells(14).Width=   35
         Sections(5).Cells(15).Name=   "CELL_18"
         Sections(5).Cells(15).Exp=   """ : """
         Sections(5).Cells(15).Width=   3
         Sections(5).Cells(16).Name=   "CELL_19"
         Sections(5).Cells(16).Exp=   "GRANDTOTAL"
         Sections(5).Cells(16).Width=   20
         Sections(5).Cells(16).PrivateStyle=   -1  'True
         Sections(5).Cells(16).Format=   "###,###,###,###,##0.00"
         Sections(5).Cells(16).Style.Name=   "<private>"
         Sections(5).Cells(16).Style.ParentName=   "total"
         Sections(5).Cells(16).Style.Font_Name=   "MS Sans Serif"
         Sections(5).Cells(16).Style.Font_Size=   8.25
         Sections(5).Cells(16).Style.Font_Bold=   -1  'True
         Sections(5).Cells(16).Style.Font_Italic=   0   'False
         Sections(5).Cells(16).Style.Font_Underline=   0   'False
         Sections(5).Cells(16).Style.Font_Strikeout=   0   'False
         Sections(5).Cells(16).Style.Font_Charset=   0
         Sections(5).Cells(16).Style.TextAlign=   2
         Sections(5).Cells(16).Style.TextVAlign=   1
         Sections(5).Cells(16).Style.TextWrap=   -1  'True
         Sections(5).Cells(16).Style.ForeColor=   0
         Sections(5).Cells(16).Style.BackColor=   16777215
         Sections(5).Cells(16).Style.NoFill=   -1  'True
         Sections(5).Cells(16).Style.BackPicFile=   ""
         Sections(5).Cells(16).Style.ForePicFile=   ""
         Sections(5).Cells(16).Style.BackPicVertPlacement=   0
         Sections(5).Cells(16).Style.BackPicHorzPlacement=   0
         Sections(5).Cells(16).Style.ForePicPlacement=   0
         Sections(5).Cells(16).Style.ForePicDrawMode=   0
         Sections(5).Cells(16).Style.MarginLeft=   6
         Sections(5).Cells(16).Style.MarginTop=   6
         Sections(5).Cells(16).Style.MarginRight=   6
         Sections(5).Cells(16).Style.MarginBottom=   6
         Sections(5).Cells(16).Style.HasBorders=   -1  'True
         Sections(5).Cells(16).Style.BorderHT=   ""
         Sections(5).Cells(16).Style.BorderHI=   ""
         Sections(5).Cells(16).Style.BorderHB=   ""
         Sections(5).Cells(16).Style.BorderVL=   ""
         Sections(5).Cells(16).Style.BorderVI=   ""
         Sections(5).Cells(16).Style.BorderVR=   ""
         Sections(5).Cells(16).Style.NoClipping=   -1  'True
         Sections(5).Cells(16).Style.RTF=   0   'False
         Sections(5).Cells(16).Style.fprops=   1
         Styles.Count    =   7
         Styles(0).Name  =   "tdb_Base"
         Styles(0).ParentName=   ""
         Styles(0).Font_Name=   "Arial"
         Styles(0).Font_Size=   8.25
         Styles(0).Font_Charset=   0
         Styles(0).TextAlign=   0
         Styles(0).TextVAlign=   1
         Styles(0).NoClipping=   -1  'True
         Styles(1).Name  =   "Tdb_FooterGarisBawah"
         Styles(1).ParentName=   "tdb_Base"
         Styles(1).Font_Name=   "Arial"
         Styles(1).Font_Size=   8.25
         Styles(1).Font_Charset=   0
         Styles(1).TextAlign=   2
         Styles(1).TextVAlign=   1
         Styles(1).BorderHT=   "Double"
         Styles(1).NoClipping=   -1  'True
         Styles(1).fprops=   163841
         Styles(2).Name  =   "tdb_PageHeader"
         Styles(2).ParentName=   "tdb_Base"
         Styles(2).Font_Name=   "Arial"
         Styles(2).Font_Size=   8.25
         Styles(2).Font_Charset=   0
         Styles(2).TextAlign=   2
         Styles(2).TextVAlign=   1
         Styles(2).NoClipping=   -1  'True
         Styles(2).fprops=   1
         Styles(3).Name  =   "tdb_TableOddRow"
         Styles(3).ParentName=   "tdb_Base"
         Styles(3).Font_Name=   "Arial"
         Styles(3).Font_Size=   8.25
         Styles(3).Font_Charset=   0
         Styles(3).BorderHT=   "Quarter"
         Styles(3).BorderHI=   "Quarter"
         Styles(3).BorderHB=   "Double"
         Styles(3).BorderVL=   "Single"
         Styles(3).BorderVI=   "Single"
         Styles(3).BorderVR=   "Single"
         Styles(3).NoClipping=   -1  'True
         Styles(4).Name  =   "tdb_TableHeader"
         Styles(4).ParentName=   "tdb_Base"
         Styles(4).Font_Name=   "Arial"
         Styles(4).Font_Size=   8.25
         Styles(4).Font_Bold=   -1  'True
         Styles(4).Font_Charset=   0
         Styles(4).TextAlign=   1
         Styles(4).TextVAlign=   1
         Styles(4).ForeColor=   4194304
         Styles(4).NoFill=   0   'False
         Styles(4).BorderHT=   "Double"
         Styles(4).BorderHI=   "Double"
         Styles(4).BorderHB=   "Double"
         Styles(4).BorderVL=   "Single"
         Styles(4).BorderVI=   "Single"
         Styles(4).BorderVR=   "Single"
         Styles(4).NoClipping=   -1  'True
         Styles(5).Name  =   "Tdb_TableFooter"
         Styles(5).ParentName=   "tdb_Base"
         Styles(5).Font_Name=   "Arial"
         Styles(5).Font_Size=   8.25
         Styles(5).Font_Charset=   0
         Styles(5).TextAlign=   2
         Styles(5).TextVAlign=   1
         Styles(5).NoClipping=   -1  'True
         Styles(5).fprops=   3
         Styles(6).Name  =   "total"
         Styles(6).ParentName=   "tdb_Base"
         Styles(6).Font_Name=   "MS Sans Serif"
         Styles(6).Font_Size=   8.25
         Styles(6).Font_Bold=   -1  'True
         Styles(6).Font_Charset=   0
         Styles(6).TextAlign=   0
         Styles(6).TextVAlign=   1
         Styles(6).NoClipping=   -1  'True
         Styles(6).fprops=   23068672
         Lines.Count     =   3
         Lines(0).Name   =   "Single"
         Lines(0).Thickness=   4
         Lines(1).Name   =   "Double"
         Lines(1).Thickness=   5
         Lines(2).Name   =   "Quarter"
         Lines(2).Thickness=   1
         Lines(2).Color  =   8421504
         Profiles.Count  =   1
         Profiles(0).Name=   "PROFILE_0"
         Profiles(0).Active=   -1  'True
         Profiles(0).PreviewNoMinimize=   -1  'True
         Profiles(0).PreviewNoMaximize=   -1  'True
         Profiles(0).PreviewNoResize=   -1  'True
         Profiles(0).PreviewMaximized=   -1  'True
         Profiles(0).PreviewNoSaveLoad=   -1  'True
         Profiles(0).PrinterMarginLeft=   10
         Profiles(0).PrinterMarginTop=   10
         Profiles(0).PrinterMarginRight=   10
         Profiles(0).PrinterMarginBottom=   10
         Profiles(0).PrinterMargins_set=   -1  'True
         Profiles(0).PrinterPaperUserSize_set=   -1  'True
      End
      Begin VB.Label Label2 
         Caption         =   "Tampilkan"
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
         TabIndex        =   10
         Top             =   1185
         Width           =   1065
      End
      Begin VB.Label Label1 
         Caption         =   "Kasir"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1770
         TabIndex        =   6
         Top             =   675
         Width           =   495
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   645
      Left            =   0
      Top             =   1725
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   1138
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
         Top             =   105
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
         Picture         =   "rptPenjualanKasir.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   5370
         TabIndex        =   3
         Top             =   105
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
         Picture         =   "rptPenjualanKasir.frx":00A6
      End
   End
End
Attribute VB_Name = "rptPenjualanKasir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB

Private Sub Check1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    keybd_event VK_TAB, 0, 0, 0
  End If
End Sub

Private Sub cKasir_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "username", "username,fullname")
  If Not dbData.EOF Then
    cKasir.Text = cKasir.Browse(dbData)
  End If
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdPreview_Click()
  GetData
End Sub

Private Sub Form_Load()
Dim n As Single

  SetIcon Me.hWnd, "SIKD"
  CenterForm Me
  dTgl(0).Value = BOM(Date)
  optKodeStock(1).Value = True
  
  TabIndex dTgl(0), n
  TabIndex dTgl(1), n
  TabIndex Check1, n
  TabIndex cKasir, n
  TabIndex optKodeStock(0), n
  TabIndex optKodeStock(1), n
  
  TabIndex cmdPreview, n
  TabIndex cmdKeluar, n
End Sub

Private Sub GetData()
Dim cSQL As String
Dim n As Integer
Dim nJumlah As Double
Dim nDiscount1 As Double
Dim nDiscount2 As Double
Dim nPajak As Double
Dim nTotal As Double
Dim nQty As Double
  
  vaArray.ReDim 0, -1, 0, 11
  cSQL = "SELECT p.nomorkasir,s.barcode,p.kodestock,p.qty,s.kodesatuan,p.harga,p.jumlah,"
  cSQL = cSQL & " t.tgl,t.subtotal,t.discount,t.total,t.username,"
  cSQL = cSQL & " s.nama as namabarang"
  cSQL = cSQL & " FROM kasir p"
  cSQL = cSQL & " LEFT JOIN totkasir t on t.nomorkasir = p.nomorkasir"
  cSQL = cSQL & " LEFT JOIN stock s on s.kodestock = p.kodestock"
  cSQL = cSQL & " WHERE t.tgl >='" & Format(dTgl(0).Value, "yyyy-MM-dd") & "'"
  cSQL = cSQL & " AND t.tgl <='" & Format(dTgl(1).Value, "yyyy-MM-dd") & "'"
  If Check1.Value = 1 Then
    cSQL = cSQL & " AND t.username = '" & cKasir.Text & "'"
  End If
  cSQL = cSQL & " ORDER BY t.nomorkasir"
  Set dbData = objData.Sql(GetDSN, cSQL)
  If Not dbData.EOF Then
    FrmPB.InitPB dbData.RecordCount
    dbData.MoveFirst
    Do While Not dbData.EOF
      FrmPB.RunPB
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      
      vaArray(n, 0) = GetNull((dbData!nomorkasir), "")
      vaArray(n, 1) = Format(GetNull((dbData!tgl), ""), "dd-MM-yyyy")
      vaArray(n, 2) = IIf(optKodeStock(0).Value = True, GetNull((dbData!KodeStock), ""), GetNull((dbData!Barcode), ""))
      vaArray(n, 3) = GetNull((dbData!Namabarang))
      vaArray(n, 4) = GetNull(dbData!qty)
      vaArray(n, 5) = GetNull((dbData!kodesatuan), "")
      vaArray(n, 6) = GetNull(dbData!Harga)
      vaArray(n, 7) = GetNull(dbData!Jumlah)
      vaArray(n, 8) = GetNull(dbData!Subtotal)
      vaArray(n, 9) = GetNull(dbData!Discount)
      vaArray(n, 10) = GetNull(dbData!Total)
      vaArray(n, 11) = GetNull(dbData!UserName)
      nQty = nQty + vaArray(n, 4)
      dbData.MoveNext
    Loop
    FrmPB.EndPB
    tdb.Parameters("TGL1").ValueExpression = dTgl(0).Value
    tdb.Parameters("TGL2").ValueExpression = dTgl(1).Value
    
    GetSUM nJumlah, nDiscount1, nDiscount2, nPajak, nTotal
    
    tdb.Parameters("TOTALJUMLAH").ValueExpression = nJumlah
    tdb.Parameters("TOTALQTY").ValueExpression = nQty
    tdb.Parameters("TOTALDISCOUNT1").ValueExpression = nDiscount1
    tdb.Parameters("TOTALPAJAK").ValueExpression = nPajak
    tdb.Parameters("GRANDTOTAL").ValueExpression = nTotal
    Set tdb.Array = vaArray
    tdb.Refresh
    tdb.PrintPreview
  End If
End Sub

Private Sub GetSUM(ByRef nSubTotal As Double, ByRef nDisc As Double, _
                    ByRef nDisc1 As Double, ByRef nPajak As Double, _
                    ByRef nTotal As Double)
  
Dim cSQL As String
  
  
  nSubTotal = 0
  nDisc = 0
  nDisc1 = 0
  nPajak = 0
  nTotal = 0
  
  cSQL = "SELECT SUM(subtotal) as SubTotal,SUM(discount) as Disc, SUM(total) as Total"
  cSQL = cSQL & " FROM totkasir"
  cSQL = cSQL & " WHERE Tgl >='" & Format(dTgl(0).Value, "yyyy-MM-dd") & "'"
  cSQL = cSQL & " AND Tgl <='" & Format(dTgl(1).Value, "yyyy-MM-dd") & "'"
  If Check1.Value = 1 Then
    cSQL = cSQL & " AND username = '" & cKasir.Text & "'"
  End If
  Set dbData = objData.Sql(GetDSN, cSQL)
  If Not dbData.EOF Then
    nSubTotal = GetNull(dbData!Subtotal)
    nDisc = GetNull(dbData!Disc)
    nDisc1 = 0
    nPajak = 0
    nTotal = GetNull(dbData!Total)
  End If
End Sub

Private Sub optKodeStock_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    keybd_event VK_TAB, 0, 0, 0
  End If
End Sub
