VERSION 5.00
Object = "{0D6235E7-DBA2-11D1-B5DF-0060976089D0}#1.0#0"; "tdbr6.ocx"
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form rptLabaRugiUpdate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Laba Rugi"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   6495
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   2010
      Left            =   0
      Top             =   0
      Width           =   6480
      _ExtentX        =   11430
      _ExtentY        =   3545
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
         Caption         =   "Ya, Tampilkan neraca konsolidasi"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2085
         TabIndex        =   2
         Top             =   990
         Width           =   3315
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Current Periode"
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
         Index           =   1
         Left            =   2820
         TabIndex        =   1
         Top             =   1650
         Width           =   1770
      End
      Begin VB.OptionButton Option1 
         Caption         =   "All"
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
         Index           =   0
         Left            =   2100
         TabIndex        =   0
         Top             =   1650
         Width           =   555
      End
      Begin BiSANumberBoxProject.BiSANumberBox nLevel 
         Height          =   330
         Left            =   285
         TabIndex        =   3
         Top             =   600
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   582
         Appearance      =   0
         Decimals        =   0
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "LEVEL"
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
      Begin BiSADateProject.BiSADate dAwal 
         Height          =   330
         Left            =   285
         TabIndex        =   4
         Top             =   195
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
         Caption         =   "TANGGAL MUTASI"
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
      Begin BiSADateProject.BiSADate dAkhir 
         Height          =   330
         Left            =   3870
         TabIndex        =   5
         Top             =   195
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
      Begin BiSATextBoxProject.BiSABrowse cCostCenter 
         Height          =   330
         Left            =   270
         TabIndex        =   6
         Top             =   1275
         Width           =   3255
         _ExtentX        =   5741
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
         Caption         =   "Cost Centre"
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
      Begin TrueDBReports60Ctl.TDBReports RptNeraca 
         Height          =   570
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1005
         Caption         =   "LABA RUGI"
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
         ConnectionString=   "DSN=Syariah"
         ConnectStringType=   3
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   "Syariah"
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         CursorLocation  =   3
         ConnectionTimeout=   15
         CommandTimeout  =   30
         RecordSource    =   "Select 'Laporan Neraca' as Kode"
         CursorType      =   3
         CommandType     =   8
         MaxRecords      =   0
         LinkType        =   0
         Master          =   ""
         CallDataRead    =   0   'False
         ConvertNullToEmpty=   -1  'True
         DesignConnection=   -1  'True
         DesignTimeout   =   5
         UnitsOfMeasurement=   1
         Vedit_ShowGrid  =   -1  'True
         Vedit_SnapToGrid=   0   'False
         Vedit_GridUnitWidth=   160
         Vedit_GridUnitHeight=   160
         Vedit_ShowCellExpressions=   -1  'True
         Norm_rect_left  =   0
         Norm_rect_top   =   0
         Norm_rect_right =   0
         Norm_rect_bottom=   0
         Virgin          =   0   'False
         Parameters.Count=   2
         Parameters(0).Name=   "cNamaKoperasi"
         Parameters(1).Name=   "cKota"
         Fields.Count    =   5
         Fields(0).Name  =   "Kode"
         Fields(0).DisplayName=   "Kode"
         Fields(0).MaxLength=   20
         Fields(1).Name  =   "Keterangan"
         Fields(1).DisplayName=   "Keterangan"
         Fields(1).MaxLength=   50
         Fields(2).Name  =   "Awal"
         Fields(2).DisplayName=   "Awal"
         Fields(2).Type  =   5
         Fields(3).Name  =   "Mutasi"
         Fields(3).DisplayName=   "Mutasi"
         Fields(3).Type  =   5
         Fields(4).Name  =   "Akhir"
         Fields(4).DisplayName=   "Akhir"
         Sections.Count  =   6
         Sections(0).Name=   "ReportHeader"
         Sections(0).Condition=   "RecNo() = 0"
         Sections(0).StyleExp=   "tdb_RepHeader"
         Sections(0).Cells.Count=   3
         Sections(0).Cells(0).Name=   "ReportHeader"
         Sections(0).Cells(0).Exp=   """LABA RUGI"""
         Sections(0).Cells(1).Name=   "NamaKoperasi"
         Sections(0).Cells(1).Exp=   "cNamaKoperasi"
         Sections(0).Cells(1).NewLine=   -1  'True
         Sections(0).Cells(2).Name=   "Antara"
         Sections(0).Cells(2).Exp=   """~~Antara"""
         Sections(0).Cells(2).StyleExp=   "'tdb_RepPeriode'"
         Sections(0).Cells(2).NewLine=   -1  'True
         Sections(0).Cells(2).Width=   100
         Sections(0).Cells(2).CallExpression=   -1  'True
         Sections(1).Name=   "PageHeader"
         Sections(1).Type=   1
         Sections(1).StyleExp=   "tdb_PageHeader"
         Sections(1).Cells.Count=   1
         Sections(1).Cells(0).Name=   "PageNumber"
         Sections(1).Cells(0).Exp=   """Page "" + CStr(PageNo())"
         Sections(1).Cells(0).Placement=   2
         Sections(2).Name=   "DetailHeader"
         Sections(2).Type=   3
         Sections(2).StyleExp=   "tdb_TableHeader"
         Sections(2).Tabulator=   "Detail"
         Sections(2).Cells.Count=   5
         Sections(2).Cells(0).Name=   "Kode"
         Sections(2).Cells(0).Exp=   """Kode"""
         Sections(2).Cells(0).Width=   10
         Sections(2).Cells(1).Name=   "Keterangan"
         Sections(2).Cells(1).Exp=   """Keterangan"""
         Sections(2).Cells(2).Name=   "Awal"
         Sections(2).Cells(2).Exp=   """1"""
         Sections(2).Cells(2).Width=   15
         Sections(2).Cells(2).CallExpression=   -1  'True
         Sections(2).Cells(3).Name=   "Mutasi"
         Sections(2).Cells(3).Exp=   """Mutasi"""
         Sections(2).Cells(3).Width=   15
         Sections(2).Cells(4).Name=   "Akhir"
         Sections(2).Cells(4).Exp=   """2"""
         Sections(2).Cells(4).CallExpression=   -1  'True
         Sections(3).Name=   "Detail"
         Sections(3).Type=   4
         Sections(3).Condition=   "Right(Kode, 1) <> 'X'"
         Sections(3).StyleExp=   "'tdb_TableOddRow'"
         Sections(3).Cells.Count=   5
         Sections(3).Cells(0).Name=   "Kode"
         Sections(3).Cells(0).Exp=   "IIF(Kode=""4Y"","""",Kode)"
         Sections(3).Cells(0).Width=   10
         Sections(3).Cells(1).Name=   "Keterangan"
         Sections(3).Cells(1).Exp=   "Keterangan"
         Sections(3).Cells(2).Name=   "Awal"
         Sections(3).Cells(2).Exp=   "Awal"
         Sections(3).Cells(2).StyleExp=   "'tdb_Number'"
         Sections(3).Cells(2).Width=   20
         Sections(3).Cells(2).CallExpression=   -1  'True
         Sections(3).Cells(3).Name=   "Mutasi"
         Sections(3).Cells(3).Exp=   "Mutasi"
         Sections(3).Cells(3).StyleExp=   "'tdb_Number'"
         Sections(3).Cells(3).Width=   20
         Sections(3).Cells(3).CallExpression=   -1  'True
         Sections(3).Cells(4).Name=   "Akhir"
         Sections(3).Cells(4).Exp=   "Akhir"
         Sections(3).Cells(4).StyleExp=   "'tdb_Number'"
         Sections(3).Cells(4).Width=   20
         Sections(3).Cells(4).CallExpression=   -1  'True
         Sections(4).Name=   "TotalAktiva"
         Sections(4).Type=   4
         Sections(4).Condition=   "Right(Kode, 1) = 'X'"
         Sections(4).StyleExp=   "'tdb_Total'"
         Sections(4).Tabulator=   "Detail"
         Sections(4).Cells.Count=   5
         Sections(4).Cells(0).Name=   "Kode"
         Sections(4).Cells(0).Exp=   "Keterangan"
         Sections(4).Cells(0).CellSpan=   2
         Sections(4).Cells(1).Name=   "Keterangan"
         Sections(4).Cells(2).Name=   "Awal"
         Sections(4).Cells(2).Exp=   "Awal"
         Sections(4).Cells(2).StyleExp=   "'tdb_NumberTotal'"
         Sections(4).Cells(2).CallExpression=   -1  'True
         Sections(4).Cells(3).Name=   "Mutasi"
         Sections(4).Cells(3).Exp=   "Mutasi"
         Sections(4).Cells(3).StyleExp=   "'tdb_NumberTotal'"
         Sections(4).Cells(3).CallExpression=   -1  'True
         Sections(4).Cells(4).Name=   "Akhir"
         Sections(4).Cells(4).Exp=   "Akhir"
         Sections(4).Cells(4).StyleExp=   "'tdb_NumberTotal'"
         Sections(4).Cells(4).CallExpression=   -1  'True
         Sections(5).Name=   "Pengesahan"
         Sections(5).Type=   5
         Sections(5).Condition=   "IsLastRec()"
         Sections(5).StyleExp=   "'tdb_Pengesahan'"
         Sections(5).Cells.Count=   6
         Sections(5).Cells(0).Name=   "CELL_4"
         Sections(5).Cells(0).NewLine=   -1  'True
         Sections(5).Cells(1).Name=   "CELL_5"
         Sections(5).Cells(1).Exp=   "cKota"
         Sections(5).Cells(1).NewLine=   -1  'True
         Sections(5).Cells(1).PrivateStyle=   -1  'True
         Sections(5).Cells(1).Style.Name=   "<private>"
         Sections(5).Cells(1).Style.ParentName=   "tdb_Pengesahan"
         Sections(5).Cells(1).Style.Font_Name=   "Arial"
         Sections(5).Cells(1).Style.Font_Size=   6.75
         Sections(5).Cells(1).Style.Font_Bold=   0   'False
         Sections(5).Cells(1).Style.Font_Italic=   0   'False
         Sections(5).Cells(1).Style.Font_Underline=   0   'False
         Sections(5).Cells(1).Style.Font_Strikeout=   0   'False
         Sections(5).Cells(1).Style.Font_Charset=   0
         Sections(5).Cells(1).Style.TextAlign=   2
         Sections(5).Cells(1).Style.TextVAlign=   0
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
         Sections(5).Cells(1).Style.MarginTop=   6
         Sections(5).Cells(1).Style.MarginRight=   6
         Sections(5).Cells(1).Style.MarginBottom=   6
         Sections(5).Cells(1).Style.HasBorders=   -1  'True
         Sections(5).Cells(1).Style.BorderHT=   ""
         Sections(5).Cells(1).Style.BorderHI=   ""
         Sections(5).Cells(1).Style.BorderHB=   ""
         Sections(5).Cells(1).Style.BorderVL=   ""
         Sections(5).Cells(1).Style.BorderVI=   ""
         Sections(5).Cells(1).Style.BorderVR=   ""
         Sections(5).Cells(1).Style.NoClipping=   -1  'True
         Sections(5).Cells(1).Style.RTF=   0   'False
         Sections(5).Cells(1).Style.fprops=   1
         Sections(5).Cells(2).Name=   "CELL_6"
         Sections(5).Cells(2).NewLine=   -1  'True
         Sections(5).Cells(3).Name=   "CELL_0"
         Sections(5).Cells(3).Exp=   """Mengetahui"""
         Sections(5).Cells(3).NewLine=   -1  'True
         Sections(5).Cells(3).Width=   33
         Sections(5).Cells(4).Name=   "CELL_1"
         Sections(5).Cells(4).Exp=   """Pemeriksa"""
         Sections(5).Cells(4).Width=   33
         Sections(5).Cells(5).Name=   "CELL_2"
         Sections(5).Cells(5).Exp=   """Pembuat"""
         Sections(5).Cells(5).Width=   33
         Styles.Count    =   26
         Styles(0).Name  =   "tdb_Base"
         Styles(0).ParentName=   ""
         Styles(0).Font_Size=   6.75
         Styles(0).Font_Charset=   0
         Styles(0).NoClipping=   -1  'True
         Styles(1).Name  =   "tdb_PageHeader"
         Styles(1).ParentName=   "tdb_Base"
         Styles(1).Font_Size=   6.75
         Styles(1).Font_Charset=   0
         Styles(1).TextAlign=   2
         Styles(1).NoClipping=   -1  'True
         Styles(1).fprops=   1
         Styles(2).Name  =   "tdb_PageFooter"
         Styles(2).ParentName=   "tdb_PageHeader"
         Styles(2).Font_Size=   6.75
         Styles(2).Font_Charset=   0
         Styles(2).NoClipping=   -1  'True
         Styles(2).fprops=   0
         Styles(3).Name  =   "tdb_GroupHeaderBase"
         Styles(3).ParentName=   "tdb_Base"
         Styles(3).Font_Name=   "Arial"
         Styles(3).Font_Size=   6.75
         Styles(3).Font_Charset=   0
         Styles(3).NoClipping=   -1  'True
         Styles(3).fprops=   2097152
         Styles(4).Name  =   "tdb_GroupHeader1"
         Styles(4).ParentName=   "tdb_GroupHeaderBase"
         Styles(4).Font_Size=   6
         Styles(4).Font_Bold=   -1  'True
         Styles(4).Font_Charset=   0
         Styles(4).NoClipping=   -1  'True
         Styles(4).fprops=   20971520
         Styles(5).Name  =   "tdb_GroupFooterBase"
         Styles(5).ParentName=   "tdb_Base"
         Styles(5).Font_Name=   "Arial"
         Styles(5).Font_Size=   6.75
         Styles(5).Font_Charset=   0
         Styles(5).TextAlign=   2
         Styles(5).NoClipping=   -1  'True
         Styles(5).fprops=   2097153
         Styles(6).Name  =   "tdb_GroupFooter1"
         Styles(6).ParentName=   "tdb_GroupFooterBase"
         Styles(6).Font_Size=   6
         Styles(6).Font_Bold=   -1  'True
         Styles(6).Font_Charset=   0
         Styles(6).NoClipping=   -1  'True
         Styles(6).fprops=   20971520
         Styles(7).Name  =   "tdb_GroupHeader2"
         Styles(7).ParentName=   "tdb_GroupHeaderBase"
         Styles(7).Font_Size=   6
         Styles(7).Font_Charset=   0
         Styles(7).NoClipping=   -1  'True
         Styles(7).fprops=   4194304
         Styles(8).Name  =   "tdb_GroupFooter2"
         Styles(8).ParentName=   "tdb_GroupFooterBase"
         Styles(8).Font_Size=   6
         Styles(8).Font_Charset=   0
         Styles(8).NoClipping=   -1  'True
         Styles(8).fprops=   4194304
         Styles(9).Name  =   "tdb_GroupHeader3"
         Styles(9).ParentName=   "tdb_GroupHeaderBase"
         Styles(9).Font_Size=   6
         Styles(9).Font_Bold=   -1  'True
         Styles(9).Font_Charset=   0
         Styles(9).NoClipping=   -1  'True
         Styles(9).fprops=   20971520
         Styles(10).Name =   "tdb_GroupFooter3"
         Styles(10).ParentName=   "tdb_GroupFooterBase"
         Styles(10).Font_Size=   6
         Styles(10).Font_Bold=   -1  'True
         Styles(10).Font_Charset=   0
         Styles(10).NoClipping=   -1  'True
         Styles(10).fprops=   20971520
         Styles(11).Name =   "tdb_GroupHeader4"
         Styles(11).ParentName=   "tdb_GroupHeaderBase"
         Styles(11).Font_Size=   6
         Styles(11).Font_Charset=   0
         Styles(11).NoClipping=   -1  'True
         Styles(11).fprops=   4194304
         Styles(12).Name =   "tdb_GroupFooter4"
         Styles(12).ParentName=   "tdb_GroupFooterBase"
         Styles(12).Font_Size=   6
         Styles(12).Font_Charset=   0
         Styles(12).NoClipping=   -1  'True
         Styles(12).fprops=   4194304
         Styles(13).Name =   "tdb_TableBase"
         Styles(13).ParentName=   "tdb_Base"
         Styles(13).Font_Name=   "Arial"
         Styles(13).Font_Size=   6.75
         Styles(13).Font_Charset=   0
         Styles(13).BorderHT=   "tdb_ThinBlack"
         Styles(13).BorderHI=   "tdb_Invisible"
         Styles(13).BorderHB=   "tdb_ThinBlack"
         Styles(13).BorderVL=   "tdb_ThinBlack"
         Styles(13).BorderVI=   "tdb_ThinGray"
         Styles(13).BorderVR=   "tdb_ThinBlack"
         Styles(13).NoClipping=   -1  'True
         Styles(13).fprops=   4161536
         Styles(14).Name =   "tdb_TableOddRow"
         Styles(14).ParentName=   "tdb_TableBase"
         Styles(14).Font_Size=   6.75
         Styles(14).Font_Charset=   0
         Styles(14).BorderHI=   "Inner"
         Styles(14).BorderVI=   "tdb_ThinBlack"
         Styles(14).NoClipping=   -1  'True
         Styles(14).fprops=   4784128
         Styles(15).Name =   "tdb_TableEvenRow"
         Styles(15).ParentName=   "tdb_TableOddRow"
         Styles(15).Font_Size=   6.75
         Styles(15).Font_Charset=   0
         Styles(15).BackColor=   8454143
         Styles(15).NoFill=   0   'False
         Styles(15).NoClipping=   -1  'True
         Styles(15).fprops=   4194352
         Styles(16).Name =   "tdb_TableHeader"
         Styles(16).ParentName=   "tdb_TableBase"
         Styles(16).Font_Size=   6.75
         Styles(16).Font_Bold=   -1  'True
         Styles(16).Font_Charset=   0
         Styles(16).TextAlign=   1
         Styles(16).BackColor=   15132390
         Styles(16).NoFill=   0   'False
         Styles(16).BorderHT=   "tdb_ThinBlack"
         Styles(16).BorderHI=   "tdb_ThinBlack"
         Styles(16).BorderHB=   "tdb_ThinBlack"
         Styles(16).BorderVL=   "tdb_ThinBlack"
         Styles(16).BorderVI=   "tdb_ThinBlack"
         Styles(16).BorderVR=   "tdb_ThinBlack"
         Styles(16).NoClipping=   -1  'True
         Styles(16).fprops=   23035961
         Styles(17).Name =   "tdb_TableFiller"
         Styles(17).ParentName=   "tdb_TableOddRow"
         Styles(17).Font_Size=   6.75
         Styles(17).Font_Charset=   0
         Styles(17).MarginTop=   0
         Styles(17).MarginBottom=   0
         Styles(17).NoClipping=   -1  'True
         Styles(17).fprops=   20480
         Styles(18).Name =   "tdb_RepHeader"
         Styles(18).ParentName=   "tdb_Base"
         Styles(18).Font_Name=   "Arial"
         Styles(18).Font_Size=   8.25
         Styles(18).Font_Bold=   -1  'True
         Styles(18).Font_Charset=   0
         Styles(18).TextAlign=   1
         Styles(18).NoClipping=   -1  'True
         Styles(18).fprops=   1096810497
         Styles(19).Name =   "tdb_Total"
         Styles(19).ParentName=   "tdb_TableBase"
         Styles(19).Font_Size=   6.75
         Styles(19).Font_Bold=   -1  'True
         Styles(19).Font_Charset=   0
         Styles(19).BackColor=   15132390
         Styles(19).NoFill=   0   'False
         Styles(19).BorderHT=   "tdb_ThinBlack"
         Styles(19).BorderHI=   "tdb_ThinBlack"
         Styles(19).BorderHB=   "tdb_ThinBlack"
         Styles(19).BorderVL=   "tdb_ThinBlack"
         Styles(19).BorderVI=   "tdb_ThinBlack"
         Styles(19).BorderVR=   "tdb_ThinBlack"
         Styles(19).NoClipping=   -1  'True
         Styles(19).fprops=   23035952
         Styles(20).Name =   "tdb_Number"
         Styles(20).ParentName=   "tdb_TableOddRow"
         Styles(20).Font_Size=   6.75
         Styles(20).Font_Charset=   0
         Styles(20).TextAlign=   2
         Styles(20).NoClipping=   -1  'True
         Styles(20).fprops=   4194305
         Styles(21).Name =   "tdb_Number_Negative"
         Styles(21).ParentName=   "tdb_Number"
         Styles(21).Font_Size=   6.75
         Styles(21).Font_Charset=   0
         Styles(21).ForeColor=   255
         Styles(21).NoClipping=   -1  'True
         Styles(21).fprops=   40
         Styles(22).Name =   "tdb_NumberTotal"
         Styles(22).ParentName=   "tdb_Total"
         Styles(22).Font_Size=   6.75
         Styles(22).Font_Charset=   0
         Styles(22).TextAlign=   2
         Styles(22).NoClipping=   -1  'True
         Styles(22).fprops=   4194305
         Styles(23).Name =   "tdb_RepPeriode"
         Styles(23).ParentName=   "tdb_RepHeader"
         Styles(23).Font_Size=   8.25
         Styles(23).Font_Charset=   0
         Styles(23).NoClipping=   -1  'True
         Styles(23).fprops=   4194304
         Styles(24).Name =   "tdb_Pengesahan"
         Styles(24).ParentName=   "tdb_Base"
         Styles(24).Font_Name=   "Arial"
         Styles(24).Font_Size=   6.75
         Styles(24).Font_Charset=   0
         Styles(24).TextAlign=   1
         Styles(24).NoClipping=   -1  'True
         Styles(24).fprops=   2097153
         Styles(25).Name =   "tdb_bold"
         Styles(25).ParentName=   "tdb_TableOddRow"
         Styles(25).Font_Size=   6.75
         Styles(25).Font_Bold=   -1  'True
         Styles(25).Font_Charset=   0
         Styles(25).NoClipping=   -1  'True
         Styles(25).fprops=   16777216
         Lines.Count     =   4
         Lines(0).Name   =   "tdb_Invisible"
         Lines(1).Name   =   "tdb_ThinBlack"
         Lines(2).Name   =   "tdb_ThinGray"
         Lines(2).Color  =   8421504
         Lines(3).Name   =   "Inner"
         Lines(3).Color  =   4210752
         Profiles.Count  =   1
         Profiles(0).Name=   "PROFILE_0"
         Profiles(0).Active=   -1  'True
         Profiles(0).PreviewNoMinimize=   -1  'True
         Profiles(0).PreviewNoMaximize=   -1  'True
         Profiles(0).PreviewNoResize=   -1  'True
         Profiles(0).PreviewMaximized=   -1  'True
         Profiles(0).PreviewNoSaveLoad=   -1  'True
         Profiles(0).PrinterMarginTop=   720
         Profiles(0).PrinterMarginBottom=   720
         Profiles(0).PrinterMargins_set=   -1  'True
         Profiles(0).PrinterPaperUserSize_set=   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "( 1-4 )"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2685
         TabIndex        =   8
         Top             =   660
         Width           =   525
      End
      Begin VB.Label Label3 
         Caption         =   "Konsolidasi"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   300
         TabIndex        =   7
         Top             =   975
         Width           =   1350
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame2 
      Height          =   630
      Left            =   0
      Top             =   1995
      Width           =   6480
      _ExtentX        =   11430
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
         Left            =   5250
         TabIndex        =   9
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
         Picture         =   "rptLabaRugiUpdate.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   4080
         TabIndex        =   10
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
         Picture         =   "rptLabaRugiUpdate.frx":00A6
      End
   End
End
Attribute VB_Name = "rptLabaRugiUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaNeraca As New XArrayDB

Private Sub cCostCenter_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "costcenter", "kodecostcenter,keterangan")
  If Not dbData.EOF Then
    cCostCenter.Text = cCostCenter.Browse(dbData)
  End If
End Sub

Private Sub Check1_Click()
  If Check1.Value = 1 Then
    cCostCenter.Enabled = False
  Else
    cCostCenter.Enabled = True
  End If
End Sub

Private Sub Check1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    keybd_event VK_TAB, 0, 0, 0
  End If
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdPreview_Click()
Dim cSQL As String
Dim n As Long
Dim cRekening As String
Dim nAwalAktiva As Double
Dim nMutasiAktiva As Double
Dim nAkhirAktiva As Double
Dim nAwalPasiva As Double
Dim nMutasiPasiva As Double
Dim nAkhirPasiva As Double
Dim nLaba As Double
Dim nAwalLaba As Double
Dim cSQLCostCenter As String
  
  vaNeraca.Clear
  vaNeraca.ReDim 0, -1, 0, 5
  
  cSQLCostCenter = ""
  If Check1.Value <> 1 Then
    cSQLCostCenter = " AND kodecostcenter = '" & cCostCenter.Text & "' "
  End If
  
  cSQL = cSQL & " Select '  ' as Kode,'PENDAPATAN       ' as Keterangan,0,0,0 "
  cSQL = cSQL & " Union "
  cSQL = cSQL & " Select '4X' as Kode,'TOTAL PENDAPATAN ' as Keterangan,0,0,0 "
  cSQL = cSQL & " Union "
  cSQL = cSQL & " Select '4Y' as Kode,'BIAYA            ' as Keterangan,0,0,0 "
  cSQL = cSQL & " Union "
  cSQL = cSQL & " Select 'XX' as Kode,'TOTAL BIAYA      ' as Keterangan,0,0,0 "
  cSQL = cSQL & " Union "
  cSQL = cSQL & " Select 'YX' as Kode,'LABA/RUGI        ' as Keterangan,0,0,0 "

  Set dbData = objData.SQL(GetDSN, cSQL)
  vaNeraca.LoadRows dbData.GetRows(dbData.RecordCount)
  
  ' Ambil Data Rekening
  Set dbData = objData.Browse(GetDSN, "akun", "kodeakun,keterangan", "LEFT(kodeakun,1)", sisGTEqual, "4", "and LEFT(kodeakun,1) <='5'", "kodeakun")
  If Not dbData.EOF Then
    FrmPB.InitPB dbData.RecordCount + 1
    dbData.MoveFirst
    Do While Not dbData.EOF
      FrmPB.RunPB
      If Level(dbData!kodeakun) <= nLevel.Value Then
        vaNeraca.InsertRows vaNeraca.UpperBound(1) + 1
        n = vaNeraca.UpperBound(1)
        vaNeraca(n, 0) = (dbData!kodeakun)
        vaNeraca(n, 1) = (RekSpace(dbData!kodeakun, dbData!keterangan))
        vaNeraca(n, 2) = 0
        vaNeraca(n, 3) = 0
        vaNeraca(n, 4) = 0
      End If
      dbData.MoveNext
    Loop
  End If
  FrmPB.EndPB
  
'   Ambil Data Saldo Awal
'  Set dbData = objData.Browse(GetDSN, "SaldoRekening")
'  If Not dbData.EOF Then
'    dbData.MoveFirst
'    Do While Not dbData.EOF
'      If Left(dbData!Rekening, 1) >= 4 And Left(dbData!Rekening, 1) <= 5 Then
'        cRekening = GetLevel(dbData!Rekening, nLevel.Value)
'        n = vaNeraca.Find(0, 0, cRekening)
'        If n >= 0 Then
'          vaNeraca(n, 2) = vaNeraca(n, 2) + GS(dbData!Rekening, dbData!Awal)
'          vaNeraca(n, 4) = vaNeraca(n, 2) + vaNeraca(n, 3)
'        End If
'      End If
'      dbData.MoveNext
'    Loop
'  End If
  If Option1(0).Value = True Then
    Set dbData = objData.Browse(GetDSN, "bukubesar", "kodeakun,tgl,SUM(debet) as Debet,SUM(kredit) as Kredit", , , , "1=1 AND tgl < '" & Format(dAwal.Value, "yyyy-MM-dd") & "' " & cSQLCostCenter & " Group by kodeakun")
  Else
    Set dbData = objData.Browse(GetDSN, "bukubesar", "kodeakun,tgl,SUM(debet) as Debet,SUM(kredit) as Kredit", , , , "1=1 AND tgl < '" & Format(dAwal.Value, "yyyy-MM-dd") & "' and tgl > '" & Year(Date) - 1 & "-12-31" & "'" & cSQLCostCenter & " Group by kodeakun")
  End If
  If Not dbData.EOF Then
    dbData.MoveFirst
    FrmPB.InitPB dbData.RecordCount
    Do While Not dbData.EOF
      FrmPB.RunPB
      cRekening = GetLevel(dbData!kodeakun, nLevel.Value)
      If Left(dbData!kodeakun, 1) = 4 Or Left(dbData!kodeakun, 1) = 5 Then
        n = vaNeraca.Find(0, 0, cRekening)
        If n >= 0 Then
          vaNeraca(n, 2) = vaNeraca(n, 2) + GS(dbData!kodeakun, dbData!debet) - GS(dbData!kodeakun, dbData!kredit)
          vaNeraca(n, 4) = vaNeraca(n, 2) + vaNeraca(n, 3)
        End If
      End If
      dbData.MoveNext
    Loop
    FrmPB.EndPB
  End If
  
  ' Ambil Data Pada Buku Besar Untuk Mutasi
  Set dbData = objData.Browse(GetDSN, "bukubesar", "kodeakun,tgl,SUM(debet) as Debet,SUM(kredit) as Kredit", , , , "1=1 AND tgl >= '" & Format(dAwal.Value, "yyyy-MM-dd") & "' AND tgl <= '" & Format(dAkhir.Value, "yyyy-MM-dd") & "' " & cSQLCostCenter & " Group by kodeakun")
  If Not dbData.EOF Then
    FrmPB.InitPB dbData.RecordCount
    dbData.MoveFirst
    Do While Not dbData.EOF
      FrmPB.RunPB
      cRekening = GetLevel(dbData!kodeakun, nLevel.Value)
      If Left(dbData!kodeakun, 1) = 5 Or Left(dbData!kodeakun, 1) = 4 Then
        n = vaNeraca.Find(0, 0, cRekening)
        If n >= 0 Then
          vaNeraca(n, 3) = vaNeraca(n, 3) + GS(dbData!kodeakun, dbData!debet) - GS(dbData!kodeakun, dbData!kredit)
          vaNeraca(n, 4) = vaNeraca(n, 2) + vaNeraca(n, 3)
        End If
      End If
      dbData.MoveNext
    Loop
    FrmPB.EndPB
  End If
  
  'Hitung Total Aktiva dan Total Pasiva
  nAwalAktiva = 0
  nMutasiAktiva = 0
  nAkhirAktiva = 0
  nAwalPasiva = 0
  nMutasiPasiva = 0
  nAkhirPasiva = 0
  For n = 0 To vaNeraca.UpperBound(1)
    If Left(vaNeraca(n, 0), 1) = "4" Then
      nAwalAktiva = nAwalAktiva + GetNull(vaNeraca(n, 2))
      nMutasiAktiva = nMutasiAktiva + GetNull(vaNeraca(n, 3))
      nAkhirAktiva = nAkhirAktiva + GetNull(vaNeraca(n, 4))
    Else
      nAwalPasiva = nAwalPasiva + GetNull(vaNeraca(n, 2))
      nMutasiPasiva = nMutasiPasiva + GetNull(vaNeraca(n, 3))
      nAkhirPasiva = nAkhirPasiva + GetNull(vaNeraca(n, 4))
    End If
  Next
  
  ' Masukkan Total Aktiva dan Total Pasiva
  vaNeraca(1, 2) = nAwalAktiva
  vaNeraca(1, 3) = nMutasiAktiva
  vaNeraca(1, 4) = nAkhirAktiva
  
  vaNeraca(3, 2) = nAwalPasiva
  vaNeraca(3, 3) = nMutasiPasiva
  vaNeraca(3, 4) = nAkhirPasiva
  
  vaNeraca(4, 2) = vaNeraca(1, 2) + vaNeraca(3, 2)
  vaNeraca(4, 3) = vaNeraca(1, 3) + vaNeraca(3, 3)
  vaNeraca(4, 4) = vaNeraca(1, 4) + vaNeraca(3, 4)

  For n = 0 To vaNeraca.UpperBound(1)
    If Left(vaNeraca(n, 0), 1) = 5 Or n = 3 Then
      vaNeraca(n, 2) = IIf(vaNeraca(n, 2) < 0, -vaNeraca(n, 2), vaNeraca(n, 2))
      vaNeraca(n, 3) = IIf(vaNeraca(n, 3) < 0, -vaNeraca(n, 3), vaNeraca(n, 3))
      vaNeraca(n, 4) = IIf(vaNeraca(n, 4) < 0, -vaNeraca(n, 4), vaNeraca(n, 4))
    End If
  Next
  
  vaNeraca.QuickSort vaNeraca.LowerBound(1), vaNeraca.UpperBound(1), 0, XORDER_ASCEND, XTYPE_DEFAULT
  
  Dim dba As New ADODB.Recordset
  Set dba = objData.Browse(GetDSN, "akun", , "jenis", sisAssign, "D")
  If Not dba.EOF Then
    Do While Not dba.EOF
      cRekening = GetLevel(dba!kodeakun, nLevel.Value)
      n = vaNeraca.Find(0, 0, cRekening)
      If n > 0 Then
        If vaNeraca(n, 4) = 0 And vaNeraca(n, 3) = 0 And vaNeraca(n, 2) = 0 Then
          vaNeraca.DeleteRows n
        End If
      End If
      dba.MoveNext
    Loop
  End If
  
  Set RptNeraca.Array = vaNeraca
  RptNeraca.Refresh
  With RptNeraca
    .Profiles(0).PrinterMarginBottom = 720
    .Profiles(0).PrinterMarginLeft = 720
    .Profiles(0).PrinterMarginRight = 720
    .Profiles(0).PrinterMarginTop = 720
    .Parameters(0).Value = aCfg(objData, msNamaPerusahaan)
'    .PageSetup
    .PrintPreview
  End With
End Sub

Private Function GS(ByVal cRekening As String, ByVal nValue As Double) As Double
  GS = IIf(Left(cRekening, 1) = "1", nValue, -nValue)
End Function

Private Sub dAkhir_Validate(Cancel As Boolean)
  If dAkhir.Value < dAwal.Value Then
    dAwal.Value = dAkhir.Value
  End If
End Sub

Private Sub dAwal_Validate(Cancel As Boolean)
  If dAkhir.Value < dAwal.Value Then
    dAwal.Value = dAkhir.Value
  End If
End Sub

Private Sub Form_Load()
Dim n As Single

  CenterForm Me
  SetIcon Me.hWnd
  dAwal.Value = BOM(Date)
  dAkhir.Value = Date
  nLevel.Value = 4
  Check1.Value = 1
  cCostCenter.Default
  cCostCenter.Enabled = False
  Option1(1).Value = True
  
  TabIndex dAwal, n
  TabIndex dAkhir, n
  TabIndex nLevel, n
  TabIndex Check1, n
  TabIndex cCostCenter, n
  TabIndex Option1(0), n
  TabIndex Option1(1), n
  
  TabIndex cmdPreview, n
  TabIndex cmdKeluar, n
End Sub

Private Sub RptNeraca_CellExpression(ByVal Section As Integer, ByVal Cell As Integer, Value As Variant)
Dim cChar As String
  If Section = 0 Then
    Value = "Per " & Day(dAkhir.Value) & " " & GetMonth(Month(dAkhir.Value)) & " " & Year(dAkhir.Value) '   Format(dAkhir.Value, "dd mmmm yyyy")
  ElseIf Section = 2 Then
    If Value = "1" Then
      Value = Format(dAwal.Value - 1, "dd-mm-yyyy")
    Else
      Value = Format(dAkhir.Value, "dd-mm-yyyy")
    End If
  Else
    Value = GetNull(Value, 0)
    cChar = IIf(Value < 0, "()", "  ")
    If Value = 0 Then
      Value = ""
    Else
      Value = Format(Abs(GetNull(Value)), "###,###,###,###,##0.00")
    End If
    Value = Left(cChar, 1) & Value & Right(cChar, 1)
  End If
End Sub

