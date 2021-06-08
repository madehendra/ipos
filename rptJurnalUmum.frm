VERSION 5.00
Object = "{0D6235E7-DBA2-11D1-B5DF-0060976089D0}#1.0#0"; "tdbr6.ocx"
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form rptJurnalUmum 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Laporan Jurnal Umum"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6975
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   6975
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1605
      Left            =   15
      Top             =   0
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   2831
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
         Left            =   2355
         TabIndex        =   5
         Top             =   795
         Width           =   3315
      End
      Begin BiSADateProject.BiSADate dDate 
         Height          =   330
         Index           =   0
         Left            =   225
         TabIndex        =   0
         Top             =   315
         Width           =   3465
         _ExtentX        =   6112
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
      Begin BiSADateProject.BiSADate dDate 
         Height          =   330
         Index           =   1
         Left            =   3855
         TabIndex        =   1
         Top             =   315
         Width           =   1965
         _ExtentX        =   3466
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
      Begin TrueDBReports60Ctl.TDBReports tdb 
         Height          =   570
         Left            =   4770
         TabIndex        =   2
         Top             =   210
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   1005
         Caption         =   "Jurnal Umum"
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
         Parameters.Count=   4
         Parameters(0).Name=   "cNamaPerusahaan"
         Parameters(1).Name=   "cCostCenter"
         Parameters(2).Name=   "TGL1"
         Parameters(3).Name=   "TGL2"
         Fields.Count    =   8
         Fields(0).Name  =   "nojurnal"
         Fields(0).DisplayName=   "faktur"
         Fields(1).Name  =   "tgl"
         Fields(1).DisplayName=   "tgl"
         Fields(2).Name  =   "kodecostcenter"
         Fields(2).DisplayName=   "kodecostcenter"
         Fields(3).Name  =   "kodeakun"
         Fields(3).DisplayName=   "kodeakun"
         Fields(4).Name  =   "namaakun"
         Fields(4).DisplayName=   "namaakun"
         Fields(5).Name  =   "keterangan"
         Fields(5).DisplayName=   "keterangan"
         Fields(6).Name  =   "debet"
         Fields(6).DisplayName=   "debet"
         Fields(7).Name  =   "kredit"
         Fields(7).DisplayName=   "kredit"
         Sections.Count  =   4
         Sections(0).Name=   "SECTION_1"
         Sections(0).Type=   1
         Sections(0).Cells.Count=   7
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
         Sections(0).Cells(2).Exp=   """JURNAL UMUM "" & cCostCenter"
         Sections(0).Cells(2).NewLine=   -1  'True
         Sections(0).Cells(2).PrivateStyle=   -1  'True
         Sections(0).Cells(2).Style.Name=   "<private>"
         Sections(0).Cells(2).Style.ParentName=   "<null>"
         Sections(0).Cells(2).Style.Font_Name=   "Verdana"
         Sections(0).Cells(2).Style.Font_Size=   11.25
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
         Sections(0).Cells(3).Name=   "CELL_7"
         Sections(0).Cells(3).Exp=   "cNamaPerusahaan"
         Sections(0).Cells(3).NewLine=   -1  'True
         Sections(0).Cells(3).PrivateStyle=   -1  'True
         Sections(0).Cells(3).Style.Name=   "<private>"
         Sections(0).Cells(3).Style.ParentName=   "<null>"
         Sections(0).Cells(3).Style.Font_Name=   "Verdana"
         Sections(0).Cells(3).Style.Font_Size=   14.25
         Sections(0).Cells(3).Style.Font_Bold=   0   'False
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
         Sections(0).Cells(3).Style.fprops=   6291457
         Sections(0).Cells(4).Name=   "CELL_3"
         Sections(0).Cells(4).Exp=   """Antara Tanggal : "" & TGL1 & "" S.D "" & TGL2"
         Sections(0).Cells(4).NewLine=   -1  'True
         Sections(0).Cells(4).PrivateStyle=   -1  'True
         Sections(0).Cells(4).Format=   "dd-Mm-yyyy"
         Sections(0).Cells(4).Style.Name=   "<private>"
         Sections(0).Cells(4).Style.ParentName=   "<null>"
         Sections(0).Cells(4).Style.Font_Name=   "Verdana"
         Sections(0).Cells(4).Style.Font_Size=   9.75
         Sections(0).Cells(4).Style.Font_Bold=   -1  'True
         Sections(0).Cells(4).Style.Font_Italic=   0   'False
         Sections(0).Cells(4).Style.Font_Underline=   0   'False
         Sections(0).Cells(4).Style.Font_Strikeout=   0   'False
         Sections(0).Cells(4).Style.Font_Charset=   1
         Sections(0).Cells(4).Style.TextAlign=   1
         Sections(0).Cells(4).Style.TextVAlign=   0
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
         Sections(0).Cells(4).Style.MarginTop=   6
         Sections(0).Cells(4).Style.MarginRight=   6
         Sections(0).Cells(4).Style.MarginBottom=   6
         Sections(0).Cells(4).Style.HasBorders=   -1  'True
         Sections(0).Cells(4).Style.BorderHT=   ""
         Sections(0).Cells(4).Style.BorderHI=   ""
         Sections(0).Cells(4).Style.BorderHB=   ""
         Sections(0).Cells(4).Style.BorderVL=   ""
         Sections(0).Cells(4).Style.BorderVI=   ""
         Sections(0).Cells(4).Style.BorderVR=   ""
         Sections(0).Cells(4).Style.NoClipping=   0   'False
         Sections(0).Cells(4).Style.RTF=   0   'False
         Sections(0).Cells(4).Style.fprops=   23068673
         Sections(0).Cells(5).Name=   "CELL_4"
         Sections(0).Cells(5).Exp=   """ """
         Sections(0).Cells(5).NewLine=   -1  'True
         Sections(0).Cells(6).Name=   "CELL_5"
         Sections(0).Cells(6).Exp=   """ """
         Sections(0).Cells(6).NewLine=   -1  'True
         Sections(1).Name=   "SECTION_5"
         Sections(1).Condition=   "HasChanged(nojurnal)"
         Sections(1).StyleExp=   "'tdb_Base'"
         Sections(1).KeepWithNext=   2
         Sections(1).Cells.Count=   7
         Sections(1).Cells(0).Name=   "CELL_8"
         Sections(1).Cells(0).NewLine=   -1  'True
         Sections(1).Cells(1).Name=   "CELL_0"
         Sections(1).Cells(1).Exp=   """NO."""
         Sections(1).Cells(1).NewLine=   -1  'True
         Sections(1).Cells(1).Width=   20
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
         Sections(1).Cells(2).Name=   "CELL_1"
         Sections(1).Cells(2).Exp=   """: "" & nojurnal"
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
         Sections(1).Cells(3).Name=   "CELL_3"
         Sections(1).Cells(3).Exp=   """COST CENTRE"""
         Sections(1).Cells(3).NewLine=   -1  'True
         Sections(1).Cells(3).Width=   20
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
         Sections(1).Cells(4).Name=   "CELL_4"
         Sections(1).Cells(4).Exp=   """: "" & kodecostcenter"
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
         Sections(1).Cells(5).Name=   "CELL_6"
         Sections(1).Cells(5).Exp=   """TANGGAL"""
         Sections(1).Cells(5).NewLine=   -1  'True
         Sections(1).Cells(5).Width=   20
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
         Sections(1).Cells(6).Name=   "CELL_7"
         Sections(1).Cells(6).Exp=   """: "" & tgl"
         Sections(1).Cells(6).Height=   4
         Sections(1).Cells(6).AutoHeight=   0   'False
         Sections(1).Cells(6).PrivateStyle=   -1  'True
         Sections(1).Cells(6).Style.Name=   "<private>"
         Sections(1).Cells(6).Style.ParentName=   "tdb_Base"
         Sections(1).Cells(6).Style.Font_Name=   "Arial"
         Sections(1).Cells(6).Style.Font_Size=   8.25
         Sections(1).Cells(6).Style.Font_Bold=   -1  'True
         Sections(1).Cells(6).Style.Font_Italic=   0   'False
         Sections(1).Cells(6).Style.Font_Underline=   0   'False
         Sections(1).Cells(6).Style.Font_Strikeout=   0   'False
         Sections(1).Cells(6).Style.Font_Charset=   0
         Sections(1).Cells(6).Style.TextAlign=   0
         Sections(1).Cells(6).Style.TextVAlign=   1
         Sections(1).Cells(6).Style.TextWrap=   -1  'True
         Sections(1).Cells(6).Style.ForeColor=   0
         Sections(1).Cells(6).Style.BackColor=   16777215
         Sections(1).Cells(6).Style.NoFill=   -1  'True
         Sections(1).Cells(6).Style.BackPicFile=   ""
         Sections(1).Cells(6).Style.ForePicFile=   ""
         Sections(1).Cells(6).Style.BackPicVertPlacement=   0
         Sections(1).Cells(6).Style.BackPicHorzPlacement=   0
         Sections(1).Cells(6).Style.ForePicPlacement=   0
         Sections(1).Cells(6).Style.ForePicDrawMode=   0
         Sections(1).Cells(6).Style.MarginLeft=   6
         Sections(1).Cells(6).Style.MarginTop=   6
         Sections(1).Cells(6).Style.MarginRight=   6
         Sections(1).Cells(6).Style.MarginBottom=   6
         Sections(1).Cells(6).Style.HasBorders=   -1  'True
         Sections(1).Cells(6).Style.BorderHT=   ""
         Sections(1).Cells(6).Style.BorderHI=   ""
         Sections(1).Cells(6).Style.BorderHB=   ""
         Sections(1).Cells(6).Style.BorderVL=   ""
         Sections(1).Cells(6).Style.BorderVI=   ""
         Sections(1).Cells(6).Style.BorderVR=   ""
         Sections(1).Cells(6).Style.NoClipping=   -1  'True
         Sections(1).Cells(6).Style.RTF=   0   'False
         Sections(1).Cells(6).Style.fprops=   16777216
         Sections(2).Name=   "DetailHeader"
         Sections(2).Type=   3
         Sections(2).StyleExp=   "tdb_TableHeader"
         Sections(2).Tabulator=   "Detail"
         Sections(2).Cells.Count=   6
         Sections(2).Cells(0).Name=   "Nomor"
         Sections(2).Cells(0).Exp=   """No."""
         Sections(2).Cells(1).Name=   "Kode"
         Sections(2).Cells(1).Exp=   """AKUN"""
         Sections(2).Cells(2).Name=   "JudulBuku"
         Sections(2).Cells(2).Exp=   """NAMA"""
         Sections(2).Cells(3).Name=   "Jumlah"
         Sections(2).Cells(3).Exp=   """KETT"""
         Sections(2).Cells(4).Name=   "Harga"
         Sections(2).Cells(4).Exp=   """DEBET"""
         Sections(2).Cells(5).Name=   "Discount"
         Sections(2).Cells(5).Exp=   """KREDIT"""
         Sections(3).Name=   "Detail"
         Sections(3).Type=   4
         Sections(3).StyleExp=   "'tdb_TableOddRow'"
         Sections(3).Cells.Count=   6
         Sections(3).Cells(0).Name=   "No"
         Sections(3).Cells(0).Exp=   "Sum(1,WillChange(nojurnal))"
         Sections(3).Cells(0).Width=   4
         Sections(3).Cells(1).Name=   "KodeAkun"
         Sections(3).Cells(1).Exp=   "kodeakun"
         Sections(3).Cells(1).Width=   10
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
         Sections(3).Cells(1).Style.TextAlign=   0
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
         Sections(3).Cells(2).Name=   "NamaAkun"
         Sections(3).Cells(2).Exp=   "namaakun"
         Sections(3).Cells(2).Width=   30
         Sections(3).Cells(2).PrivateStyle=   -1  'True
         Sections(3).Cells(2).Style.Name=   "<private>"
         Sections(3).Cells(2).Style.ParentName=   "tdb_TableOddRow"
         Sections(3).Cells(2).Style.Font_Name=   "Arial"
         Sections(3).Cells(2).Style.Font_Size=   8.25
         Sections(3).Cells(2).Style.Font_Bold=   0   'False
         Sections(3).Cells(2).Style.Font_Italic=   0   'False
         Sections(3).Cells(2).Style.Font_Underline=   0   'False
         Sections(3).Cells(2).Style.Font_Strikeout=   0   'False
         Sections(3).Cells(2).Style.Font_Charset=   0
         Sections(3).Cells(2).Style.TextAlign=   3
         Sections(3).Cells(2).Style.TextVAlign=   0
         Sections(3).Cells(2).Style.TextWrap=   0   'False
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
         Sections(3).Cells(2).Style.BorderHT=   "Quarter"
         Sections(3).Cells(2).Style.BorderHI=   "Quarter"
         Sections(3).Cells(2).Style.BorderHB=   "Double"
         Sections(3).Cells(2).Style.BorderVL=   "Single"
         Sections(3).Cells(2).Style.BorderVI=   "Single"
         Sections(3).Cells(2).Style.BorderVR=   "Single"
         Sections(3).Cells(2).Style.NoClipping=   -1  'True
         Sections(3).Cells(2).Style.RTF=   0   'False
         Sections(3).Cells(2).Style.fprops=   4
         Sections(3).Cells(3).Name=   "Keterangan"
         Sections(3).Cells(3).Exp=   "keterangan"
         Sections(3).Cells(4).Name=   "Debet"
         Sections(3).Cells(4).Exp=   "debet"
         Sections(3).Cells(4).Width=   13
         Sections(3).Cells(4).Format=   "###,###,##0.00"
         Sections(3).Cells(5).Name=   "Kredit"
         Sections(3).Cells(5).Exp=   "kredit"
         Sections(3).Cells(5).Width=   13
         Sections(3).Cells(5).Format=   "###,###,##0.00"
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
      Begin BiSATextBoxProject.BiSABrowse cCostCenter 
         Height          =   330
         Left            =   225
         TabIndex        =   6
         Top             =   1080
         Width           =   3555
         _ExtentX        =   6271
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
         Left            =   285
         TabIndex        =   7
         Top             =   780
         Width           =   1350
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   630
      Left            =   15
      Top             =   1590
      Width           =   6975
      _ExtentX        =   12303
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
         Left            =   5805
         TabIndex        =   3
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
         Picture         =   "rptJurnalUmum.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   5355
         TabIndex        =   4
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
         Picture         =   "rptJurnalUmum.frx":00A6
      End
   End
End
Attribute VB_Name = "rptJurnalUmum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB

Private Sub cCostCenter_ButtonClick()
  Set dbData = objData.Browse(GetDSN, "costcenter", "kodecostcenter,keterangan", "kodecostcenter", sisContent, cCostCenter.Text)
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
    SendKeysA VK_TAB, True
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
  dDate(0).Value = BOM(Date)
  cCostCenter.Default
  TabIndex dDate(0), n
  TabIndex dDate(1), n
  TabIndex Check1, n
  TabIndex cCostCenter, n
  TabIndex cmdPreview, n
  TabIndex cmdKeluar, n
End Sub

Private Sub GetData()
Dim n As Integer
Dim cWhere As String

  cWhere = ""
  vaArray.ReDim 0, -1, 0, 7
  cWhere = " AND t.kodecostcenter = '" & cCostCenter.Text & "'"
  Set dbData = objData.Browse(GetDSN, "jurnalumum u", "u.nomorjurnalumum,t.kodecostcenter,u.tgl,u.kodeakun,a.keterangan as namaakun,u.keterangan,u.debet,u.kredit", , , , " 1=1 AND u.tgl >= '" & Format(dDate(0).Value, "yyyy-MM-dd") & "' AND u.tgl <= '" & Format(dDate(1).Value, "yyyy-MM-dd") & "'" & IIf(Check1.Value <> 1, cWhere, ""), "u.nomorjurnalumum", Array("LEFT JOIN totjurnalumum t on t.nomorjurnalumum = u.nomorjurnalumum", "LEFT JOIN akun a on a.kodeakun = u.kodeakun"))
  If Not dbData.EOF Then
    FrmPB.InitPB dbData.RecordCount
    Do While Not dbData.EOF
      FrmPB.RunPB
      'no jurnal
      'tgl
      'kode costcenter
      'kodeakun
      'namaakun
      'keterangan
      'debet
      'kredit
      
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = GetNull(dbData!nomorjurnalumum)
      vaArray(n, 1) = Format(GetNull(dbData!tgl), "dd-MM-yyyy")
      vaArray(n, 2) = GetNull(dbData!kodecostcenter)
      vaArray(n, 3) = GetNull(dbData!kodeakun)
      vaArray(n, 4) = GetNull(dbData!namaakun)
      vaArray(n, 5) = GetNull(dbData!keterangan)
      vaArray(n, 6) = GetNull(dbData!debet)
      vaArray(n, 7) = GetNull(dbData!kredit)
      dbData.MoveNext
    Loop
    FrmPB.EndPB
    tdb.Parameters("cNamaPerusahaan").ValueExpression = "'" & aCfg(objData, msNamaPerusahaan) & "'"
    If Check1.Value <> 1 Then
      tdb.Parameters("cCostCenter").ValueExpression = "'" & UCase(cCostCenter.Text) & "'"
    Else
      tdb.Parameters("cCostCenter").ValueExpression = "'KONSOLIDASI'"
    End If
    tdb.Parameters("TGL1").ValueExpression = "'" & Format(dDate(0).Value, "dd-MM-yyyy") & "'"
    tdb.Parameters("TGL2").ValueExpression = "'" & Format(dDate(1).Value, "dd-MM-yyyy") & "'"
    Set tdb.Array = vaArray
    tdb.Refresh
    tdb.PrintPreview
  End If
End Sub
