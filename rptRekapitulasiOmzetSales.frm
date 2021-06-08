VERSION 5.00
Object = "{0D6235E7-DBA2-11D1-B5DF-0060976089D0}#1.0#0"; "tdbr6.ocx"
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{FE28459D-12F1-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Date.ocx"
Begin VB.Form rptRekapitulasiOmzetSales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rekapitulasi Omzet Salesman"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   7020
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   1080
      Left            =   15
      Top             =   0
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   1905
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
      Begin BiSATextBoxProject.BiSABrowse cKode 
         Height          =   330
         Index           =   0
         Left            =   225
         TabIndex        =   0
         Top             =   180
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   582
         Text            =   "123"
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
         MaxLength       =   3
         Button          =   -1  'True
         Caption         =   "KODE SALESMAN"
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
      Begin BiSATextBoxProject.BiSABrowse cKode 
         Height          =   330
         Index           =   1
         Left            =   3855
         TabIndex        =   1
         Top             =   180
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   582
         Text            =   "123"
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
         MaxLength       =   3
         Button          =   -1  'True
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
      Begin BiSADateProject.BiSADate dAwal 
         Height          =   330
         Left            =   225
         TabIndex        =   2
         Top             =   585
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
      Begin BiSADateProject.BiSADate dAkhir 
         Height          =   330
         Left            =   3855
         TabIndex        =   3
         Top             =   600
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
      Begin TrueDBReports60Ctl.TDBReports RptOmzetSales 
         Height          =   570
         Left            =   0
         TabIndex        =   4
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
         Parameters.Count=   6
         Parameters(0).Name=   "cKode1"
         Parameters(0).ValueExpression=   """"""
         Parameters(1).Name=   "cKode2"
         Parameters(1).ValueExpression=   """"""
         Parameters(2).Name=   "dTgl1"
         Parameters(2).ValueExpression=   """"""
         Parameters(3).Name=   "dTgl2"
         Parameters(3).ValueExpression=   """"""
         Parameters(4).Name=   "cFkt1"
         Parameters(4).ValueExpression=   """"""
         Parameters(5).Name=   "cFkt2"
         Parameters(5).ValueExpression=   """"""
         Fields.Count    =   14
         Fields(0).Name  =   "Tgl"
         Fields(0).DisplayName=   "Tgl"
         Fields(0).Type  =   7
         Fields(1).Name  =   "KodeSales"
         Fields(1).DisplayName=   "KodeSales"
         Fields(2).Name  =   "Nama"
         Fields(2).DisplayName=   "Nama"
         Fields(3).Name  =   "Alamat"
         Fields(3).DisplayName=   "Alamat"
         Fields(4).Name  =   "Kota"
         Fields(4).DisplayName=   "Kota"
         Fields(5).Name  =   "SubTotal1"
         Fields(5).DisplayName=   "SubTotal"
         Fields(5).Type  =   5
         Fields(6).Name  =   "Disc1"
         Fields(6).DisplayName=   "Disc"
         Fields(6).Type  =   5
         Fields(7).Name  =   "Total1"
         Fields(7).DisplayName=   "Total"
         Fields(7).Type  =   5
         Fields(8).Name  =   "SubTotal2"
         Fields(8).DisplayName=   "SubTotal2"
         Fields(8).Type  =   5
         Fields(9).Name  =   "Disc2"
         Fields(9).DisplayName=   "Disc2"
         Fields(9).Type  =   5
         Fields(10).Name =   "Total2"
         Fields(10).DisplayName=   "Total2"
         Fields(10).Type =   5
         Fields(11).Name =   "Pembelian"
         Fields(11).DisplayName=   "Pembelian"
         Fields(12).Name =   "Komisi"
         Fields(12).DisplayName=   "Komisi"
         Fields(13).Name =   "Bersih"
         Fields(13).DisplayName=   "Bersih"
         Fields(13).Type =   5
         Sections.Count  =   7
         Sections(0).Name=   "SECTION_1"
         Sections(0).Type=   1
         Sections(0).Cells.Count=   1
         Sections(0).Cells(0).Name=   "CELL_0"
         Sections(0).Cells(0).Exp=   """Page : "" & PageNo()"
         Sections(0).Cells(0).PrivateStyle=   -1  'True
         Sections(0).Cells(0).Style.Name=   "<private>"
         Sections(0).Cells(0).Style.ParentName=   "tdb_Base"
         Sections(0).Cells(0).Style.Font_Name=   "Arial"
         Sections(0).Cells(0).Style.Font_Size=   8.25
         Sections(0).Cells(0).Style.Font_Bold=   0   'False
         Sections(0).Cells(0).Style.Font_Italic=   0   'False
         Sections(0).Cells(0).Style.Font_Underline=   0   'False
         Sections(0).Cells(0).Style.Font_Strikeout=   0   'False
         Sections(0).Cells(0).Style.Font_Charset=   0
         Sections(0).Cells(0).Style.TextAlign=   2
         Sections(0).Cells(0).Style.TextVAlign=   1
         Sections(0).Cells(0).Style.TextWrap=   0   'False
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
         Sections(0).Cells(0).Style.MarginTop=   4
         Sections(0).Cells(0).Style.MarginRight=   6
         Sections(0).Cells(0).Style.MarginBottom=   4
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
         Sections(1).Name=   "SECTION_5"
         Sections(1).Type=   1
         Sections(1).Condition=   "RecNo()=0"
         Sections(1).StyleExp=   "'tdb_Base'"
         Sections(1).SpacingAfter=   5
         Sections(1).Cells.Count=   3
         Sections(1).Cells(0).Name=   "CELL_0"
         Sections(1).Cells(0).Exp=   """LAPORAN REKAPITULASI OMZET PENJUALAN SALES"""
         Sections(1).Cells(0).PrivateStyle=   -1  'True
         Sections(1).Cells(0).Style.Name=   "<private>"
         Sections(1).Cells(0).Style.ParentName=   "tdb_Base"
         Sections(1).Cells(0).Style.Font_Name=   "Arial"
         Sections(1).Cells(0).Style.Font_Size=   14.25
         Sections(1).Cells(0).Style.Font_Bold=   -1  'True
         Sections(1).Cells(0).Style.Font_Italic=   0   'False
         Sections(1).Cells(0).Style.Font_Underline=   0   'False
         Sections(1).Cells(0).Style.Font_Strikeout=   0   'False
         Sections(1).Cells(0).Style.Font_Charset=   0
         Sections(1).Cells(0).Style.TextAlign=   1
         Sections(1).Cells(0).Style.TextVAlign=   1
         Sections(1).Cells(0).Style.TextWrap=   0   'False
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
         Sections(1).Cells(0).Style.MarginTop=   4
         Sections(1).Cells(0).Style.MarginRight=   6
         Sections(1).Cells(0).Style.MarginBottom=   4
         Sections(1).Cells(0).Style.HasBorders=   -1  'True
         Sections(1).Cells(0).Style.BorderHT=   ""
         Sections(1).Cells(0).Style.BorderHI=   ""
         Sections(1).Cells(0).Style.BorderHB=   ""
         Sections(1).Cells(0).Style.BorderVL=   ""
         Sections(1).Cells(0).Style.BorderVI=   ""
         Sections(1).Cells(0).Style.BorderVR=   ""
         Sections(1).Cells(0).Style.NoClipping=   0   'False
         Sections(1).Cells(0).Style.RTF=   0   'False
         Sections(1).Cells(0).Style.fprops=   88080385
         Sections(1).Cells(1).Name=   "CELL_1"
         Sections(1).Cells(1).Exp=   """ """
         Sections(1).Cells(1).NewLine=   -1  'True
         Sections(1).Cells(2).Name=   "CELL_2"
         Sections(1).Cells(2).Exp=   """ """
         Sections(1).Cells(2).NewLine=   -1  'True
         Sections(2).Name=   "SECTION_0"
         Sections(2).Type=   1
         Sections(2).Condition=   "RecNo()=0"
         Sections(2).StyleExp=   "'tdb_Base'"
         Sections(2).SpacingAfter=   3
         Sections(2).Cells.Count=   6
         Sections(2).Cells(0).Name=   "CELL_1"
         Sections(2).Cells(0).Exp=   """KODE SALES"""
         Sections(2).Cells(0).NewLine=   -1  'True
         Sections(2).Cells(0).Width=   15
         Sections(2).Cells(0).PrivateStyle=   -1  'True
         Sections(2).Cells(0).Style.Name=   "<private>"
         Sections(2).Cells(0).Style.ParentName=   "tdb_Base"
         Sections(2).Cells(0).Style.Font_Name=   "Arial"
         Sections(2).Cells(0).Style.Font_Size=   8.25
         Sections(2).Cells(0).Style.Font_Bold=   -1  'True
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
         Sections(2).Cells(0).Style.MarginTop=   4
         Sections(2).Cells(0).Style.MarginRight=   6
         Sections(2).Cells(0).Style.MarginBottom=   4
         Sections(2).Cells(0).Style.HasBorders=   -1  'True
         Sections(2).Cells(0).Style.BorderHT=   ""
         Sections(2).Cells(0).Style.BorderHI=   ""
         Sections(2).Cells(0).Style.BorderHB=   ""
         Sections(2).Cells(0).Style.BorderVL=   ""
         Sections(2).Cells(0).Style.BorderVI=   ""
         Sections(2).Cells(0).Style.BorderVR=   ""
         Sections(2).Cells(0).Style.NoClipping=   0   'False
         Sections(2).Cells(0).Style.RTF=   0   'False
         Sections(2).Cells(0).Style.fprops=   16777216
         Sections(2).Cells(1).Name=   "CELL_2"
         Sections(2).Cells(1).Exp=   """: "" & ""[ ""  & cKode1 & "" ] "" & "" s.d "" &  "" [ "" & cKode2 & "" ]"""
         Sections(2).Cells(1).PrivateStyle=   -1  'True
         Sections(2).Cells(1).Style.Name=   "<private>"
         Sections(2).Cells(1).Style.ParentName=   "tdb_Base"
         Sections(2).Cells(1).Style.Font_Name=   "Arial"
         Sections(2).Cells(1).Style.Font_Size=   8.25
         Sections(2).Cells(1).Style.Font_Bold=   -1  'True
         Sections(2).Cells(1).Style.Font_Italic=   0   'False
         Sections(2).Cells(1).Style.Font_Underline=   0   'False
         Sections(2).Cells(1).Style.Font_Strikeout=   0   'False
         Sections(2).Cells(1).Style.Font_Charset=   0
         Sections(2).Cells(1).Style.TextAlign=   0
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
         Sections(2).Cells(1).Style.MarginTop=   4
         Sections(2).Cells(1).Style.MarginRight=   6
         Sections(2).Cells(1).Style.MarginBottom=   4
         Sections(2).Cells(1).Style.HasBorders=   -1  'True
         Sections(2).Cells(1).Style.BorderHT=   ""
         Sections(2).Cells(1).Style.BorderHI=   ""
         Sections(2).Cells(1).Style.BorderHB=   ""
         Sections(2).Cells(1).Style.BorderVL=   ""
         Sections(2).Cells(1).Style.BorderVI=   ""
         Sections(2).Cells(1).Style.BorderVR=   ""
         Sections(2).Cells(1).Style.NoClipping=   0   'False
         Sections(2).Cells(1).Style.RTF=   0   'False
         Sections(2).Cells(1).Style.fprops=   16777216
         Sections(2).Cells(2).Name=   "CELL_4"
         Sections(2).Cells(2).Exp=   "'ANTARA TANGGAL'"
         Sections(2).Cells(2).NewLine=   -1  'True
         Sections(2).Cells(2).Width=   15
         Sections(2).Cells(2).PrivateStyle=   -1  'True
         Sections(2).Cells(2).Style.Name=   "<private>"
         Sections(2).Cells(2).Style.ParentName=   "tdb_Base"
         Sections(2).Cells(2).Style.Font_Name=   "Arial"
         Sections(2).Cells(2).Style.Font_Size=   8.25
         Sections(2).Cells(2).Style.Font_Bold=   -1  'True
         Sections(2).Cells(2).Style.Font_Italic=   0   'False
         Sections(2).Cells(2).Style.Font_Underline=   0   'False
         Sections(2).Cells(2).Style.Font_Strikeout=   0   'False
         Sections(2).Cells(2).Style.Font_Charset=   0
         Sections(2).Cells(2).Style.TextAlign=   0
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
         Sections(2).Cells(2).Style.MarginTop=   4
         Sections(2).Cells(2).Style.MarginRight=   6
         Sections(2).Cells(2).Style.MarginBottom=   4
         Sections(2).Cells(2).Style.HasBorders=   -1  'True
         Sections(2).Cells(2).Style.BorderHT=   ""
         Sections(2).Cells(2).Style.BorderHI=   ""
         Sections(2).Cells(2).Style.BorderHB=   ""
         Sections(2).Cells(2).Style.BorderVL=   ""
         Sections(2).Cells(2).Style.BorderVI=   ""
         Sections(2).Cells(2).Style.BorderVR=   ""
         Sections(2).Cells(2).Style.NoClipping=   0   'False
         Sections(2).Cells(2).Style.RTF=   0   'False
         Sections(2).Cells(2).Style.fprops=   16777216
         Sections(2).Cells(3).Name=   "CELL_5"
         Sections(2).Cells(3).Exp=   """: "" & ""[ ""  & dTgl1 & "" ] "" & "" s.d "" &  "" [ "" & dTgl2 & "" ]"""
         Sections(2).Cells(3).PrivateStyle=   -1  'True
         Sections(2).Cells(3).Style.Name=   "<private>"
         Sections(2).Cells(3).Style.ParentName=   "tdb_Base"
         Sections(2).Cells(3).Style.Font_Name=   "Arial"
         Sections(2).Cells(3).Style.Font_Size=   8.25
         Sections(2).Cells(3).Style.Font_Bold=   -1  'True
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
         Sections(2).Cells(3).Style.MarginTop=   4
         Sections(2).Cells(3).Style.MarginRight=   6
         Sections(2).Cells(3).Style.MarginBottom=   4
         Sections(2).Cells(3).Style.HasBorders=   -1  'True
         Sections(2).Cells(3).Style.BorderHT=   ""
         Sections(2).Cells(3).Style.BorderHI=   ""
         Sections(2).Cells(3).Style.BorderHB=   ""
         Sections(2).Cells(3).Style.BorderVL=   ""
         Sections(2).Cells(3).Style.BorderVI=   ""
         Sections(2).Cells(3).Style.BorderVR=   ""
         Sections(2).Cells(3).Style.NoClipping=   0   'False
         Sections(2).Cells(3).Style.RTF=   0   'False
         Sections(2).Cells(3).Style.fprops=   16777216
         Sections(2).Cells(4).Name=   "CELL_7"
         Sections(2).Cells(4).Exp=   """Tanggal Cetak """
         Sections(2).Cells(4).NewLine=   -1  'True
         Sections(2).Cells(4).Width=   15
         Sections(2).Cells(4).PrivateStyle=   -1  'True
         Sections(2).Cells(4).Style.Name=   "<private>"
         Sections(2).Cells(4).Style.ParentName=   "tdb_Base"
         Sections(2).Cells(4).Style.Font_Name=   "Arial"
         Sections(2).Cells(4).Style.Font_Size=   8.25
         Sections(2).Cells(4).Style.Font_Bold=   -1  'True
         Sections(2).Cells(4).Style.Font_Italic=   0   'False
         Sections(2).Cells(4).Style.Font_Underline=   0   'False
         Sections(2).Cells(4).Style.Font_Strikeout=   0   'False
         Sections(2).Cells(4).Style.Font_Charset=   0
         Sections(2).Cells(4).Style.TextAlign=   0
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
         Sections(2).Cells(4).Style.MarginTop=   4
         Sections(2).Cells(4).Style.MarginRight=   6
         Sections(2).Cells(4).Style.MarginBottom=   4
         Sections(2).Cells(4).Style.HasBorders=   -1  'True
         Sections(2).Cells(4).Style.BorderHT=   ""
         Sections(2).Cells(4).Style.BorderHI=   ""
         Sections(2).Cells(4).Style.BorderHB=   ""
         Sections(2).Cells(4).Style.BorderVL=   ""
         Sections(2).Cells(4).Style.BorderVI=   ""
         Sections(2).Cells(4).Style.BorderVR=   ""
         Sections(2).Cells(4).Style.NoClipping=   0   'False
         Sections(2).Cells(4).Style.RTF=   0   'False
         Sections(2).Cells(4).Style.fprops=   16777216
         Sections(2).Cells(5).Name=   "CELL_8"
         Sections(2).Cells(5).Exp=   """: "" &   Format(Date,""dd-MM-yyyy"")"
         Sections(2).Cells(5).PrivateStyle=   -1  'True
         Sections(2).Cells(5).Style.Name=   "<private>"
         Sections(2).Cells(5).Style.ParentName=   "tdb_Base"
         Sections(2).Cells(5).Style.Font_Name=   "Arial"
         Sections(2).Cells(5).Style.Font_Size=   8.25
         Sections(2).Cells(5).Style.Font_Bold=   -1  'True
         Sections(2).Cells(5).Style.Font_Italic=   0   'False
         Sections(2).Cells(5).Style.Font_Underline=   0   'False
         Sections(2).Cells(5).Style.Font_Strikeout=   0   'False
         Sections(2).Cells(5).Style.Font_Charset=   0
         Sections(2).Cells(5).Style.TextAlign=   0
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
         Sections(2).Cells(5).Style.MarginTop=   4
         Sections(2).Cells(5).Style.MarginRight=   6
         Sections(2).Cells(5).Style.MarginBottom=   4
         Sections(2).Cells(5).Style.HasBorders=   -1  'True
         Sections(2).Cells(5).Style.BorderHT=   ""
         Sections(2).Cells(5).Style.BorderHI=   ""
         Sections(2).Cells(5).Style.BorderHB=   ""
         Sections(2).Cells(5).Style.BorderVL=   ""
         Sections(2).Cells(5).Style.BorderVI=   ""
         Sections(2).Cells(5).Style.BorderVR=   ""
         Sections(2).Cells(5).Style.NoClipping=   0   'False
         Sections(2).Cells(5).Style.RTF=   0   'False
         Sections(2).Cells(5).Style.fprops=   16777216
         Sections(3).Name=   "SECTION_4"
         Sections(3).Type=   3
         Sections(3).StyleExp=   "'tdb_Header'"
         Sections(3).Tabulator=   "tdb_Header"
         Sections(3).Cells.Count=   11
         Sections(3).Cells(0).Name=   "CELL_0"
         Sections(3).Cells(0).Exp=   """No"""
         Sections(3).Cells(0).Merge=   3
         Sections(3).Cells(1).Name=   "CELL_1"
         Sections(3).Cells(1).Exp=   """SALESMAN"""
         Sections(3).Cells(1).Merge=   3
         Sections(3).Cells(2).Name=   "CELL_3"
         Sections(3).Cells(2).Exp=   """PENJUALAN"""
         Sections(3).Cells(2).CellSpan=   3
         Sections(3).Cells(3).Name=   "CELL_4"
         Sections(3).Cells(4).Name=   "CELL_5"
         Sections(3).Cells(5).Name=   "CELL_6"
         Sections(3).Cells(5).Exp=   """RETUR PENJUALAN"""
         Sections(3).Cells(5).CellSpan=   3
         Sections(3).Cells(6).Name=   "CELL_7"
         Sections(3).Cells(7).Name=   "CELL_8"
         Sections(3).Cells(8).Name=   "CELL_10"
         Sections(3).Cells(8).Exp=   """Pembelian"""
         Sections(3).Cells(8).Merge=   3
         Sections(3).Cells(9).Name=   "CELL_11"
         Sections(3).Cells(9).Exp=   """Komisi"""
         Sections(3).Cells(9).Merge=   3
         Sections(3).Cells(10).Name=   "CELL_9"
         Sections(3).Cells(10).Exp=   """OMZET BERSIH"""
         Sections(3).Cells(10).Merge=   3
         Sections(4).Name=   "tdb_Header"
         Sections(4).Type=   3
         Sections(4).StyleExp=   "'tdb_Header'"
         Sections(4).Tabulator=   "tdb_Body"
         Sections(4).Cells.Count=   11
         Sections(4).Cells(0).Name=   "CELL_7"
         Sections(4).Cells(0).Exp=   """No"""
         Sections(4).Cells(0).Merge=   3
         Sections(4).Cells(1).Name=   "CELL_0"
         Sections(4).Cells(1).Exp=   """SALESMAN"""
         Sections(4).Cells(1).Merge=   3
         Sections(4).Cells(2).Name=   "CELL_3"
         Sections(4).Cells(2).Exp=   """SubTotal"""
         Sections(4).Cells(3).Name=   "CELL_4"
         Sections(4).Cells(3).Exp=   """Disc"""
         Sections(4).Cells(4).Name=   "CELL_5"
         Sections(4).Cells(4).Exp=   """Total"""
         Sections(4).Cells(5).Name=   "CELL_8"
         Sections(4).Cells(5).Exp=   """SubTotal"""
         Sections(4).Cells(6).Name=   "CELL_9"
         Sections(4).Cells(6).Exp=   """Disc"""
         Sections(4).Cells(7).Name=   "CELL_10"
         Sections(4).Cells(7).Exp=   """SubTotal"""
         Sections(4).Cells(8).Name=   "CELL_12"
         Sections(4).Cells(8).Exp=   """Pembelian"""
         Sections(4).Cells(8).Merge=   3
         Sections(4).Cells(9).Name=   "CELL_13"
         Sections(4).Cells(9).Exp=   """Komisi"""
         Sections(4).Cells(9).Merge=   3
         Sections(4).Cells(10).Name=   "CELL_11"
         Sections(4).Cells(10).Exp=   """OMZET BERSIH"""
         Sections(4).Cells(10).Merge=   3
         Sections(5).Name=   "tdb_Body"
         Sections(5).Type=   4
         Sections(5).StyleExp=   "'tdb_Body'"
         Sections(5).Cells.Count=   11
         Sections(5).Cells(0).Name=   "CELL_6"
         Sections(5).Cells(0).Exp=   "Sum(1,False)"
         Sections(5).Cells(0).Width=   3
         Sections(5).Cells(0).PrivateStyle=   -1  'True
         Sections(5).Cells(0).Style.Name=   "<private>"
         Sections(5).Cells(0).Style.ParentName=   "tdb_Body"
         Sections(5).Cells(0).Style.Font_Name=   "Arial"
         Sections(5).Cells(0).Style.Font_Size=   8.25
         Sections(5).Cells(0).Style.Font_Bold=   0   'False
         Sections(5).Cells(0).Style.Font_Italic=   0   'False
         Sections(5).Cells(0).Style.Font_Underline=   0   'False
         Sections(5).Cells(0).Style.Font_Strikeout=   0   'False
         Sections(5).Cells(0).Style.Font_Charset=   0
         Sections(5).Cells(0).Style.TextAlign=   0
         Sections(5).Cells(0).Style.TextVAlign=   1
         Sections(5).Cells(0).Style.TextWrap=   0   'False
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
         Sections(5).Cells(0).Style.MarginTop=   6
         Sections(5).Cells(0).Style.MarginRight=   6
         Sections(5).Cells(0).Style.MarginBottom=   6
         Sections(5).Cells(0).Style.HasBorders=   -1  'True
         Sections(5).Cells(0).Style.BorderHT=   "tdb_Quart"
         Sections(5).Cells(0).Style.BorderHI=   "tdb_Quart"
         Sections(5).Cells(0).Style.BorderHB=   "tdb_Double"
         Sections(5).Cells(0).Style.BorderVL=   "tdb_Single"
         Sections(5).Cells(0).Style.BorderVI=   "tdb_Single"
         Sections(5).Cells(0).Style.BorderVR=   "tdb_Single"
         Sections(5).Cells(0).Style.NoClipping=   0   'False
         Sections(5).Cells(0).Style.RTF=   0   'False
         Sections(5).Cells(0).Style.fprops=   1
         Sections(5).Cells(1).Name=   "CELL_0"
         Sections(5).Cells(1).Exp=   "Nama"
         Sections(5).Cells(1).PrivateStyle=   -1  'True
         Sections(5).Cells(1).Style.Name=   "<private>"
         Sections(5).Cells(1).Style.ParentName=   "tdb_Body"
         Sections(5).Cells(1).Style.Font_Name=   "Arial"
         Sections(5).Cells(1).Style.Font_Size=   8.25
         Sections(5).Cells(1).Style.Font_Bold=   0   'False
         Sections(5).Cells(1).Style.Font_Italic=   0   'False
         Sections(5).Cells(1).Style.Font_Underline=   0   'False
         Sections(5).Cells(1).Style.Font_Strikeout=   0   'False
         Sections(5).Cells(1).Style.Font_Charset=   0
         Sections(5).Cells(1).Style.TextAlign=   0
         Sections(5).Cells(1).Style.TextVAlign=   1
         Sections(5).Cells(1).Style.TextWrap=   0   'False
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
         Sections(5).Cells(1).Style.BorderHT=   "tdb_Quart"
         Sections(5).Cells(1).Style.BorderHI=   "tdb_Quart"
         Sections(5).Cells(1).Style.BorderHB=   "tdb_Double"
         Sections(5).Cells(1).Style.BorderVL=   "tdb_Single"
         Sections(5).Cells(1).Style.BorderVI=   "tdb_Single"
         Sections(5).Cells(1).Style.BorderVR=   "tdb_Single"
         Sections(5).Cells(1).Style.NoClipping=   0   'False
         Sections(5).Cells(1).Style.RTF=   0   'False
         Sections(5).Cells(1).Style.fprops=   1
         Sections(5).Cells(2).Name=   "CELL_3"
         Sections(5).Cells(2).Exp=   "SubTotal1"
         Sections(5).Cells(2).Width=   10
         Sections(5).Cells(2).CallExpression=   -1  'True
         Sections(5).Cells(2).Format=   "###,###,###,###,##0.00"
         Sections(5).Cells(3).Name=   "CELL_4"
         Sections(5).Cells(3).Exp=   "Disc1"
         Sections(5).Cells(3).Width=   7
         Sections(5).Cells(3).CallExpression=   -1  'True
         Sections(5).Cells(3).Format=   "###,###,###,###,##0.00"
         Sections(5).Cells(4).Name=   "CELL_5"
         Sections(5).Cells(4).Exp=   "Total1"
         Sections(5).Cells(4).Width=   10
         Sections(5).Cells(4).CallExpression=   -1  'True
         Sections(5).Cells(4).Format=   "###,###,###,###,##0.00"
         Sections(5).Cells(5).Name=   "CELL_7"
         Sections(5).Cells(5).Exp=   "SubTotal2"
         Sections(5).Cells(5).Width=   10
         Sections(5).Cells(5).Format=   "###,###,###,###,##0.00"
         Sections(5).Cells(6).Name=   "CELL_8"
         Sections(5).Cells(6).Exp=   "Disc2"
         Sections(5).Cells(6).Width=   7
         Sections(5).Cells(6).Format=   "###,###,###,###,##0.00"
         Sections(5).Cells(7).Name=   "CELL_9"
         Sections(5).Cells(7).Exp=   "Total2"
         Sections(5).Cells(7).Width=   10
         Sections(5).Cells(7).Format=   "###,###,###,###,##0.00"
         Sections(5).Cells(8).Name=   "CELL_11"
         Sections(5).Cells(8).Exp=   "Pembelian"
         Sections(5).Cells(8).Width=   10
         Sections(5).Cells(8).Format=   "###,###,###,###,##0.00"
         Sections(5).Cells(9).Name=   "CELL_12"
         Sections(5).Cells(9).Exp=   "Komisi"
         Sections(5).Cells(9).Width=   10
         Sections(5).Cells(9).Format=   "###,###,###,###,##0.00"
         Sections(5).Cells(10).Name=   "CELL_10"
         Sections(5).Cells(10).Exp=   "Bersih"
         Sections(5).Cells(10).Width=   10
         Sections(5).Cells(10).Format=   "###,###,###,###,##0.00"
         Sections(6).Name=   "SECTION_7"
         Sections(6).Condition=   "IsLastRec()"
         Sections(6).StyleExp=   "'Tdb_Footer'"
         Sections(6).Tabulator=   "tdb_Body"
         Sections(6).Cells.Count=   11
         Sections(6).Cells(0).Name=   "CELL_0"
         Sections(6).Cells(0).Exp=   """Grand Total"""
         Sections(6).Cells(0).CellSpan=   2
         Sections(6).Cells(1).Name=   "CELL_1"
         Sections(6).Cells(2).Name=   "CELL_4"
         Sections(6).Cells(2).Exp=   "Sum(SubTotal1,False)"
         Sections(6).Cells(2).Format=   "###,###,###,###,##0.00"
         Sections(6).Cells(3).Name=   "CELL_5"
         Sections(6).Cells(3).Exp=   "Sum(Disc1,False)"
         Sections(6).Cells(3).Format=   "###,###,###,###,##0.00"
         Sections(6).Cells(4).Name=   "CELL_6"
         Sections(6).Cells(4).Exp=   "Sum(Total1,False)"
         Sections(6).Cells(4).Format=   "###,###,###,###,##0.00"
         Sections(6).Cells(5).Name=   "CELL_7"
         Sections(6).Cells(5).Exp=   "Sum(SubTotal2,False)"
         Sections(6).Cells(5).Format=   "###,###,###,###,##0.00"
         Sections(6).Cells(6).Name=   "CELL_8"
         Sections(6).Cells(6).Exp=   "Sum(Disc2,False)"
         Sections(6).Cells(6).Format=   "###,###,###,###,##0.00"
         Sections(6).Cells(7).Name=   "CELL_9"
         Sections(6).Cells(7).Exp=   "Sum(Total2,False)"
         Sections(6).Cells(7).Format=   "###,###,###,###,##0.00"
         Sections(6).Cells(8).Name=   "CELL_11"
         Sections(6).Cells(8).Exp=   "Sum(Pembelian,False)"
         Sections(6).Cells(8).Format=   "###,###,###,###,##0.00"
         Sections(6).Cells(9).Name=   "CELL_12"
         Sections(6).Cells(9).Exp=   "Sum(Komisi,False)"
         Sections(6).Cells(9).Format=   "###,###,###,###,##0.00"
         Sections(6).Cells(10).Name=   "CELL_10"
         Sections(6).Cells(10).Exp=   "Sum(Bersih,False)"
         Sections(6).Cells(10).Format=   "###,###,###,###,##0.00"
         Styles.Count    =   5
         Styles(0).Name  =   "tdb_Base"
         Styles(0).ParentName=   ""
         Styles(0).Font_Name=   "Arial"
         Styles(0).Font_Size=   8.25
         Styles(0).Font_Charset=   0
         Styles(0).TextAlign=   0
         Styles(0).TextVAlign=   1
         Styles(0).TextWrap=   0   'False
         Styles(0).MarginTop=   4
         Styles(0).MarginBottom=   4
         Styles(1).Name  =   "Tdb_Footer"
         Styles(1).ParentName=   "tdb_Body"
         Styles(1).Font_Name=   "Arial"
         Styles(1).Font_Size=   8.25
         Styles(1).Font_Bold=   -1  'True
         Styles(1).Font_Charset=   0
         Styles(1).TextAlign=   0
         Styles(1).TextVAlign=   1
         Styles(1).TextWrap=   0   'False
         Styles(1).MarginTop=   4
         Styles(1).MarginBottom=   4
         Styles(1).BorderHT=   "tdb_Double"
         Styles(1).BorderHI=   "tdb_Double"
         Styles(1).fprops=   16875520
         Styles(2).Name  =   "Tdb_GarisBawah"
         Styles(2).ParentName=   "tdb_Base"
         Styles(2).Font_Name=   "Arial"
         Styles(2).Font_Size=   8.25
         Styles(2).Font_Bold=   -1  'True
         Styles(2).Font_Charset=   0
         Styles(2).TextAlign=   0
         Styles(2).TextVAlign=   1
         Styles(2).TextWrap=   0   'False
         Styles(2).MarginTop=   4
         Styles(2).MarginBottom=   4
         Styles(2).BorderHB=   "tdb_Double"
         Styles(2).fprops=   16908288
         Styles(3).Name  =   "tdb_Body"
         Styles(3).ParentName=   "tdb_Base"
         Styles(3).Font_Name=   "Arial"
         Styles(3).Font_Size=   8.25
         Styles(3).Font_Charset=   0
         Styles(3).TextAlign=   2
         Styles(3).TextVAlign=   1
         Styles(3).TextWrap=   0   'False
         Styles(3).BorderHT=   "tdb_Quart"
         Styles(3).BorderHI=   "tdb_Quart"
         Styles(3).BorderHB=   "tdb_Double"
         Styles(3).BorderVL=   "tdb_Single"
         Styles(3).BorderVI=   "tdb_Single"
         Styles(3).BorderVR=   "tdb_Single"
         Styles(3).fprops=   25153539
         Styles(4).Name  =   "tdb_Header"
         Styles(4).ParentName=   "tdb_Base"
         Styles(4).Font_Name=   "Arial"
         Styles(4).Font_Size=   8.25
         Styles(4).Font_Bold=   -1  'True
         Styles(4).Font_Charset=   0
         Styles(4).TextAlign=   1
         Styles(4).TextVAlign=   1
         Styles(4).TextWrap=   0   'False
         Styles(4).BorderHT=   "tdb_Double"
         Styles(4).BorderHI=   "tdb_Double"
         Styles(4).BorderHB=   "tdb_Double"
         Styles(4).BorderVL=   "tdb_Single"
         Styles(4).BorderVI=   "tdb_Single"
         Styles(4).BorderVR=   "tdb_Single"
         Styles(4).fprops=   23056385
         Lines.Count     =   3
         Lines(0).Name   =   "tdb_Single"
         Lines(0).Thickness=   4
         Lines(1).Name   =   "tdb_Double"
         Lines(1).Thickness=   5
         Lines(2).Name   =   "tdb_Quart"
         Lines(2).Thickness=   1
         Profiles.Count  =   1
         Profiles(0).Name=   "PROFILE_0"
         Profiles(0).Active=   -1  'True
         Profiles(0).PreviewNoMinimize=   -1  'True
         Profiles(0).PreviewNoMaximize=   -1  'True
         Profiles(0).PreviewNoResize=   -1  'True
         Profiles(0).PreviewMaximized=   -1  'True
         Profiles(0).PreviewNoSaveLoad=   -1  'True
         Profiles(0).PrinterMarginLeft=   12
         Profiles(0).PrinterMarginTop=   10
         Profiles(0).PrinterMarginRight=   10
         Profiles(0).PrinterMarginBottom=   10
         Profiles(0).PrinterLandscape=   -1  'True
         Profiles(0).PrinterMargins_set=   -1  'True
         Profiles(0).PrinterLandscape_set=   -1  'True
         Profiles(0).PrinterPaperUserSize_set=   -1  'True
      End
   End
   Begin BiSAFramProject.BiSAFrame BiSAFrame3 
      Height          =   645
      Left            =   15
      Top             =   1080
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
      BorderStyle     =   4
      BackColor       =   -2147483633
      Begin BiSAButtonProject.BiSAButton cmdKeluar 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   5790
         TabIndex        =   5
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
         Picture         =   "rptRekapitulasiOmzetSales.frx":0000
      End
      Begin BiSAButtonProject.BiSAButton cmdPreview 
         Height          =   435
         Left            =   5370
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
         Picture         =   "rptRekapitulasiOmzetSales.frx":00A6
      End
   End
End
Attribute VB_Name = "rptRekapitulasiOmzetSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbData As New ADODB.Recordset
Dim dbData1 As New ADODB.Recordset
Dim dbData2 As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data
Dim vaArray As New XArrayDB
Dim n As Single

Private Sub cKode_ButtonClick(Index As Integer)
  Set dbData = objData.PICK(GetDSN, "salesman", "kodesalesman", cKode(Index), "kodesalesman,nama")
  If Not dbData.EOF Then
    cKode(Index).Text = GetNull(dbData!kodesalesman, "")
  End If
End Sub

Private Sub cKode_Validate(Index As Integer, Cancel As Boolean)
  If cKode(Index).LastKey = 13 Then
    cKode_ButtonClick (Index)
  End If
End Sub

Private Sub cmdKeluar_Click()
  Unload Me
End Sub

Private Sub cmdPreview_Click()
  GetData
End Sub

Private Sub Form_Load()

  CenterForm Me
  SetIcon Me.hWnd
  GetMinMax "salesman", cKode, "kodesalesman"
  dAwal.Value = BOM(Date)
  dAkhir.Value = Date
  
  TabIndex cKode(0), n
  TabIndex cKode(1), n
  TabIndex dAwal, n
  TabIndex dAkhir, n
  TabIndex cmdPreview, n
  
End Sub

Private Sub GetData()
Dim cSQL As String
Dim Disc As Double
  
  cSQL = " SELECT t.tgl,t.nomorpenjualan,t.kodesalesman as kodesalesman,s.nama,s.alamat,SUM(t.subtotal) as subtotal1,"
  cSQL = cSQL & " SUM(t.discount)as discount1, SUM(t.total) as total1,SUM(t.komisi) as komisi"
  cSQL = cSQL & " FROM totpenjualan t"
  cSQL = cSQL & " LEFT JOIN salesman s on s.kodesalesman=t.kodesalesman"
  cSQL = cSQL & " WHERE t.kodesalesman >= '" & cKode(0).Text & "' AND t.kodesalesman <= '" & cKode(1).Text & "'"
  cSQL = cSQL & " AND t.tgl >= '" & Format(dAwal.Value, "yyyy-MM-dd") & "' AND t.Tgl <='" & Format(dAkhir.Value, "yyyy-MM-dd") & "'"
  cSQL = cSQL & " GROUP By t.kodesalesman"
  cSQL = cSQL & " ORDER By t.kodesalesman"
  Set dbData = objData.Sql(GetDSN, cSQL)
  If Not dbData.EOF Then
    FrmPB.InitPB dbData.RecordCount
    n = 0
    vaArray.ReDim 0, dbData.RecordCount - 1, 0, 13
    Do While Not dbData.EOF
      FrmPB.RunPB
      vaArray(n, 0) = (dbData!tgl)
      vaArray(n, 1) = (dbData!kodesalesman)
      vaArray(n, 2) = (dbData!nama)
      vaArray(n, 3) = (dbData!alamat)
      vaArray(n, 4) = ""
      vaArray(n, 5) = (dbData!subtotal1)
      vaArray(n, 6) = (dbData!discount1)
      vaArray(n, 7) = (dbData!total1)
      vaArray(n, 12) = GetNull(dbData!komisi)
      Set dbData1 = objData.Browse(GetDSN, "totrtnpenjualan t", "SUM(t.subtotal)as subtotal,SUM(t.discount)as discount1, SUM(t.total) as total", "tp.kodesalesman", sisAssign, vaArray(n, 1), " AND t.tgl >= '" & Format(dAwal.Value, "yyyy-mm-dd") & "' AND t.tgl <='" & Format(dAkhir.Value, "yyyy-mm-dd") & "' GROUP By tp.kodesalesman", "tp.kodesalesman", Array("LEFT JOIN totpenjualan tp on tp.nomorpenjualan = t.nomorpenjualan"))
      If Not dbData1.EOF Then
        vaArray(n, 8) = GetNull((dbData1!Subtotal), 0)
        vaArray(n, 9) = GetNull((dbData1!discount1), 0)
        vaArray(n, 10) = GetNull((dbData1!Total), 0)
      Else
        vaArray(n, 8) = 0
        vaArray(n, 9) = 0
        vaArray(n, 10) = 0
      End If
    
      Set dbData2 = objData.Browse(GetDSN, "penjualan p", "SUM(p.qty*p.hb) as pembelian", "t.kodesalesman", sisAssign, vaArray(n, 1), " AND T.Tgl >= '" & Format(dAwal.Value, "yyyy-MM-dd") & "' AND t.Tgl <= '" & Format(dAkhir.Value, "yyyy-MM-dd") & "'", , Array("LEFT JOIN totpenjualan t on t.nomorpenjualan = p.nomorpenjualan"))
      If Not dbData2.EOF Then
        vaArray(n, 11) = GetNull(dbData2!pembelian)
      Else
        vaArray(n, 11) = 0
      End If
      
      Dim nTemporary As Double
      Dim nTempKomisi As Double
      
      nTemporary = vaArray(n, 7) - vaArray(n, 10) - vaArray(n, 11) + vaArray(n, 12)
      nTempKomisi = nTemporary
       
      vaArray(n, 13) = nTempKomisi
      
      dbData.MoveNext
      n = n + 1
    Loop
    FrmPB.EndPB
    
    With RptOmzetSales
      .Parameters("cKode1").ValueExpression = "'" & cKode(0).Text & "'"
      .Parameters("cKode2").ValueExpression = "'" & cKode(1).Text & "'"
      .Parameters("dTgl1").ValueExpression = "'" & Format(dAwal.Value, "dd-MM-yyyy") & "'"
      .Parameters("dTgl2").ValueExpression = "'" & Format(dAkhir.Value, "dd-MM-yyyy") & "'"
       Set .Array = vaArray
      .Refresh
      .PrintPreview
    End With
  Else
    MsgBox "Data tidak ada.", vbInformation, "Laporan Omzet Sales"
    Exit Sub
  End If
End Sub

