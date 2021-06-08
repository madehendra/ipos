VERSION 5.00
Object = "{34C98750-1217-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Button.ocx"
Object = "{45D2FD98-1218-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA Frame.ocx"
Object = "{80D06F5A-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA NumberBox.ocx"
Object = "{80D0704C-0C2B-11D8-A794-0008C7CAB078}#1.0#0"; "BiSA TextBox.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form trFormKasir 
   BackColor       =   &H80000004&
   BorderStyle     =   0  'None
   Caption         =   "Form Kasir"
   ClientHeight    =   7890
   ClientLeft      =   60
   ClientTop       =   45
   ClientWidth     =   13710
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   13710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin BiSAFramProject.BiSAFrame BiSAFrame1 
      Height          =   4185
      Left            =   135
      Top             =   105
      Width           =   13425
      _ExtentX        =   23680
      _ExtentY        =   7382
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
      Begin VB.Frame Frame1 
         Height          =   3840
         Left            =   8220
         TabIndex        =   5
         Top             =   240
         Width           =   5085
         Begin VB.CheckBox chkKartu 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Pembayaran Dengan Kartu"
            Height          =   300
            Left            =   1740
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   330
            Width           =   2370
         End
         Begin BiSATextBoxProject.BiSABrowse cKodeKartu 
            Height          =   330
            Left            =   315
            TabIndex        =   7
            Top             =   735
            Width           =   3675
            _ExtentX        =   6482
            _ExtentY        =   582
            Text            =   "12345678"
            BorderStyle     =   0
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
            Appearance      =   0
            Caption         =   "Kartu ID"
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
         Begin BiSATextBoxProject.BiSABrowse cNamaKartu 
            Height          =   330
            Left            =   330
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   1080
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   582
            Text            =   "12345678"
            BorderStyle     =   0
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
            Caption         =   "Kartu"
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
         Begin BiSATextBoxProject.BiSABrowse cNomorKartu 
            Height          =   330
            Left            =   345
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   1455
            Width           =   4380
            _ExtentX        =   7726
            _ExtentY        =   582
            Text            =   "12345678"
            BorderStyle     =   0
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
            Caption         =   "Nomor"
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
         Begin BiSATextBoxProject.BiSABrowse cNomorTrace 
            Height          =   330
            Left            =   345
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   1830
            Width           =   3675
            _ExtentX        =   6482
            _ExtentY        =   582
            Text            =   "12345678"
            BorderStyle     =   0
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
            Caption         =   "Trace No Trx"
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
         Begin BiSANumberBoxProject.BiSANumberBox nFee 
            Height          =   330
            Left            =   1560
            TabIndex        =   12
            Top             =   3000
            Width           =   2550
            _ExtentX        =   4498
            _ExtentY        =   582
            Appearance      =   0
            Enabled         =   0   'False
            BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483633
            Caption         =   "Fee %"
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
         Begin BiSANumberBoxProject.BiSANumberBox nTotalKartu 
            Height          =   330
            Left            =   1575
            TabIndex        =   13
            Top             =   3375
            Width           =   3390
            _ExtentX        =   5980
            _ExtentY        =   582
            Appearance      =   0
            MinValue        =   0
            Enabled         =   0   'False
            BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483633
            Caption         =   "Total"
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
         Begin BiSATextBoxProject.BiSABrowse cNamaPemegangKartu 
            Height          =   330
            Left            =   345
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   2190
            Width           =   4605
            _ExtentX        =   8123
            _ExtentY        =   582
            Text            =   "12345678"
            BorderStyle     =   0
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
            Caption         =   "Nama"
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
         Begin BiSANumberBoxProject.BiSANumberBox nDP 
            Height          =   330
            Left            =   330
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   2565
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   582
            Appearance      =   0
            BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Tunai/DP"
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
      End
      Begin BiSANumberBoxProject.BiSANumberBox nJumlahYgHarusDibayar 
         Height          =   405
         Left            =   2985
         TabIndex        =   1
         Top             =   480
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   714
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
         CaptionWidth    =   2500
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
      Begin BiSANumberBoxProject.BiSANumberBox nTunai 
         Height          =   750
         Left            =   375
         TabIndex        =   3
         Top             =   1755
         Width           =   7395
         _ExtentX        =   13044
         _ExtentY        =   1323
         BorderStyle     =   0
         Appearance      =   0
         Decimals        =   0
         xxxx            =   900000000
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Tunai"
         CaptionWidth    =   2500
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin BiSANumberBoxProject.BiSANumberBox nKembalian 
         Height          =   765
         Left            =   375
         TabIndex        =   4
         Top             =   2580
         Width           =   7410
         _ExtentX        =   13070
         _ExtentY        =   1349
         BorderStyle     =   0
         Appearance      =   0
         Decimals        =   0
         Enabled         =   0   'False
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Kembali"
         CaptionWidth    =   2500
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin BiSANumberBoxProject.BiSANumberBox nVoucher 
         Height          =   750
         Left            =   375
         TabIndex        =   2
         Top             =   945
         Width           =   7395
         _ExtentX        =   13044
         _ExtentY        =   1323
         BorderStyle     =   0
         Appearance      =   0
         Decimals        =   0
         Enabled         =   0   'False
         BeginProperty font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Caption         =   "Voucher"
         CaptionWidth    =   2500
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   405
         TabIndex        =   0
         Top             =   450
         Width           =   1500
      End
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Height          =   2835
      Left            =   165
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   4380
      Width           =   13395
      _ExtentX        =   23627
      _ExtentY        =   5001
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   4
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "FAKTUR"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "TGL"
      Columns(2).DataField=   ""
      Columns(2).NumberFormat=   "dd-MM-yyyy"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "KETERANGAN"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "DEBET"
      Columns(4).DataField=   ""
      Columns(4).NumberFormat=   "###,###,###,###,##0.00"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "FLAG"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "FLAGID"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   7
      Splits(0)._UserFlags=   0
      Splits(0).AllowSizing=   -1  'True
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).FetchRowStyle=   -1  'True
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=7"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=609"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=529"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=197124"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=3704"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3625"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=197124"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=3440"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=3360"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=197121"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=7514"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=7435"
      Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=197124"
      Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(21)=   "Column(4).Width=3096"
      Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=3016"
      Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=197122"
      Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(26)=   "Column(5).Width=1826"
      Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=1746"
      Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=197124"
      Splits(0)._ColumnProps(30)=   "Column(5).FetchStyle=1"
      Splits(0)._ColumnProps(31)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(32)=   "Column(6).Width=2725"
      Splits(0)._ColumnProps(33)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(34)=   "Column(6)._WidthInPix=2646"
      Splits(0)._ColumnProps(35)=   "Column(6)._ColStyle=197124"
      Splits(0)._ColumnProps(36)=   "Column(6).Order=7"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      BorderStyle     =   0
      ColumnFooters   =   -1  'True
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      Caption         =   "VOUCHER -TOP UP"
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
      _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=46,.parent=13,.bgcolor=&H80000005&"
      _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
      _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
      _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17,.bgcolor=&H8000000F&"
      _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=70,.parent=13,.alignment=2"
      _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=67,.parent=14"
      _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=68,.parent=15"
      _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=69,.parent=17"
      _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
      _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
      _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
      _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
      _StyleDefs(53)  =   "Splits(0).Columns(4).Style:id=58,.parent=13,.alignment=1"
      _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=55,.parent=14"
      _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=56,.parent=15"
      _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=57,.parent=17"
      _StyleDefs(57)  =   "Splits(0).Columns(5).Style:id=54,.parent=13"
      _StyleDefs(58)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
      _StyleDefs(59)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
      _StyleDefs(60)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
      _StyleDefs(61)  =   "Splits(0).Columns(6).Style:id=66,.parent=13"
      _StyleDefs(62)  =   "Splits(0).Columns(6).HeadingStyle:id=63,.parent=14"
      _StyleDefs(63)  =   "Splits(0).Columns(6).FooterStyle:id=64,.parent=15"
      _StyleDefs(64)  =   "Splits(0).Columns(6).EditorStyle:id=65,.parent=17"
      _StyleDefs(65)  =   "Named:id=33:Normal"
      _StyleDefs(66)  =   ":id=33,.parent=0"
      _StyleDefs(67)  =   "Named:id=34:Heading"
      _StyleDefs(68)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(69)  =   ":id=34,.wraptext=-1"
      _StyleDefs(70)  =   "Named:id=35:Footing"
      _StyleDefs(71)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(72)  =   "Named:id=36:Selected"
      _StyleDefs(73)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(74)  =   "Named:id=37:Caption"
      _StyleDefs(75)  =   ":id=37,.parent=34,.alignment=2,.bold=0,.fontsize=825,.italic=0,.underline=0"
      _StyleDefs(76)  =   ":id=37,.strikethrough=0,.charset=0"
      _StyleDefs(77)  =   ":id=37,.fontname=Tahoma"
      _StyleDefs(78)  =   "Named:id=38:HighlightRow"
      _StyleDefs(79)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(80)  =   "Named:id=39:EvenRow"
      _StyleDefs(81)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(82)  =   "Named:id=40:OddRow"
      _StyleDefs(83)  =   ":id=40,.parent=33"
      _StyleDefs(84)  =   "Named:id=41:RecordSelector"
      _StyleDefs(85)  =   ":id=41,.parent=34"
      _StyleDefs(86)  =   "Named:id=42:FilterBar"
      _StyleDefs(87)  =   ":id=42,.parent=33"
   End
   Begin BiSAButtonProject.BiSAButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   435
      Left            =   12225
      TabIndex        =   17
      Top             =   7305
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   767
      Caption         =   "     &Cancel"
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
      Picture         =   "trFormKasir.frx":0000
   End
   Begin BiSAButtonProject.BiSAButton cmdOK 
      Height          =   435
      Left            =   10890
      TabIndex        =   16
      Top             =   7305
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   767
      Caption         =   "    &OK"
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
      Picture         =   "trFormKasir.frx":00A6
   End
   Begin BiSAButtonProject.BiSAButton BiSAButton1 
      Height          =   435
      Left            =   10365
      TabIndex        =   15
      Top             =   7305
      Width           =   495
      _ExtentX        =   873
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
      Picture         =   "trFormKasir.frx":032C
   End
End
Attribute VB_Name = "trFormKasir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vaArray As New XArrayDB
Dim dbData As New ADODB.Recordset
Dim objData As New CodeSuiteLibrary.Data

Private Sub BiSAButton1_Click()
  GetRefreshVoucher
End Sub

Private Sub chkKartu_Click()
  If chkKartu.value = 1 Then
    getTotalKartu
  Else
    nTotalKartu.Default
  End If
End Sub

Private Sub getTotalKartu()
Dim nMinimalBayarPakaiKartu As Double

  nTotalKartu.value = 0
  nMinimalBayarPakaiKartu = aCfg(objData, msMinKartu)
  If nJumlahYgHarusDibayar.value - nVoucher.value >= nMinimalBayarPakaiKartu Then
    nTotalKartu.value = (nJumlahYgHarusDibayar.value - nVoucher.value - nDP.value) + ((nJumlahYgHarusDibayar.value - nVoucher.value - nDP.value) * nFee.value / 100)
  Else
    chkKartu.value = 0
    InitKartu
    MsgBox "Maaf. Minimal pembayaran dengan menggunakan kartu adalah RP " & aCfg(objData, msMinKartu), vbInformation, "Perhatian"
  End If
End Sub

Private Sub cmdCancel_Click()
  trPenjualan.lSign = 1
  Unload Me
End Sub

Private Sub cmdOK_Click()
'  If (nKembalian.Value >= 0 And nTunai.Value > 0 And nVoucher.Value <= nJumlahYgHarusDibayar.Value) or (nJumlahYgHarusDibayar.Value <= nVoucher.Value+nTunai.Value) Then
  
    If nVoucher.value + nTunai.value >= nJumlahYgHarusDibayar.value Then
      trPenjualan.lSign = 0 'sukses untuk disimpan
      'trPenjualan.nKasirTotal = nJumlahYgHarusDibayar.Value
      trPenjualan.nKasirBayar = nTunai.value
      trPenjualan.nKasirKembalian = nKembalian.value
      trPenjualan.nKasirVoucher = nVoucher.value
      Set trPenjualan.vaVoucher = vaArray
          
      If chkKartu.value = 1 Then
        If Trim(cKodeKartu.Text) <> "" And nTotalKartu.value > 0 And Trim(aCfg(objData, msRekeningFeeKartu)) <> "" Then
          If GetCekKelengkapanKartu = True Then
            trPenjualan.nKasirTotalKartu = nTotalKartu.value
            trPenjualan.nKasirKodeKartu = cKodeKartu.Text
            trPenjualan.nKasirFeeKartu = nFee.value
            trPenjualan.nKasirFeeTotalKartu = (nJumlahYgHarusDibayar.value - nVoucher.value) * (nFee.value / 100)
            trPenjualan.nKasirNoKartu = cNomorKartu.value
            trPenjualan.nKasirNoTraceKartu = cNomorTrace.value
            trPenjualan.nKasirNamaDiKartu = cNamaPemegangKartu.Text
            trPenjualan.nKasirKeteranganKartu = cNamaKartu.Text
            trPenjualan.nDPKasir = nDP.value
            Unload Me
          Else
            MsgBox "Maaf isi terlebih dahulu kelengkapan Kartu" & vbCrLf & " 1. Nomor Kartu" & vbCrLf & " 2. Nama Pemegang Kartu" & vbCrLf & " 3. No Trace", vbCritical, "Error"
          End If
        Else
          MsgBox "Maaf. pembayaran dengan kartu tidak bisa dilanjutkan. Silahkan di cek:" & vbCrLf & "1. Apakah kode Kartu sudah benar?" & vbCrLf & "2. Apakah ada total yg dibayar?" & vbCrLf & "3. Silahkan dicek konfigurasi akunting untuk Fee Kartu", vbCritical, "Error"
        End If
      End If
      
      If chkKartu.value <> 1 Then
        Unload Me
      End If
      
    Else
      MsgBox "Err. Pembayaran", vbCritical
    End If
End Sub

Private Function GetCekKelengkapanKartu() As Boolean
  If Trim(cNomorKartu.Text) = "" Or Trim(cNomorTrace.Text) = "" Or Trim(cNamaPemegangKartu.Text) = "" Then
    GetCekKelengkapanKartu = False
  Else
    'cek total pembayaran kartu
    
    GetCekKelengkapanKartu = True
  End If
End Function

Private Sub cNamaKartu_ButtonClick()
Dim db As New ADODB.Recordset

  Set db = objData.Browse(GetDSN, "kartu", , "namakartu", sisContent, cNamaKartu.Text, " and status = 1")
  If Not db.EOF Then
    cKodeKartu.Text = cNamaKartu.Browse(db)
    cNamaKartu.Text = GetNull(db!namakartu)
    nFee.value = GetNull(db!fee)
    getTotalKartu
  End If
End Sub

Private Sub Form_Activate()
  GetRefreshVoucher
End Sub

Private Sub Form_Load()
Dim n As Single

  SetIcon Me.hWnd, "SIKD"
  CenterForm Me
  TabIndex nJumlahYgHarusDibayar, n
  TabIndex nTunai, n
  TabIndex nKembalian, n
  TabIndex cmdOK, n
  InitKartu
End Sub

Private Sub InitKartu()
  chkKartu.value = 0
  cKodeKartu.Default
  cNamaKartu.Default
  cNomorKartu.Default
  cNomorTrace.Default
  cNamaPemegangKartu.Default
End Sub

Private Sub Form_Unload(Cancel As Integer)
  nJumlahYgHarusDibayar.Default
  nTunai.Default
  nKembalian.Default
  nVoucher.Default
End Sub

Private Sub nDP_Validate(Cancel As Boolean)
Dim nMinimalBayarPakaiKartu As Double

  Cancel = True
  nMinimalBayarPakaiKartu = aCfg(objData, msMinKartu)
  If nJumlahYgHarusDibayar.value - nVoucher.value - nDP.value >= nMinimalBayarPakaiKartu Then
   Cancel = False
  End If
  getTotalKartu
End Sub

Private Sub nTunai_Change()
  nKembalian.value = nTunai.value + nVoucher.value - nJumlahYgHarusDibayar.value
End Sub

Private Sub nTunai_Validate(Cancel As Boolean)
  If nTunai.value + nVoucher.value < nJumlahYgHarusDibayar.value Then
    MsgBox "Pembayaran masih kurang", vbCritical
    Cancel = True
  End If
End Sub

Private Sub GetRefreshVoucher()
Dim n As Single



  vaArray.ReDim 0, -1, 0, 6
  nTotDebet = 0
  
  Set dbData = objData.Browse(GetDSN, "membertopup", "nomormembertopup,tgl,keterangan,debet,lstatus,lstatusid", "kodeanggota", sisAssign, trPenjualan.cCustomer.Text, " and lstatus ='0' ")
  If Not dbData.EOF Then
    Do While Not dbData.EOF
      vaArray.InsertRows vaArray.UpperBound(1) + 1
      n = vaArray.UpperBound(1)
      vaArray(n, 0) = IIf(GetNull(dbData!lStatus) = sisFlag.Posting, TDBGrid1.Columns(0).BackColor = vbButtonFace, 0)
      vaArray(n, 1) = GetNull(dbData!nomormembertopup)
      vaArray(n, 2) = GetNull(dbData!tgl)
      vaArray(n, 3) = GetNull(dbData!keterangan)
      vaArray(n, 4) = GetNull(dbData!debet)
      vaArray(n, 5) = GetNull(dbData!lStatus)
      vaArray(n, 6) = GetNull(dbData!lstatusid)

      dbData.MoveNext
    Loop
  End If

  TDBGrid1.FetchRowStyle = True
  Set TDBGrid1.Array = vaArray
  TDBGrid1.ReBind
  TDBGrid1.Refresh

End Sub

Private Sub TDBGrid1_AfterColUpdate(ByVal ColIndex As Integer)
  If ColIndex = 0 Then
      TDBGrid1.Update
      GetSumTunai
    End If
  TDBGrid1.Update
  Me.Refresh
End Sub

Private Sub GetSumTunai()
Dim n As Integer
Dim nDebet As Double

  nDebet = 0
  For n = 0 To vaArray.UpperBound(1)
    If vaArray(n, 0) = -1 Then
      nDebet = nDebet + vaArray(n, 4)
    End If
  Next n
  nVoucher.value = nDebet
  nTunai.value = nJumlahYgHarusDibayar.value - nVoucher.value
  nKembalian.value = -(nJumlahYgHarusDibayar.value - nTunai.value - nVoucher.value)
  If chkKartu.value = 1 Then
    getTotalKartu
  Else
    InitKartu
  End If
End Sub
