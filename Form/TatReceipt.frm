VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0F1F1508-C40A-101B-AD04-00AA00575482}#1.0#0"; "mhrinp32.ocx"
Begin VB.Form FrmTatReceipt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tat Receipt"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8715
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4875
   ScaleWidth      =   8715
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   4870
      Left            =   15
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   8700
      _Version        =   65536
      _ExtentX        =   15346
      _ExtentY        =   8590
      _StockProps     =   77
      TintColor       =   16711935
      Alignment       =   0
      AutoSize        =   0   'False
      BevelSize       =   0
      BevelStyle      =   0
      BorderColor     =   -2147483642
      BorderStyle     =   1
      FillColor       =   -2147483633
      FontStyle       =   0
      FontTransparent =   0   'False
      LightColor      =   -2147483643
      ShadowColor     =   -2147483632
      TextColor       =   -2147483640
      WallPaper       =   0
      NoPrefix        =   0   'False
      FormatString    =   ""
      Caption         =   ""
      Picture         =   "TatReceipt.frx":0000
      Begin TabDlg.SSTab SSTab1 
         Height          =   4630
         Left            =   120
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   120
         Width           =   8460
         _ExtentX        =   14923
         _ExtentY        =   8176
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabHeight       =   520
         ShowFocusRect   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "&List"
         TabPicture(0)   =   "TatReceipt.frx":001C
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "DataGrid1"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Text1"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "&Details"
         TabPicture(1)   =   "TatReceipt.frx":0038
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Mh3dFrame2"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   600
            TabIndex        =   10
            Top             =   4160
            Width           =   7760
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   3640
            Left            =   120
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   450
            Width           =   8235
            _ExtentX        =   14526
            _ExtentY        =   6429
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            BackColor       =   16776960
            HeadLines       =   1
            RowHeight       =   18
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   3
            BeginProperty Column00 
               DataField       =   "Name"
               Caption         =   "Voucher No."
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "Date"
               Caption         =   "Voucher Date"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "dd-MM-yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2057
                  SubFormatType   =   3
               EndProperty
            EndProperty
            BeginProperty Column02 
               DataField       =   "Particulars"
               Caption         =   "Particulars"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2057
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               MarqueeStyle    =   3
               ScrollBars      =   3
               AllowRowSizing  =   0   'False
               AllowSizing     =   0   'False
               Locked          =   -1  'True
               BeginProperty Column00 
                  Alignment       =   1
                  Locked          =   -1  'True
                  ColumnWidth     =   1019.906
               EndProperty
               BeginProperty Column01 
                  Locked          =   -1  'True
                  ColumnWidth     =   1170.142
               EndProperty
               BeginProperty Column02 
                  Locked          =   -1  'True
                  ColumnWidth     =   5474.835
               EndProperty
            EndProperty
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   3670
            Left            =   -74880
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   480
            Width           =   8235
            _Version        =   65536
            _ExtentX        =   14526
            _ExtentY        =   6473
            _StockProps     =   77
            Enabled         =   0   'False
            TintColor       =   16711935
            Alignment       =   0
            AutoSize        =   0   'False
            BevelSize       =   0
            BevelStyle      =   0
            BorderColor     =   -2147483642
            BorderStyle     =   1
            FillColor       =   -2147483633
            FontStyle       =   0
            FontTransparent =   0   'False
            LightColor      =   -2147483643
            ShadowColor     =   -2147483632
            TextColor       =   -2147483640
            WallPaper       =   0
            NoPrefix        =   0   'False
            FormatString    =   ""
            Caption         =   ""
            Picture         =   "TatReceipt.frx":0054
            Begin VB.TextBox MhRealInput1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
               Height          =   330
               Left            =   6900
               MaxLength       =   13
               TabIndex        =   18
               Text            =   "0"
               Top             =   1400
               Visible         =   0   'False
               Width           =   1005
            End
            Begin MhinrelLib.MhRealInput MhRealInput5 
               Height          =   265
               Left            =   6900
               TabIndex        =   16
               TabStop         =   0   'False
               Top             =   3290
               Width           =   1005
               _Version        =   65536
               _ExtentX        =   1773
               _ExtentY        =   467
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Comic Sans MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               FillColor       =   16777215
               MaxReal         =   9999999
               MinReal         =   0
               ReadOnly        =   -1  'True
               SpinChangeReal  =   0
               CaretColor      =   -2147483642
               DecimalPlaces   =   0
               VAlignment      =   2
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel19 
               Height          =   265
               Left            =   120
               TabIndex        =   17
               Top             =   3290
               Width           =   8010
               _Version        =   65536
               _ExtentX        =   14129
               _ExtentY        =   467
               _StockProps     =   77
               BackColor       =   32896
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Comic Sans MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   0   'False
               TintColor       =   16711935
               Caption         =   ""
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "TatReceipt.frx":0070
               Picture         =   "TatReceipt.frx":008C
            End
            Begin VB.TextBox Text5 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   435
               MaxLength       =   40
               TabIndex        =   5
               Top             =   1400
               Visible         =   0   'False
               Width           =   6480
            End
            Begin VB.TextBox Text2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1560
               MaxLength       =   10
               TabIndex        =   0
               Top             =   105
               Width           =   1530
            End
            Begin VB.TextBox Text4 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1560
               MaxLength       =   40
               TabIndex        =   2
               Top             =   630
               Width           =   3930
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
               Height          =   330
               Left            =   120
               TabIndex        =   13
               Top             =   105
               Width           =   1455
               _Version        =   65536
               _ExtentX        =   2566
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Voucher No."
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "TatReceipt.frx":00A8
               Picture         =   "TatReceipt.frx":00C4
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Left            =   5475
               TabIndex        =   14
               Top             =   105
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Voucher Date"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "TatReceipt.frx":00E0
               Picture         =   "TatReceipt.frx":00FC
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
               Height          =   330
               Left            =   120
               TabIndex        =   15
               Top             =   630
               Width           =   1455
               _Version        =   65536
               _ExtentX        =   2566
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Remarks"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "TatReceipt.frx":0118
               Picture         =   "TatReceipt.frx":0134
            End
            Begin MSDataGridLib.DataGrid DataGrid2 
               Height          =   2400
               Left            =   120
               TabIndex        =   4
               Top             =   1150
               Width           =   8010
               _ExtentX        =   14129
               _ExtentY        =   4233
               _Version        =   393216
               AllowUpdate     =   0   'False
               AllowArrows     =   -1  'True
               Appearance      =   0
               BackColor       =   16776960
               HeadLines       =   1
               RowHeight       =   20
               TabAction       =   2
               FormatLocked    =   -1  'True
               BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ColumnCount     =   2
               BeginProperty Column00 
                  DataField       =   "PrinterName"
                  Caption         =   "Printer Name"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column01 
                  DataField       =   "Quantity"
                  Caption         =   "   Quantity"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   1
                     Format          =   "0"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2057
                     SubFormatType   =   1
                  EndProperty
               EndProperty
               SplitCount      =   1
               BeginProperty Split0 
                  MarqueeStyle    =   3
                  ScrollBars      =   3
                  AllowRowSizing  =   0   'False
                  AllowSizing     =   0   'False
                  Locked          =   -1  'True
                  BeginProperty Column00 
                     ColumnAllowSizing=   -1  'True
                     Locked          =   -1  'True
                     ColumnWidth     =   6465.26
                  EndProperty
                  BeginProperty Column01 
                     Alignment       =   1
                     ColumnAllowSizing=   -1  'True
                     Locked          =   -1  'True
                     ColumnWidth     =   975.118
                  EndProperty
               EndProperty
            End
            Begin TDBDate6Ctl.TDBDate MhDateInput1 
               Height          =   330
               Left            =   7035
               TabIndex        =   1
               Top             =   105
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
               _ExtentY        =   582
               Calendar        =   "TatReceipt.frx":0150
               Caption         =   "TatReceipt.frx":0268
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "TatReceipt.frx":02D4
               Keys            =   "TatReceipt.frx":02F2
               Spin            =   "TatReceipt.frx":0350
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               CursorPosition  =   0
               DataProperty    =   0
               DisplayFormat   =   "dd-mm-yyyy"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               FirstMonth      =   1
               ForeColor       =   -2147483640
               Format          =   "dd-mm-yyyy"
               HighlightText   =   0
               IMEMode         =   3
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxDate         =   2958465
               MinDate         =   -657434
               MousePointer    =   0
               MoveOnLRKey     =   0
               OLEDragMode     =   0
               OLEDropMode     =   0
               PromptChar      =   " "
               ReadOnly        =   0
               ShowContextMenu =   1
               ShowLiterals    =   0
               TabAction       =   0
               Text            =   "  -  -    "
               ValidateMode    =   0
               ValueVT         =   1
               Value           =   39849
               CenturyMode     =   0
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Left            =   5475
               TabIndex        =   19
               Top             =   630
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               Caption         =   " Rate"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "TatReceipt.frx":0378
               Picture         =   "TatReceipt.frx":0394
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput2 
               Height          =   330
               Left            =   7035
               TabIndex        =   3
               Top             =   630
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
               _ExtentY        =   582
               Calculator      =   "TatReceipt.frx":03B0
               Caption         =   "TatReceipt.frx":03D0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "TatReceipt.frx":043C
               Keys            =   "TatReceipt.frx":045A
               Spin            =   "TatReceipt.frx":04A4
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "#########0.00"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "#########0.00"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   9999999999.99
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   2088828933
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin VB.Line Line1 
               X1              =   0
               X2              =   8280
               Y1              =   525
               Y2              =   525
            End
            Begin VB.Line Line2 
               X1              =   0
               X2              =   8280
               Y1              =   1050
               Y2              =   1050
            End
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00808000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Find"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   330
            Left            =   120
            TabIndex        =   11
            Top             =   4160
            Width           =   495
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   18
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Add"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Edit"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Delete"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Save"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Cancel"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Refresh"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Filter"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Print"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Print Preview"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Mail"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "First"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Previous"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Next"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Last"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exit"
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   4
      Left            =   2760
      Top             =   2280
   End
End
Attribute VB_Name = "FrmTatReceipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CxnTatReceipt As New ADODB.Connection
Dim rstCompanyMaster As New ADODB.Recordset
Dim rstTatRVList As New ADODB.Recordset
Dim rstTatRVParent As New ADODB.Recordset
Dim WithEvents rstTatRVChild As ADODB.Recordset
Attribute rstTatRVChild.VB_VarHelpID = -1
Dim rstPrinterList As New ADODB.Recordset
Dim PrinterCode As String
Dim PrevStr As String
Dim dblBookMark As Double
Dim blnRecordExist As Boolean
Dim BalanceQuantity As Long
Dim OutputTo As String
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    
    CenterForm Me
    BusySystemIndicator True
    CxnTatReceipt.CursorLocation = adUseClient
    CxnTatReceipt.Open CxnDatabase.ConnectionString
    rstCompanyMaster.Open "Select PrintName, Address1, Address2, Address3, Address4, Phone, Fax, EMail, Website From CompanyMaster", CxnTatReceipt, adOpenKeyset, adLockReadOnly
    rstPrinterList.Open "Select Name As Col0, Code From AccountMaster Where Type In ('05','06') Order by Name", CxnTatReceipt, adOpenKeyset, adLockReadOnly
    rstTatRVList.Open "Select TatRVParent.Code, TatRVParent.Name, Date, Particulars From TatRVParent Order By TatRVParent.Name", CxnTatReceipt, adOpenKeyset, adLockOptimistic
    rstTatRVParent.CursorLocation = adUseClient
    Set rstTatRVChild = New ADODB.Recordset
    If rstTatRVList.RecordCount > 0 Then rstTatRVList.MoveLast
    Set DataGrid1.DataSource = rstTatRVList
    BusySystemIndicator False
    SSTab1.Tab = 0
    If Not (rstTatRVList.EOF Or rstTatRVList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    rstTatRVList.ActiveConnection = Nothing
    rstPrinterList.ActiveConnection = Nothing
    SetButtonsForNoRecord
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    Call CloseForm(FrmTatReceipt)
End Sub
Private Sub Form_Activate()
    EnableChildMenu True
    Text1.SetFocus
End Sub
Private Sub Form_Deactivate()
    DisableChildMenu
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyEscape Then
        If SSTab1.Tab = 0 Then
            Call CloseForm(FrmTatReceipt)
        Else
            If Toolbar1.Buttons.Item(1).Enabled Then
                SSTab1.Tab = 0
            Else
                If Me.ActiveControl.Name <> "Text5" And Me.ActiveControl.Name <> "MhRealInput1" Then
                    If MsgBox("Are you sure to Quit?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Quit !") <> vbYes Then
                        Me.ActiveControl.SetFocus
                    Else
                        Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
                    End If
                End If
            End If
            If Me.ActiveControl.Name <> "Text5" And Me.ActiveControl.Name <> "MhRealInput1" Then
                KeyCode = 0
            End If
        End If
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyA And Toolbar1.Buttons.Item(1).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(1)
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyE And Toolbar1.Buttons.Item(2).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(2)
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyD And Toolbar1.Buttons.Item(3).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(3)
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyS And Toolbar1.Buttons.Item(4).Enabled Then
        If Me.ActiveControl.Name <> "Text5" And Me.ActiveControl.Name <> "MhRealInput1" Then
            Toolbar1_ButtonClick Toolbar1.Buttons.Item(4)
        End If
        KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyF5 And Toolbar1.Buttons.Item(6).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
        KeyCode = 0
    ElseIf Shift = vbAltMask And KeyCode = vbKeyP And Toolbar1.Buttons.Item(1).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(9)
        KeyCode = 0
    ElseIf Shift = vbAltMask And KeyCode = vbKeyV And Toolbar1.Buttons.Item(1).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(10)
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyF And Toolbar1.Buttons.Item(1).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(13)
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyP And Toolbar1.Buttons.Item(1).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(14)
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyN And Toolbar1.Buttons.Item(1).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(15)
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyL And Toolbar1.Buttons.Item(1).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(16)
        KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyReturn Then
        If Toolbar1.Buttons.Item(1).Enabled Then
            SSTab1.Tab = 1
            SSTab1.SetFocus
        Else
            If Me.ActiveControl.Name <> "MhRealInput1" Then
                SendKeys "{TAB}"
            End If
        End If
        If Me.ActiveControl.Name <> "MhRealInput1" Then
            KeyCode = 0
        End If
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Toolbar1.Buttons.Item(4).Enabled Then
        Call Form_KeyDown(vbKeyEscape, 0)
        Cancel = 1
    Else
        Call CloseForm(FrmTatReceipt)
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CloseRecordset(rstCompanyMaster)
    Call CloseRecordset(rstTatRVList)
    Call CloseRecordset(rstTatRVParent)
    Call CloseRecordset(rstTatRVChild)
    Call CloseRecordset(rstPrinterList)
    Call CloseConnection(CxnTatReceipt)
    ShowProgressInStatusBar False
    DisableChildMenu
End Sub
Private Sub Text1_Change()
    If rstTatRVList.RecordCount = 0 Then Exit Sub
    rstTatRVList.MoveFirst
    If Text1.Text <> "" Then
        rstTatRVList.Find "[Name] Like '%" & FixQuote(Text1.Text) & "%'"
        If rstTatRVList.EOF Then
            rstTatRVList.MoveFirst
            If PrevStr <> "" And Len(Text1.Text) > 1 Then
                If dblBookMark <> 0 Then
                    rstTatRVList.Bookmark = dblBookMark
                End If
            Else
                PrevStr = ""
            End If
            Beep
            DisplayError ("Spelling Error")
            Text1.Text = PrevStr
            SendKeys "{End}"
        Else
            PrevStr = Text1.Text
            dblBookMark = DataGrid1.Bookmark
        End If
    Else
        PrevStr = ""
    End If
    If Not (rstTatRVList.EOF Or rstTatRVList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim KeyProcessed As Boolean
    
    If rstTatRVList.RecordCount = 0 Then Exit Sub
    If Shift = 0 And KeyCode = vbKeyUp Then
        With rstTatRVList
            .MovePrevious
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyBack Then
        With rstTatRVList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyDown Then
        With rstTatRVList
            .MoveNext
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageUp Then
        With rstTatRVList
            .Move (-1) * (DataGrid1.VisibleRows - 1)
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageUp Then
        With rstTatRVList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageDown Then
        With rstTatRVList
            .Move DataGrid1.VisibleRows - 1
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageDown Then
        With rstTatRVList
            .MoveLast
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    End If
    If KeyProcessed Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
        KeyProcessed = False
        KeyCode = 0
    End If
End Sub
Private Sub SSTab1_Click(PreviousTab As Integer)
    If Toolbar1.Buttons.Item(1).Enabled Then
        If SSTab1.Tab = 1 Then
            ViewRecord
        Else
            If Not (rstTatRVList.EOF Or rstTatRVList.BOF) Then
                With DataGrid1.SelBookmarks
                    If .Count <> 0 Then .Remove 0
                    .Add DataGrid1.Bookmark
                End With
            End If
            Text1.SetFocus
        End If
        SSTab1.TabEnabled(0) = True
    Else
        SSTab1.TabEnabled(0) = False
        Text2.SetFocus
    End If
End Sub
Public Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim HiLiteRecord As Boolean
    Dim UpdateFlag As Integer
    
    If Button.Index = 1 Then
        If rstTatRVParent.State = adStateOpen Then
           rstTatRVParent.Close
        End If
        rstTatRVParent.Open "Select * From TatRVParent Where Code = ''", CxnTatReceipt, adOpenKeyset, adLockOptimistic
        ClearFields ("P")
        ClearFields ("C")
        Call LoadPrinterList("")
        If rstTatRVChild.State = adStateClosed Then
            SSTab1.Tab = 0
            Exit Sub
        End If
        If AddRecord(rstTatRVParent) Then
            Text2.Text = GenerateCode(CxnTatReceipt, "Select Max(Val(Name)) From TatRVParent", 10, Space(1))
            MhDateInput1.Text = Format(Date, "dd-MM-yyyy")
            Call SetButtons(False)
            SSTab1.Tab = 1
            Text2.SetFocus
            blnRecordExist = False
            CxnTatReceipt.BeginTrans
        End If
    ElseIf Button.Index = 2 Then
        If rstTatRVList.RecordCount = 0 Then Exit Sub
        SSTab1.Tab = 1
        EditRecord
    ElseIf Button.Index = 3 Then
        If rstTatRVList.RecordCount = 0 Then Exit Sub
        If AllowTransactionsDeletion = 0 Then
            Call DisplayError("You don't have the rights to Delete this Voucher")
            Exit Sub
        End If
        SSTab1.Tab = 1
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            On Error Resume Next
            MdiMainMenu.MousePointer = vbHourglass
            CxnTatReceipt.Execute "Delete From TatRVParent Where Code = '" & rstTatRVList.Fields("Code").Value & "'"
            MdiMainMenu.MousePointer = vbNormal
            If Err.Number = 0 Then
                rstTatRVList.Delete
                rstTatRVList.MoveNext
                If rstTatRVList.RecordCount > 0 And rstTatRVList.EOF Then
                    rstTatRVList.MoveLast
                End If
                ShowProgressInStatusBar True
                Timer1.Enabled = True
            Else
                DisplayError ("Failed to delete the record")
            End If
            On Error GoTo 0
        End If
        SetButtons (True)
        SetButtonsForNoRecord
        SSTab1.Tab = 0
        HiLiteRecord = True
    ElseIf Button.Index = 4 Then
        If CheckMandatoryFields Then Exit Sub
        MakeTextBoxInvisible (False)
        If blnRecordExist And AllowTransactionsModification = 0 Then
            Call DisplayError("You don't have the rights to Edit this Voucher")
            Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
            Exit Sub
        End If
        SaveFields
        UpdateFlag = 0
        If UpdateRecord(rstTatRVParent) Then
            If UpdatePrinterList("D") Then
                 UpdateFlag = 1
                 If rstTatRVChild.RecordCount <> 0 Then
                      rstTatRVChild.MoveFirst
                      Do While Not rstTatRVChild.EOF
                          If Val(rstTatRVChild.Fields("Quantity").Value) <> 0 Then
                               If Not UpdatePrinterList("U") Then
                                    UpdateFlag = 0
                                    Exit Do
                                End If
                          End If
                          rstTatRVChild.MoveNext
                      Loop
                 End If
            End If
        End If
        If UpdateFlag Then
            AddToList
            CxnTatReceipt.CommitTrans
            If rstTatRVParent.State = adStateOpen Then
                rstTatRVParent.Close
            End If
            rstTatRVParent.CursorLocation = adUseClient
            Call SetButtons(True)
            SSTab1.Tab = 0
            ShowProgressInStatusBar True
            Timer1.Enabled = True
        Else
            DisplayError ("Failed to save the record")
            Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
        End If
    ElseIf Button.Index = 5 Then
        If CancelRecordUpdate(rstTatRVParent) Then
            CxnTatReceipt.RollbackTrans
            If rstTatRVParent.State = adStateOpen Then
                rstTatRVParent.Close
            End If
            rstTatRVParent.CursorLocation = adUseClient
            Call SetButtons(True)
            SetButtonsForNoRecord
            SSTab1.Tab = 0
        End If
    ElseIf Button.Index = 6 Then
        SSTab1.Tab = 0
        Set DataGrid1.DataSource = Nothing
        rstTatRVList.ActiveConnection = CxnTatReceipt
        Do While Not RefreshRecord(rstTatRVList)
        Loop
        Set DataGrid1.DataSource = rstTatRVList
        rstTatRVList.ActiveConnection = Nothing
        If rstTatRVList.RecordCount > 0 Then rstTatRVList.MoveLast
        rstPrinterList.ActiveConnection = CxnTatReceipt
        Do While Not RefreshRecord(rstPrinterList)
        Loop
        rstPrinterList.ActiveConnection = Nothing
        HiLiteRecord = True
    ElseIf Button.Index = 7 Then
        SSTab1.Tab = 0
        HiLiteRecord = True
    ElseIf Button.Index = 9 Then
        If rstTatRVList.RecordCount = 0 Then Exit Sub
        OutputTo = "P"
        PrintTatReceipt
        HiLiteRecord = True
    ElseIf Button.Index = 10 Then
        If rstTatRVList.RecordCount = 0 Then Exit Sub
        OutputTo = "S"
        PrintTatReceipt
        HiLiteRecord = True
    ElseIf Button.Index = 13 Then
        If rstTatRVList.RecordCount > 0 Then rstTatRVList.MoveFirst
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 14 Then
        If rstTatRVList.RecordCount > 0 Then
            rstTatRVList.MovePrevious
            If rstTatRVList.BOF Then
                rstTatRVList.MoveNext
            End If
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 15 Then
        If rstTatRVList.RecordCount > 0 Then
            rstTatRVList.MoveNext
            If rstTatRVList.EOF Then
                rstTatRVList.MovePrevious
            End If
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 16 Then
        If rstTatRVList.RecordCount > 0 Then rstTatRVList.MoveLast
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 18 Then
        Call CloseForm(FrmTatReceipt)
        HiLiteRecord = False
    End If
    If HiLiteRecord Then
        If Not (rstTatRVList.EOF Or rstTatRVList.BOF) Then
            With DataGrid1.SelBookmarks
                If .Count <> 0 Then .Remove 0
                .Add DataGrid1.Bookmark
            End With
        End If
        Text1.SetFocus
    End If
End Sub
Private Sub DataGrid1_DblClick()
    If Toolbar1.Buttons.Item(2).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(2)
    End If
End Sub
Private Sub SetButtons(bVal As Boolean)
    Toolbar1.Buttons.Item(1).Enabled = bVal
    Toolbar1.Buttons.Item(2).Enabled = bVal
    Toolbar1.Buttons.Item(3).Enabled = bVal
    Toolbar1.Buttons.Item(4).Enabled = Not bVal
    Toolbar1.Buttons.Item(5).Enabled = Not bVal
    Toolbar1.Buttons.Item(6).Enabled = bVal
    Toolbar1.Buttons.Item(7).Enabled = bVal
    Toolbar1.Buttons.Item(9).Enabled = bVal
    Toolbar1.Buttons.Item(10).Enabled = bVal
    Toolbar1.Buttons.Item(13).Enabled = bVal
    Toolbar1.Buttons.Item(14).Enabled = bVal
    Toolbar1.Buttons.Item(15).Enabled = bVal
    Toolbar1.Buttons.Item(16).Enabled = bVal
    Toolbar1.Buttons.Item(18).Enabled = bVal
    Mh3dFrame2.Enabled = Not bVal
End Sub
Private Sub SetButtonsForNoRecord()
    If rstTatRVList.RecordCount = 0 Then
        Toolbar1.Buttons.Item(2).Enabled = False
        Toolbar1.Buttons.Item(3).Enabled = False
        Toolbar1.Buttons.Item(9).Enabled = False
        Toolbar1.Buttons.Item(10).Enabled = False
        Toolbar1.Buttons.Item(13).Enabled = False
        Toolbar1.Buttons.Item(14).Enabled = False
        Toolbar1.Buttons.Item(15).Enabled = False
        Toolbar1.Buttons.Item(16).Enabled = False
    End If
End Sub
Private Sub Text2_Validate(Cancel As Boolean)
    If rstTatRVParent.EOF Or rstTatRVParent.BOF Then Exit Sub
    If CheckEmpty(Text2, True) Then
        Cancel = True
    ElseIf CheckDuplicate(CxnTatReceipt, "TatRVParent", "Code", "[Name]", Trim(Text2.Text), rstTatRVParent.Fields("Code").Value, False) Then
        Cancel = True
    End If
End Sub
Private Sub MhDateInput1_Validate(Cancel As Boolean)
    If Not IsDate(GetDate(MhDateInput1.Text)) Then
        Cancel = True
    ElseIf Format(GetDate(MhDateInput1.Text), "yyyymmdd") < Format(FinancialYearFrom, "yyyymmdd") Or Format(GetDate(MhDateInput1.Text), "yyyymmdd") > Format(FinancialYearTo, "yyyymmdd") Then
        Cancel = True
    End If
End Sub
Private Sub Text4_Validate(Cancel As Boolean)
    If rstTatRVChild.RecordCount = 0 Then
        SendKeys "^"
        Call AddRecord(rstTatRVChild)
        Call ClearFields("C")
        Call DataGrid2_KeyDown(vbKeyE, vbCtrlMask)
    End If
End Sub
Private Sub ViewRecord()
    ClearFields ("P")
    ClearFields ("C")
    If rstTatRVList.EOF Then
        If rstTatRVChild.State = adStateOpen Then
            rstTatRVChild.Close
        End If
        Exit Sub
    End If
    FindRecord
    LoadFields
End Sub
Private Sub FindRecord()
    If rstTatRVParent.State = adStateOpen Then
       rstTatRVParent.Close
    End If
    rstTatRVParent.Open "Select * From TatRVParent Where Code = '" & FixQuote(rstTatRVList.Fields("Code").Value) & "'", CxnTatReceipt, adOpenKeyset, adLockOptimistic
    If rstTatRVParent.RecordCount = 0 Then
       Call DisplayError("This Record has been deleted by Another User ! Click Ok To Refresh the Recordset")
       Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
    End If
End Sub
Private Sub ClearFields(ByVal strType As String)
    If strType = "P" Then
        Text2.Text = ""
        Text4.Text = ""
        MhDateInput1.Text = Format(Date, "dd-MM-yyyy")
        MhRealInput5.Text = 0#
        MhRealInput2.Text = "0.00"
    ElseIf strType = "C" Then
        Text5.Text = ""
        MhRealInput1.Text = "0"
    End If
End Sub
Private Sub LoadFields()
    If rstTatRVParent.EOF Or rstTatRVParent.BOF Then Exit Sub
    Text2.Text = rstTatRVParent.Fields("Name").Value
    MhDateInput1.Text = Format(rstTatRVParent.Fields("Date").Value, "dd-MM-yyyy")
    MhRealInput2.Text = Format(rstTatRVParent.Fields("Rate").Value, "0.00")
    Text4.Text = rstTatRVParent.Fields("Remarks").Value
    Call LoadPrinterList(rstTatRVParent.Fields("Code").Value)
    If rstTatRVChild.State = adStateOpen Then
        CalculateTotal
    End If
End Sub
Private Sub EditRecord()
    On Error GoTo ErrorHandler
    
    If rstTatRVParent.RecordCount = 0 Then Exit Sub
    If rstTatRVChild.State = adStateClosed Then
        SSTab1.Tab = 0
        Exit Sub
    End If
    If rstTatRVParent.State = adStateOpen Then
       rstTatRVParent.Close
    End If
    rstTatRVParent.CursorLocation = adUseServer
    rstTatRVParent.Open "Select * From TatRVParent Where Code = '" & FixQuote(rstTatRVList.Fields("Code").Value) & "'", CxnTatReceipt, adOpenKeyset, adLockPessimistic
    MdiMainMenu.MousePointer = vbHourglass
    rstTatRVParent.Fields("Printstatus") = "N"
    MdiMainMenu.MousePointer = vbNormal
    AddToList
    Call SetButtons(False)
    SSTab1.TabEnabled(0) = False
    Text2.SetFocus
    blnRecordExist = True
    CxnTatReceipt.BeginTrans
    Exit Sub
ErrorHandler:
    If Err.Number = -2147467259 Then
       Call DisplayError("Failed to Edit the record")
    End If
    MdiMainMenu.MousePointer = vbNormal
    SSTab1.Tab = 0
End Sub
Private Sub SaveFields()
    If rstTatRVParent.EOF Or rstTatRVParent.BOF Then Exit Sub
    If Not blnRecordExist Then
        rstTatRVParent.Fields("Code").Value = GenerateCode(CxnTatReceipt, "Select Max(Code) From TatRVParent", 6, "0")
        rstTatRVParent.Fields("CreatedBy").Value = UserCode
        rstTatRVParent.Fields("CreatedOn").Value = Now()
        rstTatRVParent.Fields("Recordstatus").Value = "N"
    Else
        rstTatRVParent.Fields("ModifiedBy").Value = UserCode
        rstTatRVParent.Fields("ModifiedOn").Value = Now()
        rstTatRVParent.Fields("Recordstatus").Value = "M"
    End If
    rstTatRVParent.Fields("Name").Value = Pad(Trim(Text2.Text), Space(1), 10, "L")
    rstTatRVParent.Fields("Date").Value = GetDate(MhDateInput1.Text)
    rstTatRVParent.Fields("Particulars").Value = "Received " & Format(Val(MhRealInput5.Text), "0") & " Tats From " & Format(rstTatRVChild.RecordCount, 0) & " Printers"
    rstTatRVParent.Fields("Rate").Value = Val(MhRealInput2.Text)
    rstTatRVParent.Fields("Remarks").Value = Trim(Text4.Text)
    rstTatRVParent.Fields("PrintStatus").Value = "N"
End Sub
Private Sub AddToList()
    On Error Resume Next
    
    rstTatRVList.MoveFirst
    rstTatRVList.Find "[Code] = '" & rstTatRVParent.Fields("Code").Value & "'"
    If rstTatRVList.EOF Then
       rstTatRVList.AddNew
       rstTatRVList.Fields("Code").Value = rstTatRVParent.Fields("Code").Value
    End If
    rstTatRVList.Fields("Name").Value = Pad(rstTatRVParent.Fields("Name").Value, Space(1), 10, "L")
    rstTatRVList.Fields("Date").Value = rstTatRVParent.Fields("Date").Value
    rstTatRVList.Fields("Particulars").Value = Trim(rstTatRVParent.Fields("Particulars").Value)
    rstTatRVList.Update
    rstTatRVList.Sort = "Name Asc"
    rstTatRVList.Find "[Code] = '" & rstTatRVParent.Fields("Code").Value & "'"
End Sub
Private Function CheckMandatoryFields() As Boolean
    If CheckEmpty(Text2.Text, False) Then
       DisplayError ("Voucher No. cannot be blank")
       Text2.SetFocus
       CheckMandatoryFields = True
    ElseIf CheckDuplicate(CxnTatReceipt, "TatRVParent", "Code", "[Name]", Trim(Text2.Text), rstTatRVParent.Fields("Code").Value, False) Then
        Text2.SetFocus
        CheckMandatoryFields = True
    End If
End Function
Private Sub Timer1_Timer()
    On Error Resume Next
    
    MdiMainMenu.ProgressBar1.Value = MdiMainMenu.ProgressBar1.Value + 10
    If MdiMainMenu.ProgressBar1.Value = 100 Then
       Timer1.Enabled = False
       ShowProgressInStatusBar False
    End If
End Sub
Private Sub LoadPrinterList(ByVal strVoucherCode As String)
    On Error GoTo ErrorHandler
    
    If rstTatRVChild.State = adStateOpen Then
       rstTatRVChild.Close
    End If
    rstTatRVChild.Open "Select Printer, AccountMaster.Name As PrinterName, Quantity From AccountMaster, TatRVChild Where TatRVChild.Printer = AccountMaster.Code And TatRVChild.Code = '" & strVoucherCode & "' Order by AccountMaster.Name", CxnTatReceipt, adOpenKeyset, adLockOptimistic
    rstTatRVChild.ActiveConnection = Nothing
    Set DataGrid2.DataSource = rstTatRVChild
    Exit Sub
ErrorHandler:
    DisplayError ("Failed to Load Printer List")
End Sub
Private Sub DataGrid2_DblClick()
    Call DataGrid2_KeyDown(vbKeyE, vbCtrlMask)
End Sub
Private Sub DataGrid2_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask And KeyCode = vbKeyE Then
        If rstTatRVChild.RecordCount = 0 Then
            KeyCode = 0
            Exit Sub
        End If
        If Val(CheckNull(rstTatRVChild.Fields("Quantity").Value)) <> 0 Then
            PrinterCode = rstTatRVChild.Fields("Printer").Value
            Text5.Text = rstTatRVChild.Fields("PrinterName").Value
            MhRealInput1.Text = Format(Val(rstTatRVChild.Fields("Quantity").Value), "0")
        End If
        With DataGrid2
            Text5.Visible = True
            Text5.Move .Left + .Columns(0).Left, .Top + .RowTop(.Row), .Columns(0).Width + 10, .RowHeight + 30
            MhRealInput1.Visible = True
            MhRealInput1.Move .Left + .Columns(1).Left, .Top + .RowTop(.Row), .Columns(1).Width + 10, .RowHeight + 30
        End With
        DataGrid2.Enabled = False
        Text5.SetFocus
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyA Then
        SendKeys "^"
        Call AddRecord(rstTatRVChild)
        Call ClearFields("C")
        Call DataGrid2_KeyDown(vbKeyE, vbCtrlMask)
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyD Then
        If rstTatRVChild.RecordCount = 0 Then Exit Sub
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            Set DataGrid2.DataSource = Nothing
            rstTatRVChild.Delete
            rstTatRVChild.MoveNext
            Set DataGrid2.DataSource = rstTatRVChild
            CalculateTotal
            DataGrid2.SetFocus
        End If
        If rstTatRVChild.RecordCount = 0 Then
            Call ClearFields("C")
        End If
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyS And Toolbar1.Buttons.Item(4).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(4)
    ElseIf Shift = vbShiftMask And KeyCode = vbKeyTab Then
       MhRealInput2.SetFocus
    ElseIf Shift = 0 And KeyCode = vbKeyReturn Then
        Text2.SetFocus
        KeyCode = 0
    End If
End Sub
Private Sub DataGrid2_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim menusel As String
    
    If Button = vbRightButton Then
       menusel = DisplayPopupMenu(Me.hwnd)
        Select Case menusel
            Case 1
                Call DataGrid2_KeyDown(vbKeyA, vbCtrlMask)
            Case 2
                Call DataGrid2_KeyDown(vbKeyE, vbCtrlMask)
            Case 3
                Call DataGrid2_KeyDown(vbKeyD, vbCtrlMask)
            Case Else
        End Select
    End If
End Sub
Private Sub Text5_Change()
    If Text5.Text = " " Then
        Text5.Text = "?"
        SendKeys "{TAB}"
    End If
End Sub
Private Sub Text5_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyEscape Then
        MakeTextBoxInvisible (True)
    End If
End Sub
Private Sub Text5_Validate(Cancel As Boolean)
    Dim SearchString As String
    
    SearchString = FixQuote(Text5.Text)
    If rstPrinterList.RecordCount = 0 Then
        DisplayError ("No Record in Printer Master")
        Cancel = True
        Exit Sub
    Else
        rstPrinterList.MoveFirst
    End If
    rstPrinterList.Find "[Col0] = '" & RTrim(SearchString) & "'"
    If rstPrinterList.EOF Then
        SelectionType = "S"
        PrinterCode = ""
        Call LoadSelectionList(rstPrinterList, "List of Printers...", "Name")
        SearchOrder = 0
        Call DisplaySelectionList(Text5, PrinterCode)
        Call CloseForm(FrmSelectionList)
        If CheckEmpty(Text5.Text, False) Then
            Text5.Text = "?"
        End If
        If RTrim(PrinterCode) <> "" Then
            SendKeys "{TAB}"
        End If
        Cancel = True
        Exit Sub
    ElseIf (rstTatRVChild.Fields("PrinterName").Value <> Text5.Text) Or (CheckEmpty(rstTatRVChild.Fields("PrinterName").Value, False)) Then
        If CheckDuplicatePrinter Then
            Call DisplayError("Duplicate Entry")
            Text5.SelStart = 0
            Text5.SelLength = Len(Text5.Text)
            Cancel = True
            Exit Sub
        End If
    End If
    PrinterCode = rstPrinterList.Fields("Code").Value
    BalanceQuantity = CalculateTatBalance(PrinterCode, CheckNull(rstTatRVParent.Fields("Code").Value))
    MsgBox "Quantity at Press : " & Format(Str(BalanceQuantity), "#0")
End Sub
Private Sub MhRealInput1_GotFocus()
    FocusSelect Me.ActiveControl
End Sub
Private Sub MhRealInput1_KeyPress(KeyAscii As Integer)
    ValidateKey MhRealInput1, KeyAscii, 0
End Sub
Private Sub MhRealInput1_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyReturn Then
        If Not ValidateNumber(Me.ActiveControl, 0) Then Exit Sub
        If Val(MhRealInput1.Text) > 0 And Val(MhRealInput1.Text) <= BalanceQuantity Then
            rstTatRVChild.Fields("Printer").Value = PrinterCode
            rstTatRVChild.Fields("PrinterName").Value = Trim(Text5.Text)
            rstTatRVChild.Fields("Quantity").Value = Format(Val(MhRealInput1.Text), "0")
            rstTatRVChild.Update
            MakeTextBoxInvisible (False)
            CalculateTotal
            If rstTatRVChild.AbsolutePosition = rstTatRVChild.RecordCount Then
                Call DataGrid2_KeyDown(vbKeyA, vbCtrlMask)
            End If
        Else
            If Val(MhRealInput1.Text) > 0 Then
                Call DisplayError("Quantity cann't be greater than " & Format(BalanceQuantity, 0))
                MhRealInput1.SetFocus
                FocusSelect Me.ActiveControl
            End If
        End If
    ElseIf Shift = 0 And KeyCode = vbKeyEscape Then
       MakeTextBoxInvisible (True)
    End If
End Sub
Private Sub MhRealInput1_Validate(Cancel As Boolean)
    Cancel = True
End Sub
Private Sub MakeTextBoxInvisible(ByVal KeyEscPressed As Boolean)
    If KeyEscPressed Then
        If Not (rstTatRVChild.EOF Or rstTatRVChild.BOF) Then
            If Val(CheckNull(rstTatRVChild.Fields("Quantity").Value)) = 0 Then
                rstTatRVChild.Delete
                rstTatRVChild.MoveNext
                If rstTatRVChild.RecordCount > 0 Then rstTatRVChild.MoveFirst
            End If
        End If
    End If
    Text5.Visible = False
    MhRealInput1.Visible = False
    DataGrid2.Enabled = True
    DataGrid2.SetFocus
End Sub
Private Sub CalculateTotal()
    Dim dblBookMark As Double
    
    MhRealInput5.Text = 0
    If rstTatRVChild.RecordCount <> 0 Then
        If Not (rstTatRVChild.EOF Or rstTatRVChild.BOF) Then
            dblBookMark = rstTatRVChild.Bookmark
        End If
        rstTatRVChild.MoveFirst
        Do While Not rstTatRVChild.EOF
            MhRealInput5.Text = Val(MhRealInput5.Text) + Val(rstTatRVChild.Fields("Quantity").Value)
            rstTatRVChild.MoveNext
        Loop
        If dblBookMark <> 0 Then
            rstTatRVChild.Bookmark = dblBookMark
       Else
            rstTatRVChild.MoveLast
       End If
    End If
End Sub
Private Function CheckDuplicatePrinter() As Boolean
    Dim dblBookMark As Double
    
    If rstTatRVChild.RecordCount = 0 Then Exit Function
    If Not (rstTatRVChild.EOF Or rstTatRVChild.BOF) Then
       dblBookMark = rstTatRVChild.Bookmark
    End If
    rstTatRVChild.MoveFirst
    Do While Not rstTatRVChild.EOF
          If rstTatRVChild.Fields("PrinterName").Value = Trim(Text5.Text) Then
             CheckDuplicatePrinter = True
             Exit Do
          End If
          rstTatRVChild.MoveNext
    Loop
    If dblBookMark <> 0 Then
       rstTatRVChild.Bookmark = dblBookMark
    Else
       rstTatRVChild.MoveLast
    End If
End Function
Private Function UpdatePrinterList(ByVal strOption As String) As Boolean
    On Error GoTo ErrorHandler
    
    UpdatePrinterList = True
    If strOption = "D" Then
        CxnTatReceipt.Execute "Delete From TatRVChild WHERE Code = '" & rstTatRVParent.Fields("Code").Value & "'"
    Else
        CxnTatReceipt.Execute "Insert Into TatRVChild Values ('" & rstTatRVParent.Fields("Code").Value & "','" & rstTatRVChild.Fields("Printer").Value & "'," & rstTatRVChild.Fields("Quantity").Value & ")"
    End If
    Exit Function
ErrorHandler:
    UpdatePrinterList = False
End Function
Private Function CalculateTatBalance(ByVal strPrinterCode As String, ByVal strVoucherCode As String) As Long
    Dim rstTatBalance As New ADODB.Recordset
    On Error GoTo ErrorHandler
    
    If rstTatBalance.State = adStateOpen Then
        rstTatBalance.Close
    End If
    rstTatBalance.Open "Select Format((Select Sum(Tat) From PaperIOChild Where Account=M.Code),0) As Col0,Format((Select Sum(Tat) From PaperMVParent,PaperMVChild Where PaperMVParent.Code=PaperMVChild.Code And PaperMVParent.AccountFrom=M.Code),0) As Col1,Format((Select Sum(Tat) From PaperMVParent,PaperMVChild Where PaperMVParent.Code=PaperMVChild.Code And PaperMVParent.AccountTo=M.Code),0) As Col2,Format((Select Sum(Quantity) From TatRVChild Where TatRVChild.Code<>'" & strVoucherCode & "' And Printer=M.Code),0) As Col3,Format((Select Sum(OpBalTat) From PaperChild Where Account=M.Code),0) As Col4 From AccountMaster M Where Code='" & strPrinterCode & "'", CxnTatReceipt, adOpenKeyset, adLockReadOnly
    CalculateTatBalance = Val(CheckNull(rstTatBalance.Fields("Col0").Value)) - Val(CheckNull(rstTatBalance.Fields("Col1").Value)) + Val(CheckNull(rstTatBalance.Fields("Col2").Value)) - Val(CheckNull(rstTatBalance.Fields("Col3").Value)) + Val(CheckNull(rstTatBalance.Fields("Col4").Value))
    Call CloseRecordset(rstTatBalance)
    Exit Function
ErrorHandler:
    Call CloseRecordset(rstTatBalance)
End Function
Private Sub PrintTatReceipt()
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    rptTatReceipt.Text2.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptTatReceipt.Text3.SetText Trim(rstCompanyMaster.Fields("Address1").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address2").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address3").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address4").Value)
    If (Not CheckEmpty(rstCompanyMaster.Fields("Phone").Value, False)) And (Not CheckEmpty(rstCompanyMaster.Fields("Fax").Value, False)) Then
        rptTatReceipt.Text24.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & Space(1) & "Fax : " & Trim(rstCompanyMaster.Fields("Fax").Value)
    ElseIf Not CheckEmpty(rstCompanyMaster.Fields("Fax").Value, False) Then
        rptTatReceipt.Text24.SetText "Fax : " & Trim(rstCompanyMaster.Fields("Fax").Value)
    ElseIf Not CheckEmpty(rstCompanyMaster.Fields("Phone").Value, False) Then
        rptTatReceipt.Text24.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value)
    Else
        rptTatReceipt.Section5.Suppress = True
    End If
    If rstTatRVChild.State = adStateOpen Then
        rstTatRVChild.Close
    End If
    rstTatRVChild.Open "SELECT TRIM(P.Name) As VchNo,[Date] As VchDate,Trim(PrintName) As PrinterName,Quantity,Rate,Remarks FROM (TatRVParent P INNER JOIN TatRVChild C ON P.Code=C.Code) INNER JOIN AccountMaster M ON C.Printer=M.Code WHERE P.Code='" & rstTatRVList.Fields("Code").Value & "' ORDER BY M.PrintName", CxnTatReceipt, adOpenKeyset, adLockOptimistic
    rptTatReceipt.Text27.SetText "for " & Trim(rstTatRVChild.Fields("PrinterName").Value)
    rptTatReceipt.Text9.SetText "for " & Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptTatReceipt.Database.SetDataSource rstTatRVChild, 3, 1
    Screen.MousePointer = vbNormal
    If OutputTo = "S" Then
        Set FrmReportViewer.Report = rptTatReceipt
        FrmReportViewer.Show vbModal
    Else
        rptTatReceipt.PrintOut
    End If
    Set rptTatReceipt = Nothing
    On Error GoTo 0
End Sub
