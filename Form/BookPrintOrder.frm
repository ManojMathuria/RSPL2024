VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0F1F1508-C40A-101B-AD04-00AA00575482}#1.0#0"; "mhrinp32.ocx"
Object = "{FAD0952A-804F-4061-84BA-88D0F2AA07A8}#1.0#0"; "vsflex8d.ocx"
Begin VB.Form FrmBookPrintOrder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Book Print Order"
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   16725
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "BookPrintOrder.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7725
   ScaleWidth      =   16725
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   7710
      Left            =   15
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   0
      Width           =   16635
      _Version        =   65536
      _ExtentX        =   29342
      _ExtentY        =   13600
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
      Picture         =   "BookPrintOrder.frx":0442
      Begin TabDlg.SSTab SSTab1 
         Height          =   7485
         Left            =   120
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   120
         Width           =   16410
         _ExtentX        =   28945
         _ExtentY        =   13203
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
         TabPicture(0)   =   "BookPrintOrder.frx":045E
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Text1"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "DataGrid1"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "&Details"
         TabPicture(1)   =   "BookPrintOrder.frx":047A
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Mh3dFrame2"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   6495
            Left            =   120
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   450
            Width           =   16185
            _ExtentX        =   28549
            _ExtentY        =   11456
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
            ColumnCount     =   15
            BeginProperty Column00 
               DataField       =   "Name"
               Caption         =   "   Order No."
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
               Caption         =   "Order Date"
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
               DataField       =   "BookName"
               Caption         =   "Book Name"
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
            BeginProperty Column03 
               DataField       =   "BPODStatus"
               Caption         =   "BP"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   5
                  Format          =   ""
                  HaveTrueFalseNull=   1
                  TrueValue       =   "v"
                  FalseValue      =   "x"
                  NullValue       =   ""
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   7
               EndProperty
            EndProperty
            BeginProperty Column04 
               DataField       =   "TPODStatus"
               Caption         =   "TP"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   5
                  Format          =   ""
                  HaveTrueFalseNull=   1
                  TrueValue       =   "v"
                  FalseValue      =   "x"
                  NullValue       =   ""
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   7
               EndProperty
            EndProperty
            BeginProperty Column05 
               DataField       =   "TLODStatus"
               Caption         =   "TL"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   5
                  Format          =   ""
                  HaveTrueFalseNull=   1
                  TrueValue       =   "v"
                  FalseValue      =   "x"
                  NullValue       =   ""
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   7
               EndProperty
            EndProperty
            BeginProperty Column06 
               DataField       =   "BBODStatus"
               Caption         =   "BB"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   5
                  Format          =   ""
                  HaveTrueFalseNull=   1
                  TrueValue       =   "v"
                  FalseValue      =   "x"
                  NullValue       =   ""
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2057
                  SubFormatType   =   7
               EndProperty
            EndProperty
            BeginProperty Column07 
               DataField       =   "ReceivedQuantity"
               Caption         =   "    Recd Qty"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column08 
               DataField       =   "BookPrinterName"
               Caption         =   "Book Printer"
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
            BeginProperty Column09 
               DataField       =   "TitlePrinterName"
               Caption         =   "Title Printer"
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
            BeginProperty Column10 
               DataField       =   "LaminatorName"
               Caption         =   "Laminator"
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
            BeginProperty Column11 
               DataField       =   "BinderName"
               Caption         =   "Binder"
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
            BeginProperty Column12 
               DataField       =   "BookStatus"
               Caption         =   "Book Status"
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
            BeginProperty Column13 
               DataField       =   "AdvanceRecvdDate"
               Caption         =   "A.Recvd Date"
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
            BeginProperty Column14 
               DataField       =   "RefNo"
               Caption         =   "Ref No"
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
                  ColumnWidth     =   1080
               EndProperty
               BeginProperty Column01 
                  Locked          =   -1  'True
                  ColumnWidth     =   1019.906
               EndProperty
               BeginProperty Column02 
                  Locked          =   -1  'True
                  ColumnWidth     =   3644.788
               EndProperty
               BeginProperty Column03 
                  Alignment       =   2
                  ColumnAllowSizing=   0   'False
                  Locked          =   -1  'True
                  ColumnWidth     =   285.165
               EndProperty
               BeginProperty Column04 
                  Alignment       =   2
                  ColumnAllowSizing=   0   'False
                  Locked          =   -1  'True
                  ColumnWidth     =   285.165
               EndProperty
               BeginProperty Column05 
                  Alignment       =   2
                  ColumnAllowSizing=   0   'False
                  Locked          =   -1  'True
                  ColumnWidth     =   285.165
               EndProperty
               BeginProperty Column06 
                  Alignment       =   2
                  ColumnAllowSizing=   0   'False
                  Locked          =   -1  'True
                  ColumnWidth     =   285.165
               EndProperty
               BeginProperty Column07 
                  Alignment       =   1
                  Locked          =   -1  'True
                  ColumnWidth     =   975.118
               EndProperty
               BeginProperty Column08 
                  Locked          =   -1  'True
                  ColumnWidth     =   1094.74
               EndProperty
               BeginProperty Column09 
                  Locked          =   -1  'True
                  ColumnWidth     =   1065.26
               EndProperty
               BeginProperty Column10 
                  Locked          =   -1  'True
                  ColumnWidth     =   900.284
               EndProperty
               BeginProperty Column11 
               EndProperty
               BeginProperty Column12 
                  ColumnWidth     =   1124.787
               EndProperty
               BeginProperty Column13 
                  Locked          =   -1  'True
                  ColumnWidth     =   1154.835
               EndProperty
               BeginProperty Column14 
                  ColumnWidth     =   854.929
               EndProperty
            EndProperty
         End
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
            TabIndex        =   24
            Top             =   7020
            Width           =   15705
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   3150
            Left            =   -74880
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   480
            Width           =   8235
            _Version        =   65536
            _ExtentX        =   14526
            _ExtentY        =   5556
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
            Picture         =   "BookPrintOrder.frx":0496
            Begin VB.TextBox Text10 
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
               Left            =   7035
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   3
               TabStop         =   0   'False
               Top             =   630
               Width           =   1095
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
               Left            =   1440
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   4
               TabStop         =   0   'False
               Top             =   950
               Width           =   1050
            End
            Begin VB.TextBox Text9 
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
               Left            =   3120
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   5
               TabStop         =   0   'False
               Top             =   950
               Width           =   1170
            End
            Begin VB.CommandButton Command4 
               Caption         =   "...."
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
               Left            =   7800
               TabIndex        =   15
               ToolTipText     =   "Book Binding Order"
               Top             =   2210
               Width           =   330
            End
            Begin VB.CommandButton Command3 
               Caption         =   "...."
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
               Left            =   7800
               TabIndex        =   13
               ToolTipText     =   "Title Lamination Order"
               Top             =   1890
               Width           =   330
            End
            Begin VB.CommandButton Command2 
               Caption         =   "...."
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
               Left            =   7800
               TabIndex        =   11
               ToolTipText     =   "Title Printing Order"
               Top             =   1580
               Width           =   330
            End
            Begin VB.CommandButton Command1 
               Caption         =   "...."
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               Left            =   7800
               TabIndex        =   9
               ToolTipText     =   "Book Printing Order"
               Top             =   1270
               Width           =   330
            End
            Begin VB.TextBox Text8 
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
               Left            =   1440
               MaxLength       =   40
               TabIndex        =   14
               Top             =   2210
               Width           =   6360
            End
            Begin VB.TextBox Text7 
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
               Left            =   1440
               MaxLength       =   40
               TabIndex        =   12
               Top             =   1890
               Width           =   6360
            End
            Begin VB.TextBox Text6 
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
               Left            =   1440
               MaxLength       =   40
               TabIndex        =   10
               Top             =   1580
               Width           =   6360
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
               Left            =   1440
               MaxLength       =   10
               TabIndex        =   0
               Top             =   105
               Width           =   1050
            End
            Begin VB.TextBox Text5 
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
               Left            =   1440
               MaxLength       =   40
               TabIndex        =   8
               Top             =   1260
               Width           =   6360
            End
            Begin VB.TextBox Text3 
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
               Left            =   1440
               MaxLength       =   60
               TabIndex        =   2
               Top             =   630
               Width           =   4530
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
               Height          =   330
               Left            =   120
               TabIndex        =   27
               Top             =   105
               Width           =   1335
               _Version        =   65536
               _ExtentX        =   2355
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
               Caption         =   " Order No."
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "BookPrintOrder.frx":04B2
               Picture         =   "BookPrintOrder.frx":04CE
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Left            =   5955
               TabIndex        =   28
               Top             =   105
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
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
               Caption         =   " Order Date"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "BookPrintOrder.frx":04EA
               Picture         =   "BookPrintOrder.frx":0506
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   330
               Left            =   120
               TabIndex        =   29
               Top             =   630
               Width           =   1335
               _Version        =   65536
               _ExtentX        =   2355
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
               Caption         =   " Book Name"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "BookPrintOrder.frx":0522
               Picture         =   "BookPrintOrder.frx":053E
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Left            =   120
               TabIndex        =   30
               Top             =   1260
               Width           =   1335
               _Version        =   65536
               _ExtentX        =   2355
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
               Caption         =   " Book Printer"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "BookPrintOrder.frx":055A
               Picture         =   "BookPrintOrder.frx":0576
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
               Height          =   330
               Left            =   120
               TabIndex        =   31
               Top             =   1580
               Width           =   1335
               _Version        =   65536
               _ExtentX        =   2355
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
               Caption         =   " Title Printer"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "BookPrintOrder.frx":0592
               Picture         =   "BookPrintOrder.frx":05AE
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
               Height          =   330
               Left            =   120
               TabIndex        =   32
               Top             =   1890
               Width           =   1335
               _Version        =   65536
               _ExtentX        =   2355
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
               Caption         =   " Laminator"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "BookPrintOrder.frx":05CA
               Picture         =   "BookPrintOrder.frx":05E6
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel7 
               Height          =   330
               Left            =   120
               TabIndex        =   33
               Top             =   2210
               Width           =   1335
               _Version        =   65536
               _ExtentX        =   2355
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
               Caption         =   " Book Binder"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "BookPrintOrder.frx":0602
               Picture         =   "BookPrintOrder.frx":061E
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
               Height          =   330
               Left            =   120
               TabIndex        =   34
               Top             =   950
               Width           =   1335
               _Version        =   65536
               _ExtentX        =   2355
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
               Caption         =   " Book Size"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "BookPrintOrder.frx":063A
               Picture         =   "BookPrintOrder.frx":0656
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
               Height          =   330
               Left            =   2475
               TabIndex        =   35
               Top             =   950
               Width           =   660
               _Version        =   65536
               _ExtentX        =   1164
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
               Caption         =   " Alias"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "BookPrintOrder.frx":0672
               Picture         =   "BookPrintOrder.frx":068E
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel9 
               Height          =   330
               Left            =   5955
               TabIndex        =   36
               Top             =   950
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
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
               Caption         =   " Pages"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "BookPrintOrder.frx":06AA
               Picture         =   "BookPrintOrder.frx":06C6
            End
            Begin MhinrelLib.MhRealInput MhRealInput1 
               Height          =   330
               Left            =   4920
               TabIndex        =   6
               TabStop         =   0   'False
               Top             =   950
               Width           =   1050
               _Version        =   65536
               _ExtentX        =   1852
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               FillColor       =   16777215
               MaxReal         =   999999
               MinReal         =   0
               ReadOnly        =   -1  'True
               SpinChangeReal  =   0
               CaretColor      =   -2147483642
               DecimalPlaces   =   2
               VAlignment      =   2
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel10 
               Height          =   330
               Left            =   4275
               TabIndex        =   37
               Top             =   950
               Width           =   660
               _Version        =   65536
               _ExtentX        =   1164
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
               Caption         =   " Forms"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "BookPrintOrder.frx":06E2
               Picture         =   "BookPrintOrder.frx":06FE
            End
            Begin MhinrelLib.MhRealInput MhRealInput2 
               Height          =   330
               Left            =   7035
               TabIndex        =   7
               TabStop         =   0   'False
               Top             =   950
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TintColor       =   16711935
               FillColor       =   16777215
               MaxReal         =   999999
               MinReal         =   0
               ReadOnly        =   -1  'True
               SpinChangeReal  =   0
               CaretColor      =   -2147483642
               DecimalPlaces   =   0
               VAlignment      =   2
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel12 
               Height          =   330
               Left            =   5955
               TabIndex        =   38
               Top             =   630
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
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
               Caption         =   " Color"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "BookPrintOrder.frx":071A
               Picture         =   "BookPrintOrder.frx":0736
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
               Calendar        =   "BookPrintOrder.frx":0752
               Caption         =   "BookPrintOrder.frx":086A
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookPrintOrder.frx":08D6
               Keys            =   "BookPrintOrder.frx":08F4
               Spin            =   "BookPrintOrder.frx":0952
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
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel13 
               Height          =   330
               Left            =   120
               TabIndex        =   39
               Top             =   2720
               Width           =   1335
               _Version        =   65536
               _ExtentX        =   2355
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
               Caption         =   " PO Email Status"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "BookPrintOrder.frx":097A
               Picture         =   "BookPrintOrder.frx":0996
            End
            Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame4 
               Height          =   330
               Left            =   1440
               TabIndex        =   40
               TabStop         =   0   'False
               Top             =   2720
               Width           =   6690
               _Version        =   65536
               _ExtentX        =   11800
               _ExtentY        =   582
               _StockProps     =   77
               TintColor       =   16711935
               Alignment       =   0
               AutoSize        =   0   'False
               BevelSize       =   0
               BevelStyle      =   0
               BorderColor     =   -2147483642
               BorderStyle     =   1
               FillColor       =   16777215
               FontStyle       =   0
               FontTransparent =   0   'False
               LightColor      =   -2147483643
               ShadowColor     =   -2147483632
               TextColor       =   -2147483640
               WallPaper       =   0
               NoPrefix        =   0   'False
               FormatString    =   ""
               Caption         =   ""
               Picture         =   "BookPrintOrder.frx":09B2
               Begin VB.CheckBox chkBB 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Book Binding"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Left            =   5190
                  TabIndex        =   19
                  Top             =   60
                  Width           =   1380
               End
               Begin VB.CheckBox chkTL 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Title Lamination"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Left            =   3315
                  TabIndex        =   18
                  Top             =   60
                  Width           =   1630
               End
               Begin VB.CheckBox chkTP 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Title Printing"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Left            =   1700
                  TabIndex        =   17
                  Top             =   60
                  Width           =   1455
               End
               Begin VB.CheckBox chkBP 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Book Printing"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Left            =   100
                  TabIndex        =   16
                  Top             =   60
                  Width           =   1425
               End
            End
            Begin VB.Line Line2 
               X1              =   0
               X2              =   8280
               Y1              =   2630
               Y2              =   2630
            End
            Begin VB.Line Line1 
               X1              =   0
               X2              =   8280
               Y1              =   525
               Y2              =   525
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
            TabIndex        =   25
            Top             =   7020
            Width           =   495
         End
      End
      Begin VSFlex8DAOCtl.VSFlexGrid VSFlexGrid1 
         Height          =   5415
         Left            =   0
         TabIndex        =   41
         Top             =   0
         Visible         =   0   'False
         Width           =   2655
         _cx             =   4683
         _cy             =   9551
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   16725
      _ExtentX        =   29501
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
Attribute VB_Name = "FrmBookPrintOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CxnBookPrintOrder As New ADODB.Connection
Dim rstCompanyMaster As New ADODB.Recordset
Dim rstBookPOList As New ADODB.Recordset
Dim rstBookPOParent As New ADODB.Recordset
Dim rstBookPOChild05 As New ADODB.Recordset
Dim rstBookPOChild06 As New ADODB.Recordset
Dim rstBookPOChild07 As New ADODB.Recordset
Dim rstBookPOChild08 As New ADODB.Recordset
Dim rstBookPOChild0801 As New ADODB.Recordset
Dim rstCorrections As New ADODB.Recordset
Public rstBookList As New ADODB.Recordset
Dim rstBookPrinterList As New ADODB.Recordset
Dim rstTitlePrinterList As New ADODB.Recordset
Dim rstLaminatorList As New ADODB.Recordset
Dim rstBinderList As New ADODB.Recordset
Dim PaperCode As String
Dim BookCode As String

Dim BookPrinterCode As String
Dim TitlePrinterCode As String
Dim LaminatorCode As String
Dim BinderCode As String
Dim SortOrder As String
Dim PrevStr As String
Dim dblBookMark As Double
Dim blnRecordExist As Boolean
Dim oOutlook As New Outlook.Application
Dim EMailID As String
Dim Attachment As String
Dim Message As String
Dim OutputTo As String
Public BookPOType As String
Dim LRange As Long, URange As Long, PrintReport As Integer
Dim PaperCode_For_Balance As String
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    Unload FrmBookPOPrintUtility
    CenterForm Me
    BusySystemIndicator True
    WheelHook DataGrid1
    
    If BookPOType = "F" Then
        Me.Caption = "Book Print Order [Fresh]"
    ElseIf BookPOType = "R" Then
        Me.Caption = "Book Print Order [Repair]"
    Else
        Me.Caption = "Cost Sheet"
    End If
    
    CxnBookPrintOrder.CursorLocation = adUseClient
    CxnBookPrintOrder.Open CxnDatabase.ConnectionString
     
    'Need to be check here
    
    rstCompanyMaster.Open "Select PrintName, Address1, Address2, Address3, Address4, Phone, Fax, EMail, Website From CompanyMaster", CxnBookPrintOrder, adOpenKeyset, adLockReadOnly
    
    rstBookList.Open "Select M1.Name As Col0,BusyCode As Col1,M3.Name As BoardName,M2.Name As SizeName,M2.Code As SizeCode,FormType,Forms,Pages,OneColorPages,TwoColorPages,FourColorPages,OneColor¼Forms,OneColor½Forms,[OneColor1F/BForms],[OneColor1W/TForms],OneColorForms,TwoColor¼Forms,TwoColor½Forms,[TwoColor1F/BForms],[TwoColor1W/TForms],TwoColorForms,FourColor¼Forms,FourColor½Forms,[FourColor1F/BForms],[FourColor1W/TForms],FourColorForms,OneColorPlateType,TwoColorPlateType,FourColorPlateType,DuplexPrinting,BindingType,LaminationType,TitlePlateType,BindingForms01,BindingForms02,TitleFrontColor,TitleBackColor,TitlePlateType,[Qty/Pkt],[Pkt/Box],[LooseQty/Box],AddOnRate01,AddOnRate02,BookPrinter,TitlePrinter,Laminator,BinderFresh,BinderRepair,M1.Code From BookMaster M1,GeneralMaster M2,GeneralMaster M3 Where M1.[Size] = M2.Code AND M1.Board=M3.Code Order by M1.Name", CxnBookPrintOrder, adOpenKeyset, adLockOptimistic
    
    rstBookPrinterList.Open "Select Name As Col0, Code From AccountMaster Where Type = '05' Order by Name", CxnBookPrintOrder, adOpenKeyset, adLockReadOnly
        
    rstTitlePrinterList.Open "Select Name As Col0, Code From AccountMaster Where Type = '06' Order by Name", CxnBookPrintOrder, adOpenKeyset, adLockReadOnly
    
    rstLaminatorList.Open "Select Name As Col0, Code From AccountMaster Where Type = '07' Order by Name", CxnBookPrintOrder, adOpenKeyset, adLockReadOnly
    
    rstBinderList.Open "Select Name As Col0, Code From AccountMaster Where Type = '08' Order by Name", CxnBookPrintOrder, adOpenKeyset, adLockReadOnly
    
    
    Dim aaa As String
    
    aaa = "SELECT T.Code,T.Name,Date,M.Name As BookName,BPODStatus,TPODStatus,TLODStatus,BBODStatus,ReceivedQuantity,(SELECT Name FROM AccountMaster WHERE Code=T.BookPrinter) As BookPrinterName,(SELECT Name FROM AccountMaster WHERE Code=T.TitlePrinter) As TitlePrinterName,(SELECT Name FROM AccountMaster WHERE Code=T.Laminator) As LaminatorName,(SELECT Name FROM AccountMaster WHERE Code=T.Binder) As BinderName,(SELECT BookStatus FROM BookPOChild05 WHERE Code=T.Code) As BookStatus,(SELECT Name FROM PrintPVParent WHERE Code=(Select Ref  FROM BookPOChild05 WHERE Code=T.Code)) As RefNo,(SELECT AdvanceRecvdDate FROM BookPOChild08 WHERE Code=T.Code) As AdvanceRecvdDate FROM BookPOParent T INNER JOIN BookMaster M ON T.Book=M.Code WHERE T.Type = '" & BookPOType & "' AND LEFT(T.Code,1)<>'*' ORDER BY T.Name"
    
    
    rstBookPOList.Open "SELECT T.Code,T.Name,Date,M.Name As BookName,BPODStatus,TPODStatus,TLODStatus,BBODStatus,ReceivedQuantity,(SELECT Name FROM AccountMaster WHERE Code=T.BookPrinter) As BookPrinterName,(SELECT Name FROM AccountMaster WHERE Code=T.TitlePrinter) As TitlePrinterName,(SELECT Name FROM AccountMaster WHERE Code=T.Laminator) As LaminatorName,(SELECT Name FROM AccountMaster WHERE Code=T.Binder) As BinderName,(SELECT BookStatus FROM BookPOChild05 WHERE Code=T.Code) As BookStatus,(SELECT Name FROM PrintPVParent WHERE Code=(Select Ref  FROM BookPOChild05 WHERE Code=T.Code)) As RefNo,(SELECT AdvanceRecvdDate FROM BookPOChild08 WHERE Code=T.Code) As AdvanceRecvdDate FROM BookPOParent T INNER JOIN BookMaster M ON T.Book=M.Code WHERE T.Type = '" & BookPOType & "' AND LEFT(T.Code,1)<>'*' ORDER BY T.Name", CxnBookPrintOrder, adOpenKeyset, adLockOptimistic
    
    
    
    rstBookPOParent.CursorLocation = adUseClient
    
    rstBookPOList.Filter = adFilterNone
            
    If rstBookPOList.RecordCount > 0 Then rstBookPOList.MoveLast
    
    Set DataGrid1.DataSource = rstBookPOList
    
    BusySystemIndicator False
    
    SSTab1.Tab = 0
    SortOrder = "Name"
            
    If Not (rstBookPOList.EOF Or rstBookPOList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    rstBookPOList.ActiveConnection = Nothing
    rstBookList.ActiveConnection = Nothing
    rstBookPrinterList.ActiveConnection = Nothing
    rstTitlePrinterList.ActiveConnection = Nothing
    rstLaminatorList.ActiveConnection = Nothing
    rstBinderList.ActiveConnection = Nothing
    SetButtonsForNoRecord
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    CloseForm Me
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
            CloseForm Me
        Else
            If Toolbar1.Buttons.Item(1).Enabled Then
                SSTab1.Tab = 0
            Else
                If MsgBox("Are you sure to Quit?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Quit !") <> vbYes Then
                    Me.ActiveControl.SetFocus
                Else
                    Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
                End If
            End If
            KeyCode = 0
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
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(4)
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
            Sendkeys "{TAB}"
        End If
        KeyCode = 0
        
    End If
    
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    If Toolbar1.Buttons.Item(4).Enabled Then
        Call Form_KeyDown(vbKeyEscape, 0)
        Cancel = 1
    Else
        CloseForm Me
    End If
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    WheelUnHook
    Call CloseRecordset(rstCompanyMaster)
    Call CloseRecordset(rstBookPOList)
    Call CloseRecordset(rstBookPOParent)
    Call CloseRecordset(rstBookPOChild05)
    Call CloseRecordset(rstBookPOChild06)
    Call CloseRecordset(rstBookPOChild07)
    Call CloseRecordset(rstBookPOChild08)
    Call CloseRecordset(rstBookPOChild0801)
    Call CloseRecordset(rstCorrections)
    Call CloseRecordset(rstBookList)
    Call CloseRecordset(rstBookPrinterList)
    Call CloseRecordset(rstTitlePrinterList)
    Call CloseRecordset(rstLaminatorList)
    Call CloseRecordset(rstBinderList)
    'Call CloseRecordset(rstPaperRegister)
    Call CloseConnection(CxnBookPrintOrder)
    OutputTo = ""
    ShowProgressInStatusBar False
    DisableChildMenu
End Sub

Private Sub Text1_Change()
    If rstBookPOList.RecordCount = 0 Then Exit Sub
    rstBookPOList.MoveFirst
    If Text1.Text <> "" Then
        If SortOrder = "Name" Then
           rstBookPOList.Find "[" & SortOrder & "] Like '%" & FixQuote(Text1.Text) & "%'"
        Else
           rstBookPOList.Find "[" & SortOrder & "] Like '" & FixQuote(Text1.Text) & "%'"
        End If
        If rstBookPOList.EOF Then
            rstBookPOList.MoveFirst
            If PrevStr <> "" And Len(Text1.Text) > 1 Then
                If dblBookMark <> 0 Then
                    rstBookPOList.Bookmark = dblBookMark
                End If
            Else
                PrevStr = ""
            End If
            Beep
            DisplayError ("Spelling Error")
            Text1.Text = PrevStr
            Sendkeys "{End}"
        Else
            PrevStr = Text1.Text
            dblBookMark = DataGrid1.Bookmark
        End If
    Else
        PrevStr = ""
    End If
    
    If Not (rstBookPOList.EOF Or rstBookPOList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim KeyProcessed As Boolean
    If rstBookPOList.RecordCount = 0 Then Exit Sub
    If Shift = 0 And KeyCode = vbKeyUp Then
        With rstBookPOList
            .MovePrevious
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyBack Then
        With rstBookPOList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyDown Then
        With rstBookPOList
            .MoveNext
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageUp Then
        With rstBookPOList
            .Move (-1) * (DataGrid1.VisibleRows - 1)
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageUp Then
        With rstBookPOList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageDown Then
        With rstBookPOList
            .Move DataGrid1.VisibleRows - 1
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageDown Then
        With rstBookPOList
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
            Me.Width = 8805
            Mh3dFrame1.Width = 8700
            SSTab1.Width = 8460
            CenterForm Me
            ViewRecord
        Else
            Me.Width = 16815 '13155
            Mh3dFrame1.Width = 16635 '13040
            SSTab1.Width = 16410 '12810
            CenterForm Me
            If Not (rstBookPOList.EOF Or rstBookPOList.BOF) Then
                With DataGrid1.SelBookmarks
                    If .Count <> 0 Then .Remove 0
                    .Add DataGrid1.Bookmark
                End With
            End If
            Text1.SetFocus
        End If
        SSTab1.TabEnabled(0) = True
    Else
        Me.Width = 8805
        Mh3dFrame1.Width = 8700
        SSTab1.Width = 8460
        CenterForm Me
        SSTab1.TabEnabled(0) = False
        Text2.SetFocus
    End If
    
End Sub
Public Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim HiLiteRecord As Boolean
    Dim UpdateFlag As Integer, i As Integer, CellVal As Variant
    If Button.Index = 1 Then
        If rstBookPOParent.State = adStateOpen Then
           rstBookPOParent.Close
        End If
        rstBookPOParent.Open "Select * From BookPOParent Where Code = ''", CxnBookPrintOrder, adOpenKeyset, adLockOptimistic
        ClearFields
        Call LoadOrder("")
        If rstBookPOChild05.State = adStateClosed Or rstBookPOChild06.State = adStateClosed Or rstBookPOChild07.State = adStateClosed Or rstBookPOChild08.State = adStateClosed Or rstBookPOChild0801.State = adStateClosed Then
            SSTab1.Tab = 0
            Exit Sub
        End If
        Me.Tag = "A"
        If AddRecord(rstBookPOParent) Then
            Text2.Text = GenerateCode(CxnBookPrintOrder, "Select Max(Val(Name)) From BookPOParent Where Type = '" & BookPOType & "' AND LEFT(Code,1)<>'*'", 10, Space(1))
            MhDateInput1.Text = Format(Date, "dd-MM-yyyy")
            Call SetButtons(False)
            SSTab1.Tab = 1
            Text2.SetFocus
            blnRecordExist = False
            CxnBookPrintOrder.BeginTrans
        End If
    ElseIf Button.Index = 2 Then
        If rstBookPOList.RecordCount = 0 Then Exit Sub
        SSTab1.Tab = 1
        Me.Tag = "E"
        EditRecord
    ElseIf Button.Index = 3 Then
        If rstBookPOList.RecordCount = 0 Then Exit Sub
        If AllowTransactionsDeletion = 0 Then
            Call DisplayError("You don't have the rights to Delete this Voucher")
            Exit Sub
        End If
        SSTab1.Tab = 1
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            On Error Resume Next
            MdiMainMenu.MousePointer = vbHourglass
            CxnBookPrintOrder.Execute "Delete From BookPOParent Where Code = '" & rstBookPOList.Fields("Code").Value & "'"
            MdiMainMenu.MousePointer = vbNormal
            If Err.Number = 0 Then
                rstBookPOList.Delete
                rstBookPOList.MoveNext
                If rstBookPOList.RecordCount > 0 And rstBookPOList.EOF Then
                    rstBookPOList.MoveLast
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
'        If blnRecordExist And AllowTransactionsModification = 0 Then
'            Call DisplayError("You don't have the rights to Edit this Voucher")
'            Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
'            Exit Sub
'        End If
        SaveFields
        UpdateFlag = 0
        If UpdateRecord(rstBookPOParent) Then
            If UpdateOrder("D") Then
                UpdateFlag = 1
                If rstBookPOChild05.RecordCount <> 0 Then
                    rstBookPOChild05.MoveFirst
                    Do While Not rstBookPOChild05.EOF
                        If Val(rstBookPOChild05.Fields("ActualQuantity").Value) <> 0 Then
                            If Not UpdateOrder("U", "1") Then UpdateFlag = 0: Exit Do
                        End If
                        rstBookPOChild05.MoveNext
                    Loop
                End If
                If UpdateFlag Then
                    If rstBookPOChild06.RecordCount <> 0 Then
                        rstBookPOChild06.MoveFirst
                        Do While Not rstBookPOChild06.EOF
                            If Val(rstBookPOChild06.Fields("ActualQuantity").Value) <> 0 Then
                                If Not UpdateOrder("U", "2") Then UpdateFlag = 0: Exit Do
                            End If
                            rstBookPOChild06.MoveNext
                        Loop
                    End If
                End If
                If UpdateFlag Then
                    If rstBookPOChild07.RecordCount <> 0 Then
                        rstBookPOChild07.MoveFirst
                        Do While Not rstBookPOChild07.EOF
                            If Val(rstBookPOChild07.Fields("ActualQuantity").Value) <> 0 Then
                                If Not UpdateOrder("U", "3") Then UpdateFlag = 0: Exit Do
                            End If
                            rstBookPOChild07.MoveNext
                        Loop
                    End If
                End If
                If UpdateFlag Then
                    If rstBookPOChild08.RecordCount <> 0 Then
                        rstBookPOChild08.MoveFirst
                        Do While Not rstBookPOChild08.EOF
                            If Val(rstBookPOChild08.Fields("ActualQuantity").Value) <> 0 Then
                                If Not UpdateOrder("U", "4") Then UpdateFlag = 0: Exit Do
                            End If
                            rstBookPOChild08.MoveNext
                        Loop
                        If UpdateFlag Then
                            If rstBookPOChild0801.RecordCount > 0 Then rstBookPOChild0801.MoveFirst
                            Do While Not rstBookPOChild0801.EOF
                                If Val(rstBookPOChild0801.Fields("Quantity").Value) <> 0 Then
                                    If Not UpdateOrder("U", "0") Then UpdateFlag = 0: Exit Do
                                End If
                                rstBookPOChild0801.MoveNext
                            Loop
                        End If
                    End If
                End If
                If UpdateFlag Then
                    With FrmCorrectionRegister
                        Dim SNo As Variant
                        i = 1
                        On Error Resume Next
                        For i = 1 To .fpSpread3.DataRowCnt
                            .fpSpread3.GetText 1, i, SNo
                            If Val(SNo) = 1 Then
                                .fpSpread3.GetText 4, i, SNo
                                CxnBookPrintOrder.Execute "UPDATE BookChild02 SET RectifiedOn='" & Trim(Text2.Text) & "/" & Format(GetDate(MhDateInput1.Text), "dd-MM-yyyy") & "' WHERE Code='" & BookCode & "' AND SNo=" & Val(SNo)
                                If Err.Number <> 0 Then UpdateFlag = 0: Exit For
                            End If
                        Next
                        On Error GoTo 0
                    End With
                End If
            End If
        End If
        Call CloseForm(FrmCorrectionRegister)
        If UpdateFlag Then
            AddToList
            CxnBookPrintOrder.CommitTrans
            UpdateLastPrinterBinder
            If rstBookPOParent.State = adStateOpen Then
                rstBookPOParent.Close
            End If
            rstBookPOParent.CursorLocation = adUseClient
            Call SetButtons(True)
            SSTab1.Tab = 0
            ShowProgressInStatusBar True
            Timer1.Enabled = True
            LockFields (False)
        Else
            DisplayError ("Failed to save the record")
            Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
        End If
        Me.Tag = ""
    ElseIf Button.Index = 5 Then
        If CancelRecordUpdate(rstBookPOParent) Then
            CxnBookPrintOrder.RollbackTrans
            If rstBookPOParent.State = adStateOpen Then rstBookPOParent.Close
            rstBookPOParent.CursorLocation = adUseClient
            Call SetButtons(True)
            SetButtonsForNoRecord
            SSTab1.Tab = 0
            LockFields (False)
            Call CloseForm(FrmCorrectionRegister)
        End If
        Me.Tag = ""
        
    ElseIf Button.Index = 6 Then
        SSTab1.Tab = 0
        Set DataGrid1.DataSource = Nothing
        rstBookPOList.ActiveConnection = CxnBookPrintOrder
        Do While Not RefreshRecord(rstBookPOList)
        Loop
        Set DataGrid1.DataSource = rstBookPOList
        rstBookPOList.ActiveConnection = Nothing
        If rstBookPOList.RecordCount > 0 Then rstBookPOList.MoveLast
        rstBookList.ActiveConnection = CxnBookPrintOrder
        Do While Not RefreshRecord(rstBookList)
        Loop
        rstBookList.ActiveConnection = Nothing
        rstBookPrinterList.ActiveConnection = CxnBookPrintOrder
        Do While Not RefreshRecord(rstBookPrinterList)
        Loop
        rstBookPrinterList.ActiveConnection = Nothing
        rstTitlePrinterList.ActiveConnection = CxnBookPrintOrder
        Do While Not RefreshRecord(rstTitlePrinterList)
        Loop
        rstTitlePrinterList.ActiveConnection = Nothing
        rstLaminatorList.ActiveConnection = CxnBookPrintOrder
        Do While Not RefreshRecord(rstLaminatorList)
        Loop
        rstLaminatorList.ActiveConnection = Nothing
        rstBinderList.ActiveConnection = CxnBookPrintOrder
        Do While Not RefreshRecord(rstBinderList)
        Loop
        rstBinderList.ActiveConnection = Nothing
        HiLiteRecord = True
    ElseIf Button.Index = 7 Then
        SSTab1.Tab = 0
        With FrmFilter
            .Combo1.AddItem "Supplier", 0
            .Combo1.ListIndex = 0
            Set .srcForm = Me
            .Show vbModal
        End With
        HiLiteRecord = True
    ElseIf Button.Index = 9 Then
        OutputTo = "P"
        If BookPOType <> "O" Then DisplayMenu Else PrintCostSheet (rstBookPOList.Fields("Code").Value)
    ElseIf Button.Index = 10 Then
        OutputTo = "S"
        If BookPOType <> "O" Then DisplayMenu Else PrintCostSheet (rstBookPOList.Fields("Code").Value)
    ElseIf Button.Index = 13 Then
        If rstBookPOList.RecordCount > 0 Then rstBookPOList.MoveFirst
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 14 Then
        If rstBookPOList.RecordCount > 0 Then
            rstBookPOList.MovePrevious
            If rstBookPOList.BOF Then
                rstBookPOList.MoveNext
            End If
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 15 Then
        If rstBookPOList.RecordCount > 0 Then
            rstBookPOList.MoveNext
            If rstBookPOList.EOF Then
                rstBookPOList.MovePrevious
            End If
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 16 Then
        If rstBookPOList.RecordCount > 0 Then rstBookPOList.MoveLast
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 18 Then
        CloseForm Me
        HiLiteRecord = False
    End If
    
    If HiLiteRecord Then
        If Not (rstBookPOList.EOF Or rstBookPOList.BOF) Then
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
Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
    If ColIndex = 0 Then
       If SortOrder <> "Name" Then
          SortOrder = "Name"
          rstBookPOList.Sort = "Name Asc"
       End If
    ElseIf ColIndex = 2 Then
       If SortOrder <> "BookName" Then
          SortOrder = "BookName"
          rstBookPOList.Sort = "BookName Asc"
       End If
    ElseIf ColIndex = 3 Then
       If SortOrder <> "BoardName" Then
          SortOrder = "BoardName"
          rstBookPOList.Sort = "BoardName,BookName Asc"
       End If
    ElseIf ColIndex = 8 Then
       If SortOrder <> "BookPrinterName" Then
          SortOrder = "BookPrinterName"
          rstBookPOList.Sort = "BookPrinterName,BookName Asc"
       End If
    ElseIf ColIndex = 9 Then
       If SortOrder <> "TitlePrinterName" Then
          SortOrder = "TitlePrinterName"
          rstBookPOList.Sort = "TitlePrinterName,BookName Asc"
       End If
    ElseIf ColIndex = 10 Then
       If SortOrder <> "LaminatorName" Then
          SortOrder = "LaminatorName"
          rstBookPOList.Sort = "LaminatorName,BookName Asc"
       End If
    ElseIf ColIndex = 11 Then
       
       If SortOrder <> "BinderName" Then
          SortOrder = "BinderName"
          rstBookPOList.Sort = "BinderName,BookName Asc"
       End If
       
    End If
    DataGrid1.ClearSelCols
    If Not (rstBookPOList.EOF Or rstBookPOList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    Text1.Text = ""
    Text1.SetFocus
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
    If rstBookPOList.RecordCount = 0 Then
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
    If rstBookPOParent.EOF Or rstBookPOParent.BOF Then Exit Sub
    If CheckEmpty(Text2, True) Then
        Cancel = True
    ElseIf CheckDuplicate(CxnBookPrintOrder, "BookPOParent", "Code", "[Name]+[Type]", Trim(Text2.Text) & BookPOType, rstBookPOParent.Fields("Code").Value, False) Then
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
Private Sub Text3_Change()
    If Text3.Text = " " Then
        Text3.Text = "?"
        Sendkeys "{TAB}"
    End If
End Sub
Private Sub Text3_Validate(Cancel As Boolean)
    Dim SearchString As String
    
    SearchString = FixQuote(Text3.Text)
    If rstBookList.RecordCount = 0 Then
        DisplayError ("No Record in Book Master")
        Cancel = True
        Exit Sub
    Else
        rstBookList.MoveFirst
    End If
    rstBookList.Find "[Col0] = '" & RTrim(SearchString) & "'"
    If rstBookList.EOF Then
        SelectionType = "S"
        BookCode = ""
        Call LoadSelectionList(rstBookList, "List of Books...", "Name", "Alias")
        SearchOrder = 0
        Call DisplaySelectionList(Text3, BookCode)
        Call CloseForm(FrmSelectionList)
        If CheckEmpty(Text3.Text, False) Then
            Text3.Text = "?"
        End If
        If RTrim(BookCode) <> "" Then
            Sendkeys "{TAB}"
        End If
        Cancel = True
    Else
        BookCode = rstBookList.Fields("Code").Value
        
        Text9.Text = rstBookList.Fields("Col1").Value
        If Val(rstBookList.Fields("TwoColorPages").Value) = 0 And Val(rstBookList.Fields("FourColorPages").Value) = 0 Then
            Text10.Text = "1 Color"
        ElseIf Val(rstBookList.Fields("OneColorPages").Value) = 0 And Val(rstBookList.Fields("FourColorPages").Value) = 0 Then
            Text10.Text = "2 Color"
        ElseIf Val(rstBookList.Fields("OneColorPages").Value) = 0 And Val(rstBookList.Fields("TwoColorPages").Value) = 0 Then
            Text10.Text = "4 Color"
        ElseIf Val(rstBookList.Fields("OneColorPages").Value) = 0 And Val(rstBookList.Fields("TwoColorPages").Value) = 0 And Val(rstBookList.Fields("FourColorPages").Value) = 0 Then
            Text10.Text = "6 Color"
        Else
            Text10.Text = "Multi Color"
        End If
        Text4.Text = rstBookList.Fields("SizeName").Value & "/" & IIf(rstBookList.Fields("FormType").Value = "1", "08", IIf(rstBookList.Fields("FormType").Value = "2", "16", IIf(rstBookList.Fields("FormType").Value = "3", "04", IIf(rstBookList.Fields("FormType").Value = "4", "12", IIf(rstBookList.Fields("FormType").Value = "5", "24", IIf(rstBookList.Fields("FormType").Value = "6", "32", IIf(rstBookList.Fields("FormType").Value = "7", "64", "06")))))))

        MhRealInput1.Text = Val(rstBookList.Fields("Forms").Value)
        MhRealInput2.Text = Val(rstBookList.Fields("Pages").Value)
        
        
        If Me.Tag = "A" Then
            If Trim(rstBookList.Fields("BookPrinter")) <> "" Then
                rstBookPrinterList.MoveFirst
                BookPrinterCode = rstBookList.Fields("BookPrinter")
                rstBookPrinterList.Find "[Code] = '" & RTrim(BookPrinterCode) & "'"
                If Not rstBookPrinterList.EOF Then
                    Text5.Text = rstBookPrinterList.Fields("Col0").Value
                End If
            End If
            If Trim(rstBookList.Fields("TitlePrinter")) <> "" Then
                rstTitlePrinterList.MoveFirst
                TitlePrinterCode = rstBookList.Fields("TitlePrinter")
                rstTitlePrinterList.Find "[Code] = '" & RTrim(TitlePrinterCode) & "'"
                If Not rstTitlePrinterList.EOF Then
                    Text6.Text = rstTitlePrinterList.Fields("Col0").Value
                End If
            End If
            If Trim(rstBookList.Fields("Laminator")) <> "" Then
                rstLaminatorList.MoveFirst
                LaminatorCode = rstBookList.Fields("Laminator")
                rstLaminatorList.Find "[Code] = '" & RTrim(LaminatorCode) & "'"
                If Not rstLaminatorList.EOF Then
                    Text7.Text = rstLaminatorList.Fields("Col0").Value
                End If
            End If
            If BookPOType = "F" Or BookPOType = "O" Then
                If Trim(rstBookList.Fields("BinderFresh")) <> "" Then
                    rstBinderList.MoveFirst
                    BinderCode = rstBookList.Fields("BinderFresh")
                    rstBinderList.Find "[Code] = '" & RTrim(BinderCode) & "'"
                    If Not rstBinderList.EOF Then
                        Text8.Text = rstBinderList.Fields("Col0").Value
                    End If
                End If
            ElseIf BookPOType = "R" Then
                If Trim(rstBookList.Fields("BinderRepair")) <> "" Then
                    rstBinderList.MoveFirst
                    BinderCode = rstBookList.Fields("BinderRepair")
                    rstBinderList.Find "[Code] = '" & RTrim(BinderCode) & "'"
                    If Not rstBinderList.EOF Then
                        Text8.Text = rstBinderList.Fields("Col0").Value
                    End If
                End If
            End If
        End If
        Call CheckCorrections
    End If
End Sub
Private Sub Text5_Change()
    If Text5.Text = " " Then
        Text5.Text = "?"
        Sendkeys "{TAB}"
    ElseIf CheckEmpty(Text5, False) Then
        BookPrinterCode = ""
    End If
    Me.Text5.Tag = "M"
End Sub
Private Sub Text5_Validate(Cancel As Boolean)
    Dim SearchString As String
    If CheckEmpty(Text5, False) Then Exit Sub
    SearchString = FixQuote(Text5.Text)
    If rstBookPrinterList.RecordCount = 0 Then
        DisplayError ("No Record in Book Printer Master")
        Cancel = True
        Exit Sub
    Else
        rstBookPrinterList.MoveFirst
    End If
    rstBookPrinterList.Find "[Col0] = '" & RTrim(SearchString) & "'"
    If rstBookPrinterList.EOF Then
        SelectionType = "S"
        BookPrinterCode = ""
        Call LoadSelectionList(rstBookPrinterList, "List of Book Printers...", "Name")
        SearchOrder = 0
        Call DisplaySelectionList(Text5, BookPrinterCode)
        Call CloseForm(FrmSelectionList)
        If CheckEmpty(Text5.Text, False) Then Text5.Text = "?"
        If RTrim(BookPrinterCode) <> "" Then Sendkeys "{TAB}"
        Cancel = True
    Else
        BookPrinterCode = rstBookPrinterList.Fields("Code").Value
        If Me.Text5.Tag = "M" Then
            If Not CheckEmpty(Text5.Text, False) Then Command1_Click
            Me.Text5.Tag = ""
        End If
    End If
End Sub
Private Sub Text6_Change()
    If Text6.Text = " " Then
        Text6.Text = "?"
        Sendkeys "{TAB}"
    ElseIf CheckEmpty(Text6, False) Then
        TitlePrinterCode = ""
    End If
    Me.Text6.Tag = "M"
End Sub
Private Sub Text6_Validate(Cancel As Boolean)
    Dim SearchString As String
    If CheckEmpty(Text6, False) Then Exit Sub
    SearchString = FixQuote(Text6.Text)
    If rstTitlePrinterList.RecordCount = 0 Then
        DisplayError ("No Record in Title Printer Master")
        Cancel = True
        Exit Sub
    Else
        rstTitlePrinterList.MoveFirst
    End If
    rstTitlePrinterList.Find "[Col0] = '" & RTrim(SearchString) & "'"
    If rstTitlePrinterList.EOF Then
        SelectionType = "S"
        TitlePrinterCode = ""
        Call LoadSelectionList(rstTitlePrinterList, "List of Title Printers...", "Name")
        SearchOrder = 0
        Call DisplaySelectionList(Text6, TitlePrinterCode)
        Call CloseForm(FrmSelectionList)
        If CheckEmpty(Text6.Text, False) Then Text6.Text = "?"
        If RTrim(TitlePrinterCode) <> "" Then Sendkeys "{TAB}"
        Cancel = True
    Else
        TitlePrinterCode = rstTitlePrinterList.Fields("Code").Value
        If Me.Text6.Tag = "M" Then
            If Not CheckEmpty(Text6.Text, False) Then Command2_Click
            Me.Text6.Tag = ""
        End If
    End If
End Sub
Private Sub Text7_Change()
    If Text7.Text = " " Then
        Text7.Text = "?"
        Sendkeys "{TAB}"
    ElseIf CheckEmpty(Text7, False) Then
        LaminatorCode = ""
    End If
    Me.Text7.Tag = "M"
End Sub
Private Sub Text7_Validate(Cancel As Boolean)
    Dim SearchString As String
    If CheckEmpty(Text7, False) Then Exit Sub
    SearchString = FixQuote(Text7.Text)
    If rstLaminatorList.RecordCount = 0 Then
        DisplayError ("No Record in Laminator Master")
        Cancel = True
        Exit Sub
    Else
        rstLaminatorList.MoveFirst
    End If
    rstLaminatorList.Find "[Col0] = '" & RTrim(SearchString) & "'"
    If rstLaminatorList.EOF Then
        SelectionType = "S"
        LaminatorCode = ""
        Call LoadSelectionList(rstLaminatorList, "List of Laminators...", "Name")
        SearchOrder = 0
        Call DisplaySelectionList(Text7, LaminatorCode)
        Call CloseForm(FrmSelectionList)
        If CheckEmpty(Text7.Text, False) Then Text7.Text = "?"
        If RTrim(LaminatorCode) <> "" Then Sendkeys "{TAB}"
        Cancel = True
    Else
        LaminatorCode = rstLaminatorList.Fields("Code").Value
        If Me.Text7.Tag = "M" Then
            If Not CheckEmpty(Text7.Text, False) Then Command3_Click
            Me.Text7.Tag = ""
        End If
    End If
End Sub
Private Sub Text8_Change()
    If Text8.Text = " " Then
        Text8.Text = "?"
        Sendkeys "{TAB}"
    ElseIf CheckEmpty(Text8, False) Then
        BinderCode = ""
    End If
    Me.Text8.Tag = "M"
End Sub
Private Sub Text8_Validate(Cancel As Boolean)
    Dim SearchString As String
    If CheckEmpty(Text8, False) Then Exit Sub
    SearchString = FixQuote(Text8.Text)
    If rstBinderList.RecordCount = 0 Then
        DisplayError ("No Record in Binder Master")
        Cancel = True
        Exit Sub
    Else
        rstBinderList.MoveFirst
    End If
    rstBinderList.Find "[Col0] = '" & RTrim(SearchString) & "'"
    If rstBinderList.EOF Then
        SelectionType = "S"
        BinderCode = ""
        Call LoadSelectionList(rstBinderList, "List of Binders...", "Name")
        SearchOrder = 0
        Call DisplaySelectionList(Text8, BinderCode)
        Call CloseForm(FrmSelectionList)
        If CheckEmpty(Text8.Text, False) Then Text8.Text = "?"
        If RTrim(BinderCode) <> "" Then Sendkeys "{TAB}"
        Cancel = True
    Else
        BinderCode = rstBinderList.Fields("Code").Value
        If Me.Text8.Tag = "M" Then
            If Not CheckEmpty(Text8.Text, False) Then Command4_Click
            Me.Text8.Tag = ""
        End If
    End If
End Sub
Private Sub ViewRecord()
    ClearFields
    If rstBookPOList.EOF Then
        If rstBookPOChild05.State = adStateOpen Then
            rstBookPOChild05.Close
        End If
        If rstBookPOChild06.State = adStateOpen Then
            rstBookPOChild06.Close
        End If
        If rstBookPOChild07.State = adStateOpen Then
            rstBookPOChild07.Close
        End If
        If rstBookPOChild08.State = adStateOpen Then
            rstBookPOChild08.Close
        End If
        If rstBookPOChild0801.State = adStateOpen Then
            rstBookPOChild0801.Close
        End If
        Exit Sub
    End If
    FindRecord
    LoadFields
End Sub
Private Sub FindRecord()
    If rstBookPOParent.State = adStateOpen Then
       rstBookPOParent.Close
    End If
    rstBookPOParent.Open "Select * From BookPOParent Where Code = '" & FixQuote(rstBookPOList.Fields("Code").Value) & "'", CxnBookPrintOrder, adOpenKeyset, adLockOptimistic
    If rstBookPOParent.RecordCount = 0 Then
       Call DisplayError("This Record has been deleted by Another User ! Click Ok To Refresh the Recordset")
       Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
    End If
End Sub
Private Sub ClearFields()
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text7.Text = ""
    Text8.Text = ""
    Text9.Text = ""
    Text10.Text = ""
    MhDateInput1.Text = Format(Date, "dd-MM-yyyy")
    MhRealInput1.Text = 0#
    MhRealInput2.Text = 0
    BookPrinterCode = ""
    TitlePrinterCode = ""
    LaminatorCode = ""
    BinderCode = ""
    chkBP.Value = 0
    chkTP.Value = 0
    chkTL.Value = 0
    chkBB.Value = 0
End Sub
Private Sub LoadFields()
    
    If rstBookPOParent.EOF Or rstBookPOParent.BOF Then Exit Sub
    Text2.Text = rstBookPOParent.Fields("Name").Value
    MhDateInput1.Text = Format(rstBookPOParent.Fields("Date").Value, "dd-MM-yyyy")
    BookCode = rstBookPOParent.Fields("Book").Value
    If rstBookList.RecordCount > 0 Then rstBookList.MoveFirst
    rstBookList.Find "[Code] = '" & BookCode & "'"
    If Not rstBookList.EOF Then
        Text3.Text = rstBookList.Fields("Col0").Value
        Text9.Text = rstBookList.Fields("Col1").Value
        If Val(rstBookList.Fields("TwoColorPages").Value) = 0 And Val(rstBookList.Fields("FourColorPages").Value) = 0 Then
            Text10.Text = "1 Color"
        ElseIf Val(rstBookList.Fields("OneColorPages").Value) = 0 And Val(rstBookList.Fields("FourColorPages").Value) = 0 Then
            Text10.Text = "2 Color"
        ElseIf Val(rstBookList.Fields("OneColorPages").Value) = 0 And Val(rstBookList.Fields("TwoColorPages").Value) = 0 Then
            Text10.Text = "4 Color"
        ElseIf Val(rstBookList.Fields("OneColorPages").Value) = 0 And Val(rstBookList.Fields("TwoColorPages").Value) = 0 And Val(rstBookList.Fields("FourColorPages").Value) = 0 Then
            Text10.Text = "6 Color"
        Else
            Text10.Text = "Multi Color"
        End If
        Text4.Text = rstBookList.Fields("SizeName").Value & "/" & IIf(rstBookList.Fields("FormType").Value = "1", "08", IIf(rstBookList.Fields("FormType").Value = "2", "16", IIf(rstBookList.Fields("FormType").Value = "3", "04", IIf(rstBookList.Fields("FormType").Value = "4", "12", IIf(rstBookList.Fields("FormType").Value = "5", "24", IIf(rstBookList.Fields("FormType").Value = "6", "32", IIf(rstBookList.Fields("FormType").Value = "7", "64", "06")))))))
        MhRealInput1.Text = Val(rstBookList.Fields("Forms").Value)
        MhRealInput2.Text = Val(rstBookList.Fields("Pages").Value)
    End If
    
    BookPrinterCode = rstBookPOParent.Fields("BookPrinter").Value
    If rstBookPrinterList.RecordCount > 0 Then rstBookPrinterList.MoveFirst
    rstBookPrinterList.Find "[Code] = '" & BookPrinterCode & "'"
    If Not rstBookPrinterList.EOF Then Text5.Text = rstBookPrinterList.Fields("Col0").Value
    TitlePrinterCode = rstBookPOParent.Fields("TitlePrinter").Value
    If rstTitlePrinterList.RecordCount > 0 Then rstTitlePrinterList.MoveFirst
    rstTitlePrinterList.Find "[Code] = '" & TitlePrinterCode & "'"
    If Not rstTitlePrinterList.EOF Then Text6.Text = rstTitlePrinterList.Fields("Col0").Value
    LaminatorCode = rstBookPOParent.Fields("Laminator").Value
    If rstLaminatorList.RecordCount > 0 Then rstLaminatorList.MoveFirst
    rstLaminatorList.Find "[Code] = '" & LaminatorCode & "'"
    If Not rstLaminatorList.EOF Then Text7.Text = rstLaminatorList.Fields("Col0").Value
    BinderCode = rstBookPOParent.Fields("Binder").Value
    If rstBinderList.RecordCount > 0 Then rstBinderList.MoveFirst
    rstBinderList.Find "[Code] = '" & BinderCode & "'"
    If Not rstBinderList.EOF Then Text8.Text = rstBinderList.Fields("Col0").Value
    chkBP.Value = IIf(rstBookPOParent.Fields("BPODStatus").Value, 1, 0)
    chkTP.Value = IIf(rstBookPOParent.Fields("TPODStatus").Value, 1, 0)
    chkTL.Value = IIf(rstBookPOParent.Fields("TLODStatus").Value, 1, 0)
    chkBB.Value = IIf(rstBookPOParent.Fields("BBODStatus").Value, 1, 0)
    Text5.Tag = "": Text6.Tag = "": Text7.Tag = "": Text8.Tag = ""
    Call LoadOrder(rstBookPOParent.Fields("Code").Value)

End Sub
Private Sub EditRecord()
    On Error GoTo ErrorHandler
    If rstBookPOParent.RecordCount = 0 Then Exit Sub
    If rstBookPOChild05.State = adStateClosed Or rstBookPOChild06.State = adStateClosed Or rstBookPOChild07.State = adStateClosed Or rstBookPOChild08.State = adStateClosed Or rstBookPOChild0801.State = adStateClosed Then
        SSTab1.Tab = 0
        Exit Sub
    End If
    If rstBookPOParent.State = adStateOpen Then rstBookPOParent.Close
    rstBookPOParent.CursorLocation = adUseServer
    rstBookPOParent.Open "Select * From BookPOParent Where Code = '" & FixQuote(rstBookPOList.Fields("Code").Value) & "'", CxnBookPrintOrder, adOpenKeyset, adLockPessimistic
    MdiMainMenu.MousePointer = vbHourglass
    rstBookPOParent.Fields("Printstatus") = "N"
    MdiMainMenu.MousePointer = vbNormal
    AddToList
    Call SetButtons(False)
    SSTab1.TabEnabled(0) = False
    Text2.SetFocus
    blnRecordExist = True
    If AllowTransactionsModification = 0 Then
        LockFields (True)
        Text1.Locked = False:        Text5.Locked = False:        Text6.Locked = False:        Text7.Locked = False:        Text8.Locked = False
    End If
    CxnBookPrintOrder.BeginTrans
    Exit Sub
ErrorHandler:
    If Err.Number = -2147467259 Then
       Call DisplayError("Failed to Edit the record")
    End If
    MdiMainMenu.MousePointer = vbNormal
    SSTab1.Tab = 0
End Sub
Private Sub SaveFields()
    If rstBookPOParent.EOF Or rstBookPOParent.BOF Then Exit Sub
    If Not blnRecordExist Then
        rstBookPOParent.Fields("Code").Value = GenerateCode(CxnBookPrintOrder, "Select Max(Code) From BookPOParent", 6, "0")
        rstBookPOParent.Fields("CreatedBy").Value = UserCode
        rstBookPOParent.Fields("CreatedOn").Value = Now()
        rstBookPOParent.Fields("Recordstatus").Value = "N"
    Else
        rstBookPOParent.Fields("ModifiedBy").Value = UserCode
        rstBookPOParent.Fields("ModifiedOn").Value = Now()
        rstBookPOParent.Fields("Recordstatus").Value = "M"
    End If
    rstBookPOParent.Fields("Name").Value = Pad(Trim(Text2.Text), Space(1), 10, "L")
    rstBookPOParent.Fields("Date").Value = GetDate(MhDateInput1.Text)
    rstBookPOParent.Fields("Book").Value = BookCode
    rstBookPOParent.Fields("BookPrinter").Value = BookPrinterCode
    rstBookPOParent.Fields("TitlePrinter").Value = TitlePrinterCode
    rstBookPOParent.Fields("Laminator").Value = LaminatorCode
    rstBookPOParent.Fields("Binder").Value = BinderCode
    rstBookPOParent.Fields("BPODStatus").Value = chkBP.Value
    rstBookPOParent.Fields("TPODStatus").Value = chkTP.Value
    rstBookPOParent.Fields("TLODStatus").Value = chkTL.Value
    rstBookPOParent.Fields("BBODStatus").Value = chkBB.Value
    rstBookPOParent.Fields("Type").Value = BookPOType
    rstBookPOParent.Fields("PrintStatus").Value = "N"
    
End Sub
Private Sub AddToList()
    On Error Resume Next
    rstBookPOList.MoveFirst
    rstBookPOList.Find "[Code] = '" & rstBookPOParent.Fields("Code").Value & "'"
    If rstBookPOList.EOF Then
       rstBookPOList.AddNew
       rstBookPOList.Fields("Code").Value = rstBookPOParent.Fields("Code").Value
       rstBookPOList.Fields("BPODStatus").Value = 0: rstBookPOList.Fields("TPODStatus").Value = 0: rstBookPOList.Fields("TLODStatus").Value = 0: rstBookPOList.Fields("BBODStatus").Value = 0
    End If
    rstBookPOList.Fields("Name").Value = Pad(rstBookPOParent.Fields("Name").Value, Space(1), 10, "L")
    rstBookPOList.Fields("Date").Value = rstBookPOParent.Fields("Date").Value
    rstBookList.MoveFirst
    rstBookList.Find "[Code] = '" & rstBookPOParent.Fields("Book").Value & "'"
    rstBookPOList.Fields("BookName").Value = Trim(rstBookList.Fields("Col0").Value)
    rstBookPOList.Fields("BoardName").Value = Trim(rstBookList.Fields("BoardName").Value)
    rstBookPOList.Fields("ReceivedQuantity").Value = Val(rstBookPOParent.Fields("ReceivedQuantity").Value)
    rstBookPOList.Fields("BookPrinterName").Value = Trim(Text5.Text)
    rstBookPOList.Fields("TitlePrinterName").Value = Trim(Text6.Text)
    rstBookPOList.Fields("LaminatorName").Value = Trim(Text7.Text)
    rstBookPOList.Fields("BinderName").Value = Trim(Text8.Text)
    rstBookPOList.Fields("BPODStatus").Value = chkBP.Value
    rstBookPOList.Fields("TPODStatus").Value = chkTP.Value
    rstBookPOList.Fields("TLODStatus").Value = chkTL.Value
    rstBookPOList.Fields("BBODStatus").Value = chkBB.Value
    rstBookPOList.Update
    rstBookPOList.Sort = SortOrder & " Asc"
    rstBookPOList.Find "[Code] = '" & rstBookPOParent.Fields("Code").Value & "'"
End Sub
Private Function CheckMandatoryFields() As Boolean
    If CheckEmpty(Text2.Text, False) Then
       DisplayError ("Order No. cannot be blank")
       Text2.SetFocus
       CheckMandatoryFields = True
    ElseIf CheckDuplicate(CxnBookPrintOrder, "BookPOParent", "Code", "[Name]+[Type]", Trim(Text2.Text) & BookPOType, rstBookPOParent.Fields("Code").Value, False) Then
        Text2.SetFocus
        CheckMandatoryFields = True
    ElseIf CheckEmpty(Text3.Text, False) Then
       Text3.SetFocus
       CheckMandatoryFields = True
    ElseIf Not CheckExists(Text3, "Col0", rstBookList, BookCode) Then
        Text3.SetFocus
        CheckMandatoryFields = True
    Else
        If Not CheckEmpty(Text5.Text, False) Then
            If Not CheckExists(Text5, "Col0", rstBookPrinterList, BookPrinterCode) Then
                Text5.SetFocus
                CheckMandatoryFields = True
            End If
        End If
        If Not CheckEmpty(Text6.Text, False) Then
            If Not CheckExists(Text6, "Col0", rstTitlePrinterList, TitlePrinterCode) Then
                Text6.SetFocus
                CheckMandatoryFields = True
            End If
        End If
        If Not CheckEmpty(Text7.Text, False) Then
            If Not CheckExists(Text7, "Col0", rstLaminatorList, LaminatorCode) Then
                Text7.SetFocus
                CheckMandatoryFields = True
            End If
        End If
        If Not CheckEmpty(Text8.Text, False) Then
            If Not CheckExists(Text8, "Col0", rstBinderList, BinderCode) Then
                Text8.SetFocus
                CheckMandatoryFields = True
            End If
        End If
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
Private Sub LoadOrder(ByVal strOrderCode As String)
    On Error GoTo ErrorHandler
    If rstBookPOChild05.State = adStateOpen Then
       rstBookPOChild05.Close
    End If
    rstBookPOChild05.Open "Select * From BookPOChild05 Where Code = '" & strOrderCode & "'", CxnBookPrintOrder, adOpenKeyset, adLockOptimistic
    
    rstBookPOChild05.ActiveConnection = Nothing
    If rstBookPOChild06.State = adStateOpen Then
       rstBookPOChild06.Close
    End If
    rstBookPOChild06.Open "Select * From BookPOChild06 Where Code = '" & strOrderCode & "'", CxnBookPrintOrder, adOpenKeyset, adLockOptimistic
    rstBookPOChild06.ActiveConnection = Nothing
    If rstBookPOChild07.State = adStateOpen Then
       rstBookPOChild07.Close
    End If
    rstBookPOChild07.Open "Select * From BookPOChild07 Where Code = '" & strOrderCode & "'", CxnBookPrintOrder, adOpenKeyset, adLockOptimistic
    rstBookPOChild07.ActiveConnection = Nothing
    If rstBookPOChild08.State = adStateOpen Then
       rstBookPOChild08.Close
    End If
    rstBookPOChild08.Open "Select * From BookPOChild08 Where Code = '" & strOrderCode & "'", CxnBookPrintOrder, adOpenKeyset, adLockOptimistic
    rstBookPOChild08.ActiveConnection = Nothing
    If rstBookPOChild0801.State = adStateOpen Then
       rstBookPOChild0801.Close
    End If
    
    rstBookPOChild0801.Open "Select * From BookPOChild0801 T Where Code = '" & strOrderCode & "'", CxnBookPrintOrder, adOpenKeyset, adLockOptimistic
    rstBookPOChild0801.ActiveConnection = Nothing
    Exit Sub
ErrorHandler:
    DisplayError ("Failed to Load Print Order")
End Sub
Private Function UpdateOrder(ByVal strOption As String, Optional ByVal POType As String) As Boolean
    'On Error GoTo ErrorHandler
    UpdateOrder = True
    If strOption = "D" Then
        CxnBookPrintOrder.Execute "Delete From BookPOChild05 Where Code = '" & rstBookPOParent.Fields("Code").Value & "'"
        CxnBookPrintOrder.Execute "Delete From BookPOChild06 Where Code = '" & rstBookPOParent.Fields("Code").Value & "'"
        CxnBookPrintOrder.Execute "Delete From BookPOChild07 Where Code = '" & rstBookPOParent.Fields("Code").Value & "'"
        CxnBookPrintOrder.Execute "Delete From BookPOChild08 Where Code = '" & rstBookPOParent.Fields("Code").Value & "'"
        CxnBookPrintOrder.Execute "Delete From BookPOChild0801 Where Code = '" & rstBookPOParent.Fields("Code").Value & "'"
    Else
        If POType = "1" And Not CheckEmpty(Text5.Text, False) Then    'Book Printer
       
        CxnBookPrintOrder.Execute "Insert Into BookPOChild05 Values ('" & rstBookPOParent.Fields("Code").Value & "',#" & Format(rstBookPOChild05.Fields("OrderDate").Value, "mm-dd-yyyy") & "#,#" & Format(rstBookPOChild05.Fields("TargetDate").Value, "mm-dd-yyyy") & "#," & IIf(IsNull(rstBookPOChild05.Fields("ExtendDate").Value), "Null", "#" & Format(rstBookPOChild05.Fields("ExtendDate").Value, "mm-dd-yyyy") & "#") & ",'" & rstBookPOChild05.Fields("Processing").Value & "','" & rstBookPOChild05.Fields("Ref").Value & "'," & Val(rstBookPOChild05.Fields("ActualQuantity").Value) & "," & Val(rstBookPOChild05.Fields("BillingQuantity01").Value) & "," & Val(rstBookPOChild05.Fields("BillingQuantity02").Value) & "," & _
                                                           Val(rstBookPOChild05.Fields("Pages1").Value) & "," & Val(rstBookPOChild05.Fields("Forms1").Value) & "," & Val(rstBookPOChild05.Fields("Forms1-¼").Value) & "," & Val(rstBookPOChild05.Fields("Forms1-½").Value) & "," & Val(rstBookPOChild05.Fields("Forms1-1").Value) & ",'" & rstBookPOChild05.Fields("PlateType1").Value & "'," & _
                                                           Val(rstBookPOChild05.Fields("TotalForms1-¼").Value) & "," & Val(rstBookPOChild05.Fields("TotalForms1-½").Value) & "," & Val(rstBookPOChild05.Fields("TotalForms1-1").Value) & "," & Val(rstBookPOChild05.Fields("TotalPlates1-¼").Value) & "," & Val(rstBookPOChild05.Fields("TotalPlates1-½").Value) & "," & Val(rstBookPOChild05.Fields("TotalPlates1-1").Value) & "," & Val(rstBookPOChild05.Fields("PrintRate1").Value) & "," & Val(rstBookPOChild05.Fields("PrintAmount1").Value) & "," & Val(rstBookPOChild05.Fields("PlateRate1").Value) & "," & Val(rstBookPOChild05.Fields("PlateAmount1").Value) & ",'" & _
                                                           rstBookPOChild05.Fields("Paper1").Value & "'," & Val(rstBookPOChild05.Fields("PaperWastage1%").Value) & "," & Val(rstBookPOChild05.Fields("PaperConsumptionOther1").Value) & "," & Val(rstBookPOChild05.Fields("PaperConsumptionSheets1").Value) & "," & Val(rstBookPOChild05.Fields("Forms/Sheet1-1").Value) & "," & Val(rstBookPOChild05.Fields("Forms/Sheet2-1").Value) & "," & _
                                                           Val(rstBookPOChild05.Fields("Pages2").Value) & "," & Val(rstBookPOChild05.Fields("Forms2").Value) & "," & Val(rstBookPOChild05.Fields("Forms2-¼").Value) & "," & Val(rstBookPOChild05.Fields("Forms2-½").Value) & "," & Val(rstBookPOChild05.Fields("Forms2-1").Value) & ",'" & rstBookPOChild05.Fields("PlateType2").Value & "'," & _
                                                           Val(rstBookPOChild05.Fields("TotalForms2-¼").Value) & "," & Val(rstBookPOChild05.Fields("TotalForms2-½").Value) & "," & Val(rstBookPOChild05.Fields("TotalForms2-1").Value) & "," & Val(rstBookPOChild05.Fields("TotalPlates2-¼").Value) & "," & Val(rstBookPOChild05.Fields("TotalPlates2-½").Value) & "," & Val(rstBookPOChild05.Fields("TotalPlates2-1").Value) & "," & Val(rstBookPOChild05.Fields("PrintRate2").Value) & "," & Val(rstBookPOChild05.Fields("PrintAmount2").Value) & "," & Val(rstBookPOChild05.Fields("PlateRate2").Value) & "," & Val(rstBookPOChild05.Fields("PlateAmount2").Value) & ",'" & _
                                                           rstBookPOChild05.Fields("Paper2").Value & "'," & Val(rstBookPOChild05.Fields("PaperWastage2%").Value) & "," & Val(rstBookPOChild05.Fields("PaperConsumptionOther2").Value) & "," & Val(rstBookPOChild05.Fields("PaperConsumptionSheets2").Value) & "," & Val(rstBookPOChild05.Fields("Forms/Sheet1-2").Value) & "," & Val(rstBookPOChild05.Fields("Forms/Sheet2-2").Value) & "," & _
                                                           Val(rstBookPOChild05.Fields("Pages4").Value) & "," & Val(rstBookPOChild05.Fields("Forms4").Value) & "," & Val(rstBookPOChild05.Fields("Forms4-¼").Value) & "," & Val(rstBookPOChild05.Fields("Forms4-½").Value) & "," & Val(rstBookPOChild05.Fields("Forms4-1").Value) & ",'" & rstBookPOChild05.Fields("PlateType4").Value & "'," & _
                                                           Val(rstBookPOChild05.Fields("TotalForms4-¼").Value) & "," & Val(rstBookPOChild05.Fields("TotalForms4-½").Value) & "," & Val(rstBookPOChild05.Fields("TotalForms4-1").Value) & "," & Val(rstBookPOChild05.Fields("TotalPlates4-¼").Value) & "," & Val(rstBookPOChild05.Fields("TotalPlates4-½").Value) & "," & Val(rstBookPOChild05.Fields("TotalPlates4-1").Value) & "," & Val(rstBookPOChild05.Fields("PrintRate4").Value) & "," & Val(rstBookPOChild05.Fields("PrintAmount4").Value) & "," & Val(rstBookPOChild05.Fields("PlateRate4").Value) & "," & Val(rstBookPOChild05.Fields("PlateAmount4").Value) & ",'" & _
                                                           rstBookPOChild05.Fields("Paper4").Value & "'," & Val(rstBookPOChild05.Fields("PaperWastage4%").Value) & "," & Val(rstBookPOChild05.Fields("PaperConsumptionOther4").Value) & "," & Val(rstBookPOChild05.Fields("PaperConsumptionSheets4").Value) & "," & Val(rstBookPOChild05.Fields("Forms/Sheet1-4").Value) & "," & Val(rstBookPOChild05.Fields("Forms/Sheet2-4").Value) & "," & _
                                                           Val(rstBookPOChild05.Fields("TotalPaperConsumption").Value) & ",'" & rstBookPOChild05.Fields("Remarks").Value & "','" & rstBookPOChild05.Fields("BillNo").Value & "'," & IIf(IsNull(rstBookPOChild05.Fields("BillDate").Value), "Null", "#" & Format(rstBookPOChild05.Fields("BillDate").Value, "mm-dd-yyyy") & "#") & "," & Val(rstBookPOChild05.Fields("Adjustment").Value) & "," & Val(rstBookPOChild05.Fields("VAT%").Value) & "," & Val(rstBookPOChild05.Fields("VAT").Value) & "," & Val(rstBookPOChild05.Fields("BillAmount").Value) & "," & Val(rstBookPOChild05.Fields("PaidAmount").Value) & ",'" & rstBookPOChild05.Fields("Status").Value & "','" & rstBookPOChild05.Fields("Narration").Value & "'," & IIf(IsNull(rstBookPOChild05.Fields("BillFeedDate").Value), "Null", "#" & Format(rstBookPOChild05.Fields("BillFeedDate").Value, "mm-dd-yyyy") & "#") & ",'" & rstBookPOChild05.Fields("AdjustmentRemarks").Value & "'," & _
                                                           IIf(IsNull(rstBookPOChild05.Fields("ComputerName").Value), "Null", "'" & rstBookPOChild05.Fields("ComputerName").Value & "'") & "," & IIf(IsNull(rstBookPOChild05!UnitCost), 0, rstBookPOChild05!UnitCost) & ",'" & rstBookPOChild05.Fields("BookStatus").Value & "','" & rstBookPOChild05.Fields("POCode").Value & "','" & rstBookPOChild05.Fields("PlateMaking").Value & "'," & IIf(IsNull(rstBookPOChild05.Fields("PDFSendToProduction").Value), "Null", "#" & Format(rstBookPOChild05.Fields("PDFSendToProduction").Value, "mm-dd-yyyy") & "#") & "," & IIf(IsNull(rstBookPOChild05.Fields("PDFSendToPrinter").Value), "Null", "#" & Format(rstBookPOChild05.Fields("PDFSendToPrinter").Value, "mm-dd-yyyy") & "#") & ")"
                
        ElseIf POType = "2" And Not CheckEmpty(Text6.Text, False) Then    'Title Printer
            CxnBookPrintOrder.Execute "Insert Into BookPOChild06 Values ('" & rstBookPOParent.Fields("Code").Value & "',#" & Format(rstBookPOChild06.Fields("OrderDate").Value, "mm-dd-yyyy") & "#,#" & Format(rstBookPOChild06.Fields("TargetDate").Value, "mm-dd-yyyy") & "#," & IIf(IsNull(rstBookPOChild06.Fields("ExtendDate").Value), "Null", "#" & Format(rstBookPOChild06.Fields("ExtendDate").Value, "mm-dd-yyyy") & "#") & ",'" & rstBookPOChild06.Fields("Processing").Value & "','" & rstBookPOChild06.Fields("Ref").Value & "'," & rstBookPOChild06.Fields("FrontPrintingType").Value & "," & rstBookPOChild06.Fields("BackPrintingType").Value & ",'" & rstBookPOChild06.Fields("PlateType").Value & "'," & Val(rstBookPOChild06.Fields("ActualQuantity").Value) & "," & Val(rstBookPOChild06.Fields("BillingQuantity").Value) & "," & Val(rstBookPOChild06.Fields("Titles/Sheet1").Value) & "," & Val(rstBookPOChild06.Fields("TotalForms").Value) & "," & Val(rstBookPOChild06.Fields("TotalPlates").Value) & "," & _
                                                          Val(rstBookPOChild06.Fields("PrintRate").Value) & "," & Val(rstBookPOChild06.Fields("PrintAmount").Value) & "," & Val(rstBookPOChild06.Fields("PlateRate").Value) & "," & Val(rstBookPOChild06.Fields("PlateAmount").Value) & ",'" & rstBookPOChild06.Fields("Paper").Value & "'," & _
                                                          Val(rstBookPOChild06.Fields("Titles/Sheet2").Value) & "," & Val(rstBookPOChild06.Fields("PaperWastage%").Value) & "," & Val(rstBookPOChild06.Fields("PaperConsumptionOther").Value) & "," & Val(rstBookPOChild06.Fields("PaperConsumptionSheets").Value) & ",'" & rstBookPOChild06.Fields("Remarks").Value & "','" & rstBookPOChild06.Fields("BillNo").Value & "'," & IIf(IsNull(rstBookPOChild06.Fields("BillDate").Value), "Null", "#" & Format(rstBookPOChild06.Fields("BillDate").Value, "mm-dd-yyyy") & "#") & "," & Val(rstBookPOChild06.Fields("Adjustment").Value) & "," & Val(rstBookPOChild06.Fields("VAT%").Value) & "," & Val(rstBookPOChild06.Fields("VAT").Value) & "," & Val(rstBookPOChild06.Fields("BillAmount").Value) & "," & Val(rstBookPOChild06.Fields("PaidAmount").Value) & ",'" & rstBookPOChild06.Fields("Status").Value & "','" & rstBookPOChild06.Fields("Narration").Value & "'," & _
                                                          IIf(IsNull(rstBookPOChild06.Fields("BillFeedDate").Value), "Null", "#" & Format(rstBookPOChild06.Fields("BillFeedDate").Value, "mm-dd-yyyy") & "#") & ",'" & rstBookPOChild06.Fields("AdjustmentRemarks").Value & "'," & IIf(IsNull(rstBookPOChild06.Fields("ComputerName").Value), "Null", "'" & rstBookPOChild06.Fields("ComputerName").Value & "'") & "," & IIf(IsNull(rstBookPOChild06!UnitCost), 0, rstBookPOChild06!UnitCost) & "," & IIf(IsNull(rstBookPOChild06.Fields("CreatedOn").Value), "Null", "#" & Format(rstBookPOChild06.Fields("CreatedOn").Value, "mm-dd-yyyy") & "#") & ",'" & rstBookPOChild06.Fields("PlateMaking").Value & "'," & IIf(IsNull(rstBookPOChild06.Fields("PDFSendToProduction").Value), "Null", "#" & Format(rstBookPOChild06.Fields("PDFSendToProduction").Value, "mm-dd-yyyy") & "#") & "," & _
                                                          IIf(IsNull(rstBookPOChild06.Fields("PDFSendToProduction").Value), "Null", "#" & Format(rstBookPOChild06.Fields("PDFSendToProduction").Value, "mm-dd-yyyy") & "#") & ")"
        ElseIf POType = "3" And Not CheckEmpty(Text7.Text, False) Then    'Laminator
            CxnBookPrintOrder.Execute "Insert Into BookPOChild07 Values ('" & rstBookPOParent.Fields("Code").Value & "',#" & Format(rstBookPOChild07.Fields("OrderDate").Value, "mm-dd-yyyy") & "#,#" & Format(rstBookPOChild07.Fields("TargetDate").Value, "mm-dd-yyyy") & "#," & IIf(IsNull(rstBookPOChild07.Fields("ExtendDate").Value), "Null", "#" & Format(rstBookPOChild07.Fields("ExtendDate").Value, "mm-dd-yyyy") & "#") & ",'" & rstBookPOChild07.Fields("LaminationType").Value & "'," & Val(rstBookPOChild07.Fields("ActualQuantity").Value) & "," & Val(rstBookPOChild07.Fields("QuantityToBinder").Value) & "," & Val(rstBookPOChild07.Fields("QuantityToOffice").Value) & "," & Val(rstBookPOChild07.Fields("BillingQuantity").Value) & "," & Val(rstBookPOChild07.Fields("LaminationRate").Value) & "," & Val(rstBookPOChild07.Fields("LaminationAmount").Value) & ",'" & _
                                                          rstBookPOChild07.Fields("Remarks").Value & "','" & rstBookPOChild07.Fields("BillNo").Value & "'," & IIf(IsNull(rstBookPOChild07.Fields("BillDate").Value), "Null", "#" & Format(rstBookPOChild07.Fields("BillDate").Value, "mm-dd-yyyy") & "#") & "," & Val(rstBookPOChild07.Fields("Adjustment").Value) & "," & Val(rstBookPOChild07.Fields("VAT%").Value) & "," & Val(rstBookPOChild07.Fields("VAT").Value) & "," & Val(rstBookPOChild07.Fields("BillAmount").Value) & "," & Val(rstBookPOChild07.Fields("PaidAmount").Value) & "," & IIf(IsNull(rstBookPOChild07.Fields("BillFeedDate").Value), "Null", "#" & Format(rstBookPOChild07.Fields("BillFeedDate").Value, "mm-dd-yyyy") & "#") & ",'" & rstBookPOChild07.Fields("AdjustmentRemarks").Value & "'," & IIf(IsNull(rstBookPOChild07.Fields("ComputerName").Value), "Null", "'" & rstBookPOChild07.Fields("ComputerName").Value & "'") & "," & _
                                                           IIf(IsNull(rstBookPOChild07.Fields("CreatedOn").Value), "Null", "#" & Format(rstBookPOChild07.Fields("CreatedOn").Value, "mm-dd-yyyy") & "#") & ")"
        ElseIf POType = "4" And Not CheckEmpty(Text8.Text, False) Then    'Binder
            CxnBookPrintOrder.Execute "Insert Into BookPOChild08 Values ('" & rstBookPOParent.Fields("Code").Value & "',#" & Format(rstBookPOChild08.Fields("OrderDate").Value, "mm-dd-yyyy") & "#,#" & Format(rstBookPOChild08.Fields("TargetDate").Value, "mm-dd-yyyy") & "#," & IIf(IsNull(rstBookPOChild08.Fields("ExtendDate").Value), "Null", "#" & Format(rstBookPOChild08.Fields("ExtendDate").Value, "mm-dd-yyyy") & "#") & ",'" & rstBookPOChild08.Fields("BindingType").Value & "'," & Val(rstBookPOChild08.Fields("BindingForms").Value) & "," & Val(rstBookPOChild08.Fields("ExtraForms").Value) & "," & Val(rstBookPOChild08.Fields("ActualQuantity").Value) & "," & Val(rstBookPOChild08.Fields("BillingQuantity").Value) & "," & Val(rstBookPOChild08.Fields("AdjustQuantity").Value) & "," & Val(rstBookPOChild08.Fields("FormFoldRate").Value) & "," & Val(rstBookPOChild08.Fields("FormStitchRate").Value) & "," & Val(rstBookPOChild08.Fields("FormPasteRate").Value) & "," & _
                                                          Val(rstBookPOChild08.Fields("Rate/Book").Value) & "," & Val(rstBookPOChild08.Fields("LooseQty/Box").Value) & "," & Val(rstBookPOChild08.Fields("ExtraLooseQty").Value) & "," & Val(rstBookPOChild08.Fields("TotalLooseQty").Value) & "," & Val(rstBookPOChild08.Fields("Qty/Pkt").Value) & "," & Val(rstBookPOChild08.Fields("TotalPkts").Value) & "," & Val(rstBookPOChild08.Fields("Pkt/Box").Value) & "," & Val(rstBookPOChild08.Fields("TotalBoxes").Value) & "," & _
                                                          Val(rstBookPOChild08.Fields("PktPackRate").Value) & "," & Val(rstBookPOChild08.Fields("BoxPackRate").Value) & "," & Val(rstBookPOChild08.Fields("CartageRate").Value) & ",'" & rstBookPOChild08.Fields("Remarks").Value & "','" & rstBookPOChild08.Fields("BillNo").Value & "'," & IIf(IsNull(rstBookPOChild08.Fields("BillDate").Value), "Null", "#" & Format(rstBookPOChild08.Fields("BillDate").Value, "mm-dd-yyyy") & "#") & "," & Val(rstBookPOChild08.Fields("Adjustment").Value) & "," & Val(rstBookPOChild08.Fields("VAT%").Value) & "," & Val(rstBookPOChild08.Fields("VAT").Value) & "," & Val(rstBookPOChild08.Fields("BillAmount").Value) & "," & Val(rstBookPOChild08.Fields("PaidAmount").Value) & ",'" & rstBookPOChild08.Fields("Status").Value & "','" & rstBookPOChild08.Fields("Narration").Value & "','" & rstBookPOChild08.Fields("DNDetails").Value & "','" & _
                                                          rstBookPOChild08.Fields("CNDetails").Value & "'," & IIf(IsNull(rstBookPOChild08.Fields("BillFeedDate").Value), "Null", "#" & Format(rstBookPOChild08.Fields("BillFeedDate").Value, "mm-dd-yyyy") & "#") & ",'" & rstBookPOChild08.Fields("AdjustmentRemarks").Value & "'," & IIf(IsNull(rstBookPOChild08.Fields("ComputerName").Value), "Null", "'" & rstBookPOChild08.Fields("ComputerName").Value & "'") & "," & IIf(IsNull(rstBookPOChild08!UnitCost), 0, rstBookPOChild08!UnitCost) & ",'" & rstBookPOChild08!BookEdition & "'," & IIf(IsNull(rstBookPOChild08.Fields("AdvanceRecvdDate").Value), "Null", "#" & Format(rstBookPOChild08.Fields("AdvanceRecvdDate").Value, "mm-dd-yyyy") & "#") & ", '" & rstBookPOChild08.Fields("AdvanceCopyRequired").Value & "')"
        ElseIf POType = "0" And Not CheckEmpty(Text8.Text, False) Then 'Material For Binder
            CxnBookPrintOrder.Execute "Insert Into BookPOChild0801 Values ('" & rstBookPOParent.Fields("Code").Value & "','" & rstBookPOChild0801.Fields("Category").Value & "','" & rstBookPOChild0801.Fields("Item").Value & "'," & Val(rstBookPOChild0801.Fields("Quantity").Value) & ")"
        End If
    End If
    Exit Function
'ErrorHandler:
    UpdateOrder = False
End Function
Public Sub FilterRecord(ByVal SrchFor As String, ByVal SrchText As String)
    If SrchFor = "Book" Then
        rstBookPOList.Filter = "[BookName] Like '%" & SrchText & "%'"
    End If
End Sub
Private Sub Command1_Click()
    If CheckEmpty(Text5.Text, False) Or (Not CheckExists(Text5, "Col0", rstBookPrinterList, BookPrinterCode)) Then Exit Sub
    If rstBookPOChild05.RecordCount = 0 Then Call AddRecord(rstBookPOChild05)
    Set FrmBookPOChild05.rstBookPOChild05 = rstBookPOChild05
    FrmBookPOChild05.PrinterCode = BookPrinterCode
    On Error Resume Next
    Load FrmBookPOChild05
    If Err.Number <> 364 Then
        If rstBookPOParent.Fields("BookPrinter").Value <> "" Then
            If BookPOType <> "O" Then
                If blnRecordExist And AllowTransactionsModification = 0 Then
                    If Not CheckEmpty(FrmBookPOChild05.Text8.Text, False) Then
                        Dim O As Object
                        For Each O In FrmBookPOChild05
                                If TypeName(O) = "TextBox" Then
                                    O.Locked = True
                                ElseIf TypeName(O) = "TDBNumber" Then
                                    O.ReadOnly = True
                                ElseIf TypeName(O) = "ComboBox" Then
                                    O.Enabled = False
                                ElseIf TypeName(O) = "TDBDate" Then
                                    O.ReadOnly = True
                                End If
                        Next
                    End If
                End If
            End If
        End If
        FrmBookPOChild05.Show vbModal
    End If
'    If Val(rstBookPOChild05.Fields("ActualQuantity").Value) = 0 Then
'        rstBookPOChild05.Delete
'        rstBookPOChild05.MoveNext
'        If rstBookPOChild05.RecordCount > 0 Then rstBookPOChild05.MoveFirst
'    End If
    On Error GoTo 0
    If AbortPO Then Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
End Sub
Private Sub Command2_Click()
    If CheckEmpty(Text6.Text, False) Or (Not CheckExists(Text6, "Col0", rstTitlePrinterList, TitlePrinterCode)) Then Exit Sub
    If rstBookPOChild06.RecordCount = 0 Then Call AddRecord(rstBookPOChild06)
    Set FrmBookPOChild06.rstBookPOChild06 = rstBookPOChild06
    FrmBookPOChild06.PrinterCode = TitlePrinterCode
    If rstBookPOChild05.RecordCount > 0 Then FrmBookPOChild06.TextPrinterQuantity = Val(rstBookPOChild05.Fields("ActualQuantity").Value)
    On Error Resume Next
    Load FrmBookPOChild06
    If Err.Number <> 364 Then
        If rstBookPOParent.Fields("TitlePrinter").Value <> "" Then
            If BookPOType <> "O" Then
                If blnRecordExist And AllowTransactionsModification = 0 Then
                    If Not CheckEmpty(FrmBookPOChild06.Text8.Text, False) Then
                        Dim O As Object
                        For Each O In FrmBookPOChild06
                                If TypeName(O) = "TextBox" Then
                                    O.Locked = True
                                ElseIf TypeName(O) = "TDBNumber" Then
                                    O.ReadOnly = True
                                ElseIf TypeName(O) = "ComboBox" Then
                                    O.Enabled = False
                                ElseIf TypeName(O) = "TDBDate" Then
                                    O.ReadOnly = True
                                End If
                        Next
                    End If
                End If
            End If
        End If
        FrmBookPOChild06.Show vbModal
    End If
    
'    If Val(rstBookPOChild06.Fields("ActualQuantity").Value) = 0 Then
'        rstBookPOChild06.Delete
'        rstBookPOChild06.MoveNext
'        If rstBookPOChild06.RecordCount > 0 Then rstBookPOChild06.MoveFirst
'    End If

    On Error GoTo 0
    If AbortPO Then Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
End Sub
Private Sub Command3_Click()
    If CheckEmpty(Text7.Text, False) Or (Not CheckExists(Text7, "Col0", rstLaminatorList, LaminatorCode)) Then Exit Sub
    If rstBookPOChild07.RecordCount = 0 Then Call AddRecord(rstBookPOChild07)
    Set FrmBookPOChild07.rstBookPOChild07 = rstBookPOChild07
    FrmBookPOChild07.LaminatorCode = LaminatorCode
    If rstBookPOChild06.RecordCount > 0 Then FrmBookPOChild07.TitlePrinterQuantity = Val(rstBookPOChild06.Fields("ActualQuantity").Value)
    On Error Resume Next
    Load FrmBookPOChild07
    If Err.Number <> 364 Then
        If rstBookPOParent.Fields("Laminator").Value <> "" Then
            If BookPOType <> "O" Then
                If blnRecordExist And AllowTransactionsModification = 0 Then
                    If Not CheckEmpty(FrmBookPOChild07.Text8.Text, False) Then
                        Dim O As Object
                        For Each O In FrmBookPOChild07
                                If TypeName(O) = "TextBox" Then
                                    O.Locked = True
                                ElseIf TypeName(O) = "TDBNumber" Then
                                    O.ReadOnly = True
                                ElseIf TypeName(O) = "TDBDate" Then
                                    O.ReadOnly = True
                                End If
                        Next
                    End If
                End If
            End If
        End If
        FrmBookPOChild07.Show vbModal
    End If
'    If Val(rstBookPOChild07.Fields("ActualQuantity").Value) = 0 Then
'        rstBookPOChild07.Delete
'        rstBookPOChild07.MoveNext
'        If rstBookPOChild07.RecordCount > 0 Then rstBookPOChild07.MoveFirst
'    End If
    On Error GoTo 0
End Sub
Private Sub Command4_Click()

    If CheckEmpty(Text8.Text, False) Or (Not CheckExists(Text8, "Col0", rstBinderList, BinderCode)) Then Exit Sub
    If rstBookPOChild08.RecordCount = 0 Then Call AddRecord(rstBookPOChild08)
    Set FrmBookPOChild08.rstBookPOChild08 = rstBookPOChild08
    FrmBookPOChild08.BinderCode = BinderCode
    If rstBookPOChild06.RecordCount > 0 Then FrmBookPOChild08.BookPrinterQuantity = Val(rstBookPOChild06.Fields("ActualQuantity").Value)
    If rstBookPOChild05.RecordCount > 0 Then FrmBookPOChild08.BookPrinterQuantity = Val(rstBookPOChild05.Fields("ActualQuantity").Value)
    Set FrmBookPOChild0801.rstBookPOChild0801 = rstBookPOChild0801
    FrmBookPOChild0801.BinderCode = BinderCode
    FrmBookPOChild0801.BookCode = BookCode
    FrmBookPOChild0801.OrderCode = CheckNull(rstBookPOParent.Fields("Code").Value)
    FrmBookPOChild08.MhRealInput19.Text = Val(CheckNull(rstBookPOParent.Fields("ReceivedQuantity").Value))
     
    'FrmBookPOChild08.MhRealInput19.Text = Val(rstBookPOParent.Fields("ReceivedQuantity").Value)
      
    On Error Resume Next
    
    Load FrmBookPOChild08
    If Err.Number <> 364 Then
        If rstBookPOParent.Fields("Binder").Value <> "" Then
            If BookPOType <> "O" Then
                If blnRecordExist And AllowTransactionsModification = 0 Then
                    If Not CheckEmpty(FrmBookPOChild08.Text8.Text, False) Then
                        Dim O As Object
                        For Each O In FrmBookPOChild08
                                If TypeName(O) = "TextBox" Then
                                    O.Locked = True
                                ElseIf TypeName(O) = "TDBNumber" Then
                                    O.ReadOnly = True
                                ElseIf TypeName(O) = "TDBDate" Then
                                    O.ReadOnly = True
                                End If
                        Next
                    End If
                End If
            End If
        End If
        FrmBookPOChild08.Show vbModal
    End If
'    If Val(rstBookPOChild08.Fields("ActualQuantity").Value) = 0 Then
'        rstBookPOChild08.Delete
'        rstBookPOChild08.MoveNext
'        If rstBookPOChild08.RecordCount > 0 Then rstBookPOChild08.MoveFirst
'    End If
    On Error GoTo 0
    If AbortPO Then Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
End Sub
Private Sub DisplayMenu()
    Dim menusel As String

    If rstBookPOList.RecordCount = 0 Then Exit Sub
    menusel = DisplayPopupMenu(Me.hwnd, 1)
    Select Case menusel
        Case 1
            GetOrderNoRange (1)
            PrintBookPrintingOrder (rstBookPOList.Fields("Code").Value)
        Case 2
            GetOrderNoRange (2)
            PrintTitlePrintingOrder (rstBookPOList.Fields("Code").Value)
        Case 3
            GetOrderNoRange (3)
            PrintTitleLaminationOrder (rstBookPOList.Fields("Code").Value)
        Case 4
            GetOrderNoRange (4)
            PrintBookBindingOrder (rstBookPOList.Fields("Code").Value)
        Case 5
            PrintBookOrder (rstBookPOList.Fields("Code").Value)
        Case 6
            PrintBookBoxLabel (rstBookPOList.Fields("Code").Value)
    End Select
    If Not (rstBookPOList.EOF Or rstBookPOList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    Text1.SetFocus
End Sub
Public Sub PrintBookPrintingOrder(ByVal OrderCode As String, Optional ByVal Note As String, Optional ByVal OutputType As String)
    Dim oOutlookMsg As Outlook.MailItem, RecordAffected As Integer
    Dim rstPaperStock As New ADODB.Recordset
    Dim sVchCode As String, dDate As Date, sAccountCode As String, sRef As String, sPaperCode As String
    Dim OpBal As Double, Consumption As Double, Sent As Double, Bal As Double, GrandUnitCost As Double
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    If rstBookPOChild05.State = adStateOpen Then rstBookPOChild05.Close
    rstBookPOChild05.Open "SELECT P.Code,Date,BookPrinter,Ref,Paper1,Paper2,Paper4 FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE P.Code='" & rstBookPOList.Fields("Code").Value & "'", CxnDatabase, adOpenKeyset, adLockOptimistic
    If rstBookPOChild05.RecordCount > 0 Then
        sVchCode = rstBookPOChild05.Fields("Code").Value
        dDate = rstBookPOChild05.Fields("Date").Value
        sAccountCode = rstBookPOChild05.Fields("BookPrinter").Value
        
        If rstBookPOChild05.Fields("Paper1").Value <> "" Then sPaperCode = rstBookPOChild05.Fields("Paper1").Value
        If rstBookPOChild05.Fields("Paper2").Value <> "" Then sPaperCode = rstBookPOChild05.Fields("Paper2").Value
        If rstBookPOChild05.Fields("Paper4").Value <> "" Then sPaperCode = rstBookPOChild05.Fields("Paper4").Value
        sRef = rstBookPOChild05.Fields("Ref").Value
        If rstBookPOChild05.State = adStateOpen Then rstBookPOChild05.Close
        
        rstBookPOChild05.Open "SELECT " & _
                              "FORMAT((SELECT OpBalSheets FROM PaperChild WHERE Code=M.Code AND Account='" & sAccountCode & "'),0) As OpBal," & _
                              "FORMAT((SELECT SUM(INT(Quantity)*500+(Quantity-INT(Quantity))*1000) FROM MaterialSVParent P INNER JOIN MaterialSVChild C ON P.Code=C.Code WHERE Category='2' AND Item=M.Code AND Quantity>=0 AND Account='" & sAccountCode & "' AND Date<=#" & GetDate(dDate) & "#),0) As Plus1," & _
                              "FORMAT((SELECT ABS(SUM(FIX(Quantity)*500+(Quantity-FIX(Quantity))*1000)) FROM MaterialSVParent P INNER JOIN MaterialSVChild C ON P.Code=C.Code WHERE Category='2' AND Item=M.Code AND Quantity<0 AND Account='" & sAccountCode & "' AND Date<=#" & GetDate(dDate) & "#),0) As Minus1," & _
                              "FORMAT((SELECT SUM(QuantitySheets) FROM PaperMVParent P INNER JOIN PaperMVChild C ON P.Code=C.Code WHERE Paper=M.Code AND AccountFrom='" & sAccountCode & "' AND Date<=#" & GetDate(dDate) & "#),0) As Minus2," & _
                              "FORMAT((SELECT SUM(QuantitySheets) FROM PaperMVParent P INNER JOIN PaperMVChild C ON P.Code=C.Code WHERE Paper=M.Code AND AccountTo='" & sAccountCode & "' AND Date<=#" & GetDate(dDate) & "#),0) As Plus2," & _
                              "FORMAT((SELECT INT(Quantity)*500+(Quantity-INT(Quantity))*1000 FROM PaperDNParent P INNER JOIN PaperDNChild C ON P.Code=C.Code WHERE P.Account='" & sAccountCode & "' AND C.Paper=M.Code AND Date<=#" & GetDate(dDate) & "#),0) As Minus3," & _
                              "FORMAT((SELECT SUM(PaperConsumptionSheets1) FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE P.Type<>'O' AND LEFT(P.Code,1)<>'*' AND Paper1=M.Code AND BookPrinter='" & sAccountCode & "' AND P.Code<'" & sVchCode & "'),0) As Minus4," & _
                              "FORMAT((SELECT SUM(PaperConsumptionSheets2) FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE P.Type<>'O' AND LEFT(P.Code,1)<>'*' AND Paper2=M.Code AND BookPrinter='" & sAccountCode & "' AND P.Code<'" & sVchCode & "'),0) As Minus5," & _
                              "FORMAT((SELECT SUM(PaperConsumptionSheets4) FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE P.Type<>'O' AND LEFT(P.Code,1)<>'*' AND Paper4=M.Code AND BookPrinter='" & sAccountCode & "' AND P.Code<'" & sVchCode & "'),0) As Minus6," & _
                              "FORMAT((SELECT SUM(PaperConsumptionSheets) FROM BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code WHERE P.Type<>'O' AND LEFT(P.Code,1)<>'*' AND Paper=M.Code AND TitlePrinter='" & sAccountCode & "' AND P.Code<'" & sVchCode & "'),0) As Minus7," & _
                              "FORMAT((SELECT SUM(ROUND(ActualQuantity*C2.Quantity,0)) FROM (BookPOParent P INNER JOIN BookPOChild08 C1 ON P.Code=C1.Code) INNER JOIN BookPOChild0801 C2 ON C1.Code=C2.Code WHERE P.Type<>'O' AND LEFT(P.Code,1)<>'*' AND C2.Category='2' AND C2.Item=M.Code AND BookPrinter='" & sAccountCode & "' AND P.Code<'" & sVchCode & "'),0) As Minus8," & _
                              "FORMAT((SELECT SUM(QuantitySheets) FROM PaperPOParent P INNER JOIN PaperIOChild C ON P.Code=C.Code " & _
                              "WHERE C.Paper=M.Code AND Account='" & sAccountCode & "' AND P.Date<=#" & GetDate(dDate) & "# AND P.Code NOT IN (SELECT Code FROM PaperPOChildRef WHERE Ref='" & sRef & "')),0) As Plus3," & _
                              "FORMAT((SELECT SUM(QuantitySheets) FROM (PaperPOParent P INNER JOIN PaperIOChild C ON P.Code=C.Code) INNER JOIN PaperPOChildRef C1 ON P.Code=C1.Code " & _
                              "WHERE C.Paper=M.Code AND Account='" & sAccountCode & "' AND Date<=#" & GetDate(dDate) & "# AND C1.Ref='" & sRef & "'),0) As PaperSent," & _
                              "FORMAT((SELECT SUM(PaperConsumptionSheets1) FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE P.Type<>'O' AND LEFT(P.Code,1)<>'*' AND Paper1=M.Code AND BookPrinter='" & sAccountCode & "' AND P.Code='" & sVchCode & "'),0) As Consumption1," & _
                              "FORMAT((SELECT SUM(PaperConsumptionSheets2) FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE P.Type<>'O' AND LEFT(P.Code,1)<>'*' AND Paper2=M.Code AND BookPrinter='" & sAccountCode & "' AND P.Code='" & sVchCode & "'),0) As Consumption2," & _
                              "FORMAT((SELECT SUM(PaperConsumptionSheets4) FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE P.Type<>'O' AND LEFT(P.Code,1)<>'*' AND Paper4=M.Code AND BookPrinter='" & sAccountCode & "' AND P.Code='" & sVchCode & "'),0) As Consumption3 " & _
                              "FROM PaperMaster M WHERE Code='" & sPaperCode & "'", CxnDatabase, adOpenKeyset, adLockOptimistic
        OpBal = Val(rstBookPOChild05.Fields("OpBal").Value) + Val(rstBookPOChild05.Fields("Plus1").Value) + Val(rstBookPOChild05.Fields("Plus2").Value) + Val(rstBookPOChild05.Fields("Plus3").Value) - Val(rstBookPOChild05.Fields("Minus1").Value) - Val(rstBookPOChild05.Fields("Minus2").Value) - Val(rstBookPOChild05.Fields("Minus3").Value) - Val(rstBookPOChild05.Fields("Minus4").Value) - Val(rstBookPOChild05.Fields("Minus5").Value) - Val(rstBookPOChild05.Fields("Minus6").Value) - Val(rstBookPOChild05.Fields("Minus7").Value) - Val(rstBookPOChild05.Fields("Minus8").Value)
        Consumption = Val(rstBookPOChild05.Fields("Consumption1").Value) + Val(rstBookPOChild05.Fields("Consumption2").Value) + Val(rstBookPOChild05.Fields("Consumption3").Value)
        
        Sent = Val(rstBookPOChild05.Fields("PaperSent").Value)
        Bal = OpBal + Sent - Consumption
        OpBal = CLng(Fix(OpBal / 500)) + (OpBal Mod 500) / 1000
        Consumption = CLng(Fix(Consumption / 500)) + (Consumption Mod 500) / 1000
        Sent = CLng(Fix(Sent / 500)) + (Sent Mod 500) / 1000
        Bal = CLng(Fix(Bal / 500)) + (Bal Mod 500) / 1000
    
    End If
    
    If rstBookPOChild05.State = adStateOpen Then rstBookPOChild05.Close
    rstBookPOChild05.Open "Select 'BP/" & Right(Year(FinancialYearFrom), 2) + "-" + Right(Year(FinancialYearTo), 2) & "/'+Trim(P.Name) As OrderNo,OrderDate,TargetDate,Processing,(Select Trim(PrintName) From AccountMaster Where Code = P.BookPrinter) As PrinterName,(Select Trim(PrintName) From AccountMaster Where Code = P.Binder) As BinderName,Trim(M.PrintName)+iif(M.Price=0,'',' (Price : Rs. '+Format(M.Price,'0.00')+')') As BookName,M.DuplexPrinting,(Select Trim(PrintName) From GeneralMaster Where Code = M.Board) As BoardName,(Select Trim(PrintName) From GeneralMaster Where Code = M.[Size]) As BookSize,ActualQuantity,BillingQuantity01,BillingQuantity02," & _
                            "Forms1,[Forms1-¼],[Forms1-½],[Forms1-1],[TotalForms1-¼],[TotalForms1-½],[TotalForms1-1],PrintRate1,PrintAmount1,PlateRate1,(Select Trim(PrintName) From PaperMaster Where Code = C.Paper1) As Paper1Name,PlateAmount1,[PaperWastage1%],PaperConsumptionOther1," & _
                            "Forms2,[Forms2-¼],[Forms2-½],[Forms2-1],[TotalForms2-¼],[TotalForms2-½],[TotalForms2-1],PrintRate2,PrintAmount2,PlateRate2,(Select Trim(PrintName) From PaperMaster Where Code = C.Paper2) As Paper2Name,PlateAmount2,[PaperWastage2%],PaperConsumptionOther2," & _
                            "Forms4,[Forms4-¼],[Forms4-½],[Forms4-1],[TotalForms4-¼],[TotalForms4-½],[TotalForms4-1],PrintRate4,PrintAmount4,PlateRate4,(Select Trim(PrintName) From PaperMaster Where Code = C.Paper4) As Paper4Name,PlateAmount4,[PaperWastage4%],PaperConsumptionOther4," & _
                            "TotalPaperConsumption,Adjustment,[VAT%],VAT,BillAmount,C.Remarks,(Select Trim(eMail) From AccountMaster Where Code = P.BookPrinter) As EMailID,M.Narration,M.OneColorPages  As Pages1, M.TwoColorPages  As Pages2 ,M.FourColorPages  As Pages4,C.UnitCost,C.ExtendDate,C.BillNo,C.BillDate,BookStatus,M.Code As BookCode,C.Paper1,C.Paper2,C.Paper4,C.PlateMaking,C.PDFSendToProduction,C.PDFSendToPrinter  From (BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN BookMaster M ON P.Book=M.Code Where P.Code = '" & OrderCode & "'", CxnDatabase, adOpenKeyset, adLockOptimistic
    If FrmBookPOPrintUtility.PrintUtility = False Then
            If PrintReport = 3 Then 'Title Lamination
                If rstBookPOChild07.State = adStateOpen Then rstBookPOChild07.Close
                rstBookPOChild07.Open "Select 'TL/" & Right(Year(FinancialYearFrom), 2) + "-" + Right(Year(FinancialYearTo), 2) & "/'+Trim(P.Name) As OrderNo,OrderDate,TargetDate,(Select Trim(PrintName) From AccountMaster Where Code = P.Laminator) As LaminatorName,(Select Trim(PrintName) From AccountMaster Where Code = P.TitlePrinter) As PrinterName,(Select Trim(PrintName) From AccountMaster Where Code = P.Binder) As BinderName,Trim(M.PrintName)+iif(M.Price=0,'',' (Price : Rs. '+Format(M.Price,'0.00')+')') As BookName,(Select Trim(PrintName) From GeneralMaster Where Code = M.Board) As BoardName,(Select Trim(PrintName) From GeneralMaster Where Code = M.[Size]) As BookSize,(Select Trim(PrintName) From GeneralMaster Where Code = C.LaminationType) As LaminationType,ActualQuantity,QuantityToBinder,QuantityToOffice,BillingQuantity,LaminationRate,LaminationAmount,Adjustment,[VAT%],VAT,BillAmount,C.Remarks,C.BillNo,C.BillDate,(Select Trim(eMail) From AccountMaster Where Code = P.Laminator) As EMailID " & _
                                      " From BookPOParent P,BookPOChild07 C,BookMaster M Where P.Code = C.Code And P.Book = M.Code And P.Code = '" & OrderCode & "' And P.Laminator <> ''", CxnDatabase, adOpenKeyset, adLockOptimistic
            End If
            If PrintReport = 2 Then 'Title Printing
                If rstBookPOChild06.State = adStateOpen Then rstBookPOChild06.Close
                rstBookPOChild06.Open "Select 'TP/" & Right(Year(FinancialYearFrom), 2) + "-" + Right(Year(FinancialYearTo), 2) & "/'+Trim(P.Name) As OrderNo,OrderDate,TargetDate,Processing,(Select Trim(PrintName) From AccountMaster Where Code = P.TitlePrinter) As PrinterName,(Select Trim(PrintName) From AccountMaster Where Code = P.Laminator) As LaminatorName,Trim(M.PrintName)+iif(M.Price=0,'',' (Price : Rs. '+Format(M.Price,'0.00')+')') As BookName,(Select Trim(PrintName) From GeneralMaster Where Code = M.Board) As BoardName,(Select Trim(PrintName) From GeneralMaster Where Code = M.[Size]) As BookSize,C.FrontPrintingType,C.BackPrintingType,ActualQuantity,BillingQuantity,(Select Trim(PrintName) From PaperMaster Where Code = C.Paper) As PaperName,[PaperWastage%],PaperConsumptionOther,PrintRate,PlateRate,PrintAmount,PlateAmount,Adjustment,[VAT%],VAT,BillAmount,C.Remarks,M.Narration,(Select Trim(eMail) From AccountMaster Where Code = P.TitlePrinter) As EMailID " & _
                                      ",C.UnitCost,C.ExtendDate,C.BillNo,C.BillDate,C.TotalForms,C.PlateMaking,C.PDFSendToProduction,C.PDFSendToPrinter From BookPOParent P,BookPOChild06 C,BookMaster M Where P.Code = C.Code And P.Book = M.Code And P.Code = '" & OrderCode & "' And P.TitlePrinter <> ''", CxnDatabase, adOpenKeyset, adLockOptimistic
            End If
            
            If PrintReport = 4 Then 'Book Binding
                If rstBookPOChild08.State = adStateOpen Then rstBookPOChild08.Close
                rstBookPOChild08.Open "Select 'BB/" & Right(Year(FinancialYearFrom), 2) + "-" + Right(Year(FinancialYearTo), 2) & "/'+Trim(P.Name) As OrderNo,OrderDate,TargetDate,(Select Trim(PrintName) From AccountMaster Where Code = P.Binder) As BinderName,(Select Trim(PrintName) From AccountMaster Where Code = P.Laminator) As LaminatorName,(Select Trim(PrintName) From AccountMaster Where Code = P.BookPrinter) As PrinterName,Trim(M.PrintName)+iif(M.Price=0,'',' (Price : Rs. '+Format(M.Price,'0.00')+')') As BookName,(Select Trim(PrintName) From GeneralMaster Where Code = M.Board) As BoardName,(Select Trim(PrintName) From GeneralMaster Where Code = M.[Size]) As BookSize,(Select Trim(PrintName) From GeneralMaster Where Code = C.BindingType) As BindingType,ActualQuantity,BillingQuantity,M.Forms,C.BindingForms,ExtraForms,FormFoldRate,FormStitchRate,FormPasteRate,[Rate/Book],[TotalPkts],[TotalBoxes],PktPackRate,BoxPackRate,CartageRate,Adjustment,[VAT%],[VAT]," & _
                                      "BillAmount,C.Remarks,(Select Trim(eMail) From AccountMaster Where Code = P.Binder) As EMailID,M.Narration,M.OneColorPages as Page1, M.TwoColorPages as Page2,M.FourColorPages as Page4,C.[LooseQty/Box],ExtraLooseQty,TotalLooseQty,C.[Qty/Pkt],C.[Pkt/Box],ISBN,C.UnitCost,C.ExtendDate,C.BookEdition,C.BillNo,C.BillDate,C.AdvanceRecvdDate, C.AdvanceCopyRequired,(Select Warehouse1  FROM PrintPVChild Where Code in(Select Ref From BookPOChild05 WHERE Code In(Select Code From BookPOParent Where Name=P.Name)) And Book in(Select  Book From BookPOParent Where Name=P.Name))As Warehouse1,(Select Warehouse2  FROM PrintPVChild Where Code in(Select Ref From BookPOChild05 WHERE Code In(Select Code From BookPOParent Where Name=P.Name)) And Book in(Select  Book From BookPOParent Where Name=P.Name))As Warehouse2," & _
                                      "(Select Warehouse3  FROM PrintPVChild Where Code in(Select Ref From BookPOChild05 WHERE Code In(Select Code From BookPOParent Where Name=P.Name)) And Book in(Select  Book From BookPOParent Where Name=P.Name))As Warehouse3 From BookPOParent P,BookPOChild08 C,BookMaster M Where P.Code = C.Code And P.Book = M.Code And P.Code = '" & OrderCode & "' And P.Binder <> ''", CxnDatabase, adOpenKeyset, adLockOptimistic
            
            
             End If
     End If
     
          
    Screen.MousePointer = vbNormal
    If rstBookPOChild05.RecordCount = 0 Then On Error GoTo 0: Exit Sub
    rptBookPrintingOrder.Section20.Suppress = True
    rptBookPrintingOrder.Section21.Suppress = True
    rptBookPrintingOrder.Section22.Suppress = True
    rptBookPrintingOrder.Section15.Suppress = True
    
    If rstCompanyMaster.State = adStateClosed Then rstCompanyMaster.Open "Select PrintName, Address1, Address2, Address3, Address4, Phone, Fax, EMail, Website From CompanyMaster", CxnDatabase, adOpenKeyset, adLockReadOnly
    rptBookPrintingOrder.Text2.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptBookPrintingOrder.Text3.SetText Trim(rstCompanyMaster.Fields("Address1").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address2").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address3").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address4").Value)
    rptBookPrintingOrder.Text24.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & Space(1) & "Fax : " & Trim(rstCompanyMaster.Fields("Fax").Value) & Space(1) & "e-Mail : " & Trim(rstCompanyMaster.Fields("EMail").Value)
    rptBookPrintingOrder.Text25.SetText " (" & Trim(NumberToWords(rstBookPOChild05.Fields("BillAmount").Value, True)) & ")"
    rptBookPrintingOrder.Text27.SetText "for " & Trim(rstBookPOChild05.Fields("PrinterName").Value)
    rptBookPrintingOrder.Text28.SetText "for " & Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptBookPrintingOrder.Text35.SetText Trim(COMPANY_CIN)
    rptBookPrintingOrder.Text45.SetText "Plate Making   : " & Trim(rstBookPOChild05.Fields("PlateMaking").Value)
    
    rptBookPrintingOrder.Text48.SetText "Op Bal : " & Format(OpBal, "#0.000") + " Consumed : " + Format(Consumption, "#0.000") + " Sent : " + Format(Sent, "#0.000") + " Bal : " + Format(Bal, "#0.000")
    
    GrandUnitCost = GrandUnitCost + Val(rstBookPOChild05.Fields("UnitCost").Value)
    

'    rptBookPrintingOrder.Text45.SetText "Plate Making   : " & GetPlateMaking(rstBookPOChild05.Fields("BookCode").Value, OrderCode, rstBookPOChild05.Fields("OrderDate").Value, "05", IIf(Val(rstBookPOChild05.Fields("Pages1").Value) > 0, "1", IIf(Val(rstBookPOChild05.Fields("Pages2").Value) > 0, "2", "4")))
'    Dim PaperSendBalanceArray() As String
'    PaperSendBalanceArray = Split(GetPaperSend_Balance(OrderCode, "C"), "#")
'    rptBookPrintingOrder.Text46.SetText "Paper Send     : " & PaperSendBalanceArray(0)
'    rptBookPrintingOrder.Text47.SetText "Paper Balance  : " & PaperSendBalanceArray(1)

'    Dim PaperSendBalanceArray10() As String
'    PaperSendBalanceArray10 = Split(GetPaperSend_Balance(GetPOPrevious(PaperCode_For_Balance), "P"), "#")
'    rptBookPrintingOrder.Text44.SetText "Prev Paper Bal : " & PaperSendBalanceArray10(1)
       
     
     rptBookPrintingOrder.Database.SetDataSource rstBookPOChild05, 3, 1
     
     If FrmBookPOPrintUtility.PrintUtility = False Then
         If PrintReport = 3 Then    'Title Lamination
            rptBookPrintingOrder.Subreport3_Text2.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
            rptBookPrintingOrder.Subreport3_Text3.SetText Trim(rstCompanyMaster.Fields("Address1").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address2").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address3").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address4").Value)
            rptBookPrintingOrder.Subreport3_Text24.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & Space(1) & "Fax : " & Trim(rstCompanyMaster.Fields("Fax").Value) & Space(1) & "e-Mail : " & Trim(rstCompanyMaster.Fields("EMail").Value)
            rptBookPrintingOrder.Subreport3_Text25.SetText " (" & Trim(NumberToWords(rstBookPOChild07.Fields("BillAmount").Value, True)) & ")"
            rptBookPrintingOrder.Subreport3_Text27.SetText "for " & Trim(rstBookPOChild07.Fields("LaminatorName").Value)
            rptBookPrintingOrder.Subreport3_Text28.SetText "for " & Trim(rstCompanyMaster.Fields("PrintName").Value)
            rptBookPrintingOrder.Subreport3_Text31.SetText Trim(COMPANY_CIN)
            rptBookPrintingOrder.Subreport3.OpenSubreport.Database.SetDataSource rstBookPOChild07, 3, 1
            rptBookPrintingOrder.Section22.Suppress = False
            rptBookPrintingOrder.Subreport3.OpenSubreport.Database.SetDataSource rstBookPOChild07, 3, 1
          End If
          
          If PrintReport = 2 Then 'Title Printing
            rptBookPrintingOrder.Subreport2_Text2.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
            rptBookPrintingOrder.Subreport2_Text3.SetText Trim(rstCompanyMaster.Fields("Address1").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address2").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address3").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address4").Value)
            rptBookPrintingOrder.Subreport2_Text24.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & Space(1) & "Fax : " & Trim(rstCompanyMaster.Fields("Fax").Value) & Space(1) & "e-Mail : " & Trim(rstCompanyMaster.Fields("EMail").Value)
            rptBookPrintingOrder.Subreport2_Text25.SetText " (" & Trim(NumberToWords(rstBookPOChild06.Fields("BillAmount").Value, True)) & ")"
            rptBookPrintingOrder.Subreport2_Text27.SetText "for " & Trim(rstBookPOChild06.Fields("PrinterName").Value)
            rptBookPrintingOrder.Subreport2_Text28.SetText "for " & Trim(rstCompanyMaster.Fields("PrintName").Value)
            rptBookPrintingOrder.Subreport2_Text33.SetText Trim(COMPANY_CIN) 'Add here company cin no
            
            rptBookPrintingOrder.Subreport2_Text38.SetText "Plate Making : " & Trim(rstBookPOChild06.Fields("PlateMaking").Value)
            
            Dim PaperSendBalanceArray2() As String
            PaperSendBalanceArray2 = Split(GetPaperSend_Balance(OrderCode, "C"), "#")
            rptBookPrintingOrder.Subreport2_Text39.SetText "Paper Send : " & PaperSendBalanceArray2(0)
            rptBookPrintingOrder.Subreport2_Text40.SetText "Paper Bal : " & PaperSendBalanceArray2(1)
      
            
            rptBookPrintingOrder.Subreport2.OpenSubreport.Database.SetDataSource rstBookPOChild06, 3, 1
            rptBookPrintingOrder.Section21.Suppress = False
            rptBookPrintingOrder.Subreport2.OpenSubreport.Database.SetDataSource rstBookPOChild06, 3, 1
         End If
         
         If PrintReport = 4 Then 'Book Binding
            rptBookPrintingOrder.Subreport1_Text2.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
            rptBookPrintingOrder.Subreport1_Text3.SetText Trim(rstCompanyMaster.Fields("Address1").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address2").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address3").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address4").Value)
            rptBookPrintingOrder.Subreport1_Text24.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & Space(1) & "Fax : " & Trim(rstCompanyMaster.Fields("Fax").Value) & Space(1) & "e-Mail : " & Trim(rstCompanyMaster.Fields("EMail").Value)
            rptBookPrintingOrder.Subreport1_Text25.SetText " (" & Trim(NumberToWords(rstBookPOChild08.Fields("BillAmount").Value, True)) & ")"
            rptBookPrintingOrder.Subreport1_Text27.SetText "for " & Trim(rstBookPOChild05.Fields("PrinterName").Value)
            rptBookPrintingOrder.Subreport1_Text28.SetText "for " & Trim(rstCompanyMaster.Fields("PrintName").Value)
            
            rptBookPrintingOrder.Subreport1_Text46.SetText "Total Unit Cost: " & Format(GrandUnitCost + Val(rstBookPOChild08.Fields("UnitCost").Value), "#0.000")
            
            'GrandUnitCost = GrandUnitCost + Val(rstBookPOChild08.Fields("UnitCost").Value)
            
            rptBookPrintingOrder.Subreport1_Text29.SetText Trim(COMPANY_CIN)
            rptBookPrintingOrder.Section20.Suppress = False
            rptBookPrintingOrder.Subreport1.OpenSubreport.Database.SetDataSource rstBookPOChild08, 3, 1
        End If
    End If
        EMailID = rstBookPOChild05.Fields("EMailID").Value
        Attachment = Trim(rstBookPOChild05.Fields("OrderNo").Value)
        Attachment = Mid(Attachment, InStr(4, Attachment, "/") + 1)
        Message = "Dear Sir,<Br>Please find attached herewith PO #" & Trim(rstBookPOChild05.Fields("OrderNo").Value) & " for doing the needful at your end. An early finish of the job assigned to you will be highly appreciated.<Br>Kindly acknowledge the receipt of mail and confirm the date of completion of job.<Br><Br>" & IIf(Note = "", "", "<b><u>Note : " & Note & "</b></u><Br><Br>") & Trim(rstCompanyMaster.Fields("PrintName").Value) & "<Br>Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & "<Br>E-Mail : <a HRef='mailto:" & Trim(rstCompanyMaster.Fields("EMail").Value) & "'>" & Trim(rstCompanyMaster.Fields("EMail").Value) & "</a>"
        
    If OutputTo = "S" Then
        FrmReportViewer.EMailID = EMailID
        FrmReportViewer.Subject = "Book Printing Order #" & Trim(rstBookPOChild05.Fields("OrderNo").Value) + " Book : " + Trim(rstBookPOChild05.Fields("BookName").Value)
        FrmReportViewer.Attachment = Attachment
        FrmReportViewer.Message = Message
        Set FrmReportViewer.Report = rptBookPrintingOrder
        FrmReportViewer.Show vbModal
    Else
        If rstBookPOList.State = adStateClosed Then
            If EMailID = "" Or OutputType = "P" Then
                rptBookPrintingOrder.PrintOut False   ' Print Report Without Prompt
            Else
                rptBookPrintingOrder.ExportOptions.FormatType = crEFTPortableDocFormat    ' Set the Export Format As .Pdf
                rptBookPrintingOrder.ExportOptions.DestinationType = crEDTDiskFile
                rptBookPrintingOrder.ExportOptions.DiskFileName = App.Path & "\Report\" & Attachment & ".Pdf"
                rptBookPrintingOrder.Export False
                rstBookPOChild05.MoveFirst
                Set oOutlookMsg = oOutlook.CreateItem(olMailItem)
                
                With oOutlookMsg
                    .To = EMailID
                    .Subject = "Book Printing Order #" & Trim(rstBookPOChild05.Fields("OrderNo").Value) + " Book : " + Trim(rstBookPOChild05.Fields("BookName").Value)
                    .HTMLBody = "<Font Face='Calibri' Size='3'>" & Message & "</a>" & "</Font>"
                    .Attachments.Add (App.Path & "\Report\" & Attachment & ".Pdf")
                    .Importance = olImportanceHigh
                    .ReadReceiptRequested = True
                    .Send
                    If Err.Number = 0 Then CxnDatabase.Execute "UPDATE BookPOParent SET BPODStatus=1 WHERE Code='" & OrderCode & "'", RecordAffected
                    If RecordAffected = 0 Then DisplayError ("Failed to update EMail Flag (Book Print Order)")
                End With
                
                Set oOutlookMsg = Nothing
            End If
        Else
            rptBookPrintingOrder.PrintOut
        End If
    End If
    Set rptBookPrintingOrder = Nothing
    Set rptTitleLaminationOrder = Nothing
    Set rptTitlePrintingOrder = Nothing
    Set rptBookBindingOrder = Nothing
    If rstBookPOList.State = adStateClosed Then Call CloseRecordset(rstCompanyMaster): Call CloseRecordset(rstBookPOChild05): Call CloseRecordset(rstBookPOChild06): Call CloseRecordset(rstBookPOChild07): Call CloseRecordset(rstBookPOChild08)
    On Error GoTo 0
End Sub
Public Sub PrintTitlePrintingOrder(ByVal OrderCode As String, Optional ByVal Note As String, Optional ByVal OutputType As String)
    Dim oOutlookMsg As Outlook.MailItem, RecordAffected As Integer
    Dim sVchCode As String, dDate As Date, sAccountCode As String, sRef As String, sPaperCode As String
    Dim OpBal As Double, Consumption As Double, Sent As Double, Bal As Double
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    If rstBookPOChild06.State = adStateOpen Then rstBookPOChild06.Close
    rstBookPOChild06.Open "SELECT P.Code,Date,TitlePrinter,Ref,Paper FROM BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code WHERE P.Code='" & rstBookPOList.Fields("Code").Value & "'", CxnDatabase, adOpenKeyset, adLockOptimistic
    If rstBookPOChild06.RecordCount > 0 Then
        sVchCode = rstBookPOChild06.Fields("Code").Value
        dDate = rstBookPOChild06.Fields("Date").Value
        sAccountCode = rstBookPOChild06.Fields("TitlePrinter").Value
        If rstBookPOChild06.Fields("Paper").Value <> "" Then sPaperCode = rstBookPOChild06.Fields("Paper").Value
        sRef = rstBookPOChild06.Fields("Ref").Value
        If rstBookPOChild06.State = adStateOpen Then rstBookPOChild06.Close
        rstBookPOChild06.Open "SELECT " & _
                              "FORMAT((SELECT OpBalSheets FROM PaperChild WHERE Code=M.Code AND Account='" & sAccountCode & "'),0) As OpBal," & _
                              "FORMAT((SELECT SUM(INT(Quantity)*500+(Quantity-INT(Quantity))*1000) FROM MaterialSVParent P INNER JOIN MaterialSVChild C ON P.Code=C.Code WHERE Category='2' AND Item=M.Code AND Quantity>=0 AND Account='" & sAccountCode & "' AND Date<=#" & GetDate(dDate) & "#),0) As Plus1," & _
                              "FORMAT((SELECT ABS(SUM(FIX(Quantity)*500+(Quantity-FIX(Quantity))*1000)) FROM MaterialSVParent P INNER JOIN MaterialSVChild C ON P.Code=C.Code WHERE Category='2' AND Item=M.Code AND Quantity<0 AND Account='" & sAccountCode & "' AND Date<=#" & GetDate(dDate) & "#),0) As Minus1," & _
                              "FORMAT((SELECT SUM(QuantitySheets) FROM PaperMVParent P INNER JOIN PaperMVChild C ON P.Code=C.Code WHERE Paper=M.Code AND AccountFrom='" & sAccountCode & "' AND Date<=#" & GetDate(dDate) & "#),0) As Minus2," & _
                              "FORMAT((SELECT SUM(QuantitySheets) FROM PaperMVParent P INNER JOIN PaperMVChild C ON P.Code=C.Code WHERE Paper=M.Code AND AccountTo='" & sAccountCode & "' AND Date<=#" & GetDate(dDate) & "#),0) As Plus2," & _
                              "FORMAT((SELECT INT(Quantity)*500+(Quantity-INT(Quantity))*1000 FROM PaperDNParent P INNER JOIN PaperDNChild C ON P.Code=C.Code WHERE P.Account='" & sAccountCode & "' AND C.Paper=M.Code AND Date<=#" & GetDate(dDate) & "#),0) As Minus3," & _
                              "FORMAT((SELECT SUM(PaperConsumptionSheets1) FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE P.Type<>'O' AND LEFT(P.Code,1)<>'*' AND Paper1=M.Code AND BookPrinter='" & sAccountCode & "' AND P.Code<'" & sVchCode & "'),0) As Minus4," & _
                              "FORMAT((SELECT SUM(PaperConsumptionSheets2) FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE P.Type<>'O' AND LEFT(P.Code,1)<>'*' AND Paper2=M.Code AND BookPrinter='" & sAccountCode & "' AND P.Code<'" & sVchCode & "'),0) As Minus5," & _
                              "FORMAT((SELECT SUM(PaperConsumptionSheets4) FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE P.Type<>'O' AND LEFT(P.Code,1)<>'*' AND Paper4=M.Code AND BookPrinter='" & sAccountCode & "' AND P.Code<'" & sVchCode & "'),0) As Minus6," & _
                              "FORMAT((SELECT SUM(PaperConsumptionSheets) FROM BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code WHERE P.Type<>'O' AND LEFT(P.Code,1)<>'*' AND Paper=M.Code AND TitlePrinter='" & sAccountCode & "' AND P.Code<'" & sVchCode & "'),0) As Minus7," & _
                              "FORMAT((SELECT SUM(ROUND(ActualQuantity*C2.Quantity,0)) FROM (BookPOParent P INNER JOIN BookPOChild08 C1 ON P.Code=C1.Code) INNER JOIN BookPOChild0801 C2 ON C1.Code=C2.Code WHERE P.Type<>'O' AND LEFT(P.Code,1)<>'*' AND C2.Category='2' AND C2.Item=M.Code AND BookPrinter='" & sAccountCode & "' AND P.Code<'" & sVchCode & "'),0) As Minus8," & _
                              "FORMAT((SELECT SUM(QuantitySheets) FROM PaperPOParent P INNER JOIN PaperIOChild C ON P.Code=C.Code " & _
                              "WHERE C.Paper=M.Code AND Account='" & sAccountCode & "' AND P.Date<=#" & GetDate(dDate) & "# AND P.Code NOT IN (SELECT Code FROM PaperPOChildRef WHERE Ref='" & sRef & "')),0) As Plus3," & _
                              "FORMAT((SELECT SUM(QuantitySheets) FROM (PaperPOParent P INNER JOIN PaperIOChild C ON P.Code=C.Code) INNER JOIN PaperPOChildRef C1 ON P.Code=C1.Code " & _
                              "WHERE C.Paper=M.Code AND Account='" & sAccountCode & "' AND Date<=#" & GetDate(dDate) & "# AND C1.Ref='" & sRef & "'),0) As PaperSent," & _
                              "FORMAT((SELECT SUM(PaperConsumptionSheets) FROM BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code WHERE P.Type<>'O' AND LEFT(P.Code,1)<>'*' AND Paper=M.Code AND TitlePrinter='" & sAccountCode & "' AND P.Code='" & sVchCode & "'),0) As Consumption " & _
                              "FROM PaperMaster M WHERE Code='" & sPaperCode & "'", CxnDatabase, adOpenKeyset, adLockOptimistic
        OpBal = Val(rstBookPOChild06.Fields("OpBal").Value) + Val(rstBookPOChild06.Fields("Plus1").Value) + Val(rstBookPOChild06.Fields("Plus2").Value) + Val(rstBookPOChild06.Fields("Plus3").Value) - Val(rstBookPOChild06.Fields("Minus1").Value) - Val(rstBookPOChild06.Fields("Minus2").Value) - Val(rstBookPOChild06.Fields("Minus3").Value) - Val(rstBookPOChild06.Fields("Minus4").Value) - Val(rstBookPOChild06.Fields("Minus5").Value) - Val(rstBookPOChild06.Fields("Minus6").Value) - Val(rstBookPOChild06.Fields("Minus7").Value) - Val(rstBookPOChild06.Fields("Minus8").Value)
        Consumption = Val(rstBookPOChild06.Fields("Consumption").Value)
        Sent = Val(rstBookPOChild06.Fields("PaperSent").Value)
        Bal = OpBal + Sent - Consumption
        OpBal = CLng(Fix(OpBal / 500)) + (OpBal Mod 500) / 1000
        Consumption = CLng(Fix(Consumption / 500)) + (Consumption Mod 500) / 1000
        Sent = CLng(Fix(Sent / 500)) + (Sent Mod 500) / 1000
        Bal = CLng(Fix(Bal / 500)) + (Bal Mod 500) / 1000
    End If
    If rstBookPOChild06.State = adStateOpen Then rstBookPOChild06.Close
    rstBookPOChild06.Open "Select 'TP/" & Right(Year(FinancialYearFrom), 2) + "-" + Right(Year(FinancialYearTo), 2) & "/'+Trim(P.Name) As OrderNo,OrderDate,TargetDate,Processing,(Select Trim(PrintName) From AccountMaster Where Code = P.TitlePrinter) As PrinterName,(Select Trim(PrintName) From AccountMaster Where Code = P.Laminator) As LaminatorName,Trim(M.PrintName)+iif(M.Price=0,'',' (Price : Rs. '+Format(M.Price,'0.00')+')') As BookName,(Select Trim(PrintName) From GeneralMaster Where Code = M.Board) As BoardName,(Select Trim(PrintName) From GeneralMaster Where Code = M.[Size]) As BookSize,C.FrontPrintingType,C.BackPrintingType,ActualQuantity,BillingQuantity,(Select Trim(PrintName) From PaperMaster Where Code = C.Paper) As PaperName,[PaperWastage%],PaperConsumptionOther,PrintRate,PlateRate,PrintAmount,PlateAmount,Adjustment,[VAT%],VAT,BillAmount,C.Remarks,M.Narration,(Select Trim(eMail) From AccountMaster Where Code = P.TitlePrinter) As EMailID " & _
                          ",C.UnitCost,C.ExtendDate,C.BillNo,C.BillDate,C.TotalForms,C.PlateMaking From BookPOParent P,BookPOChild06 C,BookMaster M Where P.Code = C.Code And P.Book = M.Code And P.Code = '" & OrderCode & "' And P.TitlePrinter <> ''", CxnDatabase, adOpenKeyset, adLockOptimistic
    If FrmBookPOPrintUtility.PrintUtility = False Then
        If PrintReport = 1 Then 'Book Printing
        If rstBookPOChild05.State = adStateOpen Then rstBookPOChild05.Close
           rstBookPOChild05.Open "Select 'BP/" & Right(Year(FinancialYearFrom), 2) + "-" + Right(Year(FinancialYearTo), 2) & "/'+Trim(P.Name) As OrderNo,OrderDate,TargetDate,Processing,(Select Trim(PrintName) From AccountMaster Where Code = P.BookPrinter) As PrinterName,(Select Trim(PrintName) From AccountMaster Where Code = P.Binder) As BinderName,Trim(M.PrintName)+iif(M.Price=0,'',' (Price : Rs. '+Format(M.Price,'0.00')+')') As BookName,M.DuplexPrinting,(Select Trim(PrintName) From GeneralMaster Where Code = M.Board) As BoardName,(Select Trim(PrintName) From GeneralMaster Where Code = M.[Size]) As BookSize,ActualQuantity,BillingQuantity01,BillingQuantity02," & _
                                "Forms1,[Forms1-¼],[Forms1-½],[Forms1-1],[TotalForms1-¼],[TotalForms1-½],[TotalForms1-1],PrintRate1,PrintAmount1,PlateRate1,(Select Trim(PrintName) From PaperMaster Where Code = C.Paper1) As Paper1Name,PlateAmount1,[PaperWastage1%],PaperConsumptionOther1," & _
                                "Forms2,[Forms2-¼],[Forms2-½],[Forms2-1],[TotalForms2-¼],[TotalForms2-½],[TotalForms2-1],PrintRate2,PrintAmount2,PlateRate2,(Select Trim(PrintName) From PaperMaster Where Code = C.Paper2) As Paper2Name,PlateAmount2,[PaperWastage2%],PaperConsumptionOther2," & _
                                "Forms4,[Forms4-¼],[Forms4-½],[Forms4-1],[TotalForms4-¼],[TotalForms4-½],[TotalForms4-1],PrintRate4,PrintAmount4,PlateRate4,(Select Trim(PrintName) From PaperMaster Where Code = C.Paper4) As Paper4Name,PlateAmount4,[PaperWastage4%],PaperConsumptionOther4," & _
                                "TotalPaperConsumption,Adjustment,[VAT%],VAT,BillAmount,C.Remarks,(Select Trim(eMail) From AccountMaster Where Code = P.BookPrinter) As EMailID,M.Narration,M.OneColorPages  As Pages1, M.TwoColorPages  As Pages2 ,M.FourColorPages  As Pages4,C.UnitCost,C.ExtendDate,C.BillNo,C.BillDate,M.Code As BookCode,C.Paper1,C.Paper2,C.Paper4,C.PlateMaking  From (BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN BookMaster M ON P.Book=M.Code Where P.Code = '" & OrderCode & "'", CxnDatabase, adOpenKeyset, adLockOptimistic
        End If
        If PrintReport = 4 Then 'Book Binding
            If rstBookPOChild08.State = adStateOpen Then rstBookPOChild08.Close
            rstBookPOChild08.Open "Select 'BB/" & Right(Year(FinancialYearFrom), 2) + "-" + Right(Year(FinancialYearTo), 2) & "/'+Trim(P.Name) As OrderNo,OrderDate,TargetDate,(Select Trim(PrintName) From AccountMaster Where Code = P.Binder) As BinderName,(Select Trim(PrintName) From AccountMaster Where Code = P.Laminator) As LaminatorName,(Select Trim(PrintName) From AccountMaster Where Code = P.BookPrinter) As PrinterName,Trim(M.PrintName)+iif(M.Price=0,'',' (Price : Rs. '+Format(M.Price,'0.00')+')') As BookName,(Select Trim(PrintName) From GeneralMaster Where Code = M.Board) As BoardName,(Select Trim(PrintName) From GeneralMaster Where Code = M.[Size]) As BookSize,(Select Trim(PrintName) From GeneralMaster Where Code = C.BindingType) As BindingType,ActualQuantity,BillingQuantity,M.Forms,C.BindingForms,ExtraForms,FormFoldRate,FormStitchRate,FormPasteRate,[Rate/Book],[TotalPkts],[TotalBoxes],PktPackRate,BoxPackRate,CartageRate,Adjustment,[VAT%],[VAT]," & _
                                  "BillAmount,C.Remarks,(Select Trim(eMail) From AccountMaster Where Code = P.Binder) As EMailID,M.Narration,M.OneColorPages as Page1, M.TwoColorPages as Page2,M.FourColorPages as Page4,C.[LooseQty/Box],ExtraLooseQty,TotalLooseQty,C.[Qty/Pkt],C.[Pkt/Box],ISBN,C.UnitCost,C.ExtendDate,C.BookEdition,C.BillNo,C.BillDate, C.AdvanceCopyRequired As CopyReq,(Select Warehouse1  FROM PrintPVChild Where Code in(Select Ref From BookPOChild05 WHERE Code In(Select Code From BookPOParent Where Name=P.Name)) And Book in(Select  Book From BookPOParent Where Name=P.Name))As Warehouse1,(Select Warehouse2  FROM PrintPVChild Where Code in(Select Ref From BookPOChild05 WHERE Code In(Select Code From BookPOParent Where Name=P.Name)) And Book in(Select  Book From BookPOParent Where Name=P.Name))As Warehouse2," & _
                                  "(Select Warehouse3  FROM PrintPVChild Where Code in(Select Ref From BookPOChild05 WHERE Code In(Select Code From BookPOParent Where Name=P.Name)) And Book in(Select  Book From BookPOParent Where Name=P.Name))As Warehouse3 From BookPOParent P,BookPOChild08 C,BookMaster M Where P.Code = C.Code And P.Book = M.Code And P.Code = '" & OrderCode & "' And P.Binder <> ''", CxnDatabase, adOpenKeyset, adLockOptimistic
        End If
        If PrintReport = 3 Then 'Title Lamination
            If rstBookPOChild07.State = adStateOpen Then rstBookPOChild07.Close
            rstBookPOChild07.Open "Select 'TL/" & Right(Year(FinancialYearFrom), 2) + "-" + Right(Year(FinancialYearTo), 2) & "/'+Trim(P.Name) As OrderNo,OrderDate,TargetDate,(Select Trim(PrintName) From AccountMaster Where Code = P.Laminator) As LaminatorName,(Select Trim(PrintName) From AccountMaster Where Code = P.TitlePrinter) As PrinterName,(Select Trim(PrintName) From AccountMaster Where Code = P.Binder) As BinderName,Trim(M.PrintName)+iif(M.Price=0,'',' (Price : Rs. '+Format(M.Price,'0.00')+')') As BookName,(Select Trim(PrintName) From GeneralMaster Where Code = M.Board) As BoardName,(Select Trim(PrintName) From GeneralMaster Where Code = M.[Size]) As BookSize,(Select Trim(PrintName) From GeneralMaster Where Code = C.LaminationType) As LaminationType,ActualQuantity,QuantityToBinder,QuantityToOffice,BillingQuantity,LaminationRate,LaminationAmount,Adjustment,[VAT%],VAT,BillAmount,C.Remarks,C.BillNo,C.BillDate,(Select Trim(eMail) From AccountMaster Where Code = P.Laminator) As EMailID " & _
                                  " From BookPOParent P,BookPOChild07 C,BookMaster M Where P.Code = C.Code And P.Book = M.Code And P.Code = '" & OrderCode & "' And P.Laminator <> ''", CxnDatabase, adOpenKeyset, adLockOptimistic
        End If
    End If
    
    Screen.MousePointer = vbNormal
    If rstBookPOChild06.RecordCount = 0 Then On Error GoTo 0: Exit Sub
    
    rptTitlePrintingOrder.Section15.Suppress = True
    rptTitlePrintingOrder.Section18.Suppress = True
    rptTitlePrintingOrder.Section19.Suppress = True
    
    If rstCompanyMaster.State = adStateClosed Then rstCompanyMaster.Open "Select PrintName, Address1, Address2, Address3, Address4, Phone, Fax, EMail, Website From CompanyMaster", CxnDatabase, adOpenKeyset, adLockReadOnly
    rptTitlePrintingOrder.Text2.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptTitlePrintingOrder.Text3.SetText Trim(rstCompanyMaster.Fields("Address1").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address2").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address3").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address4").Value)
    rptTitlePrintingOrder.Text24.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & Space(1) & "Fax : " & Trim(rstCompanyMaster.Fields("Fax").Value) & Space(1) & "e-Mail : " & Trim(rstCompanyMaster.Fields("EMail").Value)
    rptTitlePrintingOrder.Text25.SetText " (" & Trim(NumberToWords(rstBookPOChild06.Fields("BillAmount").Value, True)) & ")"
    rptTitlePrintingOrder.Text27.SetText "for " & Trim(rstBookPOChild06.Fields("PrinterName").Value)
    rptTitlePrintingOrder.Text28.SetText "for " & Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptTitlePrintingOrder.Text33.SetText Trim(COMPANY_CIN) 'Add here company cin no
'    Dim PaperSendBalanceArray() As String
'    PaperSendBalanceArray = Split(GetPaperSend_Balance(OrderCode, "C"), "#")
'    rptTitlePrintingOrder.Text38.SetText "Paper Send : " & PaperSendBalanceArray(0)
'    rptTitlePrintingOrder.Text40.SetText "Paper Bal : " & PaperSendBalanceArray(1)
    rptTitlePrintingOrder.Text39.SetText "Plate Making   : " & Trim(rstBookPOChild06.Fields("PlateMaking").Value)
    
    rptTitlePrintingOrder.Text38.SetText "Op Bal : " & Format(OpBal, "#0.000") + " Consumed : " + Format(Consumption, "#0.000") + " Sent : " + Format(Sent, "#0.000") + " Bal : " + Format(Bal, "#0.000")
    
    rptTitlePrintingOrder.Database.SetDataSource rstBookPOChild06, 3, 1
    
    If FrmBookPOPrintUtility.PrintUtility = False Then
        If PrintReport = 3 Then    'Title Lamination
            rptTitlePrintingOrder.Subreport1_Text2.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
            rptTitlePrintingOrder.Subreport1_Text3.SetText Trim(rstCompanyMaster.Fields("Address1").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address2").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address3").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address4").Value)
            rptTitlePrintingOrder.Subreport1_Text24.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & Space(1) & "Fax : " & Trim(rstCompanyMaster.Fields("Fax").Value) & Space(1) & "e-Mail : " & Trim(rstCompanyMaster.Fields("EMail").Value)
            rptTitlePrintingOrder.Subreport1_Text25.SetText " (" & Trim(NumberToWords(rstBookPOChild07.Fields("BillAmount").Value, True)) & ")"
            rptTitlePrintingOrder.Subreport1_Text27.SetText "for " & Trim(rstBookPOChild07.Fields("LaminatorName").Value)
            rptTitlePrintingOrder.Subreport1_Text28.SetText "for " & Trim(rstCompanyMaster.Fields("PrintName").Value)
            rptTitlePrintingOrder.Subreport1_Text31.SetText Trim(COMPANY_CIN)
            rptTitlePrintingOrder.Section15.Suppress = False
            rptTitlePrintingOrder.Subreport1.OpenSubreport.Database.SetDataSource rstBookPOChild07, 3, 1
          End If
         If PrintReport = 1 Then 'Book Printing
            rptTitlePrintingOrder.Subreport2_Text2.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
            rptTitlePrintingOrder.Subreport2_Text3.SetText Trim(rstCompanyMaster.Fields("Address1").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address2").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address3").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address4").Value)
            rptTitlePrintingOrder.Subreport2_Text24.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & Space(1) & "Fax : " & Trim(rstCompanyMaster.Fields("Fax").Value) & Space(1) & "e-Mail : " & Trim(rstCompanyMaster.Fields("EMail").Value)
            rptTitlePrintingOrder.Subreport2_Text25.SetText " (" & Trim(NumberToWords(rstBookPOChild05.Fields("BillAmount").Value, True)) & ")"
            rptTitlePrintingOrder.Subreport2_Text27.SetText "for " & Trim(rstBookPOChild05.Fields("PrinterName").Value)
            rptTitlePrintingOrder.Subreport2_Text28.SetText "for " & Trim(rstCompanyMaster.Fields("PrintName").Value)
            rptTitlePrintingOrder.Subreport2_Text35.SetText Trim(COMPANY_CIN)
            rptTitlePrintingOrder.Section18.Suppress = False
            rptTitlePrintingOrder.Subreport2_Section20.Suppress = True
            rptTitlePrintingOrder.Subreport2_Section21.Suppress = True
            rptTitlePrintingOrder.Subreport2_Section22.Suppress = True
            rptTitlePrintingOrder.Subreport2_Text44.SetText "Plate Making   : " & Trim(rstBookPOChild05.Fields("PlateMaking").Value)
            
'            Dim PaperSendBalanceArray2() As String
'            PaperSendBalanceArray2 = Split(GetPaperSend_Balance(OrderCode, "C"), "#")
'            rptTitlePrintingOrder.Subreport2_Text45.SetText "Paper Send     : " & PaperSendBalanceArray2(0)
'            rptTitlePrintingOrder.Subreport2_Text46.SetText "Paper Balance  : " & PaperSendBalanceArray2(1)
            
            'Dim PaperSendBalanceArray11() As String
            'PaperSendBalanceArray11 = Split(GetPaperSend_Balance(GetPOPrevious(PaperCode_For_Balance), "P"), "#")
            'rptTitlePrintingOrder.Subreport2_Text47.SetText "Prev Paper Bal : " & PaperSendBalanceArray11(1)
            
            rptTitlePrintingOrder.Subreport2.OpenSubreport.Database.SetDataSource rstBookPOChild05, 3, 1
        
        
        End If
        If PrintReport = 4 Then 'Book Binding
            rptTitlePrintingOrder.Subreport3_Text2.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
            rptTitlePrintingOrder.Subreport3_Text3.SetText Trim(rstCompanyMaster.Fields("Address1").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address2").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address3").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address4").Value)
            rptTitlePrintingOrder.Subreport3_Text24.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & Space(1) & "Fax : " & Trim(rstCompanyMaster.Fields("Fax").Value) & Space(1) & "e-Mail : " & Trim(rstCompanyMaster.Fields("EMail").Value)
            rptTitlePrintingOrder.Subreport3_Text25.SetText " (" & Trim(NumberToWords(rstBookPOChild08.Fields("BillAmount").Value, True)) & ")"
            rptTitlePrintingOrder.Subreport3_Text27.SetText "for " & Trim(rstBookPOChild08.Fields("BinderName").Value)
            rptTitlePrintingOrder.Subreport3_Text28.SetText "for " & Trim(rstCompanyMaster.Fields("PrintName").Value)
            rptTitlePrintingOrder.Subreport3_Text29.SetText Trim(COMPANY_CIN) 'Add here company cin no
            rptTitlePrintingOrder.Section19.Suppress = False
            rptTitlePrintingOrder.Subreport3.OpenSubreport.Database.SetDataSource rstBookPOChild08, 3, 1
        End If
    End If
    EMailID = rstBookPOChild06.Fields("EMailID").Value
    Attachment = Trim(rstBookPOChild06.Fields("OrderNo").Value)
    Attachment = Mid(Attachment, InStr(4, Attachment, "/") + 1)
    Message = "Dear Sir,<Br>Please find attached herewith PO #" & Trim(rstBookPOChild06.Fields("OrderNo").Value) & " for doing the needful at your end. An early finish of the job assigned to you will be highly appreciated.<Br>Kindly acknowledge the receipt of mail and confirm the date of completion of job.<Br><Br>" & IIf(Note = "", "", "<b><u>Note : " & Note & "</b></u><Br><Br>") & Trim(rstCompanyMaster.Fields("PrintName").Value) & "<Br>Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & "<Br>E-Mail : <a HRef='mailto:" & Trim(rstCompanyMaster.Fields("EMail").Value) & "'>" & Trim(rstCompanyMaster.Fields("EMail").Value) & "</a>"
    
    If OutputTo = "S" Then
        FrmReportViewer.EMailID = EMailID
        FrmReportViewer.Subject = "Title Printing Order #" & Trim(rstBookPOChild06.Fields("OrderNo").Value) + " Book : " + Trim(rstBookPOChild06.Fields("BookName").Value)
        FrmReportViewer.Attachment = Attachment
        FrmReportViewer.Message = Message
        Set FrmReportViewer.Report = rptTitlePrintingOrder
        FrmReportViewer.Show vbModal
    Else
        If rstBookPOList.State = adStateClosed Then
            If EMailID = "" Or OutputType = "P" Then
                rptTitlePrintingOrder.PrintOut False   ' Print Report Without Prompt
            Else
                rptTitlePrintingOrder.ExportOptions.FormatType = crEFTPortableDocFormat    ' Set the Export Format As .Pdf
                rptTitlePrintingOrder.ExportOptions.DestinationType = crEDTDiskFile
                rptTitlePrintingOrder.ExportOptions.DiskFileName = App.Path & "\Report\" & Attachment & ".Pdf"
                rptTitlePrintingOrder.Export False
                rstBookPOChild06.MoveFirst
                Set oOutlookMsg = oOutlook.CreateItem(olMailItem)
                With oOutlookMsg
                    .To = EMailID
                    .Subject = "Title Printing Order #" & Trim(rstBookPOChild06.Fields("OrderNo").Value) + " Book : " + Trim(rstBookPOChild06.Fields("BookName").Value)
                    .HTMLBody = "<Font Face='Calibri' Size='3'>" & Message & "</a>" & "</Font>"
                    .Attachments.Add (App.Path & "\Report\" & Attachment & ".Pdf")
                    .Importance = olImportanceHigh
                    .ReadReceiptRequested = True
                    .Send
                    If Err.Number = 0 Then CxnDatabase.Execute "UPDATE BookPOParent SET TPODStatus=1 WHERE Code='" & OrderCode & "'", RecordAffected
                    If RecordAffected = 0 Then DisplayError ("Failed to update EMail Flag (Title Print Order)")
                End With
                Set oOutlookMsg = Nothing
            End If
        Else
            rptTitlePrintingOrder.PrintOut
        End If
    End If
    Set rptTitlePrintingOrder = Nothing
     Set rptTitleLaminationOrder = Nothing
    Set rptBookPrintingOrder = Nothing
    Set rptBookBindingOrder = Nothing
    If rstBookPOList.State = adStateClosed Then Call CloseRecordset(rstCompanyMaster): Call CloseRecordset(rstBookPOChild06): Call CloseRecordset(rstBookPOChild05): Call CloseRecordset(rstBookPOChild07): Call CloseRecordset(rstBookPOChild08)
    On Error GoTo 0
End Sub


Public Sub PrintTitleLaminationOrder(ByVal OrderCode As String, Optional ByVal Note As String, Optional ByVal OutputType As String)
    Dim oOutlookMsg As Outlook.MailItem, RecordAffected As Integer
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    If rstBookPOChild07.State = adStateOpen Then rstBookPOChild07.Close
    rstBookPOChild07.Open "Select 'TL/" & Right(Year(FinancialYearFrom), 2) + "-" + Right(Year(FinancialYearTo), 2) & "/'+Trim(P.Name) As OrderNo,OrderDate,TargetDate,(Select Trim(PrintName) From AccountMaster Where Code = P.Laminator) As LaminatorName,(Select Trim(PrintName) From AccountMaster Where Code = P.TitlePrinter) As PrinterName,(Select Trim(PrintName) From AccountMaster Where Code = P.Binder) As BinderName,Trim(M.PrintName)+iif(M.Price=0,'',' (Price : Rs. '+Format(M.Price,'0.00')+')') As BookName,(Select Trim(PrintName) From GeneralMaster Where Code = M.Board) As BoardName,(Select Trim(PrintName) From GeneralMaster Where Code = M.[Size]) As BookSize,(Select Trim(PrintName) From GeneralMaster Where Code = C.LaminationType) As LaminationType,ActualQuantity,QuantityToBinder,QuantityToOffice,BillingQuantity,LaminationRate,LaminationAmount,Adjustment,[VAT%],VAT,BillAmount,C.Remarks,C.BillNo,C.BillDate,(Select Trim(eMail) From AccountMaster Where Code = P.Laminator) As EMailID " & _
                                            " From BookPOParent P,BookPOChild07 C,BookMaster M Where P.Code = C.Code And P.Book = M.Code And P.Code = '" & OrderCode & "' And P.Laminator <> ''", CxnDatabase, adOpenKeyset, adLockOptimistic
    If FrmBookPOPrintUtility.PrintUtility = False Then
        If PrintReport = 2 Then 'Title Printing
            If rstBookPOChild06.State = adStateOpen Then rstBookPOChild06.Close
            rstBookPOChild06.Open "Select 'TP/" & Right(Year(FinancialYearFrom), 2) + "-" + Right(Year(FinancialYearTo), 2) & "/'+Trim(P.Name) As OrderNo,OrderDate,TargetDate,Processing,(Select Trim(PrintName) From AccountMaster Where Code = P.TitlePrinter) As PrinterName,(Select Trim(PrintName) From AccountMaster Where Code = P.Laminator) As LaminatorName,Trim(M.PrintName)+iif(M.Price=0,'',' (Price : Rs. '+Format(M.Price,'0.00')+')') As BookName,(Select Trim(PrintName) From GeneralMaster Where Code = M.Board) As BoardName,(Select Trim(PrintName) From GeneralMaster Where Code = M.[Size]) As BookSize,C.FrontPrintingType,C.BackPrintingType,ActualQuantity,BillingQuantity,(Select Trim(PrintName) From PaperMaster Where Code = C.Paper) As PaperName,[PaperWastage%],PaperConsumptionOther,PrintRate,PlateRate,PrintAmount,PlateAmount,Adjustment,[VAT%],VAT,BillAmount,C.Remarks,M.Narration,(Select Trim(eMail) From AccountMaster Where Code = P.TitlePrinter) As EMailID " & _
                                  ",C.UnitCost,C.ExtendDate,C.BillNo,C.BillDate,C.TotalForms,C.PlateMaking From BookPOParent P,BookPOChild06 C,BookMaster M Where P.Code = C.Code And P.Book = M.Code And P.Code = '" & OrderCode & "' And P.TitlePrinter <> ''", CxnDatabase, adOpenKeyset, adLockOptimistic
        End If
    End If
    
    Screen.MousePointer = vbNormal
    If rstBookPOChild07.RecordCount = 0 Then On Error GoTo 0: Exit Sub
    rptTitleLaminationOrder.Section15.Suppress = True
    If rstCompanyMaster.State = adStateClosed Then rstCompanyMaster.Open "Select PrintName, Address1, Address2, Address3, Address4, Phone, Fax, EMail, Website From CompanyMaster", CxnDatabase, adOpenKeyset, adLockReadOnly
    rptTitleLaminationOrder.Text2.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptTitleLaminationOrder.Text3.SetText Trim(rstCompanyMaster.Fields("Address1").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address2").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address3").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address4").Value)
    rptTitleLaminationOrder.Text24.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & Space(1) & "Fax : " & Trim(rstCompanyMaster.Fields("Fax").Value) & Space(1) & "e-Mail : " & Trim(rstCompanyMaster.Fields("EMail").Value)
    rptTitleLaminationOrder.Text25.SetText " (" & Trim(NumberToWords(rstBookPOChild07.Fields("BillAmount").Value, True)) & ")"
    rptTitleLaminationOrder.Text27.SetText "for " & Trim(rstBookPOChild07.Fields("LaminatorName").Value)
    rptTitleLaminationOrder.Text28.SetText "for " & Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptTitleLaminationOrder.Text31.SetText Trim(COMPANY_CIN)
    rptTitleLaminationOrder.Database.SetDataSource rstBookPOChild07, 3, 1
    If FrmBookPOPrintUtility.PrintUtility = False Then
        If PrintReport = 2 Then
            rptTitleLaminationOrder.Subreport1_Text2.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
            rptTitleLaminationOrder.Subreport1_Text3.SetText Trim(rstCompanyMaster.Fields("Address1").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address2").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address3").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address4").Value)
            rptTitleLaminationOrder.Subreport1_Text24.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & Space(1) & "Fax : " & Trim(rstCompanyMaster.Fields("Fax").Value) & Space(1) & "e-Mail : " & Trim(rstCompanyMaster.Fields("EMail").Value)
            rptTitleLaminationOrder.Subreport1_Text25.SetText " (" & Trim(NumberToWords(rstBookPOChild06.Fields("BillAmount").Value, True)) & ")"
            rptTitleLaminationOrder.Subreport1_Text27.SetText "for " & Trim(rstBookPOChild06.Fields("PrinterName").Value)
            rptTitleLaminationOrder.Subreport1_Text28.SetText "for " & Trim(rstCompanyMaster.Fields("PrintName").Value)
            rptTitleLaminationOrder.Subreport1_Text33.SetText Trim(COMPANY_CIN)
            
             rptTitleLaminationOrder.Subreport1_Text38.SetText "Plate Making   : " & Trim(rstBookPOChild06.Fields("PlateMaking").Value)
            
            
            Dim PaperSendBalanceArray() As String
            PaperSendBalanceArray = Split(GetPaperSend_Balance(OrderCode, "C"), "#")
            rptTitleLaminationOrder.Subreport1_Text39.SetText "Paper Send : " & PaperSendBalanceArray(0)
            rptTitleLaminationOrder.Subreport1_Text40.SetText "Paper Bal : " & PaperSendBalanceArray(1)
            
           
            
            
            
            rptTitleLaminationOrder.Section15.Suppress = False
            rptTitleLaminationOrder.Subreport1.OpenSubreport.Database.SetDataSource rstBookPOChild06, 3, 1
            
         End If
     End If
   
    EMailID = rstBookPOChild07.Fields("EMailID").Value
    Attachment = Trim(rstBookPOChild07.Fields("OrderNo").Value)
    Attachment = Mid(Attachment, InStr(4, Attachment, "/") + 1)
    Message = "Dear Sir,<Br>Please find attached herewith PO #" & Trim(rstBookPOChild07.Fields("OrderNo").Value) & " for doing the needful at your end. An early finish of the job assigned to you will be highly appreciated.<Br>Kindly acknowledge the receipt of mail and confirm the date of completion of job.<Br><Br>" & IIf(Note = "", "", "<b><u>Note : " & Note & "</b></u><Br><Br>") & Trim(rstCompanyMaster.Fields("PrintName").Value) & "<Br>Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & "<Br>E-Mail : <a HRef='mailto:" & Trim(rstCompanyMaster.Fields("EMail").Value) & "'>" & Trim(rstCompanyMaster.Fields("EMail").Value) & "</a>"
    If OutputTo = "S" Then
        FrmReportViewer.EMailID = EMailID
        FrmReportViewer.Subject = "Title Lamination Order #" & Trim(rstBookPOChild07.Fields("OrderNo").Value) + " Book : " + Trim(rstBookPOChild07.Fields("BookName").Value)
        FrmReportViewer.Attachment = Attachment
        FrmReportViewer.Message = Message
        Set FrmReportViewer.Report = rptTitleLaminationOrder
        FrmReportViewer.Show vbModal
    Else
        If rstBookPOList.State = adStateClosed Then
            If EMailID = "" Or OutputType = "P" Then
                rptTitleLaminationOrder.PrintOut False   ' Print Report Without Prompt
            Else
                rptTitleLaminationOrder.ExportOptions.FormatType = crEFTPortableDocFormat    ' Set the Export Format As .Pdf
                rptTitleLaminationOrder.ExportOptions.DestinationType = crEDTDiskFile
                rptTitleLaminationOrder.ExportOptions.DiskFileName = App.Path & "\Report\" & Attachment & ".Pdf"
                rptTitleLaminationOrder.Export False
                rstBookPOChild07.MoveFirst
                Set oOutlookMsg = oOutlook.CreateItem(olMailItem)
                With oOutlookMsg
                    .To = EMailID
                    .Subject = "Title Lamination Order #" & Trim(rstBookPOChild07.Fields("OrderNo").Value) + " Book : " + Trim(rstBookPOChild07.Fields("BookName").Value)
                    .HTMLBody = "<Font Face='Calibri' Size='3'>" & Message & "</a>" & "</Font>"
                    .Attachments.Add (App.Path & "\Report\" & Attachment & ".Pdf")
                    .Importance = olImportanceHigh
                    .ReadReceiptRequested = True
                    .Send
                    If Err.Number = 0 Then CxnDatabase.Execute "UPDATE BookPOParent SET TLODStatus=1 WHERE Code='" & OrderCode & "'", RecordAffected
                    If RecordAffected = 0 Then DisplayError ("Failed to update EMail Flag (Lamination Order)")
                End With
                Set oOutlookMsg = Nothing
            End If
        Else
            rptTitleLaminationOrder.PrintOut
        End If
    End If
    Set rptTitleLaminationOrder = Nothing
    Set rptTitlePrintingOrder = Nothing
    Set rptBookBindingOrder = Nothing
    Set rptBookPrintingOrder = Nothing
    
    If rstBookPOList.State = adStateClosed Then Call CloseRecordset(rstCompanyMaster): Call CloseRecordset(rstBookPOChild07): Call CloseRecordset(rstBookPOChild06)
    On Error GoTo 0
End Sub

Public Sub PrintBookBindingOrder(ByVal OrderCode As String, Optional ByVal Note As String, Optional ByVal OutputType As String)
    Dim oOutlookMsg As Outlook.MailItem, RecordAffected As Integer
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    
    If rstBookPOChild08.State = adStateOpen Then rstBookPOChild08.Close
    
    rstBookPOChild08.Open "Select 'BB/" & Right(Year(FinancialYearFrom), 2) + "-" + Right(Year(FinancialYearTo), 2) & "/'+Trim(P.Name) As OrderNo,OrderDate,TargetDate,(Select Trim(PrintName) From AccountMaster Where Code = P.Binder) As BinderName,(Select Trim(PrintName) From AccountMaster Where Code = P.Laminator) As LaminatorName,(Select Trim(PrintName) From AccountMaster Where Code = P.BookPrinter) As PrinterName,Trim(M.PrintName)+iif(M.Price=0,'',' (Price : Rs. '+Format(M.Price,'0.00')+')') As BookName,(Select Trim(PrintName) From GeneralMaster Where Code = M.Board) As BoardName,(Select Trim(PrintName) From GeneralMaster Where Code = M.[Size]) As BookSize,(Select Trim(PrintName) From GeneralMaster Where Code = C.BindingType) As BindingType,ActualQuantity,BillingQuantity,M.Forms,C.BindingForms,ExtraForms,FormFoldRate,FormStitchRate,FormPasteRate,[Rate/Book],[TotalPkts],[TotalBoxes],PktPackRate,BoxPackRate,CartageRate,Adjustment,[VAT%],[VAT]," & _
                                            "BillAmount,C.Remarks,(Select Trim(eMail) From AccountMaster Where Code = P.Binder) As EMailID,M.Narration,M.OneColorPages as Page1, M.TwoColorPages as Page2,M.FourColorPages as Page4,C.[LooseQty/Box],ExtraLooseQty,TotalLooseQty,C.[Qty/Pkt],C.[Pkt/Box],ISBN,C.UnitCost,C.ExtendDate,C.BookEdition,C.BillNo,C.BillDate,C.AdvanceCopyRequired,(Select Warehouse1  FROM PrintPVChild Where Code in(Select Ref From BookPOChild05 WHERE Code In(Select Code From BookPOParent Where Name=P.Name)) And Book in(Select  Book From BookPOParent Where Name=P.Name))As Warehouse1,(Select Warehouse2  FROM PrintPVChild Where Code in(Select Ref From BookPOChild05 WHERE Code In(Select Code From BookPOParent Where Name=P.Name)) And Book in(Select  Book From BookPOParent Where Name=P.Name))As Warehouse2," & _
                                            "(Select Warehouse3  FROM PrintPVChild Where Code in(Select Ref From BookPOChild05 WHERE Code In(Select Code From BookPOParent Where Name=P.Name)) And Book in(Select  Book From BookPOParent Where Name=P.Name))As Warehouse3 From BookPOParent P,BookPOChild08 C,BookMaster M Where P.Code = C.Code And P.Book = M.Code And P.Code = '" & OrderCode & "' And P.Binder <> ''", CxnDatabase, adOpenKeyset, adLockOptimistic
 
 
    If FrmBookPOPrintUtility.PrintUtility = False Then
            If PrintReport = 1 Then 'Book Printing
            If rstBookPOChild05.State = adStateOpen Then rstBookPOChild05.Close
               rstBookPOChild05.Open "Select 'BP/" & Right(Year(FinancialYearFrom), 2) + "-" + Right(Year(FinancialYearTo), 2) & "/'+Trim(P.Name) As OrderNo,OrderDate,TargetDate,Processing,(Select Trim(PrintName) From AccountMaster Where Code = P.BookPrinter) As PrinterName,(Select Trim(PrintName) From AccountMaster Where Code = P.Binder) As BinderName,Trim(M.PrintName)+iif(M.Price=0,'',' (Price : Rs. '+Format(M.Price,'0.00')+')') As BookName,M.DuplexPrinting,(Select Trim(PrintName) From GeneralMaster Where Code = M.Board) As BoardName,(Select Trim(PrintName) From GeneralMaster Where Code = M.[Size]) As BookSize,ActualQuantity,BillingQuantity01,BillingQuantity02," & _
                                    "Forms1,[Forms1-¼],[Forms1-½],[Forms1-1],[TotalForms1-¼],[TotalForms1-½],[TotalForms1-1],PrintRate1,PrintAmount1,PlateRate1,(Select Trim(PrintName) From PaperMaster Where Code = C.Paper1) As Paper1Name,PlateAmount1,[PaperWastage1%],PaperConsumptionOther1," & _
                                    "Forms2,[Forms2-¼],[Forms2-½],[Forms2-1],[TotalForms2-¼],[TotalForms2-½],[TotalForms2-1],PrintRate2,PrintAmount2,PlateRate2,(Select Trim(PrintName) From PaperMaster Where Code = C.Paper2) As Paper2Name,PlateAmount2,[PaperWastage2%],PaperConsumptionOther2," & _
                                    "Forms4,[Forms4-¼],[Forms4-½],[Forms4-1],[TotalForms4-¼],[TotalForms4-½],[TotalForms4-1],PrintRate4,PrintAmount4,PlateRate4,(Select Trim(PrintName) From PaperMaster Where Code = C.Paper4) As Paper4Name,PlateAmount4,[PaperWastage4%],PaperConsumptionOther4," & _
                                    "TotalPaperConsumption,Adjustment,[VAT%],VAT,BillAmount,C.Remarks,(Select Trim(eMail) From AccountMaster Where Code = P.BookPrinter) As EMailID,M.Narration,M.OneColorPages  As Pages1, M.TwoColorPages  As Pages2 ,M.FourColorPages  As Pages4,C.UnitCost,C.ExtendDate,C.BillNo,C.BillDate,M.Code As BookCode,C.Paper1,C.Paper2,C.Paper4,C.PlateMaking  From (BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN BookMaster M ON P.Book=M.Code Where P.Code = '" & OrderCode & "'", CxnDatabase, adOpenKeyset, adLockOptimistic
            End If
    End If
    
   
    Screen.MousePointer = vbNormal
    If rstBookPOChild08.RecordCount = 0 Then On Error GoTo 0: Exit Sub
    rptBookBindingOrder.Section23.Suppress = True
    rptBookBindingOrder.Section17.Suppress = True
    
    If rstCompanyMaster.State = adStateClosed Then rstCompanyMaster.Open "Select PrintName, Address1, Address2, Address3, Address4, Phone, Fax, EMail, Website From CompanyMaster", CxnDatabase, adOpenKeyset, adLockReadOnly
    
    '++++++++++++++++++++: Box Label Start :+++++++++++++++++++++'
    Dim oExcel As Object
    If Not FileExist(App.Path & "\Template\Box Label.xls") Then Exit Sub
    Set oExcel = CreateObject("Excel.Application")
    oExcel.Workbooks.Open (App.Path & "\Template\Box Label")
    oExcel.DisplayAlerts = False
    oExcel.Workbooks.Item(1).SaveAs (App.Path & "\Report\Box Label (" & CompCode & ")")
    oExcel.Application.DisplayAlerts = True
    oExcel.Sheets("Box Label").Select: oExcel.Sheets("Box Label").Unprotect ("shpl")
    oExcel.Application.Cells(1, "A").Value = Trim(rstCompanyMaster.Fields("PrintName").Value)
    oExcel.Application.Cells(9, "A").Value = Trim(rstCompanyMaster.Fields("PrintName").Value)
    oExcel.Application.Cells(2, "A").Value = "Phone : " + Trim(rstCompanyMaster.Fields("Phone").Value) + Space(1) + "Fax : " + Trim(rstCompanyMaster.Fields("Fax").Value)
    oExcel.Application.Cells(10, "A").Value = "Phone : " + Trim(rstCompanyMaster.Fields("Phone").Value) + Space(1) + "Fax : " + Trim(rstCompanyMaster.Fields("Fax").Value)
    oExcel.Application.Cells(3, "F").Value = Trim(rstBookPOChild08.Fields("BookName").Value)
    oExcel.Application.Cells(11, "F").Value = Trim(rstBookPOChild08.Fields("BookName").Value)
    oExcel.Application.Cells(4, "F").Value = Val(rstBookPOChild08.Fields("Qty/Pkt").Value)
    oExcel.Application.Cells(12, "F").Value = Val(rstBookPOChild08.Fields("Qty/Pkt").Value)
    oExcel.Application.Cells(5, "F").Value = Val(rstBookPOChild08.Fields("Pkt/Box").Value)
    oExcel.Application.Cells(5, "L").Value = Trim(rstBookPOChild08.Fields("ISBN").Value)
    oExcel.Application.Cells(13, "F").Value = Val(rstBookPOChild08.Fields("Pkt/Box").Value)
    oExcel.Application.Cells(13, "L").Value = Trim(rstBookPOChild08.Fields("ISBN").Value)
    oExcel.Application.Cells(6, "F").Value = Val(rstBookPOChild08.Fields("LooseQty/Box").Value)
    oExcel.Application.Cells(14, "F").Value = Val(rstBookPOChild08.Fields("LooseQty/Box").Value)
    oExcel.Application.Cells(7, "F").Value = (Val(oExcel.Application.Cells(4, "F").Value) * Val(oExcel.Application.Cells(5, "F").Value)) + Val(oExcel.Application.Cells(6, "F").Value)
    oExcel.Application.Cells(15, "F").Value = (Val(oExcel.Application.Cells(12, "F").Value) * Val(oExcel.Application.Cells(13, "F").Value)) + Val(oExcel.Application.Cells(14, "F").Value)
    oExcel.Application.Cells(7, "N").Value = Trim(rstBookPOChild08.Fields("OrderNo").Value)
    oExcel.Application.Cells(15, "N").Value = Trim(rstBookPOChild08.Fields("OrderNo").Value)
    oExcel.Sheets("Box Label").Activate
    oExcel.Sheets("Box Label").Protect ("shpl")
    oExcel.Workbooks.Item(1).Save
    
    
    '++++++++++++++++++++: Box Label End :+++++++++++++++++++++'
    rptBookBindingOrder.Text2.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptBookBindingOrder.Text3.SetText Trim(rstCompanyMaster.Fields("Address1").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address2").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address3").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address4").Value)
    rptBookBindingOrder.Text24.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & Space(1) & "Fax : " & Trim(rstCompanyMaster.Fields("Fax").Value) & Space(1) & "e-Mail : " & Trim(rstCompanyMaster.Fields("EMail").Value)
    rptBookBindingOrder.Text25.SetText " (" & Trim(NumberToWords(rstBookPOChild08.Fields("BillAmount").Value, True)) & ")"
    rptBookBindingOrder.Text27.SetText "for " & Trim(rstBookPOChild08.Fields("BinderName").Value)
    rptBookBindingOrder.Text28.SetText "for " & Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptBookBindingOrder.Text29.SetText Trim(COMPANY_CIN) 'Add here company cin no
    rptBookBindingOrder.Database.SetDataSource rstBookPOChild08, 3, 1
    
    If FrmBookPOPrintUtility.PrintUtility = False Then
        If PrintReport = 1 Then 'Book Printing
            rptBookBindingOrder.Subreport1_Text2.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
            rptBookBindingOrder.Subreport1_Text3.SetText Trim(rstCompanyMaster.Fields("Address1").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address2").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address3").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address4").Value)
            rptBookBindingOrder.Subreport1_Text24.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & Space(1) & "Fax : " & Trim(rstCompanyMaster.Fields("Fax").Value) & Space(1) & "e-Mail : " & Trim(rstCompanyMaster.Fields("EMail").Value)
            rptBookBindingOrder.Subreport1_Text25.SetText " (" & Trim(NumberToWords(rstBookPOChild05.Fields("BillAmount").Value, True)) & ")"
            rptBookBindingOrder.Subreport1_Text27.SetText "for " & Trim(rstBookPOChild05.Fields("PrinterName").Value)
            rptBookBindingOrder.Subreport1_Text28.SetText "for " & Trim(rstCompanyMaster.Fields("PrintName").Value)
            rptBookBindingOrder.Subreport1_Text35.SetText Trim(COMPANY_CIN)
            rptBookBindingOrder.Section23.Suppress = False
            rptBookBindingOrder.Subreport1_Section20.Suppress = True
            rptBookBindingOrder.Subreport1_Section21.Suppress = True
            rptBookBindingOrder.Subreport1_Section22.Suppress = True
            
            rptBookBindingOrder.Subreport1_Text44.SetText "Plate Making   : " & Trim(rstBookPOChild05.Fields("PlateMaking").Value)
            'rptBookBindingOrder.Subreport1_Text44.SetText "Plate Making   : " & GetPlateMaking(rstBookPOChild05.Fields("BookCode").Value, OrderCode, rstBookPOChild05.Fields("OrderDate").Value, "05", IIf(Val(rstBookPOChild05.Fields("Pages1").Value) > 0, "1", IIf(Val(rstBookPOChild05.Fields("Pages2").Value) > 0, "2", "4")))

            Dim PaperSendBalanceArray3() As String
            PaperSendBalanceArray3 = Split(GetPaperSend_Balance(OrderCode, "C"), "#")
            rptBookBindingOrder.Subreport1_Text45.SetText "Paper Send     : " & PaperSendBalanceArray3(0)
            rptBookBindingOrder.Subreport1_Text46.SetText "Paper Balance  : " & PaperSendBalanceArray3(1)
            
             
            'Dim PaperSendBalanceArray12() As String
            'PaperSendBalanceArray12 = Split(GetPaperSend_Balance(GetPOPrevious(PaperCode_For_Balance), "P"), "#")
            'rptBookBindingOrder.Subreport1_Text47.SetText "Prev Paper Bal : " & PaperSendBalanceArray12(1)
            
            

            rptBookBindingOrder.Subreport1.OpenSubreport.Database.SetDataSource rstBookPOChild05, 3, 1
        End If
    End If
    
    EMailID = rstBookPOChild08.Fields("EMailID").Value
    Attachment = Trim(rstBookPOChild08.Fields("OrderNo").Value)
    Attachment = Mid(Attachment, InStr(4, Attachment, "/") + 1)
    Message = "Dear Sir,<Br>Please find attached herewith PO #" & Trim(rstBookPOChild08.Fields("OrderNo").Value) & " for doing the needful at your end. An early finish of the job assigned to you will be highly appreciated.<Br>Kindly acknowledge the receipt of mail and confirm the date of completion of job.<Br><Br>" & IIf(Note = "", "", "<b><u>Note : " & Note & "</b></u><Br><Br>") & Trim(rstCompanyMaster.Fields("PrintName").Value) & "<Br>Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & "<Br>E-Mail : <a HRef='mailto:" & Trim(rstCompanyMaster.Fields("EMail").Value) & "'>" & Trim(rstCompanyMaster.Fields("EMail").Value) & "</a>"
    If OutputTo = "S" Then
        FrmReportViewer.EMailID = EMailID
        FrmReportViewer.Subject = "Book Binding Order #" & Trim(rstBookPOChild08.Fields("OrderNo").Value) + " Book : " + Trim(rstBookPOChild08.Fields("BookName").Value)
        FrmReportViewer.Attachment = Attachment
        FrmReportViewer.Message = Message
        Set FrmReportViewer.Report = rptBookBindingOrder
        FrmReportViewer.Show vbModal
    Else
        If rstBookPOList.State = adStateClosed Then
            If OutputType = "M" Then
                If CheckEmpty(EMailID, False) Then
                    rptBookBindingOrder.PrintOut False   ' Print Report Without Prompt
                    oExcel.Workbooks.Item(1).PrintOut
                Else
                    rptBookBindingOrder.ExportOptions.FormatType = crEFTPortableDocFormat    ' Set the Export Format As .Pdf
                    rptBookBindingOrder.ExportOptions.DestinationType = crEDTDiskFile
                    rptBookBindingOrder.ExportOptions.DiskFileName = App.Path & "\Report\" & Attachment & ".Pdf"
                    rptBookBindingOrder.Export False
                    rstBookPOChild08.MoveFirst
                    Set oOutlookMsg = oOutlook.CreateItem(olMailItem)
                    With oOutlookMsg
                        .To = EMailID
                        .Subject = "Book Binding Order #" & Trim(rstBookPOChild08.Fields("OrderNo").Value) + " Book : " + Trim(rstBookPOChild08.Fields("BookName").Value)
                        .HTMLBody = "<Font Face='Calibri' Size='3'>" & Message & "</a>" & "</Font>"
                        .Attachments.Add (App.Path & "\Report\" & Attachment & ".Pdf")
                        .Attachments.Add (App.Path & "\Report\Box Label (" & CompCode & ").xls")
                        .Importance = olImportanceHigh
                        .ReadReceiptRequested = True
                        .Send
                        If Err.Number = 0 Then CxnDatabase.Execute "UPDATE BookPOParent SET BBODStatus=1 WHERE Code='" & OrderCode & "'", RecordAffected
                        If RecordAffected = 0 Then DisplayError ("Failed to update EMail Flag (Book Binding Order)")
                    End With
                    Set oOutlookMsg = Nothing
                End If
            Else
                rptBookBindingOrder.PrintOut False   ' Print Report Without Prompt
            End If
        Else
            rptBookBindingOrder.PrintOut
        End If
    End If
    oExcel.Application.Quit: Set oExcel = Nothing
    Set rptBookBindingOrder = Nothing
    Set rptBookPrintingOrder = Nothing
    Set rptTitleLaminationOrder = Nothing
    Set rptTitlePrintingOrder = Nothing
    If rstBookPOList.State = adStateClosed Then Call CloseRecordset(rstCompanyMaster): Call CloseRecordset(rstBookPOChild08): Call CloseRecordset(rstBookPOChild05)
    On Error GoTo 0
End Sub

'Public Sub PrintBookBindingOrder(ByVal OrderCode As String, Optional ByVal Note As String, Optional ByVal OutputType As String)
'    Dim oOutlookMsg As Outlook.MailItem, RecordAffected As Integer
'    On Error Resume Next
'    Screen.MousePointer = vbHourglass
'
'    If rstBookPOChild08.State = adStateOpen Then rstBookPOChild08.Close
'    rstBookPOChild08.Open "Select 'BB/" & Right(Year(FinancialYearFrom), 2) + "-" + Right(Year(FinancialYearTo), 2) & "/'+Trim(P.Name) As OrderNo,OrderDate,TargetDate,(Select Trim(PrintName) From AccountMaster Where Code = P.Binder) As BinderName,(Select Trim(PrintName) From AccountMaster Where Code = P.Laminator) As LaminatorName,(Select Trim(PrintName) From AccountMaster Where Code = P.BookPrinter) As PrinterName,Trim(M.PrintName)+iif(M.Price=0,'',' (Price : Rs. '+Format(M.Price,'0.00')+')') As BookName,(Select Trim(PrintName) From GeneralMaster Where Code = M.Board) As BoardName,(Select Trim(PrintName) From GeneralMaster Where Code = M.[Size]) As BookSize,(Select Trim(PrintName) From GeneralMaster Where Code = C.BindingType) As BindingType,ActualQuantity,BillingQuantity,M.Forms,C.BindingForms,ExtraForms,FormFoldRate,FormStitchRate,FormPasteRate,[Rate/Book],[TotalPkts],[TotalBoxes],PktPackRate,BoxPackRate,CartageRate,Adjustment,[VAT%],[VAT]," & _
'                                            "BillAmount,C.Remarks,(Select Trim(eMail) From AccountMaster Where Code = P.Binder) As EMailID,M.Narration,M.OneColorPages as Page1, M.TwoColorPages as Page2,M.FourColorPages as Page4,C.[LooseQty/Box],ExtraLooseQty,TotalLooseQty,C.[Qty/Pkt],C.[Pkt/Box],ISBN,C.UnitCost,C.ExtendDate,C.BookEdition From BookPOParent P,BookPOChild08 C,BookMaster M Where P.Code = C.Code And P.Book = M.Code And P.Code = '" & OrderCode & "' And P.Binder <> ''", CxnDatabase, adOpenKeyset, adLockOptimistic
'
'    Screen.MousePointer = vbNormal
'    If rstBookPOChild08.RecordCount = 0 Then On Error GoTo 0: Exit Sub
'    If rstCompanyMaster.State = adStateClosed Then rstCompanyMaster.Open "Select PrintName, Address1, Address2, Address3, Address4, Phone, Fax, EMail, Website From CompanyMaster", CxnDatabase, adOpenKeyset, adLockReadOnly
'    '++++++++++++++++++++: Box Label Start :+++++++++++++++++++++'
'    Dim oExcel As Object
'    If Not FileExist(App.Path & "\Template\Box Label.xls") Then Exit Sub
'    Set oExcel = CreateObject("Excel.Application")
'    oExcel.Workbooks.Open (App.Path & "\Template\Box Label")
'    oExcel.DisplayAlerts = False
'    oExcel.Workbooks.Item(1).SaveAs (App.Path & "\Report\Box Label (" & CompCode & ")")
'    oExcel.Application.DisplayAlerts = True
'    oExcel.Sheets("Box Label").Select: oExcel.Sheets("Box Label").Unprotect ("shpl")
'    oExcel.Application.Cells(1, "A").Value = Trim(rstCompanyMaster.Fields("PrintName").Value)
'    oExcel.Application.Cells(9, "A").Value = Trim(rstCompanyMaster.Fields("PrintName").Value)
'    oExcel.Application.Cells(2, "A").Value = "Phone : " + Trim(rstCompanyMaster.Fields("Phone").Value) + Space(1) + "Fax : " + Trim(rstCompanyMaster.Fields("Fax").Value)
'    oExcel.Application.Cells(10, "A").Value = "Phone : " + Trim(rstCompanyMaster.Fields("Phone").Value) + Space(1) + "Fax : " + Trim(rstCompanyMaster.Fields("Fax").Value)
'    oExcel.Application.Cells(3, "F").Value = Trim(rstBookPOChild08.Fields("BookName").Value)
'    oExcel.Application.Cells(11, "F").Value = Trim(rstBookPOChild08.Fields("BookName").Value)
'    oExcel.Application.Cells(4, "F").Value = Val(rstBookPOChild08.Fields("Qty/Pkt").Value)
'    oExcel.Application.Cells(12, "F").Value = Val(rstBookPOChild08.Fields("Qty/Pkt").Value)
'    oExcel.Application.Cells(5, "F").Value = Val(rstBookPOChild08.Fields("Pkt/Box").Value)
'    oExcel.Application.Cells(5, "L").Value = Trim(rstBookPOChild08.Fields("ISBN").Value)
'    oExcel.Application.Cells(13, "F").Value = Val(rstBookPOChild08.Fields("Pkt/Box").Value)
'    oExcel.Application.Cells(13, "L").Value = Trim(rstBookPOChild08.Fields("ISBN").Value)
'    oExcel.Application.Cells(6, "F").Value = Val(rstBookPOChild08.Fields("LooseQty/Box").Value)
'    oExcel.Application.Cells(14, "F").Value = Val(rstBookPOChild08.Fields("LooseQty/Box").Value)
'    oExcel.Application.Cells(7, "F").Value = (Val(oExcel.Application.Cells(4, "F").Value) * Val(oExcel.Application.Cells(5, "F").Value)) + Val(oExcel.Application.Cells(6, "F").Value)
'    oExcel.Application.Cells(15, "F").Value = (Val(oExcel.Application.Cells(12, "F").Value) * Val(oExcel.Application.Cells(13, "F").Value)) + Val(oExcel.Application.Cells(14, "F").Value)
'    oExcel.Application.Cells(7, "N").Value = Trim(rstBookPOChild08.Fields("OrderNo").Value)
'    oExcel.Application.Cells(15, "N").Value = Trim(rstBookPOChild08.Fields("OrderNo").Value)
'    oExcel.Sheets("Box Label").Activate
'    oExcel.Sheets("Box Label").Protect ("shpl")
'    oExcel.Workbooks.Item(1).Save
'    '++++++++++++++++++++: Box Label End :+++++++++++++++++++++'
'    rptBookBindingOrder.Text2.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
'    rptBookBindingOrder.Text3.SetText Trim(rstCompanyMaster.Fields("Address1").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address2").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address3").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address4").Value)
'    rptBookBindingOrder.Text24.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & Space(1) & "Fax : " & Trim(rstCompanyMaster.Fields("Fax").Value) & Space(1) & "e-Mail : " & Trim(rstCompanyMaster.Fields("EMail").Value)
'    rptBookBindingOrder.Text25.SetText " (" & Trim(NumberToWords(rstBookPOChild08.Fields("BillAmount").Value, True)) & ")"
'    rptBookBindingOrder.Text27.SetText "for " & Trim(rstBookPOChild08.Fields("BinderName").Value)
'
'    rptBookBindingOrder.Text28.SetText "for " & Trim(rstCompanyMaster.Fields("PrintName").Value)
'    rptBookBindingOrder.Text29.SetText Trim(COMPANY_CIN) 'Add here company cin no
'
'    rptBookBindingOrder.Database.SetDataSource rstBookPOChild08, 3, 1
'    EMailID = rstBookPOChild08.Fields("EMailID").Value
'    Attachment = Trim(rstBookPOChild08.Fields("OrderNo").Value)
'    Attachment = Mid(Attachment, InStr(4, Attachment, "/") + 1)
'    Message = "Dear Sir,<Br>Please find attached herewith PO #" & Trim(rstBookPOChild08.Fields("OrderNo").Value) & " for doing the needful at your end. An early finish of the job assigned to you will be highly appreciated.<Br>Kindly acknowledge the receipt of mail and confirm the date of completion of job.<Br><Br>" & IIf(Note = "", "", "<b><u>Note : " & Note & "</b></u><Br><Br>") & Trim(rstCompanyMaster.Fields("PrintName").Value) & "<Br>Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & "<Br>E-Mail : <a HRef='mailto:" & Trim(rstCompanyMaster.Fields("EMail").Value) & "'>" & Trim(rstCompanyMaster.Fields("EMail").Value) & "</a>"
'    If OutputTo = "S" Then
'        FrmReportViewer.EMailID = EMailID
'        FrmReportViewer.Subject = "Book Binding Order #" & Trim(rstBookPOChild08.Fields("OrderNo").Value) + " Book : " + Trim(rstBookPOChild08.Fields("BookName").Value)
'        FrmReportViewer.Attachment = Attachment
'        FrmReportViewer.Message = Message
'        Set FrmReportViewer.Report = rptBookBindingOrder
'        FrmReportViewer.Show vbModal
'    Else
'        If rstBookPOList.State = adStateClosed Then
'            If OutputType = "M" Then
'                If CheckEmpty(EMailID, False) Then
'                    rptBookBindingOrder.PrintOut False   ' Print Report Without Prompt
'                    oExcel.Workbooks.Item(1).PrintOut
'                Else
'                    rptBookBindingOrder.ExportOptions.FormatType = crEFTPortableDocFormat    ' Set the Export Format As .Pdf
'                    rptBookBindingOrder.ExportOptions.DestinationType = crEDTDiskFile
'                    rptBookBindingOrder.ExportOptions.DiskFileName = App.Path & "\Report\" & Attachment & ".Pdf"
'                    rptBookBindingOrder.Export False
'                    rstBookPOChild08.MoveFirst
'                    Set oOutlookMsg = oOutlook.CreateItem(olMailItem)
'                    With oOutlookMsg
'                        .To = EMailID
'                        .Subject = "Book Binding Order #" & Trim(rstBookPOChild08.Fields("OrderNo").Value) + " Book : " + Trim(rstBookPOChild08.Fields("BookName").Value)
'                        .HTMLBody = "<Font Face='Calibri' Size='3'>" & Message & "</a>" & "</Font>"
'                        .Attachments.Add (App.Path & "\Report\" & Attachment & ".Pdf")
'                        .Attachments.Add (App.Path & "\Report\Box Label (" & CompCode & ").xls")
'                        .Importance = olImportanceHigh
'                        .ReadReceiptRequested = True
'                        .Send
'                        If Err.Number = 0 Then CxnDatabase.Execute "UPDATE BookPOParent SET BBODStatus=1 WHERE Code='" & OrderCode & "'", RecordAffected
'                        If RecordAffected = 0 Then DisplayError ("Failed to update EMail Flag (Book Binding Order)")
'                    End With
'                    Set oOutlookMsg = Nothing
'                End If
'            Else
'                rptBookBindingOrder.PrintOut False   ' Print Report Without Prompt
'            End If
'        Else
'            rptBookBindingOrder.PrintOut
'        End If
'    End If
'    oExcel.Application.Quit: Set oExcel = Nothing
'    Set rptBookBindingOrder = Nothing
'    If rstBookPOList.State = adStateClosed Then Call CloseRecordset(rstCompanyMaster): Call CloseRecordset(rstBookPOChild08)
'    On Error GoTo 0
'End Sub


Public Sub PrintBookBoxLabel(ByVal OrderCode As String, Optional ByVal Note As String, Optional ByVal OutputType As String)
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    If rstBookPOChild08.State = adStateOpen Then rstBookPOChild08.Close
    rstBookPOChild08.Open "Select 'BB/" & Right(Year(FinancialYearFrom), 2) + "-" + Right(Year(FinancialYearTo), 2) & "/'+Trim(P.Name) As OrderNo,Trim(M.PrintName)+iif(M.Price=0,'',' (Price : Rs. '+Format(M.Price,'0.00')+')') As BookName,C.[LooseQty/Box],ExtraLooseQty,TotalLooseQty,C.[Qty/Pkt],C.[Pkt/Box],ISBN From BookPOParent P,BookPOChild08 C,BookMaster M Where P.Code = C.Code And P.Book = M.Code And P.Code = '" & OrderCode & "' And P.Binder <> ''", CxnDatabase, adOpenKeyset, adLockOptimistic
    Screen.MousePointer = vbNormal
    If rstBookPOChild08.RecordCount = 0 Then On Error GoTo 0: Exit Sub
    If rstCompanyMaster.State = adStateClosed Then rstCompanyMaster.Open "Select PrintName, Address1, Address2, Address3, Address4, Phone, Fax, EMail, Website From CompanyMaster", CxnDatabase, adOpenKeyset, adLockReadOnly
    Dim oExcel As Object
    If Not FileExist(App.Path & "\Template\Box Label.xls") Then Exit Sub
    Set oExcel = CreateObject("Excel.Application")
    oExcel.Workbooks.Open (App.Path & "\Template\Box Label")
    oExcel.DisplayAlerts = False
    oExcel.Workbooks.Item(1).SaveAs (App.Path & "\Report\Box Label (" & CompCode & ")")
    oExcel.Application.DisplayAlerts = True
    oExcel.Sheets("Box Label").Select: oExcel.Sheets("Box Label").Unprotect ("shpl")
    oExcel.Application.Cells(1, "A").Value = Trim(rstCompanyMaster.Fields("PrintName").Value)
    oExcel.Application.Cells(9, "A").Value = Trim(rstCompanyMaster.Fields("PrintName").Value)
    oExcel.Application.Cells(2, "A").Value = "Phone : " + Trim(rstCompanyMaster.Fields("Phone").Value) + Space(1) + "Fax : " + Trim(rstCompanyMaster.Fields("Fax").Value)
    oExcel.Application.Cells(10, "A").Value = "Phone : " + Trim(rstCompanyMaster.Fields("Phone").Value) + Space(1) + "Fax : " + Trim(rstCompanyMaster.Fields("Fax").Value)
    oExcel.Application.Cells(3, "F").Value = Trim(rstBookPOChild08.Fields("BookName").Value)
    oExcel.Application.Cells(11, "F").Value = Trim(rstBookPOChild08.Fields("BookName").Value)
    oExcel.Application.Cells(4, "F").Value = Val(rstBookPOChild08.Fields("Qty/Pkt").Value)
    oExcel.Application.Cells(12, "F").Value = Val(rstBookPOChild08.Fields("Qty/Pkt").Value)
    oExcel.Application.Cells(5, "F").Value = Val(rstBookPOChild08.Fields("Pkt/Box").Value)
    oExcel.Application.Cells(5, "L").Value = Trim(rstBookPOChild08.Fields("ISBN").Value)
    oExcel.Application.Cells(13, "F").Value = Val(rstBookPOChild08.Fields("Pkt/Box").Value)
    oExcel.Application.Cells(13, "L").Value = Trim(rstBookPOChild08.Fields("ISBN").Value)
    oExcel.Application.Cells(6, "F").Value = Val(rstBookPOChild08.Fields("LooseQty/Box").Value)
    oExcel.Application.Cells(14, "F").Value = Val(rstBookPOChild08.Fields("LooseQty/Box").Value)
    oExcel.Application.Cells(7, "F").Value = (Val(oExcel.Application.Cells(4, "F").Value) * Val(oExcel.Application.Cells(5, "F").Value)) + Val(oExcel.Application.Cells(6, "F").Value)
    oExcel.Application.Cells(15, "F").Value = (Val(oExcel.Application.Cells(12, "F").Value) * Val(oExcel.Application.Cells(13, "F").Value)) + Val(oExcel.Application.Cells(14, "F").Value)
    oExcel.Application.Cells(7, "N").Value = Trim(rstBookPOChild08.Fields("OrderNo").Value)
    oExcel.Application.Cells(15, "N").Value = Trim(rstBookPOChild08.Fields("OrderNo").Value)
    oExcel.Sheets("Box Label").Activate
    oExcel.Sheets("Box Label").Protect ("shpl")
    oExcel.Workbooks.Item(1).Save
    If OutputTo = "S" Then oExcel.Range("A1").Activate: oExcel.Visible = True Else oExcel.Workbooks.Item(1).PrintOut: oExcel.Application.Quit
    Set oExcel = Nothing
    If rstBookPOList.State = adStateClosed Then Call CloseRecordset(rstCompanyMaster): Call CloseRecordset(rstBookPOChild08)
    On Error GoTo 0
End Sub
Public Sub PrintBookOrder(ByVal OrderCode As String, Optional ByVal Note As String, Optional ByVal OutputType As String)
    Dim oOutlookMsg As Outlook.MailItem
    On Error Resume Next
      
    Screen.MousePointer = vbHourglass
    If rstBookPOChild08.State = adStateOpen Then rstBookPOChild08.Close
    rstBookPOChild08.Open "Select 'BO/" & Right(Year(FinancialYearFrom), 2) + "-" + Right(Year(FinancialYearTo), 2) & "/'+Trim(P.Name) As OrderNo,OrderDate,TargetDate,(Select Trim(PrintName) From AccountMaster Where Code = P.Binder) As PartyName,Trim(M.PrintName)+iif(M.Price=0,'',' (Price : Rs. '+Format(M.Price,'0.00')+')') As BookName,(Select Trim(PrintName) From GeneralMaster Where Code = M.Board) As BoardName,(Select Trim(PrintName) From GeneralMaster Where Code = M.[Size]) As BookSize,M.FormType,(Select Trim(PrintName) From GeneralMaster Where Code = C.BindingType) As BindingType,(Select Trim(PrintName) From GeneralMaster Where Code = M.LaminationType) As LaminationType,TitleFrontColor,TitleBackColor,OneColorForms,TwoColorForms,FourColorForms,ActualQuantity,[Rate/Book],Adjustment,[VAT%],[VAT],BillAmount,C.Remarks,(Select Trim(eMail) From AccountMaster Where Code = P.Binder) As EMailID " & _
                                            ",M.OneColorPages as Page1, M.TwoColorPages as Page2,M.FourColorPages as Page4 From (BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code) INNER JOIN BookMaster M ON P.Book=M.Code Where P.Code = '" & OrderCode & "' And C.[Rate/Book]<>0", CxnDatabase, adOpenKeyset, adLockOptimistic
    Screen.MousePointer = vbNormal
    If rstBookPOChild08.RecordCount = 0 Then On Error GoTo 0: Exit Sub
    If rstCompanyMaster.State = adStateClosed Then rstCompanyMaster.Open "Select PrintName, Address1, Address2, Address3, Address4, Phone, Fax, EMail, Website From CompanyMaster", CxnDatabase, adOpenKeyset, adLockReadOnly
    rptBookOrder.Text2.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptBookOrder.Text3.SetText Trim(rstCompanyMaster.Fields("Address1").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address2").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address3").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address4").Value)
    rptBookOrder.Text24.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & Space(1) & "Fax : " & Trim(rstCompanyMaster.Fields("Fax").Value) & Space(1) & "e-Mail : " & Trim(rstCompanyMaster.Fields("EMail").Value)
    rptBookOrder.Text25.SetText " (" & Trim(NumberToWords(rstBookPOChild08.Fields("BillAmount").Value, True)) & ")"
    rptBookOrder.Text27.SetText "for " & Trim(rstBookPOChild08.Fields("PartyName").Value)
    rptBookOrder.Text28.SetText "for " & Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptBookOrder.Text9.SetText Trim(COMPANY_CIN) 'Add here company cin no
    rptBookOrder.Database.SetDataSource rstBookPOChild08, 3, 1
    EMailID = rstBookPOChild08.Fields("EMailID").Value
    Attachment = Trim(rstBookPOChild08.Fields("OrderNo").Value)
    Attachment = Mid(Attachment, InStr(4, Attachment, "/") + 1)
    Message = "Dear Sir,<Br>Please find attached herewith PO #" & Trim(rstBookPOChild08.Fields("OrderNo").Value) & " for doing the needful at your end. An early finish of the job assigned to you will be highly appreciated.<Br>Kindly acknowledge the receipt of mail and confirm the date of completion of job.<Br><Br>" & IIf(Note = "", "", "<b><u>Note : " & Note & "</b></u><Br><Br>") & Trim(rstCompanyMaster.Fields("PrintName").Value) & "<Br>Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & "<Br>E-Mail : <a HRef='mailto:" & Trim(rstCompanyMaster.Fields("EMail").Value) & "'>" & Trim(rstCompanyMaster.Fields("EMail").Value) & "</a>"
    If OutputTo = "S" Then
        FrmReportViewer.EMailID = EMailID
        FrmReportViewer.Subject = "Book Order #" & Trim(rstBookPOChild08.Fields("OrderNo").Value) + " Book : " + Trim(rstBookPOChild08.Fields("BookName").Value)
        FrmReportViewer.Attachment = Attachment
        FrmReportViewer.Message = Message
        Set FrmReportViewer.Report = rptBookOrder
        FrmReportViewer.Show vbModal
    Else
        If rstBookPOList.State = adStateClosed Then
            If EMailID = "" Or OutputType = "P" Then
                rptBookOrder.PrintOut False   ' Print Report Without Prompt
            Else
                rptBookOrder.ExportOptions.FormatType = crEFTPortableDocFormat    ' Set the Export Format As .Pdf
                rptBookOrder.ExportOptions.DestinationType = crEDTDiskFile
                rptBookOrder.ExportOptions.DiskFileName = App.Path & "\Report\" & Attachment & ".Pdf"
                rptBookOrder.Export False
                rstBookPOChild08.MoveFirst
                Set oOutlookMsg = oOutlook.CreateItem(olMailItem)
                With oOutlookMsg
                    .To = EMailID
                    .Subject = "Book Order #" & Trim(rstBookPOChild08.Fields("OrderNo").Value) + " Book : " + Trim(rstBookPOChild08.Fields("BookName").Value)
                    .HTMLBody = "<Font Face='Calibri' Size='3'>" & Message & "</a>" & "</Font>"
                    .Attachments.Add (App.Path & "\Report\" & Attachment & ".Pdf")
                    .Importance = olImportanceHigh
                    .ReadReceiptRequested = True
                    .Send
                    If Err.Number = 0 Then CxnDatabase.Execute "UPDATE BookPOParent SET BBODStatus=1 WHERE Code='" & OrderCode & "'"
                End With
                Set oOutlookMsg = Nothing
            End If
        Else
            rptBookOrder.PrintOut
        End If
    End If
    Set rptBookOrder = Nothing
    If rstBookPOList.State = adStateClosed Then Call CloseRecordset(rstCompanyMaster): Call CloseRecordset(rstBookPOChild08)
    On Error GoTo 0
End Sub
Public Sub PrintCostSheet(ByVal OrderCode As String)
    Dim oExcel As Object
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    If rstBookPOChild05.State = adStateOpen Then rstBookPOChild05.Close
    rstBookPOChild05.Open "SELECT M1.PrintName As BookName,M2.PrintName As BookSize,M1.FormType,M1.Royalty,C1.ActualQuantity As Quantity,(C1.Pages1+C1.Pages2+C1.Pages4) As Pages,C1.Forms1,C1.[TotalForms1-¼]+C1.[TotalForms1-½]+C1.[TotalForms1-1] As TotalForms1,C1.PrintRate1,C1.PlateType1,C1.[TotalPlates1-¼]+C1.[TotalPlates1-½]+C1.[TotalPlates1-1] As TotalPlates1,C1.PlateRate1,(SELECT STR([Weight/Ream]) FROM PaperMaster WHERE Code=C1.Paper1) As TextPaper1,(SELECT Top 1 STR([Rate/Kg]) FROM PaperPOParent P INNER JOIN PaperPOChild C ON P.code=C.Code WHERE Paper=C1.Paper1 ORDER BY P.Name DESC) As TextPaper1Rate,PaperConsumptionOther1 As TextPaperConsumption1," & _
                                            "C1.Forms2,C1.[TotalForms2-¼]+C1.[TotalForms2-½]+C1.[TotalForms2-1] As TotalForms2,C1.PrintRate2,C1.PlateType2,C1.[TotalPlates2-¼]+C1.[TotalPlates2-½]+C1.[TotalPlates2-1] As TotalPlates2,C1.PlateRate2,(SELECT STR([Weight/Ream]) FROM PaperMaster WHERE Code=C1.Paper2) As TextPaper2,(SELECT Top 1 STR([Rate/Kg]) FROM PaperPOParent P INNER JOIN PaperPOChild C ON P.code=C.Code WHERE Paper=C1.Paper2 ORDER BY P.Name DESC) As TextPaper2Rate,PaperConsumptionOther2 As TextPaperConsumption2," & _
                                            "C1.Forms4,C1.[TotalForms4-¼]+C1.[TotalForms4-½]+C1.[TotalForms4-1] As TotalForms4,C1.PrintRate4,C1.PlateType4,C1.[TotalPlates4-¼]+C1.[TotalPlates4-½]+C1.[TotalPlates4-1] As TotalPlates4,C1.PlateRate4 ,C1.Paper4,(SELECT STR([Weight/Ream]) FROM PaperMaster WHERE Code=C1.Paper4) As TextPaper4,(SELECT Top 1 STR([Rate/Kg]) FROM PaperPOParent P INNER JOIN PaperPOChild C ON P.code=C.Code WHERE Paper=C1.Paper4 ORDER BY P.Name DESC) As TextPaper4Rate,PaperConsumptionOther4 As TextPaperConsumption4," & _
                                            "C2.FrontPrintingType,C2.BackPrintingType,C2.TotalForms,C2.PrintRate,C2.PlateType,C2.TotalPlates,C2.PlateRate,C2.Paper,(SELECT STR([Weight/Ream]) FROM PaperMaster WHERE Code=C2.Paper) As TitlePaper,(SELECT Top 1 STR([Rate/Kg]) FROM PaperPOParent P INNER JOIN PaperPOChild C ON P.code=C.Code WHERE Paper=C2.Paper ORDER BY P.Name DESC) As TitlePaperRate,PaperConsumptionOther As TitlePaperConsumption," & _
                                            "(SELECT PrintName FROM GeneralMaster WHERE Code=C3.LaminationType) As LaminationType,C3.LaminationRate,(SELECT PrintName FROM GeneralMaster WHERE Code=C4.BindingType) As BindingType,(C4.BindingForms+C4.ExtraForms) As BindingForms,C4.FormFoldRate,C4.FormStitchRate,C4.FormPasteRate,C4.[Rate/Book],(TotalPkts*PktPackRate)+(TotalBoxes*BoxPackRate)+(TotalBoxes*CartageRate) As [Packing&Cartage] FROM (((((BookPOParent P INNER JOIN BookPOChild05 C1 ON (P.Code=C1.Code AND P.Code='" & OrderCode & "')) LEFT JOIN BookPOChild06 C2 ON P.Code=C2.Code) LEFT JOIN BookPOChild07 C3 ON P.Code=C3.Code) LEFT JOIN BookPOChild08 C4 ON P.Code=C4.Code) INNER JOIN BookMaster M1 ON P.Book=M1.Code) INNER JOIN GeneralMaster M2 ON M1.[Size]=M2.Code", CxnDatabase, adOpenKeyset, adLockReadOnly
    Screen.MousePointer = vbNormal
    If rstBookPOChild05.RecordCount = 0 Then
        Screen.MousePointer = vbNormal
        On Error GoTo 0
        Exit Sub
    End If
    DoEvents
    Set oExcel = CreateObject("Excel.Application")
    oExcel.Workbooks.Open (App.Path & "\Template\Cost Sheet")
    oExcel.DisplayAlerts = False
    'oExcel.Workbooks.Item(1).SaveAs (App.Path & "\Cost Sheet\" & Replace(Trim(rstBookPOList.Fields("BookName").Value), "*", "") + " (" + Format(Date, "dd-MM-yyyy") + ")")
    oExcel.Workbooks.Item(1).SaveAs (App.Path & "\Cost Sheet\" & Replace(Replace(Replace(Trim(rstBookPOList.Fields("BookName").Value), ".", ""), "/", " x "), "*", "") + " (" + Format(Date, "dd-MM-yyyy") + ")")
    oExcel.DisplayAlerts = True
    oExcel.Sheets("Sheet1").Select
    oExcel.Sheets("Sheet1").Unprotect ("shpl")
    oExcel.Application.Cells(3, 2).Value = Trim(rstBookPOChild05.Fields("BookName").Value)
    oExcel.Application.Cells(4, 2).Value = Val(rstBookPOChild05.Fields("Pages").Value)
    oExcel.Application.Cells(50, "C").Value = Val(rstBookPOChild05.Fields("Royalty").Value) / 100
    oExcel.Application.Cells(5, 2).Value = Val(rstBookPOChild05.Fields("Quantity").Value)
    oExcel.Application.Cells(6, 2).Value = Trim(rstBookPOChild05.Fields("BookSize").Value) & "/" & IIf(rstBookPOChild05.Fields("FormType").Value = "1", "08", IIf(rstBookPOChild05.Fields("FormType").Value = "2", "16", IIf(rstBookPOChild05.Fields("FormType").Value = "3", "04", IIf(rstBookPOChild05.Fields("FormType").Value = "4", "12", IIf(rstBookPOChild05.Fields("FormType").Value = "5", "24", IIf(rstBookPOChild05.Fields("FormType").Value = "6", "32", "64"))))))
    oExcel.Application.Cells(7, 2).Value = Val(rstBookPOChild05.Fields("Forms1").Value) + Val(rstBookPOChild05.Fields("Forms2").Value) + Val(rstBookPOChild05.Fields("Forms4").Value)
    oExcel.Application.Cells(8, 2).Value = Val(rstBookPOChild05.Fields("BindingForms").Value)
    oExcel.Application.Cells(9, 2).Value = Trim(rstBookPOChild05.Fields("LaminationType").Value)
    oExcel.Application.Cells(10, 2).Value = Trim(rstBookPOChild05.Fields("BindingType").Value)
    If Not IsNull(rstBookPOChild05.Fields("TextPaper1").Value) Then
        oExcel.Application.Cells(24, 3).Value = Val(rstBookPOChild05.Fields("TextPaper1").Value)
        oExcel.Application.Cells(24, 4).Value = Val(rstBookPOChild05.Fields("TextPaperConsumption1").Value)
        oExcel.Application.Cells(24, 8).Value = Val(rstBookPOChild05.Fields("TextPaper1Rate").Value)
    End If
    If Not IsNull(rstBookPOChild05.Fields("TextPaper2").Value) Then
        oExcel.Application.Cells(25, 3).Value = Val(rstBookPOChild05.Fields("TextPaper2").Value)
        oExcel.Application.Cells(25, 4).Value = Val(rstBookPOChild05.Fields("TextPaperConsumption2").Value)
        oExcel.Application.Cells(25, 8).Value = Val(rstBookPOChild05.Fields("TextPaper2Rate").Value)
    End If
    If Not IsNull(rstBookPOChild05.Fields("TextPaper4").Value) Then
        oExcel.Application.Cells(26, 3).Value = Val(rstBookPOChild05.Fields("TextPaper4").Value)
        oExcel.Application.Cells(26, 4).Value = Val(rstBookPOChild05.Fields("TextPaperConsumption4").Value)
        oExcel.Application.Cells(26, 8).Value = Val(rstBookPOChild05.Fields("TextPaper4Rate").Value)
    End If
    oExcel.Application.Cells(27, 2).Value = Trim(str(Val(rstBookPOChild05.Fields("FrontPrintingType").Value) + Val(rstBookPOChild05.Fields("BackPrintingType").Value))) & " Color"
    oExcel.Application.Cells(27, 3).Value = Val(rstBookPOChild05.Fields("TitlePaper").Value)
    oExcel.Application.Cells(27, 4).Value = Val(rstBookPOChild05.Fields("TitlePaperConsumption").Value)
    oExcel.Application.Cells(27, 8).Value = Val(rstBookPOChild05.Fields("TitlePaperRate").Value)
    If Val(rstBookPOChild05.Fields("TotalPlates1").Value) Then
        oExcel.Application.Cells(28, 3).Value = Switch(rstBookPOChild05.Fields("PlateType1").Value = "1", "Deepatch", rstBookPOChild05.Fields("PlateType1").Value = "2", "PS", rstBookPOChild05.Fields("PlateType1").Value = "3", "Wipeon")
        oExcel.Application.Cells(28, 4).Value = Val(rstBookPOChild05.Fields("TotalPlates1").Value)
        oExcel.Application.Cells(28, 5).Value = Val(rstBookPOChild05.Fields("PlateRate1").Value)
    End If
    If Val(rstBookPOChild05.Fields("TotalPlates2").Value) Then
        oExcel.Application.Cells(29, 3).Value = Switch(rstBookPOChild05.Fields("PlateType2").Value = "1", "Deepatch", rstBookPOChild05.Fields("PlateType2").Value = "2", "PS", rstBookPOChild05.Fields("PlateType2").Value = "3", "Wipeon")
        oExcel.Application.Cells(29, 4).Value = Val(rstBookPOChild05.Fields("TotalPlates2").Value)
        oExcel.Application.Cells(29, 5).Value = Val(rstBookPOChild05.Fields("PlateRate2").Value)
    End If
    If Val(rstBookPOChild05.Fields("TotalPlates4").Value) Then
        oExcel.Application.Cells(30, 3).Value = Switch(rstBookPOChild05.Fields("PlateType4").Value = "1", "Deepatch", rstBookPOChild05.Fields("PlateType4").Value = "2", "PS", rstBookPOChild05.Fields("PlateType4").Value = "3", "Wipeon")
        oExcel.Application.Cells(30, 4).Value = Val(rstBookPOChild05.Fields("TotalPlates4").Value)
        oExcel.Application.Cells(30, 5).Value = Val(rstBookPOChild05.Fields("PlateRate4").Value)
    End If
    oExcel.Application.Cells(31, 2).Value = Trim(str(Val(rstBookPOChild05.Fields("FrontPrintingType").Value) + Val(rstBookPOChild05.Fields("BackPrintingType").Value))) & " Color"
    oExcel.Application.Cells(31, 3).Value = Switch(rstBookPOChild05.Fields("PlateType").Value = "1", "Deepatch", rstBookPOChild05.Fields("PlateType").Value = "2", "PS", rstBookPOChild05.Fields("PlateType").Value = "3", "Wipeon")
    oExcel.Application.Cells(31, 4).Value = Val(rstBookPOChild05.Fields("TotalPlates").Value)
    oExcel.Application.Cells(31, 5).Value = Val(rstBookPOChild05.Fields("PlateRate").Value)
    oExcel.Application.Cells(32, 4).Value = Val(rstBookPOChild05.Fields("Forms1").Value)
    oExcel.Application.Cells(32, 5).Value = Val(rstBookPOChild05.Fields("PrintRate1").Value)
    oExcel.Application.Cells(32, 8).Value = Val(rstBookPOChild05.Fields("TotalForms1").Value)
    oExcel.Application.Cells(33, 4).Value = Val(rstBookPOChild05.Fields("Forms2").Value)
    oExcel.Application.Cells(33, 5).Value = Val(rstBookPOChild05.Fields("PrintRate2").Value)
    oExcel.Application.Cells(33, 8).Value = Val(rstBookPOChild05.Fields("TotalForms2").Value)
    oExcel.Application.Cells(34, 4).Value = Val(rstBookPOChild05.Fields("Forms4").Value)
    oExcel.Application.Cells(34, 5).Value = Val(rstBookPOChild05.Fields("PrintRate4").Value)
    oExcel.Application.Cells(34, 8).Value = Val(rstBookPOChild05.Fields("TotalForms4").Value)
    oExcel.Application.Cells(35, 2).Value = Trim(str(Val(rstBookPOChild05.Fields("FrontPrintingType").Value) + Val(rstBookPOChild05.Fields("BackPrintingType").Value))) & " Color"
    oExcel.Application.Cells(35, 5).Value = Val(rstBookPOChild05.Fields("PrintRate").Value)
    oExcel.Application.Cells(35, 8).Value = Val(rstBookPOChild05.Fields("TotalForms").Value)
    oExcel.Application.Cells(36, 5).Value = Val(rstBookPOChild05.Fields("LaminationRate").Value)
    If Val(rstBookPOChild05.Fields("FormPasteRate").Value) > 0 Then
        oExcel.Application.Cells(37, 5).Value = IIf(Val(rstBookPOChild05.Fields("FormFoldRate").Value) > 0, Val(rstBookPOChild05.Fields("FormFoldRate").Value), Val(rstBookPOChild05.Fields("FormStitchRate").Value))
        oExcel.Application.Cells(38, 5).Value = Val(rstBookPOChild05.Fields("FormPasteRate").Value) / 1000
    Else
        oExcel.Application.Cells(39, 5).Value = IIf(Val(rstBookPOChild05.Fields("FormFoldRate").Value) > 0, Val(rstBookPOChild05.Fields("FormFoldRate").Value), Val(rstBookPOChild05.Fields("FormStitchRate").Value))
        oExcel.Application.Cells(41, 5).Value = Val(rstBookPOChild05.Fields("Rate/Book").Value)
    End If
    oExcel.Application.Cells(40, "G").Value = Val(rstBookPOChild05.Fields("Packing&Cartage").Value)
    oExcel.Sheets("Sheet1").Protect ("shpl")
    oExcel.Workbooks.Item(1).Save
    If OutputTo = "S" Then
        oExcel.Range("A1").Activate
        oExcel.Application.Visible = True
    Else
        oExcel.Workbooks.Item(1).PrintOut
    End If
    Set oExcel = Nothing
    On Error GoTo 0
End Sub
Private Function UpdateLastPrinterBinder() As Boolean
    If BookPOType = "F" Then
        CxnBookPrintOrder.Execute "UPDATE BookMaster M INNER JOIN BookPOParent T ON M.Code=T.Book SET M.BookPrinter=T.BookPrinter WHERE M.Code='" & BookCode & "' AND " & _
                                                     "T.Code=(SELECT Top 1 P.Code FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE P.Type='F' AND P.Book=M.Code ORDER BY P.Code DESC)"
        rstBookList.Fields("BookPrinter").Value = BookPrinterCode
        rstBookList.Update
        CxnBookPrintOrder.Execute "UPDATE BookMaster M INNER JOIN BookPOParent T ON M.Code=T.Book SET M.TitlePrinter=T.TitlePrinter WHERE M.Code='" & BookCode & "' AND " & _
                                                     "T.Code=(SELECT Top 1 P.Code FROM BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code WHERE P.Type='F' AND P.Book=M.Code ORDER BY P.Code DESC)"
        rstBookList.Fields("TitlePrinter").Value = TitlePrinterCode
        rstBookList.Update
        CxnBookPrintOrder.Execute "UPDATE BookMaster M INNER JOIN BookPOParent T ON M.Code=T.Book SET M.Laminator=T.Laminator WHERE M.Code='" & BookCode & "' AND " & _
                                                     "T.Code=(SELECT Top 1 P.Code FROM BookPOParent P INNER JOIN BookPOChild07 C ON P.Code=C.Code WHERE P.Type='F' AND P.Book=M.Code ORDER BY P.Code DESC)"
        rstBookList.Fields("Laminator").Value = LaminatorCode
        rstBookList.Update
        CxnBookPrintOrder.Execute "UPDATE BookMaster M INNER JOIN BookPOParent T ON M.Code=T.Book SET M.BinderFresh=T.Binder WHERE M.Code='" & BookCode & "' AND " & _
                                                     "T.Code=(SELECT Top 1 P.Code FROM BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code WHERE P.Type='F' AND P.Book=M.Code ORDER BY P.Code DESC)"
        rstBookList.Fields("BinderFresh").Value = BinderCode
        rstBookList.Update
    ElseIf BookPOType = "R" Then
        CxnBookPrintOrder.Execute "UPDATE BookMaster M INNER JOIN BookPOParent T ON M.Code=T.Book SET M.BinderRepair=T.Binder WHERE M.Code='" & BookCode & "' AND " & _
                                                     "T.Code=(SELECT Top 1 P.Code FROM BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code WHERE P.Type='R' AND P.Book=M.Code ORDER BY P.Code DESC)"
        rstBookList.Fields("BinderRepair").Value = BinderCode
        rstBookList.Update
    End If
End Function
Private Sub CheckCorrections()
    If Not blnRecordExist Then
        If rstCorrections.State = adStateOpen Then rstCorrections.Close
        rstCorrections.Open "SELECT ArrivedOn,Correction,RectifiedOn,SNo FROM BookChild02 WHERE Department='P' AND Code='" & BookCode & "' AND RectifiedOn='' ORDER BY SNo", CxnBookPrintOrder, adOpenKeyset, adLockReadOnly
        rstCorrections.ActiveConnection = Nothing
        If rstCorrections.RecordCount > 0 Then
            Dim i As Integer
            Load FrmCorrectionRegister
            With FrmCorrectionRegister
                .Text2.Text = Text3.Text
                .fpSpread3.ClearRange 1, 1, .fpSpread3.MaxCols, .fpSpread3.MaxRows, True
                i = 1
                rstCorrections.MoveFirst
                Do While Not rstCorrections.EOF
                    .fpSpread3.SetText 1, i, IIf(rstCorrections.Fields("RectifiedOn").Value = "", 0, 1)
                    .fpSpread3.SetText 2, i, Trim(rstCorrections.Fields("Correction").Value)
                    .fpSpread3.SetText 3, i, Format(rstCorrections.Fields("ArrivedOn").Value, "dd-mm-yyyy")
                    .fpSpread3.SetText 4, i, Format(rstCorrections.Fields("SNo").Value, "###########0")
                    i = i + 1
                    rstCorrections.MoveNext
                Loop
                .fpSpread3.SetActiveCell 1, 1
                .Show vbModal
            End With
        End If
    End If
End Sub
Private Sub LockFields(ByVal bVal As Boolean)
    Dim O As Object
    For Each O In Me
        If TypeName(O) = "TextBox" Then O.Locked = bVal
    Next
End Sub
Private Function GetOrderNoRange(ByVal rType As Integer)
    PrintReport = 0
    FrmNumberRange.Report = rType
    Load FrmNumberRange
    FrmNumberRange.Show vbModal
    If FrmNumberRange.Text1.Text <> "" Then LRange = Val(FrmNumberRange.Text1.Text): URange = Val(FrmNumberRange.Text2.Text)
   If FrmNumberRange.Check2.Value = 1 Then PrintReport = Left(FrmNumberRange.Combo1.Text, 1)
   Call CloseForm(FrmNumberRange)
End Function

Private Function GetPlateMaking(ByVal BookCode As String, ByVal OrderCode As String, ByVal OrderDate As String, ByVal OrderType As String, ByVal PlateType As String) As String
Dim i As Integer
    On Error GoTo ErrorHandler
    
    Dim PrintType As String
    Dim NoofPrint As Integer
    
    Dim DatabaseName As String
    Dim CxnImporter As New ADODB.Connection
    Dim rstImporter As New ADODB.Recordset
    Dim rstPSPlateRegister As New ADODB.Recordset
    
    Dim rstPSPlateRegister2 As New ADODB.Recordset
 
     
    DatabaseName = Trim(ReadFromFile("Saral Database Name"))
    CxnImporter.CursorLocation = adUseClient
    If CxnImporter.State = adStateOpen Then
        CxnImporter.Close
    End If
    i = InStr(1, DatabaseName, ",")
    DatabaseName = "Saral.00" & Val(CompCode) - 1
    CxnImporter.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DatabasePath & "\" & DatabaseName & ";Persist Security Info=False;Jet OLEDB:Database Password=RSPLILoveMyINDIA"
     
    rstImporter.Open "SELECT P.Name As OrderNo,C.OrderDate,M1.PrintName As PrinterName,C.Processing,C.PlateType" & PlateType & ",C.ActualQuantity As Quantity,C.PlateRate" & PlateType & " As Rate,C.BillNo,C.BillDate,M2.PrintName,C.Remarks FROM ((BookPOParent P INNER JOIN BookPOChild" & OrderType & " C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON P." & IIf(OrderType = "06", "TitlePrinter", "BookPrinter") & "=M1.Code) INNER JOIN BookMaster M2 ON P.Book=M2.Code WHERE P.Type<>'O' AND LEFT(P.Code,1)<>'*' AND M2.Code='" & BookCode & "' AND C.OrderDate>=#" & GetDate(Format(DateAdd("d", -365, CDate(OrderDate)), "dd-mm-yyyy")) & "# " & _
                                 "ORDER BY M1.PrintName,C.OrderDate", CxnImporter, adOpenKeyset, adLockReadOnly
        
    rstImporter.ActiveConnection = Nothing
    
    
    rstPSPlateRegister.Open "SELECT P.Name As OrderNo,C.OrderDate,M1.PrintName As PrinterName,C.Processing,C.PlateType" & PlateType & ",C.ActualQuantity As Quantity,C.PlateRate" & PlateType & " As Rate,C.BillNo,C.BillDate,M2.PrintName,C.Remarks FROM ((BookPOParent P INNER JOIN BookPOChild" & OrderType & " C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON P." & IIf(OrderType = "06", "TitlePrinter", "BookPrinter") & "=M1.Code) INNER JOIN BookMaster M2 ON P.Book=M2.Code WHERE P.Type<>'O' AND LEFT(P.Code,1)<>'*' AND M2.Code='" & BookCode & "' AND C.Code<'" & IIf(OrderCode = "", "999999", OrderCode) & "' AND C.OrderDate<=#" & GetDate(Format(OrderDate, "dd-mm-yyyy")) & "# " & _
                                 "ORDER BY M1.PrintName,C.OrderDate", CxnDatabase, adOpenKeyset, adLockReadOnly
    

    rstPSPlateRegister.ActiveConnection = Nothing
    
    rstPSPlateRegister2.Open "SELECT P.Name As OrderNo,C.OrderDate,M1.PrintName As PrinterName,C.Processing,C.PlateType" & PlateType & ",C.ActualQuantity As Quantity,C.PlateRate" & PlateType & " As Rate,C.BillNo,C.BillDate,M2.PrintName,C.Remarks FROM ((BookPOParent P INNER JOIN BookPOChild" & OrderType & " C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON P." & IIf(OrderType = "06", "TitlePrinter", "BookPrinter") & "=M1.Code) INNER JOIN BookMaster M2 ON P.Book=M2.Code WHERE P.Type<>'O' AND LEFT(P.Code,1)<>'*' AND M2.Code='" & BookCode & "' AND C.OrderDate>=#" & GetDate(Format(OrderDate, "dd-mm-yyyy")) & "# " & _
                                 "ORDER BY M1.PrintName,C.OrderDate", CxnDatabase, adOpenKeyset, adLockReadOnly
        
    rstPSPlateRegister2.ActiveConnection = Nothing
    
    If rstPSPlateRegister2.Fields("Processing").Value = "N" Then
        PrintType = "Ist Print"
    Else
        
        NoofPrint = (rstImporter.RecordCount) + rstPSPlateRegister.RecordCount + 1
        If NoofPrint = 0 Then
            PrintType = ""
        ElseIf NoofPrint = 1 Then
            PrintType = "Ist Print"
        ElseIf NoofPrint = 2 Then
            PrintType = "2nd Print"
        ElseIf NoofPrint = 3 Then
            PrintType = "3rd Print"
        Else
            PrintType = "Ist Print"
        End If
    End If
    GetPlateMaking = PrintType
     Exit Function
ErrorHandler:

End Function

Private Function GetPaperSend_Balance(ByVal OrderCode As String, ByVal BalFor As String) As String
Dim i As Integer
    On Error GoTo ErrorHandler
    Dim PaperSendQty As String
    Dim PaperBalanceQty As String
    
    Dim rstPaperRegister As New ADODB.Recordset
    If rstPaperRegister.State = adStateOpen Then rstPaperRegister.Close
    
    If BalFor = "P" Then
       rstPaperRegister.Open "SELECT SendQuantity,BalanceQuantity,Paper FROM (PaperPOChildRef T INNER JOIN PaperMaster M1 ON T.Paper=M1.Code) INNER JOIN BookMaster M2 ON T.Book=M2.Code WHERE T.Code In('" & OrderCode & "') ORDER BY M1.Name,M2.Name", CxnDatabase, adOpenKeyset, adLockReadOnly
    Else
       rstPaperRegister.Open "SELECT SendQuantity,BalanceQuantity,Paper FROM (PaperPOChildRef T INNER JOIN PaperMaster M1 ON T.Paper=M1.Code) INNER JOIN BookMaster M2 ON T.Book=M2.Code WHERE T.Code In(Select POCode From BookPOChild05 Where Code='" & OrderCode & "') ORDER BY M1.Name,M2.Name", CxnDatabase, adOpenKeyset, adLockReadOnly
    End If
    
    rstPaperRegister.ActiveConnection = Nothing
     If rstPaperRegister.RecordCount > 0 Then
        PaperCode_For_Balance = rstPaperRegister.Fields("Paper").Value
        
        Do While Not rstPaperRegister.EOF
           PaperSendQty = PaperSendQty & rstPaperRegister.Fields("SendQuantity").Value & ","
           PaperBalanceQty = PaperBalanceQty & rstPaperRegister.Fields("BalanceQuantity").Value & ","
           rstPaperRegister.MoveNext
        Loop
        If Len(PaperSendQty) > 0 Then
           PaperSendQty = Mid(PaperSendQty, 1, Len(PaperSendQty) - 1)
        End If
        If Len(PaperBalanceQty) > 0 Then
           PaperBalanceQty = Mid(PaperBalanceQty, 1, Len(PaperBalanceQty) - 1)
        End If
        If Len(PaperBalanceQty) > 0 Then
            GetPaperSend_Balance = PaperSendQty & "#" & PaperBalanceQty
        Else
            GetPaperSend_Balance = PaperSendQty
        End If
     Else
        
        PaperSendQty = 0
        PaperBalanceQty = 0
        GetPaperSend_Balance = PaperSendQty & "#" & PaperBalanceQty
        
     End If
    Exit Function
ErrorHandler:
End Function
Private Function GetPOPrevious(ByVal PaperCode As String) As String
    On Error GoTo ErrorHandler
    Dim rstPreviousPaperRegister As New ADODB.Recordset
    If rstPreviousPaperRegister.State = adStateOpen Then rstPreviousPaperRegister.Close
    rstPreviousPaperRegister.Open "SELECT Top 1 P.Code,P.Name,P.Date FROM PaperPOChildRef T INNER JOIN PaperPOParent P ON T.Code=P.Code WHERE T.Paper='" & PaperCode & "'  ORDER BY P.Date Desc,P.Code Desc", CxnDatabase, adOpenKeyset, adLockReadOnly
    rstPreviousPaperRegister.ActiveConnection = Nothing
    If rstPreviousPaperRegister.RecordCount > 0 Then GetPOPrevious = rstPreviousPaperRegister.Fields("Code").Value Else GetPOPrevious = ""
    Exit Function
ErrorHandler:
End Function
