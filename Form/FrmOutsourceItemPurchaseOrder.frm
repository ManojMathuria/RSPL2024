VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0F1F1508-C40A-101B-AD04-00AA00575482}#1.0#0"; "mhrinp32.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmOutsourceItemPurchaseOrder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Outsource Item Purchase Order"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9705
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmOutsourceItemPurchaseOrder.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7155
   ScaleWidth      =   9705
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   7095
      Left            =   15
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   0
      Width           =   9660
      _Version        =   65536
      _ExtentX        =   17039
      _ExtentY        =   12515
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
      Picture         =   "FrmOutsourceItemPurchaseOrder.frx":0442
      Begin TabDlg.SSTab SSTab1 
         Height          =   6900
         Left            =   120
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   120
         Width           =   9420
         _ExtentX        =   16616
         _ExtentY        =   12171
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
         TabPicture(0)   =   "FrmOutsourceItemPurchaseOrder.frx":045E
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "DataGrid1"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Text1"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Command1"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).ControlCount=   4
         TabCaption(1)   =   "&Details"
         TabPicture(1)   =   "FrmOutsourceItemPurchaseOrder.frx":047A
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Mh3dFrame2"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         Begin VB.CommandButton Command1 
            DisabledPicture =   "FrmOutsourceItemPurchaseOrder.frx":0496
            Height          =   375
            Left            =   8220
            Picture         =   "FrmOutsourceItemPurchaseOrder.frx":08A8
            Style           =   1  'Graphical
            TabIndex        =   46
            Top             =   6450
            Width           =   1095
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
            MaxLength       =   40
            TabIndex        =   21
            Top             =   6480
            Width           =   7625
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   5955
            Left            =   120
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   450
            Width           =   9195
            _ExtentX        =   16219
            _ExtentY        =   10504
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
            ColumnCount     =   4
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
               DataField       =   "SupplierName"
               Caption         =   "Supplier Name"
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
               DataField       =   "BillAmount"
               Caption         =   "Order Amount"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0.00"
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
                  ColumnWidth     =   1080
               EndProperty
               BeginProperty Column01 
                  Locked          =   -1  'True
                  ColumnWidth     =   1019.906
               EndProperty
               BeginProperty Column02 
                  Locked          =   -1  'True
                  ColumnWidth     =   5235.024
               EndProperty
               BeginProperty Column03 
                  Alignment       =   1
                  Locked          =   -1  'True
                  ColumnWidth     =   1260.284
               EndProperty
            EndProperty
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   6345
            Left            =   -74880
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   480
            Width           =   9195
            _Version        =   65536
            _ExtentX        =   16219
            _ExtentY        =   11192
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
            Picture         =   "FrmOutsourceItemPurchaseOrder.frx":0CBA
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
               Left            =   9600
               MaxLength       =   40
               TabIndex        =   44
               Top             =   120
               Visible         =   0   'False
               Width           =   1440
            End
            Begin VB.TextBox TxtAdNar 
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
               MaxLength       =   254
               TabIndex        =   13
               Top             =   5910
               Width           =   7520
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput15 
               Height          =   330
               Left            =   7560
               TabIndex        =   36
               TabStop         =   0   'False
               Top             =   4860
               Width           =   1520
               _Version        =   65536
               _ExtentX        =   2681
               _ExtentY        =   582
               Calculator      =   "FrmOutsourceItemPurchaseOrder.frx":0CD6
               Caption         =   "FrmOutsourceItemPurchaseOrder.frx":0CF6
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "FrmOutsourceItemPurchaseOrder.frx":0D62
               Keys            =   "FrmOutsourceItemPurchaseOrder.frx":0D80
               Spin            =   "FrmOutsourceItemPurchaseOrder.frx":0DCA
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
               ReadOnly        =   1
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput12 
               Height          =   330
               Left            =   7560
               TabIndex        =   37
               TabStop         =   0   'False
               Top             =   4545
               Width           =   1520
               _Version        =   65536
               _ExtentX        =   2681
               _ExtentY        =   582
               Calculator      =   "FrmOutsourceItemPurchaseOrder.frx":0DF2
               Caption         =   "FrmOutsourceItemPurchaseOrder.frx":0E12
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "FrmOutsourceItemPurchaseOrder.frx":0E7E
               Keys            =   "FrmOutsourceItemPurchaseOrder.frx":0E9C
               Spin            =   "FrmOutsourceItemPurchaseOrder.frx":0EE6
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
               ReadOnly        =   1
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput14 
               Height          =   330
               Left            =   1560
               TabIndex        =   9
               Top             =   4860
               Width           =   1530
               _Version        =   65536
               _ExtentX        =   2699
               _ExtentY        =   582
               Calculator      =   "FrmOutsourceItemPurchaseOrder.frx":0F0E
               Caption         =   "FrmOutsourceItemPurchaseOrder.frx":0F2E
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "FrmOutsourceItemPurchaseOrder.frx":0F9A
               Keys            =   "FrmOutsourceItemPurchaseOrder.frx":0FB8
               Spin            =   "FrmOutsourceItemPurchaseOrder.frx":1002
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
               MinValue        =   -9999999999.99
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput11 
               Height          =   330
               Left            =   1560
               TabIndex        =   8
               Top             =   4545
               Width           =   1530
               _Version        =   65536
               _ExtentX        =   2699
               _ExtentY        =   582
               Calculator      =   "FrmOutsourceItemPurchaseOrder.frx":102A
               Caption         =   "FrmOutsourceItemPurchaseOrder.frx":104A
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "FrmOutsourceItemPurchaseOrder.frx":10B6
               Keys            =   "FrmOutsourceItemPurchaseOrder.frx":10D4
               Spin            =   "FrmOutsourceItemPurchaseOrder.frx":111E
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "#0.00"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "#0.00"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   99.99
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   5
               Value           =   5
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput4 
               Height          =   330
               Left            =   10590
               TabIndex        =   16
               Top             =   750
               Visible         =   0   'False
               Width           =   420
               _Version        =   65536
               _ExtentX        =   741
               _ExtentY        =   582
               Calculator      =   "FrmOutsourceItemPurchaseOrder.frx":1146
               Caption         =   "FrmOutsourceItemPurchaseOrder.frx":1166
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "FrmOutsourceItemPurchaseOrder.frx":11D2
               Keys            =   "FrmOutsourceItemPurchaseOrder.frx":11F0
               Spin            =   "FrmOutsourceItemPurchaseOrder.frx":123A
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
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput3 
               Height          =   330
               Left            =   9990
               TabIndex        =   15
               Top             =   750
               Visible         =   0   'False
               Width           =   615
               _Version        =   65536
               _ExtentX        =   1085
               _ExtentY        =   582
               Calculator      =   "FrmOutsourceItemPurchaseOrder.frx":1262
               Caption         =   "FrmOutsourceItemPurchaseOrder.frx":1282
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "FrmOutsourceItemPurchaseOrder.frx":12EE
               Keys            =   "FrmOutsourceItemPurchaseOrder.frx":130C
               Spin            =   "FrmOutsourceItemPurchaseOrder.frx":1356
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "#######0.0000"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "#######0.0000"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   99999999.9999
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   1179649
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput1 
               Height          =   330
               Left            =   9540
               TabIndex        =   14
               Top             =   750
               Visible         =   0   'False
               Width           =   465
               _Version        =   65536
               _ExtentX        =   820
               _ExtentY        =   582
               Calculator      =   "FrmOutsourceItemPurchaseOrder.frx":137E
               Caption         =   "FrmOutsourceItemPurchaseOrder.frx":139E
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "FrmOutsourceItemPurchaseOrder.frx":140A
               Keys            =   "FrmOutsourceItemPurchaseOrder.frx":1428
               Spin            =   "FrmOutsourceItemPurchaseOrder.frx":1472
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "######0"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "######0"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   9999999
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   1
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
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
               Left            =   1560
               MaxLength       =   10
               TabIndex        =   10
               Top             =   5400
               Width           =   1530
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
               Width           =   900
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
               MaxLength       =   254
               TabIndex        =   5
               Top             =   950
               Width           =   7520
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
               Left            =   1560
               MaxLength       =   40
               TabIndex        =   4
               Top             =   630
               Width           =   7520
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel9 
               Height          =   330
               Left            =   120
               TabIndex        =   24
               Top             =   5400
               Width           =   1455
               _Version        =   65536
               _ExtentX        =   2566
               _ExtentY        =   582
               _StockProps     =   77
               BackColor       =   32896
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
               Caption         =   " Bill No."
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "FrmOutsourceItemPurchaseOrder.frx":149A
               Picture         =   "FrmOutsourceItemPurchaseOrder.frx":14B6
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
               Height          =   330
               Left            =   120
               TabIndex        =   25
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
               Caption         =   " Order No."
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "FrmOutsourceItemPurchaseOrder.frx":14D2
               Picture         =   "FrmOutsourceItemPurchaseOrder.frx":14EE
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Left            =   2450
               TabIndex        =   26
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
               Picture         =   "FrmOutsourceItemPurchaseOrder.frx":150A
               Picture         =   "FrmOutsourceItemPurchaseOrder.frx":1526
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
               Height          =   330
               Left            =   120
               TabIndex        =   27
               Top             =   4545
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
               Caption         =   " GST (%)"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "FrmOutsourceItemPurchaseOrder.frx":1542
               Picture         =   "FrmOutsourceItemPurchaseOrder.frx":155E
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   330
               Left            =   120
               TabIndex        =   28
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
               Caption         =   " Supplier Name"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "FrmOutsourceItemPurchaseOrder.frx":157A
               Picture         =   "FrmOutsourceItemPurchaseOrder.frx":1596
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel13 
               Height          =   330
               Left            =   4540
               TabIndex        =   29
               Top             =   105
               Width           =   1300
               _Version        =   65536
               _ExtentX        =   2293
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
               Caption         =   " Delivery Date"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "FrmOutsourceItemPurchaseOrder.frx":15B2
               Picture         =   "FrmOutsourceItemPurchaseOrder.frx":15CE
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
               Height          =   330
               Left            =   120
               TabIndex        =   30
               Top             =   950
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
               Picture         =   "FrmOutsourceItemPurchaseOrder.frx":15EA
               Picture         =   "FrmOutsourceItemPurchaseOrder.frx":1606
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel10 
               Height          =   330
               Left            =   120
               TabIndex        =   31
               Top             =   4860
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
               Caption         =   " Adjustment"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "FrmOutsourceItemPurchaseOrder.frx":1622
               Picture         =   "FrmOutsourceItemPurchaseOrder.frx":163E
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
               Height          =   330
               Left            =   6360
               TabIndex        =   32
               Top             =   4860
               Width           =   1215
               _Version        =   65536
               _ExtentX        =   2143
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
               Caption         =   " Net Amount"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "FrmOutsourceItemPurchaseOrder.frx":165A
               Picture         =   "FrmOutsourceItemPurchaseOrder.frx":1676
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel16 
               Height          =   330
               Left            =   6360
               TabIndex        =   33
               Top             =   4545
               Width           =   1215
               _Version        =   65536
               _ExtentX        =   2143
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
               Caption         =   " GST"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "FrmOutsourceItemPurchaseOrder.frx":1692
               Picture         =   "FrmOutsourceItemPurchaseOrder.frx":16AE
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel20 
               Height          =   330
               Left            =   6360
               TabIndex        =   34
               Top             =   5400
               Width           =   1215
               _Version        =   65536
               _ExtentX        =   2143
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
               Caption         =   " Paid Amount"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "FrmOutsourceItemPurchaseOrder.frx":16CA
               Picture         =   "FrmOutsourceItemPurchaseOrder.frx":16E6
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel21 
               Height          =   330
               Left            =   3075
               TabIndex        =   35
               Top             =   5400
               Width           =   1620
               _Version        =   65536
               _ExtentX        =   2857
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
               Caption         =   " Bill Date"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "FrmOutsourceItemPurchaseOrder.frx":1702
               Picture         =   "FrmOutsourceItemPurchaseOrder.frx":171E
            End
            Begin TDBDate6Ctl.TDBDate MhDateInput1 
               Height          =   330
               Left            =   3530
               TabIndex        =   1
               Top             =   105
               Width           =   1050
               _Version        =   65536
               _ExtentX        =   1861
               _ExtentY        =   582
               Calendar        =   "FrmOutsourceItemPurchaseOrder.frx":173A
               Caption         =   "FrmOutsourceItemPurchaseOrder.frx":1852
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "FrmOutsourceItemPurchaseOrder.frx":18BE
               Keys            =   "FrmOutsourceItemPurchaseOrder.frx":18DC
               Spin            =   "FrmOutsourceItemPurchaseOrder.frx":193A
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
            Begin TDBDate6Ctl.TDBDate MhDateInput3 
               Height          =   330
               Left            =   5835
               TabIndex        =   2
               Top             =   105
               Width           =   1050
               _Version        =   65536
               _ExtentX        =   1861
               _ExtentY        =   582
               Calendar        =   "FrmOutsourceItemPurchaseOrder.frx":1962
               Caption         =   "FrmOutsourceItemPurchaseOrder.frx":1A7A
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "FrmOutsourceItemPurchaseOrder.frx":1AE6
               Keys            =   "FrmOutsourceItemPurchaseOrder.frx":1B04
               Spin            =   "FrmOutsourceItemPurchaseOrder.frx":1B62
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
            Begin TDBDate6Ctl.TDBDate MhDateInput2 
               Height          =   330
               Left            =   4680
               TabIndex        =   11
               Top             =   5400
               Width           =   1695
               _Version        =   65536
               _ExtentX        =   2990
               _ExtentY        =   582
               Calendar        =   "FrmOutsourceItemPurchaseOrder.frx":1B8A
               Caption         =   "FrmOutsourceItemPurchaseOrder.frx":1CA2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "FrmOutsourceItemPurchaseOrder.frx":1D0E
               Keys            =   "FrmOutsourceItemPurchaseOrder.frx":1D2C
               Spin            =   "FrmOutsourceItemPurchaseOrder.frx":1D8A
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
            Begin TDBNumber6Ctl.TDBNumber MhRealInput16 
               Height          =   330
               Left            =   7560
               TabIndex        =   12
               Top             =   5400
               Width           =   1520
               _Version        =   65536
               _ExtentX        =   2681
               _ExtentY        =   582
               Calculator      =   "FrmOutsourceItemPurchaseOrder.frx":1DB2
               Caption         =   "FrmOutsourceItemPurchaseOrder.frx":1DD2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "FrmOutsourceItemPurchaseOrder.frx":1E3E
               Keys            =   "FrmOutsourceItemPurchaseOrder.frx":1E5C
               Spin            =   "FrmOutsourceItemPurchaseOrder.frx":1EA6
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
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel32 
               Height          =   330
               Left            =   120
               TabIndex        =   38
               Top             =   5910
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
               Caption         =   " Adj.Remarks"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "FrmOutsourceItemPurchaseOrder.frx":1ECE
               Picture         =   "FrmOutsourceItemPurchaseOrder.frx":1EEA
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel17 
               Height          =   285
               Left            =   120
               TabIndex        =   39
               Top             =   4080
               Width           =   8955
               _Version        =   65536
               _ExtentX        =   15796
               _ExtentY        =   503
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
               Caption         =   ""
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "FrmOutsourceItemPurchaseOrder.frx":1F06
               Picture         =   "FrmOutsourceItemPurchaseOrder.frx":1F22
               Begin TDBNumber6Ctl.TDBNumber MhRealInput20 
                  Height          =   285
                  Left            =   7680
                  TabIndex        =   40
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   1040
                  _Version        =   65536
                  _ExtentX        =   1834
                  _ExtentY        =   503
                  Calculator      =   "FrmOutsourceItemPurchaseOrder.frx":1F3E
                  Caption         =   "FrmOutsourceItemPurchaseOrder.frx":1F5E
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "FrmOutsourceItemPurchaseOrder.frx":1FCA
                  Keys            =   "FrmOutsourceItemPurchaseOrder.frx":1FE8
                  Spin            =   "FrmOutsourceItemPurchaseOrder.frx":2032
                  AlignHorizontal =   1
                  AlignVertical   =   0
                  Appearance      =   0
                  BackColor       =   16777215
                  BorderStyle     =   1
                  BtnPositioning  =   0
                  ClipMode        =   0
                  ClearAction     =   0
                  DecimalPoint    =   "."
                  DisplayFormat   =   "#########0.000"
                  EditMode        =   1
                  Enabled         =   -1
                  ErrorBeep       =   0
                  ForeColor       =   255
                  Format          =   "#########0.00"
                  HighlightText   =   0
                  MarginBottom    =   1
                  MarginLeft      =   1
                  MarginRight     =   1
                  MarginTop       =   1
                  MaxValue        =   999999.999
                  MinValue        =   0
                  MousePointer    =   0
                  MoveOnLRKey     =   0
                  NegativeColor   =   255
                  OLEDragMode     =   0
                  OLEDropMode     =   0
                  ReadOnly        =   1
                  Separator       =   ""
                  ShowContextMenu =   1
                  ValueVT         =   5
                  Value           =   0
                  MaxValueVT      =   5
                  MinValueVT      =   5
               End
            End
            Begin MhinrelLib.MhRealInput MhRealInput6 
               Height          =   255
               Left            =   7230
               TabIndex        =   41
               TabStop         =   0   'False
               Top             =   2520
               Width           =   1590
               _Version        =   65536
               _ExtentX        =   2805
               _ExtentY        =   450
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
               Text            =   "0.000"
               MaxReal         =   9999999
               MinReal         =   0
               ReadOnly        =   -1  'True
               SpinChangeReal  =   0
               CaretColor      =   -2147483642
               DecimalPlaces   =   2
               VAlignment      =   2
            End
            Begin MhinrelLib.MhRealInput MhRealInput5 
               Height          =   255
               Left            =   4980
               TabIndex        =   42
               TabStop         =   0   'False
               Top             =   2520
               Width           =   1185
               _Version        =   65536
               _ExtentX        =   2090
               _ExtentY        =   450
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
               Text            =   "0.000"
               MaxReal         =   9999999
               MinReal         =   0
               ReadOnly        =   -1  'True
               SpinChangeReal  =   0
               CaretColor      =   -2147483642
               DecimalPlaces   =   3
               VAlignment      =   2
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel19 
               Height          =   255
               Left            =   120
               TabIndex        =   43
               Top             =   2520
               Width           =   8960
               _Version        =   65536
               _ExtentX        =   15804
               _ExtentY        =   450
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
               Picture         =   "FrmOutsourceItemPurchaseOrder.frx":205A
               Picture         =   "FrmOutsourceItemPurchaseOrder.frx":2076
            End
            Begin FPSpreadADO.fpSpread fpSpread1 
               Height          =   1095
               Left            =   120
               TabIndex        =   6
               Top             =   1440
               Width           =   8955
               _Version        =   524288
               _ExtentX        =   15796
               _ExtentY        =   1931
               _StockProps     =   64
               EditEnterAction =   5
               EditModeReplace =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaxCols         =   5
               MaxRows         =   1000
               ScrollBars      =   2
               SpreadDesigner  =   "FrmOutsourceItemPurchaseOrder.frx":2092
            End
            Begin FPSpreadADO.fpSpread fpSpread2 
               Height          =   1095
               Left            =   120
               TabIndex        =   7
               Top             =   3000
               Width           =   8955
               _Version        =   524288
               _ExtentX        =   15796
               _ExtentY        =   1931
               _StockProps     =   64
               EditEnterAction =   5
               EditModeReplace =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaxCols         =   5
               MaxRows         =   1000
               ScrollBars      =   2
               SpreadDesigner  =   "FrmOutsourceItemPurchaseOrder.frx":27DD
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel131 
               Height          =   330
               Left            =   6870
               TabIndex        =   45
               Top             =   105
               Width           =   1190
               _Version        =   65536
               _ExtentX        =   2099
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
               Caption         =   " Extend Date"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "FrmOutsourceItemPurchaseOrder.frx":2EDC
               Picture         =   "FrmOutsourceItemPurchaseOrder.frx":2EF8
            End
            Begin TDBDate6Ctl.TDBDate MhDateInput31 
               Height          =   330
               Left            =   8020
               TabIndex        =   3
               Top             =   105
               Width           =   1050
               _Version        =   65536
               _ExtentX        =   1861
               _ExtentY        =   582
               Calendar        =   "FrmOutsourceItemPurchaseOrder.frx":2F14
               Caption         =   "FrmOutsourceItemPurchaseOrder.frx":302C
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "FrmOutsourceItemPurchaseOrder.frx":3098
               Keys            =   "FrmOutsourceItemPurchaseOrder.frx":30B6
               Spin            =   "FrmOutsourceItemPurchaseOrder.frx":3114
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
            Begin VB.Line Line3 
               Index           =   1
               X1              =   0
               X2              =   9180
               Y1              =   2880
               Y2              =   2880
            End
            Begin VB.Line Line6 
               X1              =   0
               X2              =   11560
               Y1              =   0
               Y2              =   0
            End
            Begin VB.Line Line4 
               X1              =   0
               X2              =   9180
               Y1              =   5820
               Y2              =   5820
            End
            Begin VB.Line Line5 
               X1              =   0
               X2              =   9180
               Y1              =   5295
               Y2              =   5295
            End
            Begin VB.Line Line1 
               X1              =   0
               X2              =   9180
               Y1              =   525
               Y2              =   525
            End
            Begin VB.Line Line2 
               X1              =   0
               X2              =   9180
               Y1              =   1370
               Y2              =   1370
            End
            Begin VB.Line Line3 
               Index           =   0
               X1              =   0
               X2              =   9180
               Y1              =   4440
               Y2              =   4440
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
            TabIndex        =   22
            Top             =   6480
            Width           =   495
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   9705
      _ExtentX        =   17119
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
Attribute VB_Name = "FrmOutsourceItemPurchaseOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CxnOutsourceItemPurchaseOrder As New ADODB.Connection
Dim rstCompanyMaster As New ADODB.Recordset
Dim rstOutsourceItemPOList As New ADODB.Recordset
Dim rstOutsourceItemPOParent As New ADODB.Recordset
Dim rstOutsourceItemPOChild As New ADODB.Recordset
Dim rstAccountList As New ADODB.Recordset
Dim rstSupplierList As New ADODB.Recordset
Dim rstOutsourceItemList As New ADODB.Recordset
Dim rstLastPurchaseRate As New ADODB.Recordset
Dim SupplierCode As String
Dim OutsourceItemCode As String
Dim AccountCode As String
Dim SortOrder As String
Dim PrevStr As String
Dim dblBookMark As Double
Dim blnRecordExist As Boolean
Dim EMailID As String
Dim Attachment As String
Dim Message As String
Dim OutputTo As String
Dim EditMode As Boolean
Dim rstMaterialIOList As New ADODB.Recordset
Dim rstMaterialIOParent As New ADODB.Recordset
Dim rstMaterialIOChild As New ADODB.Recordset
Dim RefCode As String
Dim RefCodeAndQty As String
Dim oOutlook As New Outlook.Application

Private Sub Command1_Click()
    Call Export2Excel
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandler
    CenterForm Me
    WheelHook DataGrid1
    BusySystemIndicator True
    CxnOutsourceItemPurchaseOrder.CursorLocation = adUseClient
    CxnOutsourceItemPurchaseOrder.Open CxnDatabase.ConnectionString
    rstCompanyMaster.Open "Select PrintName, Address1, Address2, Address3, Address4, Phone, Fax, EMail, Website From CompanyMaster", CxnOutsourceItemPurchaseOrder, adOpenKeyset, adLockReadOnly
    rstOutsourceItemList.Open "Select Name As Col0,Code From OutsourceItemMaster Order By Name", CxnOutsourceItemPurchaseOrder, adOpenKeyset, adLockReadOnly
    rstAccountList.Open "Select Name As Col0,Code From AccountMaster Where Type In ('08','09') Order By Name", CxnOutsourceItemPurchaseOrder, adOpenKeyset, adLockReadOnly
    rstSupplierList.Open "Select Name As Col0, Code From AccountMaster Where Type = '01' Order By Name", CxnOutsourceItemPurchaseOrder, adOpenKeyset, adLockReadOnly
    rstOutsourceItemPOList.Open "Select T.Code,T.Name,T.Date,M.Name As SupplierName,T.BillAmount From OutsourceItemPOParent T,AccountMaster M Where T.Supplier = M.Code Order By T.Name", CxnOutsourceItemPurchaseOrder, adOpenKeyset, adLockOptimistic
    rstOutsourceItemPOParent.CursorLocation = adUseClient
    rstMaterialIOParent.CursorLocation = adUseClient
    rstOutsourceItemPOList.Filter = adFilterNone
    If rstOutsourceItemPOList.RecordCount > 0 Then rstOutsourceItemPOList.MoveLast
    Set DataGrid1.DataSource = rstOutsourceItemPOList
    BusySystemIndicator False
    SSTab1.Tab = 0
    SortOrder = "Name"
    
    If Not (rstOutsourceItemPOList.EOF Or rstOutsourceItemPOList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    rstOutsourceItemPOList.ActiveConnection = Nothing
    rstAccountList.ActiveConnection = Nothing
    rstOutsourceItemList.ActiveConnection = Nothing
    rstSupplierList.ActiveConnection = Nothing
    SetButtonsForNoRecord
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    Unload Me
End Sub
Private Sub Form_Activate()
    EnableChildMenu True, True
    Text1.SetFocus
End Sub
Private Sub Form_Deactivate()
    DisableChildMenu
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If Shift = 0 And KeyCode = vbKeyEscape Then
        If SSTab1.Tab = 0 Then
            Unload Me
        Else
            If Toolbar1.Buttons.Item(1).Enabled Then
                SSTab1.Tab = 0
            Else
                If Me.ActiveControl.Name <> "Text5" And Me.ActiveControl.Name <> "MhRealInput1" And Me.ActiveControl.Name <> "MhRealInput3" And Me.ActiveControl.Name <> "MhRealInput4" Then
                    If MsgBox("Are you sure to Quit?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Quit !") <> vbYes Then
                        Me.ActiveControl.SetFocus
                    Else
                        Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
                    End If
                End If
            End If
            If Me.ActiveControl.Name <> "Text5" And Me.ActiveControl.Name <> "MhRealInput1" And Me.ActiveControl.Name <> "MhRealInput3" And Me.ActiveControl.Name <> "MhRealInput4" Then
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
        If Me.ActiveControl.Name <> "Text5" And Me.ActiveControl.Name <> "MhRealInput1" And Me.ActiveControl.Name <> "MhRealInput3" And Me.ActiveControl.Name <> "MhRealInput4" Then
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
        
      ElseIf Shift = vbAltMask And KeyCode = vbKeyM And Toolbar1.Buttons.Item(1).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(11)
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
''*******Added By Shamshad
        If Toolbar1.Buttons.Item(1).Enabled Then
            SSTab1.Tab = 1: SSTab1.SetFocus
        Else
           If Me.ActiveControl.Name <> "fpSpread1" And Me.ActiveControl.Name <> "fpSpread2" Then SendKeys "{TAB}"
        End If
        If Me.ActiveControl.Name <> "fpSpread1" And Me.ActiveControl.Name <> "fpSpread2" Then KeyCode = 0
   End If
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Toolbar1.Buttons.Item(4).Enabled Then
        Call Form_KeyDown(vbKeyEscape, 0)
        Cancel = 1
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    WheelUnHook
    Call CloseRecordset(rstCompanyMaster)
    Call CloseRecordset(rstOutsourceItemPOList)
    Call CloseRecordset(rstOutsourceItemPOParent)
    Call CloseRecordset(rstMaterialIOParent)
    Call CloseRecordset(rstOutsourceItemPOChild)
    Call CloseRecordset(rstMaterialIOChild)
    Call CloseRecordset(rstOutsourceItemList)
    Call CloseRecordset(rstSupplierList)
    Call CloseRecordset(rstLastPurchaseRate)
    Call CloseRecordset(rstAccountList)
    Call CloseConnection(CxnOutsourceItemPurchaseOrder)
    ShowProgressInStatusBar False
    DisableChildMenu
End Sub
Private Sub Export2Excel()
On Error GoTo er
Dim oExcel As Object
Set oExcel = CreateObject("Excel.Application")
Dim oWorkBook As Object
Dim oWorkSheet As Object
Dim i As Integer, K As Integer, M As Integer, j As Integer
Dim lRow As Long
Dim LastRow As Long
Dim LastCol As Long
oExcel.Visible = False

If rstOutsourceItemPOList.RecordCount = 0 Then On Error GoTo 0: Screen.MousePointer = vbNormal: Exit Sub
If Not FileExist(App.Path & "\Book1.xlsx") Then DisplayError ("Template File Missing"): Exit Sub
Set oExcel = CreateObject("Excel.Application")
oExcel.Workbooks.Open App.Path & "\Template\Blank.xlsx"
Set oWorkSheet = oExcel.Workbooks("Blank.xlsx").Sheets("Sheet1")
 oWorkSheet.Cells(1, 1).Value = "Code"
 oWorkSheet.Cells(1, 2).Value = "Order No"
 oWorkSheet.Cells(1, 3).Value = "Order Date"
 oWorkSheet.Cells(1, 4).Value = "Supplier Name"
 oWorkSheet.Cells(1, 5).Value = "Order Amount"
i = 2
rstOutsourceItemPOList.MoveFirst
Do While Not rstOutsourceItemPOList.EOF
    oWorkSheet.Cells(i, "A").Value = rstOutsourceItemPOList.Fields("Code").Value
    oWorkSheet.Cells(i, "B").Value = rstOutsourceItemPOList.Fields("Name").Value
    oWorkSheet.Cells(i, "C").Value = Format(rstOutsourceItemPOList.Fields("Date").Value, "dd/MM/yyyy")
    oWorkSheet.Cells(i, "D").Value = rstOutsourceItemPOList.Fields("SupplierName").Value
    oWorkSheet.Cells(i, "E").Value = rstOutsourceItemPOList.Fields("BillAmount").Value
 i = i + 1
    rstOutsourceItemPOList.MoveNext
Loop

oExcel.Range("A:A").EntireColumn.Hidden = True

Screen.MousePointer = vbNormal
oExcel.Sheets("Sheet1").Activate
oExcel.Columns("A:L").EntireColumn.AutoFit
oExcel.Workbooks.Item(1).Save
oExcel.Visible = True
Set oExcel = Nothing
er:
If Err.Number = 1004 Then
Exit Sub
End If
End Sub
Private Sub Text1_Change()
    If rstOutsourceItemPOList.RecordCount = 0 Then Exit Sub
    rstOutsourceItemPOList.MoveFirst
    If Text1.Text <> "" Then
        If SortOrder = "Name" Then
           rstOutsourceItemPOList.Find "[" & SortOrder & "] Like '%" & FixQuote(Text1.Text) & "%'"
        Else
           rstOutsourceItemPOList.Find "[" & SortOrder & "] Like '" & FixQuote(Text1.Text) & "%'"
        End If
        If rstOutsourceItemPOList.EOF Then
            rstOutsourceItemPOList.MoveFirst
            If PrevStr <> "" And Len(Text1.Text) > 1 Then
                If dblBookMark <> 0 Then
                    rstOutsourceItemPOList.Bookmark = dblBookMark
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
    If Not (rstOutsourceItemPOList.EOF Or rstOutsourceItemPOList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim KeyProcessed As Boolean
    
    If rstOutsourceItemPOList.RecordCount = 0 Then Exit Sub
    If Shift = 0 And KeyCode = vbKeyUp Then
        With rstOutsourceItemPOList
            .MovePrevious
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyBack Then
        With rstOutsourceItemPOList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyDown Then
        With rstOutsourceItemPOList
            .MoveNext
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageUp Then
        With rstOutsourceItemPOList
            .Move (-1) * (DataGrid1.VisibleRows - 1)
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageUp Then
        With rstOutsourceItemPOList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageDown Then
        With rstOutsourceItemPOList
            .Move DataGrid1.VisibleRows - 1
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageDown Then
        With rstOutsourceItemPOList
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
            If Not (rstOutsourceItemPOList.EOF Or rstOutsourceItemPOList.BOF) Then
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
    Dim CellVal01 As Variant, CellVal02 As Variant, CellVal03 As Variant, i As Integer
    If Button.Index = 1 Then
        If rstOutsourceItemPOParent.State = adStateOpen Then
           rstOutsourceItemPOParent.Close
        End If
       
       If rstMaterialIOParent.State = adStateOpen Then
           rstMaterialIOParent.Close
        End If
                
        rstOutsourceItemPOParent.Open "Select * From OutsourceItemPOParent Where Code = ''", CxnOutsourceItemPurchaseOrder, adOpenKeyset, adLockOptimistic
        rstMaterialIOParent.Open "Select * From MaterialIOParent Where Code = ''", CxnOutsourceItemPurchaseOrder, adOpenKeyset, adLockOptimistic
        
        ClearFields ("P")
        ClearFields ("C")
        Call LoadOutsourceItemList("")
        If (rstOutsourceItemPOChild.State = adStateClosed) And (rstMaterialIOChild.State = adStateClosed) Then
            SSTab1.Tab = 0
            Exit Sub
        End If
        If AddRecord(rstOutsourceItemPOParent) Then
            Text2.Text = GenerateCode(CxnOutsourceItemPurchaseOrder, "Select Max(Val(Name)) From OutsourceItemPOParent", 10, Space(1))
            MhDateInput1.Text = Format(Date, "dd-MM-yyyy")
            Call SetButtons(False)
            SSTab1.Tab = 1
            Text2.SetFocus
            blnRecordExist = False
            CxnOutsourceItemPurchaseOrder.BeginTrans
        End If
    ElseIf Button.Index = 2 Then
        If rstOutsourceItemPOList.RecordCount = 0 Then Exit Sub
        SSTab1.Tab = 1
        EditRecord
    ElseIf Button.Index = 3 Then
        If rstOutsourceItemPOList.RecordCount = 0 Then Exit Sub
        If AllowTransactionsDeletion = 0 Then
            Call DisplayError("You don't have the rights to Delete this Voucher")
            Exit Sub
        End If
        SSTab1.Tab = 1
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            On Error Resume Next
            MdiMainMenu.MousePointer = vbHourglass
            CxnOutsourceItemPurchaseOrder.Execute "Delete From OutsourceItemPOParent Where Code = '" & rstOutsourceItemPOList.Fields("Code").Value & "'"
            CxnOutsourceItemPurchaseOrder.Execute "Delete From OutsourceItemPOChild Where Code = '" & rstOutsourceItemPOList.Fields("Code").Value & "'"
            CxnOutsourceItemPurchaseOrder.Execute "Delete From MaterialIOParent Where Code = '" & rstOutsourceItemPOList.Fields("Code").Value & "'"
            MdiMainMenu.MousePointer = vbNormal
            If Err.Number = 0 Then
                rstOutsourceItemPOList.Delete
                rstOutsourceItemPOList.MoveNext
                If rstOutsourceItemPOList.RecordCount > 0 And rstOutsourceItemPOList.EOF Then
                    rstOutsourceItemPOList.MoveLast
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
        If (UpdateRecord(rstOutsourceItemPOParent)) And (UpdateRecord(rstMaterialIOParent)) Then
            If UpdateOutsourceItemList("D") Then
                UpdateFlag = 1
                With fpSpread1
                    For i = 1 To .DataRowCnt
                        .SetActiveCell 4, i
                        .GetText 4, i, CellVal01
                        .GetText 5, i, CellVal02
                        If Val(CellVal01) <> 0 And CellVal02 <> "" Then
                            If Not UpdateOutsourceItemList("I1") Then UpdateFlag = 0: Exit For
                        End If
                    Next
                End With
                If UpdateFlag = 1 Then
                    With fpSpread2
                        For i = 1 To .DataRowCnt
                            .SetActiveCell 3, i
                            .GetText 3, i, CellVal01
                            .GetText 4, i, CellVal02
                            .GetText 5, i, CellVal03
                            If Val(CellVal01) <> 0 And CellVal02 <> "" And CellVal03 <> "" Then
                                If Not UpdateOutsourceItemList("I2") Then UpdateFlag = 0: Exit For
                            End If
                        Next
                    End With
                End If
            End If
        End If
        If UpdateFlag Then
            AddToList
            CxnOutsourceItemPurchaseOrder.CommitTrans
            If rstOutsourceItemPOParent.State = adStateOpen Then
                rstOutsourceItemPOParent.Close
            End If
            rstOutsourceItemPOParent.CursorLocation = adUseClient
            Call SetButtons(True)
            SSTab1.Tab = 0
            ShowProgressInStatusBar True
            Timer1.Enabled = True
            LockFields (False)
        Else
            DisplayError ("Failed to save the record")
            Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
        End If
    ElseIf Button.Index = 5 Then
        If CancelRecordUpdate(rstOutsourceItemPOParent) Then
            CxnOutsourceItemPurchaseOrder.RollbackTrans
            If rstOutsourceItemPOParent.State = adStateOpen Then
                rstOutsourceItemPOParent.Close
            End If
            rstOutsourceItemPOParent.CursorLocation = adUseClient
            Call SetButtons(True)
            SetButtonsForNoRecord
            SSTab1.Tab = 0
            LockFields (False)
        End If
    ElseIf Button.Index = 6 Then
        SSTab1.Tab = 0
        Set DataGrid1.DataSource = Nothing
        rstOutsourceItemPOList.ActiveConnection = CxnOutsourceItemPurchaseOrder
        Do While Not RefreshRecord(rstOutsourceItemPOList)
        Loop
        Set DataGrid1.DataSource = rstOutsourceItemPOList
        rstOutsourceItemPOList.ActiveConnection = Nothing
        If rstOutsourceItemPOList.RecordCount > 0 Then rstOutsourceItemPOList.MoveLast
        rstSupplierList.ActiveConnection = CxnOutsourceItemPurchaseOrder
        Do While Not RefreshRecord(rstSupplierList)
        Loop
        rstSupplierList.ActiveConnection = Nothing
        rstOutsourceItemList.ActiveConnection = CxnOutsourceItemPurchaseOrder
        Do While Not RefreshRecord(rstOutsourceItemList)
        Loop
        rstOutsourceItemList.ActiveConnection = Nothing
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
        If rstOutsourceItemPOList.RecordCount = 0 Then Exit Sub
        OutputTo = "P"
        Call PrintOutsourceItemPurchaseOrder(rstOutsourceItemPOList.Fields("Code").Value, "P")
        HiLiteRecord = True
    ElseIf Button.Index = 10 Then
        If rstOutsourceItemPOList.RecordCount = 0 Then Exit Sub
        OutputTo = "S"
        Call PrintOutsourceItemPurchaseOrder(rstOutsourceItemPOList.Fields("Code").Value, "S")
        HiLiteRecord = True
    ElseIf Button.Index = 11 Then
        If rstOutsourceItemPOList.RecordCount = 0 Then Exit Sub
        Call PrintOutsourceItemPurchaseOrder(rstOutsourceItemPOList.Fields("Code").Value, "M")
        HiLiteRecord = True
      
    ElseIf Button.Index = 13 Then
        If rstOutsourceItemPOList.RecordCount > 0 Then rstOutsourceItemPOList.MoveFirst
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 14 Then
        If rstOutsourceItemPOList.RecordCount > 0 Then
            rstOutsourceItemPOList.MovePrevious
            If rstOutsourceItemPOList.BOF Then
                rstOutsourceItemPOList.MoveNext
            End If
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 15 Then
        If rstOutsourceItemPOList.RecordCount > 0 Then
            rstOutsourceItemPOList.MoveNext
            If rstOutsourceItemPOList.EOF Then
                rstOutsourceItemPOList.MovePrevious
            End If
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 16 Then
        If rstOutsourceItemPOList.RecordCount > 0 Then rstOutsourceItemPOList.MoveLast
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 18 Then
        Unload Me
        HiLiteRecord = False
    End If
    If HiLiteRecord Then
        If Not (rstOutsourceItemPOList.EOF Or rstOutsourceItemPOList.BOF) Then
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
          rstOutsourceItemPOList.Sort = "Name Asc"
       End If
    ElseIf ColIndex = 2 Then
       If SortOrder <> "SupplierName" Then
          SortOrder = "SupplierName"
          rstOutsourceItemPOList.Sort = "SupplierName Asc"
       End If
    End If
    DataGrid1.ClearSelCols
    If Not (rstOutsourceItemPOList.EOF Or rstOutsourceItemPOList.BOF) Then
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
    Toolbar1.Buttons.Item(11).Enabled = bVal
    Toolbar1.Buttons.Item(13).Enabled = bVal
    Toolbar1.Buttons.Item(14).Enabled = bVal
    Toolbar1.Buttons.Item(15).Enabled = bVal
    Toolbar1.Buttons.Item(16).Enabled = bVal
    Toolbar1.Buttons.Item(18).Enabled = bVal
    Mh3dFrame2.Enabled = Not bVal
End Sub
Private Sub SetButtonsForNoRecord()
    If rstOutsourceItemPOList.RecordCount = 0 Then
        Toolbar1.Buttons.Item(2).Enabled = False
        Toolbar1.Buttons.Item(3).Enabled = False
        Toolbar1.Buttons.Item(9).Enabled = False
        Toolbar1.Buttons.Item(10).Enabled = False
        Toolbar1.Buttons.Item(11).Enabled = False
        Toolbar1.Buttons.Item(13).Enabled = False
        Toolbar1.Buttons.Item(14).Enabled = False
        Toolbar1.Buttons.Item(15).Enabled = False
        Toolbar1.Buttons.Item(16).Enabled = False
    End If
End Sub
Private Sub Text2_Validate(Cancel As Boolean)
    If rstOutsourceItemPOParent.EOF Or rstOutsourceItemPOParent.BOF Then Exit Sub
    If CheckEmpty(Text2, True) Then
        Cancel = True
    ElseIf CheckDuplicate(CxnOutsourceItemPurchaseOrder, "OutsourceItemPOParent", "Code", "[Name]", Trim(Text2.Text), rstOutsourceItemPOParent.Fields("Code").Value, False) Then
        Cancel = True
    End If
End Sub
Private Sub MhDateInput1_Validate(Cancel As Boolean)
    If Not IsDate(GetDate(MhDateInput1.Text)) Then
        Cancel = True
    ElseIf Format(GetDate(MhDateInput1.Text), "yyyymmdd") < Format(FinancialYearFrom, "yyyymmdd") Or Format(GetDate(MhDateInput1.Text), "yyyymmdd") > Format(FinancialYearTo, "yyyymmdd") Then
        Cancel = True
    ElseIf Not blnRecordExist Then
        MhDateInput3.Text = Format(DateAdd("d", 10, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
    End If
End Sub
Private Sub MhDateInput2_Validate(Cancel As Boolean)
    If MhDateInput2.ValueIsNull Then Exit Sub
    If Not IsDate(GetDate(MhDateInput2.Text)) Then
        Cancel = True
'    ElseIf Format(GetDate(MhDateInput2.Text), "yyyymmdd") < Format(FinancialYearFrom, "yyyymmdd") Or Format(GetDate(MhDateInput2.Text), "yyyymmdd") > Format(FinancialYearTo, "yyyymmdd") Then
'        Cancel = True
    End If
End Sub
Private Sub Text3_Change()
    If Text3.Text = " " Then
        Text3.Text = "?"
        SendKeys "{TAB}"
    End If
End Sub
Private Sub Text3_Validate(Cancel As Boolean)
    Dim SearchString As String
    
    SearchString = FixQuote(Text3.Text)
    If rstSupplierList.RecordCount = 0 Then
        DisplayError ("No Record in Supplier Master")
        Cancel = True
        Exit Sub
    Else
        rstSupplierList.MoveFirst
    End If
    rstSupplierList.Find "[Col0] = '" & RTrim(SearchString) & "'"
    If rstSupplierList.EOF Then
        SelectionType = "S"
        SupplierCode = ""
        Call LoadSelectionList(rstSupplierList, "List of Suppliers...", "Name")
        SearchOrder = 0
        Call DisplaySelectionList(Text3, SupplierCode)
        Call CloseForm(FrmSelectionList)
        If CheckEmpty(Text3.Text, False) Then
            Text3.Text = "?"
        End If
        If RTrim(SupplierCode) <> "" Then
            SendKeys "{TAB}"
        End If
        Cancel = True
    Else
        SupplierCode = rstSupplierList.Fields("Code").Value
    End If
End Sub

Private Sub Text4_Validate(Cancel As Boolean)

'    If rstOutsourceItemPOChild.RecordCount = 0 Then
'        SendKeys "^"
'        Call AddRecord(rstOutsourceItemPOChild)
'        Call ClearFields("C")
'        'Call DataGrid2_KeyDown(vbKeyE, vbCtrlMask)
'
'    End If
End Sub
Private Sub MhDateInput3_Validate(Cancel As Boolean)
    If Not IsDate(GetDate(MhDateInput3.Text)) Then
        Cancel = True
    End If
End Sub

Private Sub MhRealInput11_Validate(Cancel As Boolean)
    MhRealInput12.Text = Format(Val(MhRealInput6.Text) * Val(MhRealInput11.Text) / 100, "0.00")
    Call CalculateTotal("N")
End Sub
Private Sub MhRealInput14_Validate(Cancel As Boolean)
    Call CalculateTotal("N")
End Sub
Private Sub ViewRecord()
    ClearFields ("P")
    ClearFields ("C")
    If rstOutsourceItemPOList.EOF Then
        If rstOutsourceItemPOChild.State = adStateOpen Then
            rstOutsourceItemPOChild.Close
        End If
        Exit Sub
    End If
    FindRecord
    LoadFields
End Sub
Private Sub FindRecord()
    
    If rstOutsourceItemPOParent.State = adStateOpen Then
       rstOutsourceItemPOParent.Close
    End If
    rstOutsourceItemPOParent.Open "Select * From OutsourceItemPOParent Where Code = '" & FixQuote(rstOutsourceItemPOList.Fields("Code").Value) & "'", CxnOutsourceItemPurchaseOrder, adOpenKeyset, adLockOptimistic
    
    ''*********Added By Shamshad Alam
    If rstMaterialIOParent.State = adStateOpen Then
       rstMaterialIOParent.Close
    End If
    rstMaterialIOParent.Open "Select * From MaterialIOParent Where Code = '" & FixQuote(rstOutsourceItemPOList.Fields("Code").Value) & "'", CxnOutsourceItemPurchaseOrder, adOpenKeyset, adLockOptimistic
    ''*******End *******************************
    'If (rstOutsourceItemPOParent.RecordCount = 0) And (rstMaterialIOParent.RecordCount = 0) Then
    If rstOutsourceItemPOParent.RecordCount = 0 Then
       Call DisplayError("This Record has been deleted by Another User ! Click Ok To Refresh the Recordset")
       Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
    End If
End Sub
Private Sub ClearFields(ByVal strType As String)
    If strType = "P" Then
        Text2.Text = ""
        Text3.Text = ""
        Text4.Text = ""
        Text8.Text = ""
        MhDateInput1.Text = Format(Date, "dd-MM-yyyy")
        MhDateInput2.Text = "  -  -    "
        MhDateInput3.Text = Format(DateAdd("d", 10, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
        MhDateInput31.Text = "  -  -    "    'Extend Date
        MhRealInput5.Text = 0#
        MhRealInput6.Text = 0#
        MhRealInput11.Text = "5.00"
        MhRealInput12.Text = "0.00"
        MhRealInput14.Text = "0.00"
        MhRealInput15.Text = "0.00"
        MhRealInput16.Text = "0.00"
        MhRealInput20.Text = "0.00"
        
        TxtAdNar.Text = ""
        fpSpread1.ClearRange 1, 1, fpSpread1.MaxCols, fpSpread1.MaxRows, True: fpSpread1.SetActiveCell 1, 1
        fpSpread2.ClearRange 1, 1, fpSpread2.MaxCols, fpSpread2.MaxRows, True: fpSpread2.SetActiveCell 1, 1
    ElseIf strType = "C" Then
        Text5.Text = ""
        MhRealInput1.Text = "0"
        MhRealInput3.Text = "0.0000"
        MhRealInput4.Text = "0.00"
    End If
End Sub
Private Sub LoadFields()
       
    If rstOutsourceItemPOParent.EOF Or rstOutsourceItemPOParent.BOF Then Exit Sub
    Text2.Text = rstOutsourceItemPOParent.Fields("Name").Value
    MhDateInput1.Text = Format(rstOutsourceItemPOParent.Fields("Date").Value, "dd-MM-yyyy")
    SupplierCode = rstOutsourceItemPOParent.Fields("Supplier").Value
    If rstSupplierList.RecordCount > 0 Then rstSupplierList.MoveFirst
    rstSupplierList.Find "[Code] = '" & SupplierCode & "'"
    If Not rstSupplierList.EOF Then
       Text3.Text = rstSupplierList.Fields("Col0").Value
    End If
    
    MhDateInput3.Text = Format(rstOutsourceItemPOParent.Fields("DeliveryDate").Value, "dd-MM-yyyy")
    
    If Not IsNull(rstOutsourceItemPOParent.Fields("ExtendDate").Value) Then
       MhDateInput31.Text = Format(rstOutsourceItemPOParent.Fields("ExtendDate").Value, "dd-MM-yyyy")
    End If
    
    Text4.Text = rstOutsourceItemPOParent.Fields("Remarks").Value
    MhRealInput11.Text = Format(Val(rstOutsourceItemPOParent.Fields("VAT%").Value), "0.00")
    MhRealInput12.Text = Format(Val(rstOutsourceItemPOParent.Fields("VAT").Value), "0.00")
    MhRealInput14.Text = Format(Val(rstOutsourceItemPOParent.Fields("Adjustment").Value), "0.00")
    MhRealInput15.Text = Format(Val(rstOutsourceItemPOParent.Fields("BillAmount").Value), "0.00")
    Text8.Text = rstOutsourceItemPOParent.Fields("BillNo").Value
    
    If Not IsNull(rstOutsourceItemPOParent.Fields("BillDate").Value) Then
         MhDateInput2.Text = Format(rstOutsourceItemPOParent.Fields("BillDate").Value, "dd-MM-yyyy")
    End If
    
    MhRealInput16.Text = Format(Val(rstOutsourceItemPOParent.Fields("PaidAmount").Value), "0.00")
    TxtAdNar.Text = rstOutsourceItemPOParent.Fields("AdjustmentRemarks").Value
    Call LoadOutsourceItemList(rstOutsourceItemPOParent.Fields("Code").Value)
    If rstOutsourceItemPOChild.State = adStateOpen Then CalculateTotal ("G")
    
End Sub
Private Sub EditRecord()
    On Error GoTo ErrorHandler
    If rstOutsourceItemPOParent.RecordCount = 0 Then Exit Sub
    If rstOutsourceItemPOChild.State = adStateClosed Then
        SSTab1.Tab = 0
        Exit Sub
    End If
    If rstOutsourceItemPOParent.State = adStateOpen Then
       rstOutsourceItemPOParent.Close
    End If
       
    rstOutsourceItemPOParent.CursorLocation = adUseServer
    rstOutsourceItemPOParent.Open "Select * From OutsourceItemPOParent Where Code = '" & FixQuote(rstOutsourceItemPOList.Fields("Code").Value) & "'", CxnOutsourceItemPurchaseOrder, adOpenKeyset, adLockPessimistic
    rstOutsourceItemPOParent.Fields("Printstatus") = "N"
    
    If rstMaterialIOParent.State = adStateOpen Then
       rstMaterialIOParent.Close
    End If
    
    rstMaterialIOParent.CursorLocation = adUseServer
    rstMaterialIOParent.Open "Select * From MaterialIOParent Where Code = '" & FixQuote(rstOutsourceItemPOList.Fields("Code").Value) & "'", CxnOutsourceItemPurchaseOrder, adOpenKeyset, adLockPessimistic
    
    MdiMainMenu.MousePointer = vbHourglass
    MdiMainMenu.MousePointer = vbNormal
    
    
    AddToList
    Call SetButtons(False)
    SSTab1.TabEnabled(0) = False
    Text2.SetFocus
    blnRecordExist = True
    If AllowTransactionsModification = 0 Then
        If Not CheckEmpty(Text8.Text, False) Then LockFields (True)
        Text1.Locked = False
    End If
    CxnOutsourceItemPurchaseOrder.BeginTrans
    Exit Sub
ErrorHandler:
    If Err.Number = -2147467259 Then
       Call DisplayError("Failed to Edit the record")
    End If
    MdiMainMenu.MousePointer = vbNormal
    SSTab1.Tab = 0
End Sub

Private Sub SaveFields()
    If rstOutsourceItemPOParent.EOF Or rstOutsourceItemPOParent.BOF Then Exit Sub
    Dim lpBuff As String * 1024
    GetComputerName lpBuff, Len(lpBuff)
    If Not blnRecordExist Then
        rstOutsourceItemPOParent.Fields("Code").Value = GenerateCode(CxnOutsourceItemPurchaseOrder, "Select Max(Code) From OutsourceItemPOParent", 6, "0")
        rstOutsourceItemPOParent.Fields("CreatedBy").Value = UserCode
        rstOutsourceItemPOParent.Fields("CreatedOn").Value = Now()
        rstOutsourceItemPOParent.Fields("Recordstatus").Value = "N"
    Else
        rstOutsourceItemPOParent.Fields("ModifiedBy").Value = UserCode
        rstOutsourceItemPOParent.Fields("ModifiedOn").Value = Now()
        rstOutsourceItemPOParent.Fields("Recordstatus").Value = "M"
    End If
    rstOutsourceItemPOParent.Fields("Name").Value = Pad(Trim(Text2.Text), Space(1), 10, "L")
    rstOutsourceItemPOParent.Fields("Date").Value = GetDate(MhDateInput1.Text)
    rstOutsourceItemPOParent.Fields("Supplier").Value = SupplierCode
    rstOutsourceItemPOParent.Fields("DeliveryDate").Value = GetDate(MhDateInput3.Text)
    
    If Not IsDate(MhDateInput31.Text) Then rstOutsourceItemPOParent.Fields("ExtendDate").Value = Null Else rstOutsourceItemPOParent.Fields("ExtendDate").Value = GetDate(MhDateInput31.Text)
    
    rstOutsourceItemPOParent.Fields("Remarks").Value = Trim(Text4.Text)
    rstOutsourceItemPOParent.Fields("VAT%").Value = Format(Val(MhRealInput11.Text), "0.00")
    rstOutsourceItemPOParent.Fields("VAT").Value = Format(Val(MhRealInput12.Text), "0.00")
    rstOutsourceItemPOParent.Fields("Adjustment").Value = Format(Val(MhRealInput14.Text), "0.00")
    rstOutsourceItemPOParent.Fields("BillAmount").Value = Format(Val(MhRealInput15.Text), "0.00")
    rstOutsourceItemPOParent.Fields("BillNo").Value = Trim(Text8.Text)
    
    If Not IsDate(MhDateInput2.Text) Then
         rstOutsourceItemPOParent.Fields("BillDate").Value = Null
    Else
         rstOutsourceItemPOParent.Fields("BillDate").Value = GetDate(MhDateInput2.Text)
    End If
    
    rstOutsourceItemPOParent.Fields("PaidAmount").Value = Format(Val(MhRealInput16.Text), "0.00")
    rstOutsourceItemPOParent.Fields("AdjustmentRemarks").Value = IIf(Val(MhRealInput14.Text) <> 0, TxtAdNar.Text, "")
    If Not CheckEmpty(Text8.Text, False) Then If IsNull(rstOutsourceItemPOParent.Fields("BillFeedDate").Value) Then rstOutsourceItemPOParent.Fields("BillFeedDate").Value = Now()
    If Not CheckEmpty(Text8.Text, False) Then If IsNull(rstOutsourceItemPOParent.Fields("ComputerName").Value) Then rstOutsourceItemPOParent.Fields("ComputerName").Value = Left(lpBuff, (InStr(1, lpBuff, vbNullChar)) - 1)
    rstOutsourceItemPOParent.Fields("PrintStatus").Value = "N"
    
    ''**********Entry in MaterialIOParent*************
    rstMaterialIOParent.Find "[Code] = '" & rstOutsourceItemPOParent.Fields("Code").Value & "'"
    If rstMaterialIOParent.EOF Then
        rstMaterialIOParent.AddNew
        rstMaterialIOParent.Fields("Code").Value = rstOutsourceItemPOParent.Fields("Code").Value ' GenerateCode(CxnOutsourceItemPurchaseOrder, "Select Max(Code) From OutsourceItemPOParent", 6, "0")
        rstMaterialIOParent.Fields("CreatedBy").Value = UserCode
        rstMaterialIOParent.Fields("CreatedOn").Value = Now()
        rstMaterialIOParent.Fields("Recordstatus").Value = "N"
    Else
        rstMaterialIOParent.Fields("ModifiedBy").Value = UserCode
        rstMaterialIOParent.Fields("ModifiedOn").Value = Now()
        rstMaterialIOParent.Fields("Recordstatus").Value = "M"
    End If
    rstMaterialIOParent.Fields("Name").Value = Pad(Trim(Text2.Text), Space(1), 10, "L")
    rstMaterialIOParent.Fields("Date").Value = GetDate(MhDateInput1.Text)
    rstMaterialIOParent.Fields("Source").Value = SupplierCode
    rstMaterialIOParent.Fields("Type").Value = "1"
    rstMaterialIOParent.Fields("Remarks").Value = Trim(Text4.Text)
    rstMaterialIOParent.Fields("PrintStatus").Value = "N"
   
End Sub

Private Sub AddToList()
    On Error Resume Next
    rstOutsourceItemPOList.MoveFirst
    rstOutsourceItemPOList.Find "[Code] = '" & rstOutsourceItemPOParent.Fields("Code").Value & "'"
    If rstOutsourceItemPOList.EOF Then
       rstOutsourceItemPOList.AddNew
       rstOutsourceItemPOList.Fields("Code").Value = rstOutsourceItemPOParent.Fields("Code").Value
    End If
    rstOutsourceItemPOList.Fields("Name").Value = Pad(rstOutsourceItemPOParent.Fields("Name").Value, Space(1), 10, "L")
    rstOutsourceItemPOList.Fields("Date").Value = rstOutsourceItemPOParent.Fields("Date").Value
    rstSupplierList.MoveFirst
    rstSupplierList.Find "[Code] = '" & rstOutsourceItemPOParent.Fields("Supplier").Value & "'"
    rstOutsourceItemPOList.Fields("SupplierName").Value = Trim(rstSupplierList.Fields("Col0").Value)
    rstOutsourceItemPOList.Fields("BillAmount").Value = rstOutsourceItemPOParent.Fields("BillAmount").Value
    rstOutsourceItemPOList.Update
    rstOutsourceItemPOList.Sort = SortOrder & " Asc"
    rstOutsourceItemPOList.Find "[Code] = '" & rstOutsourceItemPOParent.Fields("Code").Value & "'"
End Sub
Private Function CheckMandatoryFields() As Boolean
    If CheckEmpty(Text2.Text, False) Then
       DisplayError ("Order No. cannot be blank")
       Text2.SetFocus
       CheckMandatoryFields = True
    ElseIf CheckEmpty(Text3.Text, False) Then
       Text3.SetFocus
       CheckMandatoryFields = True
    ElseIf Not CheckExists(Text3, "Col0", rstSupplierList, SupplierCode) Then
        Text3.SetFocus
        CheckMandatoryFields = True
    ElseIf CheckDuplicate(CxnOutsourceItemPurchaseOrder, "OutsourceItemPOParent", "Code", "[Name]", Trim(Text2.Text), rstOutsourceItemPOParent.Fields("Code").Value, False) Then
        Text2.SetFocus
        CheckMandatoryFields = True
    End If
    
    If Val(MhRealInput14.Text) <> 0 Then
       If CheckEmpty(TxtAdNar.Text, False) Then
         TxtAdNar.SetFocus
         CheckMandatoryFields = True
         Exit Function
       End If
     End If
    ''***Comment By Shamshad******
    'If Val(MhRealInput16.Text) <> 0 Then If Val(MhRealInput16.Text) <> Val(MhRealInput15.Text) Then MhRealInput14.SetFocus: CheckMandatoryFields = True: Exit Function: Exit Function
    If Val(MhRealInput14.Text) <> 0 Then If CheckEmpty(TxtAdNar.Text, False) Then TxtAdNar.SetFocus: CheckMandatoryFields = True
End Function
Private Sub Timer1_Timer()
    On Error Resume Next
    
    MdiMainMenu.ProgressBar1.Value = MdiMainMenu.ProgressBar1.Value + 10
    If MdiMainMenu.ProgressBar1.Value = 100 Then
       Timer1.Enabled = False
       ShowProgressInStatusBar False
    End If
End Sub
Private Sub LoadOutsourceItemListOld(ByVal strOrderCode As String)
    On Error GoTo ErrorHandler
    
    If rstOutsourceItemPOChild.State = adStateOpen Then
       rstOutsourceItemPOChild.Close
    End If
    rstOutsourceItemPOChild.Open "Select C.OutsourceItem,M.Name As OutsourceItemName,C.Quantity,C.Rate,C.Amount From OutsourceItemMaster M,OutsourceItemPOChild C Where C.OutsourceItem = M.Code And C.Code = '" & strOrderCode & "' Order By M.Name", CxnOutsourceItemPurchaseOrder, adOpenKeyset, adLockOptimistic
    rstOutsourceItemPOChild.ActiveConnection = Nothing
    'Set DataGrid2.DataSource = rstOutsourceItemPOChild
    Exit Sub
ErrorHandler:
    DisplayError ("Failed to Load OutsourceItem List")
End Sub

Private Sub LoadOutsourceItemList(ByVal strOrderCode As String)
    Dim i As Integer
    On Error GoTo ErrorHandler
    If rstOutsourceItemPOChild.State = adStateOpen Then
       rstOutsourceItemPOChild.Close
    End If
    rstOutsourceItemPOChild.Open "Select C.OutsourceItem,M.Name As OutsourceItemName,C.Quantity,C.Rate,C.Amount From OutsourceItemMaster M,OutsourceItemPOChild C Where C.OutsourceItem = M.Code And C.Code = '" & strOrderCode & "' Order By M.Name", CxnOutsourceItemPurchaseOrder, adOpenKeyset, adLockOptimistic
    rstOutsourceItemPOChild.ActiveConnection = Nothing
    If rstOutsourceItemPOChild.RecordCount > 0 Then rstOutsourceItemPOChild.MoveFirst
    i = 0
    Do While Not rstOutsourceItemPOChild.EOF
        i = i + 1
        With fpSpread1
            .SetText 1, i, rstOutsourceItemPOChild.Fields("OutsourceItemName").Value
            .SetText 2, i, Val(rstOutsourceItemPOChild.Fields("Quantity").Value)
            .SetText 3, i, Val(rstOutsourceItemPOChild.Fields("Rate").Value)
            .SetText 4, i, Val(rstOutsourceItemPOChild.Fields("Amount").Value)
            .SetText 5, i, rstOutsourceItemPOChild.Fields("OutsourceItem").Value
            
        End With
        rstOutsourceItemPOChild.MoveNext
    Loop
    If rstMaterialIOChild.State = adStateOpen Then rstMaterialIOChild.Close
    rstMaterialIOChild.Open "Select C.Item,M1.Name As OutsourceItemName,C.Godown,M2.Name As GodownName,C.Quantity From OutsourceItemMaster M1,MaterialIOChild C,AccountMaster M2, OutsourceItemPOParent T Where C.Item = M1.Code And C.Godown = M2.Code And C.Code = T.Code And C.Code = '" & strOrderCode & "' Order By M1.Name", CxnOutsourceItemPurchaseOrder, adOpenKeyset, adLockOptimistic
    rstMaterialIOChild.ActiveConnection = Nothing
    If rstMaterialIOChild.RecordCount > 0 Then rstMaterialIOChild.MoveFirst
    i = 0
    Do While Not rstMaterialIOChild.EOF
        i = i + 1
        With fpSpread2
            .SetText 1, i, rstMaterialIOChild.Fields("OutsourceItemName").Value
            .SetText 2, i, rstMaterialIOChild.Fields("GodownName").Value
            .SetText 3, i, Val(rstMaterialIOChild.Fields("Quantity").Value)
            .SetText 4, i, rstMaterialIOChild.Fields("Item").Value
            .SetText 5, i, rstMaterialIOChild.Fields("Godown").Value
            
        End With
        rstMaterialIOChild.MoveNext
    Loop
    Exit Sub
ErrorHandler:
    DisplayError ("Failed to Load OutsourceItem List")
End Sub

Private Sub Text5_Validate(Cancel As Boolean)
    Dim SearchString As String
    Dim LastPurchaseRate As Double
    
    SearchString = FixQuote(Text5.Text)
    If rstOutsourceItemList.RecordCount = 0 Then
        DisplayError ("No Record in OutsourceItem Master")
        Cancel = True
        Exit Sub
    Else
        rstOutsourceItemList.MoveFirst
    End If
    rstOutsourceItemList.Find "[Col0] = '" & RTrim(SearchString) & "'"
    If rstOutsourceItemList.EOF Then
        SelectionType = "S"
        OutsourceItemCode = ""
        Call LoadSelectionList(rstOutsourceItemList, "List of Outsource Items...", "Name")
        SearchOrder = 0
        Call DisplaySelectionList(Text5, OutsourceItemCode)
        Call CloseForm(FrmSelectionList)
        If CheckEmpty(Text5.Text, False) Then
            Text5.Text = "?"
        End If
        If RTrim(OutsourceItemCode) <> "" Then
            SendKeys "{TAB}"
        End If
        Cancel = True
        Exit Sub
    ElseIf (rstOutsourceItemPOChild.Fields("OutsourceItemName").Value <> Text5.Text) Or (CheckEmpty(rstOutsourceItemPOChild.Fields("OutsourceItemName").Value, False)) Then
        If CheckDuplicateOutsourceItem Then
            Call DisplayError("Duplicate Entry")
            Text5.SelStart = 0
            Text5.SelLength = Len(Text5.Text)
            Cancel = True
            Exit Sub
        End If
    End If
    OutsourceItemCode = rstOutsourceItemList.Fields("Code").Value
    LastPurchaseRate = GetLastPurchaseRate
    If LastPurchaseRate > 0 Then
        MsgBox "Last Purchase Rate : Rs." & Format(LastPurchaseRate, "###.00") & " !!!", vbInformation, App.Title
    End If
End Sub

Private Sub fpSpread1_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask And KeyCode = vbKeyD Then
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            fpSpread1.DeleteRows fpSpread1.ActiveRow, 1: fpSpread1.SetFocus
            CalculateTotal ("G"): CalculateTotal ("N")
        End If
    ElseIf KeyCode = vbKeySpace Then
         Dim SearchString As Variant, LastPurchaseRate As Double
         SearchString = FixQuote(Text5.Text)
         With fpSpread1
            If .ActiveCol = 1 Then
                .GetText .ActiveCol, .ActiveRow, SearchString
                Text5.Text = FixQuote(SearchString)
                If rstOutsourceItemList.RecordCount = 0 Then DisplayError ("No Record in OutsourceItem Master"): .SetActiveCell 1, .ActiveRow: Exit Sub Else rstOutsourceItemList.MoveFirst
                rstOutsourceItemList.Find "[Col0] = '" & RTrim(SearchString) & "'"
                SelectionType = "S"
                OutsourceItemCode = ""
                Call LoadSelectionList(rstOutsourceItemList, "List of Outsource Items...", "Name")
                SearchOrder = 0
                Call DisplaySelectionList(Text5, OutsourceItemCode)
                Call CloseForm(FrmSelectionList)
                If OutsourceItemCode = "" Then
                    .SetActiveCell 1, .ActiveRow
                Else
                    rstOutsourceItemList.MoveFirst: rstOutsourceItemList.Find "[Code] ='" & OutsourceItemCode & "'"
                    .SetText 1, .ActiveRow, Text5.Text
                    .SetText 5, .ActiveRow, OutsourceItemCode
                    LastPurchaseRate = GetLastPurchaseRate
                    If LastPurchaseRate > 0 Then MsgBox "Last Purchase Rate : Rs." & Format(LastPurchaseRate, "###0.00") & " !!!", vbInformation, App.Title
                    .SetFocus
                    SendKeys "{ENTER}"
                End If
            End If
        End With
    End If
End Sub

Private Sub fpSpread2_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If Shift = vbCtrlMask And KeyCode = vbKeyD Then
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            fpSpread2.DeleteRows fpSpread2.ActiveRow, 1: fpSpread2.SetFocus
            CalculateTotal ("G")
        End If
    ElseIf KeyCode = vbKeySpace Then
        Dim Paper As Variant, Account As Variant
        With fpSpread2
            .GetText 1, .ActiveRow, Paper
            If .ActiveCol = 1 Then
                If Paper = "" Then
                    fpSpread1.GetText 1, fpSpread1.ActiveRow, Paper
                    .SetText 1, .ActiveRow, Paper
                     fpSpread1.GetText 5, fpSpread1.ActiveRow, Paper
                    .SetText 4, .ActiveRow, Paper
                    If Paper <> "" Then SendKeys "{ENTER}"
                End If
                
            ElseIf .ActiveCol = 2 Then
                If Paper <> "" Then
                    .GetText 2, .ActiveRow, Account
                    Text5.Text = FixQuote(Account)
                    If rstAccountList.RecordCount = 0 Then DisplayError ("No Record in Account Master"): .SetActiveCell 1, .ActiveRow: Exit Sub Else rstAccountList.MoveFirst
                    rstAccountList.Find "[Col0] = '" & RTrim(Account) & "'"
                    SelectionType = "S"
                    AccountCode = ""
                    Call LoadSelectionList(rstAccountList, "List of Accounts...", "Name")
                    SearchOrder = 0
                    Call DisplaySelectionList(Text5, AccountCode)
                    Call CloseForm(FrmSelectionList)
                    If AccountCode = "" Then
                        .SetActiveCell 2, .ActiveRow
                    Else
                        rstAccountList.MoveFirst: rstAccountList.Find "[Code] ='" & AccountCode & "'"
                        .SetText 2, .ActiveRow, Text5.Text
                         Dim accCode As Variant
                         accCode = ""
                         accCode = AccountCode
                        .SetText 5, .ActiveRow, accCode
                        SendKeys "{ENTER}"
                    End If
                End If
            ElseIf .ActiveCol = 3 Then
            End If
        End With
    End If
    End Sub
Private Sub fpSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim Qty As Variant, Rate As Variant, Paper As Variant, tAmt As Double
    With fpSpread1
        If Col = 1 Or Col = 2 Or Col = 3 Then
            .GetText 1, Row, Paper
            .GetText 2, Row, Qty
            .GetText 3, Row, Rate
            tAmt = Qty * Rate
            If Paper = "" Then .SetText 2, Row, "": .SetText 4, Row, "" Else: .SetText 2, Row, Qty: .SetText 4, Row, tAmt:  CalculateTotal ("G"): CalculateTotal ("N")
            '.SetText 4, Row, tAmt: CalculateTotal ("G"): CalculateTotal ("N")
        End If
    End With
End Sub
Private Sub fpSpread2_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim Paper As Variant, Qty As Variant
    With fpSpread2
        If Col = 3 Then
          .GetText 1, Row, Paper
          .GetText 3, Row, Qty
          If Paper = "" Then .SetText 3, Row, "" Else .SetText 3, Row, Qty
          CalculateTotal ("G")
        End If
    End With
End Sub
Private Sub fpSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    EditMode = IIf(Mode = 1, True, False)
End Sub
Private Sub fpSpread2_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    EditMode = IIf(Mode = 1, True, False)
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
Private Sub MhRealInput1_Validate(Cancel As Boolean)
    Call CalculateAmount
End Sub
Private Sub MhRealInput1_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyEscape Then
        MakeTextBoxInvisible (True)
    End If
End Sub
Private Sub MhRealInput3_Validate(Cancel As Boolean)
    Call CalculateAmount
End Sub
Private Sub MhRealInput3_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyEscape Then
        MakeTextBoxInvisible (True)
    End If
End Sub
Private Sub MhRealInput4_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyReturn Then
        If Val(MhRealInput4.Text) > 0 Then
            rstOutsourceItemPOChild.Fields("OutsourceItem").Value = OutsourceItemCode
            rstOutsourceItemPOChild.Fields("OutsourceItemName").Value = Trim(Text5.Text)
            rstOutsourceItemPOChild.Fields("Quantity").Value = Format(Val(MhRealInput1.Text), "0")
            rstOutsourceItemPOChild.Fields("Rate").Value = Format(Val(MhRealInput3.Text), "0.0000")
            rstOutsourceItemPOChild.Fields("Amount").Value = Format(Val(MhRealInput4.Text), "0.00")
            rstOutsourceItemPOChild.Update
            MakeTextBoxInvisible (False)
            CalculateTotal ("G"): CalculateTotal ("N")
            If rstOutsourceItemPOChild.AbsolutePosition = rstOutsourceItemPOChild.RecordCount Then
                'Call DataGrid2_KeyDown(vbKeyA, vbCtrlMask)
            End If
        End If
    ElseIf Shift = 0 And KeyCode = vbKeyEscape Then
       MakeTextBoxInvisible (True)
    End If
End Sub
Private Sub MhRealInput4_Validate(Cancel As Boolean)
    Cancel = True
End Sub
Private Sub MakeTextBoxInvisible(ByVal KeyEscPressed As Boolean)
    If KeyEscPressed Then
        If Not (rstOutsourceItemPOChild.EOF Or rstOutsourceItemPOChild.BOF) Then
            If Val(CheckNull(rstOutsourceItemPOChild.Fields("Quantity").Value)) = 0 Then
                rstOutsourceItemPOChild.Delete
                rstOutsourceItemPOChild.MoveNext
                If rstOutsourceItemPOChild.RecordCount > 0 Then rstOutsourceItemPOChild.MoveFirst
            End If
        End If
    End If
    Text5.Visible = False
    MhRealInput1.Visible = False
    MhRealInput3.Visible = False
    MhRealInput4.Visible = False
    'DataGrid2.Enabled = True
    'DataGrid2.SetFocus
End Sub
Private Sub CalculateAmount()
    MhRealInput4.Text = Format(Val(MhRealInput1.Text) * Val(MhRealInput3.Text), "0.00")
End Sub
Private Sub CalculateTotalOld(ByVal strType As String)
    Dim dblBookMark As Double
    If strType = "G" Then
        MhRealInput5.Text = 0
        MhRealInput6.Text = 0
        If rstOutsourceItemPOChild.RecordCount <> 0 Then
            If Not (rstOutsourceItemPOChild.EOF Or rstOutsourceItemPOChild.BOF) Then
                dblBookMark = rstOutsourceItemPOChild.Bookmark
            End If
            rstOutsourceItemPOChild.MoveFirst
            Do While Not rstOutsourceItemPOChild.EOF
                MhRealInput5.Text = Val(MhRealInput5.Text) + Val(rstOutsourceItemPOChild.Fields("Quantity").Value)
                MhRealInput6.Text = Val(MhRealInput6.Text) + Val(rstOutsourceItemPOChild.Fields("Amount").Value)
                rstOutsourceItemPOChild.MoveNext
            Loop
            If dblBookMark <> 0 Then
                rstOutsourceItemPOChild.Bookmark = dblBookMark
           Else
                rstOutsourceItemPOChild.MoveLast
           End If
        End If
        
        MhRealInput12.Text = Format(Val(MhRealInput6.Text) * Val(MhRealInput11.Text) / 100, "0.00")
    Else
        MhRealInput15.Text = Format(Val(MhRealInput6.Text) + Val(MhRealInput12.Text) + Val(MhRealInput14.Text), "0.00")
    End If
End Sub
Private Sub CalculateTotal(ByVal strType As String)
    Dim dblBookMark As Double
    Dim Qty01 As Variant, Amt As Variant, i As Integer, Qty As Long
    If strType = "G" Then
        MhRealInput5.Text = 0
        MhRealInput6.Text = 0
        MhRealInput20.Value = 0
        Qty = 0
        With fpSpread1
            For i = 1 To .DataRowCnt
                .GetText 2, i, Qty01: .GetText 4, i, Amt
                 Qty = Qty + Int(Val(Qty01))
                MhRealInput5.Text = Val(MhRealInput5.Text) + Qty01
                MhRealInput6.Text = Val(MhRealInput6.Text) + Amt
            Next
        End With
        Qty = 0
        With fpSpread2
            For i = 1 To .DataRowCnt
                .GetText 3, i, Qty01:
                Qty = Qty + Int(Val(Qty01))
                MhRealInput20.Value = Val(MhRealInput20.Text) + Qty01
            Next
        End With
        MhRealInput12.Text = Format(Val(MhRealInput6.Text) * Val(MhRealInput11.Text) / 100, "0.00")
    Else
        MhRealInput15.Text = Format(Val(MhRealInput6.Text) + Val(MhRealInput12.Text) + Val(MhRealInput14.Text), "0.00")
    End If
End Sub
Private Function CheckDuplicateOutsourceItem() As Boolean
    Dim dblBookMark As Double
    If rstOutsourceItemPOChild.RecordCount = 0 Then Exit Function
    If Not (rstOutsourceItemPOChild.EOF Or rstOutsourceItemPOChild.BOF) Then
       dblBookMark = rstOutsourceItemPOChild.Bookmark
    End If
    rstOutsourceItemPOChild.MoveFirst
    Do While Not rstOutsourceItemPOChild.EOF
          If rstOutsourceItemPOChild.Fields("OutsourceItemName").Value = Trim(Text5.Text) Then
             CheckDuplicateOutsourceItem = True
             Exit Do
          End If
          rstOutsourceItemPOChild.MoveNext
    Loop
    If dblBookMark <> 0 Then
       rstOutsourceItemPOChild.Bookmark = dblBookMark
    Else
       rstOutsourceItemPOChild.MoveLast
    End If
End Function
Private Function GetLastPurchaseRate() As Double
    On Error GoTo ErrorHandler
    
    If rstLastPurchaseRate.State = adStateOpen Then
       rstLastPurchaseRate.Close
    End If
    rstLastPurchaseRate.Open "Select Top 1 [Rate] From OutsourceItemPOParent P, OutsourceItemPOChild C Where P.Code = C.Code And C.OutsourceItem = '" & OutsourceItemCode & "' And P.Code < '" & IIf(IsNull(rstOutsourceItemPOParent.Fields("Code").Value), "999999", rstOutsourceItemPOParent.Fields("Code").Value) & "' Order By P.Name Desc", CxnOutsourceItemPurchaseOrder, adOpenKeyset, adLockReadOnly
    If rstLastPurchaseRate.RecordCount > 0 Then
        GetLastPurchaseRate = Val(rstLastPurchaseRate.Fields("Rate").Value)
    End If
    Exit Function
ErrorHandler:
    DisplayError ("Failed to get Last Purchase Rate")
End Function
Private Function UpdateOutsourceItemList(ByVal strOption As String) As Boolean
    Dim CellVal(1 To 5) As Variant, Sheets As Long
On Error GoTo ErrorHandler
    UpdateOutsourceItemList = True
    If strOption = "D" And (Not blnRecordExist) Then Exit Function
    If strOption = "D" Then
        CxnOutsourceItemPurchaseOrder.Execute "DELETE FROM OutsourceItemPOChild WHERE Code='" & rstOutsourceItemPOParent.Fields("Code").Value & "'"
        CxnOutsourceItemPurchaseOrder.Execute "DELETE FROM MaterialIOChild WHERE Code='" & rstOutsourceItemPOParent.Fields("Code").Value & "'"
    ElseIf strOption = "I1" Then
        With fpSpread1
            .GetText 2, .ActiveRow, CellVal(1)  'Quantity
            .GetText 3, .ActiveRow, CellVal(2)  'Rate
            .GetText 4, .ActiveRow, CellVal(3)  'Amount
            .GetText 5, .ActiveRow, CellVal(4)  'Item Code
        End With
        CxnOutsourceItemPurchaseOrder.Execute "INSERT INTO OutsourceItemPOChild VALUES ('" & rstOutsourceItemPOParent.Fields("Code").Value & "','" & CellVal(4) & "'," & Val(CellVal(1)) & "," & Val(CellVal(2)) & "," & Val(CellVal(3)) & ")"
    Else
        With fpSpread2
            .GetText 3, .ActiveRow, CellVal(1)  'Quantity
            .GetText 4, .ActiveRow, CellVal(2)  'Item Code
            .GetText 5, .ActiveRow, CellVal(3)  'Account Code
        End With
        'strRef = Pad(RTrim(Val(CellVal(1))), "0", 6, "L")
        CxnOutsourceItemPurchaseOrder.Execute "INSERT INTO MaterialIOChild VALUES ('" & rstOutsourceItemPOParent.Fields("Code").Value & "','1','" & CellVal(2) & "','" & CellVal(3) & "','0'," & Val(CellVal(1)) & ")"
                                                                                        'code ,                                             cateogry, Item,              godown,          ref, Quantity
    End If
    Exit Function
ErrorHandler:
    UpdateOutsourceItemList = False
End Function

Public Sub FilterRecord(ByVal SrchFor As String, ByVal SrchText As String)
    If SrchFor = "Supplier" Then
        rstOutsourceItemPOList.Filter = "[SupplierName] Like '%" & SrchText & "%'"
    End If
End Sub
Private Sub PrintOutsourceItemPurchaseOrder(ByVal VchNo As String, ByVal OutputTo As String)
    On Error Resume Next
    Dim oOutlookMsg As Outlook.MailItem
    Dim rstOutsourceItemPOChild2 As New ADODB.Recordset
    Dim rstMaterialIOChild2 As New ADODB.Recordset
    Screen.MousePointer = vbHourglass
    rptOutsourceItemPurchaseOrder.Text2.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptOutsourceItemPurchaseOrder.Text3.SetText Trim(rstCompanyMaster.Fields("Address1").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address2").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address3").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address4").Value)
    rptOutsourceItemPurchaseOrder.Text24.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & Space(1) & "Fax : " & Trim(rstCompanyMaster.Fields("Fax").Value) & Space(1) & "e-Mail : " & Trim(rstCompanyMaster.Fields("EMail").Value)
    If rstOutsourceItemPOChild2.State = adStateOpen Then rstOutsourceItemPOChild2.Close
    rstOutsourceItemPOChild2.Open "Select 'MI/" & Right(Year(FinancialYearFrom), 2) + "-" + Right(Year(FinancialYearTo), 2) & "/'+Trim(P.Name) As OrderNo,[Date] As OrderDate,DeliveryDate,(Select Trim(PrintName) From AccountMaster Where Code = P.Supplier) As SupplierName,[VAT%],VAT,Adjustment,BillAmount,Remarks,Trim(PrintName) As OutsourceItemName,Quantity,Rate,Amount,BillNo,BillDate,(Select Trim(EMail) From AccountMaster Where Code = P.Supplier) As EMailID From (OutsourceItemPOParent P Inner Join OutsourceItemPOChild C ON (P.Code=C.Code AND P.Code='" & rstOutsourceItemPOList.Fields("Code").Value & "')) INNER JOIN OutsourceItemMaster M ON C.OutsourceItem=M.Code ORDER BY M.PrintName", CxnOutsourceItemPurchaseOrder, adOpenKeyset, adLockOptimistic
        
    'Dim rsss As String
    'rsss = "Select Trim(P.Name) As OrderNo,[Date] As OrderDate,(Select Trim(PrintName) From AccountMaster Where Code = P.Source) As Godown,Trim(PrintName) As OutsourceItemName,(Select Trim(PrintName) From AccountMaster Where Code = C.Godown) As GodownName,Quantity,Remarks,(Select Trim(Address1) From AccountMaster Where Code = C.Godown) As Address1,(Select Trim(Address2) From AccountMaster Where Code = C.Godown) As Address2,(Select Trim(Address3) From AccountMaster Where Code = C.Godown) As Address3,(Select Trim(Address4) From AccountMaster Where Code = C.Godown) As Address4 From (MaterialIOParent As P Inner Join MaterialIOChild As C On (P.Code = C.Code And P.Code = '" & rstOutsourceItemPOList.Fields("Code").Value & "')) Inner Join OutsourceItemMaster M On C.Item = M.Code Order By M.PrintName"
        
    rstMaterialIOChild2.Open "Select Trim(P.Name) As OrderNo,[Date] As OrderDate,(Select Trim(PrintName) From AccountMaster Where Code = P.Source) As Godown,Trim(PrintName) As OutsourceItemName,(Select Trim(PrintName) From AccountMaster Where Code = C.Godown) As GodownName,Quantity,Remarks,(Select Trim(Address1) From AccountMaster Where Code = C.Godown) As Address1,(Select Trim(Address2) From AccountMaster Where Code = C.Godown) As Address2,(Select Trim(Address3) From AccountMaster Where Code = C.Godown) As Address3,(Select Trim(Address4) From AccountMaster Where Code = C.Godown) As Address4 From (MaterialIOParent As P Inner Join MaterialIOChild As C On (P.Code = C.Code And P.Code = '" & rstOutsourceItemPOList.Fields("Code").Value & "')) Inner Join OutsourceItemMaster M On C.Item = M.Code Order By M.PrintName", CxnOutsourceItemPurchaseOrder, adOpenKeyset, adLockOptimistic
    
    '(SELECT IIF(Address1 is null,"",Address1)+IIF(Address1 & Address2 is null,"", " ")+IIF(Address2 is null,"",Address2)+ IIF(Address1 & Address2 & Address3 is null,"","")+IIF(Address3 is null," ",Address3) + IIF(Address1 & Address2 & Address3 & Address4 is null,""," ")+IIF(Address4 is null," ",Address4) AS Address FROM AccountMaster Where Code = '000007')
   
    rptOutsourceItemPurchaseOrder.Text20.SetText "Add : GST @" + Format(rstOutsourceItemPOChild2.Fields("VAT%").Value, "0.00") + "%"
    rptOutsourceItemPurchaseOrder.Text28.SetText " (" & Trim(NumberToWords(rstOutsourceItemPOChild2.Fields("BillAmount").Value, True)) & ")"
    rptOutsourceItemPurchaseOrder.Text27.SetText "for " & Trim(rstOutsourceItemPOChild2.Fields("SupplierName").Value)
    rptOutsourceItemPurchaseOrder.Text9.SetText "for " & Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptOutsourceItemPurchaseOrder.Text8.SetText Trim(COMPANY_CIN) 'Add here company cin no
    
    If rstMaterialIOChild2.RecordCount = 0 Then
        rptOutsourceItemPurchaseOrder.Section17.Suppress = True
    End If
 
    rptOutsourceItemPurchaseOrder.Database.SetDataSource rstOutsourceItemPOChild2, 3, 1
    rptOutsourceItemPurchaseOrder.Subreport1.OpenSubreport.Database.SetDataSource rstMaterialIOChild2, 3, 1
    
    Screen.MousePointer = vbNormal
    rstOutsourceItemPOChild2.ActiveConnection = Nothing: rstMaterialIOChild2.ActiveConnection = Nothing
    
    Dim sBill As String
    
    EMailID = rstOutsourceItemPOChild2.Fields("EMailID").Value
    Attachment = Trim(rstOutsourceItemPOChild2.Fields("OrderNo").Value)
    sBill = Trim(rstOutsourceItemPOChild2.Fields("OrderNo").Value)
    Attachment = Mid(Attachment, InStr(4, Attachment, "/") + 1)
   
    Message = "Dear Sir,<Br>Please find attached herewith PO #" & Trim(rstOutsourceItemPOChild2.Fields("OrderNo").Value) & " for doing the needful at your end. An early execution of the order will be highly appreciated.<Br>Kindly acknowledge the receipt of mail and confirm the date of execution of order.<Br><Br>" & Trim(rstCompanyMaster.Fields("PrintName").Value) & "<Br>Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & "<Br>E-Mail : <a HRef='mailto:" & Trim(rstCompanyMaster.Fields("EMail").Value) & "'>" & Trim(rstCompanyMaster.Fields("EMail").Value) & "</a>"
    If OutputTo = "S" Then
        FrmReportViewer.EMailID = EMailID
        FrmReportViewer.Subject = "Paper Purchase Order #" & Trim(rstOutsourceItemPOChild2.Fields("OrderNo").Value)
        FrmReportViewer.Attachment = Attachment
        FrmReportViewer.Message = Message
        Set FrmReportViewer.Report = rptOutsourceItemPurchaseOrder
        FrmReportViewer.Show vbModal
    ElseIf OutputTo = "P" Then
           rptOutsourceItemPurchaseOrder.PrintOut False
    Else
        rptOutsourceItemPurchaseOrder.ExportOptions.FormatType = crEFTPortableDocFormat    ' Set the Export Format As .Pdf
        rptOutsourceItemPurchaseOrder.ExportOptions.DestinationType = crEDTDiskFile
        rptOutsourceItemPurchaseOrder.ExportOptions.DiskFileName = App.Path & "\Report\" & Attachment & ".Pdf"
        rptOutsourceItemPurchaseOrder.Export False
        'rptOutsourceItemPurchaseOrder.MoveFirst
        Set oOutlookMsg = oOutlook.CreateItem(olMailItem)
        With oOutlookMsg
            .To = Trim(EMailID)
            .Subject = "Paper Purchase Order #" & sBill
            .HTMLBody = "<Font Face='Calibri' Size='3'>" & Message & "</a>" & "</Font>"
            .Attachments.Add (App.Path & "\Report\" & Attachment & ".Pdf")
            .Importance = olImportanceHigh
            .ReadReceiptRequested = True
            If CheckEmpty(.To, False) Then .Display Else .Send
        End With
        Set oOutlookMsg = Nothing
        End If
        Set rptOutsourceItemPurchaseOrder = Nothing
    Call CloseRecordset(rstOutsourceItemPOChild2): Call CloseRecordset(rstMaterialIOChild2)
    On Error GoTo 0
End Sub
Private Sub LockFields(ByVal bVal As Boolean)
    Dim O As Object
    For Each O In Me
        If TypeName(O) = "TextBox" Then
            O.Locked = bVal
        ElseIf TypeName(O) = "TDBNumber" Then
            O.ReadOnly = bVal
        ElseIf TypeName(O) = "TDBDate" Then
            O.ReadOnly = bVal
        End If
    Next
End Sub

Private Sub LoadRefList(ByVal strOutsourceItemCode As String, ByVal strSupplierCode As String, ByVal strOrderCode As String)
    
'    Dim BalanceQuantity As Long
'    On Error GoTo ErrorHandler
'
'    If rstRefList.State = adStateOpen Then
'        rstRefList.Close
'    End If
'
'    rstRefList.Open "Select P.Name,Format(Quantity,0) As ReceivedQuantity,Format((Select Sum(Quantity) From MaterialIOChild Where MaterialIOChild.Ref=P.Code And MaterialIOChild.Item=C.OutsourceItem And MaterialIOChild.Code<>'" & strOrderCode & "'),0) As IssuedQuantity,Quantity As BalanceQuantity,Remarks As Col0,P.Code From OutsourceItemPOParent P Inner Join OutsourceItemPOChild C On (P.Code=C.Code And P.Supplier='" & strSupplierCode & "' And C.OutsourceItem='" & strOutsourceItemCode & "') Order By P.Name", CxnOutsourceItemPurchaseOrder, adOpenKeyset, adLockOptimistic
'    rstRefList.ActiveConnection = Nothing
'    Do While Not rstRefList.EOF
'        BalanceQuantity = (Val(CheckNull(rstRefList.Fields("ReceivedQuantity").Value)) - Val(CheckNull(rstRefList.Fields("IssuedQuantity").Value))) - CalculateQuantityIssued(strOutsourceItemCode)
'        If BalanceQuantity <> 0 Then
'            rstRefList.Fields("Col0").Value = Trim(rstRefList.Fields("Name").Value) + " Quantity : " + Format(BalanceQuantity, "0")
'            rstRefList.Fields("BalanceQuantity").Value = BalanceQuantity
'
'            rstRefList.Update
'        Else
'            rstRefList.Delete
'        End If
'        rstRefList.MoveNext
'    Loop
'    Exit Sub
'ErrorHandler:
'    DisplayError ("Failed to Load Ref List")
End Sub

Private Function CalculateQuantityIssued(ByVal strOutsourceItemCode As String) As Long
'    Dim dblBookMark As Double
'
'    If rstMaterialIOChild.RecordCount = 0 Then Exit Function
'    If Not (rstMaterialIOChild.EOF Or rstMaterialIOChild.BOF) Then
'       dblBookMark = rstMaterialIOChild.Bookmark
'    End If
'    rstMaterialIOChild.MoveFirst
'    Do While Not rstMaterialIOChild.EOF
'        If rstMaterialIOChild.Bookmark <> dblBookMark Then
'            If Trim(rstMaterialIOChild.Fields("RefNo").Value) = Trim(rstRefList.Fields("Name").Value) And rstMaterialIOChild.Fields("Item").Value = strOutsourceItemCode Then
'                CalculateQuantityIssued = CalculateQuantityIssued + Val(rstMaterialIOChild.Fields("Quantity").Value)
'            End If
'        End If
'        rstMaterialIOChild.MoveNext
'    Loop
'    If dblBookMark <> 0 Then
'        rstMaterialIOChild.Bookmark = dblBookMark
'    Else
'        rstMaterialIOChild.MoveLast
'    End If
End Function

Private Function CheckDuplicateEntry(ByVal ItemName As String, ByVal RefNo As String, ByVal GodownName As String) As Boolean
'    Dim dblBookMark As Double
'
'    If rstMaterialIOChild.RecordCount = 0 Then Exit Function
'    If Not (rstMaterialIOChild.EOF Or rstMaterialIOChild.BOF) Then
'       dblBookMark = rstMaterialIOChild.Bookmark
'    End If
'    rstMaterialIOChild.MoveFirst
'    Do While Not rstMaterialIOChild.EOF
'          If rstMaterialIOChild.Fields("OutsourceItemName").Value = Trim(ItemName) And Trim(rstMaterialIOChild.Fields("RefNo").Value) = Trim(RefNo) And Trim(rstMaterialIOChild.Fields("GodownName").Value) = Trim(GodownName) Then
'             CheckDuplicateEntry = True
'             Exit Do
'          End If
'          rstMaterialIOChild.MoveNext
'    Loop
'    If dblBookMark <> 0 Then
'       rstMaterialIOChild.Bookmark = dblBookMark
'    Else
'       rstMaterialIOChild.MoveLast
'    End If
End Function
