VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmPaperDeliveryChallan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Paper Delivery Challan"
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11385
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "PaperDeliveryChallan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9120
   ScaleWidth      =   11385
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   9090
      Left            =   15
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   0
      Width           =   11190
      _Version        =   65536
      _ExtentX        =   19738
      _ExtentY        =   16034
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
      Picture         =   "PaperDeliveryChallan.frx":0442
      Begin TabDlg.SSTab SSTab1 
         Height          =   8865
         Left            =   120
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   120
         Width           =   10965
         _ExtentX        =   19341
         _ExtentY        =   15637
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
         TabPicture(0)   =   "PaperDeliveryChallan.frx":045E
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
         TabPicture(1)   =   "PaperDeliveryChallan.frx":047A
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Mh3dFrame2"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         Begin VB.CommandButton Command1 
            DisabledPicture =   "PaperDeliveryChallan.frx":0496
            Height          =   375
            Left            =   12280
            Picture         =   "PaperDeliveryChallan.frx":08A8
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   8390
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
            Left            =   605
            MaxLength       =   40
            TabIndex        =   19
            Top             =   8430
            Width           =   10260
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   7905
            Left            =   120
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   450
            Width           =   10740
            _ExtentX        =   18944
            _ExtentY        =   13944
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
               Caption         =   "   Challan No."
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
               Caption         =   "Challan Date"
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
               DataField       =   "ConsigneeName"
               Caption         =   "Consignee Name"
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
               Caption         =   "     Order Amount"
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
                  ColumnWidth     =   9074.835
               EndProperty
               BeginProperty Column03 
                  Alignment       =   1
                  Locked          =   -1  'True
               EndProperty
            EndProperty
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   8460
            Left            =   -74880
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   360
            Width           =   10740
            _Version        =   65536
            _ExtentX        =   18944
            _ExtentY        =   14922
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
            Picture         =   "PaperDeliveryChallan.frx":0CBA
            Begin VB.TextBox Text3 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   9360
               MaxLength       =   139
               TabIndex        =   2
               Top             =   45
               Width           =   1275
            End
            Begin VB.TextBox txtCarrierName 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   3120
               MaxLength       =   139
               TabIndex        =   6
               Top             =   2880
               Width           =   7515
            End
            Begin VB.TextBox txtCarrierAddress 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   3120
               MaxLength       =   139
               TabIndex        =   7
               Top             =   3195
               Width           =   7515
            End
            Begin VB.TextBox txtVehicleNo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   3120
               MaxLength       =   139
               TabIndex        =   8
               Top             =   3510
               Width           =   2785
            End
            Begin VB.TextBox txtDestinationGoods 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   8295
               MaxLength       =   139
               TabIndex        =   9
               Top             =   3510
               Width           =   2345
            End
            Begin VB.TextBox txtDestinationAddress 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   3120
               MaxLength       =   139
               TabIndex        =   10
               Top             =   3825
               Width           =   7515
            End
            Begin VB.TextBox txtEWayBillNo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   3120
               MaxLength       =   139
               TabIndex        =   11
               Top             =   4140
               Width           =   2805
            End
            Begin VB.TextBox txtTransitFormNumber 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   3120
               MaxLength       =   139
               TabIndex        =   13
               Top             =   4455
               Width           =   2805
            End
            Begin VB.TextBox txtRefNo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   8295
               MaxLength       =   139
               TabIndex        =   14
               Top             =   4455
               Width           =   2345
            End
            Begin VB.TextBox txtConsigneeName 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   3120
               MaxLength       =   139
               TabIndex        =   3
               Top             =   1680
               Width           =   7515
            End
            Begin VB.TextBox txtConsigneeAddress 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   3120
               MaxLength       =   139
               TabIndex        =   4
               Top             =   1995
               Width           =   7515
            End
            Begin VB.TextBox txtConsigneeRegNo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   3120
               MaxLength       =   139
               TabIndex        =   5
               Top             =   2310
               Width           =   2785
            End
            Begin VB.TextBox txtConsignerAddress 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   3120
               Locked          =   -1  'True
               MaxLength       =   139
               TabIndex        =   26
               Text            =   "4582-83/15, OPP.MTNL OFFICE,DARYA GANJ, NEW DELHI - 110002"
               Top             =   795
               Width           =   7515
            End
            Begin VB.TextBox txtConsignerName 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   3120
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   44
               Text            =   "RACHNA SAGAR (P) LTD."
               Top             =   480
               Width           =   7515
            End
            Begin VB.TextBox txtConsignerRegNo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               DataSource      =   "Adodc1"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   3120
               Locked          =   -1  'True
               MaxLength       =   139
               TabIndex        =   25
               Text            =   "07AAACR5864Q1ZO"
               Top             =   1110
               Width           =   7515
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
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1560
               MaxLength       =   10
               TabIndex        =   0
               Top             =   50
               Width           =   1530
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
               Height          =   330
               Left            =   120
               TabIndex        =   22
               Top             =   45
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
               Picture         =   "PaperDeliveryChallan.frx":0CD6
               Picture         =   "PaperDeliveryChallan.frx":0CF2
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Left            =   5925
               TabIndex        =   23
               Top             =   45
               Width           =   1140
               _Version        =   65536
               _ExtentX        =   2011
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
               Picture         =   "PaperDeliveryChallan.frx":0D0E
               Picture         =   "PaperDeliveryChallan.frx":0D2A
            End
            Begin TDBDate6Ctl.TDBDate MhDateInput1 
               Height          =   330
               Left            =   7050
               TabIndex        =   1
               Top             =   45
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
               _ExtentY        =   582
               Calendar        =   "PaperDeliveryChallan.frx":0D46
               Caption         =   "PaperDeliveryChallan.frx":0E5E
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PaperDeliveryChallan.frx":0ECA
               Keys            =   "PaperDeliveryChallan.frx":0EE8
               Spin            =   "PaperDeliveryChallan.frx":0F46
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
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   330
               Index           =   1
               Left            =   120
               TabIndex        =   27
               Top             =   480
               Width           =   3015
               _Version        =   65536
               _ExtentX        =   5318
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
               Caption         =   " Consignor's / Owner Name"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "PaperDeliveryChallan.frx":0F6E
               Picture         =   "PaperDeliveryChallan.frx":0F8A
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
               Height          =   330
               Index           =   0
               Left            =   120
               TabIndex        =   28
               Top             =   795
               Width           =   3015
               _Version        =   65536
               _ExtentX        =   5318
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
               Caption         =   " Consignor's Address"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "PaperDeliveryChallan.frx":0FA6
               Picture         =   "PaperDeliveryChallan.frx":0FC2
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
               Height          =   330
               Index           =   1
               Left            =   120
               TabIndex        =   29
               Top             =   1110
               Width           =   3015
               _Version        =   65536
               _ExtentX        =   5318
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
               Caption         =   " Consignor's Registration No."
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "PaperDeliveryChallan.frx":0FDE
               Picture         =   "PaperDeliveryChallan.frx":0FFA
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
               Height          =   330
               Index           =   2
               Left            =   120
               TabIndex        =   30
               Top             =   1680
               Width           =   3015
               _Version        =   65536
               _ExtentX        =   5318
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
               Caption         =   " Consignee Name"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "PaperDeliveryChallan.frx":1016
               Picture         =   "PaperDeliveryChallan.frx":1032
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
               Height          =   330
               Index           =   3
               Left            =   120
               TabIndex        =   31
               Top             =   1995
               Width           =   3015
               _Version        =   65536
               _ExtentX        =   5318
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
               Caption         =   " Consignor's Address"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "PaperDeliveryChallan.frx":104E
               Picture         =   "PaperDeliveryChallan.frx":106A
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
               Height          =   330
               Index           =   4
               Left            =   120
               TabIndex        =   32
               Top             =   2310
               Width           =   3015
               _Version        =   65536
               _ExtentX        =   5318
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
               Caption         =   " Consignee GSTIN Registration No."
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "PaperDeliveryChallan.frx":1086
               Picture         =   "PaperDeliveryChallan.frx":10A2
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
               Height          =   330
               Index           =   5
               Left            =   120
               TabIndex        =   33
               Top             =   2880
               Width           =   3015
               _Version        =   65536
               _ExtentX        =   5318
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
               Caption         =   " Carrier Name"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "PaperDeliveryChallan.frx":10BE
               Picture         =   "PaperDeliveryChallan.frx":10DA
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
               Height          =   330
               Index           =   7
               Left            =   120
               TabIndex        =   34
               Top             =   3510
               Width           =   3015
               _Version        =   65536
               _ExtentX        =   5318
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
               Caption         =   " Vehicle No."
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "PaperDeliveryChallan.frx":10F6
               Picture         =   "PaperDeliveryChallan.frx":1112
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
               Height          =   330
               Index           =   8
               Left            =   5895
               TabIndex        =   35
               Top             =   3510
               Width           =   2415
               _Version        =   65536
               _ExtentX        =   4260
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
               Caption         =   " Destination of goods."
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "PaperDeliveryChallan.frx":112E
               Picture         =   "PaperDeliveryChallan.frx":114A
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
               Height          =   330
               Index           =   9
               Left            =   120
               TabIndex        =   36
               Top             =   3825
               Width           =   3015
               _Version        =   65536
               _ExtentX        =   5318
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
               Caption         =   " Destination Address."
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "PaperDeliveryChallan.frx":1166
               Picture         =   "PaperDeliveryChallan.frx":1182
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
               Height          =   330
               Index           =   10
               Left            =   120
               TabIndex        =   37
               Top             =   4140
               Width           =   3015
               _Version        =   65536
               _ExtentX        =   5318
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
               Caption         =   " E-Way Bill No."
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "PaperDeliveryChallan.frx":119E
               Picture         =   "PaperDeliveryChallan.frx":11BA
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
               Height          =   330
               Index           =   11
               Left            =   5895
               TabIndex        =   38
               Top             =   4140
               Width           =   2415
               _Version        =   65536
               _ExtentX        =   4260
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
               Caption         =   " E-Way Bill Date."
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "PaperDeliveryChallan.frx":11D6
               Picture         =   "PaperDeliveryChallan.frx":11F2
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
               Height          =   330
               Index           =   13
               Left            =   120
               TabIndex        =   39
               Top             =   4455
               Width           =   3015
               _Version        =   65536
               _ExtentX        =   5318
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
               Caption         =   " Transit Form Number"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "PaperDeliveryChallan.frx":120E
               Picture         =   "PaperDeliveryChallan.frx":122A
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
               Height          =   330
               Index           =   14
               Left            =   5895
               TabIndex        =   40
               Top             =   4455
               Width           =   2415
               _Version        =   65536
               _ExtentX        =   4260
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
               Caption         =   " GSTIN Invoice Reference No."
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "PaperDeliveryChallan.frx":1246
               Picture         =   "PaperDeliveryChallan.frx":1262
            End
            Begin TDBDate6Ctl.TDBDate MhDateInput2 
               Height          =   330
               Left            =   8295
               TabIndex        =   12
               Top             =   4140
               Width           =   2340
               _Version        =   65536
               _ExtentX        =   4119
               _ExtentY        =   582
               Calendar        =   "PaperDeliveryChallan.frx":127E
               Caption         =   "PaperDeliveryChallan.frx":1396
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PaperDeliveryChallan.frx":1402
               Keys            =   "PaperDeliveryChallan.frx":1420
               Spin            =   "PaperDeliveryChallan.frx":147E
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
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
               Height          =   285
               Left            =   120
               TabIndex        =   41
               Top             =   6120
               Width           =   10515
               _Version        =   65536
               _ExtentX        =   18547
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
               Picture         =   "PaperDeliveryChallan.frx":14A6
               Picture         =   "PaperDeliveryChallan.frx":14C2
               Begin TDBNumber6Ctl.TDBNumber MhRealInput1 
                  Height          =   285
                  Left            =   4680
                  TabIndex        =   42
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   1440
                  _Version        =   65536
                  _ExtentX        =   2549
                  _ExtentY        =   503
                  Calculator      =   "PaperDeliveryChallan.frx":14DE
                  Caption         =   "PaperDeliveryChallan.frx":14FE
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "PaperDeliveryChallan.frx":156A
                  Keys            =   "PaperDeliveryChallan.frx":1588
                  Spin            =   "PaperDeliveryChallan.frx":15D2
                  AlignHorizontal =   1
                  AlignVertical   =   0
                  Appearance      =   0
                  BackColor       =   16777215
                  BorderStyle     =   1
                  BtnPositioning  =   0
                  ClipMode        =   0
                  ClearAction     =   0
                  DecimalPoint    =   "."
                  DisplayFormat   =   "#####0.000"
                  EditMode        =   1
                  Enabled         =   -1
                  ErrorBeep       =   0
                  ForeColor       =   255
                  Format          =   "#####0.000"
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
               Begin TDBNumber6Ctl.TDBNumber MhRealInput3 
                  Height          =   285
                  Left            =   9230
                  TabIndex        =   43
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   1005
                  _Version        =   65536
                  _ExtentX        =   1764
                  _ExtentY        =   503
                  Calculator      =   "PaperDeliveryChallan.frx":15FA
                  Caption         =   "PaperDeliveryChallan.frx":161A
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "PaperDeliveryChallan.frx":1686
                  Keys            =   "PaperDeliveryChallan.frx":16A4
                  Spin            =   "PaperDeliveryChallan.frx":16EE
                  AlignHorizontal =   1
                  AlignVertical   =   0
                  Appearance      =   0
                  BackColor       =   16777215
                  BorderStyle     =   1
                  BtnPositioning  =   0
                  ClipMode        =   0
                  ClearAction     =   0
                  DecimalPoint    =   "."
                  DisplayFormat   =   "######0.00"
                  EditMode        =   1
                  Enabled         =   -1
                  ErrorBeep       =   0
                  ForeColor       =   255
                  Format          =   "######0.00"
                  HighlightText   =   0
                  MarginBottom    =   1
                  MarginLeft      =   1
                  MarginRight     =   1
                  MarginTop       =   1
                  MaxValue        =   9999999.99
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
               Begin TDBNumber6Ctl.TDBNumber MhRealInput2 
                  Height          =   285
                  Left            =   7345
                  TabIndex        =   45
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   1185
                  _Version        =   65536
                  _ExtentX        =   2081
                  _ExtentY        =   503
                  Calculator      =   "PaperDeliveryChallan.frx":1716
                  Caption         =   "PaperDeliveryChallan.frx":1736
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "PaperDeliveryChallan.frx":17A2
                  Keys            =   "PaperDeliveryChallan.frx":17C0
                  Spin            =   "PaperDeliveryChallan.frx":180A
                  AlignHorizontal =   1
                  AlignVertical   =   0
                  Appearance      =   0
                  BackColor       =   16777215
                  BorderStyle     =   1
                  BtnPositioning  =   0
                  ClipMode        =   0
                  ClearAction     =   0
                  DecimalPoint    =   "."
                  DisplayFormat   =   "#####0.000"
                  EditMode        =   1
                  Enabled         =   -1
                  ErrorBeep       =   0
                  ForeColor       =   255
                  Format          =   "#####0.000"
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
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
               Height          =   330
               Left            =   7800
               TabIndex        =   46
               Top             =   6915
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
               Caption         =   " Sub Total"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "PaperDeliveryChallan.frx":1832
               Picture         =   "PaperDeliveryChallan.frx":184E
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
               Height          =   330
               Index           =   0
               Left            =   7800
               TabIndex        =   47
               Top             =   7545
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
               Caption         =   " SGST @ 6.00%"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "PaperDeliveryChallan.frx":186A
               Picture         =   "PaperDeliveryChallan.frx":1886
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel15 
               Height          =   330
               Left            =   7800
               TabIndex        =   48
               Top             =   7230
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
               Caption         =   " CGST @ 6.00%"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "PaperDeliveryChallan.frx":18A2
               Picture         =   "PaperDeliveryChallan.frx":18BE
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel16 
               Height          =   330
               Left            =   7800
               TabIndex        =   49
               Top             =   6600
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
               Caption         =   " Cartage"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "PaperDeliveryChallan.frx":18DA
               Picture         =   "PaperDeliveryChallan.frx":18F6
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput4 
               Height          =   330
               Left            =   9360
               TabIndex        =   50
               TabStop         =   0   'False
               Top             =   6600
               Width           =   1290
               _Version        =   65536
               _ExtentX        =   2275
               _ExtentY        =   582
               Calculator      =   "PaperDeliveryChallan.frx":1912
               Caption         =   "PaperDeliveryChallan.frx":1932
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PaperDeliveryChallan.frx":199E
               Keys            =   "PaperDeliveryChallan.frx":19BC
               Spin            =   "PaperDeliveryChallan.frx":1A06
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "######0.00"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   255
               Format          =   "######0.00"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   9999999.99
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   1
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   1964834821
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput6 
               Height          =   330
               Left            =   9360
               TabIndex        =   51
               Top             =   7230
               Width           =   1290
               _Version        =   65536
               _ExtentX        =   2275
               _ExtentY        =   582
               Calculator      =   "PaperDeliveryChallan.frx":1A2E
               Caption         =   "PaperDeliveryChallan.frx":1A4E
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PaperDeliveryChallan.frx":1ABA
               Keys            =   "PaperDeliveryChallan.frx":1AD8
               Spin            =   "PaperDeliveryChallan.frx":1B22
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "######0.00"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "######0.00"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   9999999.99
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   1964834821
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput7 
               Height          =   330
               Left            =   9360
               TabIndex        =   52
               TabStop         =   0   'False
               Top             =   7545
               Width           =   1290
               _Version        =   65536
               _ExtentX        =   2275
               _ExtentY        =   582
               Calculator      =   "PaperDeliveryChallan.frx":1B4A
               Caption         =   "PaperDeliveryChallan.frx":1B6A
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PaperDeliveryChallan.frx":1BD6
               Keys            =   "PaperDeliveryChallan.frx":1BF4
               Spin            =   "PaperDeliveryChallan.frx":1C3E
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "######0.00"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   255
               Format          =   "######0.00"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   9999999.99
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   1
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   1964834821
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
               Height          =   330
               Index           =   1
               Left            =   7800
               TabIndex        =   53
               Top             =   7860
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
               Caption         =   " Total (Rounded) :"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "PaperDeliveryChallan.frx":1C66
               Picture         =   "PaperDeliveryChallan.frx":1C82
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput8 
               Height          =   330
               Left            =   9360
               TabIndex        =   54
               TabStop         =   0   'False
               Top             =   7860
               Width           =   1290
               _Version        =   65536
               _ExtentX        =   2275
               _ExtentY        =   582
               Calculator      =   "PaperDeliveryChallan.frx":1C9E
               Caption         =   "PaperDeliveryChallan.frx":1CBE
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PaperDeliveryChallan.frx":1D2A
               Keys            =   "PaperDeliveryChallan.frx":1D48
               Spin            =   "PaperDeliveryChallan.frx":1D92
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "######0.00"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   255
               Format          =   "######0.00"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   9999999.99
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   1
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   1964834821
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
               Height          =   330
               Index           =   12
               Left            =   8160
               TabIndex        =   55
               Top             =   45
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
               Caption         =   " P.O Number"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "PaperDeliveryChallan.frx":1DBA
               Picture         =   "PaperDeliveryChallan.frx":1DD6
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
               Height          =   330
               Index           =   6
               Left            =   120
               TabIndex        =   56
               Top             =   3195
               Width           =   3015
               _Version        =   65536
               _ExtentX        =   5318
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
               Caption         =   " Carrier Address"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "PaperDeliveryChallan.frx":1DF2
               Picture         =   "PaperDeliveryChallan.frx":1E0E
            End
            Begin FPSpreadADO.fpSpread fpSpread1 
               Height          =   1095
               Left            =   120
               TabIndex        =   57
               Top             =   5040
               Width           =   10515
               _Version        =   524288
               _ExtentX        =   18547
               _ExtentY        =   1940
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
               MaxCols         =   7
               MaxRows         =   1000
               ScrollBars      =   2
               SpreadDesigner  =   "PaperDeliveryChallan.frx":1E2A
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput9 
               Height          =   330
               Left            =   5760
               TabIndex        =   58
               Top             =   7710
               Visible         =   0   'False
               Width           =   930
               _Version        =   65536
               _ExtentX        =   1640
               _ExtentY        =   582
               Calculator      =   "PaperDeliveryChallan.frx":2737
               Caption         =   "PaperDeliveryChallan.frx":2757
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PaperDeliveryChallan.frx":27C3
               Keys            =   "PaperDeliveryChallan.frx":27E1
               Spin            =   "PaperDeliveryChallan.frx":282B
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
               ValueVT         =   1967259653
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput10 
               Height          =   330
               Left            =   5760
               TabIndex        =   59
               Top             =   8025
               Visible         =   0   'False
               Width           =   930
               _Version        =   65536
               _ExtentX        =   1640
               _ExtentY        =   582
               Calculator      =   "PaperDeliveryChallan.frx":2853
               Caption         =   "PaperDeliveryChallan.frx":2873
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PaperDeliveryChallan.frx":28DF
               Keys            =   "PaperDeliveryChallan.frx":28FD
               Spin            =   "PaperDeliveryChallan.frx":2947
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
               ValueVT         =   1967259653
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput12 
               Height          =   330
               Left            =   1560
               TabIndex        =   60
               Top             =   7080
               Visible         =   0   'False
               Width           =   810
               _Version        =   65536
               _ExtentX        =   1429
               _ExtentY        =   582
               Calculator      =   "PaperDeliveryChallan.frx":296F
               Caption         =   "PaperDeliveryChallan.frx":298F
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PaperDeliveryChallan.frx":29FB
               Keys            =   "PaperDeliveryChallan.frx":2A19
               Spin            =   "PaperDeliveryChallan.frx":2A63
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
               ValueVT         =   1967259653
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput13 
               Height          =   330
               Left            =   3480
               TabIndex        =   61
               Top             =   7080
               Visible         =   0   'False
               Width           =   810
               _Version        =   65536
               _ExtentX        =   1429
               _ExtentY        =   582
               Calculator      =   "PaperDeliveryChallan.frx":2A8B
               Caption         =   "PaperDeliveryChallan.frx":2AAB
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PaperDeliveryChallan.frx":2B17
               Keys            =   "PaperDeliveryChallan.frx":2B35
               Spin            =   "PaperDeliveryChallan.frx":2B7F
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "######0.00"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "######0.00"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   9999999.99
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   1967259653
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Left            =   4320
               TabIndex        =   62
               Top             =   7080
               Visible         =   0   'False
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
               Caption         =   " Cartage/Kg"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "PaperDeliveryChallan.frx":2BA7
               Picture         =   "PaperDeliveryChallan.frx":2BC3
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput11 
               Height          =   330
               Left            =   5760
               TabIndex        =   63
               Top             =   7080
               Visible         =   0   'False
               Width           =   1530
               _Version        =   65536
               _ExtentX        =   2699
               _ExtentY        =   582
               Calculator      =   "PaperDeliveryChallan.frx":2BDF
               Caption         =   "PaperDeliveryChallan.frx":2BFF
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PaperDeliveryChallan.frx":2C6B
               Keys            =   "PaperDeliveryChallan.frx":2C89
               Spin            =   "PaperDeliveryChallan.frx":2CD3
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
               ValueVT         =   1967259653
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel10 
               Height          =   330
               Index           =   0
               Left            =   120
               TabIndex        =   64
               Top             =   7080
               Visible         =   0   'False
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
               Caption         =   " Reams/Bundle"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "PaperDeliveryChallan.frx":2CFB
               Picture         =   "PaperDeliveryChallan.frx":2D17
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel10 
               Height          =   330
               Index           =   1
               Left            =   2280
               TabIndex        =   65
               Top             =   7080
               Visible         =   0   'False
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
               Caption         =   " Total Bundle"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "PaperDeliveryChallan.frx":2D33
               Picture         =   "PaperDeliveryChallan.frx":2D4F
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput5 
               Height          =   330
               Left            =   9360
               TabIndex        =   66
               Top             =   6915
               Width           =   1290
               _Version        =   65536
               _ExtentX        =   2275
               _ExtentY        =   582
               Calculator      =   "PaperDeliveryChallan.frx":2D6B
               Caption         =   "PaperDeliveryChallan.frx":2D8B
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PaperDeliveryChallan.frx":2DF7
               Keys            =   "PaperDeliveryChallan.frx":2E15
               Spin            =   "PaperDeliveryChallan.frx":2E5F
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "######0.00"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "######0.00"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   9999999.99
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
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
               Height          =   330
               Index           =   15
               Left            =   5880
               TabIndex        =   67
               Top             =   2310
               Width           =   2415
               _Version        =   65536
               _ExtentX        =   4260
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
               Caption         =   " Destination Type."
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "PaperDeliveryChallan.frx":2E87
               Picture         =   "PaperDeliveryChallan.frx":2EA3
            End
            Begin MSForms.ComboBox Combo1 
               Height          =   330
               Left            =   8295
               TabIndex        =   68
               Top             =   2310
               Width           =   2345
               VariousPropertyBits=   545282075
               BackColor       =   16777215
               BorderStyle     =   1
               DisplayStyle    =   7
               Size            =   "4136;582"
               ListRows        =   3
               MatchEntry      =   0
               ShowDropButtonWhen=   1
               SpecialEffect   =   0
               FontName        =   "Calibri"
               FontHeight      =   195
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Line Line2 
               Index           =   2
               X1              =   0
               X2              =   13240
               Y1              =   4920
               Y2              =   4920
            End
            Begin VB.Line Line2 
               Index           =   1
               X1              =   0
               X2              =   13240
               Y1              =   2760
               Y2              =   2760
            End
            Begin VB.Line Line1 
               X1              =   0
               X2              =   13240
               Y1              =   405
               Y2              =   405
            End
            Begin VB.Line Line2 
               Index           =   0
               X1              =   0
               X2              =   13240
               Y1              =   1575
               Y2              =   1575
            End
            Begin VB.Line Line4 
               X1              =   0
               X2              =   13240
               Y1              =   6480
               Y2              =   6480
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   330
            Left            =   120
            TabIndex        =   20
            Top             =   8430
            Width           =   495
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   11385
      _ExtentX        =   20082
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
Attribute VB_Name = "FrmPaperDeliveryChallan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CxnPaperPurchaseOrder As New ADODB.Connection
Dim rstPaperPOList As New ADODB.Recordset
Dim rstPaperPOParent As New ADODB.Recordset
Dim rstPaperPOChild As New ADODB.Recordset
Dim rstBookRef As New ADODB.Recordset
Dim rstSupplierList As New ADODB.Recordset
Dim rstPaperList As New ADODB.Recordset
Dim rstAccountList As New ADODB.Recordset
Dim rstBookList As New ADODB.Recordset
Dim rstLastPurchaseRate As New ADODB.Recordset
Dim SupplierCode As String, POCode As String, CompanyCode As String, AccountCode As String, PaperCode As String, BookCode As String, SizeCode As String, RefCode As String
Dim SortOrder, PrevStr
Dim dblBookMark As Double
Dim blnRecordExist As Boolean
Dim EditMode As Boolean
Dim oOutlook As New Outlook.Application
Dim EMailID As String
Dim Attachment As String
Dim Message As String
Public OrderType
Public PrinterCode As String
Dim rstPrinterRates As New ADODB.Recordset
Dim rstPlaningRef As New ADODB.Recordset
Dim CartridgeVat As Double
Dim rstRefList As New ADODB.Recordset
Dim rstPaperPurchaseList As New ADODB.Recordset
Dim rstCompanyMasterList As New ADODB.Recordset
Dim OutputType As String
Dim rstPaperPurchaseIssueList As New ADODB.Recordset

'UPDATE PaperMaster INNER JOIN PaperMasterUpdate ON PaperMaster.Code = PaperMasterUpdate.Code
'Set PaperMaster.HSNCode = PaperMasterUpdate.HSNCode

Private Sub Form_Load()
    On Error GoTo ErrorHandler
    CenterForm Me
    WheelHook DataGrid1
    BusySystemIndicator True
    Dim Cn As Integer
    Me.Caption = "Paper Delivery Challan [" & IIf(OrderType = "1", "Book", "Title") & "]" '"Paper Delivery Challan"
    CompanyCode = "000001"
    CxnPaperPurchaseOrder.CursorLocation = adUseClient
    CxnPaperPurchaseOrder.Open CxnDatabase.ConnectionString
    rstPaperPurchaseList.Open "SELECT TRIM(T.Name) + '-' + TRIM(M.Name) As Col0,T.Code, T.Name As POCode,TRIM(M.Name) As Supplier, M.Address1, M.Address2, M.Address3, M.Address4,M.TIN,M.Phone, M.Mobile, M.EMail,M.Code As SupplierCode,M.Name As SupplierName,[Reams/Bundle],Bundles,[VAT%],VAT,[Cartage/Bundle],Cartage,BillAmount FROM PaperPOParent T INNER JOIN AccountMaster M ON T.Supplier=M.Code WHERE OrderType='" & OrderType & "' ORDER BY T.Name", CxnPaperPurchaseOrder, adOpenKeyset, adLockOptimistic
    rstPaperPOList.Open "SELECT T.Code,T.Name,Date,M.Name As ConsigneeName,BillAmount FROM DeliveryChallanParent T INNER JOIN AccountMaster M ON T.Consignee=M.Code WHERE OrderType='" & OrderType & "' ORDER BY T.Name", CxnPaperPurchaseOrder, adOpenKeyset, adLockOptimistic
    rstPaperPurchaseIssueList.Open "SELECT  T.Code, TRIM(M.Name) As Supplier, M.Address1, M.Address2, M.Address3, M.Address4,M.TIN,M.Phone, M.Mobile, M.EMail,M.Code As SupplierCode,M.Name As SupplierName FROM PaperIOChild T INNER JOIN AccountMaster M  ON T.Account=M.Code ", CxnPaperPurchaseOrder, adOpenKeyset, adLockOptimistic
    
    Combo1.AddItem "Delhi", 0
    Combo1.AddItem "Out of Station", 1
   
    rstPaperPOParent.CursorLocation = adUseClient
    rstPaperPOList.Filter = adFilterNone
    If rstPaperPOList.RecordCount > 0 Then rstPaperPOList.MoveLast
    Set DataGrid1.DataSource = rstPaperPOList
    BusySystemIndicator False
    SSTab1.Tab = 0
    SortOrder = "Name"
    If Not (rstPaperPOList.EOF Or rstPaperPOList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    
    rstPaperPOList.ActiveConnection = Nothing
    rstPaperPurchaseList.ActiveConnection = Nothing
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
                If Not EditMode Then
                    If MsgBox("Are you sure to Quit?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Quit !") <> vbYes Then
                        Me.ActiveControl.SetFocus
                    Else
                        Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
                    End If
                End If
            End If
            If Not EditMode Then KeyCode = 0
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
        If Not EditMode Then Toolbar1_ButtonClick Toolbar1.Buttons.Item(4)
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
    
        If Toolbar1.Buttons.Item(1).Enabled Then
            SSTab1.Tab = 1: SSTab1.SetFocus
        Else
           If Me.ActiveControl.Name <> "fpSpread1" Then SendKeys "{TAB}"
        End If
        If Me.ActiveControl.Name <> "fpSpread1" Then KeyCode = 0
    
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Toolbar1.Buttons.Item(4).Enabled Then Call Form_KeyDown(vbKeyEscape, 0): Cancel = 1
End Sub
Private Sub Form_Unload(Cancel As Integer)
    WheelUnHook
    Call CloseRecordset(rstPaperPOList)
    Call CloseRecordset(rstPaperPOParent)
    Call CloseRecordset(rstPaperPOChild)
    Call CloseRecordset(rstPaperPurchaseIssueList)
    
    Call CloseConnection(CxnPaperPurchaseOrder)
    ShowProgressInStatusBar False
    DisableChildMenu
End Sub

Private Sub Text1_Change()
    If rstPaperPOList.RecordCount = 0 Then Exit Sub
    rstPaperPOList.MoveFirst
    If Text1.Text <> "" Then
        If SortOrder = "Name" Then
           rstPaperPOList.Find "[" & SortOrder & "] Like '%" & FixQuote(Text1.Text) & "%'"
        Else
           rstPaperPOList.Find "[" & SortOrder & "] Like '" & FixQuote(Text1.Text) & "%'"
        End If
        If rstPaperPOList.EOF Then
            rstPaperPOList.MoveFirst
            If PrevStr <> "" And Len(Text1.Text) > 1 Then
                If dblBookMark <> 0 Then
                    rstPaperPOList.Bookmark = dblBookMark
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
    If Not (rstPaperPOList.EOF Or rstPaperPOList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim KeyProcessed As Boolean
    If rstPaperPOList.RecordCount = 0 Then Exit Sub
    If Shift = 0 And KeyCode = vbKeyUp Then
        With rstPaperPOList
            .MovePrevious
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyBack Then
        With rstPaperPOList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyDown Then
        With rstPaperPOList
            .MoveNext
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageUp Then
        With rstPaperPOList
            .Move (-1) * (DataGrid1.VisibleRows - 1)
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageUp Then
        With rstPaperPOList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageDown Then
        With rstPaperPOList
            .Move DataGrid1.VisibleRows - 1
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageDown Then
        With rstPaperPOList
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
            If Not (rstPaperPOList.EOF Or rstPaperPOList.BOF) Then
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
        Text3.SetFocus
    End If
End Sub
Public Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim HiLiteRecord As Boolean
    Dim UpdateFlag As Integer
    Dim CellVal01 As Variant, CellVal02 As Variant, CellVal03 As Variant, CellVal04 As Variant, i As Integer
    If Button.Index = 1 Then
        If rstPaperPOParent.State = adStateOpen Then rstPaperPOParent.Close
        rstPaperPOParent.Open "SELECT * FROM DeliveryChallanParent WHERE Code=''", CxnPaperPurchaseOrder, adOpenKeyset, adLockOptimistic
        ClearFields
        Call LoadPaperList("")

        If AddRecord(rstPaperPOParent) Then
            Text2.Text = GenerateCode(CxnPaperPurchaseOrder, "SELECT MAX(VAL(Name)) FROM DeliveryChallanParent WHERE OrderType='" & OrderType & "'", 10, Space(1))
            MhDateInput1.Text = Format(Date, "dd-MM-yyyy")
            Call SetButtons(False)
            SSTab1.Tab = 1
            Text3.SetFocus
            blnRecordExist = False
            CxnPaperPurchaseOrder.BeginTrans
        End If
    ElseIf Button.Index = 2 Then
        If rstPaperPOList.RecordCount = 0 Then Exit Sub
        SSTab1.Tab = 1
        EditRecord
    ElseIf Button.Index = 3 Then
        If rstPaperPOList.RecordCount = 0 Then Exit Sub
        If AllowTransactionsDeletion = 0 Then Call DisplayError("You don't have the rights to Delete this Voucher"): Exit Sub
        SSTab1.Tab = 1
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            On Error Resume Next
            MdiMainMenu.MousePointer = vbHourglass
            CxnPaperPurchaseOrder.Execute "DELETE FROM DeliveryChallanParent WHERE Code='" & rstPaperPOList.Fields("Code").Value & "'"
            
            MdiMainMenu.MousePointer = vbNormal
            If Err.Number = 0 Then
                rstPaperPOList.Delete
                rstPaperPOList.MoveNext
                If rstPaperPOList.RecordCount > 0 And rstPaperPOList.EOF Then rstPaperPOList.MoveLast
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
        SaveFields
        UpdateFlag = 0
        If UpdateRecord(rstPaperPOParent) Then
            If UpdatePaperList("D") Then
                UpdateFlag = 1
                With fpSpread1
                    For i = 1 To .DataRowCnt
                        .SetActiveCell 6, i
                        .GetText 6, i, CellVal01
                        .GetText 7, i, CellVal02
                        If Val(CellVal01) <> 0 And CellVal02 <> "" Then
                            If Not UpdatePaperList("I1") Then UpdateFlag = 0: Exit For
                        End If
                    Next
                End With
                
                If UpdateFlag = 1 Then
                 
                End If
                If UpdateFlag = 1 Then
               
              End If
                
            End If
        End If
        If UpdateFlag Then
            AddToList
            CxnPaperPurchaseOrder.CommitTrans
            If rstPaperPOParent.State = adStateOpen Then rstPaperPOParent.Close
            rstPaperPOParent.CursorLocation = adUseClient
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
        If CancelRecordUpdate(rstPaperPOParent) Then
            CxnPaperPurchaseOrder.RollbackTrans
            If rstPaperPOParent.State = adStateOpen Then rstPaperPOParent.Close
            rstPaperPOParent.CursorLocation = adUseClient
            Call SetButtons(True)
            SetButtonsForNoRecord
            SSTab1.Tab = 0
            LockFields (False)
        End If
    ElseIf Button.Index = 6 Then
        SSTab1.Tab = 0
        Set DataGrid1.DataSource = Nothing
        rstPaperPOList.ActiveConnection = CxnPaperPurchaseOrder
        Do While Not RefreshRecord(rstPaperPOList)
        Loop
        Set DataGrid1.DataSource = rstPaperPOList
        rstPaperPOList.ActiveConnection = Nothing
        If rstPaperPOList.RecordCount > 0 Then rstPaperPOList.MoveLast
        rstSupplierList.ActiveConnection = CxnPaperPurchaseOrder
        
        Do While Not RefreshRecord(rstSupplierList)
        Loop
        
        rstSupplierList.ActiveConnection = Nothing
        rstPaperList.ActiveConnection = CxnPaperPurchaseOrder
        
        Do While Not RefreshRecord(rstPaperList)
        Loop
        
        rstPaperList.ActiveConnection = Nothing
        rstAccountList.ActiveConnection = CxnPaperPurchaseOrder
        
        Do While Not RefreshRecord(rstAccountList)
        Loop
        
        rstAccountList.ActiveConnection = Nothing
        rstBookList.ActiveConnection = CxnPaperPurchaseOrder
        Do While Not RefreshRecord(rstBookList)
        Loop
        
        rstBookList.ActiveConnection = Nothing
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
        If rstPaperPOList.RecordCount = 0 Then Exit Sub
        Call DisplayMenu("P")
        OutputType = "P"
        'Call PrintPaperDeliveryChallan(rstPaperPOList.Fields("Code").Value, OrderType, "", OutputType, 1)
        HiLiteRecord = True
    ElseIf Button.Index = 10 Then
        If rstPaperPOList.RecordCount = 0 Then Exit Sub
        Call DisplayMenu("S")
        OutputType = "S"
        'Call PrintPaperDeliveryChallan(rstPaperPOList.Fields("Code").Value, OrderType, "", OutputType, 1)
        HiLiteRecord = True
    ElseIf Button.Index = 11 Then
        If rstPaperPOList.RecordCount = 0 Then Exit Sub
        OutputType = "M"
        Call PrintPaperDeliveryChallan(rstPaperPOList.Fields("Code").Value, OrderType, "", OutputType, 1)
        HiLiteRecord = True
    ElseIf Button.Index = 13 Then
        If rstPaperPOList.RecordCount > 0 Then rstPaperPOList.MoveFirst
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 14 Then
        If rstPaperPOList.RecordCount > 0 Then
            rstPaperPOList.MovePrevious
            If rstPaperPOList.BOF Then rstPaperPOList.MoveNext
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 15 Then
        If rstPaperPOList.RecordCount > 0 Then
            rstPaperPOList.MoveNext
            If rstPaperPOList.EOF Then rstPaperPOList.MovePrevious
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 16 Then
        If rstPaperPOList.RecordCount > 0 Then rstPaperPOList.MoveLast
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 18 Then
        Unload Me
        HiLiteRecord = False
    End If
    
    If HiLiteRecord Then
        If Not (rstPaperPOList.EOF Or rstPaperPOList.BOF) Then
            With DataGrid1.SelBookmarks
                If .Count <> 0 Then .Remove 0
                .Add DataGrid1.Bookmark
            End With
        End If
        Text1.SetFocus
        
    End If
    
End Sub

Private Sub DataGrid1_DblClick()
    If Toolbar1.Buttons.Item(2).Enabled Then Toolbar1_ButtonClick Toolbar1.Buttons.Item(2)
End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
    If ColIndex = 0 Or ColIndex = 2 Then
        SortOrder = DataGrid1.Columns(ColIndex).DataField
        rstPaperPOList.Sort = "[" + SortOrder & "] Asc"
    End If
    DataGrid1.ClearSelCols
    If Not (rstPaperPOList.EOF Or rstPaperPOList.BOF) Then
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
    If rstPaperPOList.RecordCount = 0 Then
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
    
    If rstPaperPOParent.EOF Or rstPaperPOParent.BOF Then Exit Sub
    If CheckEmpty(Text2, True) Then
        Cancel = True
    ElseIf CheckDuplicate(CxnPaperPurchaseOrder, "DeliveryChallanParent", "Code", "[Name]+OrderType", Trim(Text2.Text) & "1", rstPaperPOParent.Fields("Code").Value, False) Then
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
    If Text3.Text = " " Then Text3.Text = "?": SendKeys "{TAB}"
End Sub
Private Sub Text3_Validate(Cancel As Boolean)
    Dim SearchString As String
    SearchString = FixQuote(Text3.Text)
    If rstPaperPurchaseList.RecordCount = 0 Then DisplayError ("No Record in Supplier Master"): Cancel = True: Exit Sub Else rstPaperPurchaseList.MoveFirst
    rstPaperPurchaseList.Find "[Col0] = '" & RTrim(SearchString) & "'"
    If rstPaperPurchaseList.EOF Then
        SelectionType = "S"
        POCode = ""
        Call LoadSelectionList(rstPaperPurchaseList, "List of Purchase Order...", "PO With Name")
        SearchOrder = 0
        
        Call DisplaySelectionList(Text3, POCode)
        Call CloseForm(FrmSelectionList)
        If CheckEmpty(Text3.Text, False) Then Text3.Text = "?"
        If RTrim(POCode) <> "" Then SendKeys "{TAB}"
        Cancel = True
        
    Else
        
        POCode = rstPaperPurchaseList.Fields("Code").Value
        Text3.Text = Trim(rstPaperPurchaseList.Fields("POCode").Value)
        Text3.Tag = rstPaperPurchaseList.Fields("Code").Value
        
        If rstPaperPurchaseIssueList.RecordCount > 0 Then rstPaperPurchaseIssueList.MoveFirst
        rstPaperPurchaseIssueList.Find "[Code] = " & Text3.Tag & ""
        SupplierCode = rstPaperPurchaseIssueList.Fields("SupplierCode").Value
        txtConsigneeName.Text = rstPaperPurchaseIssueList.Fields("Supplier").Value
        txtConsigneeAddress.Text = rstPaperPurchaseIssueList.Fields("Address1").Value & " " & rstPaperPurchaseIssueList.Fields("Address2").Value & " " & rstPaperPurchaseIssueList.Fields("Address3").Value & " " & rstPaperPurchaseIssueList.Fields("Address4").Value
        txtConsigneeRegNo.Text = rstPaperPurchaseIssueList.Fields("TIN").Value
        
'        If Not IsNull(rstPaperPOParent.Fields("Address1").Value) Then
'            txtConsigneeAddress.Text = rstPaperPurchaseIssueList.Fields("Address1").Value
'        End If
        
        
        MhRealInput12.Value = Val(rstPaperPurchaseList.Fields("Reams/Bundle").Value)
        MhRealInput13.Value = Val(rstPaperPurchaseList.Fields("Bundles").Value)
        
        MhRealInput11.Value = Val(rstPaperPurchaseList.Fields("Cartage/Bundle").Value)
        MhRealInput4.Value = Val(rstPaperPurchaseList.Fields("Cartage").Value)
        
        
        MhRealInput9.Value = Val(rstPaperPurchaseList.Fields("VAT%").Value) / 2
        MhRealInput6.Value = Val(rstPaperPurchaseList.Fields("VAT").Value) / 2
        MhRealInput10.Value = Val(rstPaperPurchaseList.Fields("VAT%").Value) / 2
        MhRealInput7.Value = Val(rstPaperPurchaseList.Fields("VAT").Value) / 2
        MhRealInput8.Value = Val(rstPaperPurchaseList.Fields("BillAmount").Value)
        LoadPaperFromPO (Text3.Tag)
        CalculateTotal ("G")
        MhRealInput5.Value = Val(MhRealInput3.Value) + Val(MhRealInput4.Value) 'Sub total

    End If
End Sub
Private Sub MhDateInput2_Validate(Cancel As Boolean)
    If MhDateInput2.ValueIsNull Then Exit Sub
    If Not IsDate(GetDate(MhDateInput2.Text)) Then Cancel = True
End Sub
Private Sub ViewRecord()
    ClearFields
    If rstPaperPOList.EOF Then Exit Sub
    FindRecord
    LoadFields
End Sub
Private Sub FindRecord()
    If rstPaperPOParent.State = adStateOpen Then rstPaperPOParent.Close
    rstPaperPOParent.Open "SELECT * FROM DeliveryChallanParent WHERE Code='" & FixQuote(rstPaperPOList.Fields("Code").Value) & "'", CxnPaperPurchaseOrder, adOpenKeyset, adLockOptimistic
    If rstPaperPOParent.RecordCount = 0 Then
       Call DisplayError("This Record has been deleted by Another User ! Click Ok To Refresh the Recordset")
       Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
    End If
End Sub
Private Sub ClearFields()
    CartridgeVat = 0
    txtConsigneeName.Text = ""
    txtConsigneeAddress.Text = ""
    txtConsigneeRegNo.Text = ""
    Combo1.ListIndex = 0
    txtCarrierName.Text = ""
    txtCarrierAddress.Text = ""
    txtVehicleNo.Text = ""
    txtDestinationGoods.Text = ""
    txtDestinationAddress.Text = ""
    txtEWayBillNo.Text = ""
    txtTransitFormNumber.Text = ""
    txtRefNo.Text = ""
    
    MhDateInput1.Text = Format(Date, "dd-MM-yyyy") 'Challan Date
    MhDateInput2.Text = "  -  -    " 'EWay Bill Date
    MhRealInput1.Value = 0 'Total Quantity (Ream)
    MhRealInput2.Value = 0 'Total Quantity (Kg)
    MhRealInput3.Value = 0 'Total Gross Amount
    MhRealInput4.Value = 0  'Reams/bundle
    MhRealInput5.Value = 0  'Total bundles
    MhRealInput6.Value = 0   'Cartage/Kg
    
    MhRealInput7.Value = 0 'CGST
    MhRealInput8.Value = 0 'SGST
    
    MhRealInput9.Value = 6   'CGST%
    MhRealInput10.Value = 6 'SGST%
'    MhRealInput9.Visible = False    'CGST%
'    MhRealInput10.Visible = False 'SGST%
    
    fpSpread1.ClearRange 1, 1, fpSpread1.MaxCols, fpSpread1.MaxRows, True: fpSpread1.SetActiveCell 1, 1

End Sub
Private Sub LoadFields()
    
    If rstPaperPOParent.EOF Or rstPaperPOParent.BOF Then Exit Sub
    Text2.Text = rstPaperPOParent.Fields("Name").Value
    MhDateInput1.Text = Format(rstPaperPOParent.Fields("Date").Value, "dd-MM-yyyy")

    Dim PO As String
    PO = rstPaperPOParent.Fields("PO").Value
    '***************Consignee Details********************************************
    If rstPaperPurchaseList.RecordCount > 0 Then rstPaperPurchaseList.MoveFirst
    rstPaperPurchaseList.Find "[Code] = '" & PO & "'"
    Text3.Text = Trim(rstPaperPurchaseList.Fields("POCode").Value)
    Text3.Tag = Trim(rstPaperPurchaseList.Fields("Code").Value)
    POCode = Trim(rstPaperPurchaseList.Fields("Code").Value)
    SupplierCode = rstPaperPurchaseList.Fields("SupplierCode").Value
    If Not IsNull(rstPaperPOParent.Fields("ConsigneeName").Value) Then
       txtConsigneeName.Text = rstPaperPOParent.Fields("ConsigneeName").Value
    End If
    If Not IsNull(rstPaperPOParent.Fields("ConsigneeAddress").Value) Then
      txtConsigneeAddress.Text = rstPaperPOParent.Fields("ConsigneeAddress").Value
    End If
    If Not IsNull(rstPaperPOParent.Fields("ConsigneeGSTIN").Value) Then
       txtConsigneeRegNo.Text = rstPaperPOParent.Fields("ConsigneeGSTIN").Value
    End If
    If rstPaperPOParent.Fields("DType").Value = "1" Then
       Combo1.ListIndex = 0
    Else
       Combo1.ListIndex = 1
    End If
    
    '****************End********************************************************
    txtCarrierName.Text = rstPaperPOParent.Fields("CarrierName").Value
    txtCarrierAddress.Text = rstPaperPOParent.Fields("CarrierAddress").Value
    txtVehicleNo.Text = rstPaperPOParent.Fields("VehicleNo").Value
    txtDestinationGoods.Text = rstPaperPOParent.Fields("DestinationGoods").Value
    txtDestinationAddress.Text = rstPaperPOParent.Fields("DestinationAddress").Value
    txtEWayBillNo.Text = rstPaperPOParent.Fields("EWayBillNo").Value
    If Not IsNull(rstPaperPOParent.Fields("EWayBillDate").Value) Then MhDateInput2.Text = Format(rstPaperPOParent.Fields("EWayBillDate").Value, "dd-MM-yyyy")
    txtTransitFormNumber.Text = rstPaperPOParent.Fields("TransitFormNo").Value
    txtRefNo.Text = rstPaperPOParent.Fields("GSTRefNo").Value
    
    MhRealInput12.Value = Val(rstPaperPOParent.Fields("Reams_Bundle").Value)
    MhRealInput13.Value = Val(rstPaperPOParent.Fields("Bundles").Value)
    
    MhRealInput11.Value = Val(rstPaperPOParent.Fields("Cartage_Bundle").Value)
    MhRealInput4.Value = Val(rstPaperPOParent.Fields("Cartage").Value)
    
    MhRealInput9.Value = Val(rstPaperPOParent.Fields("CGST%").Value)
    MhRealInput6.Value = Val(rstPaperPOParent.Fields("CGST").Value)
    MhRealInput10.Value = Val(rstPaperPOParent.Fields("SGST%").Value)
    MhRealInput7.Value = Val(rstPaperPOParent.Fields("SGST").Value)
    MhRealInput8.Value = Val(rstPaperPOParent.Fields("BillAmount").Value)
    Call LoadPaperList(rstPaperPOParent.Fields("Code").Value)
    CalculateTotal ("G")
    MhRealInput5.Value = Val(MhRealInput3.Value) + Val(MhRealInput4.Value) 'Sub total
End Sub

Private Sub EditRecord()
    On Error GoTo ErrorHandler
    If rstPaperPOParent.RecordCount = 0 Then Exit Sub
    Set rstPaperPOParent = Nothing
    If rstPaperPOParent.State = adStateOpen Then rstPaperPOParent.Close
    
    rstPaperPOParent.CursorLocation = adUseServer
    rstPaperPOParent.Open "SELECT * FROM DeliveryChallanParent WHERE Code='" & FixQuote(rstPaperPOList.Fields("Code").Value) & "'", CxnPaperPurchaseOrder, adOpenKeyset, adLockPessimistic
    MdiMainMenu.MousePointer = vbHourglass
    rstPaperPOParent.Fields("Printstatus") = "N"
    MdiMainMenu.MousePointer = vbNormal
    
    AddToList
    
    Call SetButtons(False)
    SSTab1.TabEnabled(0) = False
    txtConsigneeName.SetFocus
    blnRecordExist = True
    If AllowTransactionsModification = 0 Then
        Text1.Locked = False
    End If
    CxnPaperPurchaseOrder.BeginTrans
    Exit Sub
ErrorHandler:
    If Err.Number = -2147467259 Then
       Call DisplayError("Failed to Edit the record")
    End If
    MdiMainMenu.MousePointer = vbNormal
    SSTab1.Tab = 0
End Sub
Private Sub SaveFields()
    If rstPaperPOParent.EOF Or rstPaperPOParent.BOF Then Exit Sub
    Dim lpBuff As String * 1024
    
    GetComputerName lpBuff, Len(lpBuff)
    
    If Not blnRecordExist Then
        
        rstPaperPOParent.Fields("Code").Value = GenerateCode(CxnPaperPurchaseOrder, "SELECT MAX(Code) FROM DeliveryChallanParent", 6, "0")
        
        rstPaperPOParent.Fields("CreatedBy").Value = UserCode
        rstPaperPOParent.Fields("CreatedOn").Value = Now()
        rstPaperPOParent.Fields("Recordstatus").Value = "N"
    Else
        
        rstPaperPOParent.Fields("ModifiedBy").Value = UserCode
        rstPaperPOParent.Fields("ModifiedOn").Value = Now()
        rstPaperPOParent.Fields("Recordstatus").Value = "M"
        
    End If
    rstPaperPOParent.Fields("Name").Value = Pad(Trim(Text2.Text), Space(1), 10, "L")
    rstPaperPOParent.Fields("Date").Value = GetDate(MhDateInput1.Text)
    rstPaperPOParent.Fields("OrderType").Value = OrderType
   
    rstPaperPOParent.Fields("PO").Value = Text3.Tag
    
    
    rstPaperPOParent.Fields("Consignee").Value = SupplierCode
    rstPaperPOParent.Fields("Consignor").Value = CompanyCode
    
    rstPaperPOParent.Fields("ConsigneeName").Value = Trim(txtConsigneeName.Text)
    rstPaperPOParent.Fields("ConsigneeAddress").Value = Trim(txtConsigneeAddress.Text)
    rstPaperPOParent.Fields("ConsigneeGSTIN").Value = Trim(txtConsigneeRegNo.Text)
    
    rstPaperPOParent.Fields("CarrierName").Value = txtCarrierName.Text
    rstPaperPOParent.Fields("CarrierAddress").Value = txtCarrierAddress.Text
    rstPaperPOParent.Fields("VehicleNo").Value = txtVehicleNo.Text
    rstPaperPOParent.Fields("DestinationGoods").Value = txtDestinationGoods.Text
    rstPaperPOParent.Fields("DestinationAddress").Value = txtDestinationAddress.Text
    rstPaperPOParent.Fields("EWayBillNo").Value = txtEWayBillNo.Text
    
    rstPaperPOParent.Fields("TransitFormNo").Value = txtTransitFormNumber.Text
    rstPaperPOParent.Fields("GSTRefNo").Value = txtRefNo.Text

    
    rstPaperPOParent.Fields("Reams_Bundle").Value = Format(Val(MhRealInput12.Text), "0.00")
    rstPaperPOParent.Fields("Bundles").Value = Format(Val(MhRealInput13.Text), "0")
    rstPaperPOParent.Fields("Cartage_Bundle").Value = Format(Val(MhRealInput11.Text), "0.00")
    rstPaperPOParent.Fields("Cartage").Value = Format(Val(MhRealInput4.Text), "0.00")
    
    rstPaperPOParent.Fields("CGST%").Value = Format(Val(MhRealInput9.Text), "0.00")
    rstPaperPOParent.Fields("CGST").Value = Format(Val(MhRealInput6.Text), "0.00")
    rstPaperPOParent.Fields("SGST%").Value = Format(Val(MhRealInput10.Text), "0.00")
    rstPaperPOParent.Fields("SGST").Value = Format(Val(MhRealInput7.Text), "0.00")
    rstPaperPOParent.Fields("BillAmount").Value = Format(Val(MhRealInput8.Text), "0.00")
    If Not IsDate(MhDateInput2.Text) Then rstPaperPOParent.Fields("EWayBillDate").Value = Null Else rstPaperPOParent.Fields("EWayBillDate").Value = GetDate(MhDateInput2.Text)
    If IsNull(rstPaperPOParent.Fields("ComputerName").Value) Then rstPaperPOParent.Fields("ComputerName").Value = Left(lpBuff, (InStr(1, lpBuff, vbNullChar)) - 1)
    rstPaperPOParent.Fields("PrintStatus").Value = "N"
    
    If Combo1.Text = "Delhi" Then
       rstPaperPOParent.Fields("DType").Value = 1
    Else
       rstPaperPOParent.Fields("DType").Value = 2
    End If

End Sub
Private Sub AddToList()
    On Error Resume Next
    
    rstPaperPOList.MoveFirst
    rstPaperPOList.Find "[Code] = '" & rstPaperPOParent.Fields("Code").Value & "'"
    If rstPaperPOList.EOF Then rstPaperPOList.AddNew
    rstPaperPOList.Fields("Code").Value = rstPaperPOParent.Fields("Code").Value
    rstPaperPOList.Fields("Name").Value = Pad(rstPaperPOParent.Fields("Name").Value, Space(1), 10, "L")
    rstPaperPOList.Fields("Date").Value = rstPaperPOParent.Fields("Date").Value

'    rstSupplierList.MoveFirst
'    rstSupplierList.Find "[Code] = '" & rstPaperPOParent.Fields("Supplier").Value & "'"
'    rstPaperPOList.Fields("SupplierName").Value = Trim(rstSupplierList.Fields("Col0").Value)
'    rstPaperPOList.Fields("BillAmount").Value = rstPaperPOParent.Fields("BillAmount").Value
'
    rstPaperPOList.Update
    rstPaperPOList.Sort = SortOrder & " Asc"
    rstPaperPOList.Find "[Code] = '" & rstPaperPOParent.Fields("Code").Value & "'"
    
    End Sub
Private Function CheckMandatoryFields() As Boolean
    
    If CheckEmpty(Text2.Text, False) Then
        DisplayError ("Order No. cannot be blank")
        Text2.SetFocus
        CheckMandatoryFields = True: Exit Function
    ElseIf CheckEmpty(Text3.Text, False) Then
        Text3.SetFocus
        CheckMandatoryFields = True: Exit Function
'    ElseIf Not CheckExists(Text3, "Col0", rstSupplierList, SupplierCode) Then
'        Text3.SetFocus
'        CheckMandatoryFields = True: Exit Function
'    ElseIf CheckDuplicate(CxnPaperPurchaseOrder, "PaperPOParent", "Code", "[Name]+OrderType", Trim(Text2.Text) & OrderType, rstPaperPOParent.Fields("Code").Value, False) Then
'        Text2.SetFocus
'        CheckMandatoryFields = True: Exit Function
  
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
Private Sub LoadPaperList(ByVal strOrderCode As String)
    Dim i As Integer
    On Error GoTo ErrorHandler
    If rstPaperPOChild.State = adStateOpen Then rstPaperPOChild.Close
    rstPaperPOChild.Open "SELECT Paper As PaperCode,M.Name As PaperName,QuantityOther,M.[Weight/Ream],QuantityKg,[RateKg],Amount FROM DeliveryChallanChild T INNER JOIN PaperMaster M ON T.Paper=M.Code WHERE T.Code='" & strOrderCode & "' ORDER BY M.Name", CxnPaperPurchaseOrder, adOpenKeyset, adLockOptimistic
    rstPaperPOChild.ActiveConnection = Nothing
    If rstPaperPOChild.RecordCount > 0 Then rstPaperPOChild.MoveFirst
    i = 0
    
    Do While Not rstPaperPOChild.EOF
        i = i + 1
        With fpSpread1
            .SetText 1, i, rstPaperPOChild.Fields("PaperName").Value
            .SetText 2, i, Val(rstPaperPOChild.Fields("QuantityOther").Value)
            .SetText 3, i, Val(rstPaperPOChild.Fields("Weight/Ream").Value)
            .SetText 4, i, Val(rstPaperPOChild.Fields("QuantityKg").Value)
            .SetText 5, i, Val(rstPaperPOChild.Fields("RateKg").Value)
            .SetText 6, i, Val(rstPaperPOChild.Fields("Amount").Value)
            .SetText 7, i, rstPaperPOChild.Fields("PaperCode").Value
            
        End With
        
        rstPaperPOChild.MoveNext
    Loop
    Exit Sub
ErrorHandler:
    DisplayError ("Failed to Load Paper List")
End Sub

Private Sub LoadPaperFromPO(ByVal strOrderCode As String)
    
    Dim rstPaperChilld As New ADODB.Recordset
        
    Dim i As Integer
    On Error GoTo ErrorHandler
    If rstPaperChilld.State = adStateOpen Then rstPaperChilld.Close
    
    
    'Dim aaa As String
    'aaa = "SELECT Paper As PaperCode,M.Name As PaperName,QuantityOther,M.[Weight/Ream],QuantityKg,[Rate/Kg],Amount FROM PaperPOChild T INNER JOIN PaperMaster M ON T.Paper=M.Code WHERE T.Code='" & strOrderCode & "' ORDER BY M.Name"
    
    rstPaperChilld.Open "SELECT Paper As PaperCode,M.Name As PaperName,QuantityOther,M.[Weight/Ream],QuantityKg,[Rate/Kg],Amount FROM PaperPOChild T INNER JOIN PaperMaster M ON T.Paper=M.Code WHERE T.Code='" & strOrderCode & "' ORDER BY M.Name", CxnPaperPurchaseOrder, adOpenKeyset, adLockOptimistic
    
    
    rstPaperChilld.ActiveConnection = Nothing
    
    If rstPaperChilld.RecordCount > 0 Then rstPaperChilld.MoveFirst
    i = 0
    Do While Not rstPaperChilld.EOF
        i = i + 1
        With fpSpread1
            .SetText 1, i, rstPaperChilld.Fields("PaperName").Value
            .SetText 2, i, Val(rstPaperChilld.Fields("QuantityOther").Value)
            .SetText 3, i, Val(rstPaperChilld.Fields("Weight/Ream").Value)
            .SetText 4, i, Val(rstPaperChilld.Fields("QuantityKg").Value)
            .SetText 5, i, Val(rstPaperChilld.Fields("Rate/Kg").Value)
            .SetText 6, i, Val(rstPaperChilld.Fields("Amount").Value)
            .SetText 7, i, rstPaperChilld.Fields("PaperCode").Value
        End With
        rstPaperChilld.MoveNext
    Loop
    Exit Sub
ErrorHandler:
    DisplayError ("Failed to Load Paper List")
End Sub
Private Function UpdatePaperList(ByVal ActionType As String) As Boolean
    Dim CellVal(1 To 17) As Variant, Sheets As Long
    On Error GoTo ErrorHandler
    UpdatePaperList = True
    If ActionType = "D" And (Not blnRecordExist) Then Exit Function
    If ActionType = "D" Then
        CxnPaperPurchaseOrder.Execute "DELETE FROM DeliveryChallanChild WHERE Code='" & rstPaperPOParent.Fields("Code").Value & "'"
      
    ElseIf ActionType = "I1" Then
        With fpSpread1
            .GetText 2, .ActiveRow, CellVal(1)  'Quantity
            .GetText 4, .ActiveRow, CellVal(2)  'Weight
            .GetText 5, .ActiveRow, CellVal(3)  'Rate
            .GetText 6, .ActiveRow, CellVal(4)  'Amount
            .GetText 7, .ActiveRow, CellVal(5)  'Paper
        End With
        Sheets = Int(Val(CellVal(1))) * 500 + (Val(CellVal(1)) - Int(Val(CellVal(1)))) * 1000
        CxnPaperPurchaseOrder.Execute "INSERT INTO DeliveryChallanChild VALUES ('" & rstPaperPOParent.Fields("Code").Value & "','" & CellVal(5) & "'," & Val(CellVal(1)) & "," & Sheets & "," & Val(CellVal(2)) & "," & Val(CellVal(3)) & "," & Val(CellVal(4)) & ")"
    End If
    Exit Function
ErrorHandler:
    UpdatePaperList = False
End Function
Public Sub FilterRecord(ByVal SrchFor As String, ByVal SrchText As String)
    If SrchFor = "Supplier" Then rstPaperPOList.Filter = "[SupplierName] Like '%" & SrchText & "%'"
End Sub
Private Sub DisplayMenu(ByVal OutputType As String)
   
    Dim menusel As String
    If rstPaperPOList.RecordCount = 0 Then Exit Sub
    menusel = DisplayPopupMenu(Me.hwnd, 3)
    Select Case menusel
        Case 1
            Call PrintPaperDeliveryChallan(rstPaperPOList.Fields("Code").Value, OrderType, "", OutputType, 1)
        Case 2
            Call PrintPaperDeliveryChallan(rstPaperPOList.Fields("Code").Value, OrderType, "", OutputType, 2)
        Case 3
            Call PrintPaperDeliveryChallan(rstPaperPOList.Fields("Code").Value, OrderType, "", OutputType, 3)
    End Select
    If Not (rstPaperPOList.EOF Or rstPaperPOList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    Text1.SetFocus
End Sub
Private Sub LockFields(ByVal bVal As Boolean)
    Dim O As Object
    For Each O In Me
        If TypeName(O) = "TextBox" Then
            O.Locked = bVal
        ElseIf TypeName(O) = "TDBNumber" Then
            O.ReadOnly = bVal
        ElseIf TypeName(O) = "fpSpread" Then
            O.Enabled = Not bVal
        End If
    Next
End Sub

Private Sub CalculateTotal(ByVal strType As String)
    Dim Qty01 As Variant, Qty02 As Variant, Amt As Variant, TCAmt As Variant
    Dim i As Integer
    Dim Qty As Long
    If strType = "G" Then   'Calculate Cartage & VAT
        MhRealInput1.Value = 0: MhRealInput2.Value = 0: MhRealInput3.Value = 0
        Qty = 0
        With fpSpread1
            For i = 1 To .DataRowCnt
                .GetText 2, i, Qty01: .GetText 4, i, Qty02: .GetText 6, i, Amt
                Qty = Qty + Int(Val(Qty01)) * 500 + (Val(Qty01) - Int(Val(Qty01))) * 1000
                MhRealInput2.Value = Val(MhRealInput2.Text) + Qty02
                MhRealInput3.Value = Val(MhRealInput3.Text) + Amt
            Next
            MhRealInput1.Value = Int(Qty / 500) + (Qty Mod 500) / 1000
        End With
    End If
End Sub
Public Sub PrintPaperDeliveryChallan(ByVal OrderCode As String, ByVal OrderType As String, Optional ByVal Note As String, Optional ByVal OutputType As String, Optional ByVal VchType As Integer)
    Dim rstCompanyMaster As New ADODB.Recordset, rstPurchaseOrder As New ADODB.Recordset, rstPurchaseOrderChild As New ADODB.Recordset, Prefix As String
    Dim FileName As String
    Dim strQry As String
    Dim oOutlookMsg As Outlook.MailItem, RecordAffected As Integer
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    Prefix = IIf(OrderType = "1", "PB", "PT") & "/" & Right(Year(FinancialYearFrom), 2) + "-" + Right(Year(FinancialYearTo), 2) & "/"
    rstCompanyMaster.Open "SELECT PrintName,Address1,Address2,Address3,Address4,Phone,Fax,eMail FROM CompanyMaster", CxnDatabase, adOpenKeyset, adLockReadOnly
    strQry = "SELECT '" & Prefix & "'+TRIM(P.Name) As ChallanNo,[P.Date] As ChallanDate,(Select TRIM(Name) From PaperPOParent Where Code=P.PO ) As PO,(Select TRIM(Date) From PaperPOParent Where Code=P.PO ) As PODate,TRIM(P.ConsigneeName) As SupplierName,TRIM(P.ConsigneeAddress) As Address,P.ConsigneeGSTIN As GSTIN,[P.CGST%] AS [CGST%],P.CGST,[P.SGST%] As [SGST%],P.SGST,P.Cartage_Bundle,P.Cartage,P.Reams_Bundle,P.Bundles,P.BillAmount,P.Remarks,TRIM(M2.PrintName) As PaperName,C.QuantityOther,C.QuantityKg,C.RateKg," & _
             "C.Amount,M2.HSNCode,P.CarrierName,P.CarrierAddress,P.VehicleNo,P.DestinationGoods,P.DestinationAddress,P.EWayBillNo,P.EWayBillDate,P.TransitFormNo,P.GSTRefNo,P.DType,(Select Email From AccountMaster Where Code=(Select Supplier From PaperPOParent Where Code=P.PO)) As EmailID FROM ((DeliveryChallanParent P LEFT JOIN DeliveryChallanChild C ON P.Code=C.Code) LEFT JOIN AccountMaster M1 ON M1.Code=P.Consignee) LEFT JOIN PaperMaster M2 ON M2.Code=C.Paper WHERE P.Code='" & OrderCode & "' ORDER BY M2.PrintName"
    rstPurchaseOrder.Open strQry, CxnDatabase, adOpenKeyset, adLockOptimistic
   
    rstPurchaseOrderChild.Open "SELECT '" & Prefix & "'+TRIM(P.Name) As OrderNo,[Date] As OrderDate,TRIM(M3.PrintName) As Godown,TRIM(M2.PrintName) As PaperName,TRIM(M1.PrintName) As PrinterName,'' As RefNo,QuantityOther As Quantity,Tat,'' As Remarks,M1.Address1 As PrinterAdd1,M1.Address2 As PrinterAdd2,M1.Address3 As PrinterAdd3,M1.Address4 As PrinterAdd4,TRIM(M1.eMail) As PrinterMail FROM (((PaperPOParent P INNER JOIN PaperIOChild C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON C.Account=M1.Code) INNER JOIN PaperMaster M2 ON C.Paper=M2.Code) INNER JOIN AccountMaster M3 ON P.Supplier=M3.Code WHERE P.Code='" & OrderCode & "' ORDER BY M2.PrintName", CxnDatabase, adOpenKeyset, adLockOptimistic
    Screen.MousePointer = vbNormal
    rstPurchaseOrder.ActiveConnection = Nothing: rstPurchaseOrderChild.ActiveConnection = Nothing
    'If VchType = 1 Then
        rptPaperDeliveryChallan.Text1.SetText IIf(OrderType = "1", "Book", "Title") & " Paper Delivery Challan" & IIf(VchType = 1, "(Original)", IIf(VchType = 2, "(Duplicate)", "(Triplicate)"))
        rptPaperDeliveryChallan.Text8.SetText Trim(COMPANY_CIN) 'Add here company CIN No
        rptPaperDeliveryChallan.Text16.SetText Trim(COMPANY_PAN) 'Add here company PAN No
        rptPaperDeliveryChallan.Text3.SetText Trim(rstPurchaseOrder.Fields("Address").Value)
        rptPaperDeliveryChallan.Text32.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
        rptPaperDeliveryChallan.Text33.SetText Trim(rstCompanyMaster.Fields("Address1").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address2").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address3").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address4").Value)
        rptPaperDeliveryChallan.Text34.SetText Trim(COMPANY_GSTIN) 'Add here company GST IN
        rptPaperDeliveryChallan.Text20.SetText "Add : GST @" + Format(rstPurchaseOrder.Fields("VAT%").Value, "0.00") + "%"
        rptPaperDeliveryChallan.Text28.SetText " (" & Trim(NumberToWords(rstPurchaseOrder.Fields("BillAmount").Value, True)) & ")"
        rptPaperDeliveryChallan.Text27.SetText "for " & Trim(rstPurchaseOrder.Fields("SupplierName").Value)
        rptPaperDeliveryChallan.Text9.SetText "for " & Trim(rstCompanyMaster.Fields("PrintName").Value)
        rptPaperDeliveryChallan.Database.SetDataSource rstPurchaseOrder, 3, 1
       
        EMailID = ""
        EMailID = Replace(rstPurchaseOrder.Fields("EmailID").Value, Chr(39), "") 'Replace single qote with space
        Attachment = Trim(rstPurchaseOrder.Fields("ChallanNo").Value)
        Attachment = Mid(Attachment, InStr(4, Attachment, "/") + 1)
        Message = "Dear Sir,<Br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Please find attached herewith PO #" & Trim(rstPurchaseOrder.Fields("ChallanNo").Value) & " for doing the needful at your end. An early execution of the order will be highly appreciated.<Br>Kindly acknowledge the receipt of mail and confirm the date of execution of order.<Br><Br>" & IIf(Note = "", "", "<b><u>Note : " & Note & "</b></u><Br><Br>") & "Thanks & Regards<Br>Production Department<Br>" & Trim(rstCompanyMaster.Fields("PrintName").Value) & "<Br>Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & "<Br>E-Mail : <a HRef='mailto:" & Trim(rstCompanyMaster.Fields("EMail").Value) & "'>" & Trim(rstCompanyMaster.Fields("EMail").Value) & "</a>"
         
      
        If OutputType = "S" Then
            FrmReportViewer.EMailID = Trim(EMailID)
            FrmReportViewer.Subject = IIf(OrderType = "1", "Book", "Title") & " Delivery Challan #" & Trim(rstPurchaseOrder.Fields("ChallanNo").Value)
            FrmReportViewer.Attachment = Attachment
            FrmReportViewer.Message = Message
            Set FrmReportViewer.Report = rptPaperDeliveryChallan
            FrmReportViewer.Show vbModal
         ElseIf OutputType = "P" Then
            rptPaperDeliveryChallan.PrintOut False    'Print Report Without Prompt
         Else
            rptPaperDeliveryChallan.ExportOptions.FormatType = crEFTPortableDocFormat    ' Set the Export Format As .Pdf
            rptPaperDeliveryChallan.ExportOptions.DestinationType = crEDTDiskFile
            rptPaperDeliveryChallan.ExportOptions.DiskFileName = App.Path & "\Report\" & Attachment & ".Pdf"
            rptPaperDeliveryChallan.Export False
            rstPurchaseOrder.MoveFirst
            Set oOutlookMsg = oOutlook.CreateItem(olMailItem)
            With oOutlookMsg
                .To = Trim(EMailID)
                .Subject = IIf(OrderType = "1", "Book", "Title") & " Delivery Challan #" & Trim(rstPurchaseOrder.Fields("ChallanNo").Value)
                .HTMLBody = "<Font Face='Calibri' Size='3'>" & Message & "</a>" & "</Font>"
                .Attachments.Add (App.Path & "\Report\" & Attachment & ".Pdf")
                .Importance = olImportanceHigh
                .ReadReceiptRequested = True
                If CheckEmpty(.To, False) Then .Display Else .Send
            End With
            Set oOutlookMsg = Nothing
         End If
         Set rptPaperDeliveryChallan = Nothing
    'End If
    Call CloseRecordset(rstPurchaseOrder): Call CloseRecordset(rstCompanyMaster): Call CloseRecordset(rstPurchaseOrderChild)
    On Error GoTo 0
End Sub




