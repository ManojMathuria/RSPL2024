VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0F1F1508-C40A-101B-AD04-00AA00575482}#1.0#0"; "mhrinp32.ocx"
Begin VB.Form FrmOutsourceItemIssueOrder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Outsource Item Issue Order"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8790
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "OutsourceItemIssueOrder.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4875
   ScaleWidth      =   8790
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   4875
      Left            =   15
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   8700
      _Version        =   65536
      _ExtentX        =   15346
      _ExtentY        =   8599
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
      Picture         =   "OutsourceItemIssueOrder.frx":0442
      Begin TabDlg.SSTab SSTab1 
         Height          =   4635
         Left            =   120
         TabIndex        =   11
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
         TabPicture(0)   =   "OutsourceItemIssueOrder.frx":045E
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "DataGrid1"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Text1"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "&Details"
         TabPicture(1)   =   "OutsourceItemIssueOrder.frx":047A
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
            TabIndex        =   13
            Top             =   4160
            Width           =   7760
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   3640
            Left            =   120
            TabIndex        =   12
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
                  ColumnWidth     =   5564.977
               EndProperty
            EndProperty
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   4005
            Left            =   -74880
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   480
            Width           =   8235
            _Version        =   65536
            _ExtentX        =   14526
            _ExtentY        =   7064
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
            Picture         =   "OutsourceItemIssueOrder.frx":0496
            Begin TDBNumber6Ctl.TDBNumber MhRealInput3 
               Height          =   330
               Left            =   6885
               TabIndex        =   8
               Top             =   1710
               Visible         =   0   'False
               Width           =   1005
               _Version        =   65536
               _ExtentX        =   1773
               _ExtentY        =   582
               Calculator      =   "OutsourceItemIssueOrder.frx":04B2
               Caption         =   "OutsourceItemIssueOrder.frx":04D2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "OutsourceItemIssueOrder.frx":053E
               Keys            =   "OutsourceItemIssueOrder.frx":055C
               Spin            =   "OutsourceItemIssueOrder.frx":05A6
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "###########0"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "###########0"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   999999999999
               MinValue        =   -999999999999
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   1965490181
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin MhinrelLib.MhRealInput MhRealInput4 
               Height          =   255
               Left            =   6890
               TabIndex        =   21
               TabStop         =   0   'False
               Top             =   3625
               Width           =   1005
               _Version        =   65536
               _ExtentX        =   1773
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
               MaxReal         =   9999999
               MinReal         =   0
               ReadOnly        =   -1  'True
               SpinChangeReal  =   0
               CaretColor      =   -2147483642
               DecimalPlaces   =   0
               VAlignment      =   2
            End
            Begin VB.TextBox Text9 
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
               Left            =   5730
               MaxLength       =   10
               TabIndex        =   7
               Top             =   1710
               Visible         =   0   'False
               Width           =   1165
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
               Left            =   3120
               MaxLength       =   40
               TabIndex        =   6
               Top             =   1710
               Visible         =   0   'False
               Width           =   2625
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel19 
               Height          =   255
               Left            =   120
               TabIndex        =   20
               Top             =   3625
               Width           =   8010
               _Version        =   65536
               _ExtentX        =   14129
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
               Picture         =   "OutsourceItemIssueOrder.frx":05CE
               Picture         =   "OutsourceItemIssueOrder.frx":05EA
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
               Left            =   435
               MaxLength       =   40
               TabIndex        =   5
               Top             =   1710
               Visible         =   0   'False
               Width           =   2700
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
               Width           =   1650
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
               MaxLength       =   139
               TabIndex        =   3
               Top             =   950
               Width           =   6690
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
               MaxLength       =   40
               TabIndex        =   2
               Top             =   630
               Width           =   6690
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
               Height          =   330
               Left            =   120
               TabIndex        =   16
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
               Picture         =   "OutsourceItemIssueOrder.frx":0606
               Picture         =   "OutsourceItemIssueOrder.frx":0622
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Left            =   5955
               TabIndex        =   17
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
               Picture         =   "OutsourceItemIssueOrder.frx":063E
               Picture         =   "OutsourceItemIssueOrder.frx":065A
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   330
               Left            =   120
               TabIndex        =   18
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
               Caption         =   " Supplier Name"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "OutsourceItemIssueOrder.frx":0676
               Picture         =   "OutsourceItemIssueOrder.frx":0692
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
               Height          =   330
               Left            =   120
               TabIndex        =   19
               Top             =   945
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
               Caption         =   " Remarks"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "OutsourceItemIssueOrder.frx":06AE
               Picture         =   "OutsourceItemIssueOrder.frx":06CA
            End
            Begin MSDataGridLib.DataGrid DataGrid2 
               Height          =   2415
               Left            =   120
               TabIndex        =   4
               Top             =   1470
               Width           =   8010
               _ExtentX        =   14129
               _ExtentY        =   4260
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
               ColumnCount     =   4
               BeginProperty Column00 
                  DataField       =   "OutsourceItemName"
                  Caption         =   "Outsource Item Name"
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
                  DataField       =   "GodownName"
                  Caption         =   "Godown Name"
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
               BeginProperty Column02 
                  DataField       =   "RefNo"
                  Caption         =   "Ref.No."
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
                     ColumnWidth     =   2684.977
                  EndProperty
                  BeginProperty Column01 
                     Locked          =   -1  'True
                     ColumnWidth     =   2610.142
                  EndProperty
                  BeginProperty Column02 
                     Alignment       =   1
                     Locked          =   -1  'True
                     ColumnWidth     =   1154.835
                  EndProperty
                  BeginProperty Column03 
                     Alignment       =   1
                     ColumnAllowSizing=   -1  'True
                     Locked          =   -1  'True
                     ColumnWidth     =   989.858
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
               Calendar        =   "OutsourceItemIssueOrder.frx":06E6
               Caption         =   "OutsourceItemIssueOrder.frx":07FE
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "OutsourceItemIssueOrder.frx":086A
               Keys            =   "OutsourceItemIssueOrder.frx":0888
               Spin            =   "OutsourceItemIssueOrder.frx":08E6
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
            Begin VB.Line Line1 
               X1              =   0
               X2              =   8280
               Y1              =   525
               Y2              =   525
            End
            Begin VB.Line Line2 
               X1              =   0
               X2              =   8280
               Y1              =   1365
               Y2              =   1365
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
            TabIndex        =   14
            Top             =   4160
            Width           =   495
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   8790
      _ExtentX        =   15505
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
Attribute VB_Name = "FrmOutsourceItemIssueOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CxnMaterialIssueOrder As New ADODB.Connection
Dim rstCompanyMaster As New ADODB.Recordset
Dim rstMaterialIOList As New ADODB.Recordset
Dim rstMaterialIOParent As New ADODB.Recordset
Dim rstMaterialIOChild As ADODB.Recordset
Dim rstAccountList As New ADODB.Recordset
Dim rstSupplierList As New ADODB.Recordset
Dim rstRefList As New ADODB.Recordset
Dim rstOutsourceItemList As New ADODB.Recordset
Dim AccountCode As String
Dim SupplierCode As String
Dim RefCode As String
Dim OutsourceItemCode As String
Dim SortOrder As String
Dim PrevStr As String
Dim dblBookMark As Double
Dim blnRecordExist As Boolean
Dim OutputTo As String
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    
    CenterForm Me
    BusySystemIndicator True
    CxnMaterialIssueOrder.CursorLocation = adUseClient
    CxnMaterialIssueOrder.Open CxnDatabase.ConnectionString
    rstCompanyMaster.Open "Select PrintName, Address1, Address2, Address3, Address4, Phone, Fax, EMail, Website From CompanyMaster", CxnMaterialIssueOrder, adOpenKeyset, adLockReadOnly
    
    rstOutsourceItemList.Open "Select Name As Col0,Code From OutsourceItemMaster Order By Name", CxnMaterialIssueOrder, adOpenKeyset, adLockReadOnly
    
    rstAccountList.Open "Select Name As Col0,Code From AccountMaster Where Type In ('08','09') Order By Name", CxnMaterialIssueOrder, adOpenKeyset, adLockReadOnly
    
    rstSupplierList.Open "Select Name As Col0, Code From AccountMaster Where Type = '01' Order by Name", CxnMaterialIssueOrder, adOpenKeyset, adLockReadOnly
    rstMaterialIOList.Open "Select T.Code,T.Name,T.Date,M.Name As SupplierName From MaterialIOParent T, AccountMaster M Where T.Source = M.Code And T.Type='1' Order By T.Name", CxnMaterialIssueOrder, adOpenKeyset, adLockOptimistic
    rstMaterialIOParent.CursorLocation = adUseClient
    Set rstMaterialIOChild = New ADODB.Recordset
    rstMaterialIOList.Filter = adFilterNone
    If rstMaterialIOList.RecordCount > 0 Then rstMaterialIOList.MoveLast
    Set DataGrid1.DataSource = rstMaterialIOList
    BusySystemIndicator False
    SSTab1.Tab = 0
    SortOrder = "Name"
    If Not (rstMaterialIOList.EOF Or rstMaterialIOList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    rstMaterialIOList.ActiveConnection = Nothing
    rstOutsourceItemList.ActiveConnection = Nothing
    rstAccountList.ActiveConnection = Nothing
    rstSupplierList.ActiveConnection = Nothing
    SetButtonsForNoRecord
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    Unload Me
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
            Unload Me
        Else
            If Toolbar1.Buttons.Item(1).Enabled Then
                SSTab1.Tab = 0
            Else
                If Me.ActiveControl.Name <> "Text5" And Me.ActiveControl.Name <> "Text8" And Me.ActiveControl.Name <> "Text9" And Me.ActiveControl.Name <> "MhRealInput3" Then
                    If MsgBox("Are you sure to Quit?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Quit !") <> vbYes Then
                        Me.ActiveControl.SetFocus
                    Else
                        Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
                    End If
                End If
            End If
            If Me.ActiveControl.Name <> "Text5" And Me.ActiveControl.Name <> "Text8" And Me.ActiveControl.Name <> "Text9" And Me.ActiveControl.Name <> "MhRealInput3" Then
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
        If Me.ActiveControl.Name <> "Text5" And Me.ActiveControl.Name <> "Text8" And Me.ActiveControl.Name <> "Text9" And Me.ActiveControl.Name <> "MhRealInput3" Then
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
            If Me.ActiveControl.Name <> "MhRealInput3" Then
                SendKeys "{TAB}"
            End If
        End If
        If Me.ActiveControl.Name <> "MhRealInput3" Then
            KeyCode = 0
        End If
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Toolbar1.Buttons.Item(4).Enabled Then
        Call Form_KeyDown(vbKeyEscape, 0)
        Cancel = 1
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CloseRecordset(rstCompanyMaster)
    Call CloseRecordset(rstMaterialIOList)
    Call CloseRecordset(rstMaterialIOParent)
    Call CloseRecordset(rstMaterialIOChild)
    Call CloseRecordset(rstOutsourceItemList)
    Call CloseRecordset(rstAccountList)
    Call CloseRecordset(rstSupplierList)
    Call CloseRecordset(rstRefList)
    Call CloseConnection(CxnMaterialIssueOrder)
    ShowProgressInStatusBar False
    DisableChildMenu
End Sub
Private Sub Text1_Change()
    If rstMaterialIOList.RecordCount = 0 Then Exit Sub
    rstMaterialIOList.MoveFirst
    If Text1.Text <> "" Then
        If SortOrder = "Name" Then
           rstMaterialIOList.Find "[" & SortOrder & "] Like '%" & FixQuote(Text1.Text) & "%'"
        Else
           rstMaterialIOList.Find "[" & SortOrder & "] Like '" & FixQuote(Text1.Text) & "%'"
        End If
        If rstMaterialIOList.EOF Then
            rstMaterialIOList.MoveFirst
            If PrevStr <> "" And Len(Text1.Text) > 1 Then
                If dblBookMark <> 0 Then
                    rstMaterialIOList.Bookmark = dblBookMark
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
    If Not (rstMaterialIOList.EOF Or rstMaterialIOList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim KeyProcessed As Boolean
    
    If rstMaterialIOList.RecordCount = 0 Then Exit Sub
    If Shift = 0 And KeyCode = vbKeyUp Then
        With rstMaterialIOList
            .MovePrevious
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyBack Then
        With rstMaterialIOList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyDown Then
        With rstMaterialIOList
            .MoveNext
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageUp Then
        With rstMaterialIOList
            .Move (-1) * (DataGrid1.VisibleRows - 1)
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageUp Then
        With rstMaterialIOList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageDown Then
        With rstMaterialIOList
            .Move DataGrid1.VisibleRows - 1
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageDown Then
        With rstMaterialIOList
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
            If Not (rstMaterialIOList.EOF Or rstMaterialIOList.BOF) Then
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
        If rstMaterialIOParent.State = adStateOpen Then
           rstMaterialIOParent.Close
        End If
        rstMaterialIOParent.Open "Select * From MaterialIOParent Where Code = ''", CxnMaterialIssueOrder, adOpenKeyset, adLockOptimistic
        ClearFields ("P")
        ClearFields ("C")
        Call LoadOutsourceItemList("")
        If rstMaterialIOChild.State = adStateClosed Then
            SSTab1.Tab = 0
            Exit Sub
        End If
        If AddRecord(rstMaterialIOParent) Then
            Text2.Text = GenerateCode(CxnMaterialIssueOrder, "Select Max(Val(Name)) From MaterialIOParent Where Type='1'", 10, Space(1))
            MhDateInput1.Text = Format(Date, "dd-MM-yyyy")
            Call SetButtons(False)
            SSTab1.Tab = 1
            Text2.SetFocus
            blnRecordExist = False
            CxnMaterialIssueOrder.BeginTrans
        End If
    ElseIf Button.Index = 2 Then
        If rstMaterialIOList.RecordCount = 0 Then Exit Sub
        SSTab1.Tab = 1
        EditRecord
    ElseIf Button.Index = 3 Then
        If rstMaterialIOList.RecordCount = 0 Then Exit Sub
        If AllowTransactionsDeletion = 0 Then
            Call DisplayError("You don't have the rights to Delete this Voucher")
            Exit Sub
        End If
        SSTab1.Tab = 1
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            On Error Resume Next
            MdiMainMenu.MousePointer = vbHourglass
            CxnMaterialIssueOrder.Execute "Delete From MaterialIOParent Where Code = '" & rstMaterialIOList.Fields("Code").Value & "'"
            MdiMainMenu.MousePointer = vbNormal
            If Err.Number = 0 Then
                rstMaterialIOList.Delete
                rstMaterialIOList.MoveNext
                If rstMaterialIOList.RecordCount > 0 And rstMaterialIOList.EOF Then
                    rstMaterialIOList.MoveLast
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
        If UpdateRecord(rstMaterialIOParent) Then
            If UpdateOutsourceItemList("D") Then
                 UpdateFlag = 1
                 If rstMaterialIOChild.RecordCount <> 0 Then
                      rstMaterialIOChild.MoveFirst
                      Do While Not rstMaterialIOChild.EOF
                          If Val(rstMaterialIOChild.Fields("Quantity").Value) <> 0 Then
                               If Not UpdateOutsourceItemList("U") Then
                                    UpdateFlag = 0
                                    Exit Do
                                End If
                          End If
                          rstMaterialIOChild.MoveNext
                      Loop
                 End If
            End If
        End If
        If UpdateFlag Then
            AddToList
            CxnMaterialIssueOrder.CommitTrans
            If rstMaterialIOParent.State = adStateOpen Then
                rstMaterialIOParent.Close
            End If
            rstMaterialIOParent.CursorLocation = adUseClient
            Call SetButtons(True)
            SSTab1.Tab = 0
            ShowProgressInStatusBar True
            Timer1.Enabled = True
        Else
            DisplayError ("Failed to save the record")
            Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
        End If
    ElseIf Button.Index = 5 Then
        If CancelRecordUpdate(rstMaterialIOParent) Then
            CxnMaterialIssueOrder.RollbackTrans
            If rstMaterialIOParent.State = adStateOpen Then
                rstMaterialIOParent.Close
            End If
            rstMaterialIOParent.CursorLocation = adUseClient
            Call SetButtons(True)
            SetButtonsForNoRecord
            SSTab1.Tab = 0
        End If
    ElseIf Button.Index = 6 Then
        SSTab1.Tab = 0
        Set DataGrid1.DataSource = Nothing
        rstMaterialIOList.ActiveConnection = CxnMaterialIssueOrder
        Do While Not RefreshRecord(rstMaterialIOList)
        Loop
        Set DataGrid1.DataSource = rstMaterialIOList
        rstMaterialIOList.ActiveConnection = Nothing
        If rstMaterialIOList.RecordCount > 0 Then rstMaterialIOList.MoveLast
        rstAccountList.ActiveConnection = CxnMaterialIssueOrder
        Do While Not RefreshRecord(rstAccountList)
        Loop
        rstAccountList.ActiveConnection = Nothing
        rstSupplierList.ActiveConnection = CxnMaterialIssueOrder
        Do While Not RefreshRecord(rstSupplierList)
        Loop
        rstSupplierList.ActiveConnection = Nothing
        rstOutsourceItemList.ActiveConnection = CxnMaterialIssueOrder
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
        If rstMaterialIOList.RecordCount = 0 Then Exit Sub
        OutputTo = "P"
        PrintOutsourceItemIssueOrder
        HiLiteRecord = True
    ElseIf Button.Index = 10 Then
        If rstMaterialIOList.RecordCount = 0 Then Exit Sub
        OutputTo = "S"
        PrintOutsourceItemIssueOrder
        HiLiteRecord = True
    ElseIf Button.Index = 13 Then
        If rstMaterialIOList.RecordCount > 0 Then rstMaterialIOList.MoveFirst
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 14 Then
        If rstMaterialIOList.RecordCount > 0 Then
            rstMaterialIOList.MovePrevious
            If rstMaterialIOList.BOF Then
                rstMaterialIOList.MoveNext
            End If
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 15 Then
        If rstMaterialIOList.RecordCount > 0 Then
            rstMaterialIOList.MoveNext
            If rstMaterialIOList.EOF Then
                rstMaterialIOList.MovePrevious
            End If
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 16 Then
        If rstMaterialIOList.RecordCount > 0 Then rstMaterialIOList.MoveLast
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 18 Then
        Unload Me
        HiLiteRecord = False
    End If
    If HiLiteRecord Then
        If Not (rstMaterialIOList.EOF Or rstMaterialIOList.BOF) Then
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
          rstMaterialIOList.Sort = "Name Asc"
       End If
    ElseIf ColIndex = 2 Then
       If SortOrder <> "SupplierName" Then
          SortOrder = "SupplierName"
          rstMaterialIOList.Sort = "SupplierName Asc"
       End If
    End If
    DataGrid1.ClearSelCols
    If Not (rstMaterialIOList.EOF Or rstMaterialIOList.BOF) Then
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
    If rstMaterialIOList.RecordCount = 0 Then
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
    If rstMaterialIOParent.EOF Or rstMaterialIOParent.BOF Then Exit Sub
    If CheckEmpty(Text2, True) Then
        Cancel = True
    ElseIf CheckDuplicate(CxnMaterialIssueOrder, "MaterialIOParent", "Code", "[Name]+[Type]", Trim(Text2.Text) & "1", rstMaterialIOParent.Fields("Code").Value, False) Then
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
    If rstMaterialIOChild.RecordCount = 0 Then
        SendKeys "^"
        Call AddRecord(rstMaterialIOChild)
        Call ClearFields("C")
        Call DataGrid2_KeyDown(vbKeyE, vbCtrlMask)
    End If
End Sub
Private Sub ViewRecord()
    ClearFields ("P")
    ClearFields ("C")
    If rstMaterialIOList.EOF Then
        If rstMaterialIOChild.State = adStateOpen Then
            rstMaterialIOChild.Close
        End If
        Exit Sub
    End If
    FindRecord
    LoadFields
End Sub
Private Sub FindRecord()
    If rstMaterialIOParent.State = adStateOpen Then
       rstMaterialIOParent.Close
    End If
    rstMaterialIOParent.Open "Select * From MaterialIOParent Where Code = '" & FixQuote(rstMaterialIOList.Fields("Code").Value) & "'", CxnMaterialIssueOrder, adOpenKeyset, adLockOptimistic
    If rstMaterialIOParent.RecordCount = 0 Then
       Call DisplayError("This Record has been deleted by Another User ! Click Ok To Refresh the Recordset")
       Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
    End If
End Sub
Private Sub ClearFields(ByVal strType As String)
    If strType = "P" Then
        Text2.Text = ""
        Text3.Text = ""
        Text4.Text = ""
        MhDateInput1.Text = Format(Date, "dd-MM-yyyy")
        MhRealInput4.Text = 0#
    ElseIf strType = "C" Then
        Text5.Text = ""
        Text8.Text = ""
        Text9.Text = ""
        MhRealInput3.Text = "0"
    End If
End Sub
Private Sub LoadFields()
    If rstMaterialIOParent.EOF Or rstMaterialIOParent.BOF Then Exit Sub
    Text2.Text = rstMaterialIOParent.Fields("Name").Value
    MhDateInput1.Text = Format(rstMaterialIOParent.Fields("Date").Value, "dd-MM-yyyy")
    SupplierCode = rstMaterialIOParent.Fields("Source").Value
    If rstSupplierList.RecordCount > 0 Then rstSupplierList.MoveFirst
    rstSupplierList.Find "[Code] = '" & SupplierCode & "'"
    If Not rstSupplierList.EOF Then
       Text3.Text = rstSupplierList.Fields("Col0").Value
    End If
    Text4.Text = rstMaterialIOParent.Fields("Remarks").Value
    Call LoadOutsourceItemList(rstMaterialIOParent.Fields("Code").Value)
    If rstMaterialIOChild.State = adStateOpen Then
        CalculateTotal
    End If
End Sub
Private Sub EditRecord()
    On Error GoTo ErrorHandler
    
    If rstMaterialIOParent.RecordCount = 0 Then Exit Sub
    If rstMaterialIOChild.State = adStateClosed Then
        SSTab1.Tab = 0
        Exit Sub
    End If
    If rstMaterialIOParent.State = adStateOpen Then
       rstMaterialIOParent.Close
    End If
    rstMaterialIOParent.CursorLocation = adUseServer
    rstMaterialIOParent.Open "Select * From MaterialIOParent Where Code = '" & FixQuote(rstMaterialIOList.Fields("Code").Value) & "'", CxnMaterialIssueOrder, adOpenKeyset, adLockPessimistic
    MdiMainMenu.MousePointer = vbHourglass
    rstMaterialIOParent.Fields("Printstatus") = "N"
    MdiMainMenu.MousePointer = vbNormal
    AddToList
    Call SetButtons(False)
    SSTab1.TabEnabled(0) = False
    Text2.SetFocus
    blnRecordExist = True
    CxnMaterialIssueOrder.BeginTrans
    Exit Sub
ErrorHandler:
    If Err.Number = -2147467259 Then
       Call DisplayError("Failed to Edit the record")
    End If
    MdiMainMenu.MousePointer = vbNormal
    SSTab1.Tab = 0
End Sub
Private Sub SaveFields()
    If rstMaterialIOParent.EOF Or rstMaterialIOParent.BOF Then Exit Sub
    If Not blnRecordExist Then
        rstMaterialIOParent.Fields("Code").Value = GenerateCode(CxnMaterialIssueOrder, "Select Max(Code) From MaterialIOParent", 6, "0")
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
    
    rstMaterialIOList.MoveFirst
    rstMaterialIOList.Find "[Code] = '" & rstMaterialIOParent.Fields("Code").Value & "'"
    If rstMaterialIOList.EOF Then
       rstMaterialIOList.AddNew
       rstMaterialIOList.Fields("Code").Value = rstMaterialIOParent.Fields("Code").Value
    End If
    rstMaterialIOList.Fields("Name").Value = Pad(rstMaterialIOParent.Fields("Name").Value, Space(1), 10, "L")
    rstMaterialIOList.Fields("Date").Value = rstMaterialIOParent.Fields("Date").Value
    rstSupplierList.MoveFirst
    rstSupplierList.Find "[Code] = '" & rstMaterialIOParent.Fields("Source").Value & "'"
    rstMaterialIOList.Fields("SupplierName").Value = Trim(rstSupplierList.Fields("Col0").Value)
    rstMaterialIOList.Update
    rstMaterialIOList.Sort = SortOrder & " Asc"
    rstMaterialIOList.Find "[Code] = '" & rstMaterialIOParent.Fields("Code").Value & "'"
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
    ElseIf CheckDuplicate(CxnMaterialIssueOrder, "MaterialIOParent", "Code", "[Name]+[Type]", Trim(Text2.Text) & "1", rstMaterialIOParent.Fields("Code").Value, False) Then
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
Private Sub LoadOutsourceItemList(ByVal strOrderCode As String)
    On Error GoTo ErrorHandler
    
    If rstMaterialIOChild.State = adStateOpen Then
       rstMaterialIOChild.Close
    End If
    rstMaterialIOChild.Open "Select C.Item,M1.Name As OutsourceItemName,C.Godown,M2.Name As GodownName,C.Ref,T.Name As RefNo,C.Quantity From OutsourceItemMaster M1,MaterialIOChild C,AccountMaster M2, OutsourceItemPOParent T Where C.Item = M1.Code And C.Godown = M2.Code And C.Ref = T.Code And C.Code = '" & strOrderCode & "' Order By M1.Name", CxnMaterialIssueOrder, adOpenKeyset, adLockOptimistic
    rstMaterialIOChild.ActiveConnection = Nothing
    Set DataGrid2.DataSource = rstMaterialIOChild
    Exit Sub
ErrorHandler:
    DisplayError ("Failed to Load OutsourceItem List")
End Sub
Private Sub DataGrid2_DblClick()
    Call DataGrid2_KeyDown(vbKeyE, vbCtrlMask)
End Sub
Private Sub DataGrid2_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask And KeyCode = vbKeyE Then
        If rstMaterialIOChild.RecordCount = 0 Then
            KeyCode = 0
            Exit Sub
        End If
        If Val(CheckNull(rstMaterialIOChild.Fields("Quantity").Value)) <> 0 Then
            OutsourceItemCode = rstMaterialIOChild.Fields("Item").Value
            Text5.Text = rstMaterialIOChild.Fields("OutsourceItemName").Value
            AccountCode = rstMaterialIOChild.Fields("Godown").Value
            Text8.Text = rstMaterialIOChild.Fields("GodownName").Value
            RefCode = rstMaterialIOChild.Fields("Ref").Value
            Text9.Text = rstMaterialIOChild.Fields("RefNo").Value
            MhRealInput3.Text = Format(Val(rstMaterialIOChild.Fields("Quantity").Value), "0")
        End If
        With DataGrid2
            Text5.Visible = True
            Text5.Move .Left + .Columns(0).Left, .Top + .RowTop(.Row), .Columns(0).Width + 10, .RowHeight + 30
            Text8.Visible = True
            Text8.Move .Left + .Columns(1).Left, .Top + .RowTop(.Row), .Columns(1).Width + 10, .RowHeight + 30
            Text9.Visible = True
            Text9.Move .Left + .Columns(2).Left, .Top + .RowTop(.Row), .Columns(2).Width + 10, .RowHeight + 30
            MhRealInput3.Visible = True
            MhRealInput3.Move .Left + .Columns(3).Left, .Top + .RowTop(.Row), .Columns(3).Width + 10, .RowHeight + 30
        End With
        DataGrid2.Enabled = False
        Text5.SetFocus
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyA Then
        SendKeys "^"
        Call AddRecord(rstMaterialIOChild)
        Call ClearFields("C")
        Call DataGrid2_KeyDown(vbKeyE, vbCtrlMask)
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyD Then
        If rstMaterialIOChild.RecordCount = 0 Then Exit Sub
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            Set DataGrid2.DataSource = Nothing
            rstMaterialIOChild.Delete
            rstMaterialIOChild.MoveNext
            Set DataGrid2.DataSource = rstMaterialIOChild
            CalculateTotal
            DataGrid2.SetFocus
        End If
        If rstMaterialIOChild.RecordCount = 0 Then
            Call ClearFields("C")
        End If
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyS And Toolbar1.Buttons.Item(4).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(4)
    ElseIf Shift = vbShiftMask And KeyCode = vbKeyTab Then
       Text4.SetFocus
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
    If rstOutsourceItemList.RecordCount = 0 Then
        DisplayError ("No Record in OutsourceItem Master")
        Cancel = True
        Exit Sub
    Else
        rstOutsourceItemList.MoveFirst
    End If
    rstOutsourceItemList.Find "[Col0] = '" & RTrim(SearchString) & "'"
    If rstOutsourceItemList.EOF Then
        If Not CheckEmpty(OutsourceItemCode, False) Then
            rstOutsourceItemList.MoveFirst
            rstOutsourceItemList.Find "[Code] = '" & OutsourceItemCode & "'"
            Text5.Text = rstOutsourceItemList.Fields("Col0").Value
        End If
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
    ElseIf (rstMaterialIOChild.Fields("OutsourceItemName").Value <> Text5.Text) Or (CheckEmpty(rstMaterialIOChild.Fields("OutsourceItemName").Value, False)) Then
        If CheckDuplicateEntry Then
            Call DisplayError("Duplicate Entry")
            Text5.SelStart = 0
            Text5.SelLength = Len(Text5.Text)
            Cancel = True
            Exit Sub
        End If
    End If
    OutsourceItemCode = rstOutsourceItemList.Fields("Code").Value
    Call LoadRefList(OutsourceItemCode, SupplierCode, CheckNull(rstMaterialIOParent.Fields("Code").Value))
End Sub
Private Sub Text8_Change()
    If Text8.Text = " " Then
        Text8.Text = "?"
        SendKeys "{TAB}"
    End If
End Sub
Private Sub Text8_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyEscape Then
        MakeTextBoxInvisible (True)
    End If
End Sub
Private Sub Text8_Validate(Cancel As Boolean)
    Dim SearchString As String
    
    SearchString = FixQuote(Text8.Text)
    If rstAccountList.RecordCount = 0 Then
        DisplayError ("No Record in Godown Master")
        Cancel = True
        Exit Sub
    Else
        rstAccountList.MoveFirst
    End If
    rstAccountList.Find "[Col0] = '" & RTrim(SearchString) & "'"
    If rstAccountList.EOF Then
        SelectionType = "S"
        AccountCode = ""
        Call LoadSelectionList(rstAccountList, "List of Godowns...", "Name")
        SearchOrder = 0
        Call DisplaySelectionList(Text8, AccountCode)
        Call CloseForm(FrmSelectionList)
        If CheckEmpty(Text8.Text, False) Then
            Text8.Text = "?"
        End If
        If RTrim(AccountCode) <> "" Then
            SendKeys "{TAB}"
        End If
        Cancel = True
        Exit Sub
    ElseIf (rstMaterialIOChild.Fields("GodownName").Value <> Text8.Text) Or (CheckEmpty(rstMaterialIOChild.Fields("GodownName").Value, False)) Then
        If CheckDuplicateEntry Then
            Call DisplayError("Duplicate Entry")
            Text8.SelStart = 0
            Text8.SelLength = Len(Text8.Text)
            Cancel = True
            Exit Sub
        End If
    End If
    AccountCode = rstAccountList.Fields("Code").Value
End Sub
Private Sub Text9_Change()
    If Text9.Text = " " Then
        Text9.Text = "?"
        SendKeys "{TAB}"
    End If
End Sub
Private Sub Text9_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyEscape Then
        MakeTextBoxInvisible (True)
    End If
End Sub
Private Sub Text9_Validate(Cancel As Boolean)
    
    Dim SearchString As String
    SearchString = FixQuote(Text9.Text)
    If rstRefList.RecordCount = 0 Then
        DisplayError ("No Pending Order")
        Cancel = True
        Exit Sub
    Else
        rstRefList.MoveFirst
    End If
    rstRefList.MoveFirst
    
    rstRefList.Find "[Name] = '" & Pad(Trim(SearchString), Space(1), 10, "L") & "'"
    
    If rstRefList.EOF Then
        SelectionType = "S"
        RefCode = ""
        Call LoadSelectionList(rstRefList, "List of Pending Orders...", "Order No.")
        SearchOrder = 0
        Call DisplaySelectionList(Text9, RefCode)
        Call CloseForm(FrmSelectionList)
        If CheckEmpty(Text9.Text, False) Then
            Text9.Text = "?"
        End If
        If RTrim(RefCode) <> "" Then
            SendKeys "{TAB}"
        End If
        Cancel = True
        Exit Sub
    ElseIf (Trim(rstMaterialIOChild.Fields("RefNo").Value) <> Trim(Text9.Text)) Or (CheckEmpty(rstMaterialIOChild.Fields("RefNo").Value, False)) Then
        If CheckDuplicateEntry Then
            Call DisplayError("Duplicate Entry")
            Text8.SelStart = 0
            Text8.SelLength = Len(Text9.Text)
            Cancel = True
            Exit Sub
        End If
    End If
    RefCode = rstRefList.Fields("Code").Value
    Text9.Text = Trim(rstRefList.Fields("Name").Value)
    If Val(CheckNull(rstMaterialIOChild.Fields("Quantity").Value)) = 0 Then
        MhRealInput3.Text = Format(Val(Right(Trim(rstRefList.Fields("Col0").Value), InStr(1, StrReverse(Trim(rstRefList.Fields("Col0").Value)), ":") - 1)), "0")
    End If
End Sub
Private Sub MhRealInput3_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyReturn Then
        If Val(MhRealInput3.Text) > Val(rstRefList.Fields("BalanceQuantity").Value) Then
            If Val(MhRealInput3.Text) > 0 Then
                Call DisplayError("Quantity cann't be greater than " & Format(Val(Right(Trim(rstRefList.Fields("Col0").Value), InStr(1, StrReverse(Trim(rstRefList.Fields("Col0").Value)), ":") - 1)), "#0.000"))
            End If
            MhRealInput3.SetFocus
        Else
            rstMaterialIOChild.Fields("Item").Value = OutsourceItemCode
            rstMaterialIOChild.Fields("OutsourceItemName").Value = Trim(Text5.Text)
            rstMaterialIOChild.Fields("Godown").Value = AccountCode
            rstMaterialIOChild.Fields("GodownName").Value = Trim(Text8.Text)
            rstMaterialIOChild.Fields("Ref").Value = RefCode
            rstMaterialIOChild.Fields("RefNo").Value = Pad(Trim(Text9.Text), Space(1), 10, "L")
            rstMaterialIOChild.Fields("Quantity").Value = Format(Val(MhRealInput3.Text), "0")
            rstMaterialIOChild.Update
            MakeTextBoxInvisible (False)
            CalculateTotal
            If rstMaterialIOChild.AbsolutePosition = rstMaterialIOChild.RecordCount Then
                Call DataGrid2_KeyDown(vbKeyA, vbCtrlMask)
            End If
        End If
    ElseIf Shift = 0 And KeyCode = vbKeyEscape Then
       MakeTextBoxInvisible (True)
    End If
End Sub
Private Sub MhRealInput3_Validate(Cancel As Boolean)
    Cancel = True
End Sub
Private Sub MakeTextBoxInvisible(ByVal KeyEscPressed As Boolean)
    If KeyEscPressed Then
        If Not (rstMaterialIOChild.EOF Or rstMaterialIOChild.BOF) Then
            If Val(CheckNull(rstMaterialIOChild.Fields("Quantity").Value)) = 0 Then
                rstMaterialIOChild.Delete
                rstMaterialIOChild.MoveNext
                If rstMaterialIOChild.RecordCount > 0 Then rstMaterialIOChild.MoveFirst
            End If
        End If
    End If
    Text5.Visible = False
    Text8.Visible = False
    Text9.Visible = False
    MhRealInput3.Visible = False
    DataGrid2.Enabled = True
    DataGrid2.SetFocus
End Sub
Private Sub CalculateTotal()
    Dim dblBookMark As Double
    
    MhRealInput4.Text = 0
    If rstMaterialIOChild.RecordCount <> 0 Then
        If Not (rstMaterialIOChild.EOF Or rstMaterialIOChild.BOF) Then
            dblBookMark = rstMaterialIOChild.Bookmark
        End If
        rstMaterialIOChild.MoveFirst
        Do While Not rstMaterialIOChild.EOF
            MhRealInput4.Text = Val(MhRealInput4.Text) + Val(rstMaterialIOChild.Fields("Quantity").Value)
            rstMaterialIOChild.MoveNext
        Loop
        If dblBookMark <> 0 Then
            rstMaterialIOChild.Bookmark = dblBookMark
       Else
            rstMaterialIOChild.MoveLast
       End If
    End If
End Sub
Private Function CheckDuplicateEntry() As Boolean
    Dim dblBookMark As Double
    
    If rstMaterialIOChild.RecordCount = 0 Then Exit Function
    If Not (rstMaterialIOChild.EOF Or rstMaterialIOChild.BOF) Then
       dblBookMark = rstMaterialIOChild.Bookmark
    End If
    rstMaterialIOChild.MoveFirst
    Do While Not rstMaterialIOChild.EOF
          If rstMaterialIOChild.Fields("OutsourceItemName").Value = Trim(Text5.Text) And Trim(rstMaterialIOChild.Fields("RefNo").Value) = Trim(Text9.Text) And Trim(rstMaterialIOChild.Fields("GodownName").Value) = Trim(Text8.Text) Then
             CheckDuplicateEntry = True
             Exit Do
          End If
          rstMaterialIOChild.MoveNext
    Loop
    If dblBookMark <> 0 Then
       rstMaterialIOChild.Bookmark = dblBookMark
    Else
       rstMaterialIOChild.MoveLast
    End If
End Function
Private Function UpdateOutsourceItemList(ByVal strOption As String) As Boolean
    On Error GoTo ErrorHandler
    
    UpdateOutsourceItemList = True
    If strOption = "D" Then
        CxnMaterialIssueOrder.Execute "Delete From MaterialIOChild WHERE Code = '" & rstMaterialIOParent.Fields("Code").Value & "'"
    Else
        CxnMaterialIssueOrder.Execute "Insert Into MaterialIOChild Values ('" & rstMaterialIOParent.Fields("Code").Value & "','1','" & rstMaterialIOChild.Fields("Item").Value & "','" & rstMaterialIOChild.Fields("Godown").Value & "','" & rstMaterialIOChild.Fields("Ref").Value & "'," & rstMaterialIOChild.Fields("Quantity").Value & ")"
    End If
    Exit Function
ErrorHandler:
    UpdateOutsourceItemList = False
End Function
Public Sub FilterRecord(ByVal SrchFor As String, ByVal SrchText As String)
    If SrchFor = "Supplier" Then
        rstMaterialIOList.Filter = "[SupplierName] Like '%" & SrchText & "%'"
    End If
End Sub
Private Sub LoadRefList(ByVal strOutsourceItemCode As String, ByVal strSupplierCode As String, ByVal strOrderCode As String)
    Dim BalanceQuantity As Long
    On Error GoTo ErrorHandler
    
    If rstRefList.State = adStateOpen Then
        rstRefList.Close
    End If
    rstRefList.Open "Select P.Name,Format(Quantity,0) As ReceivedQuantity,Format((Select Sum(Quantity) From MaterialIOChild Where MaterialIOChild.Ref=P.Code And MaterialIOChild.Item=C.OutsourceItem And MaterialIOChild.Code<>'" & strOrderCode & "'),0) As IssuedQuantity,Quantity As BalanceQuantity,Remarks As Col0,P.Code From OutsourceItemPOParent P Inner Join OutsourceItemPOChild C On (P.Code=C.Code And P.Supplier='" & strSupplierCode & "' And C.OutsourceItem='" & strOutsourceItemCode & "') Order By P.Name", CxnMaterialIssueOrder, adOpenKeyset, adLockOptimistic
    rstRefList.ActiveConnection = Nothing
    Do While Not rstRefList.EOF
        BalanceQuantity = (Val(CheckNull(rstRefList.Fields("ReceivedQuantity").Value)) - Val(CheckNull(rstRefList.Fields("IssuedQuantity").Value))) - CalculateQuantityIssued(strOutsourceItemCode)
        If BalanceQuantity <> 0 Then
            rstRefList.Fields("Col0").Value = Trim(rstRefList.Fields("Name").Value) + " Quantity : " + Format(BalanceQuantity, "0")
            rstRefList.Fields("BalanceQuantity").Value = BalanceQuantity
            rstRefList.Update
        Else
            rstRefList.Delete
        End If
        rstRefList.MoveNext
    Loop
    Exit Sub
ErrorHandler:
    DisplayError ("Failed to Load Ref List")
End Sub
Private Function CalculateQuantityIssued(ByVal strOutsourceItemCode As String) As Long
    Dim dblBookMark As Double
    
    If rstMaterialIOChild.RecordCount = 0 Then Exit Function
    If Not (rstMaterialIOChild.EOF Or rstMaterialIOChild.BOF) Then
       dblBookMark = rstMaterialIOChild.Bookmark
    End If
    rstMaterialIOChild.MoveFirst
    Do While Not rstMaterialIOChild.EOF
        If rstMaterialIOChild.Bookmark <> dblBookMark Then
            If Trim(rstMaterialIOChild.Fields("RefNo").Value) = Trim(rstRefList.Fields("Name").Value) And rstMaterialIOChild.Fields("Item").Value = strOutsourceItemCode Then
                CalculateQuantityIssued = CalculateQuantityIssued + Val(rstMaterialIOChild.Fields("Quantity").Value)
            End If
        End If
        rstMaterialIOChild.MoveNext
    Loop
    If dblBookMark <> 0 Then
        rstMaterialIOChild.Bookmark = dblBookMark
    Else
        rstMaterialIOChild.MoveLast
    End If
End Function
Private Sub PrintOutsourceItemIssueOrder()
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    rptOutsourceItemIssueOrder.Text2.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptOutsourceItemIssueOrder.Text3.SetText Trim(rstCompanyMaster.Fields("Address1").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address2").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address3").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address4").Value)
    
    If (Not CheckEmpty(rstCompanyMaster.Fields("Phone").Value, False)) And (Not CheckEmpty(rstCompanyMaster.Fields("Fax").Value, False)) Then
        rptOutsourceItemIssueOrder.Text24.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & Space(1) & "Fax : " & Trim(rstCompanyMaster.Fields("Fax").Value)
    ElseIf Not CheckEmpty(rstCompanyMaster.Fields("Fax").Value, False) Then
        rptOutsourceItemIssueOrder.Text24.SetText "Fax : " & Trim(rstCompanyMaster.Fields("Fax").Value)
    ElseIf Not CheckEmpty(rstCompanyMaster.Fields("Phone").Value, False) Then
        rptOutsourceItemIssueOrder.Text24.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value)
    Else
        rptOutsourceItemIssueOrder.Section5.Suppress = True
    End If
    If rstMaterialIOChild.State = adStateOpen Then
        rstMaterialIOChild.Close
    End If
    
    Dim rsss As String
    
    rsss = "Select Trim(P.Name) As OrderNo,[Date] As OrderDate,(Select Trim(PrintName) From AccountMaster Where Code = P.Source) As Godown,Trim(PrintName) As OutsourceItemName,(Select Trim(PrintName) From AccountMaster Where Code = C.Godown) As GodownName,(Select Trim(Name) From OutsourceItemPOParent Where Code = C.Ref) As RefNo,Quantity,Remarks From (MaterialIOParent As P Inner Join MaterialIOChild As C On (P.Code = C.Code And P.Code = '" & rstMaterialIOList.Fields("Code").Value & "')) Inner Join OutsourceItemMaster M On C.Item = M.Code Order By M.PrintName"
    
    rstMaterialIOChild.Open "Select Trim(P.Name) As OrderNo,[Date] As OrderDate,(Select Trim(PrintName) From AccountMaster Where Code = P.Source) As Godown,Trim(PrintName) As OutsourceItemName,(Select Trim(PrintName) From AccountMaster Where Code = C.Godown) As GodownName,(Select Trim(Name) From OutsourceItemPOParent Where Code = C.Ref) As RefNo,Quantity,Remarks From (MaterialIOParent As P Inner Join MaterialIOChild As C On (P.Code = C.Code And P.Code = '" & rstMaterialIOList.Fields("Code").Value & "')) Inner Join OutsourceItemMaster M On C.Item = M.Code Order By M.PrintName", CxnMaterialIssueOrder, adOpenKeyset, adLockOptimistic
    
    rptOutsourceItemIssueOrder.Text27.SetText "for " & Trim(rstMaterialIOChild.Fields("PrinterName").Value)
    rptOutsourceItemIssueOrder.Text9.SetText "for " & Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptOutsourceItemIssueOrder.Database.SetDataSource rstMaterialIOChild, 3, 1
    
    Screen.MousePointer = vbNormal
    
    If OutputTo = "S" Then
        Set FrmReportViewer.Report = rptOutsourceItemIssueOrder
        FrmReportViewer.Show vbModal
    Else
        rptOutsourceItemIssueOrder.PrintOut
    End If
    Set rptOutsourceItemIssueOrder = Nothing
    On Error GoTo 0
End Sub
