VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0F1F1508-C40A-101B-AD04-00AA00575482}#1.0#0"; "mhrinp32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FrmPaperMovement 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Paper Movement"
   ClientHeight    =   5160
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
   Icon            =   "PaperMovement.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5160
   ScaleWidth      =   8715
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   5160
      Left            =   15
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   8700
      _Version        =   65536
      _ExtentX        =   15346
      _ExtentY        =   9102
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
      Picture         =   "PaperMovement.frx":0442
      Begin TabDlg.SSTab SSTab1 
         Height          =   4920
         Left            =   120
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   120
         Width           =   8460
         _ExtentX        =   14923
         _ExtentY        =   8678
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
         TabPicture(0)   =   "PaperMovement.frx":045E
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "DataGrid1"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Text1"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "&Details"
         TabPicture(1)   =   "PaperMovement.frx":047A
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
            Top             =   4450
            Width           =   7760
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   3930
            Left            =   120
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   450
            Width           =   8235
            _ExtentX        =   14526
            _ExtentY        =   6932
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
               DataField       =   "GodownFromName"
               Caption         =   "Godown From"
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
               DataField       =   "GodownToName"
               Caption         =   "Godown To"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
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
                  ColumnWidth     =   1019.906
               EndProperty
               BeginProperty Column01 
                  Locked          =   -1  'True
                  ColumnWidth     =   1124.787
               EndProperty
               BeginProperty Column02 
                  Locked          =   -1  'True
                  ColumnWidth     =   2924.788
               EndProperty
               BeginProperty Column03 
                  Locked          =   -1  'True
                  ColumnWidth     =   2594.835
               EndProperty
            EndProperty
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   4170
            Left            =   -74880
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   480
            Width           =   8235
            _Version        =   65536
            _ExtentX        =   14526
            _ExtentY        =   7355
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
            Picture         =   "PaperMovement.frx":0496
            Begin VB.TextBox MhRealInput1 
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
               Left            =   6050
               MaxLength       =   13
               TabIndex        =   7
               Text            =   "0.000"
               Top             =   2025
               Visible         =   0   'False
               Width           =   1005
            End
            Begin VB.TextBox MhRealInput4 
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
               Left            =   7040
               MaxLength       =   13
               TabIndex        =   8
               Text            =   "0"
               Top             =   2025
               Visible         =   0   'False
               Width           =   870
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
               Left            =   1440
               MaxLength       =   60
               TabIndex        =   3
               Top             =   950
               Width           =   6690
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
               Left            =   4320
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   28
               TabStop         =   0   'False
               Top             =   3735
               Width           =   1095
            End
            Begin MhinrelLib.MhRealInput MhRealInput6 
               Height          =   255
               Left            =   7040
               TabIndex        =   25
               TabStop         =   0   'False
               Top             =   3300
               Width           =   870
               _Version        =   65536
               _ExtentX        =   1535
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
            Begin MhinrelLib.MhRealInput MhRealInput5 
               Height          =   255
               Left            =   6050
               TabIndex        =   24
               TabStop         =   0   'False
               Top             =   3300
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
               DecimalPlaces   =   3
               VAlignment      =   2
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel19 
               Height          =   255
               Left            =   120
               TabIndex        =   26
               Top             =   3300
               Width           =   8010
               _Version        =   65536
               _ExtentX        =   14129
               _ExtentY        =   441
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
               Picture         =   "PaperMovement.frx":04B2
               Picture         =   "PaperMovement.frx":04CE
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
               TabIndex        =   6
               Top             =   2025
               Visible         =   0   'False
               Width           =   5625
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
               TabIndex        =   4
               Top             =   1260
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
               MaxLength       =   60
               TabIndex        =   2
               Top             =   630
               Width           =   6690
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
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   16
               TabStop         =   0   'False
               Top             =   3735
               Width           =   1650
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
               Height          =   330
               Left            =   120
               TabIndex        =   17
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
               Caption         =   " Voucher No."
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "PaperMovement.frx":04EA
               Picture         =   "PaperMovement.frx":0506
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Left            =   5835
               TabIndex        =   18
               Top             =   105
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
               Caption         =   " Voucher Date"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "PaperMovement.frx":0522
               Picture         =   "PaperMovement.frx":053E
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
               Height          =   330
               Left            =   120
               TabIndex        =   19
               Top             =   3735
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
               Caption         =   " Make"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "PaperMovement.frx":055A
               Picture         =   "PaperMovement.frx":0576
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   330
               Left            =   120
               TabIndex        =   20
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
               Caption         =   " Godown From"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "PaperMovement.frx":0592
               Picture         =   "PaperMovement.frx":05AE
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
               Height          =   330
               Left            =   120
               TabIndex        =   21
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
               Caption         =   " Remarks"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "PaperMovement.frx":05CA
               Picture         =   "PaperMovement.frx":05E6
            End
            Begin MSDataGridLib.DataGrid DataGrid2 
               Height          =   1765
               Left            =   120
               TabIndex        =   5
               Top             =   1780
               Width           =   8010
               _ExtentX        =   14129
               _ExtentY        =   3096
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
               ColumnCount     =   3
               BeginProperty Column00 
                  DataField       =   "PaperName"
                  Caption         =   "Paper Name"
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
                  DataField       =   "QuantityOther"
                  Caption         =   "   Quantity"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   1
                     Format          =   "0.000"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2057
                     SubFormatType   =   1
                  EndProperty
               EndProperty
               BeginProperty Column02 
                  DataField       =   "Tat"
                  Caption         =   "       Tat"
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
                     ColumnWidth     =   5609.764
                  EndProperty
                  BeginProperty Column01 
                     Alignment       =   1
                     ColumnAllowSizing=   -1  'True
                     Locked          =   -1  'True
                     ColumnWidth     =   989.858
                  EndProperty
                  BeginProperty Column02 
                     Alignment       =   1
                     ColumnAllowSizing=   -1  'True
                     Locked          =   -1  'True
                     ColumnWidth     =   854.929
                  EndProperty
               EndProperty
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel17 
               Height          =   330
               Left            =   3075
               TabIndex        =   22
               Top             =   3735
               Width           =   1260
               _Version        =   65536
               _ExtentX        =   2222
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
               Caption         =   " Size"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "PaperMovement.frx":0602
               Picture         =   "PaperMovement.frx":061E
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel18 
               Height          =   330
               Left            =   5400
               TabIndex        =   23
               Top             =   3735
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
               Caption         =   " GSM"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "PaperMovement.frx":063A
               Picture         =   "PaperMovement.frx":0656
            End
            Begin MhinrelLib.MhRealInput MhRealInput7 
               Height          =   330
               Left            =   6600
               TabIndex        =   27
               TabStop         =   0   'False
               Top             =   3735
               Width           =   1530
               _Version        =   65536
               _ExtentX        =   2699
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
               MaxReal         =   9999999
               MinReal         =   0
               ReadOnly        =   -1  'True
               SpinChangeReal  =   0
               CaretColor      =   -2147483642
               DecimalPlaces   =   0
               VAlignment      =   2
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel12 
               Height          =   330
               Left            =   120
               TabIndex        =   29
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
               Caption         =   " Godown To"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "PaperMovement.frx":0672
               Picture         =   "PaperMovement.frx":068E
            End
            Begin MSMask.MaskEdBox MhDateInput1 
               Height          =   330
               Left            =   7035
               TabIndex        =   1
               Top             =   105
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   582
               _Version        =   393216
               Appearance      =   0
               BackColor       =   16777215
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "##-##-####"
               PromptChar      =   " "
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
               Y1              =   1680
               Y2              =   1680
            End
            Begin VB.Line Line3 
               X1              =   0
               X2              =   8280
               Y1              =   3630
               Y2              =   3630
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
            Top             =   4450
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
Attribute VB_Name = "FrmPaperMovement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CxnPaperMovement As New ADODB.Connection
Dim rstCompanyMaster As New ADODB.Recordset
Dim rstPaperMVList As New ADODB.Recordset
Dim rstPaperMVParent As New ADODB.Recordset
Dim WithEvents rstPaperMVChild As ADODB.Recordset
Attribute rstPaperMVChild.VB_VarHelpID = -1
Dim rstAccountList As New ADODB.Recordset
Dim rstPaperList As New ADODB.Recordset
Dim AccountFromCode As String
Dim AccountToCode As String
Dim PaperCode As String
Dim SortOrder As String
Dim PrevStr As String
Dim dblBookMark As Double
Dim blnRecordExist As Boolean
Dim BalanceQuantity As Long
Dim EMailID As String
Dim Attachment As String
Dim Message As String
Dim OutputTo As String
Public MovementType As String
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    
    CenterForm Me
    BusySystemIndicator True
    If MovementType = "1" Then
        Me.Caption = "Paper Movement [Book]"
    Else
        Me.Caption = "Paper Movement [Title]"
        DataGrid2.Columns(2).Visible = False
        MhRealInput6.Visible = False
        DataGrid2.Columns(0).Width = 6465.26
        Text5.Width = 6475
        MhRealInput1.Left = 6895
        MhRealInput5.Left = 6895
    End If
    CxnPaperMovement.CursorLocation = adUseClient
    CxnPaperMovement.Open CxnDatabase.ConnectionString
    rstCompanyMaster.Open "Select PrintName, Address1, Address2, Address3, Address4, Phone, Fax, EMail, Website From CompanyMaster", CxnPaperMovement, adOpenKeyset, adLockReadOnly
    rstPaperList.Open "Select PaperMaster.Name As Col0, Make, GSM, [Reams/Bundle], GeneralMaster.Name As SizeName, PaperMaster.Code From PaperMaster, GeneralMaster Where PaperMaster.[Size] = GeneralMaster.Code And PaperMaster.Type = '" & MovementType & "' Order by PaperMaster.Name", CxnPaperMovement, adOpenKeyset, adLockReadOnly
    rstAccountList.Open "SELECT TRIM(Name)+' ('+CHOOSE(VAL(Type)-4,'Book Printer','Title Printer','','Binder','Godown')+')'  As Col0,Code FROM AccountMaster WHERE Type IN ('05','06','08','09') ORDER BY Name", CxnPaperMovement, adOpenKeyset, adLockReadOnly
    rstPaperMVList.Open "Select PaperMVParent.Code, PaperMVParent.Name, Date, (Select Name From AccountMaster Where Code=PaperMVParent.AccountFrom) As GodownFromName,(Select Name From AccountMaster Where Code=PaperMVParent.AccountTo) As GodownToName From PaperMVParent Where MovementType = '" & MovementType & "' Order By PaperMVParent.Name", CxnPaperMovement, adOpenKeyset, adLockOptimistic
    rstPaperMVParent.CursorLocation = adUseClient
    Set rstPaperMVChild = New ADODB.Recordset
    rstPaperMVList.Filter = adFilterNone
    If rstPaperMVList.RecordCount > 0 Then rstPaperMVList.MoveLast
    Set DataGrid1.DataSource = rstPaperMVList
    BusySystemIndicator False
    SSTab1.Tab = 0
    SortOrder = "Name"
    If Not (rstPaperMVList.EOF Or rstPaperMVList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    rstPaperMVList.ActiveConnection = Nothing
    rstPaperList.ActiveConnection = Nothing
    rstAccountList.ActiveConnection = Nothing
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
                If Me.ActiveControl.Name <> "Text5" And Me.ActiveControl.Name <> "MhRealInput1" And Me.ActiveControl.Name <> "MhRealInput4" Then
                    If MsgBox("Are you sure to Quit?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Quit !") <> vbYes Then
                        Me.ActiveControl.SetFocus
                    Else
                        Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
                    End If
                End If
            End If
            If Me.ActiveControl.Name <> "Text5" And Me.ActiveControl.Name <> "MhRealInput1" And Me.ActiveControl.Name <> "MhRealInput4" Then
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
        If Me.ActiveControl.Name <> "Text5" And Me.ActiveControl.Name <> "MhRealInput1" And Me.ActiveControl.Name <> "MhRealInput4" Then
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
            If Me.ActiveControl.Name <> IIf(MovementType = "1", "MhRealInput4", "MhRealInput1") Then
                SendKeys "{TAB}"
            End If
        End If
        If Me.ActiveControl.Name <> IIf(MovementType = "1", "MhRealInput4", "MhRealInput1") Then
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
    Call CloseRecordset(rstPaperMVList)
    Call CloseRecordset(rstPaperMVParent)
    Call CloseRecordset(rstPaperMVChild)
    Call CloseRecordset(rstPaperList)
    Call CloseRecordset(rstAccountList)
    Call CloseConnection(CxnPaperMovement)
    ShowProgressInStatusBar False
    DisableChildMenu
End Sub
Private Sub Text1_Change()
    If rstPaperMVList.RecordCount = 0 Then Exit Sub
    rstPaperMVList.MoveFirst
    If Text1.Text <> "" Then
        If SortOrder = "Name" Then
           rstPaperMVList.Find "[" & SortOrder & "] Like '%" & FixQuote(Text1.Text) & "%'"
        Else
           rstPaperMVList.Find "[" & SortOrder & "] Like '" & FixQuote(Text1.Text) & "%'"
        End If
        If rstPaperMVList.EOF Then
            rstPaperMVList.MoveFirst
            If PrevStr <> "" And Len(Text1.Text) > 1 Then
                If dblBookMark <> 0 Then
                    rstPaperMVList.Bookmark = dblBookMark
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
    If Not (rstPaperMVList.EOF Or rstPaperMVList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim KeyProcessed As Boolean
    
    If rstPaperMVList.RecordCount = 0 Then Exit Sub
    If Shift = 0 And KeyCode = vbKeyUp Then
        With rstPaperMVList
            .MovePrevious
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyBack Then
        With rstPaperMVList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyDown Then
        With rstPaperMVList
            .MoveNext
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageUp Then
        With rstPaperMVList
            .Move (-1) * (DataGrid1.VisibleRows - 1)
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageUp Then
        With rstPaperMVList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageDown Then
        With rstPaperMVList
            .Move DataGrid1.VisibleRows - 1
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageDown Then
        With rstPaperMVList
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
            If Not (rstPaperMVList.EOF Or rstPaperMVList.BOF) Then
                With DataGrid1.SelBookmarks
                    If .Count <> 0 Then .Remove 0
                    .Add DataGrid1.Bookmark
                End With
            End If
            'Error Occurs On 12.09.14
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
        If rstPaperMVParent.State = adStateOpen Then
           rstPaperMVParent.Close
        End If
        rstPaperMVParent.Open "Select * From PaperMVParent Where Code = ''", CxnPaperMovement, adOpenKeyset, adLockOptimistic
        ClearFields ("P")
        ClearFields ("C")
        Call LoadPaperList("")
        If rstPaperMVChild.State = adStateClosed Then
            SSTab1.Tab = 0
            Exit Sub
        End If
        If AddRecord(rstPaperMVParent) Then
            Text2.Text = GenerateCode(CxnPaperMovement, "Select Max(Val(Name)) From PaperMVParent Where MovementType = '" & MovementType & "'", 10, Space(1))
            MhDateInput1.Text = Format(Date, "dd-MM-yyyy")
            Call SetButtons(False)
            SSTab1.Tab = 1
            Text2.SetFocus
            blnRecordExist = False
            CxnPaperMovement.BeginTrans
        End If
    ElseIf Button.Index = 2 Then
        If rstPaperMVList.RecordCount = 0 Then Exit Sub
        SSTab1.Tab = 1
        EditRecord
    ElseIf Button.Index = 3 Then
        If rstPaperMVList.RecordCount = 0 Then Exit Sub
        If AllowTransactionsDeletion = 0 Then
            Call DisplayError("You don't have the rights to Delete this Voucher")
            Exit Sub
        End If
        SSTab1.Tab = 1
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            On Error Resume Next
            MdiMainMenu.MousePointer = vbHourglass
            CxnPaperMovement.Execute "Delete From PaperMVParent Where Code = '" & rstPaperMVList.Fields("Code").Value & "'"
            MdiMainMenu.MousePointer = vbNormal
            If Err.Number = 0 Then
                rstPaperMVList.Delete
                rstPaperMVList.MoveNext
                If rstPaperMVList.RecordCount > 0 And rstPaperMVList.EOF Then
                    rstPaperMVList.MoveLast
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
        If UpdateRecord(rstPaperMVParent) Then
            If UpdatePaperList("D") Then
                 UpdateFlag = 1
                 If rstPaperMVChild.RecordCount <> 0 Then
                      rstPaperMVChild.MoveFirst
                      Do While Not rstPaperMVChild.EOF
                          If Val(rstPaperMVChild.Fields("QuantityOther").Value) <> 0 Then
                               If Not UpdatePaperList("U") Then
                                    UpdateFlag = 0
                                    Exit Do
                                End If
                          End If
                          rstPaperMVChild.MoveNext
                      Loop
                 End If
            End If
        End If
        If UpdateFlag Then
            AddToList
            CxnPaperMovement.CommitTrans
            If rstPaperMVParent.State = adStateOpen Then
                rstPaperMVParent.Close
            End If
            rstPaperMVParent.CursorLocation = adUseClient
            Call SetButtons(True)
            SSTab1.Tab = 0
            ShowProgressInStatusBar True
            Timer1.Enabled = True
        Else
            DisplayError ("Failed to save the record")
            Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
        End If
    ElseIf Button.Index = 5 Then
        If CancelRecordUpdate(rstPaperMVParent) Then
            CxnPaperMovement.RollbackTrans
            If rstPaperMVParent.State = adStateOpen Then
                rstPaperMVParent.Close
            End If
            rstPaperMVParent.CursorLocation = adUseClient
            Call SetButtons(True)
            SetButtonsForNoRecord
            SSTab1.Tab = 0
        End If
    ElseIf Button.Index = 6 Then
        SSTab1.Tab = 0
        Set DataGrid1.DataSource = Nothing
        rstPaperMVList.ActiveConnection = CxnPaperMovement
        Do While Not RefreshRecord(rstPaperMVList)
        Loop
        Set DataGrid1.DataSource = rstPaperMVList
        rstPaperMVList.ActiveConnection = Nothing
        If rstPaperMVList.RecordCount > 0 Then rstPaperMVList.MoveLast
        rstAccountList.ActiveConnection = CxnPaperMovement
        Do While Not RefreshRecord(rstAccountList)
        Loop
        rstAccountList.ActiveConnection = Nothing
        rstPaperList.ActiveConnection = CxnPaperMovement
        Do While Not RefreshRecord(rstPaperList)
        Loop
        rstPaperList.ActiveConnection = Nothing
        HiLiteRecord = True
    ElseIf Button.Index = 7 Then
        SSTab1.Tab = 0
        With FrmFilter
            .Combo1.AddItem "Godown From", 0
            .Combo1.AddItem "Godown To", 0
            .Combo1.ListIndex = 0
            Set .srcForm = Me
            .Show vbModal
        End With
        HiLiteRecord = True
    ElseIf Button.Index = 9 Then
        If rstPaperMVList.RecordCount = 0 Then Exit Sub
        OutputTo = "P"
        PrintPaperMovement
        HiLiteRecord = True
    ElseIf Button.Index = 10 Then
        If rstPaperMVList.RecordCount = 0 Then Exit Sub
        OutputTo = "S"
        PrintPaperMovement
        HiLiteRecord = True
    ElseIf Button.Index = 13 Then
        If rstPaperMVList.RecordCount > 0 Then rstPaperMVList.MoveFirst
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 14 Then
        If rstPaperMVList.RecordCount > 0 Then
            rstPaperMVList.MovePrevious
            If rstPaperMVList.BOF Then
                rstPaperMVList.MoveNext
            End If
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 15 Then
        If rstPaperMVList.RecordCount > 0 Then
            rstPaperMVList.MoveNext
            If rstPaperMVList.EOF Then
                rstPaperMVList.MovePrevious
            End If
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 16 Then
        If rstPaperMVList.RecordCount > 0 Then rstPaperMVList.MoveLast
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 18 Then
        Unload Me
        HiLiteRecord = False
    End If
    If HiLiteRecord Then
        If Not (rstPaperMVList.EOF Or rstPaperMVList.BOF) Then
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
          rstPaperMVList.Sort = "Name Asc"
       End If
    ElseIf ColIndex = 2 Then
       If SortOrder <> "GodownFromName" Then
          SortOrder = "GodownFromName"
          rstPaperMVList.Sort = "GodownFromName Asc"
       End If
    ElseIf ColIndex = 3 Then
       If SortOrder <> "GodownToName" Then
          SortOrder = "GodownToName"
          rstPaperMVList.Sort = "GodownToName Asc"
       End If
    End If
    DataGrid1.ClearSelCols
    If Not (rstPaperMVList.EOF Or rstPaperMVList.BOF) Then
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
    If rstPaperMVList.RecordCount = 0 Then
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
    If rstPaperMVParent.EOF Or rstPaperMVParent.BOF Then Exit Sub
    If CheckEmpty(Text2, True) Then
        Cancel = True
    ElseIf CheckDuplicate(CxnPaperMovement, "PaperMVParent", "Code", "[Name]+MovementType", Trim(Text2.Text) & MovementType, rstPaperMVParent.Fields("Code").Value, False) Then
        Cancel = True
    End If
End Sub
Private Sub MhDateInput1_GotFocus()
    FocusSelect Me.ActiveControl
End Sub
Private Sub MhDateInput1_Validate(Cancel As Boolean)
    If Not ValidateDate(Me.ActiveControl) Then
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
        AccountFromCode = ""
        Call LoadSelectionList(rstAccountList, "List of Godowns...", "Name")
        SearchOrder = 0
        Call DisplaySelectionList(Text3, AccountFromCode)
        Call CloseForm(FrmSelectionList)
        If CheckEmpty(Text3.Text, False) Then
            Text3.Text = "?"
        End If
        If RTrim(AccountFromCode) <> "" Then
            SendKeys "{TAB}"
        End If
        Cancel = True
    Else
        AccountFromCode = rstAccountList.Fields("Code").Value
    End If
End Sub
Private Sub Text9_Change()
    If Text9.Text = " " Then
        Text9.Text = "?"
        SendKeys "{TAB}"
    End If
End Sub
Private Sub Text9_Validate(Cancel As Boolean)
    Dim SearchString As String
    
    SearchString = FixQuote(Text9.Text)
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
        AccountToCode = ""
        Call LoadSelectionList(rstAccountList, "List of Godowns...", "Name")
        SearchOrder = 0
        Call DisplaySelectionList(Text9, AccountToCode)
        Call CloseForm(FrmSelectionList)
        If CheckEmpty(Text9.Text, False) Then
            Text9.Text = "?"
        End If
        If RTrim(AccountToCode) <> "" Then
            SendKeys "{TAB}"
        End If
        Cancel = True
        Exit Sub
    ElseIf Text3.Text = Text9.Text Then
        Call DisplayError("Both godowns cann't be same")
        Text9.SelStart = 0
        Text9.SelLength = Len(Text9.Text)
        Cancel = True
        Exit Sub
    End If
    AccountToCode = rstAccountList.Fields("Code").Value
End Sub
Private Sub Text4_Validate(Cancel As Boolean)
    If rstPaperMVChild.RecordCount = 0 Then
        SendKeys "^"
        Call AddRecord(rstPaperMVChild)
        Call ClearFields("C")
        Call DataGrid2_KeyDown(vbKeyE, vbCtrlMask)
    End If
End Sub
Private Sub ViewRecord()
    ClearFields ("P")
    ClearFields ("C")
    If rstPaperMVList.EOF Then
        If rstPaperMVChild.State = adStateOpen Then
            rstPaperMVChild.Close
        End If
        Exit Sub
    End If
    FindRecord
    LoadFields
End Sub
Private Sub FindRecord()
    If rstPaperMVParent.State = adStateOpen Then
       rstPaperMVParent.Close
    End If
    rstPaperMVParent.Open "Select * From PaperMVParent Where Code = '" & FixQuote(rstPaperMVList.Fields("Code").Value) & "'", CxnPaperMovement, adOpenKeyset, adLockOptimistic
    If rstPaperMVParent.RecordCount = 0 Then
       Call DisplayError("This Record has been deleted by Another User ! Click Ok To Refresh the Recordset")
       Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
    End If
End Sub
Private Sub ClearFields(ByVal strType As String)
    If strType = "P" Then
        Text2.Text = ""
        Text3.Text = ""
        Text9.Text = ""
        Text4.Text = ""
        MhDateInput1.Text = Format(Date, "dd-MM-yyyy")
        MhRealInput5.Text = 0#
        MhRealInput6.Text = 0#
    ElseIf strType = "C" Then
        Text5.Text = ""
        Text6.Text = ""
        Text7.Text = ""
        MhRealInput1.Text = "0.000"
        MhRealInput4.Text = "0"
        MhRealInput7.Text = 0#
    End If
End Sub
Private Sub LoadFields()
    If rstPaperMVParent.EOF Or rstPaperMVParent.BOF Then Exit Sub
    Text2.Text = rstPaperMVParent.Fields("Name").Value
    MhDateInput1.Text = Format(rstPaperMVParent.Fields("Date").Value, "dd-MM-yyyy")
    AccountFromCode = rstPaperMVParent.Fields("AccountFrom").Value
    If rstAccountList.RecordCount > 0 Then rstAccountList.MoveFirst
    rstAccountList.Find "[Code] = '" & AccountFromCode & "'"
    If Not rstAccountList.EOF Then
       Text3.Text = rstAccountList.Fields("Col0").Value
    End If
    AccountToCode = rstPaperMVParent.Fields("AccountTo").Value
    If rstAccountList.RecordCount > 0 Then rstAccountList.MoveFirst
    rstAccountList.Find "[Code] = '" & AccountToCode & "'"
    If Not rstAccountList.EOF Then
       Text9.Text = rstAccountList.Fields("Col0").Value
    End If
    Text4.Text = rstPaperMVParent.Fields("Remarks").Value
    Call LoadPaperList(rstPaperMVParent.Fields("Code").Value)
    If rstPaperMVChild.State = adStateOpen Then
        CalculateTotal
    End If
End Sub
Private Sub EditRecord()
    On Error GoTo ErrorHandler
    
    If rstPaperMVParent.RecordCount = 0 Then Exit Sub
    If rstPaperMVChild.State = adStateClosed Then
        SSTab1.Tab = 0
        Exit Sub
    End If
    If rstPaperMVParent.State = adStateOpen Then
       rstPaperMVParent.Close
    End If
    rstPaperMVParent.CursorLocation = adUseServer
    rstPaperMVParent.Open "Select * From PaperMVParent Where Code = '" & FixQuote(rstPaperMVList.Fields("Code").Value) & "'", CxnPaperMovement, adOpenKeyset, adLockPessimistic
    MdiMainMenu.MousePointer = vbHourglass
    rstPaperMVParent.Fields("Printstatus") = "N"
    MdiMainMenu.MousePointer = vbNormal
    AddToList
    Call SetButtons(False)
    SSTab1.TabEnabled(0) = False
    Text2.SetFocus
    blnRecordExist = True
    CxnPaperMovement.BeginTrans
    Exit Sub
ErrorHandler:
    If Err.Number = -2147467259 Then
       Call DisplayError("Failed to Edit the record")
    End If
    MdiMainMenu.MousePointer = vbNormal
    SSTab1.Tab = 0
End Sub
Private Sub SaveFields()
    If rstPaperMVParent.EOF Or rstPaperMVParent.BOF Then Exit Sub
    If Not blnRecordExist Then
        rstPaperMVParent.Fields("Code").Value = GenerateCode(CxnPaperMovement, "Select Max(Code) From PaperMVParent", 6, "0")
        rstPaperMVParent.Fields("CreatedBy").Value = UserCode
        rstPaperMVParent.Fields("CreatedOn").Value = Now()
        rstPaperMVParent.Fields("Recordstatus").Value = "N"
    Else
        rstPaperMVParent.Fields("ModifiedBy").Value = UserCode
        rstPaperMVParent.Fields("ModifiedOn").Value = Now()
        rstPaperMVParent.Fields("Recordstatus").Value = "M"
    End If
    rstPaperMVParent.Fields("Name").Value = Pad(Trim(Text2.Text), Space(1), 10, "L")
    rstPaperMVParent.Fields("Date").Value = GetDate(MhDateInput1.Text)
    rstPaperMVParent.Fields("MovementType").Value = MovementType
    rstPaperMVParent.Fields("AccountFrom").Value = AccountFromCode
    rstPaperMVParent.Fields("AccountTo").Value = AccountToCode
    rstPaperMVParent.Fields("Remarks").Value = Trim(Text4.Text)
    rstPaperMVParent.Fields("PrintStatus").Value = "N"
End Sub
Private Sub AddToList()
    On Error Resume Next
    
    rstPaperMVList.MoveFirst
    rstPaperMVList.Find "[Code] = '" & rstPaperMVParent.Fields("Code").Value & "'"
    If rstPaperMVList.EOF Then
       rstPaperMVList.AddNew
       rstPaperMVList.Fields("Code").Value = rstPaperMVParent.Fields("Code").Value
    End If
    rstPaperMVList.Fields("Name").Value = Pad(rstPaperMVParent.Fields("Name").Value, Space(1), 10, "L")
    rstPaperMVList.Fields("Date").Value = rstPaperMVParent.Fields("Date").Value
    rstAccountList.MoveFirst
    rstAccountList.Find "[Code] = '" & rstPaperMVParent.Fields("AccountFrom").Value & "'"
    rstPaperMVList.Fields("GodownFromName").Value = Trim(rstAccountList.Fields("Col0").Value)
    rstAccountList.MoveFirst
    rstAccountList.Find "[Code] = '" & rstPaperMVParent.Fields("AccountTo").Value & "'"
    rstPaperMVList.Fields("GodownToName").Value = Trim(rstAccountList.Fields("Col0").Value)
    rstPaperMVList.Update
    rstPaperMVList.Sort = SortOrder & " Asc"
    rstPaperMVList.Find "[Code] = '" & rstPaperMVParent.Fields("Code").Value & "'"
End Sub
Private Function CheckMandatoryFields() As Boolean
    If CheckEmpty(Text2.Text, False) Then
       DisplayError ("Voucher No. cannot be blank")
       Text2.SetFocus
       CheckMandatoryFields = True
    ElseIf CheckEmpty(Text3.Text, False) Then
       Text3.SetFocus
       CheckMandatoryFields = True
    ElseIf Not CheckExists(Text3, "Col0", rstAccountList, AccountFromCode) Then
        Text3.SetFocus
        CheckMandatoryFields = True
    ElseIf CheckEmpty(Text9.Text, False) Then
       Text9.SetFocus
       CheckMandatoryFields = True
    ElseIf Not CheckExists(Text9, "Col0", rstAccountList, AccountToCode) Then
        Text9.SetFocus
        CheckMandatoryFields = True
    ElseIf Text3.Text = Text9.Text Then
       DisplayError ("Both godowns cann't be same")
       Text9.SetFocus
       CheckMandatoryFields = True
    ElseIf CheckDuplicate(CxnPaperMovement, "PaperMVParent", "Code", "[Name]+MovementType", Trim(Text2.Text) & MovementType, rstPaperMVParent.Fields("Code").Value, False) Then
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
Private Sub LoadPaperList(ByVal strVoucherCode As String)
    On Error GoTo ErrorHandler
    
    If rstPaperMVChild.State = adStateOpen Then
       rstPaperMVChild.Close
    End If
    rstPaperMVChild.Open "Select Paper, PaperMaster.Name As PaperName, Make, GSM, GeneralMaster.Name As SizeName, QuantityOther,Tat From PaperMaster, GeneralMaster, PaperMVChild Where PaperMVChild.Paper = PaperMaster.Code And PaperMaster.[Size] = GeneralMaster.Code And PaperMVChild.Code = '" & strVoucherCode & "' Order by PaperMaster.Name", CxnPaperMovement, adOpenKeyset, adLockOptimistic
    rstPaperMVChild.ActiveConnection = Nothing
    Set DataGrid2.DataSource = rstPaperMVChild
    Exit Sub
ErrorHandler:
    DisplayError ("Failed to Load Paper List")
End Sub
Private Sub DataGrid2_DblClick()
    Call DataGrid2_KeyDown(vbKeyE, vbCtrlMask)
End Sub
Private Sub DataGrid2_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask And KeyCode = vbKeyE Then
        If rstPaperMVChild.RecordCount = 0 Then
            KeyCode = 0
            Exit Sub
        End If
        If Val(CheckNull(rstPaperMVChild.Fields("QuantityOther").Value)) <> 0 Then
            PaperCode = rstPaperMVChild.Fields("Paper").Value
            Text5.Text = rstPaperMVChild.Fields("PaperName").Value
            MhRealInput1.Text = Format(Val(rstPaperMVChild.Fields("QuantityOther").Value), "0.000")
            MhRealInput4.Text = Format(Val(rstPaperMVChild.Fields("Tat").Value), "0")
        End If
        With DataGrid2
            Text5.Visible = True
            Text5.Move .Left + .Columns(0).Left, .Top + .RowTop(.Row), .Columns(0).Width + 10, .RowHeight + 30
            MhRealInput1.Visible = True
            MhRealInput1.Move .Left + .Columns(1).Left, .Top + .RowTop(.Row), .Columns(1).Width + 10, .RowHeight + 30
            If MovementType = "1" Then
                MhRealInput4.Visible = True
                MhRealInput4.Move .Left + .Columns(2).Left, .Top + .RowTop(.Row), .Columns(2).Width + 10, .RowHeight + 30
            End If
        End With
        DataGrid2.Enabled = False
        Text5.SetFocus
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyA Then
        SendKeys "^"
        Call AddRecord(rstPaperMVChild)
        Call ClearFields("C")
        Call DataGrid2_KeyDown(vbKeyE, vbCtrlMask)
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyD Then
        If rstPaperMVChild.RecordCount = 0 Then Exit Sub
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            Set DataGrid2.DataSource = Nothing
            rstPaperMVChild.Delete
            rstPaperMVChild.MoveNext
            Set DataGrid2.DataSource = rstPaperMVChild
            CalculateTotal
            DataGrid2.SetFocus
        End If
        If rstPaperMVChild.RecordCount = 0 Then
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
Private Sub rstPaperMVChild_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    On Error Resume Next
    
    If Not (rstPaperMVChild.EOF Or rstPaperMVChild.BOF) Then
        If Not IsNull(rstPaperMVChild.Fields("Make").Value) Then
            Text6.Text = rstPaperMVChild.Fields("Make").Value
        End If
        If Not IsNull(rstPaperMVChild.Fields("SizeName").Value) Then
            Text7.Text = rstPaperMVChild.Fields("SizeName").Value
        End If
        If Not IsNull(rstPaperMVChild.Fields("GSM").Value) Then
            MhRealInput7.Text = Val(rstPaperMVChild.Fields("GSM").Value)
        End If
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
    If rstPaperList.RecordCount = 0 Then
        DisplayError ("No Record in Paper Master")
        Cancel = True
        Exit Sub
    Else
        rstPaperList.MoveFirst
    End If
    rstPaperList.Find "[Col0] = '" & RTrim(SearchString) & "'"
    If rstPaperList.EOF Then
        SelectionType = "S"
        PaperCode = ""
        Call LoadSelectionList(rstPaperList, "List of Papers...", "Name")
        SearchOrder = 0
        Call DisplaySelectionList(Text5, PaperCode)
        Call CloseForm(FrmSelectionList)
        If CheckEmpty(Text5.Text, False) Then
            Text5.Text = "?"
        End If
        If RTrim(PaperCode) <> "" Then
            SendKeys "{TAB}"
        End If
        Cancel = True
        Exit Sub
    ElseIf (rstPaperMVChild.Fields("PaperName").Value <> Text5.Text) Or (CheckEmpty(rstPaperMVChild.Fields("PaperName").Value, False)) Then
        If CheckDuplicatePaper Then
            Call DisplayError("Duplicate Entry")
            Text5.SelStart = 0
            Text5.SelLength = Len(Text5.Text)
            Cancel = True
            Exit Sub
        End If
    End If
    PaperCode = rstPaperList.Fields("Code").Value
    Text6.Text = rstPaperList.Fields("Make").Value
    Text7.Text = rstPaperList.Fields("SizeName").Value
    MhRealInput7.Text = Val(rstPaperList.Fields("GSM").Value)
    BalanceQuantity = CalculatePaperBalance(AccountFromCode, PaperCode, CheckNull(rstPaperMVParent.Fields("Code").Value), IIf(MovementType = "1", "PMVB", "PMVT"))
    MsgBox "Quantity at Godown : " & Format(str(Int(BalanceQuantity / 500) + ((BalanceQuantity Mod 500) / 1000)), "#0.000")
End Sub
Private Sub MhRealInput1_GotFocus()
    FocusSelect Me.ActiveControl
End Sub
Private Sub MhRealInput1_KeyPress(KeyAscii As Integer)
    ValidateKey MhRealInput1, KeyAscii, 3
End Sub
Private Sub MhRealInput1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim Quantity As Long
    
    If Shift = 0 And KeyCode = vbKeyReturn Then
        If Not ValidateNumber(Me.ActiveControl, 3) Then Exit Sub
        Quantity = (Int(Val(MhRealInput1.Text)) * 500) + ((Val(MhRealInput1.Text) - Int(Val(MhRealInput1.Text))) * 1000)
        If Val(MhRealInput1.Text) <= 0 Or Quantity > BalanceQuantity Then
            If Val(MhRealInput1.Text) > 0 Then
                Call DisplayError("Quantity cann't be greater than " & Format(str(Int(BalanceQuantity / 500) + ((BalanceQuantity Mod 500) / 1000)), "#0.000"))
            End If
            MhRealInput1.SetFocus
            FocusSelect Me.ActiveControl
        Else
            rstPaperMVChild.Fields("Paper").Value = PaperCode
            rstPaperMVChild.Fields("PaperName").Value = Trim(Text5.Text)
            rstPaperMVChild.Fields("Make").Value = Trim(Text6.Text)
            rstPaperMVChild.Fields("SizeName").Value = Trim(Text7.Text)
            rstPaperMVChild.Fields("GSM").Value = Val(MhRealInput7.Text)
            rstPaperMVChild.Fields("QuantityOther").Value = Format(Val(MhRealInput1.Text), "0.000")
            rstPaperMVChild.Fields("Tat").Value = 0
            rstPaperMVChild.Update
            MakeTextBoxInvisible (False)
            CalculateTotal
            If rstPaperMVChild.AbsolutePosition = rstPaperMVChild.RecordCount Then
                Call DataGrid2_KeyDown(vbKeyA, vbCtrlMask)
            End If
        End If
    ElseIf Shift = 0 And KeyCode = vbKeyEscape Then
        MakeTextBoxInvisible (True)
    End If
End Sub
Private Sub MhRealInput1_Validate(Cancel As Boolean)
    Dim Quantity As Long
    Dim RPB As Double
    
    If Not ValidateNumber(Me.ActiveControl, 3) Then
        Cancel = True
    Else
        Quantity = (Int(Val(MhRealInput1.Text)) * 500) + ((Val(MhRealInput1.Text) - Int(Val(MhRealInput1.Text))) * 1000)
        If Val(MhRealInput1.Text) <= 0 Or Quantity > BalanceQuantity Then
            If Val(MhRealInput1.Text) > 0 Then
                Call DisplayError("Quantity cann't be greater than " & Format(str(Int(BalanceQuantity / 500) + ((BalanceQuantity Mod 500) / 1000)), "#0.000"))
            End If
            Cancel = True
            MhRealInput1.SetFocus
            FocusSelect Me.ActiveControl
        Else
            RPB = Val(rstPaperList.Fields("Reams/Bundle").Value)
            If RPB <> 0 Then
                MhRealInput4.Text = Format(Int(Val(MhRealInput1.Text) / RPB), "0")
            End If
        End If
    End If
End Sub
Private Sub MhRealInput4_GotFocus()
    FocusSelect Me.ActiveControl
End Sub
Private Sub MhRealInput4_KeyPress(KeyAscii As Integer)
    ValidateKey MhRealInput4, KeyAscii, 0
End Sub
Private Sub MhRealInput4_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyReturn Then
        If Not ValidateNumber(Me.ActiveControl, 0) Then Exit Sub
        If Val(MhRealInput4.Text) >= 0 Then
            rstPaperMVChild.Fields("Paper").Value = PaperCode
            rstPaperMVChild.Fields("PaperName").Value = Trim(Text5.Text)
            rstPaperMVChild.Fields("Make").Value = Trim(Text6.Text)
            rstPaperMVChild.Fields("SizeName").Value = Trim(Text7.Text)
            rstPaperMVChild.Fields("GSM").Value = Val(MhRealInput7.Text)
            rstPaperMVChild.Fields("QuantityOther").Value = Format(Val(MhRealInput1.Text), "0.000")
            rstPaperMVChild.Fields("Tat").Value = Format(Val(MhRealInput4.Text), "0")
            rstPaperMVChild.Update
            MakeTextBoxInvisible (False)
            CalculateTotal
            If rstPaperMVChild.AbsolutePosition = rstPaperMVChild.RecordCount Then
                Call DataGrid2_KeyDown(vbKeyA, vbCtrlMask)
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
        If Not (rstPaperMVChild.EOF Or rstPaperMVChild.BOF) Then
            If Val(CheckNull(rstPaperMVChild.Fields("QuantityOther").Value)) = 0 Then
                rstPaperMVChild.Delete
                rstPaperMVChild.MoveNext
                If rstPaperMVChild.RecordCount > 0 Then rstPaperMVChild.MoveFirst
            End If
        End If
    End If
    Text5.Visible = False
    MhRealInput1.Visible = False
    MhRealInput4.Visible = False
    DataGrid2.Enabled = True
    DataGrid2.SetFocus
End Sub
Private Sub CalculateTotal()
    Dim dblBookMark As Double
    
    MhRealInput5.Text = 0
    MhRealInput6.Text = 0
    If rstPaperMVChild.RecordCount <> 0 Then
        If Not (rstPaperMVChild.EOF Or rstPaperMVChild.BOF) Then
            dblBookMark = rstPaperMVChild.Bookmark
        End If
        rstPaperMVChild.MoveFirst
        Do While Not rstPaperMVChild.EOF
            MhRealInput5.Text = Val(MhRealInput5.Text) + Val(rstPaperMVChild.Fields("QuantityOther").Value)
            MhRealInput6.Text = Val(MhRealInput6.Text) + Val(rstPaperMVChild.Fields("Tat").Value)
            rstPaperMVChild.MoveNext
        Loop
        If dblBookMark <> 0 Then
            rstPaperMVChild.Bookmark = dblBookMark
       Else
            rstPaperMVChild.MoveLast
       End If
    End If
End Sub
Private Function CheckDuplicatePaper() As Boolean
    Dim dblBookMark As Double
    
    If rstPaperMVChild.RecordCount = 0 Then Exit Function
    If Not (rstPaperMVChild.EOF Or rstPaperMVChild.BOF) Then
       dblBookMark = rstPaperMVChild.Bookmark
    End If
    rstPaperMVChild.MoveFirst
    Do While Not rstPaperMVChild.EOF
          If rstPaperMVChild.Fields("PaperName").Value = Trim(Text5.Text) Then
             CheckDuplicatePaper = True
             Exit Do
          End If
          rstPaperMVChild.MoveNext
    Loop
    If dblBookMark <> 0 Then
       rstPaperMVChild.Bookmark = dblBookMark
    Else
       rstPaperMVChild.MoveLast
    End If
End Function
Private Function UpdatePaperList(ByVal strOption As String) As Boolean
    Dim Sheets As Long
    On Error GoTo ErrorHandler
    
    UpdatePaperList = True
    If strOption = "D" Then
        CxnPaperMovement.Execute "Delete From PaperMVChild Where Code = '" & rstPaperMVParent.Fields("Code").Value & "'"
    Else
        Sheets = (Int(Val(rstPaperMVChild.Fields("QuantityOther").Value)) * 500) + ((Val(rstPaperMVChild.Fields("QuantityOther").Value) - Int(Val(rstPaperMVChild.Fields("QuantityOther").Value))) * 1000)
        CxnPaperMovement.Execute "Insert Into PaperMVChild Values ('" & rstPaperMVParent.Fields("Code").Value & "','" & rstPaperMVChild.Fields("Paper").Value & "'," & rstPaperMVChild.Fields("QuantityOther").Value & "," & Sheets & "," & rstPaperMVChild.Fields("Tat").Value & ")"
    End If
    Exit Function
ErrorHandler:
    UpdatePaperList = False
End Function
Public Sub FilterRecord(ByVal SrchFor As String, ByVal SrchText As String)
    If SrchFor = "Godown From" Then
        rstPaperMVList.Filter = "[GodownFromName] Like '%" & SrchText & "%'"
    Else
        rstPaperMVList.Filter = "[GodownToName] Like '%" & SrchText & "%'"
    End If
End Sub
Private Sub PrintPaperMovement()
    Dim oOutlookMsg As Outlook.MailItem
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    rptPaperMovement.Text1.SetText IIf(MovementType = "1", "Book", "Title") & " Paper Movement"
    rptPaperMovement.Text2.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptPaperMovement.Text3.SetText Trim(rstCompanyMaster.Fields("Address1").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address2").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address3").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address4").Value)
    If (Not CheckEmpty(rstCompanyMaster.Fields("Phone").Value, False)) And (Not CheckEmpty(rstCompanyMaster.Fields("Fax").Value, False)) Then
        rptPaperMovement.Text24.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & Space(1) & "Fax : " & Trim(rstCompanyMaster.Fields("Fax").Value)
    ElseIf Not CheckEmpty(rstCompanyMaster.Fields("Fax").Value, False) Then
        rptPaperMovement.Text24.SetText "Fax : " & Trim(rstCompanyMaster.Fields("Fax").Value)
    ElseIf Not CheckEmpty(rstCompanyMaster.Fields("Phone").Value, False) Then
        rptPaperMovement.Text24.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value)
    Else
        rptPaperMovement.Section5.Suppress = True
    End If
    If rstPaperMVChild.State = adStateOpen Then
        rstPaperMVChild.Close
    End If
    rstPaperMVChild.Open "Select Trim(PaperMVParent.Name) As VchNo,[Date] As VchDate,(Select Trim(PrintName) From AccountMaster Where Code = PaperMVParent.AccountFrom) As GodownFrom,(Select Trim(PrintName) From AccountMaster Where Code = PaperMVParent.AccountTo) As GodownTo,Trim(PrintName) As PaperName,QuantityOther As Quantity,Remarks,(Select Trim(EMail) From AccountMaster Where Code=PaperMVParent.AccountFrom) As TransferorEMailID,(Select Trim(EMail) From AccountMaster Where Code=PaperMVParent.AccountTo) As TransfereeEMailID From (PaperMVParent Left Join PaperMVChild On (PaperMVParent.Code = PaperMVChild.Code And MovementType = '" & MovementType & "' And PaperMVParent.Code = '" & rstPaperMVList.Fields("Code").Value & "')) Left Join PaperMaster On PaperMVChild.Paper = PaperMaster.Code Order By PaperMaster.PrintName", CxnPaperMovement, adOpenKeyset, adLockOptimistic
    rptPaperMovement.Text27.SetText "for " & Trim(rstPaperMVChild.Fields("GodownTo").Value)
    rptPaperMovement.Text9.SetText "for " & Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptPaperMovement.Database.SetDataSource rstPaperMVChild, 3, 1
    Dim TransfereeEMailID As String
    EMailID = rstPaperMVChild.Fields("TransferorEMailID").Value
    TransfereeEMailID = rstPaperMVChild.Fields("TransfereeEMailID").Value
    Attachment = Trim(rstPaperMVChild.Fields("VchNo").Value)
    Message = "Dear Sir,<Br>Pls find attached herewith " & IIf(MovementType = "1", "Book", "Title") & " Paper Movement Order #" & Trim(rstPaperMVChild.Fields("VchNo").Value) & " for doing the needful at your end. An early execution of the order will be highly appreciated.<Br>Kindly acknowledge the receipt of mail.<Br><Br>" & Trim(rstCompanyMaster.Fields("PrintName").Value) & "<Br>Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & "<Br>E-Mail : <a HRef='mailto:" & Trim(rstCompanyMaster.Fields("EMail").Value) & "'>" & Trim(rstCompanyMaster.Fields("EMail").Value) & "</a><Br><Br>CC: " & Trim(rstPaperMVChild.Fields("GodownTo").Value) & "-Kindly acknowledge the receipt of the Paper"
    Screen.MousePointer = vbNormal
    If OutputTo = "S" Then
        FrmReportViewer.EMailID = EMailID
        FrmReportViewer.CCID = TransfereeEMailID
        FrmReportViewer.Subject = IIf(MovementType = "1", "Book", "Title") & " Paper Movement Order #" & Trim(rstPaperMVChild.Fields("OrderNo").Value)
        FrmReportViewer.Attachment = Attachment
        FrmReportViewer.Message = Message
        Set FrmReportViewer.Report = rptPaperMovement
        FrmReportViewer.Show vbModal
    Else
        rptPaperMovement.PrintOut
    End If
    Set rptPaperMovement = Nothing
    On Error GoTo 0
End Sub
