VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmPrintPlanning 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Planning"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14850
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "PrintPlanning.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7275
   ScaleWidth      =   14850
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   7275
      Left            =   15
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   0
      Width           =   14820
      _Version        =   65536
      _ExtentX        =   26141
      _ExtentY        =   12832
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
      Picture         =   "PrintPlanning.frx":0442
      Begin TabDlg.SSTab SSTab1 
         Height          =   7035
         Left            =   120
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   120
         Width           =   14590
         _ExtentX        =   25744
         _ExtentY        =   12409
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
         TabPicture(0)   =   "PrintPlanning.frx":045E
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "DataGrid1"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Text1"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "&Details"
         TabPicture(1)   =   "PrintPlanning.frx":047A
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
            TabIndex        =   18
            Top             =   6555
            Width           =   7875
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   6045
            Left            =   120
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   450
            Width           =   8355
            _ExtentX        =   14737
            _ExtentY        =   10663
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
               Caption         =   "Voucher Date  "
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
                  ColumnWidth     =   1140.095
               EndProperty
               BeginProperty Column01 
                  Locked          =   -1  'True
                  ColumnWidth     =   1244.976
               EndProperty
               BeginProperty Column02 
                  Locked          =   -1  'True
                  ColumnWidth     =   5385.26
               EndProperty
            EndProperty
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   6435
            Left            =   -74880
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   480
            Width           =   14355
            _Version        =   65536
            _ExtentX        =   25321
            _ExtentY        =   11351
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
            Picture         =   "PrintPlanning.frx":0496
            Begin VB.TextBox MhRealInput7 
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
               Left            =   11250
               MaxLength       =   13
               TabIndex        =   11
               Text            =   "0"
               Top             =   1200
               Visible         =   0   'False
               Width           =   780
            End
            Begin VB.TextBox MhRealInput6 
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
               Left            =   10090
               MaxLength       =   13
               TabIndex        =   10
               Text            =   "0"
               Top             =   1200
               Visible         =   0   'False
               Width           =   1170
            End
            Begin VB.TextBox MhRealInput5 
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
               Left            =   9320
               MaxLength       =   13
               TabIndex        =   9
               Text            =   "0"
               Top             =   1200
               Visible         =   0   'False
               Width           =   800
            End
            Begin VB.TextBox Text11 
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
               Left            =   6180
               MaxLength       =   139
               TabIndex        =   8
               Top             =   1200
               Visible         =   0   'False
               Width           =   3150
            End
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
               Left            =   5250
               MaxLength       =   139
               TabIndex        =   7
               Top             =   1200
               Visible         =   0   'False
               Width           =   950
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
               Left            =   12950
               MaxLength       =   13
               TabIndex        =   13
               Text            =   "0.000"
               Top             =   1200
               Visible         =   0   'False
               Width           =   1045
            End
            Begin VB.TextBox MhRealInput3 
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
               Left            =   11985
               MaxLength       =   13
               TabIndex        =   12
               Text            =   "0.00"
               Top             =   1200
               Visible         =   0   'False
               Width           =   980
            End
            Begin VB.TextBox MhRealInput2 
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
               Left            =   4620
               MaxLength       =   13
               TabIndex        =   6
               Text            =   "0.00"
               Top             =   1200
               Visible         =   0   'False
               Width           =   640
            End
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
               Left            =   3780
               MaxLength       =   13
               TabIndex        =   5
               Text            =   "0"
               Top             =   1200
               Visible         =   0   'False
               Width           =   850
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
               Left            =   5535
               Locked          =   -1  'True
               MaxLength       =   40
               TabIndex        =   26
               TabStop         =   0   'False
               Top             =   6000
               Width           =   8685
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
               TabIndex        =   24
               TabStop         =   0   'False
               Top             =   6000
               Width           =   3375
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
               MaxLength       =   60
               TabIndex        =   4
               Top             =   1200
               Visible         =   0   'False
               Width           =   3360
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
               Left            =   1440
               MaxLength       =   139
               TabIndex        =   2
               Top             =   640
               Width           =   12795
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
               Height          =   330
               Left            =   120
               TabIndex        =   21
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
               Picture         =   "PrintPlanning.frx":04B2
               Picture         =   "PrintPlanning.frx":04CE
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Left            =   11355
               TabIndex        =   22
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
               Caption         =   " Voucher Date"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "PrintPlanning.frx":04EA
               Picture         =   "PrintPlanning.frx":0506
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
               Height          =   330
               Index           =   0
               Left            =   120
               TabIndex        =   23
               Top             =   645
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
               Picture         =   "PrintPlanning.frx":0522
               Picture         =   "PrintPlanning.frx":053E
            End
            Begin MSDataGridLib.DataGrid DataGrid2 
               Height          =   4605
               Left            =   120
               TabIndex        =   3
               Top             =   1200
               Width           =   14115
               _ExtentX        =   24897
               _ExtentY        =   8123
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
               ColumnCount     =   10
               BeginProperty Column00 
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
               BeginProperty Column01 
                  DataField       =   "Quantity"
                  Caption         =   " Quantity"
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
               BeginProperty Column02 
                  DataField       =   "Forms"
                  Caption         =   " Forms"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   1
                     Format          =   "0.00"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2057
                     SubFormatType   =   1
                  EndProperty
               EndProperty
               BeginProperty Column03 
                  DataField       =   "BookSize"
                  Caption         =   "Size"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   1
                     Format          =   "0.00"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2057
                     SubFormatType   =   1
                  EndProperty
               EndProperty
               BeginProperty Column04 
                  DataField       =   "Narration"
                  Caption         =   "Remarks"
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
               BeginProperty Column05 
                  DataField       =   "Warehouse1"
                  Caption         =   "Noida"
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
               BeginProperty Column06 
                  DataField       =   "Warehouse2"
                  Caption         =   "D.Daryaganj"
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
               BeginProperty Column07 
                  DataField       =   "Warehouse3"
                  Caption         =   "8 No"
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
               BeginProperty Column08 
                  DataField       =   "PaperWastage%"
                  Caption         =   "Wastage%"
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
                  DataField       =   "PaperConsumption"
                  Caption         =   "Consption"
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
               SplitCount      =   1
               BeginProperty Split0 
                  MarqueeStyle    =   3
                  ScrollBars      =   3
                  AllowRowSizing  =   0   'False
                  AllowSizing     =   0   'False
                  Locked          =   -1  'True
                  BeginProperty Column00 
                     Locked          =   -1  'True
                     ColumnWidth     =   3344.882
                  EndProperty
                  BeginProperty Column01 
                     Alignment       =   1
                     ColumnAllowSizing=   -1  'True
                     Locked          =   -1  'True
                     ColumnWidth     =   840.189
                  EndProperty
                  BeginProperty Column02 
                     Alignment       =   1
                     Locked          =   -1  'True
                     ColumnWidth     =   629.858
                  EndProperty
                  BeginProperty Column03 
                     Alignment       =   1
                     Locked          =   -1  'True
                     ColumnWidth     =   929.764
                  EndProperty
                  BeginProperty Column04 
                     ColumnWidth     =   3135.118
                  EndProperty
                  BeginProperty Column05 
                     ColumnWidth     =   780.095
                  EndProperty
                  BeginProperty Column06 
                     ColumnWidth     =   1154.835
                  EndProperty
                  BeginProperty Column07 
                     ColumnWidth     =   764.787
                  EndProperty
                  BeginProperty Column08 
                     ColumnWidth     =   929.764
                  EndProperty
                  BeginProperty Column09 
                     Alignment       =   1
                     Locked          =   -1  'True
                     ColumnWidth     =   1035.213
                  EndProperty
               EndProperty
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
               Height          =   330
               Left            =   120
               TabIndex        =   25
               Top             =   6000
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
               Caption         =   " Size"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "PrintPlanning.frx":055A
               Picture         =   "PrintPlanning.frx":0576
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Left            =   4800
               TabIndex        =   27
               Top             =   6000
               Width           =   795
               _Version        =   65536
               _ExtentX        =   1402
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
               Caption         =   " Board"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "PrintPlanning.frx":0592
               Picture         =   "PrintPlanning.frx":05AE
            End
            Begin TDBDate6Ctl.TDBDate MhDateInput1 
               Height          =   330
               Left            =   12795
               TabIndex        =   1
               Top             =   105
               Width           =   1445
               _Version        =   65536
               _ExtentX        =   2549
               _ExtentY        =   582
               Calendar        =   "PrintPlanning.frx":05CA
               Caption         =   "PrintPlanning.frx":06E2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PrintPlanning.frx":074E
               Keys            =   "PrintPlanning.frx":076C
               Spin            =   "PrintPlanning.frx":07CA
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
            Begin FPSpreadADO.fpSpread fpSpread1 
               Height          =   1335
               Left            =   6660
               TabIndex        =   28
               Top             =   4440
               Width           =   7575
               _Version        =   524288
               _ExtentX        =   13361
               _ExtentY        =   2355
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
               MaxCols         =   22
               MaxRows         =   3
               OperationMode   =   2
               SpreadDesigner  =   "PrintPlanning.frx":07F2
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInputTotalConsumption 
               Height          =   330
               Left            =   2520
               TabIndex        =   29
               TabStop         =   0   'False
               Top             =   5400
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
               _ExtentY        =   582
               Calculator      =   "PrintPlanning.frx":14FB
               Caption         =   "PrintPlanning.frx":151B
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PrintPlanning.frx":1587
               Keys            =   "PrintPlanning.frx":15A5
               Spin            =   "PrintPlanning.frx":15EF
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "###########0.000"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "###########0.000"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   999999999999.999
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   -1
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   1962541061
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInputComsumption 
               Height          =   330
               Left            =   1440
               TabIndex        =   30
               TabStop         =   0   'False
               Top             =   5400
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
               _ExtentY        =   582
               Calculator      =   "PrintPlanning.frx":1617
               Caption         =   "PrintPlanning.frx":1637
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PrintPlanning.frx":16A3
               Keys            =   "PrintPlanning.frx":16C1
               Spin            =   "PrintPlanning.frx":170B
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "###########0.000"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "###########0.000"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   999999999999.999
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   -1
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   1962541061
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInputWastage 
               Height          =   330
               Left            =   240
               TabIndex        =   31
               TabStop         =   0   'False
               Top             =   5400
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
               _ExtentY        =   582
               Calculator      =   "PrintPlanning.frx":1733
               Caption         =   "PrintPlanning.frx":1753
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PrintPlanning.frx":17BF
               Keys            =   "PrintPlanning.frx":17DD
               Spin            =   "PrintPlanning.frx":1827
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "###########0.000"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "###########0.000"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   999999999999.999
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   -1
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin VB.Line Line3 
               X1              =   0
               X2              =   14350
               Y1              =   5890
               Y2              =   5890
            End
            Begin VB.Line Line1 
               X1              =   0
               X2              =   14350
               Y1              =   525
               Y2              =   525
            End
            Begin VB.Line Line2 
               X1              =   0
               X2              =   14350
               Y1              =   1080
               Y2              =   1080
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
            TabIndex        =   19
            Top             =   6555
            Width           =   495
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   14850
      _ExtentX        =   26194
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
Attribute VB_Name = "FrmPrintPlanning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CxnPrintPlanning As New ADODB.Connection
Dim rstCompanyMaster As New ADODB.Recordset
Dim rstPrintPVList As New ADODB.Recordset
Dim rstPrintPVParent As New ADODB.Recordset
Dim WithEvents rstPrintPVChild As ADODB.Recordset
Attribute rstPrintPVChild.VB_VarHelpID = -1
Dim rstBookList As New ADODB.Recordset
Dim rstCheckRef As New ADODB.Recordset
Dim rstPrinterRates As New ADODB.Recordset
Dim BookCode As String
Dim Size_Code As String
Dim Acc_Code As String
Dim AddTitleEntry As String
Dim TitleEntryCode As String
Dim TitleEntryName As String
Dim PrevStr As String
Dim dblBookMark As Double
Dim blnRecordExist As Boolean
Dim OutputTo As String
Public PlanningType As String
Private Sub Form_Load()

    On Error GoTo ErrorHandler
    Me.Width = 8930
    Mh3dFrame1.Width = 8810
    SSTab1.Width = 8580
    CenterForm Me
    
    BusySystemIndicator True
    
    If PlanningType = "1" Then
        DataGrid2.Columns(0).Caption = "Book Name"
        Me.Caption = "Print Planning [Book]"
    Else
        DataGrid2.Columns(0).Caption = "Title Name"
        Me.Caption = "Print Planning [Title]"
    End If
           
    CxnPrintPlanning.CursorLocation = adUseClient
    CxnPrintPlanning.Open CxnDatabase.ConnectionString
    
    rstCompanyMaster.Open "Select PrintName, Address1, Address2, Address3, Address4, Phone, Fax, EMail, Website From CompanyMaster", CxnPrintPlanning, adOpenKeyset, adLockReadOnly
    rstBookList.Open "Select TRIM(M1.Name) As Col0,M3.Name As Col1,M2.Name As SizeName,M2.Code As SizeCode,FormType,Forms,Pages,OneColorPages,TwoColorPages,FourColorPages,OneColor¼Forms,OneColor½Forms,[OneColor1F/BForms],[OneColor1W/TForms],OneColorForms,TwoColor¼Forms,TwoColor½Forms,[TwoColor1F/BForms],[TwoColor1W/TForms],TwoColorForms,FourColor¼Forms,FourColor½Forms,[FourColor1F/BForms],[FourColor1W/TForms],FourColorForms,OneColorPlateType,TwoColorPlateType,FourColorPlateType,DuplexPrinting,BindingType,LaminationType,TitlePlateType,BindingForms01,BindingForms02,TitleFrontColor,TitleBackColor,TitlePlateType,[Qty/Pkt],[Pkt/Box],[LooseQty/Box],AddOnRate01,AddOnRate02,BookPrinter,TitlePrinter,Laminator,BinderFresh,BinderRepair,M1.Code From BookMaster M1,GeneralMaster M2,GeneralMaster M3 Where M1.[Size] = M2.Code AND M1.Board=M3.Code Order by M1.Name", CxnPrintPlanning, adOpenKeyset, adLockReadOnly
        
    rstPrintPVList.Open "Select PrintPVParent.Code, PrintPVParent.Name, Date, Particulars From PrintPVParent Where PlanningType = '" & PlanningType & "' Order By PrintPVParent.Name", CxnPrintPlanning, adOpenKeyset, adLockOptimistic
    rstPrintPVParent.CursorLocation = adUseClient
    
    Set rstPrintPVChild = New ADODB.Recordset
    If rstPrintPVList.RecordCount > 0 Then rstPrintPVList.MoveLast
    Set DataGrid1.DataSource = rstPrintPVList
    BusySystemIndicator False
    SSTab1.Tab = 0
    
    If Not (rstPrintPVList.EOF Or rstPrintPVList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    
    rstPrintPVList.ActiveConnection = Nothing
    rstBookList.ActiveConnection = Nothing
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
                If Me.ActiveControl.Name <> "Text5" And Me.ActiveControl.Name <> "MhRealInput1" And Me.ActiveControl.Name <> "MhRealInput2" And Me.ActiveControl.Name <> "MhRealInput3" And Me.ActiveControl.Name <> "MhRealInput4" Then
                    If MsgBox("Are you sure to Quit?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Quit !") <> vbYes Then
                        Me.ActiveControl.SetFocus
                    Else
                        Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
                    End If
                End If
            End If
            If Me.ActiveControl.Name <> "Text5" And Me.ActiveControl.Name <> "MhRealInput1" And Me.ActiveControl.Name <> "MhRealInput2" And Me.ActiveControl.Name <> "MhRealInput3" And Me.ActiveControl.Name <> "MhRealInput4" Then
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
        If Me.ActiveControl.Name <> "Text5" And Me.ActiveControl.Name <> "MhRealInput1" And Me.ActiveControl.Name <> "MhRealInput2" And Me.ActiveControl.Name <> "MhRealInput3" And Me.ActiveControl.Name <> "MhRealInput4" Then
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
            If Me.ActiveControl.Name <> "MhRealInput4" Then
                SendKeys "{TAB}"
            End If
        End If
        If Me.ActiveControl.Name <> "MhRealInput4" Then
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
    Call CloseRecordset(rstPrintPVList)
    Call CloseRecordset(rstPrintPVParent)
    Call CloseRecordset(rstPrintPVChild)
    Call CloseRecordset(rstBookList)
    Call CloseRecordset(rstCheckRef)
    Call CloseRecordset(rstPrinterRates)
    Call CloseConnection(CxnPrintPlanning)
    ShowProgressInStatusBar False
    DisableChildMenu
End Sub

Private Sub fpSpread1_DblClick(ByVal Col As Long, ByVal Row As Long)
        Dim Content As Variant
End Sub

Private Sub Text1_Change()
    If rstPrintPVList.RecordCount = 0 Then Exit Sub
    rstPrintPVList.MoveFirst
    If Text1.Text <> "" Then
        rstPrintPVList.Find "[Name] Like '%" & FixQuote(Text1.Text) & "%'"
        If rstPrintPVList.EOF Then
            rstPrintPVList.MoveFirst
            If PrevStr <> "" And Len(Text1.Text) > 1 Then
                If dblBookMark <> 0 Then
                    rstPrintPVList.Bookmark = dblBookMark
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
    If Not (rstPrintPVList.EOF Or rstPrintPVList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim KeyProcessed As Boolean
    If rstPrintPVList.RecordCount = 0 Then Exit Sub
    
    If Shift = 0 And KeyCode = vbKeyUp Then
        With rstPrintPVList
            .MovePrevious
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyBack Then
        With rstPrintPVList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyDown Then
        With rstPrintPVList
            .MoveNext
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageUp Then
        With rstPrintPVList
            .Move (-1) * (DataGrid1.VisibleRows - 1)
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageUp Then
        With rstPrintPVList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageDown Then
        With rstPrintPVList
            .Move DataGrid1.VisibleRows - 1
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageDown Then
        With rstPrintPVList
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
            Me.Width = 14940
            Mh3dFrame1.Width = 14820
            SSTab1.Width = 14590
            CenterForm Me
            ViewRecord
        Else
            Me.Width = 8930
            Mh3dFrame1.Width = 8810
            SSTab1.Width = 8580
            CenterForm Me
            If Not (rstPrintPVList.EOF Or rstPrintPVList.BOF) Then
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
        Me.Width = 14940
        Mh3dFrame1.Width = 14820
        SSTab1.Width = 14590
        CenterForm Me
        Text2.SetFocus
    End If
    
End Sub

Public Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim HiLiteRecord As Boolean
    Dim UpdateFlag As Integer
    
    If Button.Index = 1 Then
        If rstPrintPVParent.State = adStateOpen Then
           rstPrintPVParent.Close
        End If
        rstPrintPVParent.Open "Select * From PrintPVParent Where Code = ''", CxnPrintPlanning, adOpenKeyset, adLockOptimistic
        ClearFields ("P")
        ClearFields ("C")
        Call LoadBookList("")
        
        If rstPrintPVChild.State = adStateClosed Then
            SSTab1.Tab = 0
            Exit Sub
        End If
        
        If AddRecord(rstPrintPVParent) Then
            Text2.Text = GenerateCode(CxnPrintPlanning, "Select Max(Val(Name)) From PrintPVParent Where PlanningType = '" & PlanningType & "'", 10, Space(1))
            MhDateInput1.Text = Format(Date, "dd-MM-yyyy")
            Call SetButtons(False)
            SSTab1.Tab = 1
            Text2.SetFocus
            blnRecordExist = False
            CxnPrintPlanning.BeginTrans
        End If
        
    ElseIf Button.Index = 2 Then
        If rstPrintPVList.RecordCount = 0 Then Exit Sub
        SSTab1.Tab = 1
        EditRecord
    ElseIf Button.Index = 3 Then
        If rstPrintPVList.RecordCount = 0 Then Exit Sub
        If AllowTransactionsDeletion = 0 Then
            Call DisplayError("You don't have the rights to Delete this Voucher")
            Exit Sub
        End If
        SSTab1.Tab = 1
        
        If CheckRef Then
            DisplayError ("Failed to delete the record")
        ElseIf MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            On Error Resume Next
            MdiMainMenu.MousePointer = vbHourglass
            CxnPrintPlanning.Execute "Delete From PrintPVParent Where Code = '" & rstPrintPVList.Fields("Code").Value & "'"
            MdiMainMenu.MousePointer = vbNormal
            If Err.Number = 0 Then
                rstPrintPVList.Delete
                rstPrintPVList.MoveNext
                If rstPrintPVList.RecordCount > 0 And rstPrintPVList.EOF Then
                    rstPrintPVList.MoveLast
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
        AddTitleEntry = vbNo
        If PlanningType = "1" Then
            If CheckEmpty(rstPrintPVParent.Fields("Code").Value, False) Then
                AddTitleEntry = MsgBox("Do you wish to add entry for Title also ?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Title Entry !")
            End If
        End If
        SaveFields
        UpdateFlag = 0
        If UpdateRecord(rstPrintPVParent) Then
            UpdateFlag = 1
            If AddTitleEntry = vbYes Then If Not TitleEntry Then UpdateFlag = 0
            If UpdateFlag Then
                UpdateFlag = 0
                If UpdateBookList("D") Then
                     UpdateFlag = 1
                     Call DeleteBlankRecordset(rstPrintPVChild)
                     If rstPrintPVChild.RecordCount <> 0 Then
                        rstPrintPVChild.MoveFirst
                        Do While Not rstPrintPVChild.EOF
                            If Val(rstPrintPVChild.Fields("Quantity").Value) <> 0 Then
                                 If Not UpdateBookList("U") Then UpdateFlag = 0: Exit Do
                            End If
                            rstPrintPVChild.MoveNext
                        Loop
                          
                     End If
                End If
            End If
        End If
        If UpdateFlag Then
            AddToList
            CxnPrintPlanning.CommitTrans
            If rstPrintPVParent.State = adStateOpen Then
                rstPrintPVParent.Close
            End If
            rstPrintPVParent.CursorLocation = adUseClient
            Call SetButtons(True)
            SSTab1.Tab = 0
            ShowProgressInStatusBar True
            Timer1.Enabled = True
        Else
            DisplayError ("Failed to save the record")
            Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
        End If
    ElseIf Button.Index = 5 Then
        If CancelRecordUpdate(rstPrintPVParent) Then
            CxnPrintPlanning.RollbackTrans
            If rstPrintPVParent.State = adStateOpen Then
                rstPrintPVParent.Close
            End If
            rstPrintPVParent.CursorLocation = adUseClient
            Call SetButtons(True)
            SetButtonsForNoRecord
            SSTab1.Tab = 0
        End If
    ElseIf Button.Index = 6 Then
        SSTab1.Tab = 0
        Set DataGrid1.DataSource = Nothing
        rstPrintPVList.ActiveConnection = CxnPrintPlanning
        Do While Not RefreshRecord(rstPrintPVList)
        Loop
        Set DataGrid1.DataSource = rstPrintPVList
        rstPrintPVList.ActiveConnection = Nothing
        If rstPrintPVList.RecordCount > 0 Then rstPrintPVList.MoveLast
        rstBookList.ActiveConnection = CxnPrintPlanning
        Do While Not RefreshRecord(rstBookList)
        Loop
        rstBookList.ActiveConnection = Nothing
        HiLiteRecord = True
    ElseIf Button.Index = 7 Then
        SSTab1.Tab = 0
        HiLiteRecord = True
    ElseIf Button.Index = 9 Then
        
        If rstPrintPVList.RecordCount = 0 Then Exit Sub
        OutputTo = "P"
        PrintPrintPlanning
        HiLiteRecord = True
        
    ElseIf Button.Index = 10 Then
        If rstPrintPVList.RecordCount = 0 Then Exit Sub
        OutputTo = "S"
        PrintPrintPlanning
        HiLiteRecord = True
    ElseIf Button.Index = 13 Then
        If rstPrintPVList.RecordCount > 0 Then rstPrintPVList.MoveFirst
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 14 Then
        If rstPrintPVList.RecordCount > 0 Then
            rstPrintPVList.MovePrevious
            If rstPrintPVList.BOF Then
                rstPrintPVList.MoveNext
            End If
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 15 Then
        If rstPrintPVList.RecordCount > 0 Then
            rstPrintPVList.MoveNext
            If rstPrintPVList.EOF Then
                rstPrintPVList.MovePrevious
            End If
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 16 Then
        If rstPrintPVList.RecordCount > 0 Then rstPrintPVList.MoveLast
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 18 Then
        Unload Me
        HiLiteRecord = False
    End If
    If HiLiteRecord Then
        If Not (rstPrintPVList.EOF Or rstPrintPVList.BOF) Then
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
    If rstPrintPVList.RecordCount = 0 Then
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
    If rstPrintPVParent.EOF Or rstPrintPVParent.BOF Then Exit Sub
    If CheckEmpty(Text2, True) Then
        Cancel = True
    ElseIf CheckDuplicate(CxnPrintPlanning, "PrintPVParent", "Code", "[Name]+PlanningType", Trim(Text2.Text) & PlanningType, rstPrintPVParent.Fields("Code").Value, False) Then
        Cancel = True
    End If
End Sub
Private Sub MhDateInput1_Validate(Cancel As Boolean)
    If Not ValidateDate(Me.ActiveControl) Then
        Cancel = True
    ElseIf Format(GetDate(MhDateInput1.Text), "yyyymmdd") < Format(FinancialYearFrom, "yyyymmdd") Or Format(GetDate(MhDateInput1.Text), "yyyymmdd") > Format(FinancialYearTo, "yyyymmdd") Then
        Cancel = True
    End If
End Sub
Private Sub Text4_Validate(Cancel As Boolean)
    If rstPrintPVChild.RecordCount = 0 Then
        SendKeys "^"
        Call AddRecord(rstPrintPVChild)
        Call ClearFields("C")
        Call DataGrid2_KeyDown(vbKeyE, vbCtrlMask)
    End If
End Sub
Private Sub ViewRecord()
    ClearFields ("P")
    ClearFields ("C")
    If rstPrintPVList.EOF Then
        If rstPrintPVChild.State = adStateOpen Then
            rstPrintPVChild.Close
        End If
        Exit Sub
    End If
    
    FindRecord
    LoadFields
    
End Sub
Private Sub FindRecord()
    If rstPrintPVParent.State = adStateOpen Then
       rstPrintPVParent.Close
    End If
    rstPrintPVParent.Open "Select * From PrintPVParent Where Code = '" & FixQuote(rstPrintPVList.Fields("Code").Value) & "'", CxnPrintPlanning, adOpenKeyset, adLockOptimistic
    If rstPrintPVParent.RecordCount = 0 Then
       Call DisplayError("This Record has been deleted by Another User ! Click Ok To Refresh the Recordset")
       Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
    End If
End Sub
Private Sub ClearFields(ByVal strType As String)
    If strType = "P" Then
        Text2.Text = ""
        Text4.Text = ""
        MhDateInput1.Text = Format(Date, "dd-MM-yyyy")
    ElseIf strType = "C" Then
        Text5.Text = ""
        Text3.Text = ""
        Text6.Text = ""
        MhRealInput1.Text = "0"
        MhRealInput2.Text = "0.00"
        Text10.Text = ""
        Text11.Text = ""
        MhRealInput5.Text = ""
        MhRealInput6.Text = ""
        MhRealInput7.Text = ""
        MhRealInput4.Text = "0.000"
        
    End If
End Sub
Private Sub LoadFields()
    If rstPrintPVParent.EOF Or rstPrintPVParent.BOF Then Exit Sub
    Text2.Text = rstPrintPVParent.Fields("Name").Value
    MhDateInput1.Text = Format(rstPrintPVParent.Fields("Date").Value, "dd-MM-yyyy")
    Text4.Text = rstPrintPVParent.Fields("Remarks").Value
    Call LoadBookList(rstPrintPVParent.Fields("Code").Value)
End Sub
Private Sub EditRecord()
    On Error GoTo ErrorHandler
    If rstPrintPVParent.RecordCount = 0 Then Exit Sub
    If rstPrintPVChild.State = adStateClosed Then
        SSTab1.Tab = 0
        Exit Sub
    End If
    
    If rstPrintPVParent.State = adStateOpen Then
       rstPrintPVParent.Close
    End If
    
    rstPrintPVParent.CursorLocation = adUseServer
    rstPrintPVParent.Open "Select * From PrintPVParent Where Code = '" & FixQuote(rstPrintPVList.Fields("Code").Value) & "'", CxnPrintPlanning, adOpenKeyset, adLockPessimistic
    MdiMainMenu.MousePointer = vbHourglass
    rstPrintPVParent.Fields("Printstatus") = "N"
    MdiMainMenu.MousePointer = vbNormal
    AddToList
    Call SetButtons(False)
    SSTab1.TabEnabled(0) = False
    Text2.SetFocus
    blnRecordExist = True
    CxnPrintPlanning.BeginTrans
    Exit Sub
ErrorHandler:
    If Err.Number = -2147467259 Then
       Call DisplayError("Failed to Edit the record")
    End If
    MdiMainMenu.MousePointer = vbNormal
    SSTab1.Tab = 0
    
End Sub
Private Sub SaveFields()
    If rstPrintPVParent.EOF Or rstPrintPVParent.BOF Then Exit Sub
    If Not blnRecordExist Then
        rstPrintPVParent.Fields("Code").Value = GenerateCode(CxnPrintPlanning, "Select Max(Code) From PrintPVParent", 6, "0")
        rstPrintPVParent.Fields("CreatedBy").Value = UserCode
        rstPrintPVParent.Fields("CreatedOn").Value = Now()
        rstPrintPVParent.Fields("Recordstatus").Value = "N"
    Else
        rstPrintPVParent.Fields("ModifiedBy").Value = UserCode
        rstPrintPVParent.Fields("ModifiedOn").Value = Now()
        rstPrintPVParent.Fields("Recordstatus").Value = "M"
    End If
    rstPrintPVParent.Fields("Name").Value = Pad(Trim(Text2.Text), Space(1), 10, "L")
    rstPrintPVParent.Fields("Date").Value = GetDate(MhDateInput1.Text)
    rstPrintPVParent.Fields("PlanningType").Value = PlanningType
    rstPrintPVParent.Fields("Particulars").Value = "Planned " & Format(rstPrintPVChild.RecordCount, 0) & IIf(PlanningType = "1", " Book(s)", " Title(s)") & " For Printing"
    rstPrintPVParent.Fields("Remarks").Value = Trim(Text4.Text)
    rstPrintPVParent.Fields("PrintStatus").Value = "N"
End Sub
Private Sub AddToList()
    On Error Resume Next
    
    rstPrintPVList.MoveFirst
    rstPrintPVList.Find "[Code] = '" & rstPrintPVParent.Fields("Code").Value & "'"
    If rstPrintPVList.EOF Then
       rstPrintPVList.AddNew
       rstPrintPVList.Fields("Code").Value = rstPrintPVParent.Fields("Code").Value
    End If
    rstPrintPVList.Fields("Name").Value = Pad(rstPrintPVParent.Fields("Name").Value, Space(1), 10, "L")
    rstPrintPVList.Fields("Date").Value = rstPrintPVParent.Fields("Date").Value
    rstPrintPVList.Fields("Particulars").Value = Trim(rstPrintPVParent.Fields("Particulars").Value)
    rstPrintPVList.Update
    rstPrintPVList.Sort = "Name Asc"
    rstPrintPVList.Find "[Code] = '" & rstPrintPVParent.Fields("Code").Value & "'"
End Sub
Private Function CheckMandatoryFields() As Boolean
    If CheckEmpty(Text2.Text, False) Then
       DisplayError ("Voucher No. cannot be blank")
       Text2.SetFocus
       CheckMandatoryFields = True
    ElseIf CheckDuplicate(CxnPrintPlanning, "PrintPVParent", "Code", "[Name]+PlanningType", Trim(Text2.Text) & PlanningType, rstPrintPVParent.Fields("Code").Value, False) Then
        Text2.SetFocus
        CheckMandatoryFields = True
    End If
End Function
Private Function CheckRef() As Boolean
    On Error GoTo ErrorHandler
    
    If rstCheckRef.State = adStateOpen Then
         rstCheckRef.Close
    End If
    rstCheckRef.Open "Select Ref From " & IIf(PlanningType = "1", "BookPOChild05", "BookPOChild06") & " Where Ref = '" & rstPrintPVList.Fields("Code").Value & "'", CxnPrintPlanning, adOpenKeyset, adLockReadOnly
    If rstCheckRef.RecordCount > 0 Then
        CheckRef = True
    End If
    Exit Function
ErrorHandler:
    CheckRef = True
End Function
Private Sub Timer1_Timer()
    On Error Resume Next
    
    MdiMainMenu.ProgressBar1.Value = MdiMainMenu.ProgressBar1.Value + 10
    If MdiMainMenu.ProgressBar1.Value = 100 Then
       Timer1.Enabled = False
       ShowProgressInStatusBar False
    End If
End Sub
Private Sub LoadBookList(ByVal strVoucherCode As String)
    On Error GoTo ErrorHandler
    
    If rstPrintPVChild.State = adStateOpen Then
       rstPrintPVChild.Close
    End If
    
    rstPrintPVChild.Open "Select Book, M1.Name As BookName, M2.Name As SizeName, M3.Name As BoardName, Quantity, T.Forms, [PaperWastage%], PaperConsumption,T.BookSize,T.Narration,T.Warehouse1,T.Warehouse2,T.Warehouse3 From BookMaster M1, GeneralMaster M2, GeneralMaster M3, PrintPVChild T Where T.Book = M1.Code And M1.[Size] = M2.Code And M1.Board = M3.Code And T.Code = '" & strVoucherCode & "'", CxnPrintPlanning, adOpenKeyset, adLockOptimistic
        
    rstPrintPVChild.ActiveConnection = Nothing
    Set DataGrid2.DataSource = rstPrintPVChild
    Exit Sub
    
        
ErrorHandler:
    DisplayError ("Failed to Load " & IIf(PlanningType = "1", "Books", "Titles") & " List")
End Sub
Private Sub DataGrid2_DblClick()
    Call DataGrid2_KeyDown(vbKeyE, vbCtrlMask)
End Sub
Private Sub DataGrid2_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask And KeyCode = vbKeyE Then
        If rstPrintPVChild.RecordCount = 0 Then
            KeyCode = 0
            Exit Sub
        End If
        If Val(CheckNull(rstPrintPVChild.Fields("Quantity").Value)) <> 0 Then
            BookCode = rstPrintPVChild.Fields("Book").Value
            Text5.Text = rstPrintPVChild.Fields("BookName").Value
                       
            MhRealInput1.Text = Format(Val(rstPrintPVChild.Fields("Quantity").Value), "0")
            MhRealInput2.Text = Format(Val(rstPrintPVChild.Fields("Forms").Value), "0.00")
            MhRealInput3.Text = Format(Val(rstPrintPVChild.Fields("PaperWastage%").Value), "0.00")
            MhRealInput4.Text = Format(Val(rstPrintPVChild.Fields("PaperConsumption").Value), "0.000")
            If rstPrintPVChild.Fields("BookSize").Value <> "" Then
               Text10.Text = rstPrintPVChild.Fields("BookSize").Value
            End If
            If rstPrintPVChild.Fields("Narration").Value <> "" Then
               Text11.Text = rstPrintPVChild.Fields("Narration").Value
            End If
            If rstPrintPVChild.Fields("Warehouse1").Value <> "" Then
               MhRealInput5.Text = Format(Val(rstPrintPVChild.Fields("Warehouse1").Value), "0")
            End If
            If rstPrintPVChild.Fields("Warehouse2").Value <> "" Then
               MhRealInput6.Text = Format(Val(rstPrintPVChild.Fields("Warehouse2").Value), "0")
            End If
            If rstPrintPVChild.Fields("Warehouse3").Value <> "" Then
               MhRealInput7.Text = Format(Val(rstPrintPVChild.Fields("Warehouse3").Value), "0")
            End If
            
        End If
        With DataGrid2
            
            Text5.Visible = True
            Text5.Move .Left + .Columns(0).Left, .Top + .RowTop(.Row), .Columns(0).Width + 10, .RowHeight + 30
            MhRealInput1.Visible = True
            MhRealInput1.Move .Left + .Columns(1).Left, .Top + .RowTop(.Row), .Columns(1).Width + 10, .RowHeight + 30
            MhRealInput2.Visible = True
            MhRealInput2.Move .Left + .Columns(2).Left, .Top + .RowTop(.Row), .Columns(2).Width + 10, .RowHeight + 30
            Text10.Visible = True
            Text10.Move .Left + .Columns(3).Left, .Top + .RowTop(.Row), .Columns(3).Width + 10, .RowHeight + 30
            Text11.Visible = True
            Text11.Move .Left + .Columns(4).Left, .Top + .RowTop(.Row), .Columns(4).Width + 10, .RowHeight + 30
            
            MhRealInput5.Visible = True
            MhRealInput5.Move .Left + .Columns(5).Left, .Top + .RowTop(.Row), .Columns(5).Width + 10, .RowHeight + 30
            
            MhRealInput6.Visible = True
            MhRealInput6.Move .Left + .Columns(6).Left, .Top + .RowTop(.Row), .Columns(6).Width + 10, .RowHeight + 30
            
            MhRealInput7.Visible = True
            MhRealInput7.Move .Left + .Columns(7).Left, .Top + .RowTop(.Row), .Columns(7).Width + 10, .RowHeight + 30
            
            MhRealInput3.Visible = True
            MhRealInput3.Move .Left + .Columns(8).Left, .Top + .RowTop(.Row), .Columns(8).Width + 10, .RowHeight + 30
            MhRealInput4.Visible = True
            MhRealInput4.Move .Left + .Columns(9).Left, .Top + .RowTop(.Row), .Columns(9).Width + 10, .RowHeight + 30
            
            
        End With
        
        DataGrid2.Enabled = False
        Text5.SetFocus
        KeyCode = 0
        
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyA Then
        
        SendKeys "^"
        
        Call AddRecord(rstPrintPVChild)
        Call ClearFields("C")
        Call DataGrid2_KeyDown(vbKeyE, vbCtrlMask)
        
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyD Then
        
        If rstPrintPVChild.RecordCount = 0 Then Exit Sub
        
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            
            Set DataGrid2.DataSource = Nothing
            rstPrintPVChild.Delete
            rstPrintPVChild.MoveNext
            Set DataGrid2.DataSource = rstPrintPVChild
            DataGrid2.SetFocus
            
        End If
        If rstPrintPVChild.RecordCount = 0 Then
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
Private Sub rstPrintPVChild_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    
    On Error Resume Next
    
    If Not (rstPrintPVChild.EOF Or rstPrintPVChild.BOF) Then
        If Not IsNull(rstPrintPVChild.Fields("SizeName").Value) Then
            Text6.Text = rstPrintPVChild.Fields("SizeName").Value
        End If
        If Not IsNull(rstPrintPVChild.Fields("BoardName").Value) Then
            Text3.Text = rstPrintPVChild.Fields("BoardName").Value
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
    Dim SearchString As String, Pages As Variant
    SearchString = FixQuote(Text5.Text)
    
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
        Call LoadSelectionList(rstBookList, "List of Books...", "Name", "Board")
        SearchOrder = 0
        Call DisplaySelectionList(Text5, BookCode)
        Call CloseForm(FrmSelectionList)
        If CheckEmpty(Text5.Text, False) Then
            Text5.Text = "?"
        End If
        If RTrim(BookCode) <> "" Then
            SendKeys "{TAB}"
        End If
        Cancel = True
        Exit Sub
    ElseIf (rstPrintPVChild.Fields("BookName").Value <> Text5.Text) Or (CheckEmpty(rstPrintPVChild.Fields("BookName").Value, False)) Then
        If CheckDuplicateBook Then
            Call DisplayError("Duplicate Entry")
            Text5.SelStart = 0
            Text5.SelLength = Len(Text5.Text)
            Cancel = True
            Exit Sub
        End If
    End If
    BookCode = rstBookList.Fields("Code").Value
    Text3.Text = rstBookList.Fields("Col1").Value
    Text6.Text = rstBookList.Fields("SizeName").Value
    Size_Code = rstBookList.Fields("SizeCode").Value
    If Val(MhRealInput2.Text) = 0 Then
        MhRealInput2.Text = Format(Val(rstBookList.Fields("Forms").Value), "0.00")
    End If
    If Text10.Text = "" Then
        Text10.Text = rstBookList.Fields("SizeName").Value
    End If
    
    Dim Cnt As Integer
    For Cnt = 1 To fpSpread1.MaxRows
          
          fpSpread1.SetText 1, Cnt, Val(rstBookList.Fields(IIf(Cnt = 1, "One", IIf(Cnt = 2, "Two", "Four")) & "ColorPages").Value)
          fpSpread1.SetText 2, Cnt, Val(rstBookList.Fields(IIf(Cnt = 1, "One", IIf(Cnt = 2, "Two", "Four")) & "ColorForms").Value)
          fpSpread1.SetText 3, Cnt, Val(rstBookList.Fields(IIf(Cnt = 1, "One", IIf(Cnt = 2, "Two", "Four")) & "Color¼Forms").Value)
          fpSpread1.SetText 4, Cnt, Val(rstBookList.Fields(IIf(Cnt = 1, "One", IIf(Cnt = 2, "Two", "Four")) & "Color½Forms").Value)
          fpSpread1.SetText 5, Cnt, Val(rstBookList.Fields(IIf(Cnt = 1, "One", IIf(Cnt = 2, "Two", "Four")) & "Color1F/BForms").Value) + Val(rstBookList.Fields(IIf(Cnt = 1, "One", IIf(Cnt = 2, "Two", "Four")) & "Color1W/TForms").Value)
          fpSpread1.SetText 6, Cnt, IIf(rstBookList.Fields(IIf(Cnt = 1, "One", IIf(Cnt = 2, "Two", "Four")) & "ColorPlateType").Value = "1", "Deepatch", IIf(rstBookList.Fields(IIf(Cnt = 1, "One", IIf(Cnt = 2, "Two", "Four")) & "ColorPlateType").Value = "2", "PS", IIf(rstBookList.Fields(IIf(Cnt = 1, "One", IIf(Cnt = 2, "Two", "Four")) & "ColorPlateType").Value = "3", "Wipeon", "CTP")))
          fpSpread1.SetText 7, Cnt, 0#
          fpSpread1.SetText 8, Cnt, 0#
          fpSpread1.SetText 9, Cnt, 0#
          fpSpread1.SetText 10, Cnt, 0#
              
      Next
   For Cnt = 1 To fpSpread1.MaxRows
        fpSpread1.GetText 1, Cnt, Pages
        If Val(Pages) > 0 Then
            fpSpread1.SetActiveCell 1, Cnt
            fpSpread1_DblClick 1, Cnt
            Exit For
        End If
    Next
      
End Sub
Private Sub MhRealInput3_GotFocus()
    FocusSelect Me.ActiveControl
End Sub
Private Sub MhRealInput3_KeyPress(KeyAscii As Integer)
    ValidateKey MhRealInput3, KeyAscii, 2
End Sub
Private Sub MhRealInput3_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyEscape Then
        MakeTextBoxInvisible (True)
    End If
End Sub
Private Sub MhRealInput3_Validate(Cancel As Boolean)
    If Not ValidateNumber(Me.ActiveControl, 2) Then
        Cancel = True
    ElseIf Val(MhRealInput3.Text) < 0 Or Val(MhRealInput3.Text) > 99.99 Then
        Cancel = True
        MhRealInput3.SetFocus
        FocusSelect Me.ActiveControl
    End If
End Sub
Private Sub MhRealInput4_GotFocus()
    FocusSelect Me.ActiveControl
End Sub
Private Sub MhRealInput4_KeyPress(KeyAscii As Integer)
    ValidateKey MhRealInput4, KeyAscii, 3
End Sub

Private Sub MhRealInput4_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyReturn Then
        If Not ValidateNumber(Me.ActiveControl, 3) Then Exit Sub
          If (Val(MhRealInput5.Text) + Val(MhRealInput6.Text) + Val(MhRealInput7.Text) > Val(MhRealInput1.Text)) Then
                DisplayError ("Mat. Centre Qty should be less than Quantity")
                MhRealInput5.Visible = True
                MhRealInput6.Visible = True
                MhRealInput7.Visible = True
                MhRealInput5.SetFocus
                KeyCode = 0
                Exit Sub
          End If
        If Val(MhRealInput4.Text) > 0 Then
            rstPrintPVChild.Fields("Book").Value = BookCode
            rstPrintPVChild.Fields("BookName").Value = Trim(Text5.Text)
            rstPrintPVChild.Fields("SizeName").Value = Trim(Text6.Text)
            rstPrintPVChild.Fields("Quantity").Value = Format(Val(MhRealInput1.Text), "0")
            rstPrintPVChild.Fields("Forms").Value = Format(Val(MhRealInput2.Text), "0.00")
            rstPrintPVChild.Fields("PaperWastage%").Value = Format(Val(MhRealInput3.Text), "0.00")
            rstPrintPVChild.Fields("PaperConsumption").Value = Format(Val(MhRealInput4.Text), "0.000")
            rstPrintPVChild.Fields("BookSize").Value = Trim(Text10.Text)
            rstPrintPVChild.Fields("Narration").Value = Trim(Text11.Text)
            rstPrintPVChild.Fields("Warehouse1").Value = Format(Val(MhRealInput5.Text), "0")
            rstPrintPVChild.Fields("Warehouse2").Value = Format(Val(MhRealInput6.Text), "0")
            rstPrintPVChild.Fields("Warehouse3").Value = Format(Val(MhRealInput7.Text), "0")
                                   
            rstPrintPVChild.Update
            MakeTextBoxInvisible (False)
            If rstPrintPVChild.AbsolutePosition = rstPrintPVChild.RecordCount Then
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
        If Not (rstPrintPVChild.EOF Or rstPrintPVChild.BOF) Then
            If Val(CheckNull(rstPrintPVChild.Fields("Quantity").Value)) = 0 Then
                rstPrintPVChild.Delete
                rstPrintPVChild.MoveNext
                If rstPrintPVChild.RecordCount > 0 Then rstPrintPVChild.MoveFirst
            End If
        End If
    End If
    
    
    Text5.Visible = False
    MhRealInput1.Visible = False
    MhRealInput2.Visible = False
    MhRealInput3.Visible = False
    MhRealInput4.Visible = False
    MhRealInput5.Visible = False
    MhRealInput6.Visible = False
    MhRealInput7.Visible = False
    Text10.Visible = False
    Text11.Visible = False
    
    DataGrid2.Enabled = True
    DataGrid2.SetFocus
End Sub
Private Function CheckDuplicateBook() As Boolean
    
    Dim dblBookMark As Double
    
    If rstPrintPVChild.RecordCount = 0 Then Exit Function
    If Not (rstPrintPVChild.EOF Or rstPrintPVChild.BOF) Then
       dblBookMark = rstPrintPVChild.Bookmark
    End If
    rstPrintPVChild.MoveFirst
    Do While Not rstPrintPVChild.EOF
          If rstPrintPVChild.Fields("BookName").Value = Trim(Text5.Text) Then
             CheckDuplicateBook = True
             Exit Do
          End If
          rstPrintPVChild.MoveNext
    Loop
    If dblBookMark <> 0 Then
       rstPrintPVChild.Bookmark = dblBookMark
    Else
       rstPrintPVChild.MoveLast
    End If
End Function
Private Function UpdateBookList(ByVal strOption As String) As Boolean
    On Error GoTo ErrorHandler
    UpdateBookList = True
    If strOption = "D" Then
        CxnPrintPlanning.Execute "Delete From PrintPVChild Where Code = '" & rstPrintPVParent.Fields("Code").Value & "'"
    Else
        CxnPrintPlanning.Execute "Insert Into PrintPVChild Values ('" & rstPrintPVParent.Fields("Code").Value & "','" & rstPrintPVChild.Fields("Book").Value & "'," & Val(rstPrintPVChild.Fields("Quantity").Value) & "," & Val(rstPrintPVChild.Fields("Forms").Value) & "," & Val(rstPrintPVChild.Fields("PaperWastage%").Value) & "," & Val(rstPrintPVChild.Fields("PaperConsumption").Value) & ",'" & rstPrintPVChild.Fields("BookSize").Value & "','" & rstPrintPVChild.Fields("Narration").Value & "'," & IIf(rstPrintPVChild.Fields("Warehouse1").Value <> "", rstPrintPVChild.Fields("Warehouse1").Value, 0) & "," & IIf(rstPrintPVChild.Fields("Warehouse2").Value <> "", rstPrintPVChild.Fields("Warehouse2").Value, 0) & "," & IIf(rstPrintPVChild.Fields("Warehouse3").Value <> "", rstPrintPVChild.Fields("Warehouse3").Value, 0) & ")"
        If AddTitleEntry = vbYes Then
           Call CalculateAQD(Val(rstPrintPVChild.Fields("Quantity").Value), 0, 0, 0, rstPrintPVChild.Fields("Book").Value)
           MhRealInput3.Text = MhRealInputWastage.Value
           CxnPrintPlanning.Execute "Insert Into PrintPVChild Values ('" & TitleEntryCode & "','" & rstPrintPVChild.Fields("Book").Value & "'," & Val(rstPrintPVChild.Fields("Quantity").Value) & "," & Val(rstPrintPVChild.Fields("Forms").Value) & "," & Val(MhRealInputWastage.Value) & "," & CalculateConsumption2("2", Val(rstPrintPVChild.Fields("Quantity").Value), Val(rstPrintPVChild.Fields("Forms").Value), Val(MhRealInputWastage.Value)) & ",'" & rstPrintPVChild.Fields("BookSize").Value & "','" & rstPrintPVChild.Fields("Narration").Value & "'," & IIf(rstPrintPVChild.Fields("Warehouse1").Value <> "", rstPrintPVChild.Fields("Warehouse1").Value, 0) & "," & IIf(rstPrintPVChild.Fields("Warehouse2").Value <> "", rstPrintPVChild.Fields("Warehouse2").Value, 0) & "," & IIf(rstPrintPVChild.Fields("Warehouse3").Value <> "", rstPrintPVChild.Fields("Warehouse3").Value, 0) & ")"
        Else
           If PlanningType = "1" Then CxnPrintPlanning.Execute "UPDATE PrintPVParent P INNER JOIN PrintPVChild C ON P.Code=C.Code SET C.Quantity=" & Val(rstPrintPVChild.Fields("Quantity").Value) & " WHERE P.Name='" & rstPrintPVParent.Fields("Name").Value & "' AND C.Book='" & rstPrintPVChild.Fields("Book").Value & "' AND P.PlanningType='2'"
        End If
    End If
    Exit Function
ErrorHandler:
    UpdateBookList = False
End Function
Private Function TitleEntry() As Boolean
    On Error GoTo ErrorHandler
    TitleEntry = True
    TitleEntryCode = GenerateCode(CxnPrintPlanning, "Select Max(Code) From PrintPVParent", 6, "0")
    TitleEntryName = GenerateCode(CxnPrintPlanning, "Select Max(Val(Name)) From PrintPVParent Where PlanningType = '2'", 10, Space(1))
    CxnPrintPlanning.Execute "Insert Into PrintPVParent (Code, Name, [Date], PlanningType, Particulars, Remarks, CreatedBy) Values('" & TitleEntryCode & "','" & TitleEntryName & "',#" & GetDate(MhDateInput1.Text) & "#,'2','" & "Planned " & Format(rstPrintPVChild.RecordCount, 0) & " Title(s) For Printing','','" & UserCode & "')"
    Exit Function
ErrorHandler:
    TitleEntry = False
End Function
'******************New code**********************
Public Function CalculateConsumption2(ByVal xPaperType As String, ByVal xQuantity As Long, ByVal xForms As Double, ByVal xWastage As Double) As Double
      If xPaperType = "1" Then    'Book
        CalculateConsumption2 = CLng(xQuantity * xForms * (100 + xWastage) / 100)
    Else    'Title
        CalculateConsumption2 = Format((xQuantity / 2) * ((100 + xWastage) / 100), "#0")
    End If
    
    CalculateConsumption2 = CLng(Val(CalculateConsumption2) / 2)
    CalculateConsumption2 = Int(Val(CalculateConsumption2) / 500) & "." & Format(Val(CalculateConsumption2) Mod 500, "000")
End Function

Private Sub MhRealInput1_GotFocus()
    FocusSelect Me.ActiveControl
End Sub
Private Sub MhRealInput1_KeyPress(KeyAscii As Integer)
    ValidateKey MhRealInput1, KeyAscii, 0
End Sub
Private Sub MhRealInput1_Change()
'MhRealInput4.Text = Format(CalculateConsumption2(PlanningType, Val(MhRealInput1.Text), Val(MhRealInput2.Text), Val(MhRealInput3.Text)), "0.000")

Call CalculateAQD(Val(MhRealInput1.Text), 0, 0, Size_Code, BookCode)
MhRealInput3.Text = MhRealInputWastage.Value
MhRealInput4.Text = MhRealInputTotalConsumption.Value
        
End Sub
Private Sub MhRealInput1_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyEscape Then
        MakeTextBoxInvisible (True)
    End If
End Sub
Private Sub MhRealInput1_Validate(Cancel As Boolean)
    If Not ValidateNumber(Me.ActiveControl, 0) Then
        Cancel = True
    ElseIf Val(MhRealInput1.Text) <= 0 Then
        Cancel = True
        MhRealInput1.SetFocus
        FocusSelect Me.ActiveControl
    End If
'        Call CalculateAQD(Val(MhRealInput1.Text), 0, 0, Size_Code, BookCode)
'        MhRealInput3.Text = MhRealInputWastage.Value
'        MhRealInput4.Text = MhRealInputTotalConsumption.Value
End Sub
Private Sub MhRealInput2_GotFocus()
    FocusSelect Me.ActiveControl
End Sub
Private Sub MhRealInput2_KeyPress(KeyAscii As Integer)
    ValidateKey MhRealInput2, KeyAscii, 2
End Sub
Private Sub MhRealInput2_Change()
'    Dim Forms¼ As Variant, Forms½ As Variant, Forms1 As Variant
'    fpSpread1.SetText 5, fpSpread1.ActiveRow, Val(MhRealInput2.Text)
'    Call CalculateTotalPlates(IIf(fpSpread1.ActiveRow = 1, "1", IIf(fpSpread1.ActiveRow = 2, "2", "4")), Val(MhRealInput2.Text), "1")
'    Call CalculateTotalForms(IIf(fpSpread1.ActiveRow = 1, "1", IIf(fpSpread1.ActiveRow = 2, "2", "4")), Val(MhRealInput2.Text), "1", 0, 0)
'    CalculateAmount
'    Call CalculateConsumption(IIf(fpSpread1.ActiveRow = 1, "1", IIf(fpSpread1.ActiveRow = 2, "2", "4")), Val(MhRealInput1.Text))
'    fpSpread1.GetText 3, fpSpread1.ActiveRow, Forms¼
'    fpSpread1.GetText 4, fpSpread1.ActiveRow, Forms½
'    fpSpread1.GetText 5, fpSpread1.ActiveRow, Forms1
'    fpSpread1.SetText 2, fpSpread1.ActiveRow, Val(Forms¼) * 0.25 + Val(Forms½) * 0.5 + Val(Forms1) * 1
'    Call CalculateAQD(Val(MhRealInput1.Text), 0, 0, Size_Code, BookCode)
'    MhRealInput3.Text = MhRealInputWastage.Value
'    MhRealInput4.Text = MhRealInputTotalConsumption.Value
End Sub
Private Sub MhRealInput2_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyEscape Then
        MakeTextBoxInvisible (True)
    End If
End Sub
Private Sub MhRealInput2_Validate(Cancel As Boolean)
    If Not ValidateNumber(Me.ActiveControl, 2) Then
        Cancel = True
    ElseIf Val(MhRealInput2.Text) <= 0 Then
        Cancel = True
        MhRealInput2.SetFocus
        FocusSelect Me.ActiveControl
    End If
    Call CalculateAQD(Val(MhRealInput1.Text), 0, 0, Size_Code, BookCode)
       MhRealInput3.Text = MhRealInputWastage.Value
 End Sub


Private Function CalculateConsumption(ByVal xPrintingType As String, ByVal MhRealInput1 As Variant) As Double
    
    Dim Forms¼ As Variant, Forms½ As Variant, Forms1 As Variant, WastageRate As Variant, CurrentPaperConsumption As Variant, Cnt As Integer, FS As Variant
   
    fpSpread1.GetText 3, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), Forms¼
    fpSpread1.GetText 4, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), Forms½
    fpSpread1.GetText 5, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), Forms1
    fpSpread1.GetText 9, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), WastageRate
    
    CalculateConsumption = CLng(Val(MhRealInput1) * (Val(Forms¼) * 0.25 + Val(Forms½) * 0.5 + Val(Forms1) * 1) * ((100 + Val(WastageRate)) / 100))
    CalculateConsumption = CLng(Val(CalculateConsumption) / 2)
    fpSpread1.GetText 22, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), FS
    CalculateConsumption = IIf(Val(FS) = 0.5, 2, IIf(Val(FS) = 2, 0.5, 1)) * CalculateConsumption
    CalculateConsumption = Format(CLng(Int(Val(CalculateConsumption) / 500)) + ((Val(CalculateConsumption) Mod 500) / 1000), "0.000")
    fpSpread1.SetText 10, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), CalculateConsumption
    If fpSpread1.ActiveRow = IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)) Then
        MhRealInputComsumption.Text = Format(Val(CalculateConsumption), "0.000")
    End If
    For Cnt = 1 To fpSpread1.MaxRows
        fpSpread1.GetText 10, Cnt, CurrentPaperConsumption
        MhRealInputTotalConsumption.Text = Format(IIf(Cnt = 1, 0, Val(MhRealInputTotalConsumption.Text)) + CLng((Int(Val(CurrentPaperConsumption)) * 500) + ((Val(CurrentPaperConsumption) - Int(Val(CurrentPaperConsumption))) * 1000)), "0.000")
    Next
    MhRealInputTotalConsumption.Text = Format(CLng(Int(Val(MhRealInputTotalConsumption.Text) / 500)) + ((Val(MhRealInputTotalConsumption.Text) Mod 500) / 1000), "0.000")

End Function

Private Sub CalculateAmount()
    Dim Cnt As Integer, TotalPlates¼ As Variant, TotalPlates½ As Variant, TotalPlates1 As Variant, PlateRate As Variant, TotalForms¼ As Variant, TotalForms½ As Variant, TotalForms1 As Variant, PrintRate As Variant
    
    For Cnt = 1 To fpSpread1.MaxRows
        fpSpread1.GetText 11, Cnt, TotalPlates¼
        fpSpread1.GetText 12, Cnt, TotalPlates½
        fpSpread1.GetText 13, Cnt, TotalPlates1
        fpSpread1.GetText 14, Cnt, PlateRate
        fpSpread1.GetText 15, Cnt, TotalForms¼
        fpSpread1.GetText 16, Cnt, TotalForms½
        fpSpread1.GetText 17, Cnt, TotalForms1
        fpSpread1.GetText 18, Cnt, PrintRate
        fpSpread1.SetText 7, Cnt, IIf(Cnt = 1, 1, IIf(Cnt = 2, 2, 4)) * (Val(TotalPlates¼) + Val(TotalPlates½) + Val(TotalPlates1)) * Val(PlateRate)
        fpSpread1.SetText 8, Cnt, IIf(Cnt = 1, 1, IIf(Cnt = 2, 2, 4)) * (Val(TotalForms¼) + Val(TotalForms½) + Val(TotalForms1)) * Val(PrintRate)
        If fpSpread1.ActiveRow = Cnt Then
            'MhRealInput7.Text = Format(IIf(Cnt = 1, 1, IIf(Cnt = 2, 2, 4)) * (Val(TotalPlates¼) + Val(TotalPlates½) + Val(TotalPlates1)) * Val(PlateRate), "0.00")
            'MhRealInput8.Text = Format(IIf(Cnt = 1, 1, IIf(Cnt = 2, 2, 4)) * (Val(TotalForms¼) + Val(TotalForms½) + Val(TotalForms1)) * Val(PrintRate), "0.00")
        End If
    Next
    CalculateTotalAmount
End Sub
Private Function CalculateTotalAmount() As Double
    Dim Cnt As Integer, PlateAmount As Variant, PrintAmount As Variant, TotalAmount As Double
    For Cnt = 1 To fpSpread1.MaxRows
        fpSpread1.GetText 7, Cnt, PlateAmount
        fpSpread1.GetText 8, Cnt, PrintAmount
        TotalAmount = TotalAmount + PlateAmount + PrintAmount
    Next
End Function

Private Function CalculateTotalForms(ByVal xPrintingType As String, ByVal Forms As Double, ByVal FormType As String, ByVal MhRealInput2 As Variant, ByVal MhRealInput19 As Variant) As Double
    Dim FS As Variant
    fpSpread1.GetText 21, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), FS
    CalculateTotalForms = (Int(IIf(xPrintingType = "1", Val(0), Val(0)) * IIf(FormType = "¼", 0.25, IIf(FormType = "½", 0.5, 1)) / 1000) + IIf(IIf(xPrintingType = "1", Val(0), Val(0)) * IIf(FormType = "¼", 0.25, IIf(FormType = "½", 0.5, 1)) Mod 1000 = 0, 0, 1)) * Forms
    CalculateTotalForms = IIf(Val(FS) = 0.5, 2, IIf(Val(FS) = 2, 0.5, 1)) * Val(CalculateTotalForms)
    If rstBookList.Fields("DuplexPrinting").Value = "N" Then CalculateTotalForms = 0.5 * CalculateTotalForms
    CalculateTotalForms = Int(Val(CalculateTotalForms)) + IIf(Val(CalculateTotalForms) - Int(Val(CalculateTotalForms)) = 0, 0, 1)
    If FormType = "¼" Then
        fpSpread1.SetText 15, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), CalculateTotalForms
        If fpSpread1.ActiveRow = IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)) Then
        End If
    ElseIf FormType = "½" Then
        fpSpread1.SetText 16, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), CalculateTotalForms
        If fpSpread1.ActiveRow = IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)) Then
        End If
    Else
        fpSpread1.SetText 17, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), CalculateTotalForms
        If fpSpread1.ActiveRow = IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)) Then
        End If
    End If
End Function
Private Sub CalculateAQD(ByVal xMhRealInput1 As Variant, ByVal xMhRealInput2 As Variant, ByVal xMhRealInput19 As Variant, ByVal Size_Code As String, ByVal Ac_Code As String) 'Calculate Actual Quantity Dependents
    
    Dim Q1 As Double, Q24 As Double
    'For Single Color : Actual Quantity = Billing Quantity + 10 % of Billing Quantity + 99
    Q1 = Val(xMhRealInput1) * 100 / (10 + 100) Mod 1000
    Q1 = IIf(Val(xMhRealInput1) > 99 And Q1 > 0 And Int(Q1) <= 90, Val(xMhRealInput1) - 99, Val(xMhRealInput1))  'New Actual Quantity
    Q1 = Int(Q1 * 100 / (10 + 100) / 1000) * 1000 + IIf(Q1 * 100 / (10 + 100) Mod 1000 = 0, 0, 1000)
    'For Double/Four Color : Actual Quantity = Billing Quantity - 200 + 99 OR Actual Quantity = Billing Quantity - 500 + 99
    Q24 = IIf(Int(Val(xMhRealInput1) / 1000) = 0, 1000, Int(Val(xMhRealInput1) / 1000) * 1000) + IIf(Val(xMhRealInput1) Mod 1000 <= IIf(Val(xMhRealInput1) <= 10000, 299, 599), 0, 1000)
    If Val(xMhRealInput2) = 0 Then
        'MhRealInputBillingQty1.Text = Format(Q1, "0")
    ElseIf Val(xMhRealInput2) <> Q1 Then
        If MsgBox("Variation (Single Color) between Billing Quantity (" & xMhRealInput2 & ") Vs Calculated Billing Quantity (" & Trim(str(Q1)) & ") ! Change Quantity ?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then
            'MhRealInputBillingQty1.Text = Format(Q1, "0")
        End If
    End If
    If Val(xMhRealInput19) = 0 Then
        'MhRealInputBillingQty2.Text = Format(Q24, "0")
    ElseIf Val(xMhRealInput19) <> Q24 Then
        If MsgBox("Variation (Double & Four Color) between Billing Quantity (" & xMhRealInput19 & ") Vs Calculated Billing Quantity (" & Trim(str(Q24)) & ") ! Change Quantity ?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then
            'MhRealInputBillingQty2.Text = Format(Q24, "0")
        End If
    End If
    Call CalculateBQD("S", Format(Q1, "0"), xMhRealInput1, Size_Code, Ac_Code)
    Call CalculateBQD("O", Format(Q1, "0"), xMhRealInput1, Size_Code, Ac_Code)
    Call CalculateConsumption("1", xMhRealInput1): Call CalculateConsumption("2", xMhRealInput1): Call CalculateConsumption("4", xMhRealInput1)
End Sub

Private Sub CalculateBQD(ByVal xPrintingType As String, ByVal BillingQty As Variant, ByVal ActualQty As Variant, ByVal Size_Code As String, ByVal Acc_Code As String) 'Calculate Billing Quantity Dependents
    Dim Cnt As Integer, Content As Variant, Forms As Variant
    
    For Cnt = IIf(xPrintingType = "S", 1, 2) To IIf(xPrintingType = "S", 1, fpSpread1.MaxRows)
        fpSpread1.GetText 1, Cnt, Content   'Pages
        If Val(Content) <> 0 Then
            GetPrinterRates IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4")), "B", BillingQty, ActualQty, Size_Code, Acc_Code 'Get Both Plate & Printing Rates
        End If
        fpSpread1.GetText 3, Cnt, Forms
        Call CalculateTotalForms(IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4")), Val(Forms), "¼", BillingQty, ActualQty)
        fpSpread1.GetText 4, Cnt, Forms
        Call CalculateTotalForms(IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4")), Val(Forms), "½", BillingQty, ActualQty)
        fpSpread1.GetText 5, Cnt, Forms
        Call CalculateTotalForms(IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4")), Val(Forms), "1", BillingQty, ActualQty)
    Next
    CalculateAmount
End Sub

Private Sub GetPrinterRates(ByVal xPrintingType As String, ByVal xRateType As String, ByVal MhRealInput2 As Variant, ByVal MhRealInput19 As Variant, ByVal Size_Code As String, ByVal Acc_Code As String) 'xRateType : 'B'-Both Plate & Print Rate 'L'-Only Plate Rate
    
    Dim PrintRate As Double, PlateRate As Double, PaperWastageRate As Double, CurrentRate As Variant, PlateType As Variant
    
    On Error GoTo ErrorHandler

    If rstPrinterRates.State = adStateOpen Then rstPrinterRates.Close
    rstPrinterRates.Open "Select Top 1 * From AccountChild05 Where Code = '" & Acc_Code & "' And [Size] = '" & Size_Code & "' And Range" & Trim(xPrintingType) & " >= " & IIf(xPrintingType = "1", Val(MhRealInput2), Val(MhRealInput19)) & " Order By Range" & Trim(xPrintingType), CxnDatabase, adOpenKeyset, adLockReadOnly
    
    If rstPrinterRates.RecordCount = 0 Then
        If rstPrinterRates.State = adStateOpen Then rstPrinterRates.Close
        rstPrinterRates.Open "Select Top 1 * From AccountMaster,AccountChild05 Where AccountMaster.Code = AccountChild05.Code And [Name] Like '%Rate%' And [Size] = '" & Size_Code & "' And Range" & Trim(xPrintingType) & " >= " & IIf(xPrintingType = "1", Val(MhRealInput2), Val(MhRealInput19)) & " Order By Range" & Trim(xPrintingType), CxnDatabase, adOpenKeyset, adLockReadOnly
    End If
    If rstPrinterRates.RecordCount > 0 Then
        fpSpread1.GetText 6, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), PlateType
        PlateRate = rstPrinterRates.Fields(PlateType & "PlateRate" & Trim(xPrintingType)).Value
        PrintRate = rstPrinterRates.Fields("PrintRate" & Trim(xPrintingType)).Value
        PrintRate = PrintRate + IIf(PrintRate > 0, Val(rstBookList.Fields("AddOnRate01").Value), 0)
        PaperWastageRate = Val(rstPrinterRates.Fields("PaperWastageRate" & Trim(xPrintingType)))
        MhRealInputWastage.Text = Format(PaperWastageRate, "0.000")
    End If
    fpSpread1.GetText 14, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), CurrentRate  'Plate Rate
    If CurrentRate <> PlateRate Then
       fpSpread1.SetText 14, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), PlateRate
    End If
    If xRateType = "B" Then
        fpSpread1.GetText 18, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), CurrentRate  'Print Rate
        If CurrentRate <> PrintRate And CurrentRate <> 0 Then

        Else
            fpSpread1.SetText 18, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), PrintRate
        End If
        fpSpread1.GetText 9, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), CurrentRate   'Paper Wastage Rate
        If CurrentRate <> PaperWastageRate Then
           fpSpread1.SetText 9, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), PaperWastageRate
        End If
    End If
    If fpSpread1.ActiveRow = IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)) Then
        fpSpread1.GetText 14, fpSpread1.ActiveRow, CurrentRate  'Plate Rate
        fpSpread1.GetText 18, fpSpread1.ActiveRow, CurrentRate  'Print Rate
        fpSpread1.GetText 9, fpSpread1.ActiveRow, CurrentRate   'Paper Wastage Rate
    End If
    Exit Sub
ErrorHandler:
    DisplayError ("Failed to Fetch Printer Rates")
End Sub

Private Function CalculateTotalPlates(ByVal xPrintingType As String, ByVal Forms As Double, ByVal FormType As String) As Double
    Dim FS As Variant
    fpSpread1.GetText 21, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), FS
    CalculateTotalPlates = Forms
    CalculateTotalPlates = IIf(Val(FS) = 0.5, 2, IIf(Val(FS) = 2, 0.5, 1)) * Val(CalculateTotalPlates)
    If rstBookList.Fields("DuplexPrinting").Value = "N" Then CalculateTotalPlates = 0.5 * CalculateTotalPlates
    CalculateTotalPlates = Int(Val(CalculateTotalPlates)) + IIf(Val(CalculateTotalPlates) - Int(Val(CalculateTotalPlates)) = 0.5, 1, 0)
    If FormType = "¼" Then
        fpSpread1.SetText 11, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), CalculateTotalPlates
        If fpSpread1.ActiveRow = IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)) Then
            MhRealInput3.Text = Format(CalculateTotalPlates, "0")
        End If
    ElseIf FormType = "½" Then
        fpSpread1.SetText 12, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), CalculateTotalPlates
        If fpSpread1.ActiveRow = IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)) Then
            'MhRealInput23.Text = Format(CalculateTotalPlates, "0")
        End If
    Else
        fpSpread1.SetText 13, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), CalculateTotalPlates
        If fpSpread1.ActiveRow = IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)) Then
            'MhRealInput24.Text = Format(CalculateTotalPlates, "0")
        End If
    End If
End Function

Private Sub PrintPrintPlanning()
    
    On Error Resume Next
    
    Screen.MousePointer = vbHourglass
    rptPrintPlanning.Text1.SetText IIf(PlanningType = "1", "Book", "Title") & " Print Planning"
    rptPrintPlanning.Text2.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptPrintPlanning.Text3.SetText Trim(rstCompanyMaster.Fields("Address1").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address2").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address3").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address4").Value)
    If (Not CheckEmpty(rstCompanyMaster.Fields("Phone").Value, False)) And (Not CheckEmpty(rstCompanyMaster.Fields("Fax").Value, False)) Then
        rptPrintPlanning.Text24.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & Space(1) & "Fax : " & Trim(rstCompanyMaster.Fields("Fax").Value)
    ElseIf Not CheckEmpty(rstCompanyMaster.Fields("Fax").Value, False) Then
        rptPrintPlanning.Text24.SetText "Fax : " & Trim(rstCompanyMaster.Fields("Fax").Value)
    ElseIf Not CheckEmpty(rstCompanyMaster.Fields("Phone").Value, False) Then
        rptPrintPlanning.Text24.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value)
    Else
        rptPrintPlanning.Section5.Suppress = True
    End If
    
    If rstPrintPVChild.State = adStateOpen Then
       
        rstPrintPVChild.Close
        
    End If
       
   
    rstPrintPVChild.Open "Select Trim(PrintPVParent.Name) As VchNo,[Date] As VchDate,Trim(PrintName) As BookName,(Select Trim(PrintName) From GeneralMaster Where Code = BookMaster.Board) As BoardName,(Select Trim(PrintName) From GeneralMaster Where Code = BookMaster.[Size]) As SizeName,(Select Trim(PrintName) From AccountMaster Where Code=BookMaster.BookPrinter) As BookPrinter,(Select Trim(PrintName) From AccountMaster Where Code=BookMaster.TitlePrinter) As TitlePrinter,(Select Trim(PrintName) From AccountMaster Where Code=BookMaster.Laminator) As Laminator,(Select Trim(PrintName) From AccountMaster Where Code=BookMaster.BinderFresh) As Binder,Quantity,PrintPVChild.Forms,PaperConsumption " & _
                             "From (PrintPVParent Inner Join PrintPVChild On (PrintPVParent.Code = PrintPVChild.Code And PlanningType = '" & PlanningType & "' And PrintPVParent.Code = '" & rstPrintPVList.Fields("Code").Value & "')) Inner Join BookMaster On PrintPVChild.Book = BookMaster.Code Order By BookMaster.PrintName", CxnPrintPlanning, adOpenKeyset, adLockOptimistic
    rptPrintPlanning.Database.SetDataSource rstPrintPVChild, 3, 1

    Screen.MousePointer = vbNormal
    If OutputTo = "S" Then
        
        Set FrmReportViewer.Report = rptPrintPlanning
        FrmReportViewer.Show vbModal
        
    Else
        rptPrintPlanning.PrintOut
    End If
    Set rptPrintPlanning = Nothing
    On Error GoTo 0
    
End Sub

Private Sub DeleteBlankRecordset(ByVal rst As Recordset)
    rst.MoveFirst
    Do Until rst.EOF
       If rst.EOF Then Exit Do
       rst.MoveNext
       If IsNull(rst.Fields("BookName")) Then
          rst.Delete
          rst.MoveNext
       End If
    Loop
End Sub

