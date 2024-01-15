VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmAccountMaster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Account Master"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6750
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "AccountMaster.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6030
   ScaleWidth      =   6750
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   6025
      Left            =   15
      TabIndex        =   82
      TabStop         =   0   'False
      Top             =   0
      Width           =   6735
      _Version        =   65536
      _ExtentX        =   11880
      _ExtentY        =   10627
      _StockProps     =   77
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Picture         =   "AccountMaster.frx":0442
      Begin TabDlg.SSTab SSTab1 
         Height          =   5785
         Left            =   120
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   120
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   10213
         _Version        =   393216
         Style           =   1
         Tabs            =   11
         TabsPerRow      =   8
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
         TabPicture(0)   =   "AccountMaster.frx":045E
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "DataGrid1"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Text1"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "&Details"
         TabPicture(1)   =   "AccountMaster.frx":047A
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Mh3dFrame2(0)"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "&Details"
         TabPicture(2)   =   "AccountMaster.frx":0496
         Tab(2).ControlEnabled=   0   'False
         Tab(2).ControlCount=   0
         TabCaption(3)   =   "&Details"
         TabPicture(3)   =   "AccountMaster.frx":04B2
         Tab(3).ControlEnabled=   0   'False
         Tab(3).ControlCount=   0
         TabCaption(4)   =   "&Details"
         TabPicture(4)   =   "AccountMaster.frx":04CE
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "Mh3dFrame2(3)"
         Tab(4).ControlCount=   1
         TabCaption(5)   =   "&Details"
         TabPicture(5)   =   "AccountMaster.frx":04EA
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "Mh3dFrame2(4)"
         Tab(5).ControlCount=   1
         TabCaption(6)   =   "&Details"
         TabPicture(6)   =   "AccountMaster.frx":0506
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "Mh3dFrame2(5)"
         Tab(6).ControlCount=   1
         TabCaption(7)   =   "&Details"
         TabPicture(7)   =   "AccountMaster.frx":0522
         Tab(7).ControlEnabled=   0   'False
         Tab(7).Control(0)=   "Mh3dFrame2(6)"
         Tab(7).ControlCount=   1
         TabCaption(8)   =   "&Details"
         TabPicture(8)   =   "AccountMaster.frx":053E
         Tab(8).ControlEnabled=   0   'False
         Tab(8).Control(0)=   "Mh3dFrame2(7)"
         Tab(8).ControlCount=   1
         TabCaption(9)   =   "&Details"
         TabPicture(9)   =   "AccountMaster.frx":055A
         Tab(9).ControlEnabled=   0   'False
         Tab(9).Control(0)=   "Mh3dFrame2(8)"
         Tab(9).ControlCount=   1
         TabCaption(10)  =   "&Op.Bal."
         TabPicture(10)  =   "AccountMaster.frx":0576
         Tab(10).ControlEnabled=   0   'False
         Tab(10).Control(0)=   "Mh3dFrame2(9)"
         Tab(10).ControlCount=   1
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   600
            TabIndex        =   86
            Top             =   5310
            Width           =   5775
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   4790
            Left            =   120
            TabIndex        =   85
            TabStop         =   0   'False
            Top             =   450
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   8440
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
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   "Name"
               Caption         =   "Name"
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
               DataField       =   "Alias"
               Caption         =   "Alias"
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
                  Locked          =   -1  'True
                  ColumnWidth     =   4410.142
               EndProperty
               BeginProperty Column01 
                  Locked          =   -1  'True
                  ColumnWidth     =   1275.024
               EndProperty
            EndProperty
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   2735
            Index           =   0
            Left            =   -74880
            TabIndex        =   87
            TabStop         =   0   'False
            Top             =   480
            Width           =   6255
            _Version        =   65536
            _ExtentX        =   11033
            _ExtentY        =   4824
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
            Picture         =   "AccountMaster.frx":0592
            Begin VB.TextBox Text12 
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
               Index           =   0
               Left            =   4200
               MaxLength       =   40
               TabIndex        =   8
               Top             =   1980
               Width           =   1935
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
               Index           =   0
               Left            =   1440
               MaxLength       =   80
               TabIndex        =   9
               Top             =   2300
               Width           =   2055
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
               Index           =   0
               Left            =   4200
               MaxLength       =   40
               TabIndex        =   10
               Top             =   2300
               Width           =   1935
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
               Index           =   0
               Left            =   1440
               MaxLength       =   40
               TabIndex        =   7
               Top             =   1980
               Width           =   2055
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
               Index           =   0
               Left            =   1440
               MaxLength       =   40
               TabIndex        =   6
               Top             =   1670
               Width           =   4695
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
               Index           =   0
               Left            =   1440
               MaxLength       =   40
               TabIndex        =   5
               Top             =   1355
               Width           =   4695
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
               Index           =   0
               Left            =   1440
               MaxLength       =   40
               TabIndex        =   4
               Top             =   1040
               Width           =   4695
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
               Index           =   0
               Left            =   1440
               MaxLength       =   40
               TabIndex        =   3
               Top             =   725
               Width           =   4695
            End
            Begin VB.TextBox Text2 
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
               Index           =   0
               Left            =   1440
               MaxLength       =   40
               TabIndex        =   0
               Top             =   100
               Width           =   4695
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   330
               Index           =   0
               Left            =   120
               TabIndex        =   81
               Top             =   410
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
               Caption         =   " Print Name"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "AccountMaster.frx":05AE
               Picture         =   "AccountMaster.frx":05CA
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Index           =   0
               Left            =   120
               TabIndex        =   80
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
               Caption         =   " Name"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "AccountMaster.frx":05E6
               Picture         =   "AccountMaster.frx":0602
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
               Height          =   1270
               Index           =   0
               Left            =   120
               TabIndex        =   89
               Top             =   720
               Width           =   1335
               _Version        =   65536
               _ExtentX        =   2355
               _ExtentY        =   2240
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
               Caption         =   " Address"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "AccountMaster.frx":061E
               Picture         =   "AccountMaster.frx":063A
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
               Height          =   330
               Index           =   0
               Left            =   120
               TabIndex        =   90
               Top             =   1980
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
               Caption         =   " Phone"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "AccountMaster.frx":0656
               Picture         =   "AccountMaster.frx":0672
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
               Height          =   330
               Index           =   0
               Left            =   120
               TabIndex        =   91
               Top             =   2300
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
               Caption         =   " E-mail"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "AccountMaster.frx":068E
               Picture         =   "AccountMaster.frx":06AA
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
               Height          =   330
               Index           =   0
               Left            =   3480
               TabIndex        =   92
               Top             =   2295
               Width           =   735
               _Version        =   65536
               _ExtentX        =   1296
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
               Caption         =   " TIN No."
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "AccountMaster.frx":06C6
               Picture         =   "AccountMaster.frx":06E2
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Index           =   0
               Left            =   3480
               TabIndex        =   93
               Top             =   1980
               Width           =   735
               _Version        =   65536
               _ExtentX        =   1296
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
               Caption         =   " Mobile"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "AccountMaster.frx":06FE
               Picture         =   "AccountMaster.frx":071A
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Index           =   13
               Left            =   3480
               TabIndex        =   152
               Top             =   410
               Width           =   735
               _Version        =   65536
               _ExtentX        =   1296
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
               Picture         =   "AccountMaster.frx":0736
               Picture         =   "AccountMaster.frx":0752
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
               Index           =   0
               Left            =   1440
               MaxLength       =   40
               TabIndex        =   1
               Top             =   410
               Width           =   2055
            End
            Begin VB.TextBox Text13 
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
               Index           =   0
               Left            =   4200
               MaxLength       =   40
               TabIndex        =   2
               Top             =   410
               Width           =   1935
            End
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   5015
            Index           =   4
            Left            =   -74880
            TabIndex        =   94
            TabStop         =   0   'False
            Top             =   480
            Width           =   6255
            _Version        =   65536
            _ExtentX        =   11033
            _ExtentY        =   8846
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
            Picture         =   "AccountMaster.frx":076E
            Begin VB.TextBox Text2 
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
               Index           =   4
               Left            =   1440
               MaxLength       =   40
               TabIndex        =   23
               Top             =   100
               Width           =   4695
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
               Index           =   4
               Left            =   1440
               MaxLength       =   40
               TabIndex        =   24
               Top             =   410
               Width           =   2055
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
               Index           =   4
               Left            =   1440
               MaxLength       =   40
               TabIndex        =   26
               Top             =   725
               Width           =   4695
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
               Index           =   4
               Left            =   1440
               MaxLength       =   40
               TabIndex        =   27
               Top             =   1040
               Width           =   4695
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
               Index           =   4
               Left            =   1440
               MaxLength       =   40
               TabIndex        =   28
               Top             =   1355
               Width           =   4695
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
               Index           =   4
               Left            =   1440
               MaxLength       =   40
               TabIndex        =   29
               Top             =   1670
               Width           =   4695
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
               Index           =   4
               Left            =   1440
               MaxLength       =   40
               TabIndex        =   30
               Top             =   1980
               Width           =   2055
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
               Index           =   4
               Left            =   4200
               MaxLength       =   40
               TabIndex        =   33
               Top             =   2300
               Width           =   1935
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
               Index           =   4
               Left            =   1440
               MaxLength       =   80
               TabIndex        =   32
               Top             =   2300
               Width           =   2055
            End
            Begin VB.TextBox Text12 
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
               Index           =   4
               Left            =   4200
               MaxLength       =   40
               TabIndex        =   31
               Top             =   1980
               Width           =   1935
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   320
               Index           =   4
               Left            =   120
               TabIndex        =   95
               Top             =   420
               Width           =   1335
               _Version        =   65536
               _ExtentX        =   2355
               _ExtentY        =   564
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
               Caption         =   " Print Name"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "AccountMaster.frx":078A
               Picture         =   "AccountMaster.frx":07A6
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Index           =   4
               Left            =   120
               TabIndex        =   96
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
               Caption         =   " Name"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "AccountMaster.frx":07C2
               Picture         =   "AccountMaster.frx":07DE
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
               Height          =   1270
               Index           =   4
               Left            =   120
               TabIndex        =   97
               Top             =   720
               Width           =   1335
               _Version        =   65536
               _ExtentX        =   2355
               _ExtentY        =   2240
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
               Caption         =   " Address"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "AccountMaster.frx":07FA
               Picture         =   "AccountMaster.frx":0816
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
               Height          =   330
               Index           =   4
               Left            =   120
               TabIndex        =   98
               Top             =   1980
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
               Caption         =   " Phone"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "AccountMaster.frx":0832
               Picture         =   "AccountMaster.frx":084E
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
               Height          =   330
               Index           =   4
               Left            =   120
               TabIndex        =   99
               Top             =   2300
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
               Caption         =   " E-mail"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "AccountMaster.frx":086A
               Picture         =   "AccountMaster.frx":0886
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
               Height          =   330
               Index           =   4
               Left            =   3480
               TabIndex        =   100
               Top             =   2295
               Width           =   735
               _Version        =   65536
               _ExtentX        =   1296
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
               Caption         =   " TIN No."
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "AccountMaster.frx":08A2
               Picture         =   "AccountMaster.frx":08BE
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Index           =   4
               Left            =   3480
               TabIndex        =   101
               Top             =   1980
               Width           =   735
               _Version        =   65536
               _ExtentX        =   1296
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
               Caption         =   " Mobile"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "AccountMaster.frx":08DA
               Picture         =   "AccountMaster.frx":08F6
            End
            Begin MSDataGridLib.DataGrid DataGrid2 
               Height          =   2085
               Index           =   4
               Left            =   120
               TabIndex        =   34
               Top             =   2825
               Width           =   6015
               _ExtentX        =   10610
               _ExtentY        =   3678
               _Version        =   393216
               AllowUpdate     =   0   'False
               AllowArrows     =   -1  'True
               Appearance      =   0
               BackColor       =   16776960
               HeadLines       =   1
               RowHeight       =   18
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
               ColumnCount     =   5
               BeginProperty Column00 
                  DataField       =   "SizeName"
                  Caption         =   "Book Size"
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
                  DataField       =   "Range1"
                  Caption         =   "Range (1)"
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
                  DataField       =   "Range2"
                  Caption         =   "Range (2)"
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
               BeginProperty Column03 
                  DataField       =   "Range4"
                  Caption         =   "Range (4)"
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
               BeginProperty Column04 
                  DataField       =   "Range6"
                  Caption         =   "Range (6)"
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
                     Locked          =   -1  'True
                     ColumnWidth     =   2204.788
                  EndProperty
                  BeginProperty Column01 
                     Alignment       =   1
                     Locked          =   -1  'True
                     ColumnWidth     =   810.142
                  EndProperty
                  BeginProperty Column02 
                     Alignment       =   1
                     Locked          =   -1  'True
                     ColumnWidth     =   810.142
                  EndProperty
                  BeginProperty Column03 
                     Alignment       =   1
                     Locked          =   -1  'True
                     ColumnWidth     =   810.142
                  EndProperty
                  BeginProperty Column04 
                     Alignment       =   1
                     Locked          =   -1  'True
                     ColumnWidth     =   810.142
                  EndProperty
               EndProperty
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Index           =   11
               Left            =   3480
               TabIndex        =   150
               Top             =   410
               Width           =   735
               _Version        =   65536
               _ExtentX        =   1296
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
               Picture         =   "AccountMaster.frx":0912
               Picture         =   "AccountMaster.frx":092E
            End
            Begin VB.TextBox Text13 
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
               Index           =   4
               Left            =   4200
               MaxLength       =   40
               TabIndex        =   25
               Top             =   410
               Width           =   1935
            End
            Begin VB.Line Line4 
               X1              =   0
               X2              =   6240
               Y1              =   2720
               Y2              =   2720
            End
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   5015
            Index           =   5
            Left            =   -74880
            TabIndex        =   102
            TabStop         =   0   'False
            Top             =   480
            Width           =   6255
            _Version        =   65536
            _ExtentX        =   11033
            _ExtentY        =   8846
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
            Picture         =   "AccountMaster.frx":094A
            Begin VB.TextBox Text2 
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
               Index           =   5
               Left            =   1440
               MaxLength       =   40
               TabIndex        =   35
               Top             =   100
               Width           =   4695
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
               Index           =   5
               Left            =   1440
               MaxLength       =   40
               TabIndex        =   36
               Top             =   410
               Width           =   2055
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
               Index           =   5
               Left            =   1440
               MaxLength       =   40
               TabIndex        =   38
               Top             =   725
               Width           =   4695
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
               Index           =   5
               Left            =   1440
               MaxLength       =   40
               TabIndex        =   39
               Top             =   1040
               Width           =   4695
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
               Index           =   5
               Left            =   1440
               MaxLength       =   40
               TabIndex        =   40
               Top             =   1355
               Width           =   4695
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
               Index           =   5
               Left            =   1440
               MaxLength       =   40
               TabIndex        =   41
               Top             =   1670
               Width           =   4695
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
               Index           =   5
               Left            =   1440
               MaxLength       =   40
               TabIndex        =   42
               Top             =   1980
               Width           =   2055
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
               Index           =   5
               Left            =   4200
               MaxLength       =   40
               TabIndex        =   45
               Top             =   2300
               Width           =   1935
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
               Index           =   5
               Left            =   1440
               MaxLength       =   80
               TabIndex        =   44
               Top             =   2295
               Width           =   2055
            End
            Begin VB.TextBox Text12 
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
               Index           =   5
               Left            =   4200
               MaxLength       =   40
               TabIndex        =   43
               Top             =   1980
               Width           =   1935
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   320
               Index           =   5
               Left            =   120
               TabIndex        =   103
               Top             =   420
               Width           =   1335
               _Version        =   65536
               _ExtentX        =   2355
               _ExtentY        =   564
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
               Caption         =   " Print Name"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "AccountMaster.frx":0966
               Picture         =   "AccountMaster.frx":0982
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Index           =   5
               Left            =   120
               TabIndex        =   104
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
               Caption         =   " Name"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "AccountMaster.frx":099E
               Picture         =   "AccountMaster.frx":09BA
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
               Height          =   1270
               Index           =   5
               Left            =   120
               TabIndex        =   105
               Top             =   720
               Width           =   1335
               _Version        =   65536
               _ExtentX        =   2355
               _ExtentY        =   2240
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
               Caption         =   " Address"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "AccountMaster.frx":09D6
               Picture         =   "AccountMaster.frx":09F2
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
               Height          =   330
               Index           =   5
               Left            =   120
               TabIndex        =   106
               Top             =   1980
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
               Caption         =   " Phone"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "AccountMaster.frx":0A0E
               Picture         =   "AccountMaster.frx":0A2A
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
               Height          =   330
               Index           =   5
               Left            =   120
               TabIndex        =   107
               Top             =   2300
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
               Caption         =   " E-mail"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "AccountMaster.frx":0A46
               Picture         =   "AccountMaster.frx":0A62
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
               Height          =   330
               Index           =   5
               Left            =   3480
               TabIndex        =   108
               Top             =   2295
               Width           =   735
               _Version        =   65536
               _ExtentX        =   1296
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
               Caption         =   " TIN No."
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "AccountMaster.frx":0A7E
               Picture         =   "AccountMaster.frx":0A9A
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Index           =   5
               Left            =   3480
               TabIndex        =   109
               Top             =   1980
               Width           =   735
               _Version        =   65536
               _ExtentX        =   1296
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
               Caption         =   " Mobile"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "AccountMaster.frx":0AB6
               Picture         =   "AccountMaster.frx":0AD2
            End
            Begin MSDataGridLib.DataGrid DataGrid2 
               Height          =   2085
               Index           =   5
               Left            =   120
               TabIndex        =   135
               Top             =   2825
               Width           =   6015
               _ExtentX        =   10610
               _ExtentY        =   3678
               _Version        =   393216
               AllowUpdate     =   0   'False
               AllowArrows     =   -1  'True
               Appearance      =   0
               BackColor       =   16776960
               HeadLines       =   1
               RowHeight       =   18
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
               ColumnCount     =   5
               BeginProperty Column00 
                  DataField       =   "SizeName"
                  Caption         =   "Book Size"
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
                  DataField       =   "Range1"
                  Caption         =   "Range (1)"
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
                  DataField       =   "Range2"
                  Caption         =   "Range (2)"
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
               BeginProperty Column03 
                  DataField       =   "Range4"
                  Caption         =   "Range (4)"
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
               BeginProperty Column04 
                  DataField       =   "Range6"
                  Caption         =   "Range (6)"
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
                     Locked          =   -1  'True
                     ColumnWidth     =   2204.788
                  EndProperty
                  BeginProperty Column01 
                     Alignment       =   1
                     Locked          =   -1  'True
                     ColumnWidth     =   810.142
                  EndProperty
                  BeginProperty Column02 
                     Alignment       =   1
                     Locked          =   -1  'True
                     ColumnWidth     =   810.142
                  EndProperty
                  BeginProperty Column03 
                     Alignment       =   1
                     Locked          =   -1  'True
                     ColumnWidth     =   810.142
                  EndProperty
                  BeginProperty Column04 
                     Alignment       =   1
                     Locked          =   -1  'True
                     ColumnWidth     =   810.142
                  EndProperty
               EndProperty
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Index           =   10
               Left            =   3480
               TabIndex        =   149
               Top             =   420
               Width           =   735
               _Version        =   65536
               _ExtentX        =   1296
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
               Picture         =   "AccountMaster.frx":0AEE
               Picture         =   "AccountMaster.frx":0B0A
            End
            Begin VB.TextBox Text13 
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
               Index           =   5
               Left            =   4200
               MaxLength       =   40
               TabIndex        =   37
               Top             =   410
               Width           =   1935
            End
            Begin VB.Line Line5 
               X1              =   0
               X2              =   6240
               Y1              =   2720
               Y2              =   2720
            End
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   5015
            Index           =   6
            Left            =   -74880
            TabIndex        =   110
            TabStop         =   0   'False
            Top             =   480
            Width           =   6255
            _Version        =   65536
            _ExtentX        =   11033
            _ExtentY        =   8846
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
            Picture         =   "AccountMaster.frx":0B26
            Begin VB.TextBox Text2 
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
               Index           =   6
               Left            =   1440
               MaxLength       =   40
               TabIndex        =   46
               Top             =   100
               Width           =   4695
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
               Index           =   6
               Left            =   1440
               MaxLength       =   40
               TabIndex        =   47
               Top             =   410
               Width           =   2055
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
               Index           =   6
               Left            =   1440
               MaxLength       =   40
               TabIndex        =   49
               Top             =   725
               Width           =   4695
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
               Index           =   6
               Left            =   1440
               MaxLength       =   40
               TabIndex        =   50
               Top             =   1040
               Width           =   4695
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
               Index           =   6
               Left            =   1440
               MaxLength       =   40
               TabIndex        =   51
               Top             =   1355
               Width           =   4695
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
               Index           =   6
               Left            =   1440
               MaxLength       =   40
               TabIndex        =   52
               Top             =   1670
               Width           =   4695
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
               Index           =   6
               Left            =   1440
               MaxLength       =   40
               TabIndex        =   53
               Top             =   1980
               Width           =   2055
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
               Index           =   6
               Left            =   4200
               MaxLength       =   40
               TabIndex        =   56
               Top             =   2300
               Width           =   1935
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
               Index           =   6
               Left            =   1440
               MaxLength       =   80
               TabIndex        =   55
               Top             =   2300
               Width           =   2055
            End
            Begin VB.TextBox Text12 
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
               Index           =   6
               Left            =   4200
               MaxLength       =   40
               TabIndex        =   54
               Top             =   1980
               Width           =   1935
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   320
               Index           =   6
               Left            =   120
               TabIndex        =   111
               Top             =   420
               Width           =   1335
               _Version        =   65536
               _ExtentX        =   2355
               _ExtentY        =   564
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
               Caption         =   " Print Name"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "AccountMaster.frx":0B42
               Picture         =   "AccountMaster.frx":0B5E
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Index           =   6
               Left            =   120
               TabIndex        =   112
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
               Caption         =   " Name"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "AccountMaster.frx":0B7A
               Picture         =   "AccountMaster.frx":0B96
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
               Height          =   1270
               Index           =   6
               Left            =   120
               TabIndex        =   113
               Top             =   720
               Width           =   1335
               _Version        =   65536
               _ExtentX        =   2355
               _ExtentY        =   2240
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
               Caption         =   " Address"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "AccountMaster.frx":0BB2
               Picture         =   "AccountMaster.frx":0BCE
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
               Height          =   330
               Index           =   6
               Left            =   120
               TabIndex        =   114
               Top             =   1980
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
               Caption         =   " Phone"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "AccountMaster.frx":0BEA
               Picture         =   "AccountMaster.frx":0C06
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
               Height          =   330
               Index           =   6
               Left            =   120
               TabIndex        =   115
               Top             =   2300
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
               Caption         =   " E-mail"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "AccountMaster.frx":0C22
               Picture         =   "AccountMaster.frx":0C3E
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
               Height          =   330
               Index           =   6
               Left            =   3480
               TabIndex        =   116
               Top             =   2295
               Width           =   735
               _Version        =   65536
               _ExtentX        =   1296
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
               Caption         =   " TIN No."
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "AccountMaster.frx":0C5A
               Picture         =   "AccountMaster.frx":0C76
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Index           =   6
               Left            =   3480
               TabIndex        =   117
               Top             =   1980
               Width           =   735
               _Version        =   65536
               _ExtentX        =   1296
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
               Caption         =   " Mobile"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "AccountMaster.frx":0C92
               Picture         =   "AccountMaster.frx":0CAE
            End
            Begin MSDataGridLib.DataGrid DataGrid2 
               Height          =   2085
               Index           =   6
               Left            =   120
               TabIndex        =   134
               Top             =   2825
               Width           =   6015
               _ExtentX        =   10610
               _ExtentY        =   3678
               _Version        =   393216
               AllowUpdate     =   0   'False
               AllowArrows     =   -1  'True
               Appearance      =   0
               BackColor       =   16776960
               Enabled         =   -1  'True
               HeadLines       =   1
               RowHeight       =   18
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
               ColumnCount     =   9
               BeginProperty Column00 
                  DataField       =   "SizeName"
                  Caption         =   "Book Size"
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
                  DataField       =   "LaminationTypeName"
                  Caption         =   "Lamination Type"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   "0.00"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column02 
                  DataField       =   "Rate04"
                  Caption         =   "   Rate (04)"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   1
                     Format          =   "0.00"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   1
                  EndProperty
               EndProperty
               BeginProperty Column03 
                  DataField       =   "Rate08"
                  Caption         =   "   Rate (08)"
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
                  DataField       =   "Rate12"
                  Caption         =   "   Rate (12)"
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
               BeginProperty Column05 
                  DataField       =   "Rate16"
                  Caption         =   "   Rate (16)"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   1
                     Format          =   "0.00"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   1
                  EndProperty
               EndProperty
               BeginProperty Column06 
                  DataField       =   "Rate24"
                  Caption         =   "   Rate (24)"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   1
                     Format          =   "0.00"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   1
                  EndProperty
               EndProperty
               BeginProperty Column07 
                  DataField       =   "Rate32"
                  Caption         =   "   Rate (32)"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   1
                     Format          =   "0.00"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   1
                  EndProperty
               EndProperty
               BeginProperty Column08 
                  DataField       =   "Rate64"
                  Caption         =   "   Rate (64)"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   1
                     Format          =   "0.00"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
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
                     ColumnWidth     =   1200.189
                  EndProperty
                  BeginProperty Column01 
                     Locked          =   -1  'True
                     ColumnWidth     =   1395.213
                  EndProperty
                  BeginProperty Column02 
                     Alignment       =   1
                     Locked          =   -1  'True
                     ColumnWidth     =   945.071
                  EndProperty
                  BeginProperty Column03 
                     Alignment       =   1
                     Locked          =   -1  'True
                     ColumnWidth     =   959.811
                  EndProperty
                  BeginProperty Column04 
                     Alignment       =   1
                     Locked          =   -1  'True
                     ColumnWidth     =   945.071
                  EndProperty
                  BeginProperty Column05 
                     ColumnWidth     =   945.071
                  EndProperty
                  BeginProperty Column06 
                     ColumnWidth     =   945.071
                  EndProperty
                  BeginProperty Column07 
                     ColumnWidth     =   945.071
                  EndProperty
                  BeginProperty Column08 
                     ColumnWidth     =   945.071
                  EndProperty
               EndProperty
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Index           =   9
               Left            =   3480
               TabIndex        =   148
               Top             =   420
               Width           =   735
               _Version        =   65536
               _ExtentX        =   1296
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
               Picture         =   "AccountMaster.frx":0CCA
               Picture         =   "AccountMaster.frx":0CE6
            End
            Begin VB.TextBox Text13 
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
               Index           =   6
               Left            =   4200
               MaxLength       =   40
               TabIndex        =   48
               Top             =   420
               Width           =   1935
            End
            Begin VB.Line Line6 
               X1              =   0
               X2              =   6240
               Y1              =   2720
               Y2              =   2720
            End
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   5150
            Index           =   7
            Left            =   -74880
            TabIndex        =   118
            TabStop         =   0   'False
            Top             =   480
            Width           =   6255
            _Version        =   65536
            _ExtentX        =   11033
            _ExtentY        =   9084
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
            Picture         =   "AccountMaster.frx":0D02
            Begin VB.TextBox Text2 
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
               Index           =   7
               Left            =   1440
               MaxLength       =   40
               TabIndex        =   57
               Top             =   100
               Width           =   4695
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
               Index           =   7
               Left            =   1440
               MaxLength       =   40
               TabIndex        =   58
               Top             =   410
               Width           =   2055
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
               Index           =   7
               Left            =   1440
               MaxLength       =   40
               TabIndex        =   60
               Top             =   725
               Width           =   4695
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
               Index           =   7
               Left            =   1440
               MaxLength       =   40
               TabIndex        =   61
               Top             =   1040
               Width           =   4695
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
               Index           =   7
               Left            =   1440
               MaxLength       =   40
               TabIndex        =   62
               Top             =   1355
               Width           =   4695
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
               Index           =   7
               Left            =   1440
               MaxLength       =   40
               TabIndex        =   63
               Top             =   1670
               Width           =   4695
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
               Index           =   7
               Left            =   1440
               MaxLength       =   40
               TabIndex        =   64
               Top             =   1980
               Width           =   2055
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
               Index           =   7
               Left            =   4200
               MaxLength       =   40
               TabIndex        =   67
               Top             =   2300
               Width           =   1935
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
               Index           =   7
               Left            =   1440
               MaxLength       =   80
               TabIndex        =   66
               Top             =   2300
               Width           =   2055
            End
            Begin VB.TextBox Text12 
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
               Index           =   7
               Left            =   4200
               MaxLength       =   40
               TabIndex        =   65
               Top             =   1980
               Width           =   1935
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   320
               Index           =   7
               Left            =   120
               TabIndex        =   119
               Top             =   420
               Width           =   1335
               _Version        =   65536
               _ExtentX        =   2355
               _ExtentY        =   564
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
               Caption         =   " Print Name"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "AccountMaster.frx":0D1E
               Picture         =   "AccountMaster.frx":0D3A
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Index           =   7
               Left            =   120
               TabIndex        =   120
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
               Caption         =   " Name"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "AccountMaster.frx":0D56
               Picture         =   "AccountMaster.frx":0D72
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
               Height          =   1270
               Index           =   7
               Left            =   120
               TabIndex        =   121
               Top             =   720
               Width           =   1335
               _Version        =   65536
               _ExtentX        =   2355
               _ExtentY        =   2240
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
               Caption         =   " Address"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "AccountMaster.frx":0D8E
               Picture         =   "AccountMaster.frx":0DAA
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
               Height          =   330
               Index           =   7
               Left            =   120
               TabIndex        =   122
               Top             =   1980
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
               Caption         =   " Phone"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "AccountMaster.frx":0DC6
               Picture         =   "AccountMaster.frx":0DE2
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
               Height          =   330
               Index           =   7
               Left            =   120
               TabIndex        =   123
               Top             =   2300
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
               Caption         =   " E-mail"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "AccountMaster.frx":0DFE
               Picture         =   "AccountMaster.frx":0E1A
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
               Height          =   330
               Index           =   7
               Left            =   3480
               TabIndex        =   124
               Top             =   2295
               Width           =   735
               _Version        =   65536
               _ExtentX        =   1296
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
               Caption         =   " TIN No."
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "AccountMaster.frx":0E36
               Picture         =   "AccountMaster.frx":0E52
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Index           =   7
               Left            =   3480
               TabIndex        =   125
               Top             =   1980
               Width           =   735
               _Version        =   65536
               _ExtentX        =   1296
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
               Caption         =   " Mobile"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "AccountMaster.frx":0E6E
               Picture         =   "AccountMaster.frx":0E8A
            End
            Begin MSDataGridLib.DataGrid DataGrid2 
               Height          =   2215
               Index           =   7
               Left            =   120
               TabIndex        =   68
               Top             =   2825
               Width           =   6015
               _ExtentX        =   10610
               _ExtentY        =   3916
               _Version        =   393216
               AllowUpdate     =   0   'False
               AllowArrows     =   -1  'True
               Appearance      =   0
               BackColor       =   16776960
               HeadLines       =   1
               RowHeight       =   18
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
               ColumnCount     =   9
               BeginProperty Column00 
                  DataField       =   "SizeName"
                  Caption         =   "Book Size"
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
                  DataField       =   "BindingTypeName"
                  Caption         =   "Binding Type"
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
               BeginProperty Column02 
                  DataField       =   "Range04"
                  Caption         =   "Range (04)"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   1
                     Format          =   "0.00"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   1
                  EndProperty
               EndProperty
               BeginProperty Column03 
                  DataField       =   "Range08"
                  Caption         =   "Range (08)"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   1
                     Format          =   "0.00"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   1
                  EndProperty
               EndProperty
               BeginProperty Column04 
                  DataField       =   "Range12"
                  Caption         =   "Range (12)"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   1
                     Format          =   "0.00"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   1
                  EndProperty
               EndProperty
               BeginProperty Column05 
                  DataField       =   "Range16"
                  Caption         =   "Range (16)"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   1
                     Format          =   "0.00"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   1
                  EndProperty
               EndProperty
               BeginProperty Column06 
                  DataField       =   "Range24"
                  Caption         =   "Range (24)"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   1
                     Format          =   "0.00"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   1
                  EndProperty
               EndProperty
               BeginProperty Column07 
                  DataField       =   "Range32"
                  Caption         =   "Range (32)"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   1
                     Format          =   "0.00"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   1
                  EndProperty
               EndProperty
               BeginProperty Column08 
                  DataField       =   "Range64"
                  Caption         =   "Range (64)"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   1
                     Format          =   "0.00"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
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
                     ColumnWidth     =   989.858
                  EndProperty
                  BeginProperty Column01 
                     Locked          =   -1  'True
                     ColumnWidth     =   1709.858
                  EndProperty
                  BeginProperty Column02 
                     Alignment       =   1
                     Locked          =   -1  'True
                     ColumnWidth     =   915.024
                  EndProperty
                  BeginProperty Column03 
                     Alignment       =   1
                     Locked          =   -1  'True
                     ColumnWidth     =   915.024
                  EndProperty
                  BeginProperty Column04 
                     Alignment       =   1
                     Locked          =   -1  'True
                     ColumnWidth     =   915.024
                  EndProperty
                  BeginProperty Column05 
                     ColumnWidth     =   929.764
                  EndProperty
                  BeginProperty Column06 
                     ColumnWidth     =   929.764
                  EndProperty
                  BeginProperty Column07 
                     ColumnWidth     =   929.764
                  EndProperty
                  BeginProperty Column08 
                     ColumnWidth     =   929.764
                  EndProperty
               EndProperty
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Index           =   3
               Left            =   3480
               TabIndex        =   147
               Top             =   420
               Width           =   735
               _Version        =   65536
               _ExtentX        =   1296
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
               Picture         =   "AccountMaster.frx":0EA6
               Picture         =   "AccountMaster.frx":0EC2
            End
            Begin VB.TextBox Text13 
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
               Index           =   7
               Left            =   4200
               MaxLength       =   40
               TabIndex        =   59
               Top             =   410
               Width           =   1935
            End
            Begin VB.Line Line7 
               X1              =   0
               X2              =   6240
               Y1              =   2720
               Y2              =   2720
            End
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   2735
            Index           =   8
            Left            =   -74880
            TabIndex        =   126
            TabStop         =   0   'False
            Top             =   480
            Width           =   6255
            _Version        =   65536
            _ExtentX        =   11033
            _ExtentY        =   4824
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
            Picture         =   "AccountMaster.frx":0EDE
            Begin VB.TextBox Text2 
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
               Index           =   8
               Left            =   1440
               MaxLength       =   40
               TabIndex        =   69
               Top             =   100
               Width           =   4695
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
               Index           =   8
               Left            =   1440
               MaxLength       =   40
               TabIndex        =   70
               Top             =   410
               Width           =   2055
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
               Index           =   8
               Left            =   1440
               MaxLength       =   40
               TabIndex        =   72
               Top             =   725
               Width           =   4695
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
               Index           =   8
               Left            =   1440
               MaxLength       =   40
               TabIndex        =   73
               Top             =   1040
               Width           =   4695
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
               Index           =   8
               Left            =   1440
               MaxLength       =   40
               TabIndex        =   74
               Top             =   1355
               Width           =   4695
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
               Index           =   8
               Left            =   1440
               MaxLength       =   40
               TabIndex        =   75
               Top             =   1670
               Width           =   4695
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
               Index           =   8
               Left            =   1440
               MaxLength       =   40
               TabIndex        =   76
               Top             =   1980
               Width           =   2055
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
               Index           =   8
               Left            =   4200
               MaxLength       =   40
               TabIndex        =   79
               Top             =   2300
               Width           =   1935
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
               Index           =   8
               Left            =   1440
               MaxLength       =   80
               TabIndex        =   78
               Top             =   2300
               Width           =   2055
            End
            Begin VB.TextBox Text12 
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
               Index           =   8
               Left            =   4200
               MaxLength       =   40
               TabIndex        =   77
               Top             =   1980
               Width           =   1935
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   320
               Index           =   8
               Left            =   120
               TabIndex        =   127
               Top             =   420
               Width           =   1335
               _Version        =   65536
               _ExtentX        =   2355
               _ExtentY        =   564
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
               Caption         =   " Print Name"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "AccountMaster.frx":0EFA
               Picture         =   "AccountMaster.frx":0F16
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Index           =   8
               Left            =   120
               TabIndex        =   128
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
               Caption         =   " Name"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "AccountMaster.frx":0F32
               Picture         =   "AccountMaster.frx":0F4E
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
               Height          =   1270
               Index           =   8
               Left            =   120
               TabIndex        =   129
               Top             =   720
               Width           =   1335
               _Version        =   65536
               _ExtentX        =   2355
               _ExtentY        =   2240
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
               Caption         =   " Address"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "AccountMaster.frx":0F6A
               Picture         =   "AccountMaster.frx":0F86
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
               Height          =   330
               Index           =   8
               Left            =   120
               TabIndex        =   130
               Top             =   1980
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
               Caption         =   " Phone"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "AccountMaster.frx":0FA2
               Picture         =   "AccountMaster.frx":0FBE
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
               Height          =   330
               Index           =   8
               Left            =   120
               TabIndex        =   131
               Top             =   2300
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
               Caption         =   " E-mail"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "AccountMaster.frx":0FDA
               Picture         =   "AccountMaster.frx":0FF6
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
               Height          =   330
               Index           =   8
               Left            =   3480
               TabIndex        =   132
               Top             =   2295
               Width           =   735
               _Version        =   65536
               _ExtentX        =   1296
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
               Caption         =   " TIN No."
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "AccountMaster.frx":1012
               Picture         =   "AccountMaster.frx":102E
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Index           =   8
               Left            =   3480
               TabIndex        =   133
               Top             =   1980
               Width           =   735
               _Version        =   65536
               _ExtentX        =   1296
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
               Caption         =   " Mobile"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "AccountMaster.frx":104A
               Picture         =   "AccountMaster.frx":1066
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Index           =   2
               Left            =   3480
               TabIndex        =   146
               Top             =   420
               Width           =   735
               _Version        =   65536
               _ExtentX        =   1296
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
               Picture         =   "AccountMaster.frx":1082
               Picture         =   "AccountMaster.frx":109E
            End
            Begin VB.TextBox Text13 
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
               Index           =   8
               Left            =   4200
               MaxLength       =   40
               TabIndex        =   71
               Top             =   410
               Width           =   1935
            End
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   5010
            Index           =   3
            Left            =   -74880
            TabIndex        =   136
            TabStop         =   0   'False
            Top             =   480
            Width           =   6255
            _Version        =   65536
            _ExtentX        =   11033
            _ExtentY        =   8846
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
            Picture         =   "AccountMaster.frx":10BA
            Begin VB.TextBox Text2 
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
               Index           =   3
               Left            =   1440
               MaxLength       =   40
               TabIndex        =   11
               Top             =   100
               Width           =   4695
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
               Index           =   3
               Left            =   1440
               MaxLength       =   40
               TabIndex        =   12
               Top             =   410
               Width           =   2055
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
               Index           =   3
               Left            =   1440
               MaxLength       =   40
               TabIndex        =   14
               Top             =   725
               Width           =   4695
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
               Index           =   3
               Left            =   1440
               MaxLength       =   40
               TabIndex        =   15
               Top             =   1040
               Width           =   4695
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
               Index           =   3
               Left            =   1440
               MaxLength       =   40
               TabIndex        =   16
               Top             =   1355
               Width           =   4695
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
               Index           =   3
               Left            =   1440
               MaxLength       =   40
               TabIndex        =   17
               Top             =   1670
               Width           =   4695
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
               Index           =   3
               Left            =   1440
               MaxLength       =   40
               TabIndex        =   18
               Top             =   1980
               Width           =   2055
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
               Index           =   3
               Left            =   4200
               MaxLength       =   40
               TabIndex        =   21
               Top             =   2300
               Width           =   1935
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
               Index           =   3
               Left            =   1440
               MaxLength       =   80
               TabIndex        =   20
               Top             =   2300
               Width           =   2055
            End
            Begin VB.TextBox Text12 
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
               Index           =   3
               Left            =   4200
               MaxLength       =   40
               TabIndex        =   19
               Top             =   1980
               Width           =   1935
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   320
               Index           =   1
               Left            =   120
               TabIndex        =   137
               Top             =   420
               Width           =   1335
               _Version        =   65536
               _ExtentX        =   2355
               _ExtentY        =   564
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
               Caption         =   " Print Name"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "AccountMaster.frx":10D6
               Picture         =   "AccountMaster.frx":10F2
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Index           =   1
               Left            =   120
               TabIndex        =   138
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
               Caption         =   " Name"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "AccountMaster.frx":110E
               Picture         =   "AccountMaster.frx":112A
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
               Height          =   1270
               Index           =   1
               Left            =   120
               TabIndex        =   139
               Top             =   720
               Width           =   1335
               _Version        =   65536
               _ExtentX        =   2355
               _ExtentY        =   2240
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
               Caption         =   " Address"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "AccountMaster.frx":1146
               Picture         =   "AccountMaster.frx":1162
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
               Height          =   330
               Index           =   1
               Left            =   120
               TabIndex        =   140
               Top             =   1980
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
               Caption         =   " Phone"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "AccountMaster.frx":117E
               Picture         =   "AccountMaster.frx":119A
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
               Height          =   330
               Index           =   1
               Left            =   120
               TabIndex        =   141
               Top             =   2300
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
               Caption         =   " E-mail"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "AccountMaster.frx":11B6
               Picture         =   "AccountMaster.frx":11D2
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
               Height          =   330
               Index           =   1
               Left            =   3480
               TabIndex        =   142
               Top             =   2295
               Width           =   735
               _Version        =   65536
               _ExtentX        =   1296
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
               Caption         =   " TIN No."
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "AccountMaster.frx":11EE
               Picture         =   "AccountMaster.frx":120A
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Index           =   1
               Left            =   3480
               TabIndex        =   143
               Top             =   1980
               Width           =   735
               _Version        =   65536
               _ExtentX        =   1296
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
               Caption         =   " Mobile"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "AccountMaster.frx":1226
               Picture         =   "AccountMaster.frx":1242
            End
            Begin MSDataGridLib.DataGrid DataGrid2 
               Height          =   2085
               Index           =   3
               Left            =   120
               TabIndex        =   22
               Top             =   2825
               Width           =   6015
               _ExtentX        =   10610
               _ExtentY        =   3678
               _Version        =   393216
               AllowUpdate     =   0   'False
               AllowArrows     =   -1  'True
               Appearance      =   0
               BackColor       =   16776960
               Enabled         =   -1  'True
               HeadLines       =   1
               RowHeight       =   18
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
               ColumnCount     =   5
               BeginProperty Column00 
                  DataField       =   "SizeName"
                  Caption         =   "Size Name"
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
                  DataField       =   "OnePieceRate"
                  Caption         =   "One Piece Rate"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   1
                     Format          =   "0.00"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   1
                  EndProperty
               EndProperty
               BeginProperty Column02 
                  DataField       =   "CutPieceRate"
                  Caption         =   "Cut Piece Rate"
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
                  DataField       =   "PastingRate"
                  Caption         =   "Pasting Rate"
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
                  DataField       =   "Rate/Inch²"
                  Caption         =   "Rate/Inch²"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   1
                     Format          =   "0.00"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
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
                     ColumnWidth     =   1874.835
                  EndProperty
                  BeginProperty Column01 
                     Alignment       =   1
                     Locked          =   -1  'True
                     ColumnWidth     =   1289.764
                  EndProperty
                  BeginProperty Column02 
                     Alignment       =   1
                     Locked          =   -1  'True
                     ColumnWidth     =   1230.236
                  EndProperty
                  BeginProperty Column03 
                     Alignment       =   1
                     Locked          =   -1  'True
                     ColumnWidth     =   1049.953
                  EndProperty
                  BeginProperty Column04 
                     Alignment       =   1
                     Locked          =   -1  'True
                     ColumnWidth     =   929.764
                  EndProperty
               EndProperty
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Index           =   12
               Left            =   3480
               TabIndex        =   151
               Top             =   420
               Width           =   735
               _Version        =   65536
               _ExtentX        =   1296
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
               Picture         =   "AccountMaster.frx":125E
               Picture         =   "AccountMaster.frx":127A
            End
            Begin VB.TextBox Text13 
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
               Index           =   3
               Left            =   4200
               MaxLength       =   40
               TabIndex        =   13
               Top             =   410
               Width           =   1935
            End
            Begin VB.Line Line1 
               X1              =   0
               X2              =   6240
               Y1              =   2720
               Y2              =   2720
            End
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   5060
            Index           =   9
            Left            =   -74880
            TabIndex        =   144
            TabStop         =   0   'False
            Top             =   480
            Width           =   6255
            _Version        =   65536
            _ExtentX        =   11033
            _ExtentY        =   8925
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
            Picture         =   "AccountMaster.frx":1296
            Begin FPSpreadADO.fpSpread fpSpread1 
               Height          =   4845
               Left            =   120
               TabIndex        =   145
               Top             =   105
               Width           =   6015
               _Version        =   524288
               _ExtentX        =   10610
               _ExtentY        =   8546
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
               MaxRows         =   100
               OperationMode   =   2
               SpreadDesigner  =   "AccountMaster.frx":12B2
            End
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00808000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Find"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   120
            TabIndex        =   88
            Top             =   5310
            Width           =   495
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   83
      Top             =   0
      Width           =   6750
      _ExtentX        =   11906
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
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Print"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Print Preview"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
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
Attribute VB_Name = "FrmAccountMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public AccountType As String
Dim CxnAccountMaster As New ADODB.Connection
Dim rstAccountList As New ADODB.Recordset
Dim rstAccountMaster As New ADODB.Recordset
Dim rstSizeList As New ADODB.Recordset
Dim rstBindingTypeList As New ADODB.Recordset
Dim rstLaminationTypeList As New ADODB.Recordset
Dim rstOutsourceItemList As New ADODB.Recordset
Dim rstPaperList As New ADODB.Recordset
Dim rstFreshBookList As New ADODB.Recordset
Dim rstRepairBookList As New ADODB.Recordset
Dim rstAccountChild As New ADODB.Recordset
Dim rstCheckRef As New ADODB.Recordset
Dim PrevStr As String
Dim dblBookMark As Double
Dim blnRecordExist As Boolean
Dim SizeCode As String
Dim BindingTypeCode As String
Dim LaminationTypeCode As String
Dim OutsourceItem As String
Dim Paper As String
Dim FreshBook As String
Dim RepairBook As String
Dim Title As String
Dim SortOrder As String
Dim EditMode As Boolean
Private Sub Form_Load()
    Dim Cnt As Integer
    On Error GoTo ErrorHandler
    CenterForm Me
    BusySystemIndicator True
    For Cnt = 1 To 9
        If Cnt <> Val(AccountType) Then SSTab1.TabVisible(Cnt) = False
    Next
    If AccountType <> "08" Then SSTab1.TabVisible(10) = False
    CxnAccountMaster.CursorLocation = adUseClient
    CxnAccountMaster.Open CxnDatabase.ConnectionString
    
    rstAccountList.Open "Select Name,Alias,Code From AccountMaster Where Type = '" & AccountType & "' Order By Name", CxnAccountMaster, adOpenKeyset, adLockOptimistic
    
    rstSizeList.Open "Select Name As Col0, Code From GeneralMaster Where Type = '1' Order By Name", CxnAccountMaster, adOpenKeyset, adLockReadOnly
    rstBindingTypeList.Open "Select Name As Col0, Code From GeneralMaster Where Type = '6' Order By Name", CxnAccountMaster, adOpenKeyset, adLockReadOnly
    rstLaminationTypeList.Open "Select Name As Col0, Code From GeneralMaster Where Type = '7' Order By Name", CxnAccountMaster, adOpenKeyset, adLockReadOnly
    If AccountType = "08" Then
        rstOutsourceItemList.Open "Select Name,'1'+Code As NCode From OutsourceItemMaster Order By Name", CxnAccountMaster, adOpenKeyset, adLockOptimistic
        rstPaperList.Open "Select Name,'2'+Code As NCode From PaperMaster Order By Name", CxnAccountMaster, adOpenKeyset, adLockOptimistic
        rstFreshBookList.Open "Select Name,Board,'3'+Code As NCode From BookMaster Where Type='F' Order By Name", CxnAccountMaster, adOpenKeyset, adLockOptimistic
        rstRepairBookList.Open "Select Name,'4'+Code As NCode From BookMaster Where Type='R' Order By Name", CxnAccountMaster, adOpenKeyset, adLockOptimistic
    End If
    rstAccountMaster.CursorLocation = adUseClient
    rstAccountList.Filter = adFilterNone
    Set DataGrid1.DataSource = rstAccountList
    BusySystemIndicator False
    SSTab1.Tab = 0
    SortOrder = "Name"
    If Not (rstAccountList.EOF Or rstAccountList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    rstAccountList.ActiveConnection = Nothing
    rstSizeList.ActiveConnection = Nothing
    rstOutsourceItemList.ActiveConnection = Nothing
    rstPaperList.ActiveConnection = Nothing
    rstFreshBookList.ActiveConnection = Nothing
    rstRepairBookList.ActiveConnection = Nothing
    rstBindingTypeList.ActiveConnection = Nothing
    rstLaminationTypeList.ActiveConnection = Nothing
    If AccountType = "08" Then
        Call RefreshDropDownList("A")
        fpSpread1.Col = 4
        fpSpread1.ColHidden = True
        fpSpread1.Col = 5
        fpSpread1.ColHidden = True
    End If
    SetButtonsForNoRecord
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    Unload Me
End Sub
Private Sub Form_Activate()
    EnableChildMenu
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
       End If
       If Not EditMode Then
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
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyC And Toolbar1.Buttons.Item(1).Enabled Then
       If InStr(1, "01_02_09", AccountType) = 0 Then
            If MsgBox("Are you sure to make a duplicate copy of the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Proceed !") = vbYes Then
                 Clipboard.SetText rstAccountList.Fields("Code").Value + rstAccountList.Fields("Name").Value
                 PasteRecord
            End If
       End If
       KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyS And Toolbar1.Buttons.Item(4).Enabled Then
        If Not EditMode Then
            Toolbar1_ButtonClick Toolbar1.Buttons.Item(4)
        End If
        KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyF5 And Toolbar1.Buttons.Item(6).Enabled Then
       Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
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
           SSTab1.Tab = Val(AccountType)
           SSTab1.SetFocus
        Else
           If Me.ActiveControl.Name <> "fpSpread1" Then SendKeys "{TAB}"
        End If
        If Me.ActiveControl.Name <> "fpSpread1" Then KeyCode = 0
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Toolbar1.Buttons.Item(4).Enabled Then
        Call Form_KeyDown(vbKeyEscape, 0)
        Cancel = 1
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CloseRecordset(rstAccountList)
    Call CloseRecordset(rstAccountMaster)
    Call CloseRecordset(rstSizeList)
    Call CloseRecordset(rstBindingTypeList)
    Call CloseRecordset(rstLaminationTypeList)
    Call CloseRecordset(rstAccountChild)
    Call CloseConnection(CxnAccountMaster)
    Call CloseRecordset(rstCheckRef)
    Call CloseRecordset(rstOutsourceItemList)
    Call CloseRecordset(rstPaperList)
    Call CloseRecordset(rstFreshBookList)
    Call CloseRecordset(rstRepairBookList)
    ShowProgressInStatusBar False
    DisableChildMenu
End Sub
Private Sub Text1_Change()
    If rstAccountList.RecordCount = 0 Then Exit Sub
    rstAccountList.MoveFirst
    If Text1.Text <> "" Then
        rstAccountList.Find "[" & SortOrder & "] Like '" & FixQuote(Text1.Text) & "%'"
        If rstAccountList.EOF Then
            rstAccountList.MoveFirst
            If PrevStr <> "" And Len(Text1.Text) > 1 Then
                If dblBookMark <> 0 Then
                    rstAccountList.Bookmark = dblBookMark
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
    If Not (rstAccountList.EOF Or rstAccountList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim KeyProcessed As Boolean
    
    If rstAccountList.RecordCount = 0 Then Exit Sub
    If Shift = 0 And KeyCode = vbKeyUp Then
        With rstAccountList
            .MovePrevious
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyBack Then
        With rstAccountList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyDown Then
        With rstAccountList
            .MoveNext
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageUp Then
        With rstAccountList
            .Move (-1) * (DataGrid1.VisibleRows - 1)
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageUp Then
        With rstAccountList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageDown Then
        With rstAccountList
            .Move DataGrid1.VisibleRows - 1
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageDown Then
        With rstAccountList
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
    On Error Resume Next
    
    If Toolbar1.Buttons.Item(1).Enabled Then
        If SSTab1.Tab = Val(AccountType) Then
            ViewRecord
        Else
            If Not (rstAccountList.EOF Or rstAccountList.BOF) Then
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
       If AccountType <> "08" Then
            Text2(Val(AccountType) - 1).SetFocus
        Else
            If SSTab1.Tab = 8 Then
                Mh3dFrame2(Val(AccountType) - 1).Enabled = True
                Mh3dFrame2(Val(AccountType) + 1).Enabled = False
                Text2(Val(AccountType) - 1).SetFocus
            Else
                Mh3dFrame2(Val(AccountType) - 1).Enabled = False
                Mh3dFrame2(Val(AccountType) + 1).Enabled = True
                fpSpread1.SetFocus
            End If
        End If
    End If
End Sub
Public Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim HiLiteRecord As Boolean
    Dim UpdateFlag As Integer, i As Integer
    Dim CellVal As Variant, Imported As Variant
    
    If Button.Index = 1 Then
        If rstAccountMaster.State = adStateOpen Then
           rstAccountMaster.Close
        End If
        rstAccountMaster.Open "Select * From AccountMaster Where Code = ''", CxnAccountMaster, adOpenKeyset, adLockOptimistic
        ClearFields
        Call LoadRateList("")
        If InStr(1, "01_09", AccountType) = 0 Then
            If rstAccountChild.State = adStateClosed Then
                SSTab1.Tab = 0
                Exit Sub
            End If
        End If
        If AddRecord(rstAccountMaster) Then
            Call SetButtons(False)
            SSTab1.Tab = Val(AccountType)
            Text2(Val(AccountType) - 1).SetFocus
            blnRecordExist = False
            CxnAccountMaster.BeginTrans
        End If
    ElseIf Button.Index = 2 Then
        If rstAccountList.RecordCount = 0 Then Exit Sub
        SSTab1.Tab = Val(AccountType)
        EditRecord
    ElseIf Button.Index = 3 Then
        If rstAccountList.RecordCount = 0 Then Exit Sub
        If AllowMastersDeletion = 0 Then
            Call DisplayError("You don't have the rights to Delete this Master")
            Exit Sub
        End If
        SSTab1.Tab = Val(AccountType)
        If CheckRef Then
            DisplayError ("Failed to delete the record")
        ElseIf MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            UpdateFlag = 0
            CxnAccountMaster.BeginTrans
            On Error Resume Next
            MdiMainMenu.MousePointer = vbHourglass
            blnRecordExist = True
            If UpdateRateList("D") Then
                CxnAccountMaster.Execute "Delete From AccountMaster Where Code = '" & rstAccountList.Fields("Code").Value & "'"
                If Err.Number = 0 Then
                    CxnAccountMaster.CommitTrans
                    rstAccountList.Delete
                    rstAccountList.MoveNext
                    If rstAccountList.RecordCount > 0 And rstAccountList.EOF Then rstAccountList.MoveLast
                    Call UpdateUserAction(Choose(Val(AccountType), "Supplier", "Artist", "Composer", "Processor", "Book Printer", "Title Printer", "Laminator", "Binder", "Godown") & " Master", "D", Trim(Text2(Val(AccountType) - 1).Text), CxnAccountMaster)
                    ShowProgressInStatusBar True
                    Timer1.Enabled = True
                    UpdateFlag = 1
                End If
            End If
            If UpdateFlag = 0 Then
                DisplayError ("Failed to delete the record")
                CxnAccountMaster.RollbackTrans
            End If
            MdiMainMenu.MousePointer = vbNormal
            On Error GoTo 0
        End If
        SetButtons (True)
        SetButtonsForNoRecord
        SSTab1.Tab = 0
        HiLiteRecord = True
    ElseIf Button.Index = 4 Then
        If CheckMandatoryFields Then Exit Sub
        If blnRecordExist And AllowMastersModification = 0 Then
            Call DisplayError("You don't have the rights to Edit this Master")
            Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
            Exit Sub
        End If
        SaveFields
        UpdateFlag = 0
        If UpdateRecord(rstAccountMaster) Then
            UpdateFlag = 1
            If UpdateRateList("D") Then
                If InStr(1, "01_02_09", AccountType) = 0 Then
                    If rstAccountChild.RecordCount <> 0 Then
                        rstAccountChild.MoveFirst
                        Do While Not rstAccountChild.EOF
                            If Not UpdateRateList("I") Then
                                UpdateFlag = 0
                                Exit Do
                            End If
                            rstAccountChild.MoveNext
                        Loop
                    End If
                End If
            End If
            If AccountType = "08" Then
                If UpdateFlag Then
                    If UpdateMaterialList("D") Then
                        For i = 1 To fpSpread1.DataRowCnt
                            fpSpread1.SetActiveCell 3, i
                            fpSpread1.GetText 3, i, CellVal
                            fpSpread1.GetText 5, i, Imported
                            If Val(CellVal) <> 0 And Imported = "N" Then
                                If Not UpdateMaterialList("I") Then
                                    UpdateFlag = 0
                                    Exit For
                                End If
                            End If
                        Next
                    End If
                End If
            End If
        End If
        If UpdateFlag Then
            Call UpdateUserAction(Choose(Val(AccountType), "Supplier", "Artist", "Composer", "Processor", "Book Printer", "Title Printer", "Laminator", "Binder", "Godown") & " Master", IIf(blnRecordExist, "M", "A"), Trim(Text2(Val(AccountType) - 1).Text), CxnAccountMaster)
            AddToList
            CxnAccountMaster.CommitTrans
            If rstAccountMaster.State = adStateOpen Then
                rstAccountMaster.Close
            End If
            rstAccountMaster.CursorLocation = adUseClient
            Call SetButtons(True)
            SSTab1.Tab = 0
            ShowProgressInStatusBar True
            Timer1.Enabled = True
        Else
            DisplayError ("Failed to save the record")
            Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
        End If
    ElseIf Button.Index = 5 Then
        If CancelRecordUpdate(rstAccountMaster) Then
            CxnAccountMaster.RollbackTrans
            If rstAccountMaster.State = adStateOpen Then
                rstAccountMaster.Close
            End If
            rstAccountMaster.CursorLocation = adUseClient
            Call SetButtons(True)
            SetButtonsForNoRecord
            SSTab1.Tab = 0
        End If
    ElseIf Button.Index = 6 Then
        SSTab1.Tab = 0
        Set DataGrid1.DataSource = Nothing
        rstAccountList.ActiveConnection = CxnAccountMaster
        Do While Not RefreshRecord(rstAccountList)
        Loop
        Set DataGrid1.DataSource = rstAccountList
        rstAccountList.ActiveConnection = Nothing
        rstSizeList.ActiveConnection = CxnAccountMaster
        Do While Not RefreshRecord(rstSizeList)
        Loop
        rstSizeList.ActiveConnection = Nothing
        rstBindingTypeList.ActiveConnection = CxnAccountMaster
        Do While Not RefreshRecord(rstBindingTypeList)
        Loop
        rstBindingTypeList.ActiveConnection = Nothing
        rstLaminationTypeList.ActiveConnection = CxnAccountMaster
        Do While Not RefreshRecord(rstLaminationTypeList)
        Loop
        rstLaminationTypeList.ActiveConnection = Nothing
        HiLiteRecord = True
    ElseIf Button.Index = 7 Then
        SSTab1.Tab = 0
        With FrmFilter
            .Combo1.AddItem "Name", 0
            .Combo1.ListIndex = 0
            Set .srcForm = Me
            .Show vbModal
        End With
        HiLiteRecord = True
    ElseIf Button.Index = 13 Then
        If rstAccountList.RecordCount > 0 Then rstAccountList.MoveFirst
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 14 Then
        If rstAccountList.RecordCount > 0 Then
           rstAccountList.MovePrevious
           If rstAccountList.BOF Then
              rstAccountList.MoveNext
           End If
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 15 Then
        If rstAccountList.RecordCount > 0 Then
           rstAccountList.MoveNext
           If rstAccountList.EOF Then
              rstAccountList.MovePrevious
           End If
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 16 Then
        If rstAccountList.RecordCount > 0 Then rstAccountList.MoveLast
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 18 Then
        Unload Me
        HiLiteRecord = False
    End If
    If HiLiteRecord Then
        If Not (rstAccountList.EOF Or rstAccountList.BOF) Then
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
    SortOrder = DataGrid1.Columns(ColIndex).DataField
    rstAccountList.Sort = "[" + SortOrder & "] Asc"
    DataGrid1.ClearSelCols
    If Not (rstAccountList.EOF Or rstAccountList.BOF) Then
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
    Toolbar1.Buttons.Item(13).Enabled = bVal
    Toolbar1.Buttons.Item(14).Enabled = bVal
    Toolbar1.Buttons.Item(15).Enabled = bVal
    Toolbar1.Buttons.Item(16).Enabled = bVal
    Toolbar1.Buttons.Item(18).Enabled = bVal
    Mh3dFrame2(Val(AccountType) - 1).Enabled = Not bVal
    Mh3dFrame2(9).Enabled = False
End Sub
Private Sub SetButtonsForNoRecord()
    If rstAccountList.RecordCount = 0 Then
        Toolbar1.Buttons.Item(2).Enabled = False
        Toolbar1.Buttons.Item(3).Enabled = False
        Toolbar1.Buttons.Item(13).Enabled = False
        Toolbar1.Buttons.Item(14).Enabled = False
        Toolbar1.Buttons.Item(15).Enabled = False
        Toolbar1.Buttons.Item(16).Enabled = False
    End If
End Sub
Private Sub Text2_Validate(Index As Integer, Cancel As Boolean)
    If rstAccountMaster.EOF Or rstAccountMaster.BOF Then Exit Sub
    If CheckEmpty(Text2(Val(AccountType) - 1), True) Then
        Cancel = True
    ElseIf CheckDuplicate(CxnAccountMaster, "AccountMaster", "Code", "Name", Trim(Text2(Val(AccountType) - 1).Text), rstAccountMaster.Fields("Code").Value, False) Then
        Cancel = True
    ElseIf CheckEmpty(Text3(Val(AccountType) - 1), False) Then
        Text3(Val(AccountType) - 1).Text = Text2(Val(AccountType) - 1).Text
    End If
End Sub
Private Sub Text09_Validate(Index As Integer, Cancel As Boolean)
    If InStr(1, "01_02_09", AccountType) <> 0 Then Exit Sub
    If rstAccountChild.RecordCount = 0 Then
        Call AddRecord(rstAccountChild)
        Call DataGrid2_KeyDown(Index, vbKeyE, vbCtrlMask)
    End If
End Sub
Private Sub ViewRecord()
    ClearFields
    If rstAccountList.EOF Then
        If rstAccountChild.State = adStateOpen Then rstAccountChild.Close
        Exit Sub
    End If
    FindRecord
    LoadFields
End Sub
Private Sub FindRecord()
    If rstAccountMaster.State = adStateOpen Then
       rstAccountMaster.Close
    End If
    rstAccountMaster.Open "Select * From AccountMaster Where Code = '" & FixQuote(rstAccountList.Fields("Code").Value) & "'", CxnAccountMaster, adOpenKeyset, adLockOptimistic
    If rstAccountMaster.RecordCount = 0 Then
       Call DisplayError("This Record has been deleted by Another User ! Click Ok To Refresh the Recordset")
       Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
    End If
End Sub
Private Sub ClearFields()
    Text2(Val(AccountType) - 1).Text = ""
    Text3(Val(AccountType) - 1).Text = ""
    Text4(Val(AccountType) - 1).Text = ""
    Text5(Val(AccountType) - 1).Text = ""
    Text6(Val(AccountType) - 1).Text = ""
    Text7(Val(AccountType) - 1).Text = ""
    Text8(Val(AccountType) - 1).Text = ""
    Text9(Val(AccountType) - 1).Text = ""
    Text11(Val(AccountType) - 1).Text = ""
    Text12(Val(AccountType) - 1).Text = ""
    Text13(Val(AccountType) - 1).Text = ""
    fpSpread1.ClearRange 1, 1, fpSpread1.MaxCols, fpSpread1.MaxRows, True
End Sub
Private Sub LoadFields()
    If rstAccountMaster.EOF Or rstAccountMaster.BOF Then Exit Sub
    Text2(Val(AccountType) - 1).Text = rstAccountMaster.Fields("Name").Value
    Text3(Val(AccountType) - 1).Text = rstAccountMaster.Fields("PrintName").Value
    Text4(Val(AccountType) - 1).Text = rstAccountMaster.Fields("Address1").Value
    Text5(Val(AccountType) - 1).Text = rstAccountMaster.Fields("Address2").Value
    Text6(Val(AccountType) - 1).Text = rstAccountMaster.Fields("Address3").Value
    Text7(Val(AccountType) - 1).Text = rstAccountMaster.Fields("Address4").Value
    Text8(Val(AccountType) - 1).Text = rstAccountMaster.Fields("Phone").Value
    Text12(Val(AccountType) - 1).Text = rstAccountMaster.Fields("Mobile").Value
    Text13(Val(AccountType) - 1).Text = rstAccountMaster.Fields("Alias").Value
    Text9(Val(AccountType) - 1).Text = rstAccountMaster.Fields("TIN").Value
    Text11(Val(AccountType) - 1).Text = rstAccountMaster.Fields("EMail").Value
    If AccountType = "08" Then
        Call LoadMaterialList(rstAccountMaster.Fields("Code").Value)
    End If
    Call LoadRateList(rstAccountMaster.Fields("Code").Value)
End Sub
Private Sub EditRecord()
    On Error GoTo ErrorHandler
    
    If rstAccountMaster.RecordCount = 0 Then Exit Sub
    If InStr(1, "01_09", AccountType) = 0 Then
        If rstAccountChild.State = adStateClosed Then
            SSTab1.Tab = 0
            Exit Sub
        End If
    End If
    If rstAccountMaster.State = adStateOpen Then
        rstAccountMaster.Close
    End If
    rstAccountMaster.CursorLocation = adUseServer
    rstAccountMaster.Open "Select * From AccountMaster Where Code = '" & FixQuote(rstAccountList.Fields("Code").Value) & "'", CxnAccountMaster, adOpenKeyset, adLockPessimistic
    MdiMainMenu.MousePointer = vbHourglass
    rstAccountMaster.Fields("Printstatus") = "N"
    MdiMainMenu.MousePointer = vbNormal
    AddToList
    Call SetButtons(False)
    SSTab1.TabEnabled(0) = False
    Text2(Val(AccountType) - 1).SetFocus
    blnRecordExist = True
    CxnAccountMaster.BeginTrans
    Exit Sub
ErrorHandler:
    If Err.Number = -2147467259 Then
       Call DisplayError("Failed to Edit the record")
    End If
    MdiMainMenu.MousePointer = vbNormal
    SSTab1.Tab = 0
End Sub
Private Sub SaveFields()
    If rstAccountMaster.EOF Or rstAccountMaster.BOF Then Exit Sub
    If Not blnRecordExist Then
        rstAccountMaster.Fields("Code").Value = GenerateCode(CxnAccountMaster, "Select Max(Code) From AccountMaster", 6, "0")
        rstAccountMaster.Fields("CreatedBy").Value = UserCode
        rstAccountMaster.Fields("CreatedOn").Value = Now()
        rstAccountMaster.Fields("Recordstatus").Value = "N"
    Else
        rstAccountMaster.Fields("ModifiedBy").Value = UserCode
        rstAccountMaster.Fields("ModifiedOn").Value = Now()
        rstAccountMaster.Fields("Recordstatus").Value = "M"
    End If
    rstAccountMaster.Fields("Name").Value = Trim(Text2(Val(AccountType) - 1).Text)
    rstAccountMaster.Fields("PrintName").Value = Trim(Text3(Val(AccountType) - 1).Text)
    rstAccountMaster.Fields("Address1").Value = Trim(Text4(Val(AccountType) - 1).Text)
    rstAccountMaster.Fields("Address2").Value = Trim(Text5(Val(AccountType) - 1).Text)
    rstAccountMaster.Fields("Address3").Value = Trim(Text6(Val(AccountType) - 1).Text)
    rstAccountMaster.Fields("Address4").Value = Trim(Text7(Val(AccountType) - 1).Text)
    rstAccountMaster.Fields("Phone").Value = Trim(Text8(Val(AccountType) - 1).Text)
    rstAccountMaster.Fields("Mobile").Value = Trim(Text12(Val(AccountType) - 1).Text)
    rstAccountMaster.Fields("Alias").Value = Trim(Text13(Val(AccountType) - 1).Text)
    rstAccountMaster.Fields("TIN").Value = Trim(Text9(Val(AccountType) - 1).Text)
    rstAccountMaster.Fields("EMail").Value = Trim(Text11(Val(AccountType) - 1).Text)
    rstAccountMaster.Fields("Type").Value = AccountType
    rstAccountMaster.Fields("PrintStatus").Value = "N"
End Sub
Private Sub AddToList()
    On Error Resume Next
    rstAccountList.MoveFirst
    rstAccountList.Find "[Code] = '" & rstAccountMaster.Fields("Code").Value & "'"
    If rstAccountList.EOF Then
       rstAccountList.AddNew
       rstAccountList.Fields("Code").Value = rstAccountMaster.Fields("Code").Value
    End If
    rstAccountList.Fields("Name").Value = rstAccountMaster.Fields("Name").Value
    rstAccountList.Update
    rstAccountList.Sort = SortOrder & " Asc"
    rstAccountList.Find "[Code] = '" & rstAccountMaster.Fields("Code").Value & "'"
End Sub
Private Function CheckMandatoryFields() As Boolean
    If CheckEmpty(Text2(Val(AccountType) - 1).Text, False) Then
       SSTab1.Tab = Val(AccountType)
       Text2(Val(AccountType) - 1).SetFocus
       CheckMandatoryFields = True
    ElseIf CheckDuplicate(CxnAccountMaster, "AccountMaster", "Code", "Name", Trim(Text2(Val(AccountType) - 1).Text), rstAccountMaster.Fields("Code").Value, False) Then
       SSTab1.Tab = Val(AccountType)
       Text2(Val(AccountType) - 1).SetFocus
       CheckMandatoryFields = True
       SSTab1.Tab = Val(AccountType)
    ElseIf CheckEmpty(Text3(Val(AccountType) - 1).Text, False) Then
       SSTab1.Tab = Val(AccountType)
       Text3(Val(AccountType) - 1).SetFocus
       CheckMandatoryFields = True
    ElseIf CheckItem() Then
       SSTab1.Tab = Val(AccountType + 2)
       fpSpread1.SetFocus
       CheckMandatoryFields = True
    End If
End Function
Public Sub FilterRecord(ByVal SrchFor As String, ByVal SrchText As String)
    If SrchFor = "Name" Then
        rstAccountList.Filter = "[Name] Like '%" & SrchText & "%'"
    End If
End Sub
Private Function CheckRef() As Boolean
    On Error GoTo ErrorHandler
    
    If rstCheckRef.State = adStateOpen Then
         rstCheckRef.Close
    End If
    rstCheckRef.Open "Select BookPrinter From BookPOParent Where BookPrinter = '" & rstAccountList.Fields("Code").Value & "'", CxnAccountMaster, adOpenKeyset, adLockReadOnly
    If rstCheckRef.RecordCount > 0 Then
        CheckRef = True
        Exit Function
    End If
    If rstCheckRef.State = adStateOpen Then
         rstCheckRef.Close
    End If
    rstCheckRef.Open "Select TitlePrinter From BookPOParent Where TitlePrinter = '" & rstAccountList.Fields("Code").Value & "'", CxnAccountMaster, adOpenKeyset, adLockReadOnly
    If rstCheckRef.RecordCount > 0 Then
        CheckRef = True
        Exit Function
    End If
    If rstCheckRef.State = adStateOpen Then
         rstCheckRef.Close
    End If
    rstCheckRef.Open "Select Laminator From BookPOParent Where Laminator = '" & rstAccountList.Fields("Code").Value & "'", CxnAccountMaster, adOpenKeyset, adLockReadOnly
    If rstCheckRef.RecordCount > 0 Then
        CheckRef = True
        Exit Function
    End If
    If rstCheckRef.State = adStateOpen Then
         rstCheckRef.Close
    End If
    rstCheckRef.Open "Select Binder From BookPOParent Where Binder = '" & rstAccountList.Fields("Code").Value & "'", CxnAccountMaster, adOpenKeyset, adLockReadOnly
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
Private Sub LoadRateList(ByVal strAccountCode As String)
    Dim i As Integer
    On Error GoTo ErrorHandler
    
    If rstAccountChild.State = adStateOpen Then
       rstAccountChild.Close
    End If
    If AccountType = "04" Then
        rstAccountChild.Open "Select M2.[Name] As SizeName, M1.* From AccountChild04 M1, GeneralMaster M2 Where M1.[Size] = M2.Code And M1.Code = '" & strAccountCode & "' Order By M2.Name", CxnAccountMaster, adOpenKeyset, adLockOptimistic
    ElseIf AccountType = "05" Then
        rstAccountChild.Open "Select M2.[Name] As SizeName, M1.* From AccountChild05 M1, GeneralMaster M2 Where M1.[Size] = M2.Code And M1.Code = '" & strAccountCode & "' Order By M2.Name", CxnAccountMaster, adOpenKeyset, adLockOptimistic
    ElseIf AccountType = "06" Then
        rstAccountChild.Open "Select M2.[Name] As SizeName, M1.* From AccountChild06 M1, GeneralMaster M2 Where M1.[Size] = M2.Code And M1.Code = '" & strAccountCode & "' Order By M2.Name", CxnAccountMaster, adOpenKeyset, adLockOptimistic
    ElseIf AccountType = "07" Then
        rstAccountChild.Open "Select M2.[Name] As SizeName,M3.[Name] As LaminationTypeName,M1.* From AccountChild07 M1,GeneralMaster M2,GeneralMaster M3 Where M1.[Size] = M2.Code And M1.LaminationType = M3.Code And M3.Type = '7' And M1.Code = '" & strAccountCode & "' Order By M2.Name, M3.Name", CxnAccountMaster, adOpenKeyset, adLockOptimistic
    ElseIf AccountType = "08" Then
        rstAccountChild.Open "Select M2.[Name] As SizeName,M3.[Name] As BindingTypeName,M1.* From AccountChild08 M1,GeneralMaster M2,GeneralMaster M3 Where M1.[Size] = M2.Code And M1.BindingType = M3.Code And M3.Type = '6' And M1.Code = '" & strAccountCode & "' Order By M2.Name, M3.Name", CxnAccountMaster, adOpenKeyset, adLockOptimistic
    End If
    rstAccountChild.ActiveConnection = Nothing
    If InStr(1, "01_02_09", AccountType) = 0 Then
        Set DataGrid2(Val(AccountType) - 1).DataSource = rstAccountChild
    End If
    Exit Sub
ErrorHandler:
    DisplayError ("Failed to Load Rate List")
End Sub
Private Function UpdateRateList(ByVal ActionType As String) As Boolean
    On Error GoTo ErrorHandler
    
    UpdateRateList = True
    If (ActionType = "D" And (Not blnRecordExist)) Or InStr(1, "01_09", AccountType) > 0 Then Exit Function
    If ActionType <> "I" Then
        CxnAccountMaster.Execute "Delete From AccountChild" & AccountType & " Where Code = '" & rstAccountMaster.Fields("Code").Value & "'"
    Else
        If AccountType = "04" Then
            CxnAccountMaster.Execute "Insert Into AccountChild04 Values ('" & rstAccountMaster.Fields("Code").Value & "','" & rstAccountChild.Fields("Size").Value & "'," & Val(rstAccountChild.Fields("OnePieceRate").Value) & "," & Val(rstAccountChild.Fields("CutPieceRate").Value) & "," & Val(rstAccountChild.Fields("PastingRate").Value) & "," & Val(rstAccountChild.Fields("OutputRate").Value) & ")"
        ElseIf AccountType = "05" Then
            CxnAccountMaster.Execute "Insert Into AccountChild05 Values ('" & rstAccountMaster.Fields("Code").Value & "','" & rstAccountChild.Fields("Size").Value & "'," & Val(rstAccountChild.Fields("Range1").Value) & "," & Val(rstAccountChild.Fields("Range2").Value) & "," & Val(rstAccountChild.Fields("Range4").Value) & "," & Val(rstAccountChild.Fields("Range6").Value) & "," & _
                                                            Val(rstAccountChild.Fields("PrintRate1").Value) & "," & Val(rstAccountChild.Fields("PrintRate2").Value) & "," & Val(rstAccountChild.Fields("PrintRate4").Value) & "," & Val(rstAccountChild.Fields("PrintRate6").Value) & "," & Val(rstAccountChild.Fields("PSPlateRate1").Value) & "," & Val(rstAccountChild.Fields("PSPlateRate2").Value) & "," & Val(rstAccountChild.Fields("PSPlateRate4").Value) & "," & Val(rstAccountChild.Fields("PSPlateRate6").Value) & "," & Val(rstAccountChild.Fields("DeepatchPlateRate1").Value) & "," & Val(rstAccountChild.Fields("DeepatchPlateRate2").Value) & "," & Val(rstAccountChild.Fields("DeepatchPlateRate4").Value) & "," & Val(rstAccountChild.Fields("DeepatchPlateRate6").Value) & "," & _
                                                            Val(rstAccountChild.Fields("WipeonPlateRate1").Value) & "," & Val(rstAccountChild.Fields("WipeonPlateRate2").Value) & "," & Val(rstAccountChild.Fields("WipeonPlateRate4").Value) & "," & Val(rstAccountChild.Fields("WipeonPlateRate6").Value) & "," & Val(rstAccountChild.Fields("CTPPlateRate1").Value) & "," & Val(rstAccountChild.Fields("CTPPlateRate2").Value) & "," & Val(rstAccountChild.Fields("CTPPlateRate4").Value) & "," & Val(rstAccountChild.Fields("CTPPlateRate6").Value) & "," & Val(rstAccountChild.Fields("PaperWastageRate1").Value) & "," & Val(rstAccountChild.Fields("PaperWastageRate2").Value) & "," & Val(rstAccountChild.Fields("PaperWastageRate4").Value) & "," & Val(rstAccountChild.Fields("PaperWastageRate6").Value) & ")"
        ElseIf AccountType = "06" Then
            CxnAccountMaster.Execute "Insert Into AccountChild06 Values ('" & rstAccountMaster.Fields("Code").Value & "','" & rstAccountChild.Fields("Size").Value & "'," & Val(rstAccountChild.Fields("Range1").Value) & "," & Val(rstAccountChild.Fields("Range2").Value) & "," & Val(rstAccountChild.Fields("Range4").Value) & "," & Val(rstAccountChild.Fields("Range6").Value) & "," & _
                                                            Val(rstAccountChild.Fields("PrintRate1").Value) & "," & Val(rstAccountChild.Fields("PrintRate2").Value) & "," & Val(rstAccountChild.Fields("PrintRate4").Value) & "," & Val(rstAccountChild.Fields("PrintRate6").Value) & "," & Val(rstAccountChild.Fields("PSPlateRate1").Value) & "," & Val(rstAccountChild.Fields("PSPlateRate2").Value) & "," & Val(rstAccountChild.Fields("PSPlateRate4").Value) & "," & Val(rstAccountChild.Fields("PSPlateRate6").Value) & "," & Val(rstAccountChild.Fields("DeepatchPlateRate1").Value) & "," & Val(rstAccountChild.Fields("DeepatchPlateRate2").Value) & "," & Val(rstAccountChild.Fields("DeepatchPlateRate4").Value) & "," & Val(rstAccountChild.Fields("DeepatchPlateRate6").Value) & "," & Val(rstAccountChild.Fields("WipeonPlateRate1").Value) & "," & Val(rstAccountChild.Fields("WipeonPlateRate2").Value) & "," & Val(rstAccountChild.Fields("WipeonPlateRate4").Value) & "," & Val(rstAccountChild.Fields("WipeonPlateRate6").Value) & "," & _
                                                            Val(rstAccountChild.Fields("CTPPlateRate1").Value) & "," & Val(rstAccountChild.Fields("CTPPlateRate2").Value) & "," & Val(rstAccountChild.Fields("CTPPlateRate4").Value) & "," & Val(rstAccountChild.Fields("CTPPlateRate6").Value) & "," & Val(rstAccountChild.Fields("PaperWastageRate1").Value) & "," & Val(rstAccountChild.Fields("PaperWastageRate2").Value) & "," & Val(rstAccountChild.Fields("PaperWastageRate4").Value) & "," & Val(rstAccountChild.Fields("PaperWastageRate6").Value) & ")"
        ElseIf AccountType = "07" Then
            CxnAccountMaster.Execute "Insert Into AccountChild07 Values ('" & rstAccountMaster.Fields("Code").Value & "','" & rstAccountChild.Fields("Size").Value & "','" & rstAccountChild.Fields("LaminationType").Value & "'," & Val(rstAccountChild.Fields("Rate04").Value) & "," & Val(rstAccountChild.Fields("Rate06").Value) & "," & Val(rstAccountChild.Fields("Rate08").Value) & "," & Val(rstAccountChild.Fields("Rate12").Value) & "," & Val(rstAccountChild.Fields("Rate16").Value) & "," & Val(rstAccountChild.Fields("Rate24").Value) & "," & Val(rstAccountChild.Fields("Rate32").Value) & "," & Val(rstAccountChild.Fields("Rate64").Value) & ")"
        ElseIf AccountType = "08" Then
            CxnAccountMaster.Execute "Insert Into AccountChild08 Values ('" & rstAccountMaster.Fields("Code").Value & "','" & rstAccountChild.Fields("BindingType").Value & "','" & rstAccountChild.Fields("Size").Value & "'," & Val(rstAccountChild.Fields("Range04").Value) & "," & Val(rstAccountChild.Fields("Range06").Value) & "," & Val(rstAccountChild.Fields("Range08").Value) & "," & Val(rstAccountChild.Fields("Range12").Value) & "," & Val(rstAccountChild.Fields("Range16").Value) & "," & Val(rstAccountChild.Fields("Range24").Value) & "," & Val(rstAccountChild.Fields("Range32").Value) & "," & Val(rstAccountChild.Fields("Range64").Value) & "," & _
                                                            Val(rstAccountChild.Fields("FormStitchRate04").Value) & "," & Val(rstAccountChild.Fields("FormStitchRate06").Value) & "," & Val(rstAccountChild.Fields("FormStitchRate08").Value) & "," & Val(rstAccountChild.Fields("FormStitchRate12").Value) & "," & Val(rstAccountChild.Fields("FormStitchRate16").Value) & "," & Val(rstAccountChild.Fields("FormStitchRate24").Value) & "," & Val(rstAccountChild.Fields("FormStitchRate32").Value) & "," & Val(rstAccountChild.Fields("FormStitchRate64").Value) & "," & _
                                                            Val(rstAccountChild.Fields("FormPasteRate04").Value) & "," & Val(rstAccountChild.Fields("FormPasteRate06").Value) & "," & Val(rstAccountChild.Fields("FormPasteRate08").Value) & "," & Val(rstAccountChild.Fields("FormPasteRate12").Value) & "," & Val(rstAccountChild.Fields("FormPasteRate16").Value) & "," & Val(rstAccountChild.Fields("FormPasteRate24").Value) & "," & Val(rstAccountChild.Fields("FormPasteRate32").Value) & "," & Val(rstAccountChild.Fields("FormPasteRate64").Value) & "," & _
                                                            Val(rstAccountChild.Fields("FormFoldRate04").Value) & "," & Val(rstAccountChild.Fields("FormFoldRate06").Value) & "," & Val(rstAccountChild.Fields("FormFoldRate08").Value) & "," & Val(rstAccountChild.Fields("FormFoldRate12").Value) & "," & Val(rstAccountChild.Fields("FormFoldRate16").Value) & "," & Val(rstAccountChild.Fields("FormFoldRate24").Value) & "," & Val(rstAccountChild.Fields("FormFoldRate32").Value) & "," & Val(rstAccountChild.Fields("FormFoldRate64").Value) & "," & _
                                                            Val(rstAccountChild.Fields("Rate/Book04").Value) & "," & Val(rstAccountChild.Fields("Rate/Book06").Value) & "," & Val(rstAccountChild.Fields("Rate/Book08").Value) & "," & Val(rstAccountChild.Fields("Rate/Book12").Value) & "," & Val(rstAccountChild.Fields("Rate/Book16").Value) & "," & Val(rstAccountChild.Fields("Rate/Book24").Value) & "," & Val(rstAccountChild.Fields("Rate/Book32").Value) & "," & Val(rstAccountChild.Fields("Rate/Book64").Value) & "," & Val(rstAccountChild.Fields("PktPackRate").Value) & "," & Val(rstAccountChild.Fields("BoxPackRate").Value) & ")"
        End If
    End If
    Exit Function
ErrorHandler:
    UpdateRateList = False
End Function
Private Sub DataGrid2_DblClick(Index As Integer)
    Call DataGrid2_KeyDown(Index, vbKeyE, vbCtrlMask)
End Sub
Private Sub DataGrid2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask And KeyCode = vbKeyE Then
        If rstAccountChild.RecordCount = 0 Then
            KeyCode = 0
            Exit Sub
        End If
        If AccountType = "04" Then
            Set FrmAccountChild04.rstAccountChild = rstAccountChild
            Set FrmAccountChild04.rstSizeList = rstSizeList
            FrmAccountChild04.AccountName = Trim(Text2(Val(AccountType) - 1).Text)
            On Error Resume Next
            Load FrmAccountChild04
            If Err.Number <> 364 Then
                FrmAccountChild04.Show vbModal
            End If
            On Error GoTo 0
        ElseIf AccountType = "05" Then
            Set FrmAccountChild05.rstAccountChild = rstAccountChild
            Set FrmAccountChild05.rstSizeList = rstSizeList
            FrmAccountChild05.AccountName = Trim(Text2(Val(AccountType) - 1).Text)
            On Error Resume Next
            Load FrmAccountChild05
            If Err.Number <> 364 Then
                FrmAccountChild05.Show vbModal
            End If
            On Error GoTo 0
        ElseIf AccountType = "06" Then
            Set FrmAccountChild06.rstAccountChild = rstAccountChild
            Set FrmAccountChild06.rstSizeList = rstSizeList
            FrmAccountChild06.AccountName = Trim(Text2(Val(AccountType) - 1).Text)
            On Error Resume Next
            Load FrmAccountChild06
            If Err.Number <> 364 Then
                FrmAccountChild06.Show vbModal
            End If
            On Error GoTo 0
        ElseIf AccountType = "07" Then
            Set FrmAccountChild07.rstAccountChild = rstAccountChild
            Set FrmAccountChild07.rstSizeList = rstSizeList
            Set FrmAccountChild07.rstLaminationTypeList = rstLaminationTypeList
            FrmAccountChild07.AccountName = Trim(Text2(Val(AccountType) - 1).Text)
            On Error Resume Next
            Load FrmAccountChild07
            If Err.Number <> 364 Then
                FrmAccountChild07.Show vbModal
            End If
            On Error GoTo 0
        ElseIf AccountType = "08" Then
            Set FrmAccountChild08.rstAccountChild = rstAccountChild
            Set FrmAccountChild08.rstSizeList = rstSizeList
            Set FrmAccountChild08.rstBindingTypeList = rstBindingTypeList
            FrmAccountChild08.AccountName = Trim(Text2(Val(AccountType) - 1).Text)
            On Error Resume Next
            Load FrmAccountChild08
            If Err.Number <> 364 Then FrmAccountChild08.Show vbModal
            On Error GoTo 0
        End If
        KeyCode = 0
        If CheckEmpty(rstAccountChild.Fields("Size").Value, False) Then
            rstAccountChild.Delete
            rstAccountChild.MoveNext
            If rstAccountChild.RecordCount > 0 Then rstAccountChild.MoveFirst
        ElseIf rstAccountChild.AbsolutePosition = rstAccountChild.RecordCount Then
            Call DataGrid2_KeyDown(Index, vbKeyA, vbCtrlMask)
        End If
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyA Then
        SendKeys "^"
        Call AddRecord(rstAccountChild)
        Call DataGrid2_KeyDown(Index, vbKeyE, vbCtrlMask)
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyD Then
        If rstAccountChild.RecordCount = 0 Then Exit Sub
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            Set DataGrid2(Index).DataSource = Nothing
            rstAccountChild.Delete
            rstAccountChild.MoveNext
            Set DataGrid2(Index).DataSource = rstAccountChild
            DataGrid2(Index).SetFocus
        End If
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyS And Toolbar1.Buttons.Item(4).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(4)
    ElseIf Shift = 0 And KeyCode = vbKeyReturn Then
        Text2(Val(AccountType) - 1).SetFocus
        KeyCode = 0
    ElseIf Shift = vbShiftMask And KeyCode = vbKeyTab Then
       Text11(Val(AccountType) - 1).SetFocus
    End If
End Sub
Private Sub DataGrid2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim menusel As String
    
    If Button = vbRightButton Then
        menusel = DisplayPopupMenu(Me.hwnd)
        Select Case menusel
            Case 1
                Call DataGrid2_KeyDown(Index, vbKeyA, vbCtrlMask)
            Case 2
                Call DataGrid2_KeyDown(Index, vbKeyE, vbCtrlMask)
            Case 3
                Call DataGrid2_KeyDown(Index, vbKeyD, vbCtrlMask)
        End Select
    End If
End Sub
Private Sub PasteRecord()
    On Error GoTo ErrorHandler
    
    If rstAccountMaster.State = adStateOpen Then
        rstAccountMaster.Close
    End If
    rstAccountMaster.Open "Select * From AccountMaster Where Code = '" & FixQuote(Left(Clipboard.GetText, 6)) & "'", CxnAccountMaster, adOpenKeyset, adLockOptimistic
    Set rstAccountMaster.ActiveConnection = Nothing
    Call LoadRateList(Left(Clipboard.GetText, 6))
    rstAccountMaster.Fields("Code").Value = GenerateCode(CxnAccountMaster, "Select Max(Code) From AccountMaster", 6, "0")
    rstAccountMaster.Fields("Name").Value = Trim(Mid(Clipboard.GetText, 7, 34)) & " (New)"
    rstAccountMaster.Fields("PrintName").Value = Trim(Mid(Clipboard.GetText, 7, 34)) & " (New)"
    rstAccountMaster.Update
    CxnAccountMaster.BeginTrans
    CxnAccountMaster.Execute "Insert Into AccountMaster Values ('" & rstAccountMaster.Fields("Code").Value & "','" & rstAccountMaster.Fields("Name").Value & "','" & rstAccountMaster.Fields("PrintName").Value & "','" & AccountType & "','','','','','','','','','" & UserCode & "',Now(),'',Null,'N','N')"
    If rstAccountChild.RecordCount <> 0 Then
        rstAccountChild.MoveFirst
        Do While Not rstAccountChild.EOF
            If Not UpdateRateList("I") Then
                GoTo ErrorHandler
            End If
            rstAccountChild.MoveNext
        Loop
    End If
    CxnAccountMaster.CommitTrans
    AddToList
    Clipboard.Clear
    Text1.Text = rstAccountMaster.Fields("Name").Value
    SendKeys "{END}"
    Exit Sub
ErrorHandler:
    DisplayError ("Failed to Paste the record")
    CxnAccountMaster.RollbackTrans
End Sub
Private Sub fpSpread1_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask And KeyCode = vbKeyD Then
        Dim Imported As Variant
        fpSpread1.GetText 5, fpSpread1.ActiveRow, Imported
        If Imported = "Y" Then Exit Sub
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            fpSpread1.DeleteRows fpSpread1.ActiveRow, 1
            fpSpread1.SetFocus
        End If
    ElseIf Shift = 0 And KeyCode = vbKeyF5 Then
        Call RefreshDropDownList("R")
    End If
End Sub
Private Sub fpSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim ActiveCellVal As Variant, Category As Variant
    
    fpSpread1.GetText Col, Row, ActiveCellVal
    If ActiveCellVal = "" Then
        Cancel = True
        Exit Sub
    End If
    fpSpread1.GetText 1, Row, Category
    If Col = 1 Then
        fpSpread1.Col = 2
        fpSpread1.TypeComboBoxList = IIf(Category = "Outsource Item", OutsourceItem, IIf(Category = "Paper", Paper, IIf(Category = "Repair Book", RepairBook, IIf(Category = "Fresh Book", FreshBook, Title))))
    ElseIf Col = 2 Then
        If Category = "Outsource Item" Then
           If rstOutsourceItemList.RecordCount > 0 Then rstOutsourceItemList.MoveFirst
           rstOutsourceItemList.Find "[Name]='" & FixQuote(ActiveCellVal) & "'"
           If Not rstOutsourceItemList.EOF Then
                fpSpread1.SetText 4, Row, rstOutsourceItemList.Fields("NCode").Value
           End If
        ElseIf Category = "Paper" Then
           If rstPaperList.RecordCount > 0 Then rstPaperList.MoveFirst
           rstPaperList.Find "[Name]='" & FixQuote(ActiveCellVal) & "'"
           If Not rstPaperList.EOF Then
                fpSpread1.SetText 4, Row, rstPaperList.Fields("NCode").Value
           End If
        ElseIf Category = "Repair Book" Then
           If rstRepairBookList.RecordCount > 0 Then rstRepairBookList.MoveFirst
           rstRepairBookList.Find "[Name]='" & FixQuote(ActiveCellVal) & "'"
           If Not rstRepairBookList.EOF Then
                fpSpread1.SetText 4, Row, rstRepairBookList.Fields("NCode").Value
           End If
        Else
           If rstFreshBookList.RecordCount > 0 Then rstFreshBookList.MoveFirst
           rstFreshBookList.Find "[Name]='" & FixQuote(ActiveCellVal) & "'"
           If Not rstFreshBookList.EOF Then
                fpSpread1.SetText 4, Row, rstFreshBookList.Fields("NCode").Value
           End If
        End If
    End If
End Sub
Private Function CheckItem() As Boolean
    Dim i As Integer, Item As Variant, Category As Variant
    
    CheckItem = False
    For i = 1 To fpSpread1.DataRowCnt
        fpSpread1.SetActiveCell 1, i
        fpSpread1.GetText 4, i, Item
        fpSpread1.GetText 1, i, Category
        If Category = "Outsource Item" Then
            If Left(Item, 1) <> "1" Then
                CheckItem = True
            End If
        ElseIf Category = "Paper" Then
            If Left(Item, 1) <> "2" Then
                CheckItem = True
            End If
        ElseIf Category = "Repair Book" Then
            If Left(Item, 1) <> "4" Then
                CheckItem = True
            End If
        Else
            If Left(Item, 1) <> "3" And Left(Item, 1) <> "5" Then
                CheckItem = True
            End If
        End If
        If CheckItem Then
            DisplayError "Data mismatch in row #" & Trim(str(i))
            Exit For
        End If
    Next
End Function
Private Sub LoadMaterialList(ByVal strAccountCode As String)
    Dim i As Integer
    On Error GoTo ErrorHandler
    
    If rstAccountChild.State = adStateOpen Then
       rstAccountChild.Close
    End If
    rstAccountChild.Open "SELECT Category,IIF(Category='1',(SELECT Name FROM OutsourceItemMaster WHERE Code=T.Item),IIF(Category='2',(SELECT Name FROM PaperMaster WHERE Code=T.Item),(SELECT Name FROM BookMaster WHERE Code=T.Item))) AS ItemName,OpBal,Category+Item As ItemCode,Imported FROM AccountChild0801 T WHERE Code='" & strAccountCode & "' ORDER BY Category", CxnAccountMaster, adOpenKeyset, adLockReadOnly
    rstAccountChild.ActiveConnection = Nothing
    If rstAccountChild.RecordCount > 0 Then rstAccountChild.MoveFirst
    i = 0
    Do While Not rstAccountChild.EOF
        i = i + 1
        With fpSpread1
            .SetText 1, i, IIf(rstAccountChild.Fields("Category").Value = "1", "Outsource Item", IIf(rstAccountChild.Fields("Category").Value = "2", "Paper", IIf(rstAccountChild.Fields("Category").Value = "3", "Fresh Book", IIf(rstAccountChild.Fields("Category").Value = "4", "Repair Book", "Title"))))
            .Col = 2
            .TypeComboBoxList = IIf(rstAccountChild.Fields("Category").Value = "1", OutsourceItem, IIf(rstAccountChild.Fields("Category").Value = "2", Paper, IIf(rstAccountChild.Fields("Category").Value = "4", RepairBook, IIf(rstAccountChild.Fields("Category").Value = "3", FreshBook, Title))))
            .SetText 2, i, rstAccountChild.Fields("ItemName").Value
            .SetText 3, i, Val(rstAccountChild.Fields("OpBal").Value)
            .SetText 4, i, rstAccountChild.Fields("ItemCode").Value
            .SetText 5, i, rstAccountChild.Fields("Imported").Value
        End With
        rstAccountChild.MoveNext
    Loop
    Exit Sub
ErrorHandler:
    DisplayError ("Failed to Load Material List")
End Sub
Private Function UpdateMaterialList(ByVal ActionType As String) As Boolean
    Dim CellVal(1 To 3) As Variant
    On Error GoTo ErrorHandler
    
    UpdateMaterialList = True
    If ActionType = "D" And (Not blnRecordExist) Then Exit Function
    If ActionType <> "I" Then
        CxnAccountMaster.Execute "Delete From AccountChild0801 Where Code = '" & rstAccountMaster.Fields("Code").Value & "' AND Imported='N'"
    Else
        With fpSpread1
            .GetText 1, .ActiveRow, CellVal(1)
            .GetText 3, .ActiveRow, CellVal(2)
            .GetText 4, .ActiveRow, CellVal(3)
        End With
        CxnAccountMaster.Execute "Insert Into AccountChild0801 Values ('" & rstAccountMaster.Fields("Code").Value & "','" & IIf(CellVal(1) = "Outsource Item", "1", IIf(CellVal(1) = "Paper", "2", IIf(CellVal(1) = "Fresh Book", "3", IIf(CellVal(1) = "Repair Book", "4", "5")))) & "','" & Right(CellVal(3), 6) & "'," & Val(CellVal(2)) & ",'N')"
    End If
    Exit Function
ErrorHandler:
    UpdateMaterialList = False
End Function
Private Sub fpSpread1_BeforeEditMode(ByVal Col As Long, ByVal Row As Long, ByVal UserAction As FPSpreadADO.BeforeEditModeActionConstants, CursorPos As Variant, Cancel As Variant)
    Dim Imported As Variant
    fpSpread1.GetText 5, fpSpread1.ActiveRow, Imported
    If Imported = "Y" Then Cancel = True
End Sub
Private Sub fpSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    EditMode = IIf(Mode = 1, True, False)
End Sub
Private Sub RefreshDropDownList(ByVal xType As String)
    If xType = "R" Then
        rstOutsourceItemList.ActiveConnection = CxnAccountMaster
        Do While Not RefreshRecord(rstOutsourceItemList)
        Loop
        rstOutsourceItemList.ActiveConnection = Nothing
        rstPaperList.ActiveConnection = CxnAccountMaster
        Do While Not RefreshRecord(rstPaperList)
        Loop
        rstPaperList.ActiveConnection = Nothing
        rstFreshBookList.ActiveConnection = CxnAccountMaster
        Do While Not RefreshRecord(rstFreshBookList)
        Loop
        rstFreshBookList.ActiveConnection = Nothing
        rstRepairBookList.ActiveConnection = CxnAccountMaster
        Do While Not RefreshRecord(rstRepairBookList)
        Loop
        rstRepairBookList.ActiveConnection = Nothing
        OutsourceItem = "": Paper = "": FreshBook = "": RepairBook = "": Title = ""
    End If
    Do While Not rstOutsourceItemList.EOF
        If OutsourceItem = "" Then
            OutsourceItem = rstOutsourceItemList.Fields("Name").Value
        Else
            OutsourceItem = OutsourceItem + Chr$(9) + rstOutsourceItemList.Fields("Name").Value
        End If
        rstOutsourceItemList.MoveNext
    Loop
    Do While Not rstPaperList.EOF
        If Paper = "" Then
            Paper = rstPaperList.Fields("Name").Value
        Else
            Paper = Paper + Chr$(9) + rstPaperList.Fields("Name").Value
        End If
        rstPaperList.MoveNext
    Loop
    rstFreshBookList.Filter = "[Board]='000000'"
    Do While Not rstFreshBookList.EOF
        If FreshBook = "" Then
            FreshBook = rstFreshBookList.Fields("Name").Value
        Else
            FreshBook = FreshBook + Chr$(9) + rstFreshBookList.Fields("Name").Value
        End If
        rstFreshBookList.MoveNext
    Loop
    rstFreshBookList.Filter = "[Board]<>'000000'"
    Do While Not rstFreshBookList.EOF
        If Title = "" Then
            Title = rstFreshBookList.Fields("Name").Value
        Else
            Title = Title + Chr$(9) + rstFreshBookList.Fields("Name").Value
        End If
        rstFreshBookList.MoveNext
    Loop
    rstFreshBookList.Filter = adFilterNone
    Do While Not rstRepairBookList.EOF
        If RepairBook = "" Then
            RepairBook = rstRepairBookList.Fields("Name").Value
        Else
            RepairBook = RepairBook + Chr$(9) + rstRepairBookList.Fields("Name").Value
        End If
        rstRepairBookList.MoveNext
    Loop
End Sub
