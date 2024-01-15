VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0F1F1508-C40A-101B-AD04-00AA00575482}#1.0#0"; "mhrinp32.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmBookMaster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Book Master"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6780
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "BookMaster.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7050
   ScaleWidth      =   6780
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   7020
      Left            =   15
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   0
      Width           =   6720
      _Version        =   65536
      _ExtentX        =   11853
      _ExtentY        =   12382
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
      Picture         =   "BookMaster.frx":0442
      Begin TabDlg.SSTab SSTab1 
         Height          =   6810
         Left            =   120
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   120
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   12012
         _Version        =   393216
         Style           =   1
         Tabs            =   4
         TabsPerRow      =   4
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
         TabPicture(0)   =   "BookMaster.frx":045E
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "DataGrid1"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Text1"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "&Details"
         TabPicture(1)   =   "BookMaster.frx":047A
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Mh3dFrame2"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "&BOM"
         TabPicture(2)   =   "BookMaster.frx":0496
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Mh3dFrame3"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "&Corrections"
         TabPicture(3)   =   "BookMaster.frx":04B2
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Mh3dFrame5"
         Tab(3).Control(0).Enabled=   0   'False
         Tab(3).ControlCount=   1
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
            TabIndex        =   34
            Top             =   6315
            Width           =   5775
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   5760
            Left            =   120
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   450
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   10160
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
               DataField       =   "BusyCode"
               Caption         =   "Alias"
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
               DataField       =   "ISBN"
               Caption         =   "ISBN"
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
                  ColumnWidth     =   4394.835
               EndProperty
               BeginProperty Column01 
                  Locked          =   -1  'True
                  ColumnWidth     =   1275.024
               EndProperty
               BeginProperty Column02 
                  Locked          =   -1  'True
                  ColumnWidth     =   2115.213
               EndProperty
            EndProperty
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   6180
            Left            =   -74880
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   480
            Width           =   6255
            _Version        =   65536
            _ExtentX        =   11033
            _ExtentY        =   10901
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
            Picture         =   "BookMaster.frx":04CE
            Begin VB.TextBox Text41 
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
               Left            =   4440
               MaxLength       =   17
               TabIndex        =   27
               Top             =   3850
               Width           =   1695
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
               Left            =   1560
               MaxLength       =   255
               TabIndex        =   28
               Top             =   4170
               Width           =   4575
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
               Left            =   1560
               MaxLength       =   40
               TabIndex        =   10
               Top             =   1980
               Width           =   1575
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
               Left            =   4440
               MaxLength       =   40
               TabIndex        =   3
               Top             =   725
               Width           =   1695
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
               Left            =   4440
               MaxLength       =   40
               TabIndex        =   11
               Top             =   1980
               Width           =   1695
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
               Left            =   4440
               MaxLength       =   40
               TabIndex        =   9
               Top             =   1670
               Width           =   1695
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
               MaxLength       =   40
               TabIndex        =   8
               Top             =   1670
               Width           =   1575
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
               Left            =   4440
               MaxLength       =   40
               TabIndex        =   7
               Top             =   1345
               Width           =   1695
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
               Left            =   1560
               MaxLength       =   40
               TabIndex        =   6
               Top             =   1345
               Width           =   1575
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
               Left            =   4440
               MaxLength       =   40
               TabIndex        =   5
               Top             =   1040
               Width           =   1695
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
               MaxLength       =   17
               TabIndex        =   2
               Top             =   725
               Width           =   1575
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
               MaxLength       =   60
               TabIndex        =   1
               Top             =   410
               Width           =   4575
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
               Left            =   1560
               MaxLength       =   60
               TabIndex        =   0
               Top             =   100
               Width           =   4575
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   330
               Left            =   120
               TabIndex        =   37
               Top             =   405
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
               Caption         =   " Print Name"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "BookMaster.frx":04EA
               Picture         =   "BookMaster.frx":0506
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Left            =   120
               TabIndex        =   38
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
               Caption         =   " Name"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "BookMaster.frx":0522
               Picture         =   "BookMaster.frx":053E
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
               Height          =   330
               Left            =   120
               TabIndex        =   39
               Top             =   1980
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
               Caption         =   " Lamination Type"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "BookMaster.frx":055A
               Picture         =   "BookMaster.frx":0576
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Left            =   120
               TabIndex        =   40
               Top             =   720
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
               Caption         =   " ISBN"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "BookMaster.frx":0592
               Picture         =   "BookMaster.frx":05AE
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
               Height          =   330
               Left            =   3120
               TabIndex        =   41
               Top             =   1035
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
               Caption         =   " Board"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "BookMaster.frx":05CA
               Picture         =   "BookMaster.frx":05E6
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel9 
               Height          =   330
               Left            =   3120
               TabIndex        =   42
               Top             =   1350
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
               Caption         =   " Subject"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "BookMaster.frx":0602
               Picture         =   "BookMaster.frx":061E
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel10 
               Height          =   330
               Left            =   3120
               TabIndex        =   43
               Top             =   1665
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
               Picture         =   "BookMaster.frx":063A
               Picture         =   "BookMaster.frx":0656
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
               Height          =   330
               Left            =   120
               TabIndex        =   44
               Top             =   2610
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
               Caption         =   " Binding Forms"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "BookMaster.frx":0672
               Picture         =   "BookMaster.frx":068E
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
               Height          =   330
               Left            =   120
               TabIndex        =   45
               Top             =   4480
               Width           =   6015
               _Version        =   65536
               _ExtentX        =   10610
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
               Caption         =   " Printing Form Details"
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "BookMaster.frx":06AA
               Picture         =   "BookMaster.frx":06C6
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel14 
               Height          =   330
               Left            =   120
               TabIndex        =   46
               Top             =   1040
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
               Caption         =   " Price"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "BookMaster.frx":06E2
               Picture         =   "BookMaster.frx":06FE
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel15 
               Height          =   330
               Left            =   120
               TabIndex        =   47
               Top             =   1345
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
               Caption         =   " Class"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "BookMaster.frx":071A
               Picture         =   "BookMaster.frx":0736
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel16 
               Height          =   330
               Left            =   120
               TabIndex        =   48
               Top             =   1670
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
               Caption         =   " Group"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "BookMaster.frx":0752
               Picture         =   "BookMaster.frx":076E
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel17 
               Height          =   330
               Left            =   3120
               TabIndex        =   49
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
               Caption         =   " Binding Type"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "BookMaster.frx":078A
               Picture         =   "BookMaster.frx":07A6
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel18 
               Height          =   330
               Left            =   3120
               TabIndex        =   50
               Top             =   2295
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
               Caption         =   " Form Type"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "BookMaster.frx":07C2
               Picture         =   "BookMaster.frx":07DE
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel19 
               Height          =   330
               Left            =   3120
               TabIndex        =   51
               Top             =   2610
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
               Caption         =   " Pages/Forms"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "BookMaster.frx":07FA
               Picture         =   "BookMaster.frx":0816
            End
            Begin MhinrelLib.MhRealInput MhRealInput7 
               Height          =   330
               Left            =   5280
               TabIndex        =   52
               TabStop         =   0   'False
               ToolTipText     =   "Forms"
               Top             =   2610
               Width           =   855
               _Version        =   65536
               _ExtentX        =   1499
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
               DecimalPlaces   =   2
               VAlignment      =   2
               FocusSelect     =   -1  'True
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel13 
               Height          =   330
               Left            =   3120
               TabIndex        =   53
               Top             =   725
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
               Caption         =   " Alias"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "BookMaster.frx":0832
               Picture         =   "BookMaster.frx":084E
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput4 
               Height          =   330
               Left            =   1560
               TabIndex        =   15
               Top             =   2610
               Width           =   795
               _Version        =   65536
               _ExtentX        =   1402
               _ExtentY        =   582
               Calculator      =   "BookMaster.frx":086A
               Caption         =   "BookMaster.frx":088A
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookMaster.frx":08F6
               Keys            =   "BookMaster.frx":0914
               Spin            =   "BookMaster.frx":095E
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "##0"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "##0"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   999
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
            Begin TDBNumber6Ctl.TDBNumber MhRealInput1 
               Height          =   330
               Left            =   1560
               TabIndex        =   4
               ToolTipText     =   "Printing Form"
               Top             =   1040
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               Calculator      =   "BookMaster.frx":0986
               Caption         =   "BookMaster.frx":09A6
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookMaster.frx":0A12
               Keys            =   "BookMaster.frx":0A30
               Spin            =   "BookMaster.frx":0A7A
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
               ValueVT         =   1962016773
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput15 
               Height          =   330
               Left            =   4440
               TabIndex        =   54
               TabStop         =   0   'False
               ToolTipText     =   "Pages"
               Top             =   2610
               Width           =   855
               _Version        =   65536
               _ExtentX        =   1508
               _ExtentY        =   582
               Calculator      =   "BookMaster.frx":0AA2
               Caption         =   "BookMaster.frx":0AC2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookMaster.frx":0B2E
               Keys            =   "BookMaster.frx":0B4C
               Spin            =   "BookMaster.frx":0B96
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "###0"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "###0"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   9999
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   1
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   2088828933
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel7 
               Height          =   330
               Left            =   120
               TabIndex        =   55
               Top             =   2295
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
               Caption         =   " Add-On Rates"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "BookMaster.frx":0BBE
               Picture         =   "BookMaster.frx":0BDA
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput3 
               Height          =   330
               Left            =   1560
               TabIndex        =   12
               ToolTipText     =   "Book Printing"
               Top             =   2295
               Width           =   795
               _Version        =   65536
               _ExtentX        =   1402
               _ExtentY        =   582
               Calculator      =   "BookMaster.frx":0BF6
               Caption         =   "BookMaster.frx":0C16
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookMaster.frx":0C82
               Keys            =   "BookMaster.frx":0CA0
               Spin            =   "BookMaster.frx":0CEA
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
            Begin TDBNumber6Ctl.TDBNumber MhRealInput16 
               Height          =   330
               Left            =   2340
               TabIndex        =   13
               ToolTipText     =   "Binding"
               Top             =   2295
               Width           =   795
               _Version        =   65536
               _ExtentX        =   1402
               _ExtentY        =   582
               Calculator      =   "BookMaster.frx":0D12
               Caption         =   "BookMaster.frx":0D32
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookMaster.frx":0D9E
               Keys            =   "BookMaster.frx":0DBC
               Spin            =   "BookMaster.frx":0E06
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
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel26 
               Height          =   330
               Left            =   120
               TabIndex        =   56
               Top             =   2930
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
               Caption         =   " Title Plate Type"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "BookMaster.frx":0E2E
               Picture         =   "BookMaster.frx":0E4A
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel27 
               Height          =   330
               Left            =   3120
               TabIndex        =   57
               Top             =   2930
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
               Caption         =   " Title Color"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "BookMaster.frx":0E66
               Picture         =   "BookMaster.frx":0E82
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput17 
               Height          =   330
               Left            =   4440
               TabIndex        =   18
               ToolTipText     =   "Front Color"
               Top             =   2930
               Width           =   855
               _Version        =   65536
               _ExtentX        =   1508
               _ExtentY        =   582
               Calculator      =   "BookMaster.frx":0E9E
               Caption         =   "BookMaster.frx":0EBE
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookMaster.frx":0F2A
               Keys            =   "BookMaster.frx":0F48
               Spin            =   "BookMaster.frx":0F92
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "0"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "0"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   9
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
               Value           =   4
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput18 
               Height          =   330
               Left            =   5280
               TabIndex        =   19
               ToolTipText     =   "Back Color"
               Top             =   2930
               Width           =   855
               _Version        =   65536
               _ExtentX        =   1508
               _ExtentY        =   582
               Calculator      =   "BookMaster.frx":0FBA
               Caption         =   "BookMaster.frx":0FDA
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookMaster.frx":1046
               Keys            =   "BookMaster.frx":1064
               Spin            =   "BookMaster.frx":10AE
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "0"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "0"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   9
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
            Begin FPSpreadADO.fpSpread fpSpread2 
               Height          =   1335
               Left            =   120
               TabIndex        =   29
               Top             =   4785
               Width           =   6015
               _Version        =   524288
               _ExtentX        =   10610
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
               MaxCols         =   7
               MaxRows         =   3
               OperationMode   =   2
               SpreadDesigner  =   "BookMaster.frx":10D6
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
               Height          =   330
               Left            =   120
               TabIndex        =   62
               Top             =   3240
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
               Caption         =   " Duplex Printing"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "BookMaster.frx":1859
               Picture         =   "BookMaster.frx":1875
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel12 
               Height          =   330
               Index           =   0
               Left            =   120
               TabIndex        =   63
               Top             =   3560
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
               Caption         =   " Qty/Pkt"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "BookMaster.frx":1891
               Picture         =   "BookMaster.frx":18AD
            End
            Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame4 
               Height          =   330
               Left            =   1560
               TabIndex        =   64
               TabStop         =   0   'False
               Top             =   3240
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
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
               Picture         =   "BookMaster.frx":18C9
               Begin VB.OptionButton Option2 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "No"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Left            =   915
                  TabIndex        =   21
                  Top             =   60
                  Width           =   615
               End
               Begin VB.OptionButton Option1 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Yes"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Left            =   120
                  TabIndex        =   20
                  Top             =   60
                  Value           =   -1  'True
                  Width           =   585
               End
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput5 
               Height          =   330
               Left            =   1560
               TabIndex        =   23
               ToolTipText     =   "Qty/Pkt"
               Top             =   3560
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               Calculator      =   "BookMaster.frx":18E5
               Caption         =   "BookMaster.frx":1905
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookMaster.frx":1971
               Keys            =   "BookMaster.frx":198F
               Spin            =   "BookMaster.frx":19D9
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "###0"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "###0"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   9999
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   1962016773
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel20 
               Height          =   330
               Left            =   3120
               TabIndex        =   65
               Top             =   3560
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
               Caption         =   " Pkt&&Loose/Box"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "BookMaster.frx":1A01
               Picture         =   "BookMaster.frx":1A1D
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput6 
               Height          =   330
               Left            =   4440
               TabIndex        =   24
               ToolTipText     =   "Pkt/Box"
               Top             =   3560
               Width           =   855
               _Version        =   65536
               _ExtentX        =   1508
               _ExtentY        =   582
               Calculator      =   "BookMaster.frx":1A39
               Caption         =   "BookMaster.frx":1A59
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookMaster.frx":1AC5
               Keys            =   "BookMaster.frx":1AE3
               Spin            =   "BookMaster.frx":1B2D
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "#####0"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "#####0"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   999999
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   1962016773
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput19 
               Height          =   330
               Left            =   2340
               TabIndex        =   16
               ToolTipText     =   "Extra Forms"
               Top             =   2610
               Width           =   795
               _Version        =   65536
               _ExtentX        =   1402
               _ExtentY        =   582
               Calculator      =   "BookMaster.frx":1B55
               Caption         =   "BookMaster.frx":1B75
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookMaster.frx":1BE1
               Keys            =   "BookMaster.frx":1BFF
               Spin            =   "BookMaster.frx":1C49
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "##0"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "##0"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   999
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
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel21 
               Height          =   330
               Left            =   120
               TabIndex        =   66
               Top             =   4170
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
               Picture         =   "BookMaster.frx":1C71
               Picture         =   "BookMaster.frx":1C8D
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel22 
               Height          =   330
               Left            =   3120
               TabIndex        =   67
               Top             =   3240
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
               Caption         =   " Royalty (%)"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "BookMaster.frx":1CA9
               Picture         =   "BookMaster.frx":1CC5
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput10 
               Height          =   330
               Left            =   4440
               TabIndex        =   22
               ToolTipText     =   "Binding"
               Top             =   3240
               Width           =   1695
               _Version        =   65536
               _ExtentX        =   2990
               _ExtentY        =   582
               Calculator      =   "BookMaster.frx":1CE1
               Caption         =   "BookMaster.frx":1D01
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookMaster.frx":1D6D
               Keys            =   "BookMaster.frx":1D8B
               Spin            =   "BookMaster.frx":1DD5
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
               ValueVT         =   2088828933
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput8 
               Height          =   330
               Left            =   5280
               TabIndex        =   25
               ToolTipText     =   "Loose Qty/Box"
               Top             =   3560
               Width           =   855
               _Version        =   65536
               _ExtentX        =   1508
               _ExtentY        =   582
               Calculator      =   "BookMaster.frx":1DFD
               Caption         =   "BookMaster.frx":1E1D
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookMaster.frx":1E89
               Keys            =   "BookMaster.frx":1EA7
               Spin            =   "BookMaster.frx":1EF1
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "###0"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "###0"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   9999
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   1962016773
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel12 
               Height          =   340
               Index           =   2
               Left            =   120
               TabIndex        =   68
               Top             =   3840
               Width           =   1455
               _Version        =   65536
               _ExtentX        =   2566
               _ExtentY        =   600
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
               Caption         =   " Sale Discount"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "BookMaster.frx":1F19
               Picture         =   "BookMaster.frx":1F35
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel12 
               Height          =   340
               Index           =   3
               Left            =   3120
               TabIndex        =   69
               Top             =   3840
               Width           =   1335
               _Version        =   65536
               _ExtentX        =   2355
               _ExtentY        =   600
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
               Caption         =   " HSN Code"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "BookMaster.frx":1F51
               Picture         =   "BookMaster.frx":1F6D
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput52 
               Height          =   330
               Left            =   1560
               TabIndex        =   26
               ToolTipText     =   "Qty/Pkt"
               Top             =   3850
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   582
               Calculator      =   "BookMaster.frx":1F89
               Caption         =   "BookMaster.frx":1FA9
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "BookMaster.frx":2015
               Keys            =   "BookMaster.frx":2033
               Spin            =   "BookMaster.frx":207D
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "###0"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "###0"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   9999
               MinValue        =   0
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   1962016773
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin MSForms.ComboBox Combo7 
               Height          =   330
               Left            =   1560
               TabIndex        =   17
               Top             =   2930
               Width           =   1575
               VariousPropertyBits=   545282075
               BackColor       =   16777215
               BorderStyle     =   1
               DisplayStyle    =   7
               Size            =   "2778;582"
               MatchEntry      =   0
               ShowDropButtonWhen=   1
               SpecialEffect   =   0
               FontName        =   "Calibri"
               FontHeight      =   195
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.ComboBox Combo1 
               Height          =   330
               Left            =   4440
               TabIndex        =   14
               Top             =   2295
               Width           =   1695
               VariousPropertyBits=   545282075
               BackColor       =   16777215
               BorderStyle     =   1
               DisplayStyle    =   7
               Size            =   "2990;582"
               MatchEntry      =   0
               ShowDropButtonWhen=   1
               SpecialEffect   =   0
               FontName        =   "Calibri"
               FontHeight      =   195
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame3 
            Height          =   5600
            Left            =   -74880
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   480
            Width           =   6255
            _Version        =   65536
            _ExtentX        =   11033
            _ExtentY        =   9878
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
            Picture         =   "BookMaster.frx":20A5
            Begin FPSpreadADO.fpSpread fpSpread1 
               Height          =   5385
               Left            =   120
               TabIndex        =   61
               Top             =   105
               Width           =   6015
               _Version        =   524288
               _ExtentX        =   10610
               _ExtentY        =   9499
               _StockProps     =   64
               EditEnterAction =   5
               EditModePermanent=   -1  'True
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
               MaxCols         =   4
               MaxRows         =   100
               OperationMode   =   2
               SpreadDesigner  =   "BookMaster.frx":20C1
            End
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame5 
            Height          =   5595
            Left            =   -74880
            TabIndex        =   59
            TabStop         =   0   'False
            Top             =   480
            Width           =   6255
            _Version        =   65536
            _ExtentX        =   11033
            _ExtentY        =   9878
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
            Picture         =   "BookMaster.frx":2659
            Begin FPSpreadADO.fpSpread fpSpread3 
               Height          =   5385
               Left            =   120
               TabIndex        =   60
               Top             =   105
               Width           =   6015
               _Version        =   524288
               _ExtentX        =   10610
               _ExtentY        =   9499
               _StockProps     =   64
               EditEnterAction =   5
               EditModePermanent=   -1  'True
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
               MaxCols         =   3
               MaxRows         =   100
               OperationMode   =   2
               SpreadDesigner  =   "BookMaster.frx":2675
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
            ForeColor       =   &H80000005&
            Height          =   330
            Left            =   120
            TabIndex        =   35
            Top             =   6315
            Width           =   495
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Width           =   6780
      _ExtentX        =   11959
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
Attribute VB_Name = "FrmBookMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CxnBookMaster As New ADODB.Connection
Dim rstBookList As New ADODB.Recordset
Dim rstBookMaster As New ADODB.Recordset
Dim rstBoardList As New ADODB.Recordset
Dim rstClassList As New ADODB.Recordset
Dim rstSubjectList As New ADODB.Recordset
Dim rstGroupList As New ADODB.Recordset
Dim rstSizeList As New ADODB.Recordset
Dim rstBindingTypeList As New ADODB.Recordset
Dim rstLaminationTypeList As New ADODB.Recordset
Dim rstOutsourceItemList As New ADODB.Recordset
Dim rstFreshBookList As New ADODB.Recordset
Dim rstBookChild As New ADODB.Recordset


Dim BoardCode As String
Dim ClassCode As String
Dim SubjectCode As String
Dim GroupCode As String
Dim SizeCode As String
Dim BindingTypeCode As String
Dim LaminationTypeCode As String
Dim SortOrder As String
Dim PrevStr As String
Dim dblBookMark As Double
Dim blnRecordExist As Boolean
Dim OutsourceItem As String
Dim FreshBook As String
Dim EditMode As Boolean
Public BookType As String
Private Sub Form_Load()
    
    On Error GoTo ErrorHandler
    CenterForm Me
    
    BusySystemIndicator True
    Me.Caption = IIf(BookType = "F", "Book Master [Fresh]", "Book Master [Repair]")
    CxnBookMaster.CursorLocation = adUseClient
    CxnBookMaster.Open CxnDatabase.ConnectionString
    rstBookList.Open "Select Name,BusyCode,Board,ISBN,Code From BookMaster Where Type='" & BookType & "' Order By Name", CxnBookMaster, adOpenKeyset, adLockOptimistic
    rstBoardList.Open "Select Name As Col0, Code From GeneralMaster Where Type = '2' Order By Name", CxnBookMaster, adOpenKeyset, adLockReadOnly
    rstClassList.Open "Select Name As Col0, Code From GeneralMaster Where Type = '4' Order By Name", CxnBookMaster, adOpenKeyset, adLockReadOnly
    rstSubjectList.Open "Select Name As Col0, Code From GeneralMaster Where Type = '3' Order By Name", CxnBookMaster, adOpenKeyset, adLockReadOnly
    rstGroupList.Open "Select Name As Col0, Code From GeneralMaster Where Type = '5' Order By Name", CxnBookMaster, adOpenKeyset, adLockReadOnly
    rstSizeList.Open "Select Name As Col0, Code From GeneralMaster Where Type = '1' Order By Name", CxnBookMaster, adOpenKeyset, adLockReadOnly
    rstBindingTypeList.Open "Select Name As Col0, Code From GeneralMaster Where Type = '6' Order By Name", CxnBookMaster, adOpenKeyset, adLockReadOnly
    rstLaminationTypeList.Open "Select Name As Col0, Code From GeneralMaster Where Type = '7' Order By Name", CxnBookMaster, adOpenKeyset, adLockReadOnly
    rstOutsourceItemList.Open "Select Name,'1'+Code As NCode From OutsourceItemMaster Order By Name", CxnBookMaster, adOpenKeyset, adLockOptimistic
    rstFreshBookList.Open "Select Name,'3'+Code As NCode From BookMaster Where Type='F' AND Board='000000' Order By Name", CxnBookMaster, adOpenKeyset, adLockOptimistic
    rstBookMaster.CursorLocation = adUseClient
    rstBookList.Filter = adFilterNone
    Set DataGrid1.DataSource = rstBookList
    BusySystemIndicator False
    SSTab1.Tab = 0
    SortOrder = "Name"
    If Not (rstBookList.EOF Or rstBookList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    rstBookList.ActiveConnection = Nothing
    rstBoardList.ActiveConnection = Nothing
    rstClassList.ActiveConnection = Nothing
    rstSubjectList.ActiveConnection = Nothing
    rstGroupList.ActiveConnection = Nothing
    rstSizeList.ActiveConnection = Nothing
    rstBindingTypeList.ActiveConnection = Nothing
    rstLaminationTypeList.ActiveConnection = Nothing
    rstOutsourceItemList.ActiveConnection = Nothing
    rstFreshBookList.ActiveConnection = Nothing
    
    Combo1.AddItem "4 Pages", 0
    Combo1.AddItem "6 Pages", 1
    Combo1.AddItem "8 Pages", 2
    Combo1.AddItem "12 Pages", 3
    Combo1.AddItem "16 Pages", 4
    Combo1.AddItem "24 Pages", 5
    Combo1.AddItem "32 Pages", 6
    Combo1.AddItem "64 Pages", 7
    Combo7.AddItem "Depatch", 0
    Combo7.AddItem "PS", 1
    Combo7.AddItem "Wipeon", 2
    Combo7.AddItem "CTP", 3
    
'    Combo3.AddItem "Old", 0
'    Combo3.AddItem "New", 1
'    Combo3.AddItem "Revised", 2
    
    Call RefreshDropDownList("A")
    
    fpSpread1.Col = 4
    fpSpread1.ColHidden = True
    SetButtonsForNoRecord
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    Call CloseForm(FrmBookMaster)
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
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyS And Toolbar1.Buttons.Item(4).Enabled Then
        If Not EditMode Then Toolbar1_ButtonClick Toolbar1.Buttons.Item(4)
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyC Then
        DuplicateRecord
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
           SSTab1.Tab = 1
           SSTab1.SetFocus
        Else
           If Me.ActiveControl.Name <> "fpSpread1" And Me.ActiveControl.Name <> "fpSpread2" And Me.ActiveControl.Name <> "fpSpread3" Then
              Sendkeys "{TAB}"
           End If
        End If
        If Me.ActiveControl.Name <> "fpSpread1" And Me.ActiveControl.Name <> "fpSpread2" And Me.ActiveControl.Name <> "fpSpread3" Then
            KeyCode = 0
        End If
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Toolbar1.Buttons.Item(4).Enabled Then
        Call Form_KeyDown(vbKeyEscape, 0)
        Cancel = 1
    Else
        Call CloseForm(FrmBookMaster)
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CloseRecordset(rstBookList)
    Call CloseRecordset(rstBookMaster)
    Call CloseRecordset(rstBoardList)
    Call CloseRecordset(rstClassList)
    Call CloseRecordset(rstSubjectList)
    Call CloseRecordset(rstGroupList)
    Call CloseRecordset(rstSizeList)
    Call CloseRecordset(rstBindingTypeList)
    Call CloseRecordset(rstLaminationTypeList)
    Call CloseRecordset(rstBookChild)
    Call CloseRecordset(rstOutsourceItemList)
    Call CloseRecordset(rstFreshBookList)
    Call CloseConnection(CxnBookMaster)
    ShowProgressInStatusBar False
    DisableChildMenu
End Sub
Private Sub Text1_Change()
    If rstBookList.RecordCount = 0 Then Exit Sub
    rstBookList.MoveFirst
    If Text1.Text <> "" Then
        rstBookList.Find "[" & SortOrder & "] Like '" & FixQuote(Text1.Text) & "%'"
        If rstBookList.EOF Then
            rstBookList.MoveFirst
            If PrevStr <> "" And Len(Text1.Text) > 1 Then
                If dblBookMark <> 0 Then
                    rstBookList.Bookmark = dblBookMark
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
    If Not (rstBookList.EOF Or rstBookList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim KeyProcessed As Boolean
    
    If rstBookList.RecordCount = 0 Then Exit Sub
    If Shift = 0 And KeyCode = vbKeyUp Then
        With rstBookList
            .MovePrevious
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyBack Then
        With rstBookList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyDown Then
        With rstBookList
            .MoveNext
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageUp Then
        With rstBookList
            .Move (-1) * (DataGrid1.VisibleRows - 1)
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageUp Then
        With rstBookList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageDown Then
        With rstBookList
            .Move DataGrid1.VisibleRows - 1
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageDown Then
        With rstBookList
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
        If SSTab1.Tab >= 1 Then
            ViewRecord
        Else
            If Not (rstBookList.EOF Or rstBookList.BOF) Then
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
        If SSTab1.Tab = 1 Then
            Mh3dFrame2.Enabled = True
            Mh3dFrame3.Enabled = False
            Mh3dFrame5.Enabled = False
            Text2.SetFocus
        ElseIf SSTab1.Tab = 2 Then
            Mh3dFrame2.Enabled = False
            Mh3dFrame5.Enabled = False
            Mh3dFrame3.Enabled = True
            fpSpread1.SetFocus
        Else
            Mh3dFrame5.Enabled = True
            Mh3dFrame2.Enabled = False
            Mh3dFrame3.Enabled = False
            fpSpread3.SetFocus
        End If
    End If
End Sub
Public Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim HiLiteRecord As Boolean
    Dim UpdateFlag As Integer, i As Integer
     Dim CellVal As Variant
   
    If Button.Index = 1 Then
        If rstBookMaster.State = adStateOpen Then
           rstBookMaster.Close
        End If
        rstBookMaster.Open "Select * From BookMaster Where Code = ''", CxnBookMaster, adOpenKeyset, adLockOptimistic
        ClearFields
        If AddRecord(rstBookMaster) Then
           Call SetButtons(False)
           SSTab1.Tab = 1
           Text2.SetFocus
           blnRecordExist = False
           CxnBookMaster.BeginTrans
        End If
    ElseIf Button.Index = 2 Then
        If rstBookList.RecordCount = 0 Then Exit Sub
        SSTab1.Tab = 1
        EditRecord
    ElseIf Button.Index = 3 Then
        If rstBookList.RecordCount = 0 Then Exit Sub
        If AllowMastersDeletion = 0 Or rstBookList.Fields("Board").Value = "000000" Then
            Call DisplayError("You don't have the rights to Delete this Master")
            Exit Sub
        End If
        SSTab1.Tab = 1
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            On Error Resume Next
            MdiMainMenu.MousePointer = vbHourglass
            CxnBookMaster.Execute "Delete From BookMaster WHERE Code = '" & rstBookList.Fields("Code").Value & "'"
            MdiMainMenu.MousePointer = vbNormal
            If Err.Number = 0 Then
                rstBookList.Delete
                rstBookList.MoveNext
                If rstBookList.RecordCount > 0 And rstBookList.EOF Then rstBookList.MoveLast
                Call UpdateUserAction("Book Master", "D", Trim(Text2.Text), CxnBookMaster)
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
        If blnRecordExist And AllowMastersModification = 0 Then
            Call DisplayError("You don't have the rights to Edit this Master")
            Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
            Exit Sub
        End If
        SaveFields
        UpdateFlag = 0
        If UpdateRecord(rstBookMaster) Then
            If UpdateMaterialList("D1") Then
                UpdateFlag = 1
                For i = 1 To fpSpread1.DataRowCnt
                    fpSpread1.SetActiveCell 3, i
                    fpSpread1.GetText 3, i, CellVal
                    If Val(CellVal) <> 0 Then
                        If Not UpdateMaterialList("I1") Then
                            UpdateFlag = 0
                            Exit For
                        End If
                    End If
                Next
            End If
            If UpdateFlag Then
                If UpdateMaterialList("D2") Then
                    For i = 1 To fpSpread3.DataRowCnt
                        fpSpread3.SetActiveCell 1, i
                        If Not UpdateMaterialList("I2") Then
                            UpdateFlag = 0
                            Exit For
                        End If
                    Next
                End If
            End If
        End If
        If UpdateFlag Then
            Call UpdateUserAction("Book Master", IIf(blnRecordExist, "M", "A"), Trim(Text2.Text), CxnBookMaster)
            AddToList
            CxnBookMaster.CommitTrans
            If rstBookMaster.State = adStateOpen Then rstBookMaster.Close
            rstBookMaster.CursorLocation = adUseClient
            Call SetButtons(True)
            SSTab1.Tab = 0
            ShowProgressInStatusBar True
            Timer1.Enabled = True
        Else
            DisplayError ("Failed to save the record")
            Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
        End If
    ElseIf Button.Index = 5 Then
        If CancelRecordUpdate(rstBookMaster) Then
            CxnBookMaster.RollbackTrans
            If rstBookMaster.State = adStateOpen Then
                rstBookMaster.Close
            End If
            rstBookMaster.CursorLocation = adUseClient
            Call SetButtons(True)
            SetButtonsForNoRecord
            SSTab1.Tab = 0
        End If
    ElseIf Button.Index = 6 Then
        SSTab1.Tab = 0
        Set DataGrid1.DataSource = Nothing
        rstBookList.ActiveConnection = CxnBookMaster
        Do While Not RefreshRecord(rstBookList)
        Loop
        Set DataGrid1.DataSource = rstBookList
        rstBookList.ActiveConnection = Nothing
        rstBoardList.ActiveConnection = CxnBookMaster
        Do While Not RefreshRecord(rstBoardList)
        Loop
        rstBoardList.ActiveConnection = Nothing
        rstClassList.ActiveConnection = CxnBookMaster
        Do While Not RefreshRecord(rstClassList)
        Loop
        rstClassList.ActiveConnection = Nothing
        rstSubjectList.ActiveConnection = CxnBookMaster
        Do While Not RefreshRecord(rstSubjectList)
        Loop
        rstSubjectList.ActiveConnection = Nothing
        rstGroupList.ActiveConnection = CxnBookMaster
        Do While Not RefreshRecord(rstGroupList)
        Loop
        rstGroupList.ActiveConnection = Nothing
        rstSizeList.ActiveConnection = CxnBookMaster
        Do While Not RefreshRecord(rstSizeList)
        Loop
        rstSizeList.ActiveConnection = Nothing
        rstBindingTypeList.ActiveConnection = CxnBookMaster
        Do While Not RefreshRecord(rstBindingTypeList)
        Loop
        rstBindingTypeList.ActiveConnection = Nothing
        rstLaminationTypeList.ActiveConnection = CxnBookMaster
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
        If rstBookList.RecordCount > 0 Then rstBookList.MoveFirst
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 14 Then
        If rstBookList.RecordCount > 0 Then
           rstBookList.MovePrevious
           If rstBookList.BOF Then
              rstBookList.MoveNext
           End If
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 15 Then
        If rstBookList.RecordCount > 0 Then
           rstBookList.MoveNext
           If rstBookList.EOF Then
              rstBookList.MovePrevious
           End If
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 16 Then
        If rstBookList.RecordCount > 0 Then rstBookList.MoveLast
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 18 Then
        Call CloseForm(FrmBookMaster)
        HiLiteRecord = False
    End If
    If HiLiteRecord Then
        If Not (rstBookList.EOF Or rstBookList.BOF) Then
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
    rstBookList.Sort = "[" + SortOrder & "] Asc"
    DataGrid1.ClearSelCols
    If Not (rstBookList.EOF Or rstBookList.BOF) Then
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
    Mh3dFrame2.Enabled = Not bVal
    Mh3dFrame3.Enabled = False
End Sub
Private Sub SetButtonsForNoRecord()
    If rstBookList.RecordCount = 0 Then
        Toolbar1.Buttons.Item(2).Enabled = False
        Toolbar1.Buttons.Item(3).Enabled = False
        Toolbar1.Buttons.Item(13).Enabled = False
        Toolbar1.Buttons.Item(14).Enabled = False
        Toolbar1.Buttons.Item(15).Enabled = False
        Toolbar1.Buttons.Item(16).Enabled = False
    End If
End Sub
Private Sub Text2_Validate(Cancel As Boolean)
    If rstBookMaster.EOF Or rstBookMaster.BOF Then Exit Sub
    If CheckEmpty(Text2, True) Then
        Cancel = True
    ElseIf CheckDuplicate(CxnBookMaster, "BookMaster", "Code", "Name", Text2.Text, rstBookMaster.Fields("Code").Value, False) Then
        Cancel = True
    ElseIf CheckEmpty(Text3, False) Then
        Text3.Text = Text2.Text
    End If
End Sub
Private Sub Text4_Validate(Cancel As Boolean)
    If CheckEmpty(Text4.Text, False) Then Exit Sub
    If rstBookMaster.EOF Or rstBookMaster.BOF Then Exit Sub
    If CheckDuplicate(CxnBookMaster, "BookMaster", "Code", "Isbn", Text4.Text, rstBookMaster.Fields("Code").Value, False) Then
        Cancel = True
    ElseIf Len(Text4.Text) = 13 Then
        If Not bVerifySum10(Text4.Text) Then Cancel = True
    ElseIf Len(Text4.Text) = 17 Then
        If Not bVerifySum13(Text4.Text) Then Cancel = True
    End If
End Sub
Private Sub Text11_Validate(Cancel As Boolean)
    If CheckEmpty(Text11, True) Then
        Cancel = True
    End If
End Sub
Private Sub Text5_Change()
    If Text5.Text = " " Then
        Text5.Text = "?"
        Sendkeys "{TAB}"
    End If
End Sub
Private Sub Text5_Validate(Cancel As Boolean)
    Dim SearchString As String
    
    SearchString = FixQuote(Text5.Text)
    If rstBoardList.RecordCount = 0 Then
       DisplayError ("No Record in Board Master")
       Cancel = True
       Exit Sub
    Else
       rstBoardList.MoveFirst
    End If
    rstBoardList.Find "[Col0] = '" & RTrim(SearchString) & "'"
    If rstBoardList.EOF Then
       SelectionType = "S"
       BoardCode = ""
       Call LoadSelectionList(rstBoardList, "List of Boards...", "Name")
       SearchOrder = 0
       Call DisplaySelectionList(Text5, BoardCode)
       Call CloseForm(FrmSelectionList)
       If CheckEmpty(Text5.Text, False) Then
          Text5.Text = "?"
       End If
       If RTrim(BoardCode) <> "" Then
          Sendkeys "{TAB}"
       End If
       Cancel = True
    Else
       BoardCode = rstBoardList.Fields("Code").Value
    End If
End Sub
Private Sub Text6_Change()
    If Text6.Text = " " Then
        Text6.Text = "?"
        Sendkeys "{TAB}"
    End If
End Sub
Private Sub Text6_Validate(Cancel As Boolean)
    Dim SearchString As String
    
    SearchString = FixQuote(Text6.Text)
    If rstClassList.RecordCount = 0 Then
       DisplayError ("No Record in Class Master")
       Cancel = True
       Exit Sub
    Else
       rstClassList.MoveFirst
    End If
    rstClassList.Find "[Col0] = '" & RTrim(SearchString) & "'"
    If rstClassList.EOF Then
       SelectionType = "S"
       ClassCode = ""
       Call LoadSelectionList(rstClassList, "List of Classes...", "Name")
       SearchOrder = 0
       Call DisplaySelectionList(Text6, ClassCode)
       Call CloseForm(FrmSelectionList)
       If CheckEmpty(Text6.Text, False) Then
          Text6.Text = "?"
       End If
       If RTrim(ClassCode) <> "" Then
          Sendkeys "{TAB}"
       End If
       Cancel = True
    Else
       ClassCode = rstClassList.Fields("Code").Value
    End If
End Sub
Private Sub Text7_Change()
    If Text7.Text = " " Then
        Text7.Text = "?"
        Sendkeys "{TAB}"
    ElseIf CheckEmpty(Text7, False) Then
        SubjectCode = ""
    End If
End Sub
Private Sub Text7_Validate(Cancel As Boolean)
    Dim SearchString As String
    
    If CheckEmpty(Text7, False) Then
        Exit Sub
    End If
    SearchString = FixQuote(Text7.Text)
    If rstSubjectList.RecordCount = 0 Then
       DisplayError ("No Record in Subject Master")
       Cancel = True
       Exit Sub
    Else
       rstSubjectList.MoveFirst
    End If
    rstSubjectList.Find "[Col0] = '" & RTrim(SearchString) & "'"
    If rstSubjectList.EOF Then
       SelectionType = "S"
       SubjectCode = ""
       Call LoadSelectionList(rstSubjectList, "List of Subjects...", "Name")
       SearchOrder = 0
       Call DisplaySelectionList(Text7, SubjectCode)
       Call CloseForm(FrmSelectionList)
       If CheckEmpty(Text7.Text, False) Then
          Text7.Text = "?"
       End If
       If RTrim(SubjectCode) <> "" Then
          Sendkeys "{TAB}"
       End If
       Cancel = True
    Else
       SubjectCode = rstSubjectList.Fields("Code").Value
    End If
End Sub
Private Sub Text8_Change()
    If Text8.Text = " " Then
        Text8.Text = "?"
        Sendkeys "{TAB}"
    End If
End Sub
Private Sub Text8_Validate(Cancel As Boolean)
    Dim SearchString As String
    
    SearchString = FixQuote(Text8.Text)
    If rstGroupList.RecordCount = 0 Then
       DisplayError ("No Record in Group Master")
       Cancel = True
       Exit Sub
    Else
       rstGroupList.MoveFirst
    End If
    rstGroupList.Find "[Col0] = '" & RTrim(SearchString) & "'"
    If rstGroupList.EOF Then
       SelectionType = "S"
       GroupCode = ""
       Call LoadSelectionList(rstGroupList, "List of Groups...", "Name")
       SearchOrder = 0
       Call DisplaySelectionList(Text8, GroupCode)
       Call CloseForm(FrmSelectionList)
       If CheckEmpty(Text8.Text, False) Then
          Text8.Text = "?"
       End If
       If RTrim(GroupCode) <> "" Then
          Sendkeys "{TAB}"
       End If
       Cancel = True
    Else
       GroupCode = rstGroupList.Fields("Code").Value
    End If
End Sub
Private Sub Text9_Change()
    If Text9.Text = " " Then
        Text9.Text = "?"
        Sendkeys "{TAB}"
    End If
End Sub
Private Sub Text9_Validate(Cancel As Boolean)
    Dim SearchString As String
    
    SearchString = FixQuote(Text9.Text)
    If rstSizeList.RecordCount = 0 Then
       DisplayError ("No Record in Size Master")
       Cancel = True
       Exit Sub
    Else
       rstSizeList.MoveFirst
    End If
    rstSizeList.Find "[Col0] = '" & RTrim(SearchString) & "'"
    If rstSizeList.EOF Then
       SelectionType = "S"
       SizeCode = ""
       Call LoadSelectionList(rstSizeList, "List of Sizes...", "Name")
       SearchOrder = 0
       Call DisplaySelectionList(Text9, SizeCode)
       Call CloseForm(FrmSelectionList)
       If CheckEmpty(Text9.Text, False) Then
          Text9.Text = "?"
       End If
       If RTrim(SizeCode) <> "" Then
          Sendkeys "{TAB}"
       End If
       Cancel = True
    Else
       SizeCode = rstSizeList.Fields("Code").Value
    End If
End Sub
Private Sub Combo1_Click()
    Dim Pages As Variant
    Dim Forms As Double, TotalPages As Long, TotalForms As Double, i As Integer
    
    TotalPages = 0
    TotalForms = 0
    For i = 1 To fpSpread2.DataRowCnt
        fpSpread2.GetText 2, i, Pages
        Forms = Val(Pages) / Val(Combo1.Text)
        fpSpread2.SetText 7, i, Forms
        TotalPages = TotalPages + Pages
        TotalForms = TotalForms + Forms
    Next
    MhRealInput15.Text = Format(TotalPages, "0")
    MhRealInput7.Text = Format(TotalForms, "0.00")
End Sub
Private Sub Text10_Change()
    If Text10.Text = " " Then
        Text10.Text = "?"
        Sendkeys "{TAB}"
    ElseIf CheckEmpty(Text10, False) Then
        BindingTypeCode = ""
    End If
End Sub
Private Sub Text10_Validate(Cancel As Boolean)
    Dim SearchString As String
    
    If CheckEmpty(Text10, False) Then
        Exit Sub
    End If
    SearchString = FixQuote(Text10.Text)
    If rstBindingTypeList.RecordCount = 0 Then
       DisplayError ("No Record in Binding Type Master")
       Cancel = True
       Exit Sub
    Else
       rstBindingTypeList.MoveFirst
    End If
    rstBindingTypeList.Find "[Col0] = '" & RTrim(SearchString) & "'"
    If rstBindingTypeList.EOF Then
       SelectionType = "S"
       BindingTypeCode = ""
       Call LoadSelectionList(rstBindingTypeList, "List of Binding Types...", "Name")
       SearchOrder = 0
       Call DisplaySelectionList(Text10, BindingTypeCode)
       Call CloseForm(FrmSelectionList)
       If CheckEmpty(Text10.Text, False) Then
          Text10.Text = "?"
       End If
       If RTrim(BindingTypeCode) <> "" Then
          Sendkeys "{TAB}"
       End If
       Cancel = True
    Else
       BindingTypeCode = rstBindingTypeList.Fields("Code").Value
    End If
End Sub
Private Sub Text12_Change()
    If Text12.Text = " " Then
        Text12.Text = "?"
        Sendkeys "{TAB}"
    ElseIf CheckEmpty(Text12, False) Then
        LaminationTypeCode = ""
    End If
End Sub
Private Sub Text12_Validate(Cancel As Boolean)
    Dim SearchString As String
    
    If CheckEmpty(Text12, False) Then
        Exit Sub
    End If
    SearchString = FixQuote(Text12.Text)
    If rstLaminationTypeList.RecordCount = 0 Then
       DisplayError ("No Record in Binding Type Master")
       Cancel = True
       Exit Sub
    Else
       rstLaminationTypeList.MoveFirst
    End If
    rstLaminationTypeList.Find "[Col0] = '" & RTrim(SearchString) & "'"
    If rstLaminationTypeList.EOF Then
       SelectionType = "S"
       LaminationTypeCode = ""
       Call LoadSelectionList(rstLaminationTypeList, "List of Lamination Types...", "Name")
       SearchOrder = 0
       Call DisplaySelectionList(Text12, LaminationTypeCode)
       Call CloseForm(FrmSelectionList)
       If CheckEmpty(Text12.Text, False) Then
          Text12.Text = "?"
       End If
       If RTrim(LaminationTypeCode) <> "" Then
          Sendkeys "{TAB}"
       End If
       Cancel = True
    Else
       LaminationTypeCode = rstLaminationTypeList.Fields("Code").Value
    End If
End Sub
Private Sub ViewRecord()
    ClearFields
    If rstBookList.EOF Then Exit Sub
    FindRecord
    LoadFields
End Sub
Private Sub FindRecord()
    If rstBookMaster.State = adStateOpen Then
       rstBookMaster.Close
    End If
    rstBookMaster.Open "Select * From BookMaster Where Code = '" & FixQuote(rstBookList.Fields("Code").Value) & "'", CxnBookMaster, adOpenKeyset, adLockOptimistic
    If rstBookMaster.RecordCount = 0 Then
       Call DisplayError("This Record has been deleted by Another User ! Click Ok To Refresh the Recordset")
       Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
    End If
End Sub
Private Sub ClearFields()
    fpSpread1.ClearRange 1, 1, fpSpread1.MaxCols, fpSpread1.MaxRows, True
    fpSpread2.ClearRange 1, 1, fpSpread2.MaxCols, fpSpread2.MaxRows, True
    fpSpread3.ClearRange 1, 1, fpSpread3.MaxCols, fpSpread3.MaxRows, True
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text11.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text7.Text = ""
    Text8.Text = ""
    Text9.Text = ""
    Text10.Text = ""
    Text12.Text = ""
    Text13.Text = ""
    Text41.Text = ""
    MhRealInput1.Text = "0.00"
    MhRealInput3.Text = "0.00"
    MhRealInput16.Text = "0.00"
    MhRealInput17.Text = "4"
    MhRealInput18.Text = "0"
    MhRealInput4.Text = "0"
    MhRealInput19.Text = "0"
    MhRealInput7.Text = 0#
    MhRealInput15.Text = "0"
    MhRealInput5.Text = "0"
    
    MhRealInput52.Text = "0"
    MhRealInput8.Text = "0"
    MhRealInput6.Text = "0"
    MhRealInput10.Text = "0.00"
    Option1.Value = True
    Option2.Value = False
    Combo1.ListIndex = 0
    Combo7.ListIndex = 1
    
    ClassCode = ""
    SubjectCode = ""
    GroupCode = ""
    BindingTypeCode = ""
    LaminationTypeCode = ""
End Sub
Private Sub LoadFields()
    If rstBookMaster.EOF Or rstBookMaster.BOF Then Exit Sub
    Text2.Text = rstBookMaster.Fields("Name").Value
    Text3.Text = rstBookMaster.Fields("PrintName").Value
    Text4.Text = rstBookMaster.Fields("Isbn").Value
    Text11.Text = rstBookMaster.Fields("BusyCode").Value
    MhRealInput1.Text = Format(Val(rstBookMaster.Fields("Price").Value), "0.00")
    BoardCode = rstBookMaster.Fields("Board").Value
    rstBoardList.MoveFirst
    rstBoardList.Find "[Code] = '" & BoardCode & "'"
    Text5.Text = rstBoardList.Fields("Col0").Value
    ClassCode = rstBookMaster.Fields("Class").Value
    If rstClassList.RecordCount > 0 Then rstClassList.MoveFirst
    rstClassList.Find "[Code] = '" & ClassCode & "'"
    If Not rstClassList.EOF Then
        Text6.Text = rstClassList.Fields("Col0").Value
    End If
    SubjectCode = rstBookMaster.Fields("Subject").Value
    If rstSubjectList.RecordCount > 0 Then rstSubjectList.MoveFirst
    rstSubjectList.Find "[Code] = '" & SubjectCode & "'"
    If Not rstSubjectList.EOF Then
        Text7.Text = rstSubjectList.Fields("Col0").Value
    End If
    GroupCode = rstBookMaster.Fields("Group").Value
    If rstGroupList.RecordCount > 0 Then rstGroupList.MoveFirst
    rstGroupList.Find "[Code] = '" & GroupCode & "'"
    If Not rstGroupList.EOF Then
        Text8.Text = rstGroupList.Fields("Col0").Value
    End If
    SizeCode = rstBookMaster.Fields("Size").Value
    rstSizeList.MoveFirst
    rstSizeList.Find "[Code] = '" & SizeCode & "'"
    Text9.Text = rstSizeList.Fields("Col0").Value
    BindingTypeCode = rstBookMaster.Fields("BindingType").Value
    If rstBindingTypeList.RecordCount > 0 Then rstBindingTypeList.MoveFirst
    rstBindingTypeList.Find "[Code] = '" & BindingTypeCode & "'"
    If Not rstBindingTypeList.EOF Then
        Text10.Text = rstBindingTypeList.Fields("Col0").Value
    End If
    LaminationTypeCode = rstBookMaster.Fields("LaminationType").Value
    If rstLaminationTypeList.RecordCount > 0 Then rstLaminationTypeList.MoveFirst
    rstLaminationTypeList.Find "[Code] = '" & LaminationTypeCode & "'"
    If Not rstLaminationTypeList.EOF Then
        Text12.Text = rstLaminationTypeList.Fields("Col0").Value
    End If
    MhRealInput3.Text = Format(Val(rstBookMaster.Fields("AddOnRate01").Value), "0.00")
    MhRealInput16.Text = Format(Val(rstBookMaster.Fields("AddOnRate02").Value), "0.00")
    Combo1.ListIndex = IIf(Val(rstBookMaster.Fields("FormType").Value) = 1, 2, IIf(Val(rstBookMaster.Fields("FormType").Value) = 2, 4, IIf(Val(rstBookMaster.Fields("FormType").Value) = 3, 0, IIf(Val(rstBookMaster.Fields("FormType").Value) = 4, 3, IIf(Val(rstBookMaster.Fields("FormType").Value) = 8, 1, Val(rstBookMaster.Fields("FormType").Value))))))
    MhRealInput4.Text = Format(Val(rstBookMaster.Fields("BindingForms01").Value), "0")
    MhRealInput19.Text = Format(Val(rstBookMaster.Fields("BindingForms02").Value), "0")
    MhRealInput15.Text = Format(Val(rstBookMaster.Fields("Pages").Value), "0")
    MhRealInput7.Text = Format(Val(rstBookMaster.Fields("Forms").Value), "0.00")
    fpSpread2.SetText 1, 1, IIf(rstBookMaster.Fields("OneColorPlateType").Value = "1", "Deepatch", IIf(rstBookMaster.Fields("OneColorPlateType").Value = "2", "PS", IIf(rstBookMaster.Fields("OneColorPlateType").Value = "3", "Wipeon", "CTP")))
    fpSpread2.SetText 2, 1, Val(rstBookMaster.Fields("OneColorPages").Value)
    fpSpread2.SetText 3, 1, Val(rstBookMaster.Fields("OneColor¼Forms").Value)
    fpSpread2.SetText 4, 1, Val(rstBookMaster.Fields("OneColor½Forms").Value)
    fpSpread2.SetText 5, 1, Val(rstBookMaster.Fields("OneColor1F/BForms").Value)
    fpSpread2.SetText 6, 1, Val(rstBookMaster.Fields("OneColor1W/TForms").Value)
    fpSpread2.SetText 7, 1, Val(rstBookMaster.Fields("OneColorForms").Value)
    fpSpread2.SetText 1, 2, IIf(rstBookMaster.Fields("TwoColorPlateType").Value = "1", "Deepatch", IIf(rstBookMaster.Fields("TwoColorPlateType").Value = "2", "PS", IIf(rstBookMaster.Fields("TwoColorPlateType").Value = "3", "Wipeon", "CTP")))
    fpSpread2.SetText 2, 2, Val(rstBookMaster.Fields("TwoColorPages").Value)
    fpSpread2.SetText 3, 2, Val(rstBookMaster.Fields("TwoColor¼Forms").Value)
    fpSpread2.SetText 4, 2, Val(rstBookMaster.Fields("TwoColor½Forms").Value)
    fpSpread2.SetText 5, 2, Val(rstBookMaster.Fields("TwoColor1F/BForms").Value)
    fpSpread2.SetText 6, 2, Val(rstBookMaster.Fields("TwoColor1W/TForms").Value)
    fpSpread2.SetText 7, 2, Val(rstBookMaster.Fields("TwoColorForms").Value)
    fpSpread2.SetText 1, 3, IIf(rstBookMaster.Fields("FourColorPlateType").Value = "1", "Deepatch", IIf(rstBookMaster.Fields("FourColorPlateType").Value = "2", "PS", IIf(rstBookMaster.Fields("FourColorPlateType").Value = "3", "Wipeon", "CTP")))
    fpSpread2.SetText 2, 3, Val(rstBookMaster.Fields("FourColorPages").Value)
    fpSpread2.SetText 3, 3, Val(rstBookMaster.Fields("FourColor¼Forms").Value)
    fpSpread2.SetText 4, 3, Val(rstBookMaster.Fields("FourColor½Forms").Value)
    fpSpread2.SetText 5, 3, Val(rstBookMaster.Fields("FourColor1F/BForms").Value)
    fpSpread2.SetText 6, 3, Val(rstBookMaster.Fields("FourColor1W/TForms").Value)
    fpSpread2.SetText 7, 3, Val(rstBookMaster.Fields("FourColorForms").Value)
    Combo7.ListIndex = Val(rstBookMaster.Fields("TitlePlateType").Value) - 1
    MhRealInput17.Text = Format(Val(rstBookMaster.Fields("TitleFrontColor").Value), "0")
    MhRealInput18.Text = Format(Val(rstBookMaster.Fields("TitleBackColor").Value), "0")
    MhRealInput5.Text = Format(Val(rstBookMaster.Fields("Qty/Pkt").Value), "0")
    MhRealInput8.Text = Format(Val(rstBookMaster.Fields("LooseQty/Box").Value), "0")
    MhRealInput6.Text = Format(Val(rstBookMaster.Fields("Pkt/Box").Value), "0")
    MhRealInput10.Text = Format(Val(rstBookMaster.Fields("Royalty").Value), "0.00")
    Option1.Value = IIf(rstBookMaster.Fields("DuplexPrinting").Value = "Y", True, False)
    Option2.Value = IIf(rstBookMaster.Fields("DuplexPrinting").Value = "N", True, False)
    MhRealInput52.Value = Format(Val(rstBookMaster.Fields("SaleDiscount").Value), "0")
    Text41.Text = rstBookMaster.Fields("HSNCode").Value
    'Combo3.ListIndex = IIf(rstBookMaster.Fields("Processing").Value = "O", 0, IIf(rstBookMaster.Fields("Processing").Value = "N", 1, 2))
    Text13.Text = rstBookMaster.Fields("Narration").Value
    Call LoadMaterialList(rstBookMaster.Fields("Code").Value)
End Sub
Private Sub EditRecord()
    On Error GoTo ErrorHandler
    
    If rstBookMaster.RecordCount = 0 Then Exit Sub
    If rstBookMaster.State = adStateOpen Then
       rstBookMaster.Close
    End If
    rstBookMaster.CursorLocation = adUseServer
    rstBookMaster.Open "Select * From BookMaster Where Code = '" & FixQuote(rstBookList.Fields("Code").Value) & "'", CxnBookMaster, adOpenKeyset, adLockPessimistic
    MdiMainMenu.MousePointer = vbHourglass
    rstBookMaster.Fields("Printstatus") = "N"
    MdiMainMenu.MousePointer = vbNormal
    AddToList
    Call SetButtons(False)
    SSTab1.TabEnabled(0) = False
    Text2.SetFocus
    blnRecordExist = True
    CxnBookMaster.BeginTrans
    Exit Sub
ErrorHandler:
    If Err.Number = -2147467259 Then
       Call DisplayError("Failed to Edit the record")
    End If
    MdiMainMenu.MousePointer = vbNormal
    SSTab1.Tab = 0
End Sub
Private Sub SaveFields()
    Dim Fld As Variant
    
    If rstBookMaster.EOF Or rstBookMaster.BOF Then Exit Sub
    If Not blnRecordExist Then
        rstBookMaster.Fields("Code").Value = GenerateCode(CxnBookMaster, "Select Max(Code) From BookMaster", 6, "0")
        rstBookMaster.Fields("CreatedBy").Value = UserCode
        rstBookMaster.Fields("CreatedOn").Value = Now()
        rstBookMaster.Fields("Recordstatus").Value = "N"
    Else
        rstBookMaster.Fields("ModifiedBy").Value = UserCode
        rstBookMaster.Fields("ModifiedOn").Value = Now()
        rstBookMaster.Fields("Recordstatus").Value = "M"
    End If
    rstBookMaster.Fields("Name").Value = Trim(Text2.Text)
    rstBookMaster.Fields("PrintName").Value = Trim(Text3.Text)
    rstBookMaster.Fields("Isbn").Value = Trim(Text4.Text)
    rstBookMaster.Fields("BusyCode").Value = Trim(Text11.Text)
    rstBookMaster.Fields("Price").Value = Val(MhRealInput1.Text)
    rstBookMaster.Fields("Board").Value = BoardCode
    rstBookMaster.Fields("Class").Value = ClassCode
    rstBookMaster.Fields("Subject").Value = SubjectCode
    rstBookMaster.Fields("Group").Value = GroupCode
    rstBookMaster.Fields("Size").Value = SizeCode
    rstBookMaster.Fields("BindingType").Value = BindingTypeCode
    rstBookMaster.Fields("LaminationType").Value = LaminationTypeCode
    rstBookMaster.Fields("AddOnRate01").Value = Val(MhRealInput3.Text)
    rstBookMaster.Fields("AddOnRate02").Value = Val(MhRealInput16.Text)
    rstBookMaster.Fields("FormType").Value = IIf(Combo1.ListIndex = 0, "3", IIf(Combo1.ListIndex = 1, "8", IIf(Combo1.ListIndex = 2, "1", IIf(Combo1.ListIndex = 3, "4", IIf(Combo1.ListIndex = 4, "2", Trim(str(Combo1.ListIndex)))))))
    rstBookMaster.Fields("BindingForms01").Value = Val(MhRealInput4.Text)
    rstBookMaster.Fields("BindingForms02").Value = Val(MhRealInput19.Text)
    rstBookMaster.Fields("Pages").Value = Val(MhRealInput15.Text)
    rstBookMaster.Fields("Forms").Value = Val(MhRealInput7.Text)
    fpSpread2.GetText 1, 1, Fld
    rstBookMaster.Fields("OneColorPlateType").Value = IIf(Trim(Fld) = "Deepatch", "1", IIf(Trim(Fld) = "PS", "2", IIf(Trim(Fld) = "Wipeon", "3", "4")))
    fpSpread2.GetText 2, 1, Fld
    rstBookMaster.Fields("OneColorPages").Value = Val(Fld)
    fpSpread2.GetText 3, 1, Fld
    rstBookMaster.Fields("OneColor¼Forms").Value = Val(Fld)
    fpSpread2.GetText 4, 1, Fld
    rstBookMaster.Fields("OneColor½Forms").Value = Val(Fld)
    fpSpread2.GetText 5, 1, Fld
    rstBookMaster.Fields("OneColor1F/BForms").Value = Val(Fld)
    fpSpread2.GetText 6, 1, Fld
    rstBookMaster.Fields("OneColor1W/TForms").Value = Val(Fld)
    fpSpread2.GetText 7, 1, Fld
    rstBookMaster.Fields("OneColorForms").Value = Val(Fld)
    fpSpread2.GetText 1, 2, Fld
    rstBookMaster.Fields("TwoColorPlateType").Value = IIf(Trim(Fld) = "Deepatch", "1", IIf(Trim(Fld) = "PS", "2", IIf(Trim(Fld) = "Wipeon", "3", "4")))
    fpSpread2.GetText 2, 2, Fld
    rstBookMaster.Fields("TwoColorPages").Value = Val(Fld)
    fpSpread2.GetText 3, 2, Fld
    rstBookMaster.Fields("TwoColor¼Forms").Value = Val(Fld)
    fpSpread2.GetText 4, 2, Fld
    rstBookMaster.Fields("TwoColor½Forms").Value = Val(Fld)
    fpSpread2.GetText 5, 2, Fld
    rstBookMaster.Fields("TwoColor1F/BForms").Value = Val(Fld)
    fpSpread2.GetText 6, 2, Fld
    rstBookMaster.Fields("TwoColor1W/TForms").Value = Val(Fld)
    fpSpread2.GetText 7, 2, Fld
    rstBookMaster.Fields("TwoColorForms").Value = Val(Fld)
    fpSpread2.GetText 1, 3, Fld
    rstBookMaster.Fields("FourColorPlateType").Value = IIf(Trim(Fld) = "Deepatch", "1", IIf(Trim(Fld) = "PS", "2", IIf(Trim(Fld) = "Wipeon", "3", "4")))
    fpSpread2.GetText 2, 3, Fld
    rstBookMaster.Fields("FourColorPages").Value = Val(Fld)
    fpSpread2.GetText 3, 3, Fld
    rstBookMaster.Fields("FourColor¼Forms").Value = Val(Fld)
    fpSpread2.GetText 4, 3, Fld
    rstBookMaster.Fields("FourColor½Forms").Value = Val(Fld)
    fpSpread2.GetText 5, 3, Fld
    rstBookMaster.Fields("FourColor1F/BForms").Value = Val(Fld)
    fpSpread2.GetText 6, 3, Fld
    rstBookMaster.Fields("FourColor1W/TForms").Value = Val(Fld)
    fpSpread2.GetText 7, 3, Fld
    rstBookMaster.Fields("FourColorForms").Value = Val(Fld)
    rstBookMaster.Fields("TitlePlateType").Value = Trim(str(Combo7.ListIndex + 1))
    rstBookMaster.Fields("TitleFrontColor").Value = Val(MhRealInput17.Text)
    rstBookMaster.Fields("TitleBackColor").Value = Val(MhRealInput18.Text)
    rstBookMaster.Fields("Qty/Pkt").Value = Val(MhRealInput5.Text)
    rstBookMaster.Fields("LooseQty/Box").Value = Val(MhRealInput8.Text)
    rstBookMaster.Fields("Pkt/Box").Value = Val(MhRealInput6.Text)
    rstBookMaster.Fields("Royalty").Value = Val(MhRealInput10.Text)
    rstBookMaster.Fields("DuplexPrinting").Value = IIf(Option1.Value, "Y", "N")
    
    rstBookMaster.Fields("SaleDiscount").Value = Val(MhRealInput52.Text)
    'rstBookMaster.Fields("Processing").Value = IIf(Combo3.ListIndex = 0, "O", IIf(Combo3.ListIndex = 1, "N", "R"))
    rstBookMaster.Fields("HSNCode").Value = Text41.Text
    
    rstBookMaster.Fields("Narration").Value = Text13.Text
    rstBookMaster.Fields("Type").Value = BookType
    rstBookMaster.Fields("PrintStatus").Value = "N"
End Sub
Private Sub AddToList()
    On Error Resume Next
    rstBookList.MoveFirst
    rstBookList.Find "[Code] = '" & rstBookMaster.Fields("Code").Value & "'"
    If rstBookList.EOF Then rstBookList.AddNew:               rstBookList.Fields("Code").Value = rstBookMaster.Fields("Code").Value
    rstBookList.Fields("Name").Value = rstBookMaster.Fields("Name").Value
    rstBookList.Fields("BusyCode").Value = rstBookMaster.Fields("BusyCode").Value
    rstBookList.Update
    rstBookList.Sort = SortOrder & " Asc"
    rstBookList.Find "[Code] = '" & rstBookMaster.Fields("Code").Value & "'"
End Sub
Private Function CheckMandatoryFields() As Boolean
    If CheckEmpty(Text2.Text, False) Then
        SSTab1.Tab = 1
        Text2.SetFocus
        CheckMandatoryFields = True
    ElseIf CheckDuplicate(CxnBookMaster, "BookMaster", "Code", "Name", Text2.Text, rstBookMaster.Fields("Code").Value, False) Then
        SSTab1.Tab = 1
        Text2.SetFocus
        CheckMandatoryFields = True
    ElseIf CheckEmpty(Text3.Text, False) Then
        SSTab1.Tab = 1
        Text3.SetFocus
        CheckMandatoryFields = True
    ElseIf CheckEmpty(Text11.Text, False) Then
        SSTab1.Tab = 1
        Text11.SetFocus
        CheckMandatoryFields = True
    ElseIf CheckEmpty(Text5.Text, False) Then
        SSTab1.Tab = 1
        Text5.SetFocus
        CheckMandatoryFields = True
    ElseIf Not CheckExists(Text5, "Col0", rstBoardList, BoardCode) Then
        SSTab1.Tab = 1
        Text5.SetFocus
        CheckMandatoryFields = True
    ElseIf CheckEmpty(Text6.Text, False) Then
        SSTab1.Tab = 1
        Text6.SetFocus
        CheckMandatoryFields = True
    ElseIf Not CheckExists(Text6, "Col0", rstClassList, ClassCode) Then
        SSTab1.Tab = 1
        Text6.SetFocus
        CheckMandatoryFields = True
    ElseIf CheckEmpty(Text8.Text, False) Then
        SSTab1.Tab = 1
        Text8.SetFocus
        CheckMandatoryFields = True
    ElseIf Not CheckExists(Text8, "Col0", rstGroupList, GroupCode) Then
        SSTab1.Tab = 1
        Text8.SetFocus
        CheckMandatoryFields = True
    ElseIf CheckEmpty(Text9.Text, False) Then
        SSTab1.Tab = 1
        Text9.SetFocus
        CheckMandatoryFields = True
    ElseIf Not CheckExists(Text9, "Col0", rstSizeList, SizeCode) Then
        SSTab1.Tab = 1
        Text9.SetFocus
        CheckMandatoryFields = True
    Else
        If Not CheckEmpty(Text4.Text, False) Then
            If CheckDuplicate(CxnBookMaster, "BookMaster", "Code", "Isbn", Text4.Text, rstBookMaster.Fields("Code").Value, False) Then
                SSTab1.Tab = 1
                Text4.SetFocus
                CheckMandatoryFields = True
            End If
        End If
        If Not CheckEmpty(Text7.Text, False) Then
            If Not CheckExists(Text7, "Col0", rstSubjectList, SubjectCode) Then
                SSTab1.Tab = 1
                Text7.SetFocus
                CheckMandatoryFields = True
            End If
        End If
        If Not CheckEmpty(Text10.Text, False) Then
            If Not CheckExists(Text10, "Col0", rstBindingTypeList, BindingTypeCode) Then
                SSTab1.Tab = 1
                Text10.SetFocus
                CheckMandatoryFields = True
            End If
        End If
        If Not CheckEmpty(Text12.Text, False) Then
            If Not CheckExists(Text12, "Col0", rstLaminationTypeList, LaminationTypeCode) Then
                SSTab1.Tab = 1
                Text12.SetFocus
                CheckMandatoryFields = True
            End If
        End If
        If CheckForms() Then
            SSTab1.Tab = 1
            fpSpread2.SetFocus
            CheckMandatoryFields = True
        End If
        If CheckItem() Then
            SSTab1.Tab = 2
            fpSpread1.SetFocus
            CheckMandatoryFields = True
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
Public Sub FilterRecord(ByVal SrchFor As String, ByVal SrchText As String)
    If SrchFor = "Name" Then
        rstBookList.Filter = "[Name] Like '%" & SrchText & "%'"
    End If
End Sub
Private Sub fpSpread1_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask And KeyCode = vbKeyD Then
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            fpSpread1.DeleteRows fpSpread1.ActiveRow, 1
            fpSpread1.SetFocus
        End If
    ElseIf Shift = 0 And KeyCode = vbKeyF5 Then
        Call RefreshDropDownList("R")
    End If
End Sub
Private Sub fpSpread3_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask And KeyCode = vbKeyD Then
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            fpSpread3.DeleteRows fpSpread3.ActiveRow, 1
            fpSpread3.SetFocus
        End If
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
        fpSpread1.TypeComboBoxList = IIf(Category = "Outsource Item", OutsourceItem, FreshBook)
    ElseIf Col = 2 Then
        If Category = "Outsource Item" Then
           If rstOutsourceItemList.RecordCount > 0 Then rstOutsourceItemList.MoveFirst
           rstOutsourceItemList.Find "[Name]='" & FixQuote(ActiveCellVal) & "'"
           If Not rstOutsourceItemList.EOF Then
                fpSpread1.SetText 4, Row, rstOutsourceItemList.Fields("NCode").Value
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
Private Sub fpSpread2_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim ActiveCellVal As Variant, Forms As Variant
    Dim i As Integer, TotalForms As Double
    
    fpSpread2.GetText Col, Row, ActiveCellVal
    If ActiveCellVal = "" Then
        Cancel = True
        Exit Sub
    End If
    If Col = 2 Then
        Combo1_Click
    ElseIf Col = 3 Or Col = 4 Or Col = 5 Or Col = 6 Then   'Calculate Binding Forms
        TotalForms = 0
        For i = 1 To 3
            fpSpread2.GetText 3, i, Forms
            TotalForms = TotalForms + Forms
            fpSpread2.GetText 4, i, Forms
            TotalForms = TotalForms + Forms
            fpSpread2.GetText 5, i, Forms
            If Combo1.ListIndex <= 2 Then
                Forms = Val(Forms) / 2
                Forms = Int(Forms) + IIf(Val(Forms) = Int(Val(Forms)), 0, 1)
            End If
            TotalForms = TotalForms + Forms
            fpSpread2.GetText 6, i, Forms
            TotalForms = TotalForms + Forms
        Next
        MhRealInput4.Text = Format(TotalForms, "0")
    End If
End Sub
Private Sub fpSpread3_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim ActiveCellVal As Variant, Forms As Variant
    
    fpSpread3.GetText Col, Row, ActiveCellVal
    If Col <= 2 Then
        If ActiveCellVal = "" Then
            Cancel = True
            Exit Sub
        End If
        If Col = 1 Then
            If Not IsDate(ActiveCellVal) Then
                Cancel = True
                Exit Sub
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
        Else
            If Left(Item, 1) <> "3" Then
                CheckItem = True
            End If
        End If
        If CheckItem Then
            DisplayError "Data mismatch in row #" & Trim(str(i))
            Exit For
        End If
    Next
End Function
Private Function CheckForms() As Boolean
    Dim i As Integer
    Dim Pages As Variant, Forms¼ As Variant, Forms½ As Variant, Forms1FB As Variant, Forms1WT As Variant, TotalForms As Variant
    
    CheckForms = False
    
    For i = 1 To fpSpread2.DataRowCnt
        fpSpread2.SetActiveCell 1, i
        fpSpread2.GetText 2, i, Pages
        fpSpread2.GetText 7, i, TotalForms
        If Pages / Val(Combo1.Text) <> TotalForms Then
            CheckForms = True
        End If
        If Not CheckForms Then
            fpSpread2.GetText 3, i, Forms¼
            fpSpread2.GetText 4, i, Forms½
            fpSpread2.GetText 5, i, Forms1FB
            fpSpread2.GetText 6, i, Forms1WT
            If Val(Forms¼) * 0.25 + Val(Forms½) * 0.5 + Val(Forms1FB) + Val(Forms1WT) <> TotalForms Then
                CheckForms = True
            End If
        End If
        If CheckForms Then
            DisplayError "Data mismatch in row #" & Trim(str(i))
            Exit For
        End If
    Next
    If Not CheckForms Then
        TotalForms = 0
        For i = 1 To 3
            fpSpread2.GetText 3, i, Forms¼
            fpSpread2.GetText 4, i, Forms½
            fpSpread2.GetText 5, i, Forms1FB
            If Combo1.ListIndex <= 2 Then
                Forms1FB = Val(Forms1FB) / 2
                Forms1FB = Int(Forms1FB) + IIf(Val(Forms1FB) = Int(Val(Forms1FB)), 0, 1)
            End If
            fpSpread2.GetText 6, i, Forms1WT
            TotalForms = TotalForms + Val(Forms¼) + Val(Forms½) + Val(Forms1FB) + Val(Forms1WT)
        Next
        If Val(MhRealInput4.Text) <> TotalForms Then
            DisplayError "Printing & Binding Forms Mismatch"
            CheckForms = True
        End If
    End If
End Function
Private Sub LoadMaterialList(ByVal strBookCode As String)
    Dim i As Integer
    On Error GoTo ErrorHandler
    
    If rstBookChild.State = adStateOpen Then
       rstBookChild.Close
    End If
    rstBookChild.Open "SELECT Category,IIF(Category='1',(SELECT Name FROM OutsourceItemMaster WHERE Code=T.Item),(SELECT Name FROM BookMaster WHERE Code=T.Item)) As ItemName,Quantity,Category+Item As ItemCode FROM BookChild01 T WHERE Code='" & strBookCode & "' ORDER BY Category", CxnBookMaster, adOpenKeyset, adLockReadOnly
    rstBookChild.ActiveConnection = Nothing
    If rstBookChild.RecordCount > 0 Then rstBookChild.MoveFirst
    i = 0
    Do While Not rstBookChild.EOF
        i = i + 1
        With fpSpread1
            .SetText 1, i, IIf(rstBookChild.Fields("Category").Value = "1", "Outsource Item", "Fresh Book")
            .Col = 2
            .TypeComboBoxList = IIf(rstBookChild.Fields("Category").Value = "1", OutsourceItem, FreshBook)
            .SetText 2, i, rstBookChild.Fields("ItemName").Value
            .SetText 3, i, Val(rstBookChild.Fields("Quantity").Value)
            .SetText 4, i, rstBookChild.Fields("ItemCode").Value
        End With
        rstBookChild.MoveNext
    Loop
    If rstBookChild.State = adStateOpen Then
       rstBookChild.Close
    End If
    rstBookChild.Open "SELECT ArrivedOn,Correction,RectifiedOn FROM BookChild02 T WHERE Code='" & strBookCode & "' AND Department='P' ORDER BY ArrivedOn,SNo", CxnBookMaster, adOpenKeyset, adLockReadOnly
    rstBookChild.ActiveConnection = Nothing
    If rstBookChild.RecordCount > 0 Then rstBookChild.MoveFirst
    i = 0
    Do While Not rstBookChild.EOF
        i = i + 1
        With fpSpread3
            .SetText 1, i, Format(rstBookChild.Fields("ArrivedOn").Value, "dd-mm-yyyy")
            .SetText 2, i, rstBookChild.Fields("Correction").Value
            .SetText 3, i, rstBookChild.Fields("RectifiedOn").Value
        End With
        rstBookChild.MoveNext
    Loop
    Exit Sub
ErrorHandler:
    DisplayError ("Failed to Load Material/Correction List")
End Sub
Private Function UpdateMaterialList(ByVal ActionType As String) As Boolean
    Dim CellVal(1 To 3) As Variant
    On Error GoTo ErrorHandler
    
    UpdateMaterialList = True
    If Left(ActionType, 1) = "D" And (Not blnRecordExist) Then Exit Function
    If ActionType = "D1" Then
        CxnBookMaster.Execute "Delete From BookChild01 Where Code = '" & rstBookMaster.Fields("Code").Value & "'"
    ElseIf ActionType = "D2" Then
        CxnBookMaster.Execute "DELETE FROM BookChild02 WHERE Code='" & rstBookMaster.Fields("Code").Value & "' AND Department='P'"
    ElseIf ActionType = "I1" Then
        With fpSpread1
            .GetText 1, .ActiveRow, CellVal(1)
            .GetText 3, .ActiveRow, CellVal(2)
            .GetText 4, .ActiveRow, CellVal(3)
        End With
        CxnBookMaster.Execute "Insert Into BookChild01 Values ('" & rstBookMaster.Fields("Code").Value & "','" & IIf(CellVal(1) = "Outsource Item", "1", "3") & "','" & Right(CellVal(3), 6) & "'," & Val(CellVal(2)) & ")"
    Else
        With fpSpread3
            .GetText 1, .ActiveRow, CellVal(1)
            .GetText 2, .ActiveRow, CellVal(2)
            .GetText 3, .ActiveRow, CellVal(3)
        End With
        CxnBookMaster.Execute "INSERT INTO BookChild02 VALUES ('" & rstBookMaster.Fields("Code").Value & "'," & fpSpread3.ActiveRow & ",'" & CellVal(2) & "',#" & Format(GetDate(CellVal(1)), "mm-dd-yyyy") & "#,'" & CellVal(3) & "','P')"
    End If
    Exit Function
ErrorHandler:
    UpdateMaterialList = False
End Function
Private Sub fpSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    EditMode = IIf(Mode = 1, True, False)
End Sub
Private Sub fpSpread2_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    EditMode = IIf(Mode = 1, True, False)
End Sub
Private Sub fpSpread3_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    EditMode = IIf(Mode = 1, True, False)
End Sub
Private Sub RefreshDropDownList(ByVal xType As String)
    If xType = "R" Then
        rstOutsourceItemList.ActiveConnection = CxnBookMaster
        Do While Not RefreshRecord(rstOutsourceItemList)
        Loop
        rstOutsourceItemList.ActiveConnection = Nothing
        rstFreshBookList.ActiveConnection = CxnBookMaster
        Do While Not RefreshRecord(rstFreshBookList)
        Loop
        rstFreshBookList.ActiveConnection = Nothing
        OutsourceItem = "": FreshBook = ""
    End If
    Do While Not rstOutsourceItemList.EOF
        If OutsourceItem = "" Then
            OutsourceItem = rstOutsourceItemList.Fields("Name").Value
        Else
            OutsourceItem = OutsourceItem + Chr$(9) + rstOutsourceItemList.Fields("Name").Value
        End If
        rstOutsourceItemList.MoveNext
    Loop
    Do While Not rstFreshBookList.EOF
        If FreshBook = "" Then
            FreshBook = rstFreshBookList.Fields("Name").Value
        Else
            FreshBook = FreshBook + Chr$(9) + rstFreshBookList.Fields("Name").Value
        End If
        rstFreshBookList.MoveNext
    Loop
End Sub
Private Sub DuplicateRecord()
    On Error GoTo ErrorHandler
    MdiMainMenu.MousePointer = vbHourglass
    Dim BookCode As String
    BookCode = GenerateCode(CxnBookMaster, "SELECT MAX(Code) FROM BookMaster", 6, "0")
    CxnBookMaster.BeginTrans
    CxnBookMaster.Execute "INSERT INTO BookMaster SELECT '" & BookCode & "' As Code,TRIM(LEFT(P.Name,36))+' (D)' As Name,TRIM(LEFT(P.PrintName,36))+' (D)' As PrintName,[ISBN],[BusyCode],[Price],[Board],[Subject],[Class],[Group],[Size],[AddOnRate01],[AddOnRate02],[FormType],[Pages],[Forms],[BindingForms01],[BindingForms02],[OneColorPlateType],[OneColorPages],[OneColor¼Forms],[OneColor½Forms],[OneColor1F/BForms],[OneColor1W/TForms],[OneColorForms],[TwoColorPlateType],[TwoColorPages],[TwoColor¼Forms],[TwoColor½Forms],[TwoColor1F/BForms],[TwoColor1W/TForms],[TwoColorForms],[FourColorPlateType],[FourColorPages],[FourColor¼Forms],[FourColor½Forms],[FourColor1F/BForms],[FourColor1W/TForms],[FourColorForms],[TitleFrontColor],[TitleBackColor],[TitlePlateType],[LaminationType],[BindingType],[Qty/Pkt],[Pkt/Box],[DuplexPrinting],[BookPrinter],[TitlePrinter],[Laminator],[BinderFresh],[BinderRepair],[Type],[SaleLY1003],[SaleTY0409],[StockTransferLY1003],[StockTransferTY0409],[SpecimenLY1003],[SpecimenTY0409]," & _
                                            "[PendingSO],[SaleableStock],[POLTLY1003],[POLY0409],[POLY1003],[POTY0409],[PendingPO],[ESO30],[ESO60],[ESO90],[ESO150],[Royalty],[Remarks],[Narration],'" & UserCode & "' As [CreatedBy],#" & Format(Now(), "dd-MMM-yyyy HH:MM") & "# As [CreatedOn],'' As [ModifiedBy],NULL As [ModifiedOn],[Recordstatus],[Printstatus] FROM BookMaster P WHERE P.Code='" & rstBookList.Fields("Code").Value & "'"
    CxnBookMaster.Execute "INSERT INTO BookChild01 SELECT '" & BookCode & "' As Code,Category,Item,Quantity FROM BookChild01 C WHERE C.Code='" & rstBookList.Fields("Code").Value & "'"
    CxnBookMaster.CommitTrans
    If rstBookMaster.State = adStateOpen Then rstBookMaster.Close
    rstBookMaster.Open "SELECT * FROM BookMaster WHERE Code='" & BookCode & "'", CxnBookMaster, adOpenKeyset, adLockReadOnly
    AddToList
    Text1.Text = rstBookMaster.Fields("Name").Value: Sendkeys "{END}"
    MdiMainMenu.MousePointer = vbNormal
    Call MsgBox("Successfully Duplicated the Record !", vbInformation, App.Title)
    Exit Sub
ErrorHandler:
    MdiMainMenu.MousePointer = vbNormal
    DisplayError ("Failed to Duplicate the Record")
    CxnBookMaster.RollbackTrans
End Sub
