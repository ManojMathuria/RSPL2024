VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmPaperMaster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Paper Master"
   ClientHeight    =   5160
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
   Icon            =   "PaperMaster.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5160
   ScaleWidth      =   6750
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   5160
      Left            =   15
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   0
      Width           =   6735
      _Version        =   65536
      _ExtentX        =   11880
      _ExtentY        =   9102
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
      Picture         =   "PaperMaster.frx":0442
      Begin TabDlg.SSTab SSTab1 
         Height          =   4930
         Left            =   120
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   120
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   8705
         _Version        =   393216
         Style           =   1
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
         TabPicture(0)   =   "PaperMaster.frx":045E
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "DataGrid1"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Text1"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "&Details"
         TabPicture(1)   =   "PaperMaster.frx":047A
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Mh3dFrame2"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "&Op.Bal."
         TabPicture(2)   =   "PaperMaster.frx":0496
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Mh3dFrame3"
         Tab(2).ControlCount=   1
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
            TabIndex        =   23
            Top             =   4450
            Width           =   5775
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   3930
            Left            =   120
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   450
            Width           =   6255
            _ExtentX        =   11033
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
            ColumnCount     =   1
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
            SplitCount      =   1
            BeginProperty Split0 
               MarqueeStyle    =   3
               ScrollBars      =   3
               AllowRowSizing  =   0   'False
               AllowSizing     =   0   'False
               Locked          =   -1  'True
               BeginProperty Column00 
                  Locked          =   -1  'True
                  ColumnWidth     =   5940.284
               EndProperty
            EndProperty
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   2475
            Left            =   -74880
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   480
            Width           =   6255
            _Version        =   65536
            _ExtentX        =   11033
            _ExtentY        =   4366
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
            Picture         =   "PaperMaster.frx":04B2
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
               Left            =   4800
               MaxLength       =   40
               TabIndex        =   13
               Top             =   2000
               Width           =   1335
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
               Left            =   4800
               MaxLength       =   40
               TabIndex        =   6
               Top             =   740
               Width           =   1335
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
               Left            =   4800
               MaxLength       =   40
               TabIndex        =   4
               Top             =   425
               Width           =   1335
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
               Left            =   4800
               MaxLength       =   40
               TabIndex        =   2
               Top             =   105
               Width           =   1335
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
               Left            =   1320
               MaxLength       =   40
               TabIndex        =   8
               Top             =   1370
               Width           =   4815
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
               Left            =   1320
               MaxLength       =   40
               TabIndex        =   7
               Top             =   1055
               Width           =   4815
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   330
               Left            =   120
               TabIndex        =   18
               Top             =   1370
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
               Caption         =   " Print Name"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "PaperMaster.frx":04CE
               Picture         =   "PaperMaster.frx":04EA
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Left            =   120
               TabIndex        =   16
               Top             =   1055
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
               Caption         =   " Name"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "PaperMaster.frx":0506
               Picture         =   "PaperMaster.frx":0522
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel7 
               Height          =   330
               Left            =   120
               TabIndex        =   26
               Top             =   740
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
               Caption         =   " Weight/Ream"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "PaperMaster.frx":053E
               Picture         =   "PaperMaster.frx":055A
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
               Height          =   330
               Left            =   120
               TabIndex        =   27
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
               Caption         =   " Type"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "PaperMaster.frx":0576
               Picture         =   "PaperMaster.frx":0592
            End
            Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame4 
               Height          =   330
               Index           =   0
               Left            =   1320
               TabIndex        =   28
               TabStop         =   0   'False
               Top             =   105
               Width           =   2295
               _Version        =   65536
               _ExtentX        =   4048
               _ExtentY        =   582
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
               Picture         =   "PaperMaster.frx":05AE
               Begin VB.OptionButton Option3 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Book"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   210
                  Left            =   120
                  TabIndex        =   0
                  Top             =   60
                  Width           =   735
               End
               Begin VB.OptionButton Option4 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Title"
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
                  Left            =   860
                  TabIndex        =   1
                  Top             =   60
                  Width           =   735
               End
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
               Height          =   330
               Left            =   3600
               TabIndex        =   29
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
               Caption         =   " Size"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "PaperMaster.frx":05CA
               Picture         =   "PaperMaster.frx":05E6
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel9 
               Height          =   330
               Left            =   120
               TabIndex        =   30
               Top             =   420
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
               Picture         =   "PaperMaster.frx":0602
               Picture         =   "PaperMaster.frx":061E
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel10 
               Height          =   330
               Left            =   3600
               TabIndex        =   31
               Top             =   420
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
               Caption         =   " Make"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "PaperMaster.frx":063A
               Picture         =   "PaperMaster.frx":0656
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
               Height          =   330
               Left            =   3600
               TabIndex        =   34
               Top             =   735
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
               Caption         =   " Sub-Make"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "PaperMaster.frx":0672
               Picture         =   "PaperMaster.frx":068E
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
               Height          =   330
               Index           =   0
               Left            =   120
               TabIndex        =   35
               Top             =   1680
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
               Caption         =   " Ream/Bdl (T)"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "PaperMaster.frx":06AA
               Picture         =   "PaperMaster.frx":06C6
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput3 
               Height          =   330
               Left            =   1320
               TabIndex        =   3
               Top             =   420
               Width           =   2295
               _Version        =   65536
               _ExtentX        =   4048
               _ExtentY        =   582
               Calculator      =   "PaperMaster.frx":06E2
               Caption         =   "PaperMaster.frx":0702
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PaperMaster.frx":076E
               Keys            =   "PaperMaster.frx":078C
               Spin            =   "PaperMaster.frx":07D6
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
               ValueVT         =   1245189
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput4 
               Height          =   330
               Left            =   1320
               TabIndex        =   5
               Top             =   735
               Width           =   2295
               _Version        =   65536
               _ExtentX        =   4048
               _ExtentY        =   582
               Calculator      =   "PaperMaster.frx":07FE
               Caption         =   "PaperMaster.frx":081E
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PaperMaster.frx":088A
               Keys            =   "PaperMaster.frx":08A8
               Spin            =   "PaperMaster.frx":08F2
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16777215
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "##0.000"
               EditMode        =   1
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "##0.000"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   999.999
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
            Begin TDBNumber6Ctl.TDBNumber MhRealInput5 
               Height          =   330
               Left            =   1320
               TabIndex        =   9
               Top             =   1680
               Width           =   2295
               _Version        =   65536
               _ExtentX        =   4048
               _ExtentY        =   582
               Calculator      =   "PaperMaster.frx":091A
               Caption         =   "PaperMaster.frx":093A
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PaperMaster.frx":09A6
               Keys            =   "PaperMaster.frx":09C4
               Spin            =   "PaperMaster.frx":0A0E
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
               ValueVT         =   1973092357
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
               Height          =   330
               Index           =   0
               Left            =   3600
               TabIndex        =   36
               Top             =   1680
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
               Caption         =   " Ream/Bdl (C)"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "PaperMaster.frx":0A36
               Picture         =   "PaperMaster.frx":0A52
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput6 
               Height          =   330
               Left            =   4800
               TabIndex        =   10
               Top             =   1680
               Width           =   1335
               _Version        =   65536
               _ExtentX        =   2355
               _ExtentY        =   582
               Calculator      =   "PaperMaster.frx":0A6E
               Caption         =   "PaperMaster.frx":0A8E
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PaperMaster.frx":0AFA
               Keys            =   "PaperMaster.frx":0B18
               Spin            =   "PaperMaster.frx":0B62
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
               ValueVT         =   1973092357
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
               Height          =   330
               Index           =   1
               Left            =   120
               TabIndex        =   37
               Top             =   2000
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
               Caption         =   " Paper Type"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "PaperMaster.frx":0B8A
               Picture         =   "PaperMaster.frx":0BA6
            End
            Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame4 
               Height          =   330
               Index           =   1
               Left            =   1320
               TabIndex        =   38
               TabStop         =   0   'False
               Top             =   1995
               Width           =   2295
               _Version        =   65536
               _ExtentX        =   4048
               _ExtentY        =   582
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
               Picture         =   "PaperMaster.frx":0BC2
               Begin VB.OptionButton Option6 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "UnCoated"
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
                  Left            =   1095
                  TabIndex        =   12
                  Top             =   60
                  Width           =   1095
               End
               Begin VB.OptionButton Option5 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Coated"
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
                  TabIndex        =   11
                  Top             =   60
                  Width           =   975
               End
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
               Height          =   330
               Index           =   1
               Left            =   3600
               TabIndex        =   39
               Top             =   2000
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
               Caption         =   " HSN Code"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "PaperMaster.frx":0BDE
               Picture         =   "PaperMaster.frx":0BFA
            End
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame3 
            Height          =   2875
            Left            =   -74880
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   480
            Width           =   6255
            _Version        =   65536
            _ExtentX        =   11033
            _ExtentY        =   5071
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
            Picture         =   "PaperMaster.frx":0C16
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
               Left            =   4520
               MaxLength       =   13
               TabIndex        =   33
               Text            =   "0"
               Top             =   590
               Visible         =   0   'False
               Width           =   1395
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
               Left            =   3135
               MaxLength       =   13
               TabIndex        =   17
               Text            =   "0.000"
               Top             =   590
               Visible         =   0   'False
               Width           =   1395
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
               Left            =   440
               MaxLength       =   40
               TabIndex        =   15
               Top             =   590
               Visible         =   0   'False
               Width           =   2715
            End
            Begin MSDataGridLib.DataGrid DataGrid2 
               Height          =   2650
               Left            =   120
               TabIndex        =   14
               Top             =   100
               Width           =   6015
               _ExtentX        =   10610
               _ExtentY        =   4657
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
               Caption         =   "Opening Balance"
               ColumnCount     =   3
               BeginProperty Column00 
                  DataField       =   "GodownName"
                  Caption         =   "Godown Name"
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
                  DataField       =   "OpBalOther"
                  Caption         =   "          Op.Bal."
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   1
                     Format          =   "0.000"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   1
                  EndProperty
               EndProperty
               BeginProperty Column02 
                  DataField       =   "OpBalTat"
                  Caption         =   "              Tat"
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
                     ColumnWidth     =   2700.284
                  EndProperty
                  BeginProperty Column01 
                     Alignment       =   1
                     Locked          =   -1  'True
                     ColumnWidth     =   1379.906
                  EndProperty
                  BeginProperty Column02 
                     Alignment       =   1
                     ColumnWidth     =   1379.906
                  EndProperty
               EndProperty
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
            ForeColor       =   &H80000004&
            Height          =   330
            Left            =   120
            TabIndex        =   25
            Top             =   4450
            Width           =   495
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   20
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
Attribute VB_Name = "FrmPaperMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CxnPaperMaster As New ADODB.Connection
Dim rstPaperList As New ADODB.Recordset
Dim rstPaperMaster As New ADODB.Recordset
Dim rstPaperChild As New ADODB.Recordset
Dim rstAccountList As New ADODB.Recordset
Dim rstSizeList As New ADODB.Recordset
Dim rstCheckRef As New ADODB.Recordset
Dim AccountCode As String
Dim SizeCode As String
Dim PrevStr As String
Dim dblBookMark As Double
Dim blnRecordExist As Boolean
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    
    CenterForm Me
    BusySystemIndicator True
    CxnPaperMaster.CursorLocation = adUseClient
    CxnPaperMaster.Open CxnDatabase.ConnectionString
    rstPaperList.Open "Select Name,Code From PaperMaster Order By Name", CxnPaperMaster, adOpenKeyset, adLockOptimistic
    rstSizeList.Open "Select Name As Col0, Code From GeneralMaster Where Type = '1' Order By Name", CxnPaperMaster, adOpenKeyset, adLockReadOnly
    rstAccountList.Open "Select Name As Col0, Code From AccountMaster Where Type In ('05','06','09') Order By Name", CxnPaperMaster, adOpenKeyset, adLockReadOnly
    rstPaperMaster.CursorLocation = adUseClient
    rstPaperList.Filter = adFilterNone
    Set DataGrid1.DataSource = rstPaperList
    BusySystemIndicator False
    SSTab1.Tab = 0
    If Not (rstPaperList.EOF Or rstPaperList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    rstPaperList.ActiveConnection = Nothing
    rstSizeList.ActiveConnection = Nothing
    rstAccountList.ActiveConnection = Nothing
    SetButtonsForNoRecord
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    Call CloseForm(FrmPaperMaster)
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
          Call CloseForm(FrmPaperMaster)
       Else
           If Toolbar1.Buttons.Item(1).Enabled Then
              SSTab1.Tab = 0
           Else
              If Me.ActiveControl.Name <> "MhRealInput1" And Me.ActiveControl.Name <> "MhRealInput2" And Me.ActiveControl.Name <> "Text12" Then
                   If MsgBox("Are you sure to Quit?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Quit !") <> vbYes Then
                        Me.ActiveControl.SetFocus
                   Else
                        Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
                   End If
              End If
           End If
           If Me.ActiveControl.Name <> "MhRealInput1" And Me.ActiveControl.Name <> "MhRealInput2" And Me.ActiveControl.Name <> "Text12" Then
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
       If Me.ActiveControl.Name <> "MhRealInput1" And Me.ActiveControl.Name <> "MhRealInput2" And Me.ActiveControl.Name <> "Text12" Then
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
        Call CloseForm(FrmPaperMaster)
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CloseRecordset(rstPaperList)
    Call CloseRecordset(rstPaperMaster)
    Call CloseRecordset(rstPaperChild)
    Call CloseRecordset(rstSizeList)
    Call CloseRecordset(rstAccountList)
    Call CloseConnection(CxnPaperMaster)
    Call CloseRecordset(rstCheckRef)
    ShowProgressInStatusBar False
    DisableChildMenu
End Sub
Private Sub Text1_Change()
    If rstPaperList.RecordCount = 0 Then Exit Sub
    rstPaperList.MoveFirst
    If Text1.Text <> "" Then
        rstPaperList.Find "[Name] Like '" & FixQuote(Text1.Text) & "%'"
        If rstPaperList.EOF Then
            rstPaperList.MoveFirst
            If PrevStr <> "" And Len(Text1.Text) > 1 Then
                If dblBookMark <> 0 Then
                    rstPaperList.Bookmark = dblBookMark
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
    If Not (rstPaperList.EOF Or rstPaperList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim KeyProcessed As Boolean
    
    If rstPaperList.RecordCount = 0 Then Exit Sub
    If Shift = 0 And KeyCode = vbKeyUp Then
        With rstPaperList
            .MovePrevious
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyBack Then
        With rstPaperList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyDown Then
        With rstPaperList
            .MoveNext
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageUp Then
        With rstPaperList
            .Move (-1) * (DataGrid1.VisibleRows - 1)
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageUp Then
        With rstPaperList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageDown Then
        With rstPaperList
            .Move DataGrid1.VisibleRows - 1
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageDown Then
        With rstPaperList
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
            If Not (rstPaperList.EOF Or rstPaperList.BOF) Then
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
            If Option3.Value Then
                Option3.SetFocus
            Else
                Option4.SetFocus
            End If
        Else
            Mh3dFrame2.Enabled = False
            Mh3dFrame3.Enabled = True
            DataGrid2.SetFocus
        End If
    End If
End Sub
Public Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim HiLiteRecord As Boolean
    Dim UpdateFlag As Integer
    
    If Button.Index = 1 Then
        If rstPaperMaster.State = adStateOpen Then
           rstPaperMaster.Close
        End If
        rstPaperMaster.Open "Select * From PaperMaster Where Code = ''", CxnPaperMaster, adOpenKeyset, adLockOptimistic
        ClearFields ("P")
        ClearFields ("C")
        Call LoadOpBalList("")
        If rstPaperChild.State = adStateClosed Then
            SSTab1.Tab = 0
            Exit Sub
        End If
        If AddRecord(rstPaperMaster) Then
            Call SetButtons(False)
            SSTab1.Tab = 1
            If Option3.Value Then
                Option3.SetFocus
            Else
                Option4.SetFocus
            End If
            blnRecordExist = False
            CxnPaperMaster.BeginTrans
        End If
    ElseIf Button.Index = 2 Then
        If rstPaperList.RecordCount = 0 Then Exit Sub
        SSTab1.Tab = 1
        EditRecord
    ElseIf Button.Index = 3 Then
        If rstPaperList.RecordCount = 0 Then Exit Sub
        If AllowMastersDeletion = 0 Then
            Call DisplayError("You don't have the rights to Delete this Master")
            Exit Sub
        End If
        SSTab1.Tab = 1
        If CheckRef Then
            DisplayError ("Failed to delete the record")
        ElseIf MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            On Error Resume Next
            MdiMainMenu.MousePointer = vbHourglass
            CxnPaperMaster.Execute "DELETE FROM PaperMaster WHERE Code = '" & rstPaperList.Fields("Code").Value & "'"
            MdiMainMenu.MousePointer = vbNormal
            If Err.Number = 0 Then
                rstPaperList.Delete
                rstPaperList.MoveNext
                If rstPaperList.RecordCount > 0 And rstPaperList.EOF Then rstPaperList.MoveLast
                Call UpdateUserAction("Paper Master", "D", Trim(Text2.Text), CxnPaperMaster)
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
        If blnRecordExist And AllowMastersModification = 0 Then
            Call DisplayError("You don't have the rights to Edit this Master")
            Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
            Exit Sub
        End If
        SaveFields
        UpdateFlag = 0
        If UpdateRecord(rstPaperMaster) Then
            If UpdateOpBalList("D") Then
                 UpdateFlag = 1
                 If rstPaperChild.RecordCount <> 0 Then
                      rstPaperChild.MoveFirst
                      Do While Not rstPaperChild.EOF
                          If (Val(rstPaperChild.Fields("OpBalOther").Value) <> 0 Or Val(rstPaperChild.Fields("OpBalTat").Value) <> 0) And rstPaperChild.Fields("Imported").Value = "N" Then
                               If Not UpdateOpBalList("U") Then
                                    UpdateFlag = 0
                                    Exit Do
                               End If
                          End If
                          rstPaperChild.MoveNext
                      Loop
                 End If
            End If
        End If
        If UpdateFlag Then
            Call UpdateUserAction("Paper Master", IIf(blnRecordExist, "M", "A"), Trim(Text2.Text), CxnPaperMaster)
            AddToList
            CxnPaperMaster.CommitTrans
            If rstPaperMaster.State = adStateOpen Then
                rstPaperMaster.Close
            End If
            rstPaperMaster.CursorLocation = adUseClient
            Call SetButtons(True)
            SSTab1.Tab = 0
            ShowProgressInStatusBar True
            Timer1.Enabled = True
        Else
            DisplayError ("Failed to save the record")
            Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
        End If
    ElseIf Button.Index = 5 Then
        If CancelRecordUpdate(rstPaperMaster) Then
            CxnPaperMaster.RollbackTrans
            If rstPaperMaster.State = adStateOpen Then
                rstPaperMaster.Close
            End If
            rstPaperMaster.CursorLocation = adUseClient
            Call SetButtons(True)
            SetButtonsForNoRecord
            SSTab1.Tab = 0
        End If
    ElseIf Button.Index = 6 Then
        SSTab1.Tab = 0
        Set DataGrid1.DataSource = Nothing
        rstPaperList.ActiveConnection = CxnPaperMaster
        Do While Not RefreshRecord(rstPaperList)
        Loop
        Set DataGrid1.DataSource = rstPaperList
        rstPaperList.ActiveConnection = Nothing
        rstSizeList.ActiveConnection = CxnPaperMaster
        Do While Not RefreshRecord(rstSizeList)
        Loop
        rstSizeList.ActiveConnection = Nothing
        rstAccountList.ActiveConnection = CxnPaperMaster
        Do While Not RefreshRecord(rstAccountList)
        Loop
        rstAccountList.ActiveConnection = Nothing
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
        If rstPaperList.RecordCount > 0 Then rstPaperList.MoveFirst
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 14 Then
        If rstPaperList.RecordCount > 0 Then
           rstPaperList.MovePrevious
           If rstPaperList.BOF Then
              rstPaperList.MoveNext
           End If
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 15 Then
        If rstPaperList.RecordCount > 0 Then
           rstPaperList.MoveNext
           If rstPaperList.EOF Then
              rstPaperList.MovePrevious
           End If
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 16 Then
        If rstPaperList.RecordCount > 0 Then rstPaperList.MoveLast
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 18 Then
        Call CloseForm(FrmPaperMaster)
        HiLiteRecord = False
    End If
    If HiLiteRecord Then
        If Not (rstPaperList.EOF Or rstPaperList.BOF) Then
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
    Toolbar1.Buttons.Item(13).Enabled = bVal
    Toolbar1.Buttons.Item(14).Enabled = bVal
    Toolbar1.Buttons.Item(15).Enabled = bVal
    Toolbar1.Buttons.Item(16).Enabled = bVal
    Toolbar1.Buttons.Item(18).Enabled = bVal
    Mh3dFrame2.Enabled = Not bVal
End Sub
Private Sub SetButtonsForNoRecord()
    If rstPaperList.RecordCount = 0 Then
        Toolbar1.Buttons.Item(2).Enabled = False
        Toolbar1.Buttons.Item(3).Enabled = False
        Toolbar1.Buttons.Item(13).Enabled = False
        Toolbar1.Buttons.Item(14).Enabled = False
        Toolbar1.Buttons.Item(15).Enabled = False
        Toolbar1.Buttons.Item(16).Enabled = False
    End If
End Sub
Private Sub Text2_Validate(Cancel As Boolean)
    If rstPaperMaster.EOF Or rstPaperMaster.BOF Then Exit Sub
    If CheckEmpty(Text2, True) Then
        Cancel = True
    ElseIf CheckDuplicate(CxnPaperMaster, "PaperMaster", "Code", "Name", Text2.Text, rstPaperMaster.Fields("Code").Value, False) Then
        Cancel = True
    ElseIf CheckEmpty(Text3, False) Then
        Text3.Text = Text2.Text
    End If
End Sub
Private Sub Text4_Change()
    If Text4.Text = " " Then
        Text4.Text = "?"
        SendKeys "{TAB}"
    End If
End Sub
Private Sub Text4_Validate(Cancel As Boolean)
    Dim SearchString As String
    
    SearchString = FixQuote(Text4.Text)
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
       Call DisplaySelectionList(Text4, SizeCode)
       Call CloseForm(FrmSelectionList)
       If CheckEmpty(Text4.Text, False) Then
          Text4.Text = "?"
       End If
       If RTrim(SizeCode) <> "" Then
          SendKeys "{TAB}"
       End If
       Cancel = True
    Else
       SizeCode = rstSizeList.Fields("Code").Value
    End If
End Sub
Private Sub Text5_Validate(Cancel As Boolean)
    If CheckEmpty(Text5.Text, True) Then
        Cancel = True
    End If
End Sub
Private Sub Text6_Validate(Cancel As Boolean)
    If CheckEmpty(Text6.Text, True) Then
        Cancel = True
    Else
        If CheckEmpty(Text2, False) Then
            Text2.Text = Trim(Text5.Text) + "-" + Trim(MhRealInput3.Text) + "-" + Trim(Text4.Text) + "-" + Trim(MhRealInput4.Text) + "-" + Trim(Text6.Text)
        End If
        If CheckEmpty(Text3, False) Then
            Text3.Text = Text2.Text
        End If
    End If
End Sub
Private Sub MhRealInput4_Validate(Cancel As Boolean)
    If Option4.Value Or blnRecordExist Then Exit Sub
    If Val(MhRealInput4.Text) <= 12.4 Then
        MhRealInput5.Text = "5.00"
    ElseIf Val(MhRealInput4.Text) >= 13.6 And Val(MhRealInput4.Text) <= 15.5 Then
        MhRealInput5.Text = "4.00"
    ElseIf Val(MhRealInput4.Text) >= 17 And Val(MhRealInput4.Text) <= 21.3 Then
        MhRealInput5.Text = "3.00"
    ElseIf Val(MhRealInput4.Text) >= 24 Then
        MhRealInput5.Text = "2.00"
    End If
End Sub
Private Sub ViewRecord()
    ClearFields ("P")
    ClearFields ("C")
    If rstPaperList.EOF Then Exit Sub
    FindRecord
    LoadFields
End Sub
Private Sub FindRecord()
    If rstPaperMaster.State = adStateOpen Then
       rstPaperMaster.Close
    End If
    rstPaperMaster.Open "Select * From PaperMaster Where Code = '" & FixQuote(rstPaperList.Fields("Code").Value) & "'", CxnPaperMaster, adOpenKeyset, adLockOptimistic
    If rstPaperMaster.RecordCount = 0 Then
       Call DisplayError("This Record has been deleted by Another User ! Click Ok To Refresh the Recordset")
       Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
    End If
End Sub
Private Sub ClearFields(ByVal strType As String)
    If strType = "P" Then
        Text2.Text = ""
        Text3.Text = ""
        Text4.Text = ""
        Text41.Text = ""
        Text5.Text = ""
        Text6.Text = ""
        MhRealInput3.Text = "0"
        MhRealInput4.Text = "0.000"
        MhRealInput5.Text = "0.00"
        MhRealInput6.Text = "0.00"
        Option3.Value = True
        Option4.Value = False
        
        Option4.Value = True
        Option5.Value = False
        
    ElseIf strType = "C" Then
        Text12.Text = ""
        MhRealInput1.Text = "0"
        MhRealInput2.Text = "0.000"
    End If
End Sub
Private Sub LoadFields()
    If rstPaperMaster.EOF Or rstPaperMaster.BOF Then Exit Sub
    Text2.Text = rstPaperMaster.Fields("Name").Value
    Text3.Text = rstPaperMaster.Fields("PrintName").Value
    If rstPaperMaster.Fields("Type").Value = "1" Then
        Option3.Value = True
    Else
        Option4.Value = True
    End If
    SizeCode = rstPaperMaster.Fields("Size").Value
    rstSizeList.MoveFirst
    rstSizeList.Find "[Code] = '" & SizeCode & "'"
    Text4.Text = rstSizeList.Fields("Col0").Value
    MhRealInput3.Text = Format(Val(rstPaperMaster.Fields("GSM").Value), "0")
    Text5.Text = rstPaperMaster.Fields("Make").Value
    Text6.Text = rstPaperMaster.Fields("SubMake").Value
    MhRealInput4.Text = Format(Val(rstPaperMaster.Fields("Weight/Ream").Value), "0.000")
    MhRealInput5.Text = Format(Val(rstPaperMaster.Fields("Reams/Bundle").Value), "0.00")
    MhRealInput6.Text = Format(Val(rstPaperMaster.Fields("Reams/Bundle2").Value), "0.00")
    Call LoadOpBalList(rstPaperMaster.Fields("Code").Value)
    If rstPaperMaster.Fields("PaperType").Value = "C" Then
        Option4.Value = True
    Else
        Option5.Value = True
    End If
    Text41.Text = rstPaperMaster.Fields("HSNCode").Value
End Sub
Private Sub EditRecord()
    On Error GoTo ErrorHandler
    
    If rstPaperMaster.RecordCount = 0 Then Exit Sub
    If rstPaperChild.State = adStateClosed Then
        SSTab1.Tab = 0
        Exit Sub
    End If
    If rstPaperMaster.State = adStateOpen Then
       rstPaperMaster.Close
    End If
    rstPaperMaster.CursorLocation = adUseServer
    rstPaperMaster.Open "Select * From PaperMaster Where Code = '" & FixQuote(rstPaperList.Fields("Code").Value) & "'", CxnPaperMaster, adOpenKeyset, adLockPessimistic
    MdiMainMenu.MousePointer = vbHourglass
    rstPaperMaster.Fields("Printstatus") = "N"
    MdiMainMenu.MousePointer = vbNormal
    AddToList
    Call SetButtons(False)
    SSTab1.TabEnabled(0) = False
    
    If Option3.Value Then
        Option3.SetFocus
    Else
        Option4.SetFocus
    End If
    
    If Option4.Value Then
        Option4.SetFocus
    Else
        Option5.SetFocus
    End If
    
    blnRecordExist = True
    CxnPaperMaster.BeginTrans
    Exit Sub
ErrorHandler:
    If Err.Number = -2147467259 Then
       Call DisplayError("Failed to Edit the record")
    End If
    MdiMainMenu.MousePointer = vbNormal
    SSTab1.Tab = 0
End Sub
Private Sub SaveFields()
    If rstPaperMaster.EOF Or rstPaperMaster.BOF Then Exit Sub
    If Not blnRecordExist Then
        rstPaperMaster.Fields("Code").Value = GenerateCode(CxnPaperMaster, "Select Max(Code) From PaperMaster", 6, "0")
        rstPaperMaster.Fields("CreatedBy").Value = UserCode
        rstPaperMaster.Fields("CreatedOn").Value = Now()
        rstPaperMaster.Fields("Recordstatus").Value = "N"
    Else
        rstPaperMaster.Fields("ModifiedBy").Value = UserCode
        rstPaperMaster.Fields("ModifiedOn").Value = Now()
        rstPaperMaster.Fields("Recordstatus").Value = "M"
    End If
    
    rstPaperMaster.Fields("Name").Value = Trim(Text2.Text)
    rstPaperMaster.Fields("PrintName").Value = Trim(Text3.Text)
    If Option3.Value Then
        rstPaperMaster.Fields("Type").Value = "1"
    Else
        rstPaperMaster.Fields("Type").Value = "2"
    End If
    rstPaperMaster.Fields("Size").Value = SizeCode
    rstPaperMaster.Fields("GSM").Value = Format(Val(MhRealInput3.Text), "0")
    rstPaperMaster.Fields("Make").Value = Trim(Text5.Text)
    rstPaperMaster.Fields("SubMake").Value = Trim(Text6.Text)
    rstPaperMaster.Fields("Weight/Ream").Value = Format(Val(MhRealInput4.Text), "0.000")
    rstPaperMaster.Fields("Reams/Bundle").Value = Format(Val(MhRealInput5.Text), "0.00")
    rstPaperMaster.Fields("Reams/Bundle2").Value = Format(Val(MhRealInput6.Text), "0.00")
    rstPaperMaster.Fields("PrintStatus").Value = "N"
    
    If Option4.Value Then
        rstPaperMaster.Fields("PaperType").Value = "C"
    Else
        rstPaperMaster.Fields("PaperType").Value = "U"
    End If
    rstPaperMaster.Fields("HSNCode").Value = Trim(Text41.Text)
    
End Sub
Private Sub AddToList()
    On Error Resume Next
    
    rstPaperList.MoveFirst
    rstPaperList.Find "[Code] = '" & rstPaperMaster.Fields("Code").Value & "'"
    If rstPaperList.EOF Then
       rstPaperList.AddNew
       rstPaperList.Fields("Code").Value = rstPaperMaster.Fields("Code").Value
    End If
    rstPaperList.Fields("Name").Value = rstPaperMaster.Fields("Name").Value
    rstPaperList.Update
    rstPaperList.Sort = "Name Asc"
    rstPaperList.Find "[Code] = '" & rstPaperMaster.Fields("Code").Value & "'"
End Sub
Private Function CheckMandatoryFields() As Boolean
    If CheckEmpty(Text2.Text, False) Then
        SSTab1.Tab = 1
        Text2.SetFocus
        CheckMandatoryFields = True
    ElseIf CheckDuplicate(CxnPaperMaster, "PaperMaster", "Code", "Name", Text2.Text, rstPaperMaster.Fields("Code").Value, False) Then
        SSTab1.Tab = 1
        Text2.SetFocus
        CheckMandatoryFields = True
    ElseIf CheckEmpty(Text3.Text, False) Then
        SSTab1.Tab = 1
        Text3.SetFocus
        CheckMandatoryFields = True
    ElseIf CheckEmpty(Text4.Text, False) Then
        SSTab1.Tab = 1
        Text4.SetFocus
        CheckMandatoryFields = True
    ElseIf Not CheckExists(Text4, "Col0", rstSizeList, SizeCode) Then
        SSTab1.Tab = 1
        Text4.SetFocus
        CheckMandatoryFields = True
    ElseIf CheckEmpty(Text5.Text, False) Then
        SSTab1.Tab = 1
        Text5.SetFocus
        CheckMandatoryFields = True
    ElseIf CheckEmpty(Text6.Text, False) Then
        SSTab1.Tab = 1
        Text6.SetFocus
        CheckMandatoryFields = True
    End If
End Function
Private Sub LoadOpBalList(ByVal strPaperCode As String)
    On Error GoTo ErrorHandler
    
    If rstPaperChild.State = adStateOpen Then
       rstPaperChild.Close
    End If
    
    
    rstPaperChild.Open "Select P.Account, A.Name As GodownName, P.OpBalOther, P.OpBalTat, P.Imported From PaperChild P, AccountMaster A Where P.Account = A.Code And P.Code = '" & FixQuote(strPaperCode) & "' Order By A.Name", CxnPaperMaster, adOpenKeyset, adLockOptimistic
    
    
    rstPaperChild.ActiveConnection = Nothing
    Set DataGrid2.DataSource = rstPaperChild
    Exit Sub
ErrorHandler:
    DisplayError ("Failed to Load Opening Balance")
End Sub
Private Function UpdateOpBalList(ByVal strOption As String) As Boolean
    Dim Sheets As Long
    On Error GoTo ErrorHandler
    
    UpdateOpBalList = True
    If strOption = "D" Then
        CxnPaperMaster.Execute "Delete From PaperChild WHERE Code = '" & rstPaperMaster.Fields("Code").Value & "' And Imported = 'N'"
    Else
        Sheets = (Fix(Val(rstPaperChild.Fields("OpBalOther").Value)) * 500) + ((Val(rstPaperChild.Fields("OpBalOther").Value) - Fix(Val(rstPaperChild.Fields("OpBalOther").Value))) * 1000)
        CxnPaperMaster.Execute "Insert Into PaperChild Values ('" & rstPaperMaster.Fields("Code").Value & "','" & rstPaperChild.Fields("Account").Value & "'," & rstPaperChild.Fields("OpBalOther").Value & "," & Sheets & "," & rstPaperChild.Fields("OpBalTat").Value & ",'N')"
    End If
    Exit Function
ErrorHandler:
    UpdateOpBalList = False
End Function
Private Function CheckRef() As Boolean
    On Error GoTo ErrorHandler
    
    If rstCheckRef.State = adStateOpen Then
         rstCheckRef.Close
    End If
    rstCheckRef.Open "Select Paper1 From BookPOChild05 Where Paper1 = '" & rstPaperList.Fields("Code").Value & "'", CxnPaperMaster, adOpenKeyset, adLockReadOnly
    If rstCheckRef.RecordCount > 0 Then
        CheckRef = True
        Exit Function
    End If
    If rstCheckRef.State = adStateOpen Then
         rstCheckRef.Close
    End If
    rstCheckRef.Open "Select Paper2 From BookPOChild05 Where Paper2 = '" & rstPaperList.Fields("Code").Value & "'", CxnPaperMaster, adOpenKeyset, adLockReadOnly
    If rstCheckRef.RecordCount > 0 Then
        CheckRef = True
        Exit Function
    End If
    If rstCheckRef.State = adStateOpen Then
         rstCheckRef.Close
    End If
    rstCheckRef.Open "Select Paper4 From BookPOChild05 Where Paper4 = '" & rstPaperList.Fields("Code").Value & "'", CxnPaperMaster, adOpenKeyset, adLockReadOnly
    If rstCheckRef.RecordCount > 0 Then
        CheckRef = True
        Exit Function
    End If
    If rstCheckRef.State = adStateOpen Then
         rstCheckRef.Close
    End If
'    rstCheckRef.Open "Select Paper6 From BookPOChild05 Where Paper6 = '" & rstPaperList.Fields("Code").Value & "'", CxnPaperMaster, adOpenKeyset, adLockReadOnly
'    If rstCheckRef.RecordCount > 0 Then
'        CheckRef = True
'    End If
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
Private Sub DataGrid2_DblClick()
    Call DataGrid2_KeyDown(vbKeyE, vbCtrlMask)
End Sub
Private Sub DataGrid2_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask And KeyCode = vbKeyE Then
        If rstPaperChild.RecordCount = 0 Or rstPaperChild.Fields("Imported").Value = "Y" Then Exit Sub
        If Val(CheckNull(rstPaperChild.Fields("OpBalOther").Value)) <> 0 Or Val(CheckNull(rstPaperChild.Fields("OpBalTat").Value)) <> 0 Then
            AccountCode = rstPaperChild.Fields("Account").Value
            Text12.Text = rstPaperChild.Fields("GodownName").Value
            MhRealInput2.Text = Format(Val(rstPaperChild.Fields("OpBalOther").Value), "0.000")
            MhRealInput1.Text = Format(Val(rstPaperChild.Fields("OpBalTat").Value), "0")
        End If
        With DataGrid2
            Text12.Visible = True
            Text12.Move .Left + .Columns(0).Left, .Top + .RowTop(.Row), .Columns(0).Width + 10, .RowHeight + 30
            MhRealInput2.Visible = True
            MhRealInput2.Move .Left + .Columns(1).Left, .Top + .RowTop(.Row), .Columns(1).Width + 10, .RowHeight + 30
            MhRealInput1.Visible = True
            MhRealInput1.Move .Left + .Columns(2).Left, .Top + .RowTop(.Row), .Columns(2).Width + 10, .RowHeight + 30
        End With
        DataGrid2.Enabled = False
        Text12.SetFocus
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyA Then
        SendKeys "^"
        Call AddRecord(rstPaperChild)
        Call ClearFields("C")
        Call DataGrid2_KeyDown(vbKeyE, vbCtrlMask)
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyD Then
        If rstPaperChild.RecordCount = 0 Or rstPaperChild.Fields("Imported").Value = "Y" Then Exit Sub
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            Set DataGrid2.DataSource = Nothing
            rstPaperChild.Delete
            rstPaperChild.MoveNext
            Set DataGrid2.DataSource = rstPaperChild
            DataGrid2.SetFocus
        End If
        If rstPaperChild.RecordCount = 0 Then
            Call ClearFields("C")
        End If
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyS And Toolbar1.Buttons.Item(4).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(4)
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
Private Sub Text12_Change()
    If Text12.Text = " " Then
        Text12.Text = "?"
        SendKeys "{TAB}"
    End If
End Sub
Private Sub Text12_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyEscape Then
        MakeTextBoxInvisible (True)
    End If
End Sub
Private Sub Text12_Validate(Cancel As Boolean)
    Dim SearchString As String
    
    SearchString = FixQuote(Text12.Text)
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
        Call DisplaySelectionList(Text12, AccountCode)
        Call CloseForm(FrmSelectionList)
        If CheckEmpty(Text12.Text, False) Then
            Text12.Text = "?"
        End If
        If RTrim(AccountCode) <> "" Then
            SendKeys "{TAB}"
        End If
        Cancel = True
        Exit Sub
    ElseIf (rstPaperChild.Fields("GodownName").Value <> Text12.Text) Or (CheckEmpty(rstPaperChild.Fields("GodownName").Value, False)) Then
        If CheckDuplicateGodown Then
            Call DisplayError("Duplicate Entry")
            Text12.SelStart = 0
            Text12.SelLength = Len(Text12.Text)
            Cancel = True
            Exit Sub
        End If
    End If
    AccountCode = rstAccountList.Fields("Code").Value
End Sub
Private Sub MhRealInput2_GotFocus()
    FocusSelect Me.ActiveControl
End Sub
Private Sub MhRealInput2_KeyPress(KeyAscii As Integer)
    ValidateKey MhRealInput2, KeyAscii, 3
End Sub
Private Sub MhRealInput2_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyEscape Then
        MakeTextBoxInvisible (True)
    End If
End Sub
Private Sub MhRealInput2_Validate(Cancel As Boolean)
    Dim RPB As Double
    
    If Not ValidateNumber(Me.ActiveControl, 3) Then
        Cancel = True
    Else
        If Val(CheckNull(rstPaperChild.Fields("OpBalTat").Value)) = 0 Then
            If Val(MhRealInput5.Text) <> 0 Then
                RPB = Val(MhRealInput5.Text)
                If Val(MhRealInput2.Text) * 1000 Mod RPB * 1000 > 0 Then
                    MhRealInput1.Text = Format(Int(Val(MhRealInput2.Text) / RPB) + 1, "0")
                Else
                    MhRealInput1.Text = Format(Int(Val(MhRealInput2.Text) / RPB), "0")
                End If
            End If
        End If
    End If
End Sub
Private Sub MhRealInput1_GotFocus()
    FocusSelect Me.ActiveControl
End Sub
Private Sub MhRealInput1_KeyPress(KeyAscii As Integer)
    ValidateKey MhRealInput2, KeyAscii, 0
End Sub
Private Sub MhRealInput1_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyReturn Then
        If Not ValidateNumber(Me.ActiveControl, 0) Then Exit Sub
        rstPaperChild.Fields("Account").Value = AccountCode
        rstPaperChild.Fields("GodownName").Value = Trim(Text12.Text)
        rstPaperChild.Fields("OpBalOther").Value = Format(Val(MhRealInput2.Text), "0.000")
        rstPaperChild.Fields("OpBalTat").Value = Format(Val(MhRealInput1.Text), "0")
        rstPaperChild.Fields("Imported").Value = "N"
        rstPaperChild.Update
        MakeTextBoxInvisible (False)
        If rstPaperChild.AbsolutePosition = rstPaperChild.RecordCount Then
            Call DataGrid2_KeyDown(vbKeyA, vbCtrlMask)
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
        If Not (rstPaperChild.EOF Or rstPaperChild.BOF) Then
            If Val(CheckNull(rstPaperChild.Fields("OpBalOther").Value)) = 0 And Val(CheckNull(rstPaperChild.Fields("OpBalTat").Value)) = 0 Then
                rstPaperChild.Delete
                rstPaperChild.MoveNext
                If rstPaperChild.RecordCount > 0 Then rstPaperChild.MoveFirst
            End If
        End If
    End If
    Text12.Visible = False
    MhRealInput2.Visible = False
    MhRealInput1.Visible = False
    DataGrid2.Enabled = True
    If Mh3dFrame3.Enabled Then
        DataGrid2.SetFocus
    End If
End Sub
Public Sub FilterRecord(ByVal SrchFor As String, ByVal SrchText As String)
    If SrchFor = "Name" Then
        rstPaperList.Filter = "[Name] Like '%" & SrchText & "%'"
    End If
End Sub
Private Function CheckDuplicateGodown() As Boolean
    Dim dblBookMark As Double
    
    If rstPaperChild.RecordCount = 0 Then Exit Function
    If Not (rstPaperChild.EOF Or rstPaperChild.BOF) Then
       dblBookMark = rstPaperChild.Bookmark
    End If
    rstPaperChild.MoveFirst
    Do While Not rstPaperChild.EOF
          If rstPaperChild.Fields("GodownName").Value = Trim(Text12.Text) Then
             CheckDuplicateGodown = True
             Exit Do
          End If
          rstPaperChild.MoveNext
    Loop
    If dblBookMark <> 0 Then
       rstPaperChild.Bookmark = dblBookMark
    Else
       rstPaperChild.MoveLast
    End If
End Function
