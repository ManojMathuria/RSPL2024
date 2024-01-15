VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmPaperPurchaseOrderNew 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Outsource Item Purchase Order With Issue"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15360
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
   ScaleHeight     =   7245
   ScaleWidth      =   15360
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   7170
      Left            =   15
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   0
      Width           =   15150
      _Version        =   65536
      _ExtentX        =   26723
      _ExtentY        =   12647
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
      Picture         =   "PaperPurchaseOrderNew.frx":0000
      Begin TabDlg.SSTab SSTab1 
         Height          =   6945
         Left            =   120
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   120
         Width           =   14925
         _ExtentX        =   26326
         _ExtentY        =   12250
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         Tab             =   1
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
         TabPicture(0)   =   "PaperPurchaseOrderNew.frx":001C
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Text1"
         Tab(0).Control(1)=   "DataGrid1"
         Tab(0).Control(2)=   "Label1"
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "&Details"
         TabPicture(1)   =   "PaperPurchaseOrderNew.frx":0038
         Tab(1).ControlEnabled=   -1  'True
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
            Left            =   -74395
            MaxLength       =   40
            TabIndex        =   16
            Top             =   6510
            Width           =   8700
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   6300
            Left            =   120
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   480
            Width           =   14100
            _Version        =   65536
            _ExtentX        =   24871
            _ExtentY        =   11112
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
            Picture         =   "PaperPurchaseOrderNew.frx":0054
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
               Left            =   1530
               MaxLength       =   40
               TabIndex        =   9
               Top             =   5835
               Width           =   7545
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
               MaxLength       =   30
               TabIndex        =   10
               Top             =   5380
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
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1560
               MaxLength       =   10
               TabIndex        =   0
               Top             =   105
               Width           =   2250
            End
            Begin VB.TextBox Text4 
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
               MaxLength       =   40
               TabIndex        =   4
               Top             =   950
               Width           =   7515
            End
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
               Left            =   1560
               MaxLength       =   40
               TabIndex        =   3
               Top             =   630
               Width           =   7515
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel9 
               Height          =   330
               Left            =   120
               TabIndex        =   19
               Top             =   5380
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
               Picture         =   "PaperPurchaseOrderNew.frx":0070
               Picture         =   "PaperPurchaseOrderNew.frx":008C
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
               Height          =   330
               Left            =   120
               TabIndex        =   20
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
               Picture         =   "PaperPurchaseOrderNew.frx":00A8
               Picture         =   "PaperPurchaseOrderNew.frx":00C4
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Left            =   3765
               TabIndex        =   21
               Top             =   105
               Width           =   1740
               _Version        =   65536
               _ExtentX        =   3069
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
               Picture         =   "PaperPurchaseOrderNew.frx":00E0
               Picture         =   "PaperPurchaseOrderNew.frx":00FC
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel7 
               Height          =   330
               Left            =   120
               TabIndex        =   22
               Top             =   4525
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
               Caption         =   " VAT (%)"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "PaperPurchaseOrderNew.frx":0118
               Picture         =   "PaperPurchaseOrderNew.frx":0134
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   330
               Left            =   120
               TabIndex        =   23
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
               Picture         =   "PaperPurchaseOrderNew.frx":0150
               Picture         =   "PaperPurchaseOrderNew.frx":016C
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel13 
               Height          =   330
               Left            =   6780
               TabIndex        =   24
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
               Caption         =   " Delivery Date"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "PaperPurchaseOrderNew.frx":0188
               Picture         =   "PaperPurchaseOrderNew.frx":01A4
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
               Height          =   330
               Left            =   120
               TabIndex        =   25
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
               Picture         =   "PaperPurchaseOrderNew.frx":01C0
               Picture         =   "PaperPurchaseOrderNew.frx":01DC
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
               Height          =   330
               Left            =   6780
               TabIndex        =   26
               Top             =   4875
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
               Picture         =   "PaperPurchaseOrderNew.frx":01F8
               Picture         =   "PaperPurchaseOrderNew.frx":0214
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel14 
               Height          =   330
               Left            =   120
               TabIndex        =   27
               Top             =   4868
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
               Picture         =   "PaperPurchaseOrderNew.frx":0230
               Picture         =   "PaperPurchaseOrderNew.frx":024C
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel16 
               Height          =   330
               Left            =   6780
               TabIndex        =   28
               Top             =   4530
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
               Caption         =   " VAT"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "PaperPurchaseOrderNew.frx":0268
               Picture         =   "PaperPurchaseOrderNew.frx":0284
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel20 
               Height          =   330
               Left            =   6780
               TabIndex        =   29
               Top             =   5380
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
               Picture         =   "PaperPurchaseOrderNew.frx":02A0
               Picture         =   "PaperPurchaseOrderNew.frx":02BC
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel21 
               Height          =   330
               Left            =   3045
               TabIndex        =   30
               Top             =   5385
               Width           =   1980
               _Version        =   65536
               _ExtentX        =   3492
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
               Picture         =   "PaperPurchaseOrderNew.frx":02D8
               Picture         =   "PaperPurchaseOrderNew.frx":02F4
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel32 
               Height          =   330
               Left            =   120
               TabIndex        =   31
               Top             =   5835
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
               Picture         =   "PaperPurchaseOrderNew.frx":0310
               Picture         =   "PaperPurchaseOrderNew.frx":032C
            End
            Begin TDBDate6Ctl.TDBDate MhDateInput3 
               Height          =   330
               Left            =   7980
               TabIndex        =   2
               Top             =   105
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
               _ExtentY        =   582
               Calendar        =   "PaperPurchaseOrderNew.frx":0348
               Caption         =   "PaperPurchaseOrderNew.frx":0460
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PaperPurchaseOrderNew.frx":04CC
               Keys            =   "PaperPurchaseOrderNew.frx":04EA
               Spin            =   "PaperPurchaseOrderNew.frx":0548
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
            Begin TDBDate6Ctl.TDBDate MhDateInput1 
               Height          =   330
               Left            =   5490
               TabIndex        =   1
               Top             =   105
               Width           =   1335
               _Version        =   65536
               _ExtentX        =   2355
               _ExtentY        =   582
               Calendar        =   "PaperPurchaseOrderNew.frx":0570
               Caption         =   "PaperPurchaseOrderNew.frx":0688
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PaperPurchaseOrderNew.frx":06F4
               Keys            =   "PaperPurchaseOrderNew.frx":0712
               Spin            =   "PaperPurchaseOrderNew.frx":0770
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
               Height          =   1095
               Left            =   120
               TabIndex        =   5
               Top             =   1470
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
               MaxCols         =   4
               MaxRows         =   1000
               ScrollBars      =   2
               SpreadDesigner  =   "PaperPurchaseOrderNew.frx":0798
            End
            Begin FPSpreadADO.fpSpread fpSpread2 
               Height          =   1095
               Left            =   120
               TabIndex        =   6
               Top             =   3015
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
               MaxCols         =   4
               MaxRows         =   1000
               ScrollBars      =   2
               SpreadDesigner  =   "PaperPurchaseOrderNew.frx":0EAE
            End
            Begin TDBDate6Ctl.TDBDate MhDateInput2 
               Height          =   330
               Left            =   5010
               TabIndex        =   11
               Top             =   5385
               Width           =   1815
               _Version        =   65536
               _ExtentX        =   3201
               _ExtentY        =   582
               Calendar        =   "PaperPurchaseOrderNew.frx":1538
               Caption         =   "PaperPurchaseOrderNew.frx":1650
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PaperPurchaseOrderNew.frx":16BC
               Keys            =   "PaperPurchaseOrderNew.frx":16DA
               Spin            =   "PaperPurchaseOrderNew.frx":1738
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
               Left            =   7980
               TabIndex        =   12
               Top             =   5380
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
               _ExtentY        =   582
               Calculator      =   "PaperPurchaseOrderNew.frx":1760
               Caption         =   "PaperPurchaseOrderNew.frx":1780
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PaperPurchaseOrderNew.frx":17EC
               Keys            =   "PaperPurchaseOrderNew.frx":180A
               Spin            =   "PaperPurchaseOrderNew.frx":1854
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
               ValueVT         =   1996816385
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput11 
               Height          =   330
               Left            =   1560
               TabIndex        =   7
               Top             =   4525
               Width           =   1530
               _Version        =   65536
               _ExtentX        =   2699
               _ExtentY        =   582
               Calculator      =   "PaperPurchaseOrderNew.frx":187C
               Caption         =   "PaperPurchaseOrderNew.frx":189C
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PaperPurchaseOrderNew.frx":1908
               Keys            =   "PaperPurchaseOrderNew.frx":1926
               Spin            =   "PaperPurchaseOrderNew.frx":1970
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
               ValueVT         =   1969487877
               Value           =   5
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput12 
               Height          =   330
               Left            =   7980
               TabIndex        =   32
               TabStop         =   0   'False
               Top             =   4530
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
               _ExtentY        =   582
               Calculator      =   "PaperPurchaseOrderNew.frx":1998
               Caption         =   "PaperPurchaseOrderNew.frx":19B8
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PaperPurchaseOrderNew.frx":1A24
               Keys            =   "PaperPurchaseOrderNew.frx":1A42
               Spin            =   "PaperPurchaseOrderNew.frx":1A8C
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
               ForeColor       =   255
               Format          =   "#########0.00"
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
               ValueVT         =   1996816385
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput14 
               Height          =   330
               Left            =   1560
               TabIndex        =   8
               Top             =   4868
               Width           =   1530
               _Version        =   65536
               _ExtentX        =   2699
               _ExtentY        =   582
               Calculator      =   "PaperPurchaseOrderNew.frx":1AB4
               Caption         =   "PaperPurchaseOrderNew.frx":1AD4
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PaperPurchaseOrderNew.frx":1B40
               Keys            =   "PaperPurchaseOrderNew.frx":1B5E
               Spin            =   "PaperPurchaseOrderNew.frx":1BA8
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
               MinValue        =   -9999999.99
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   ""
               ShowContextMenu =   1
               ValueVT         =   1969487877
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput15 
               Height          =   330
               Left            =   7980
               TabIndex        =   33
               TabStop         =   0   'False
               Top             =   4875
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
               _ExtentY        =   582
               Calculator      =   "PaperPurchaseOrderNew.frx":1BD0
               Caption         =   "PaperPurchaseOrderNew.frx":1BF0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PaperPurchaseOrderNew.frx":1C5C
               Keys            =   "PaperPurchaseOrderNew.frx":1C7A
               Spin            =   "PaperPurchaseOrderNew.frx":1CC4
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
               ForeColor       =   255
               Format          =   "#########0.00"
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
               ValueVT         =   1996816385
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
               Height          =   285
               Left            =   120
               TabIndex        =   34
               Top             =   2550
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
               Picture         =   "PaperPurchaseOrderNew.frx":1CEC
               Picture         =   "PaperPurchaseOrderNew.frx":1D08
               Begin TDBNumber6Ctl.TDBNumber MhRealInput17 
                  Height          =   285
                  Left            =   4845
                  TabIndex        =   36
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   1200
                  _Version        =   65536
                  _ExtentX        =   2117
                  _ExtentY        =   503
                  Calculator      =   "PaperPurchaseOrderNew.frx":1D24
                  Caption         =   "PaperPurchaseOrderNew.frx":1D44
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "PaperPurchaseOrderNew.frx":1DB0
                  Keys            =   "PaperPurchaseOrderNew.frx":1DCE
                  Spin            =   "PaperPurchaseOrderNew.frx":1E18
                  AlignHorizontal =   1
                  AlignVertical   =   0
                  Appearance      =   0
                  BackColor       =   16777215
                  BorderStyle     =   1
                  BtnPositioning  =   0
                  ClipMode        =   0
                  ClearAction     =   0
                  DecimalPoint    =   "."
                  DisplayFormat   =   ""
                  EditMode        =   1
                  Enabled         =   -1
                  ErrorBeep       =   0
                  ForeColor       =   255
                  Format          =   ""
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
                  ValueVT         =   1179653
                  Value           =   0
                  MaxValueVT      =   5
                  MinValueVT      =   5
               End
               Begin TDBNumber6Ctl.TDBNumber MhRealInput18 
                  Height          =   285
                  Left            =   7155
                  TabIndex        =   37
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   1545
                  _Version        =   65536
                  _ExtentX        =   2725
                  _ExtentY        =   503
                  Calculator      =   "PaperPurchaseOrderNew.frx":1E40
                  Caption         =   "PaperPurchaseOrderNew.frx":1E60
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "PaperPurchaseOrderNew.frx":1ECC
                  Keys            =   "PaperPurchaseOrderNew.frx":1EEA
                  Spin            =   "PaperPurchaseOrderNew.frx":1F34
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
                  ValueVT         =   1999372293
                  Value           =   0
                  MaxValueVT      =   5
                  MinValueVT      =   5
               End
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel17 
               Height          =   285
               Left            =   120
               TabIndex        =   35
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
               Picture         =   "PaperPurchaseOrderNew.frx":1F5C
               Picture         =   "PaperPurchaseOrderNew.frx":1F78
               Begin TDBNumber6Ctl.TDBNumber MhRealInput20 
                  Height          =   285
                  Left            =   7680
                  TabIndex        =   38
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   1040
                  _Version        =   65536
                  _ExtentX        =   1834
                  _ExtentY        =   503
                  Calculator      =   "PaperPurchaseOrderNew.frx":1F94
                  Caption         =   "PaperPurchaseOrderNew.frx":1FB4
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "PaperPurchaseOrderNew.frx":2020
                  Keys            =   "PaperPurchaseOrderNew.frx":203E
                  Spin            =   "PaperPurchaseOrderNew.frx":2088
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
                  ValueVT         =   1179653
                  Value           =   0
                  MaxValueVT      =   5
                  MinValueVT      =   5
               End
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
               Left            =   480
               MaxLength       =   100
               TabIndex        =   39
               TabStop         =   0   'False
               Top             =   2160
               Width           =   8595
            End
            Begin VB.Line Line6 
               X1              =   0
               X2              =   11560
               Y1              =   5775
               Y2              =   5775
            End
            Begin VB.Line Line5 
               X1              =   0
               X2              =   11560
               Y1              =   5295
               Y2              =   5295
            End
            Begin VB.Line Line1 
               X1              =   0
               X2              =   11560
               Y1              =   525
               Y2              =   525
            End
            Begin VB.Line Line2 
               X1              =   0
               X2              =   11560
               Y1              =   1365
               Y2              =   1365
            End
            Begin VB.Line Line3 
               X1              =   0
               X2              =   11560
               Y1              =   2920
               Y2              =   2920
            End
            Begin VB.Line Line4 
               X1              =   0
               X2              =   11560
               Y1              =   4445
               Y2              =   4445
            End
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   6000
            Left            =   -74880
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   480
            Width           =   9195
            _ExtentX        =   16219
            _ExtentY        =   10583
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
                  ColumnWidth     =   5249.764
               EndProperty
               BeginProperty Column03 
                  Alignment       =   1
                  Locked          =   -1  'True
                  ColumnWidth     =   1260.284
               EndProperty
            EndProperty
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
            Left            =   -74880
            TabIndex        =   17
            Top             =   6510
            Width           =   495
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   15360
      _ExtentX        =   27093
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
Attribute VB_Name = "FrmPaperPurchaseOrderNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CxnPaperPurchaseOrder As New ADODB.Connection
Dim rstPaperPOList As New ADODB.Recordset
Dim rstPaperPOParent As New ADODB.Recordset
Dim rstPaperPOChild As New ADODB.Recordset
Dim rstSupplierList As New ADODB.Recordset
Dim rstPaperList As New ADODB.Recordset
Dim rstAccountList As New ADODB.Recordset
Dim rstLastPurchaseRate As New ADODB.Recordset
Dim SupplierCode As String, AccountCode As String, PaperCode As String
Dim SortOrder, PrevStr
Dim dblBookMark As Double
Dim blnRecordExist As Boolean
Dim oOutlook As New Outlook.Application
Dim EditMode As Boolean
Dim EMailID, Attachment, Message
Public OrderType

'By Shamshad

Dim CxnOutsourceItemPurchaseOrder As New ADODB.Connection
Dim rstCompanyMaster As New ADODB.Recordset
Dim rstOutsourceItemPOList As New ADODB.Recordset
Dim rstOutsourceItemPOParent As New ADODB.Recordset
Dim rstOutsourceItemPOChild As New ADODB.Recordset
Dim rstOutsourceItemList As New ADODB.Recordset
Dim OutsourceItemCode As String
Dim OutputTo As String

Private Sub Form_Load()
'    On Error GoTo ErrorHandler
'    CenterForm Me
'    WheelHook DataGrid1
'    BusySystemIndicator True
'    Me.Caption = "Paper Purchase Order [" & IIf(OrderType = "1", "Book", "Title") & "]"
'    If OrderType = "2" Then fpSpread2.ColWidth(1) = 48.34: fpSpread2.ColWidth(4) = 0
'    CxnPaperPurchaseOrder.CursorLocation = adUseClient
'    CxnPaperPurchaseOrder.Open CxnDatabase.ConnectionString
'    rstPaperList.Open "SELECT Name As Col0,[Weight/Ream],[Reams/Bundle],Code FROM PaperMaster WHERE Type = '" & OrderType & "' ORDER BY Name", CxnPaperPurchaseOrder, adOpenKeyset, adLockReadOnly
'    rstSupplierList.Open "SELECT Name As Col0,Code FROM AccountMaster WHERE Type='01' ORDER BY Name", CxnPaperPurchaseOrder, adOpenKeyset, adLockReadOnly
'    rstAccountList.Open "SELECT TRIM(Name)+' ('+CHOOSE(VAL(Type)-4,'Book Printer','Title Printer','','Book Binder','Godown')+')' As Col0,Code FROM AccountMaster WHERE Type IN ('05','06','08','09') ORDER BY Name", CxnPaperPurchaseOrder, adOpenKeyset, adLockReadOnly
'    rstPaperPOList.Open "SELECT T.Code,T.Name,Date,M.Name As SupplierName,BillAmount FROM PaperPOParent T INNER JOIN AccountMaster M ON T.Supplier=M.Code WHERE OrderType='" & OrderType & "' ORDER BY T.Name", CxnPaperPurchaseOrder, adOpenKeyset, adLockOptimistic
'    rstPaperPOParent.CursorLocation = adUseClient
'    rstPaperPOList.Filter = adFilterNone
'    If rstPaperPOList.RecordCount > 0 Then rstPaperPOList.MoveLast
'    Set DataGrid1.DataSource = rstPaperPOList
'    BusySystemIndicator False
'    SSTab1.Tab = 0
'    SortOrder = "Name"
'    If Not (rstPaperPOList.EOF Or rstPaperPOList.BOF) Then
'        With DataGrid1.SelBookmarks
'            If .Count <> 0 Then .Remove 0
'            .Add DataGrid1.Bookmark
'        End With
'    End If
'    rstPaperPOList.ActiveConnection = Nothing
'    rstPaperList.ActiveConnection = Nothing
'    rstSupplierList.ActiveConnection = Nothing
'    rstAccountList.ActiveConnection = Nothing
'    SetButtonsForNoRecord
'    Exit Sub
'ErrorHandler:
'    BusySystemIndicator False
'    Unload Me

 ' On Error GoTo ErrorHandler
    CenterForm Me
    WheelHook DataGrid1
    BusySystemIndicator True
    CxnOutsourceItemPurchaseOrder.CursorLocation = adUseClient
    CxnOutsourceItemPurchaseOrder.Open CxnDatabase.ConnectionString
    rstCompanyMaster.Open "Select PrintName, Address1, Address2, Address3, Address4, Phone, Fax, EMail, Website From CompanyMaster", CxnOutsourceItemPurchaseOrder, adOpenKeyset, adLockReadOnly
    rstOutsourceItemList.Open "Select Name As Col0,Code From OutsourceItemMaster Order By Name", CxnOutsourceItemPurchaseOrder, adOpenKeyset, adLockReadOnly
    rstSupplierList.Open "Select Name As Col0, Code From AccountMaster Where Type = '01' Order By Name", CxnOutsourceItemPurchaseOrder, adOpenKeyset, adLockReadOnly
    rstOutsourceItemPOList.Open "Select T.Code,T.Name,T.Date,M.Name As SupplierName,T.BillAmount From OutsourceItemPOParent T,AccountMaster M Where T.Supplier = M.Code Order By T.Name", CxnOutsourceItemPurchaseOrder, adOpenKeyset, adLockOptimistic
    rstOutsourceItemPOParent.CursorLocation = adUseClient
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
    rstOutsourceItemList.ActiveConnection = Nothing
    rstSupplierList.ActiveConnection = Nothing
    SetButtonsForNoRecord
    Exit Sub
'ErrorHandler:
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
           If Me.ActiveControl.Name <> "fpSpread1" And Me.ActiveControl.Name <> "fpSpread2" Then SendKeys "{TAB}"
        End If
        If Me.ActiveControl.Name <> "fpSpread1" And Me.ActiveControl.Name <> "fpSpread2" Then KeyCode = 0
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
    Call CloseRecordset(rstPaperList)
    Call CloseRecordset(rstSupplierList)
    Call CloseRecordset(rstAccountList)
    Call CloseRecordset(rstLastPurchaseRate)
    Call CloseConnection(CxnPaperPurchaseOrder)
    ShowProgressInStatusBar False
    DisableChildMenu
End Sub

Private Sub Mh3dLabel6_Click()

End Sub

Private Sub Mh3dLabel22_Click()

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
'    Dim HiLiteRecord As Boolean
'    Dim UpdateFlag As Integer
'    Dim CellVal01 As Variant, CellVal02 As Variant, CellVal03 As Variant, i As Integer
'    If Button.Index = 1 Then
'        If rstPaperPOParent.State = adStateOpen Then rstPaperPOParent.Close
'        rstPaperPOParent.Open "SELECT * FROM PaperPOParent WHERE Code=''", CxnPaperPurchaseOrder, adOpenKeyset, adLockOptimistic
'        ClearFields
'        Call LoadPaperList("")
'        If AddRecord(rstPaperPOParent) Then
'            Text2.Text = GenerateCode(CxnPaperPurchaseOrder, "SELECT MAX(VAL(Name)) FROM PaperPOParent WHERE OrderType='" & OrderType & "'", 10, Space(1))
'            MhDateInput1.Text = Format(Date, "dd-MM-yyyy")
'            Call SetButtons(False)
'            SSTab1.Tab = 1
'            Text3.SetFocus
'            blnRecordExist = False
'            CxnPaperPurchaseOrder.BeginTrans
'        End If
'    ElseIf Button.Index = 2 Then
'        If rstPaperPOList.RecordCount = 0 Then Exit Sub
'        SSTab1.Tab = 1
'        EditRecord
'    ElseIf Button.Index = 3 Then
'        If rstPaperPOList.RecordCount = 0 Then Exit Sub
'        If AllowTransactionsDeletion = 0 Then Call DisplayError("You don't have the rights to Delete this Voucher"): Exit Sub
'        SSTab1.Tab = 1
'        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
'            On Error Resume Next
'            MdiMainMenu.MousePointer = vbHourglass
'            CxnPaperPurchaseOrder.Execute "DELETE FROM PaperPOParent WHERE Code='" & rstPaperPOList.Fields("Code").Value & "'"
'            MdiMainMenu.MousePointer = vbNormal
'            If Err.Number = 0 Then
'                rstPaperPOList.Delete
'                rstPaperPOList.MoveNext
'                If rstPaperPOList.RecordCount > 0 And rstPaperPOList.EOF Then rstPaperPOList.MoveLast
'                ShowProgressInStatusBar True
'                Timer1.Enabled = True
'            Else
'                DisplayError ("Failed to delete the record")
'            End If
'            On Error GoTo 0
'        End If
'        SetButtons (True)
'        SetButtonsForNoRecord
'        SSTab1.Tab = 0
'        HiLiteRecord = True
'    ElseIf Button.Index = 4 Then
'        If CheckMandatoryFields Then Exit Sub
'        SaveFields
'        UpdateFlag = 0
'        If UpdateRecord(rstPaperPOParent) Then
'            If UpdatePaperList("D") Then
'                UpdateFlag = 1
'                With fpSpread1
'                    For i = 1 To .DataRowCnt
'                        .SetActiveCell 6, i
'                        .GetText 6, i, CellVal01
'                        .GetText 7, i, CellVal02
'                        If Val(CellVal01) <> 0 And CellVal02 <> "" Then
'                            If Not UpdatePaperList("I1") Then UpdateFlag = 0: Exit For
'                        End If
'                    Next
'                End With
'                If UpdateFlag = 1 Then
'                    With fpSpread2
'                        For i = 1 To .DataRowCnt
'                            .SetActiveCell 3, i
'                            .GetText 3, i, CellVal01
'                            .GetText 5, i, CellVal02
'                            .GetText 6, i, CellVal03
'                            If Val(CellVal01) <> 0 And CellVal02 <> "" And CellVal03 <> "" Then
'                                If Not UpdatePaperList("I2") Then UpdateFlag = 0: Exit For
'                            End If
'                        Next
'                    End With
'                End If
'            End If
'        End If
'        If UpdateFlag Then
'            AddToList
'            CxnPaperPurchaseOrder.CommitTrans
'            If rstPaperPOParent.State = adStateOpen Then rstPaperPOParent.Close
'            rstPaperPOParent.CursorLocation = adUseClient
'            Call SetButtons(True)
'            SSTab1.Tab = 0
'            ShowProgressInStatusBar True
'            Timer1.Enabled = True
'            LockFields (False)
'        Else
'            DisplayError ("Failed to save the record")
'            Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
'        End If
'    ElseIf Button.Index = 5 Then
'        If CancelRecordUpdate(rstPaperPOParent) Then
'            CxnPaperPurchaseOrder.RollbackTrans
'            If rstPaperPOParent.State = adStateOpen Then rstPaperPOParent.Close
'            rstPaperPOParent.CursorLocation = adUseClient
'            Call SetButtons(True)
'            SetButtonsForNoRecord
'            SSTab1.Tab = 0
'            LockFields (False)
'        End If
'    ElseIf Button.Index = 6 Then
'        SSTab1.Tab = 0
'        Set DataGrid1.DataSource = Nothing
'        rstPaperPOList.ActiveConnection = CxnPaperPurchaseOrder
'        Do While Not RefreshRecord(rstPaperPOList)
'        Loop
'        Set DataGrid1.DataSource = rstPaperPOList
'        rstPaperPOList.ActiveConnection = Nothing
'        If rstPaperPOList.RecordCount > 0 Then rstPaperPOList.MoveLast
'        rstSupplierList.ActiveConnection = CxnPaperPurchaseOrder
'        Do While Not RefreshRecord(rstSupplierList)
'        Loop
'        rstSupplierList.ActiveConnection = Nothing
'        rstPaperList.ActiveConnection = CxnPaperPurchaseOrder
'        Do While Not RefreshRecord(rstPaperList)
'        Loop
'        rstPaperList.ActiveConnection = Nothing
'        rstAccountList.ActiveConnection = CxnPaperPurchaseOrder
'        Do While Not RefreshRecord(rstAccountList)
'        Loop
'        rstAccountList.ActiveConnection = Nothing
'        HiLiteRecord = True
'    ElseIf Button.Index = 7 Then
'        SSTab1.Tab = 0
'        With FrmFilter
'            .Combo1.AddItem "Supplier", 0
'            .Combo1.ListIndex = 0
'            Set .srcForm = Me
'            .Show vbModal
'        End With
'        HiLiteRecord = True
'    ElseIf Button.Index = 9 Then
'        If rstPaperPOList.RecordCount = 0 Then Exit Sub
'        Call DisplayMenu("P")
'        HiLiteRecord = True
'    ElseIf Button.Index = 10 Then
'        If rstPaperPOList.RecordCount = 0 Then Exit Sub
'        Call DisplayMenu("S")
'        HiLiteRecord = True
'    ElseIf Button.Index = 11 Then
'        If rstPaperPOList.RecordCount = 0 Then Exit Sub
'        Call DisplayMenu("M")
'        HiLiteRecord = True
'    ElseIf Button.Index = 13 Then
'        If rstPaperPOList.RecordCount > 0 Then rstPaperPOList.MoveFirst
'        HiLiteRecord = True
'        ViewRecord
'    ElseIf Button.Index = 14 Then
'        If rstPaperPOList.RecordCount > 0 Then
'            rstPaperPOList.MovePrevious
'            If rstPaperPOList.BOF Then rstPaperPOList.MoveNext
'        End If
'        HiLiteRecord = True
'        ViewRecord
'    ElseIf Button.Index = 15 Then
'        If rstPaperPOList.RecordCount > 0 Then
'            rstPaperPOList.MoveNext
'            If rstPaperPOList.EOF Then rstPaperPOList.MovePrevious
'        End If
'        HiLiteRecord = True
'        ViewRecord
'    ElseIf Button.Index = 16 Then
'        If rstPaperPOList.RecordCount > 0 Then rstPaperPOList.MoveLast
'        HiLiteRecord = True
'        ViewRecord
'    ElseIf Button.Index = 18 Then
'        Unload Me
'        HiLiteRecord = False
'    End If
'
'    If HiLiteRecord Then
'        If Not (rstPaperPOList.EOF Or rstPaperPOList.BOF) Then
'            With DataGrid1.SelBookmarks
'                If .Count <> 0 Then .Remove 0
'                .Add DataGrid1.Bookmark
'            End With
'        End If
'        Text1.SetFocus
'    End If

    Dim HiLiteRecord As Boolean
    Dim UpdateFlag As Integer
    
    If Button.Index = 1 Then
        If rstOutsourceItemPOParent.State = adStateOpen Then
           rstOutsourceItemPOParent.Close
        End If
        rstOutsourceItemPOParent.Open "Select * From OutsourceItemPOParent Where Code = ''", CxnOutsourceItemPurchaseOrder, adOpenKeyset, adLockOptimistic
        'ClearFields ("P")
        'ClearFields ("C")
        ClearFields
        Call LoadOutsourceItemList("")
        If rstOutsourceItemPOChild.State = adStateClosed Then
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
'        If blnRecordExist And AllowTransactionsModification = 0 Then
'            Call DisplayError("You don't have the rights to Edit this Voucher")
'            Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
'            Exit Sub
'        End If
        SaveFields
        UpdateFlag = 0
        If UpdateRecord(rstOutsourceItemPOParent) Then
            If UpdateOutsourceItemList("D") Then
                 UpdateFlag = 1
                 If rstOutsourceItemPOChild.RecordCount <> 0 Then
                      rstOutsourceItemPOChild.MoveFirst
                      Do While Not rstOutsourceItemPOChild.EOF
                          If Val(rstOutsourceItemPOChild.Fields("Quantity").Value) <> 0 Then
                               If Not UpdateOutsourceItemList("U") Then
                                    UpdateFlag = 0
                                    Exit Do
                                End If
                          End If
                          rstOutsourceItemPOChild.MoveNext
                      Loop
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
        PrintOutsourceItemPurchaseOrder
        HiLiteRecord = True
    ElseIf Button.Index = 10 Then
        If rstOutsourceItemPOList.RecordCount = 0 Then Exit Sub
        OutputTo = "S"
        PrintOutsourceItemPurchaseOrder
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
    'If Toolbar1.Buttons.Item(2).Enabled Then Toolbar1_ButtonClick Toolbar1.Buttons.Item(2)
    
    If Toolbar1.Buttons.Item(2).Enabled Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(2)
    End If
    
End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
'    If ColIndex = 0 Or ColIndex = 2 Then
'        SortOrder = DataGrid1.Columns(ColIndex).DataField
'        rstPaperPOList.Sort = "[" + SortOrder & "] Asc"
'    End If
'    DataGrid1.ClearSelCols
'    If Not (rstPaperPOList.EOF Or rstPaperPOList.BOF) Then
'        With DataGrid1.SelBookmarks
'            If .Count <> 0 Then .Remove 0
'            .Add DataGrid1.Bookmark
'        End With
'    End If
'    Text1.Text = ""
'    Text1.SetFocus


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
'    If rstPaperPOList.RecordCount = 0 Then
'        Toolbar1.Buttons.Item(2).Enabled = False
'        Toolbar1.Buttons.Item(3).Enabled = False
'        Toolbar1.Buttons.Item(9).Enabled = False
'        Toolbar1.Buttons.Item(10).Enabled = False
'        Toolbar1.Buttons.Item(11).Enabled = False
'        Toolbar1.Buttons.Item(13).Enabled = False
'        Toolbar1.Buttons.Item(14).Enabled = False
'        Toolbar1.Buttons.Item(15).Enabled = False
'        Toolbar1.Buttons.Item(16).Enabled = False
'    End If
     If rstOutsourceItemPOList.RecordCount = 0 Then
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
    If rstPaperPOParent.EOF Or rstPaperPOParent.BOF Then Exit Sub
    If CheckEmpty(Text2, True) Then
        Cancel = True
    ElseIf CheckDuplicate(CxnPaperPurchaseOrder, "PaperPOParent", "Code", "[Name]+OrderType", Trim(Text2.Text) & OrderType, rstPaperPOParent.Fields("Code").Value, False) Then
        Cancel = True
    End If
    
End Sub
Private Sub MhDateInput1_Validate(Cancel As Boolean)
    If Not IsDate(GetDate(MhDateInput1.Text)) Then
        Cancel = True
    ElseIf Format(GetDate(MhDateInput1.Text), "yyyymmdd") < Format(FinancialYearFrom, "yyyymmdd") Or Format(GetDate(MhDateInput1.Text), "yyyymmdd") > Format(FinancialYearTo, "yyyymmdd") Then
        Cancel = True
    ElseIf Not blnRecordExist Then
        MhDateInput3.Text = Format(DateAdd("d", 1, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
    End If
        
End Sub
Private Sub Text3_Change()
    If Text3.Text = " " Then Text3.Text = "?": SendKeys "{TAB}"
End Sub
Private Sub Text3_Validate(Cancel As Boolean)
    Dim SearchString As String
    SearchString = FixQuote(Text3.Text)
    If rstSupplierList.RecordCount = 0 Then DisplayError ("No Record in Supplier Master"): Cancel = True: Exit Sub Else rstSupplierList.MoveFirst
    rstSupplierList.Find "[Col0] = '" & RTrim(SearchString) & "'"
    If rstSupplierList.EOF Then
        SelectionType = "S"
        SupplierCode = ""
        Call LoadSelectionList(rstSupplierList, "List of Suppliers...", "Name")
        SearchOrder = 0
        Call DisplaySelectionList(Text3, SupplierCode)
        Call CloseForm(FrmSelectionList)
        If CheckEmpty(Text3.Text, False) Then Text3.Text = "?"
        If RTrim(SupplierCode) <> "" Then SendKeys "{TAB}"
        Cancel = True
    Else
        SupplierCode = rstSupplierList.Fields("Code").Value
    End If
End Sub
Private Sub MhDateInput2_Validate(Cancel As Boolean)
    If MhDateInput2.ValueIsNull Then Exit Sub
    If Not IsDate(GetDate(MhDateInput2.Text)) Then Cancel = True
End Sub
Private Sub MhDateInput3_Validate(Cancel As Boolean)
    If Not IsDate(GetDate(MhDateInput3.Text)) Then Cancel = True
End Sub
Private Sub MhDateInput4_Validate(Cancel As Boolean)
    If MhDateInput4.ValueIsNull Then Exit Sub
    If Not IsDate(GetDate(MhDateInput4.Text)) Then Cancel = True
End Sub
Private Sub MhDateInput5_Validate(Cancel As Boolean)
    If MhDateInput5.ValueIsNull Then Exit Sub
    If Not IsDate(GetDate(MhDateInput5.Text)) Then Cancel = True
End Sub
Private Sub MhDateInput6_Validate(Cancel As Boolean)
    If MhDateInput6.ValueIsNull Then Exit Sub
    If Not IsDate(GetDate(MhDateInput6.Text)) Then Cancel = True
End Sub
Private Sub MhRealInput8_Validate(Cancel As Boolean)    'Reams/bundle
    If Val(MhRealInput8.Text) > 0 Then MhRealInput9.Value = Int(Val(MhRealInput17.Text) / Val(MhRealInput8.Text)) + IIf(Int(Val(MhRealInput17.Text)) * 500 + (Val(MhRealInput17.Text) - Int(Val(MhRealInput17.Text))) * 1000 Mod Val(MhRealInput8.Text) * 500 > 0, 1, 0)    'Total bundles
End Sub
Private Sub MhRealInput9_Validate(Cancel As Boolean)    'Total bundles
    CalculateCartage
End Sub
Private Sub MhRealInput10_Validate(Cancel As Boolean)   'Cartage/Kg
    CalculateCartage
End Sub
Private Sub MhRealInput11_Validate(Cancel As Boolean)   'VAT (%)
    MhRealInput12.Value = Val(MhRealInput19.Text) * Val(MhRealInput11.Text) / 100   'VAT
    Call CalculateTotal("N")    'VAT Changed
End Sub
Private Sub MhRealInput13_Validate(Cancel As Boolean)   'Cartage
    Call CalculateTotal("N")    'Cartage Changed
    If Not blnRecordExist Then MhRealInput22.Value = MhRealInput13.Value
End Sub
Private Sub MhRealInput14_Validate(Cancel As Boolean)   'Adjustment
    Call CalculateTotal("N")    'Adjustment Changed
End Sub
Private Sub ViewRecord()
    ClearFields
    If rstPaperPOList.EOF Then Exit Sub
    FindRecord
    LoadFields
End Sub
Private Sub FindRecord()
    If rstPaperPOParent.State = adStateOpen Then rstPaperPOParent.Close
    rstPaperPOParent.Open "SELECT * FROM PaperPOParent WHERE Code='" & FixQuote(rstPaperPOList.Fields("Code").Value) & "'", CxnPaperPurchaseOrder, adOpenKeyset, adLockOptimistic
    If rstPaperPOParent.RecordCount = 0 Then
       Call DisplayError("This Record has been deleted by Another User ! Click Ok To Refresh the Recordset")
       Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
    End If
End Sub
Private Sub ClearFields()
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    'Text5.Text = ""
    Text8.Text = ""
    'Text9.Text = ""
    MhDateInput1.Text = Format(Date, "dd-MM-yyyy")
    MhDateInput2.Text = "  -  -    "    'Bill Date
    MhDateInput3.Text = Format(DateAdd("d", 1, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")    'Delivery Date
    'MhDateInput4.Text = "  -  -    "    'Delivery Start Date
    'MhDateInput5.Text = "  -  -    "    'Delivery End Date
    'MhDateInput6.Text = "  -  -    "    'Bilty Date
    MhRealInput17.Value = 0 'Total Quantity (Ream) - To be purchased
    MhRealInput18.Value = 0 'Total Quantity (Kg)
    'MhRealInput19.Value = 0 'Total Gross Amount
    'MhRealInput8.Value = 0  'Reams/bundle
    'MhRealInput9.Value = 0  'Total bundles
    'MhRealInput10.Value = 0.6   'Cartage/Kg
    MhRealInput11.Value = 5 'VAT (%)
    MhRealInput12.Value = 0 'VAT
    'MhRealInput13.Value = 0 'Total Cartage
    MhRealInput14.Value = 0 'Adjustment
    MhRealInput15.Value = 0 'Net Amount
    MhRealInput16.Value = 0 'Paid Amount
    MhRealInput20.Value = 0 'Total Quantity (Ream) - To be issued
   ' MhRealInput21.Value = 0 'Total Tat
'    MhRealInput22.Value = 0 'Bilty Amount
    TxtAdNar.Text = ""
    fpSpread1.ClearRange 1, 1, fpSpread1.MaxCols, fpSpread1.MaxRows, True: fpSpread1.SetActiveCell 1, 1
    fpSpread2.ClearRange 1, 1, fpSpread2.MaxCols, fpSpread2.MaxRows, True: fpSpread2.SetActiveCell 1, 1
End Sub
Private Sub LoadFields()
    If rstPaperPOParent.EOF Or rstPaperPOParent.BOF Then Exit Sub
    Text2.Text = rstPaperPOParent.Fields("Name").Value
    MhDateInput1.Text = Format(rstPaperPOParent.Fields("Date").Value, "dd-MM-yyyy")
    MhDateInput3.Text = Format(rstPaperPOParent.Fields("DeliveryDate").Value, "dd-MM-yyyy")
    SupplierCode = rstPaperPOParent.Fields("Supplier").Value
    If rstSupplierList.RecordCount > 0 Then rstSupplierList.MoveFirst
    rstSupplierList.Find "[Code] = '" & SupplierCode & "'"
    If Not rstSupplierList.EOF Then Text3.Text = rstSupplierList.Fields("Col0").Value
    Text4.Text = rstPaperPOParent.Fields("Remarks").Value
    MhRealInput8.Value = Val(rstPaperPOParent.Fields("Reams/Bundle").Value)
    MhRealInput9.Value = Val(rstPaperPOParent.Fields("Bundles").Value)
    MhRealInput10.Value = Val(rstPaperPOParent.Fields("Cartage/Bundle").Value)
    MhRealInput11.Value = Val(rstPaperPOParent.Fields("VAT%").Value)
    MhRealInput12.Value = Val(rstPaperPOParent.Fields("VAT").Value)
    MhRealInput13.Value = Val(rstPaperPOParent.Fields("Cartage").Value)
    MhRealInput14.Value = Val(rstPaperPOParent.Fields("Adjustment").Value)
    MhRealInput15.Value = Val(rstPaperPOParent.Fields("BillAmount").Value)
    Text8.Text = rstPaperPOParent.Fields("BillNo").Value
    Text9.Text = rstPaperPOParent.Fields("ChallanNo").Value
    If Not IsNull(rstPaperPOParent.Fields("BillDate").Value) Then MhDateInput2.Text = Format(rstPaperPOParent.Fields("BillDate").Value, "dd-MM-yyyy")
    If Not IsNull(rstPaperPOParent.Fields("DeliveryStartDate").Value) Then MhDateInput4.Text = Format(rstPaperPOParent.Fields("DeliveryStartDate").Value, "dd-MM-yyyy")
    If Not IsNull(rstPaperPOParent.Fields("DeliveryEndDate").Value) Then MhDateInput5.Text = Format(rstPaperPOParent.Fields("DeliveryEndDate").Value, "dd-MM-yyyy")
    MhRealInput16.Value = Val(rstPaperPOParent.Fields("PaidAmount").Value)
    TxtAdNar.Text = rstPaperPOParent.Fields("AdjustmentRemarks").Value
    Text5.Text = rstPaperPOParent.Fields("BiltyNo").Value
    If Not IsNull(rstPaperPOParent.Fields("BiltyDate").Value) Then MhDateInput6.Text = Format(rstPaperPOParent.Fields("BiltyDate").Value, "dd-MM-yyyy")
    MhRealInput22.Value = Val(rstPaperPOParent.Fields("BiltyAmount").Value)
    Call LoadPaperList(rstPaperPOParent.Fields("Code").Value)
    CalculateTotal ("G")
End Sub
Private Sub EditRecord()
    On Error GoTo ErrorHandler
    If rstPaperPOParent.RecordCount = 0 Then Exit Sub
    If rstPaperPOParent.State = adStateOpen Then rstPaperPOParent.Close
    rstPaperPOParent.CursorLocation = adUseServer
    rstPaperPOParent.Open "SELECT * FROM PaperPOParent WHERE Code='" & FixQuote(rstPaperPOList.Fields("Code").Value) & "'", CxnPaperPurchaseOrder, adOpenKeyset, adLockPessimistic
    MdiMainMenu.MousePointer = vbHourglass
    rstPaperPOParent.Fields("Printstatus") = "N"
    MdiMainMenu.MousePointer = vbNormal
    AddToList
    Call SetButtons(False)
    SSTab1.TabEnabled(0) = False
    Text3.SetFocus
    blnRecordExist = True
    If AllowTransactionsModification = 0 Then
        If Not CheckEmpty(Text8.Text, False) Then LockFields (True)
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
        
        rstPaperPOParent.Fields("Code").Value = GenerateCode(CxnPaperPurchaseOrder, "SELECT MAX(Code) FROM PaperPOParent", 6, "0")
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
    rstPaperPOParent.Fields("Supplier").Value = SupplierCode
    rstPaperPOParent.Fields("DeliveryDate").Value = GetDate(MhDateInput3.Text)
    rstPaperPOParent.Fields("Remarks").Value = Trim(Text4.Text)
    rstPaperPOParent.Fields("Reams/Bundle").Value = Format(Val(MhRealInput8.Text), "0.00")
    rstPaperPOParent.Fields("Bundles").Value = Format(Val(MhRealInput9.Text), "0")
    rstPaperPOParent.Fields("Cartage/Bundle").Value = Format(Val(MhRealInput10.Text), "0.00")
    rstPaperPOParent.Fields("VAT%").Value = Format(Val(MhRealInput11.Text), "0.00")
    rstPaperPOParent.Fields("VAT").Value = Format(Val(MhRealInput12.Text), "0.00")
    rstPaperPOParent.Fields("Cartage").Value = Format(Val(MhRealInput13.Text), "0.00")
    rstPaperPOParent.Fields("Adjustment").Value = Format(Val(MhRealInput14.Text), "0.00")
    rstPaperPOParent.Fields("BillAmount").Value = Format(Val(MhRealInput15.Text), "0.00")
    rstPaperPOParent.Fields("BillNo").Value = Trim(Text8.Text)
    rstPaperPOParent.Fields("ChallanNo").Value = Trim(Text9.Text)
    
    If Not IsDate(MhDateInput2.Text) Then rstPaperPOParent.Fields("BillDate").Value = Null Else rstPaperPOParent.Fields("BillDate").Value = GetDate(MhDateInput2.Text)
    If Not IsDate(MhDateInput4.Text) Then rstPaperPOParent.Fields("DeliveryStartDate").Value = Null Else rstPaperPOParent.Fields("DeliveryStartDate").Value = GetDate(MhDateInput4.Text)
    If Not IsDate(MhDateInput5.Text) Then rstPaperPOParent.Fields("DeliveryEndDate").Value = Null Else rstPaperPOParent.Fields("DeliveryEndDate").Value = GetDate(MhDateInput5.Text)
    rstPaperPOParent.Fields("PaidAmount").Value = Format(Val(MhRealInput16.Text), "0.00")
    rstPaperPOParent.Fields("AdjustmentRemarks").Value = IIf(Val(MhRealInput14.Text) <> 0, TxtAdNar.Text, "")
    rstPaperPOParent.Fields("BiltyNo").Value = Trim(Text5.Text)
    If Not IsDate(MhDateInput6.Text) Then rstPaperPOParent.Fields("BiltyDate").Value = Null Else rstPaperPOParent.Fields("BiltyDate").Value = GetDate(MhDateInput6.Text)
    rstPaperPOParent.Fields("BiltyAmount").Value = Format(Val(MhRealInput22.Text), "0.00")
    If Not CheckEmpty(Text8.Text, False) Then If IsNull(rstPaperPOParent.Fields("BillFeedDate").Value) Then rstPaperPOParent.Fields("BillFeedDate").Value = Now()
    If Not CheckEmpty(Text8.Text, False) Then If IsNull(rstPaperPOParent.Fields("ComputerName").Value) Then rstPaperPOParent.Fields("ComputerName").Value = Left(lpBuff, (InStr(1, lpBuff, vbNullChar)) - 1)
    rstPaperPOParent.Fields("PrintStatus").Value = "N"
    
    
End Sub
Private Sub AddToList()
    On Error Resume Next
    
    rstPaperPOList.MoveFirst
    rstPaperPOList.Find "[Code] = '" & rstPaperPOParent.Fields("Code").Value & "'"
    If rstPaperPOList.EOF Then rstPaperPOList.AddNew
    rstPaperPOList.Fields("Code").Value = rstPaperPOParent.Fields("Code").Value
    rstPaperPOList.Fields("Name").Value = Pad(rstPaperPOParent.Fields("Name").Value, Space(1), 10, "L")
    rstPaperPOList.Fields("Date").Value = rstPaperPOParent.Fields("Date").Value
    rstSupplierList.MoveFirst
    rstSupplierList.Find "[Code] = '" & rstPaperPOParent.Fields("Supplier").Value & "'"
    rstPaperPOList.Fields("SupplierName").Value = Trim(rstSupplierList.Fields("Col0").Value)
    rstPaperPOList.Fields("BillAmount").Value = rstPaperPOParent.Fields("BillAmount").Value
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
    ElseIf Not CheckExists(Text3, "Col0", rstSupplierList, SupplierCode) Then
        Text3.SetFocus
        CheckMandatoryFields = True: Exit Function
    ElseIf CheckDuplicate(CxnPaperPurchaseOrder, "PaperPOParent", "Code", "[Name]+OrderType", Trim(Text2.Text) & OrderType, rstPaperPOParent.Fields("Code").Value, False) Then
        Text2.SetFocus
        CheckMandatoryFields = True: Exit Function
    ElseIf Not ChkPaper() Then
        fpSpread2.SetFocus
        CheckMandatoryFields = True: Exit Function
    End If

    If Val(MhRealInput14.Text) <> 0 Then If CheckEmpty(TxtAdNar.Text, False) Then TxtAdNar.SetFocus: CheckMandatoryFields = True: Exit Function
    If Val(MhRealInput16.Text) <> 0 Then If Val(MhRealInput16.Text) <> Val(MhRealInput15.Text) Then MhRealInput14.SetFocus: CheckMandatoryFields = True: Exit Function: Exit Function
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
Private Sub LoadPaperList(ByVal strOrderCode As String)
    Dim i As Integer
    On Error GoTo ErrorHandler
    If rstPaperPOChild.State = adStateOpen Then rstPaperPOChild.Close
    rstPaperPOChild.Open "SELECT Paper As PaperCode,M.Name As PaperName,QuantityOther,M.[Weight/Ream],QuantityKg,[Rate/Kg],Amount FROM PaperPOChild T INNER JOIN PaperMaster M ON T.Paper=M.Code WHERE T.Code='" & strOrderCode & "' ORDER BY M.Name", CxnPaperPurchaseOrder, adOpenKeyset, adLockOptimistic
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
            .SetText 5, i, Val(rstPaperPOChild.Fields("Rate/Kg").Value)
            .SetText 6, i, Val(rstPaperPOChild.Fields("Amount").Value)
            .SetText 7, i, rstPaperPOChild.Fields("PaperCode").Value
        End With
        rstPaperPOChild.MoveNext
    Loop
    If rstPaperPOChild.State = adStateOpen Then rstPaperPOChild.Close
    rstPaperPOChild.Open "SELECT Paper As PaperCode,M1.Name As PaperName,Account As AccountCode,M2.Name As AccountName,QuantityOther,Tat,Narration FROM (PaperIOChild T INNER JOIN PaperMaster M1 ON T.Paper=M1.Code) INNER JOIN AccountMaster M2 ON T.Account=M2.Code WHERE T.Code='" & strOrderCode & "' ORDER BY M1.Name,M2.Name", CxnPaperPurchaseOrder, adOpenKeyset, adLockOptimistic
    rstPaperPOChild.ActiveConnection = Nothing
    If rstPaperPOChild.RecordCount > 0 Then rstPaperPOChild.MoveFirst
    i = 0
    Do While Not rstPaperPOChild.EOF
        i = i + 1
        With fpSpread2
            .SetText 1, i, rstPaperPOChild.Fields("PaperName").Value
            .SetText 2, i, rstPaperPOChild.Fields("AccountName").Value
            .SetText 3, i, Val(rstPaperPOChild.Fields("QuantityOther").Value)
            .SetText 4, i, Val(rstPaperPOChild.Fields("Tat").Value)
            .SetText 5, i, rstPaperPOChild.Fields("AccountCode").Value
            .SetText 6, i, rstPaperPOChild.Fields("PaperCode").Value
            .SetText 7, i, rstPaperPOChild.Fields("Narration").Value
        End With
        rstPaperPOChild.MoveNext
    Loop
    Exit Sub
ErrorHandler:
    DisplayError ("Failed to Load Paper List")
End Sub
Private Sub CalculateCartage()
    If Val(MhRealInput10.Text) <> 0 Then
        MhRealInput13.Value = Round(Val(MhRealInput18.Text) * Val(MhRealInput10.Text), 0)   'Total Cartage
        If Not blnRecordExist Then MhRealInput22.Value = MhRealInput13.Value
        CalculateTotal ("N")
    End If
End Sub
Private Sub CalculateTotal(ByVal strType As String)
    Dim Qty01 As Variant, Qty02 As Variant, Amt As Variant, i As Integer, Qty As Long
    If strType = "G" Then   'Calculate Cartage & VAT
        MhRealInput17.Value = 0: MhRealInput18.Value = 0: MhRealInput19.Value = 0: MhRealInput20.Value = 0: MhRealInput21.Value = 0
        Qty = 0
        With fpSpread1
            For i = 1 To .DataRowCnt
                .GetText 2, i, Qty01: .GetText 4, i, Qty02: .GetText 6, i, Amt
                Qty = Qty + Int(Val(Qty01)) * 500 + (Val(Qty01) - Int(Val(Qty01))) * 1000
                MhRealInput18.Value = Val(MhRealInput18.Text) + Qty02
                MhRealInput19.Value = Val(MhRealInput19.Text) + Amt
            Next
            MhRealInput17.Value = Int(Qty / 500) + (Qty Mod 500) / 1000
        End With
        Qty = 0
        With fpSpread2
            For i = 1 To .DataRowCnt
                .GetText 3, i, Qty01: .GetText 4, i, Qty02
                Qty = Qty + Int(Val(Qty01)) * 500 + (Val(Qty01) - Int(Val(Qty01))) * 1000
                MhRealInput21.Value = Val(MhRealInput21.Text) + Qty02
            Next
            MhRealInput20.Value = Int(Qty / 500) + (Qty Mod 500) / 1000
        End With
        MhRealInput8_Validate False 'Calculate Total bundles
        MhRealInput12.Value = Val(MhRealInput19.Text) * Val(MhRealInput11.Text) / 100   'VAT
    Else
        MhRealInput15.Value = Round(Val(MhRealInput19.Text) + Val(MhRealInput12.Text) + Val(MhRealInput13.Text) + Val(MhRealInput14.Text), 0)
    End If
End Sub
Private Function GetLastPurchaseRate() As Double
    On Error GoTo ErrorHandler
    If rstLastPurchaseRate.State = adStateOpen Then rstLastPurchaseRate.Close
    rstLastPurchaseRate.Open "SELECT TOP 1 [Rate/Kg] FROM PaperPOParent P INNER JOIN PaperPOChild C ON P.Code=C.Code WHERE Paper='" & PaperCode & "' AND P.Code < '" & IIf(IsNull(rstPaperPOParent.Fields("Code").Value), "999999", rstPaperPOParent.Fields("Code").Value) & "' ORDER BY P.Name DESC", CxnPaperPurchaseOrder, adOpenKeyset, adLockReadOnly
    If rstLastPurchaseRate.RecordCount > 0 Then GetLastPurchaseRate = Val(rstLastPurchaseRate.Fields("Rate/Kg").Value)
    Exit Function
ErrorHandler:
    DisplayError ("Failed to fetch Last Purchase Rate")
End Function
Private Function UpdatePaperList(ByVal ActionType As String) As Boolean
    Dim CellVal(1 To 5) As Variant, Sheets As Long
    On Error GoTo ErrorHandler
    UpdatePaperList = True
    If ActionType = "D" And (Not blnRecordExist) Then Exit Function
    If ActionType = "D" Then
        CxnPaperPurchaseOrder.Execute "DELETE FROM PaperPOChild WHERE Code='" & rstPaperPOParent.Fields("Code").Value & "'"
        CxnPaperPurchaseOrder.Execute "DELETE FROM PaperIOChild WHERE Code='" & rstPaperPOParent.Fields("Code").Value & "'"
    ElseIf ActionType = "I1" Then
        With fpSpread1
            .GetText 2, .ActiveRow, CellVal(1)  'Quantity
            .GetText 4, .ActiveRow, CellVal(2)  'Weight
            .GetText 5, .ActiveRow, CellVal(3)  'Rate
            .GetText 6, .ActiveRow, CellVal(4)  'Amount
            .GetText 7, .ActiveRow, CellVal(5)  'Paper
        End With
        Sheets = Int(Val(CellVal(1))) * 500 + (Val(CellVal(1)) - Int(Val(CellVal(1)))) * 1000
        CxnPaperPurchaseOrder.Execute "INSERT INTO PaperPOChild VALUES ('" & rstPaperPOParent.Fields("Code").Value & "','" & CellVal(5) & "'," & Val(CellVal(1)) & "," & Sheets & "," & Val(CellVal(2)) & "," & Val(CellVal(3)) & "," & Val(CellVal(4)) & ")"
    Else
        With fpSpread2
            .GetText 3, .ActiveRow, CellVal(1)  'Quantity
            .GetText 4, .ActiveRow, CellVal(2)  'Tat
            .GetText 5, .ActiveRow, CellVal(3)  'Account
            .GetText 6, .ActiveRow, CellVal(4)  'Paper
            .GetText 7, .ActiveRow, CellVal(5)  'Narration
        End With
        Sheets = Int(Val(CellVal(1))) * 500 + (Val(CellVal(1)) - Int(Val(CellVal(1)))) * 1000
        CxnPaperPurchaseOrder.Execute "INSERT INTO PaperIOChild VALUES ('" & rstPaperPOParent.Fields("Code").Value & "','" & CellVal(4) & "','" & CellVal(3) & "'," & Val(CellVal(1)) & "," & Sheets & "," & Val(CellVal(2)) & ",'" & CellVal(5) & "')"
    End If
    Exit Function
ErrorHandler:
    UpdatePaperList = False
End Function
Public Sub FilterRecord(ByVal SrchFor As String, ByVal SrchText As String)
    If SrchFor = "Supplier" Then rstPaperPOList.Filter = "[SupplierName] Like '%" & SrchText & "%'"
End Sub
Public Sub PrintPaperPurchaseOrder(ByVal OrderCode As String, ByVal OrderType As String, Optional ByVal Note As String, Optional ByVal OutputType As String, Optional ByVal VchType As Integer)
    
    Dim rstCompanyMaster As New ADODB.Recordset, rstPurchaseOrder As New ADODB.Recordset, rstPurchaseOrderChild As New ADODB.Recordset, Prefix As String
    Dim oOutlookMsg As Outlook.MailItem, FileName As String
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    Prefix = IIf(OrderType = "1", "PB", "PT") & "/" & Right(Year(FinancialYearFrom), 2) + "-" + Right(Year(FinancialYearTo), 2) & "/"
    rstCompanyMaster.Open "SELECT PrintName,Address1,Address2,Address3,Address4,Phone,Fax,eMail FROM CompanyMaster", CxnDatabase, adOpenKeyset, adLockReadOnly
    rstPurchaseOrder.Open "SELECT '" & Prefix & "'+TRIM(P.Name) As OrderNo,[Date] As OrderDate,DeliveryDate,TRIM(M1.PrintName) As SupplierName,[VAT%],VAT,P.Cartage,Adjustment,BillAmount,Remarks,TRIM(M2.PrintName) As PaperName,'',QuantityOther,[Weight/Ream],QuantityKg,[Rate/Kg],(SELECT TOP 1 '" & Prefix & "'+TRIM(P1.Name)+'/'+FORMAT(P1.Date,'dd-MM-yyyy')+'/'+FORMAT([Rate/Kg],'0.00') FROM PaperPOParent P1 INNER JOIN PaperPOChild C1 ON P1.Code=C1.Code WHERE C1.Paper=C.Paper AND P1.Code<P.Code ORDER BY P1.Name DESC) As LastPurchaseRate,Amount,TRIM(eMail) As SupplierMail FROM ((PaperPOParent P LEFT JOIN PaperPOChild C ON P.Code=C.Code) LEFT JOIN AccountMaster M1 ON M1.Code=P.Supplier) LEFT JOIN PaperMaster M2 ON M2.Code=C.Paper WHERE P.Code='" & OrderCode & "' ORDER BY M2.PrintName", CxnDatabase, adOpenKeyset, adLockOptimistic
    rstPurchaseOrderChild.Open "SELECT '" & Prefix & "'+TRIM(P.Name) As OrderNo,[Date] As OrderDate,TRIM(M3.PrintName) As Godown,TRIM(M2.PrintName) As PaperName,TRIM(M1.PrintName) As PrinterName,'' As RefNo,QuantityOther As Quantity,Tat,'' As Remarks,M1.Address1 As PrinterAdd1,M1.Address2 As PrinterAdd2,M1.Address3 As PrinterAdd3,M1.Address4 As PrinterAdd4,TRIM(M1.eMail) As PrinterMail FROM (((PaperPOParent P INNER JOIN PaperIOChild C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON C.Account=M1.Code) INNER JOIN PaperMaster M2 ON C.Paper=M2.Code) INNER JOIN AccountMaster M3 ON P.Supplier=M3.Code WHERE P.Code='" & OrderCode & "' ORDER BY M2.PrintName", CxnDatabase, adOpenKeyset, adLockOptimistic
    Screen.MousePointer = vbNormal
    rstPurchaseOrder.ActiveConnection = Nothing: rstPurchaseOrderChild.ActiveConnection = Nothing
    
    If VchType = 1 Then
        rptPaperPurchaseOrder.Text1.SetText IIf(OrderType = "1", "Book", "Title") & " Paper Purchase Order"
        rptPaperPurchaseOrder.Text2.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
        rptPaperPurchaseOrder.Text3.SetText Trim(rstCompanyMaster.Fields("Address1").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address2").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address3").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address4").Value)
        rptPaperPurchaseOrder.Text24.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & Space(1) & "Fax : " & Trim(rstCompanyMaster.Fields("Fax").Value) & Space(1) & "e-Mail : " & Trim(rstCompanyMaster.Fields("eMail").Value)
        rptPaperPurchaseOrder.Text20.SetText "Add : VAT @" + Format(rstPurchaseOrder.Fields("VAT%").Value, "0.00") + "%"
        rptPaperPurchaseOrder.Text28.SetText " (" & Trim(NumberToWords(rstPurchaseOrder.Fields("BillAmount").Value, True)) & ")"
        rptPaperPurchaseOrder.Text27.SetText "for " & Trim(rstPurchaseOrder.Fields("SupplierName").Value)
        rptPaperPurchaseOrder.Text9.SetText "for " & Trim(rstCompanyMaster.Fields("PrintName").Value)
        
        
    '   **************  By Shamshad Alam **********************************************
        rptPaperPurchaseOrder.Text8.SetText Trim(COMPANY_CIN) 'Add here company cin no
        
        rptPaperPurchaseOrder.Database.SetDataSource rstPurchaseOrder, 3, 1
        rptPaperPurchaseOrder.Subreport1.OpenSubreport.Database.SetDataSource rstPurchaseOrderChild, 3, 1
        If OutputType = "S" Then
            Set FrmReportViewer.Report = rptPaperPurchaseOrder
            FrmReportViewer.Show vbModal
        ElseIf OutputType = "P" Then
            rptPaperPurchaseOrder.PrintOut False    'Print Report Without Prompt
        Else
            Set oOutlookMsg = oOutlook.CreateItem(olMailItem)
            With oOutlookMsg
                .To = rstPurchaseOrder.Fields("SupplierMail").Value
                .Subject = IIf(OrderType = "1", "Book", "Title") & " Paper Purchase Order #" & Trim(rstPurchaseOrder.Fields("OrderNo").Value)
                .HTMLBody = "<Font Face='Calibri' Size='3'>Dear Sir,<Br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Please find attached herewith PO #" & Trim(rstPurchaseOrder.Fields("OrderNo").Value) & " for doing the needful at your end. An early execution of the order will be highly appreciated.<Br>Kindly acknowledge the receipt of mail and confirm the date of execution of order.<Br><Br>" & IIf(Note = "", "", "<b><u>Note : " & Note & "</b></u><Br><Br>") & "Thanks & Regards<Br>Production Department<Br>" & Trim(rstCompanyMaster.Fields("PrintName").Value) & "<Br>Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & "<Br>E-Mail : <a HRef='mailto:" & Trim(rstCompanyMaster.Fields("EMail").Value) & "'>" & Trim(rstCompanyMaster.Fields("EMail").Value) & "</a></Font>"
                rptPaperPurchaseOrder.ExportOptions.FormatType = crEFTPortableDocFormat    ' Set the Export Format As .Pdf
                rptPaperPurchaseOrder.ExportOptions.DestinationType = crEDTDiskFile
                FileName = FixAPIString(GetTemporaryFileName): FileName = Mid(FileName, 1, Len(FileName) - 4) & ".Pdf"
                rptPaperPurchaseOrder.ExportOptions.DiskFileName = FileName
                rptPaperPurchaseOrder.Export False
                .Attachments.Add (FileName)
                .Importance = olImportanceHigh
                .ReadReceiptRequested = True
                If CheckEmpty(.To, False) Then .Display Else .Send
            End With
            Set oOutlookMsg = Nothing
        End If
        Set rptPaperPurchaseOrder = Nothing
    Else
        Dim PrinterMail As String
        Do While Not rstPurchaseOrderChild.EOF
            If Trim(rstPurchaseOrderChild.Fields("PrinterMail").Value) <> "" Then PrinterMail = PrinterMail + IIf(PrinterMail = "", "", ";") & Trim(rstPurchaseOrderChild.Fields("PrinterMail").Value)
            rstPurchaseOrderChild.MoveNext
        Loop
        rstPurchaseOrderChild.MoveFirst
        rptPaperIssueOrder.Text1.SetText IIf(OrderType = "1", "Book", "Title") & " Paper Issue Voucher"
        rptPaperIssueOrder.Text2.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
        rptPaperIssueOrder.Text3.SetText Trim(rstCompanyMaster.Fields("Address1").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address2").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address3").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address4").Value)
        rptPaperIssueOrder.Text24.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & Space(1) & "Fax : " & Trim(rstCompanyMaster.Fields("Fax").Value) & Space(1) & "e-Mail : " & Trim(rstCompanyMaster.Fields("eMail").Value)
        rptPaperIssueOrder.Text9.SetText "for " & Trim(rstCompanyMaster.Fields("PrintName").Value)
        rptPaperIssueOrder.Database.SetDataSource rstPurchaseOrderChild, 3, 1
        If OutputType = "S" Then
            Set FrmReportViewer.Report = rptPaperIssueOrder
            FrmReportViewer.Show vbModal
        ElseIf OutputType = "P" Then
            rptPaperIssueOrder.PrintOut False    'Print Report Without Prompt
        Else
            Set oOutlookMsg = oOutlook.CreateItem(olMailItem)
            With oOutlookMsg
                .To = PrinterMail
                .Subject = IIf(OrderType = "1", "Book", "Title") & " Paper Issue Voucher #" & Trim(rstPurchaseOrderChild.Fields("OrderNo").Value)
                .HTMLBody = "<Font Face='Calibri' Size='3'>Dear Sir,<Br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Please find attached herewith Paper Issue Voucher #" & Trim(rstPurchaseOrderChild.Fields("OrderNo").Value) & " for doing the needful at your end.<Br><b>Kindly acknowledge the receipt of the mail</b>.<Br><Br>" & IIf(Note = "", "", "<b><u>Note : " & Note & "</b></u><Br><Br>") & "Thanks & Regards<Br>Production Department<Br>" & Trim(rstCompanyMaster.Fields("PrintName").Value) & "<Br>Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & "<Br>E-Mail : <a HRef='mailto:" & Trim(rstCompanyMaster.Fields("EMail").Value) & "'>" & Trim(rstCompanyMaster.Fields("EMail").Value) & "</a></Font>"
                rptPaperIssueOrder.ExportOptions.FormatType = crEFTPortableDocFormat    ' Set the Export Format As .Pdf
                rptPaperIssueOrder.ExportOptions.DestinationType = crEDTDiskFile
                FileName = FixAPIString(GetTemporaryFileName): FileName = Mid(FileName, 1, Len(FileName) - 4) & ".Pdf"
                rptPaperIssueOrder.ExportOptions.DiskFileName = FileName
                rptPaperIssueOrder.Export False
                .Attachments.Add (FileName)
                .Importance = olImportanceHigh
                .ReadReceiptRequested = True
                If CheckEmpty(.To, False) Then .Display Else .Send
            End With
            Set oOutlookMsg = Nothing
        End If
        Set rptPaperIssueOrder = Nothing
    End If
    Call CloseRecordset(rstPurchaseOrder): Call CloseRecordset(rstCompanyMaster): Call CloseRecordset(rstPurchaseOrderChild)
    
    On Error GoTo 0
    
End Sub
Private Sub fpSpread1_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask And KeyCode = vbKeyD Then
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            fpSpread1.DeleteRows fpSpread1.ActiveRow, 1: fpSpread1.SetFocus
            CalculateTotal ("G"): CalculateTotal ("N")
        End If
    ElseIf KeyCode = vbKeySpace Then
        Dim Paper As Variant, LastPurchaseRate As Double
        With fpSpread1
            If .ActiveCol = 1 Then
                .GetText .ActiveCol, .ActiveRow, Paper
                Text6.Text = FixQuote(Paper)
                If rstPaperList.RecordCount = 0 Then DisplayError ("No Record in Paper Master"): .SetActiveCell 1, .ActiveRow: Exit Sub Else rstPaperList.MoveFirst
                rstPaperList.Find "[Col0] = '" & RTrim(Paper) & "'"
                SelectionType = "S"
                PaperCode = ""
                Call LoadSelectionList(rstPaperList, "List of Papers...", "Name")
                SearchOrder = 0
                Call DisplaySelectionList(Text6, PaperCode)
                Call CloseForm(FrmSelectionList)
                If PaperCode = "" Then
                    .SetActiveCell 1, .ActiveRow
                Else
                    rstPaperList.MoveFirst: rstPaperList.Find "[Code] ='" & PaperCode & "'"
                    .SetText 1, .ActiveRow, Text6.Text
                    .SetText 3, .ActiveRow, Val(rstPaperList.Fields("Weight/Ream").Value)
                    .SetText 7, .ActiveRow, PaperCode
                    'By Shams
                    'If Not blnRecordExist Then MhRealInput8.Value = Val(rstPaperList.Fields("Reams/Bundle").Value)
                    'End
                    LastPurchaseRate = GetLastPurchaseRate
                    If LastPurchaseRate > 0 Then MsgBox "Last Purchase Rate : Rs." & Format(LastPurchaseRate, "###0.00") & " !!!", vbInformation, App.Title
                    .SetFocus
                    SendKeys "{ENTER}"
                End If
            End If
        End With
    End If
End Sub
Private Sub fpSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim Qty As Variant, Rate As Variant, Wt As Variant, Paper As Variant, GrWt As Double
    With fpSpread1
        If Col = 1 Or Col = 2 Or Col = 5 Then
            .GetText 1, Row, Paper
            .GetText 2, Row, Qty
            .GetText 3, Row, Wt
            .GetText 5, Row, Rate
            GrWt = Fix(Qty) * Wt
            If Qty - Fix(Qty) > 0 Then GrWt = GrWt + ((Qty - Fix(Qty)) * 1000) * (Wt / 500)
            If Paper = "" Then .SetText 4, Row, "": .SetText 6, Row, "" Else .SetText 4, Row, GrWt: .SetText 6, Row, GrWt * Rate: CalculateTotal ("G"): CalculateTotal ("N")
        End If
    End With
End Sub
Private Sub fpSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    EditMode = IIf(Mode = 1, True, False)
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
                    fpSpread1.GetText 7, fpSpread1.ActiveRow, Paper
                    .SetText 6, .ActiveRow, Paper
                    If Paper <> "" Then SendKeys "{ENTER}"
                End If
            ElseIf .ActiveCol = 2 Then
                If Paper <> "" Then
                    .GetText 2, .ActiveRow, Account
                    Text6.Text = FixQuote(Account)
                    If rstAccountList.RecordCount = 0 Then DisplayError ("No Record in Account Master"): .SetActiveCell 1, .ActiveRow: Exit Sub Else rstAccountList.MoveFirst
                    rstAccountList.Find "[Col0] = '" & RTrim(Account) & "'"
                    SelectionType = "S"
                    AccountCode = ""
                    Call LoadSelectionList(rstAccountList, "List of Accounts...", "Name")
                    SearchOrder = 0
                    Call DisplaySelectionList(Text6, AccountCode)
                    Call CloseForm(FrmSelectionList)
                    If AccountCode = "" Then
                        .SetActiveCell 2, .ActiveRow
                    Else
                        rstAccountList.MoveFirst: rstAccountList.Find "[Code] ='" & AccountCode & "'"
                        .SetText 2, .ActiveRow, Text6.Text
                        .SetText 5, .ActiveRow, AccountCode
                        SendKeys "{ENTER}"
                    End If
                End If
            End If
        End With
    End If
End Sub
Private Sub fpSpread2_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim Paper As Variant, Qty As Variant
    With fpSpread2
        If Col = 3 Or Col = 4 Then
            If Col = 3 And OrderType = "1" And MhRealInput8.Value > 0 Then
                .GetText 1, Row, Paper
                .GetText 3, Row, Qty
                If Paper = "" Then .SetText 4, Row, "" Else .SetText 4, Row, Int(Qty / MhRealInput8.Value)
            End If
            CalculateTotal ("G")
        End If
    End With
End Sub
Private Sub fpSpread2_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    EditMode = IIf(Mode = 1, True, False)
End Sub
Private Function ChkPaper() As Boolean
    Dim i As Integer, K As Integer, Paper01 As Variant, Qty01 As Variant, Paper02 As Variant, Qty02 As Variant, Qty As Long, Price As Variant
    ChkPaper = True
    For i = 1 To fpSpread1.DataRowCnt
        fpSpread1.GetText 1, i, Paper01
        fpSpread1.GetText 2, i, Qty01
        fpSpread1.GetText 5, i, Price
        If Val(Price) = 0 Then DisplayError ("Price of Paper at row #" & Trim(Str(i)) & " is zero"): ChkPaper = False: Exit Function
        Qty = 0
        With fpSpread2
            For K = 1 To .DataRowCnt
                .GetText 1, K, Paper02
                If Paper01 = Paper02 Then
                    .GetText 3, K, Qty02
                    Qty = Qty + Int(Val(Qty02)) * 500 + (Val(Qty02) - Int(Val(Qty02))) * 1000
                End If
            Next
        End With
        If Val(Int(Val(Qty01)) * 500 + (Val(Qty01) - Int(Val(Qty01))) * 1000) <> Qty Then DisplayError ("Purchased vs Issued quantity difference for Paper - " & Paper01): ChkPaper = False: Exit Function
    Next
End Function
Private Sub DisplayMenu(ByVal OutputType As String)
    Dim menusel As String
    If rstPaperPOList.RecordCount = 0 Then Exit Sub
    menusel = DisplayPopupMenu(Me.hwnd, 2)
    Select Case menusel
        Case 1
            Call PrintPaperPurchaseOrder(rstPaperPOList.Fields("Code").Value, OrderType, "", OutputType, 1)
        Case 2
            Call PrintPaperPurchaseOrder(rstPaperPOList.Fields("Code").Value, OrderType, "", OutputType, 2)
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
