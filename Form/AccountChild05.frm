VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Begin VB.Form FrmAccountChild05 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Book Printer Rate Detail"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6315
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "AccountChild05.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "FrmLogin"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   6315
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Height          =   375
      Left            =   5845
      Picture         =   "AccountChild05.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Cancel"
      Top             =   465
      Width           =   375
   End
   Begin VB.CommandButton cmdProceed 
      Height          =   375
      Left            =   5845
      Picture         =   "AccountChild05.frx":0544
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Save"
      Top             =   105
      Width           =   375
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
      Height          =   3390
      Left            =   120
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   105
      Width           =   5610
      _Version        =   65536
      _ExtentX        =   9895
      _ExtentY        =   5980
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
      Picture         =   "AccountChild05.frx":0646
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
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   31
         Top             =   100
         Width           =   3690
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
         Left            =   1800
         MaxLength       =   40
         TabIndex        =   0
         Top             =   425
         Width           =   3690
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
         Height          =   330
         Left            =   120
         TabIndex        =   33
         Top             =   420
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
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
         Picture         =   "AccountChild05.frx":0662
         Picture         =   "AccountChild05.frx":067E
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
         Height          =   330
         Left            =   120
         TabIndex        =   34
         Top             =   105
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
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
         Caption         =   " Printer Name"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "AccountChild05.frx":069A
         Picture         =   "AccountChild05.frx":06B6
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
         Height          =   330
         Left            =   120
         TabIndex        =   35
         Top             =   1680
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
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
         Caption         =   " PS Plate Rate"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "AccountChild05.frx":06D2
         Picture         =   "AccountChild05.frx":06EE
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
         Height          =   330
         Left            =   120
         TabIndex        =   36
         Top             =   1995
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
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
         Caption         =   " Depatch Plate Rate"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "AccountChild05.frx":070A
         Picture         =   "AccountChild05.frx":0726
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel7 
         Height          =   330
         Left            =   120
         TabIndex        =   37
         Top             =   2945
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
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
         Caption         =   " Paper Wastage Rate"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "AccountChild05.frx":0742
         Picture         =   "AccountChild05.frx":075E
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel9 
         Height          =   330
         Left            =   120
         TabIndex        =   38
         Top             =   735
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
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
         Caption         =   " Printing Type"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "AccountChild05.frx":077A
         Picture         =   "AccountChild05.frx":0796
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
         Height          =   330
         Left            =   1800
         TabIndex        =   39
         Top             =   735
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
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
         Caption         =   " 1 Color "
         Alignment       =   1
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "AccountChild05.frx":07B2
         Picture         =   "AccountChild05.frx":07CE
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
         Height          =   330
         Left            =   3480
         TabIndex        =   40
         Top             =   735
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
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
         Caption         =   " 4 Color "
         Alignment       =   1
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "AccountChild05.frx":07EA
         Picture         =   "AccountChild05.frx":0806
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
         Height          =   330
         Left            =   2640
         TabIndex        =   41
         Top             =   735
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
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
         Caption         =   " 2 Color "
         Alignment       =   1
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "AccountChild05.frx":0822
         Picture         =   "AccountChild05.frx":083E
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel10 
         Height          =   330
         Left            =   120
         TabIndex        =   42
         Top             =   1365
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
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
         Caption         =   " Print Rate"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "AccountChild05.frx":085A
         Picture         =   "AccountChild05.frx":0876
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
         Height          =   330
         Left            =   4320
         TabIndex        =   43
         Top             =   735
         Width           =   1170
         _Version        =   65536
         _ExtentX        =   2064
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
         Caption         =   " 6 Color "
         Alignment       =   1
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "AccountChild05.frx":0892
         Picture         =   "AccountChild05.frx":08AE
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel12 
         Height          =   330
         Left            =   120
         TabIndex        =   44
         Top             =   1050
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
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
         Caption         =   " Range"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "AccountChild05.frx":08CA
         Picture         =   "AccountChild05.frx":08E6
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput1 
         Height          =   330
         Left            =   1800
         TabIndex        =   1
         Top             =   1050
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   582
         Calculator      =   "AccountChild05.frx":0902
         Caption         =   "AccountChild05.frx":0922
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild05.frx":098E
         Keys            =   "AccountChild05.frx":09AC
         Spin            =   "AccountChild05.frx":09F6
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
         ValueVT         =   1245189
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput2 
         Height          =   330
         Left            =   2640
         TabIndex        =   2
         Top             =   1050
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   582
         Calculator      =   "AccountChild05.frx":0A1E
         Caption         =   "AccountChild05.frx":0A3E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild05.frx":0AAA
         Keys            =   "AccountChild05.frx":0AC8
         Spin            =   "AccountChild05.frx":0B12
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
         ValueVT         =   1245189
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput3 
         Height          =   330
         Left            =   3480
         TabIndex        =   3
         Top             =   1050
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   582
         Calculator      =   "AccountChild05.frx":0B3A
         Caption         =   "AccountChild05.frx":0B5A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild05.frx":0BC6
         Keys            =   "AccountChild05.frx":0BE4
         Spin            =   "AccountChild05.frx":0C2E
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
         ValueVT         =   1245189
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput13 
         Height          =   330
         Left            =   4320
         TabIndex        =   4
         Top             =   1050
         Width           =   1170
         _Version        =   65536
         _ExtentX        =   2064
         _ExtentY        =   582
         Calculator      =   "AccountChild05.frx":0C56
         Caption         =   "AccountChild05.frx":0C76
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild05.frx":0CE2
         Keys            =   "AccountChild05.frx":0D00
         Spin            =   "AccountChild05.frx":0D4A
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
         ValueVT         =   1245189
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput4 
         Height          =   330
         Left            =   1800
         TabIndex        =   5
         Top             =   1365
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   582
         Calculator      =   "AccountChild05.frx":0D72
         Caption         =   "AccountChild05.frx":0D92
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild05.frx":0DFE
         Keys            =   "AccountChild05.frx":0E1C
         Spin            =   "AccountChild05.frx":0E66
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
         ValueVT         =   1245189
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput5 
         Height          =   330
         Left            =   2640
         TabIndex        =   6
         Top             =   1365
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   582
         Calculator      =   "AccountChild05.frx":0E8E
         Caption         =   "AccountChild05.frx":0EAE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild05.frx":0F1A
         Keys            =   "AccountChild05.frx":0F38
         Spin            =   "AccountChild05.frx":0F82
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
         ValueVT         =   1245189
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput6 
         Height          =   330
         Left            =   3480
         TabIndex        =   7
         Top             =   1365
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   582
         Calculator      =   "AccountChild05.frx":0FAA
         Caption         =   "AccountChild05.frx":0FCA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild05.frx":1036
         Keys            =   "AccountChild05.frx":1054
         Spin            =   "AccountChild05.frx":109E
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
         ValueVT         =   1245189
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput14 
         Height          =   330
         Left            =   4320
         TabIndex        =   8
         Top             =   1365
         Width           =   1170
         _Version        =   65536
         _ExtentX        =   2064
         _ExtentY        =   582
         Calculator      =   "AccountChild05.frx":10C6
         Caption         =   "AccountChild05.frx":10E6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild05.frx":1152
         Keys            =   "AccountChild05.frx":1170
         Spin            =   "AccountChild05.frx":11BA
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
         ValueVT         =   1245189
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput7 
         Height          =   330
         Left            =   1800
         TabIndex        =   9
         Top             =   1680
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   582
         Calculator      =   "AccountChild05.frx":11E2
         Caption         =   "AccountChild05.frx":1202
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild05.frx":126E
         Keys            =   "AccountChild05.frx":128C
         Spin            =   "AccountChild05.frx":12D6
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
         ValueVT         =   1245189
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput8 
         Height          =   330
         Left            =   2640
         TabIndex        =   10
         Top             =   1680
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   582
         Calculator      =   "AccountChild05.frx":12FE
         Caption         =   "AccountChild05.frx":131E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild05.frx":138A
         Keys            =   "AccountChild05.frx":13A8
         Spin            =   "AccountChild05.frx":13F2
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
         ValueVT         =   1245189
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput9 
         Height          =   330
         Left            =   3480
         TabIndex        =   11
         Top             =   1680
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   582
         Calculator      =   "AccountChild05.frx":141A
         Caption         =   "AccountChild05.frx":143A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild05.frx":14A6
         Keys            =   "AccountChild05.frx":14C4
         Spin            =   "AccountChild05.frx":150E
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
         ValueVT         =   1245189
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput15 
         Height          =   330
         Left            =   4320
         TabIndex        =   12
         Top             =   1680
         Width           =   1170
         _Version        =   65536
         _ExtentX        =   2064
         _ExtentY        =   582
         Calculator      =   "AccountChild05.frx":1536
         Caption         =   "AccountChild05.frx":1556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild05.frx":15C2
         Keys            =   "AccountChild05.frx":15E0
         Spin            =   "AccountChild05.frx":162A
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
         ValueVT         =   1245189
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput10 
         Height          =   330
         Left            =   1800
         TabIndex        =   13
         Top             =   1995
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   582
         Calculator      =   "AccountChild05.frx":1652
         Caption         =   "AccountChild05.frx":1672
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild05.frx":16DE
         Keys            =   "AccountChild05.frx":16FC
         Spin            =   "AccountChild05.frx":1746
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
         ValueVT         =   1245189
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput11 
         Height          =   330
         Left            =   2640
         TabIndex        =   14
         Top             =   1995
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   582
         Calculator      =   "AccountChild05.frx":176E
         Caption         =   "AccountChild05.frx":178E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild05.frx":17FA
         Keys            =   "AccountChild05.frx":1818
         Spin            =   "AccountChild05.frx":1862
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
         ValueVT         =   1245189
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput12 
         Height          =   330
         Left            =   3480
         TabIndex        =   15
         Top             =   1995
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   582
         Calculator      =   "AccountChild05.frx":188A
         Caption         =   "AccountChild05.frx":18AA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild05.frx":1916
         Keys            =   "AccountChild05.frx":1934
         Spin            =   "AccountChild05.frx":197E
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
         ValueVT         =   1245189
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput16 
         Height          =   330
         Left            =   4320
         TabIndex        =   16
         Top             =   1995
         Width           =   1170
         _Version        =   65536
         _ExtentX        =   2064
         _ExtentY        =   582
         Calculator      =   "AccountChild05.frx":19A6
         Caption         =   "AccountChild05.frx":19C6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild05.frx":1A32
         Keys            =   "AccountChild05.frx":1A50
         Spin            =   "AccountChild05.frx":1A9A
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
         ValueVT         =   1245189
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput17 
         Height          =   330
         Left            =   1800
         TabIndex        =   17
         Top             =   2315
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   582
         Calculator      =   "AccountChild05.frx":1AC2
         Caption         =   "AccountChild05.frx":1AE2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild05.frx":1B4E
         Keys            =   "AccountChild05.frx":1B6C
         Spin            =   "AccountChild05.frx":1BB6
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
         ValueVT         =   1245189
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput18 
         Height          =   330
         Left            =   2640
         TabIndex        =   18
         Top             =   2315
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   582
         Calculator      =   "AccountChild05.frx":1BDE
         Caption         =   "AccountChild05.frx":1BFE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild05.frx":1C6A
         Keys            =   "AccountChild05.frx":1C88
         Spin            =   "AccountChild05.frx":1CD2
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput19 
         Height          =   330
         Left            =   3480
         TabIndex        =   19
         Top             =   2315
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   582
         Calculator      =   "AccountChild05.frx":1CFA
         Caption         =   "AccountChild05.frx":1D1A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild05.frx":1D86
         Keys            =   "AccountChild05.frx":1DA4
         Spin            =   "AccountChild05.frx":1DEE
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput20 
         Height          =   330
         Left            =   4320
         TabIndex        =   20
         Top             =   2315
         Width           =   1170
         _Version        =   65536
         _ExtentX        =   2064
         _ExtentY        =   582
         Calculator      =   "AccountChild05.frx":1E16
         Caption         =   "AccountChild05.frx":1E36
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild05.frx":1EA2
         Keys            =   "AccountChild05.frx":1EC0
         Spin            =   "AccountChild05.frx":1F0A
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput21 
         Height          =   330
         Left            =   1800
         TabIndex        =   25
         Top             =   2945
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   582
         Calculator      =   "AccountChild05.frx":1F32
         Caption         =   "AccountChild05.frx":1F52
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild05.frx":1FBE
         Keys            =   "AccountChild05.frx":1FDC
         Spin            =   "AccountChild05.frx":2026
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "##0.00"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "##0.00"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999.99
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput22 
         Height          =   330
         Left            =   2640
         TabIndex        =   26
         Top             =   2945
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   582
         Calculator      =   "AccountChild05.frx":204E
         Caption         =   "AccountChild05.frx":206E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild05.frx":20DA
         Keys            =   "AccountChild05.frx":20F8
         Spin            =   "AccountChild05.frx":2142
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "##0.00"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "##0.00"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999.99
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput23 
         Height          =   330
         Left            =   3480
         TabIndex        =   27
         Top             =   2945
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   582
         Calculator      =   "AccountChild05.frx":216A
         Caption         =   "AccountChild05.frx":218A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild05.frx":21F6
         Keys            =   "AccountChild05.frx":2214
         Spin            =   "AccountChild05.frx":225E
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "##0.00"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "##0.00"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999.99
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput24 
         Height          =   330
         Left            =   4320
         TabIndex        =   28
         Top             =   2945
         Width           =   1170
         _Version        =   65536
         _ExtentX        =   2064
         _ExtentY        =   582
         Calculator      =   "AccountChild05.frx":2286
         Caption         =   "AccountChild05.frx":22A6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild05.frx":2312
         Keys            =   "AccountChild05.frx":2330
         Spin            =   "AccountChild05.frx":237A
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "##0.00"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "##0.00"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999.99
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
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel13 
         Height          =   330
         Left            =   120
         TabIndex        =   45
         Top             =   2315
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
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
         Caption         =   " Wipeon Plate Rate"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "AccountChild05.frx":23A2
         Picture         =   "AccountChild05.frx":23BE
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel14 
         Height          =   330
         Left            =   120
         TabIndex        =   46
         Top             =   2630
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
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
         Caption         =   " CTP Plate Rate"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "AccountChild05.frx":23DA
         Picture         =   "AccountChild05.frx":23F6
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput25 
         Height          =   330
         Left            =   1800
         TabIndex        =   21
         Top             =   2630
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   582
         Calculator      =   "AccountChild05.frx":2412
         Caption         =   "AccountChild05.frx":2432
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild05.frx":249E
         Keys            =   "AccountChild05.frx":24BC
         Spin            =   "AccountChild05.frx":2506
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
         ValueVT         =   1245189
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput26 
         Height          =   330
         Left            =   2640
         TabIndex        =   22
         Top             =   2630
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   582
         Calculator      =   "AccountChild05.frx":252E
         Caption         =   "AccountChild05.frx":254E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild05.frx":25BA
         Keys            =   "AccountChild05.frx":25D8
         Spin            =   "AccountChild05.frx":2622
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
         ValueVT         =   1245189
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput27 
         Height          =   330
         Left            =   3480
         TabIndex        =   23
         Top             =   2630
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   582
         Calculator      =   "AccountChild05.frx":264A
         Caption         =   "AccountChild05.frx":266A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild05.frx":26D6
         Keys            =   "AccountChild05.frx":26F4
         Spin            =   "AccountChild05.frx":273E
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
         ValueVT         =   1245189
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput28 
         Height          =   330
         Left            =   4320
         TabIndex        =   24
         Top             =   2630
         Width           =   1170
         _Version        =   65536
         _ExtentX        =   2064
         _ExtentY        =   582
         Calculator      =   "AccountChild05.frx":2766
         Caption         =   "AccountChild05.frx":2786
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild05.frx":27F2
         Keys            =   "AccountChild05.frx":2810
         Spin            =   "AccountChild05.frx":285A
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
         ValueVT         =   1245189
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
   End
End
Attribute VB_Name = "FrmAccountChild05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rstAccountChild As New ADODB.Recordset
Public rstSizeList As New ADODB.Recordset
Public AccountName As String
Dim SizeCode As String
Private Sub Form_Load()
    CenterForm Me
    Text2.Text = Trim(AccountName)
    ClearFields
    LoadFields
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyReturn Then
       SendKeys "{TAB}"
       KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyS Then
       cmdProceed_Click
       KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyEscape Then
       cmdCancel_Click
       KeyCode = 0
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        Call CloseForm(FrmAccountChild05)
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set rstAccountChild = Nothing
    Set rstSizeList = Nothing
End Sub
Private Sub ClearFields()
    Text3.Text = ""
    MhRealInput1.Text = "0"
    MhRealInput2.Text = "0"
    MhRealInput3.Text = "0"
    MhRealInput4.Text = "0.00"
    MhRealInput5.Text = "0.00"
    MhRealInput6.Text = "0.00"
    MhRealInput7.Text = "0.00"
    MhRealInput8.Text = "0.00"
    MhRealInput9.Text = "0.00"
    MhRealInput10.Text = "0.00"
    MhRealInput11.Text = "0.00"
    MhRealInput12.Text = "0.00"
    MhRealInput13.Text = "0"
    MhRealInput14.Text = "0.00"
    MhRealInput15.Text = "0.00"
    MhRealInput16.Text = "0.00"
    MhRealInput17.Text = "0.00"
    MhRealInput18.Text = "0.00"
    MhRealInput19.Text = "0.00"
    MhRealInput20.Text = "0.00"
    MhRealInput21.Text = "0.00"
    MhRealInput22.Text = "0.00"
    MhRealInput23.Text = "0.00"
    MhRealInput24.Text = "0.00"
    MhRealInput25.Text = "0.00"
    MhRealInput26.Text = "0.00"
    MhRealInput27.Text = "0.00"
    MhRealInput28.Text = "0.00"
End Sub
Private Sub LoadFields()
    If rstAccountChild.RecordCount = 0 Then Exit Sub
    If Not CheckEmpty(rstAccountChild.Fields("Size").Value, False) Then
        Text3.Text = rstAccountChild.Fields("SizeName").Value
        MhRealInput1.Text = Format(Val(rstAccountChild.Fields("Range1").Value), "0")
        MhRealInput2.Text = Format(Val(rstAccountChild.Fields("Range2").Value), "0")
        MhRealInput3.Text = Format(Val(rstAccountChild.Fields("Range4").Value), "0")
        MhRealInput13.Text = Format(Val(rstAccountChild.Fields("Range6").Value), "0")
        MhRealInput4.Text = Format(Val(rstAccountChild.Fields("PrintRate1").Value), "0.00")
        MhRealInput5.Text = Format(Val(rstAccountChild.Fields("PrintRate2").Value), "0.00")
        MhRealInput6.Text = Format(Val(rstAccountChild.Fields("PrintRate4").Value), "0.00")
        MhRealInput14.Text = Format(Val(rstAccountChild.Fields("PrintRate6").Value), "0.00")
        MhRealInput7.Text = Format(Val(rstAccountChild.Fields("PSPlateRate1").Value), "0.00")
        MhRealInput8.Text = Format(Val(rstAccountChild.Fields("PSPlateRate2").Value), "0.00")
        MhRealInput9.Text = Format(Val(rstAccountChild.Fields("PSPlateRate4").Value), "0.00")
        MhRealInput15.Text = Format(Val(rstAccountChild.Fields("PSPlateRate6").Value), "0.00")
        MhRealInput10.Text = Format(Val(rstAccountChild.Fields("DeepatchPlateRate1").Value), "0.00")
        MhRealInput11.Text = Format(Val(rstAccountChild.Fields("DeepatchPlateRate2").Value), "0.00")
        MhRealInput12.Text = Format(Val(rstAccountChild.Fields("DeepatchPlateRate4").Value), "0.00")
        MhRealInput16.Text = Format(Val(rstAccountChild.Fields("DeepatchPlateRate6").Value), "0.00")
        MhRealInput17.Text = Format(Val(rstAccountChild.Fields("WipeonPlateRate1").Value), "0.00")
        MhRealInput18.Text = Format(Val(rstAccountChild.Fields("WipeonPlateRate2").Value), "0.00")
        MhRealInput19.Text = Format(Val(rstAccountChild.Fields("WipeonPlateRate4").Value), "0.00")
        MhRealInput20.Text = Format(Val(rstAccountChild.Fields("WipeonPlateRate6").Value), "0.00")
        MhRealInput25.Text = Format(Val(rstAccountChild.Fields("CTPPlateRate1").Value), "0.00")
        MhRealInput26.Text = Format(Val(rstAccountChild.Fields("CTPPlateRate2").Value), "0.00")
        MhRealInput27.Text = Format(Val(rstAccountChild.Fields("CTPPlateRate4").Value), "0.00")
        MhRealInput28.Text = Format(Val(rstAccountChild.Fields("CTPPlateRate6").Value), "0.00")
        MhRealInput21.Text = Format(Val(rstAccountChild.Fields("PaperWastageRate1").Value), "0.00")
        MhRealInput22.Text = Format(Val(rstAccountChild.Fields("PaperWastageRate2").Value), "0.00")
        MhRealInput23.Text = Format(Val(rstAccountChild.Fields("PaperWastageRate4").Value), "0.00")
        MhRealInput24.Text = Format(Val(rstAccountChild.Fields("PaperWastageRate6").Value), "0.00")
    End If
End Sub
Private Sub SaveFields()
    rstAccountChild.Fields("Size").Value = SizeCode
    rstAccountChild.Fields("SizeName").Value = Trim(Text3.Text)
    rstAccountChild.Fields("Range1").Value = Val(MhRealInput1.Text)
    rstAccountChild.Fields("Range2").Value = Val(MhRealInput2.Text)
    rstAccountChild.Fields("Range4").Value = Val(MhRealInput3.Text)
    rstAccountChild.Fields("Range6").Value = Val(MhRealInput13.Text)
    rstAccountChild.Fields("PrintRate1").Value = Val(MhRealInput4.Text)
    rstAccountChild.Fields("PrintRate2").Value = Val(MhRealInput5.Text)
    rstAccountChild.Fields("PrintRate4").Value = Val(MhRealInput6.Text)
    rstAccountChild.Fields("PrintRate6").Value = Val(MhRealInput14.Text)
    rstAccountChild.Fields("PSPlateRate1").Value = Val(MhRealInput7.Text)
    rstAccountChild.Fields("PSPlateRate2").Value = Val(MhRealInput8.Text)
    rstAccountChild.Fields("PSPlateRate4").Value = Val(MhRealInput9.Text)
    rstAccountChild.Fields("PSPlateRate6").Value = Val(MhRealInput15.Text)
    rstAccountChild.Fields("DeepatchPlateRate1").Value = Val(MhRealInput10.Text)
    rstAccountChild.Fields("DeepatchPlateRate2").Value = Val(MhRealInput11.Text)
    rstAccountChild.Fields("DeepatchPlateRate4").Value = Val(MhRealInput12.Text)
    rstAccountChild.Fields("DeepatchPlateRate6").Value = Val(MhRealInput16.Text)
    rstAccountChild.Fields("WipeonPlateRate1").Value = Val(MhRealInput17.Text)
    rstAccountChild.Fields("WipeonPlateRate2").Value = Val(MhRealInput18.Text)
    rstAccountChild.Fields("WipeonPlateRate4").Value = Val(MhRealInput19.Text)
    rstAccountChild.Fields("WipeonPlateRate6").Value = Val(MhRealInput20.Text)
    rstAccountChild.Fields("CTPPlateRate1").Value = Val(MhRealInput25.Text)
    rstAccountChild.Fields("CTPPlateRate2").Value = Val(MhRealInput26.Text)
    rstAccountChild.Fields("CTPPlateRate4").Value = Val(MhRealInput27.Text)
    rstAccountChild.Fields("CTPPlateRate6").Value = Val(MhRealInput28.Text)
    rstAccountChild.Fields("PaperWastageRate1").Value = Val(MhRealInput21.Text)
    rstAccountChild.Fields("PaperWastageRate2").Value = Val(MhRealInput22.Text)
    rstAccountChild.Fields("PaperWastageRate4").Value = Val(MhRealInput23.Text)
    rstAccountChild.Fields("PaperWastageRate6").Value = Val(MhRealInput24.Text)
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
        Call DisplaySelectionList(Text3, SizeCode)
        Call CloseForm(FrmSelectionList)
        If CheckEmpty(Text3.Text, False) Then
            Text3.Text = "?"
        End If
        If RTrim(SizeCode) <> "" Then
            SendKeys "{TAB}"
        End If
        Cancel = True
        Exit Sub
    ElseIf (rstAccountChild.Fields("SizeName").Value <> Trim(Text3.Text)) Or (CheckEmpty(rstAccountChild.Fields("SizeName").Value, False)) Then
        If CheckDuplicateEntry Then
            Call DisplayError("Duplicate Entry")
            FocusSelect Me.ActiveControl
            Cancel = True
            Exit Sub
        End If
    End If
    SizeCode = rstSizeList.Fields("Code").Value
End Sub
Private Sub MhRealInput1_Validate(Cancel As Boolean)
    If (Val(CheckNull(rstAccountChild.Fields("Range1").Value)) <> Val(MhRealInput1.Text)) Or (CheckEmpty(rstAccountChild.Fields("SizeName").Value, False)) Then
        If CheckDuplicateEntry Then
            Call DisplayError("Duplicate Entry")
            Me.SetFocus
            Cancel = True
        End If
    End If
End Sub
Private Sub MhRealInput2_Validate(Cancel As Boolean)
    If (Val(CheckNull(rstAccountChild.Fields("Range2").Value)) <> Val(MhRealInput2.Text)) Or (CheckEmpty(rstAccountChild.Fields("SizeName").Value, False)) Then
        If CheckDuplicateEntry Then
            Call DisplayError("Duplicate Entry")
            Me.SetFocus
            Cancel = True
        End If
    End If
End Sub
Private Sub MhRealInput3_Validate(Cancel As Boolean)
    If (Val(CheckNull(rstAccountChild.Fields("Range4").Value)) <> Val(MhRealInput3.Text)) Or (CheckEmpty(rstAccountChild.Fields("SizeName").Value, False)) Then
        If CheckDuplicateEntry Then
            Call DisplayError("Duplicate Entry")
            Me.SetFocus
            Cancel = True
        End If
    End If
End Sub
Private Sub MhRealInput13_Validate(Cancel As Boolean)
    If (Val(CheckNull(rstAccountChild.Fields("Range6").Value)) <> Val(MhRealInput13.Text)) Or (CheckEmpty(rstAccountChild.Fields("SizeName").Value, False)) Then
        If CheckDuplicateEntry Then
            Call DisplayError("Duplicate Entry")
            Me.SetFocus
            Cancel = True
        End If
    End If
End Sub
Private Sub cmdProceed_Click()
    Dim Control As Object
    
    If CheckMandatoryFields Then Exit Sub
    SaveFields
    Me.Tag = "T"
    For Each Control In Me
        If Left(Control.Name, 6) = "MhReal" Then
            If Val(Control.Text) <> 0 Then
                Me.Tag = "F"
            End If
        End If
    Next
    If Me.Tag = "T" Then
        rstAccountChild.Fields("Size").Value = ""
    End If
    rstAccountChild.Update
    Call CloseForm(FrmAccountChild05)
End Sub
Private Sub cmdCancel_Click()
    Call CloseForm(FrmAccountChild05)
End Sub
Private Function CheckMandatoryFields() As Boolean
    If CheckEmpty(Text3.Text, False) Then
        Text3.SetFocus
        CheckMandatoryFields = True
    ElseIf Not CheckExists(Text3, "Col0", rstSizeList, SizeCode) Then
        Text3.SetFocus
        CheckMandatoryFields = True
    ElseIf Val(MhRealInput1.Text) < 0 Or Val(MhRealInput1.Text) > 999999 Then
        MhRealInput1.SetFocus
        FocusSelect Me.ActiveControl
        CheckMandatoryFields = True
    ElseIf Val(MhRealInput2.Text) < 0 Or Val(MhRealInput2.Text) > 999999 Then
        MhRealInput2.SetFocus
        FocusSelect Me.ActiveControl
        CheckMandatoryFields = True
    ElseIf Val(MhRealInput3.Text) < 0 Or Val(MhRealInput3.Text) > 999999 Then
        MhRealInput3.SetFocus
        FocusSelect Me.ActiveControl
        CheckMandatoryFields = True
    ElseIf Val(MhRealInput13.Text) < 0 Or Val(MhRealInput13.Text) > 999999 Then
        MhRealInput13.SetFocus
        FocusSelect Me.ActiveControl
        CheckMandatoryFields = True
    End If
End Function
Private Function CheckDuplicateEntry() As Boolean
    Dim dblBookMark As Double
    
    If rstAccountChild.RecordCount = 0 Then Exit Function
    If Not (rstAccountChild.EOF Or rstAccountChild.BOF) Then
       dblBookMark = rstAccountChild.Bookmark
    End If
    rstAccountChild.MoveFirst
    Do While Not rstAccountChild.EOF
          If rstAccountChild.Fields("SizeName").Value = Trim(Text3.Text) And Val(CheckNull(rstAccountChild.Fields("Range1").Value)) = Val(MhRealInput1.Text) And Val(CheckNull(rstAccountChild.Fields("Range2").Value)) = Val(MhRealInput2.Text) And Val(CheckNull(rstAccountChild.Fields("Range4").Value)) = Val(MhRealInput3.Text) And Val(CheckNull(rstAccountChild.Fields("Range6").Value)) = Val(MhRealInput13.Text) Then
             CheckDuplicateEntry = True
             Exit Do
          End If
          rstAccountChild.MoveNext
    Loop
    If dblBookMark <> 0 Then
       rstAccountChild.Bookmark = dblBookMark
    Else
       rstAccountChild.MoveLast
    End If
End Function
