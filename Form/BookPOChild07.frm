VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Begin VB.Form FrmBookPOChild07 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Title Lamination Order Details"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8280
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "BookPOChild07.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "FrmLogin"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   8280
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Height          =   375
      Left            =   7815
      Picture         =   "BookPOChild07.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Cancel"
      Top             =   465
      Width           =   375
   End
   Begin VB.CommandButton cmdProceed 
      Height          =   375
      Left            =   7815
      Picture         =   "BookPOChild07.frx":0544
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Save"
      Top             =   105
      Width           =   375
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
      Height          =   4410
      Left            =   120
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   105
      Width           =   7575
      _Version        =   65536
      _ExtentX        =   13361
      _ExtentY        =   7779
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
      Picture         =   "BookPOChild07.frx":0646
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
         Left            =   1680
         MaxLength       =   139
         TabIndex        =   14
         Top             =   3600
         Width           =   5775
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
         Left            =   1680
         MaxLength       =   139
         TabIndex        =   13
         Top             =   3290
         Width           =   5775
      End
      Begin VB.TextBox Text5 
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
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   105
         Width           =   1095
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
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   640
         Width           =   3495
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
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   10
         Top             =   2755
         Width           =   1095
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
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   60
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   965
         Width           =   5775
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
         Left            =   1680
         MaxLength       =   40
         TabIndex        =   2
         Top             =   1280
         Width           =   5775
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
         Height          =   330
         Left            =   2760
         TabIndex        =   19
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
         Caption         =   " Order Date"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild07.frx":0662
         Picture         =   "BookPOChild07.frx":067E
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
         Height          =   330
         Left            =   120
         TabIndex        =   20
         Top             =   960
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
         Caption         =   " Title Name"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild07.frx":069A
         Picture         =   "BookPOChild07.frx":06B6
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel9 
         Height          =   330
         Left            =   120
         TabIndex        =   21
         Top             =   1590
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
         Caption         =   " Actual Quantity"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild07.frx":06D2
         Picture         =   "BookPOChild07.frx":06EE
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
         Height          =   330
         Left            =   120
         TabIndex        =   22
         Top             =   1905
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
         Caption         =   " Billing Quantity"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild07.frx":070A
         Picture         =   "BookPOChild07.frx":0726
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
         Height          =   330
         Left            =   2760
         TabIndex        =   23
         Top             =   1905
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
         Caption         =   " Rate"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild07.frx":0742
         Picture         =   "BookPOChild07.frx":075E
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
         Height          =   330
         Left            =   2760
         TabIndex        =   24
         Top             =   2220
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
         Caption         =   " Adjustment"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild07.frx":077A
         Picture         =   "BookPOChild07.frx":0796
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel10 
         Height          =   330
         Left            =   120
         TabIndex        =   25
         Top             =   1275
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
         Caption         =   " Lamination Type"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild07.frx":07B2
         Picture         =   "BookPOChild07.frx":07CE
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
         Height          =   330
         Left            =   2760
         TabIndex        =   26
         Top             =   1590
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
         Caption         =   " Qt.To Binder"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild07.frx":07EA
         Picture         =   "BookPOChild07.frx":0806
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel13 
         Height          =   330
         Left            =   5160
         TabIndex        =   27
         Top             =   1905
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
         Caption         =   " Amount"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild07.frx":0822
         Picture         =   "BookPOChild07.frx":083E
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel15 
         Height          =   330
         Left            =   5160
         TabIndex        =   28
         Top             =   2220
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
         Caption         =   " Total Amount"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild07.frx":085A
         Picture         =   "BookPOChild07.frx":0876
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel19 
         Height          =   330
         Left            =   120
         TabIndex        =   29
         Top             =   2755
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
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
         Picture         =   "BookPOChild07.frx":0892
         Picture         =   "BookPOChild07.frx":08AE
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel20 
         Height          =   330
         Left            =   5160
         TabIndex        =   30
         Top             =   2760
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
         Picture         =   "BookPOChild07.frx":08CA
         Picture         =   "BookPOChild07.frx":08E6
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel23 
         Height          =   330
         Left            =   2760
         TabIndex        =   31
         Top             =   2760
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
         Caption         =   " Bill Date"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild07.frx":0902
         Picture         =   "BookPOChild07.frx":091E
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel24 
         Height          =   330
         Left            =   5160
         TabIndex        =   32
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
         Caption         =   " Target Date"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild07.frx":093A
         Picture         =   "BookPOChild07.frx":0956
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel25 
         Height          =   330
         Left            =   120
         TabIndex        =   34
         Top             =   645
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
         Caption         =   " Laminator Name"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild07.frx":0972
         Picture         =   "BookPOChild07.frx":098E
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel26 
         Height          =   330
         Left            =   5160
         TabIndex        =   35
         Top             =   1590
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
         Caption         =   " Qt.To Off."
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild07.frx":09AA
         Picture         =   "BookPOChild07.frx":09C6
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel27 
         Height          =   330
         Left            =   120
         TabIndex        =   37
         Top             =   105
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
         Caption         =   " Order No."
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild07.frx":09E2
         Picture         =   "BookPOChild07.frx":09FE
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel28 
         Height          =   330
         Left            =   120
         TabIndex        =   38
         Top             =   3290
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
         Caption         =   " Remarks"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild07.frx":0A1A
         Picture         =   "BookPOChild07.frx":0A36
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel7 
         Height          =   330
         Left            =   120
         TabIndex        =   39
         Top             =   2220
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
         Caption         =   " VAT"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild07.frx":0A52
         Picture         =   "BookPOChild07.frx":0A6E
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput2 
         Height          =   330
         Left            =   1680
         TabIndex        =   6
         Top             =   1905
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild07.frx":0A8A
         Caption         =   "BookPOChild07.frx":0AAA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild07.frx":0B16
         Keys            =   "BookPOChild07.frx":0B34
         Spin            =   "BookPOChild07.frx":0B7E
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
         Left            =   1680
         TabIndex        =   3
         Top             =   1590
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild07.frx":0BA6
         Caption         =   "BookPOChild07.frx":0BC6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild07.frx":0C32
         Keys            =   "BookPOChild07.frx":0C50
         Spin            =   "BookPOChild07.frx":0C9A
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
         Left            =   4080
         TabIndex        =   4
         Top             =   1590
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild07.frx":0CC2
         Caption         =   "BookPOChild07.frx":0CE2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild07.frx":0D4E
         Keys            =   "BookPOChild07.frx":0D6C
         Spin            =   "BookPOChild07.frx":0DB6
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
         ValueVT         =   2088828933
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput4 
         Height          =   330
         Left            =   6360
         TabIndex        =   5
         Top             =   1590
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild07.frx":0DDE
         Caption         =   "BookPOChild07.frx":0DFE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild07.frx":0E6A
         Keys            =   "BookPOChild07.frx":0E88
         Spin            =   "BookPOChild07.frx":0ED2
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
         ValueVT         =   2088828933
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput5 
         Height          =   330
         Left            =   4080
         TabIndex        =   7
         Top             =   1905
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild07.frx":0EFA
         Caption         =   "BookPOChild07.frx":0F1A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild07.frx":0F86
         Keys            =   "BookPOChild07.frx":0FA4
         Spin            =   "BookPOChild07.frx":0FEE
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput6 
         Height          =   330
         Left            =   6360
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   1905
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild07.frx":1016
         Caption         =   "BookPOChild07.frx":1036
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild07.frx":10A2
         Keys            =   "BookPOChild07.frx":10C0
         Spin            =   "BookPOChild07.frx":110A
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
         ValueVT         =   2088828933
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput9 
         Height          =   330
         Left            =   4080
         TabIndex        =   9
         Top             =   2220
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild07.frx":1132
         Caption         =   "BookPOChild07.frx":1152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild07.frx":11BE
         Keys            =   "BookPOChild07.frx":11DC
         Spin            =   "BookPOChild07.frx":1226
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
         ValueVT         =   2088828933
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput10 
         Height          =   330
         Left            =   6360
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   2220
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild07.frx":124E
         Caption         =   "BookPOChild07.frx":126E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild07.frx":12DA
         Keys            =   "BookPOChild07.frx":12F8
         Spin            =   "BookPOChild07.frx":1342
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
         ValueVT         =   2088828933
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput17 
         Height          =   330
         Left            =   2160
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   2220
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   582
         Calculator      =   "BookPOChild07.frx":136A
         Caption         =   "BookPOChild07.frx":138A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild07.frx":13F6
         Keys            =   "BookPOChild07.frx":1414
         Spin            =   "BookPOChild07.frx":145E
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
         ValueVT         =   2088828933
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput15 
         Height          =   330
         Left            =   1680
         TabIndex        =   8
         Top             =   2220
         Width           =   495
         _Version        =   65536
         _ExtentX        =   873
         _ExtentY        =   582
         Calculator      =   "BookPOChild07.frx":1486
         Caption         =   "BookPOChild07.frx":14A6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild07.frx":1512
         Keys            =   "BookPOChild07.frx":1530
         Spin            =   "BookPOChild07.frx":157A
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
         ValueVT         =   1975123973
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBDate6Ctl.TDBDate MhDateInput1 
         Height          =   330
         Left            =   4080
         TabIndex        =   0
         Top             =   105
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calendar        =   "BookPOChild07.frx":15A2
         Caption         =   "BookPOChild07.frx":16BA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild07.frx":1726
         Keys            =   "BookPOChild07.frx":1744
         Spin            =   "BookPOChild07.frx":17A2
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
         Left            =   6360
         TabIndex        =   1
         Top             =   105
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calendar        =   "BookPOChild07.frx":17CA
         Caption         =   "BookPOChild07.frx":18E2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild07.frx":194E
         Keys            =   "BookPOChild07.frx":196C
         Spin            =   "BookPOChild07.frx":19CA
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
         Left            =   4080
         TabIndex        =   11
         Top             =   2760
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calendar        =   "BookPOChild07.frx":19F2
         Caption         =   "BookPOChild07.frx":1B0A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild07.frx":1B76
         Keys            =   "BookPOChild07.frx":1B94
         Spin            =   "BookPOChild07.frx":1BF2
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
         Left            =   6360
         TabIndex        =   12
         Top             =   2760
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild07.frx":1C1A
         Caption         =   "BookPOChild07.frx":1C3A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild07.frx":1CA6
         Keys            =   "BookPOChild07.frx":1CC4
         Spin            =   "BookPOChild07.frx":1D0E
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
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel32 
         Height          =   330
         Index           =   0
         Left            =   120
         TabIndex        =   43
         Top             =   3600
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
         Caption         =   " Adj.Remarks"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild07.frx":1D36
         Picture         =   "BookPOChild07.frx":1D52
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel321 
         Height          =   330
         Left            =   5160
         TabIndex        =   44
         Top             =   645
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
         Caption         =   " Extend Date"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild07.frx":1D6E
         Picture         =   "BookPOChild07.frx":1D8A
      End
      Begin TDBDate6Ctl.TDBDate MhDateInput31 
         Height          =   330
         Left            =   6360
         TabIndex        =   45
         Top             =   640
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calendar        =   "BookPOChild07.frx":1DA6
         Caption         =   "BookPOChild07.frx":1EBE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild07.frx":1F2A
         Keys            =   "BookPOChild07.frx":1F48
         Spin            =   "BookPOChild07.frx":1FA6
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
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel32 
         Height          =   330
         Index           =   1
         Left            =   120
         TabIndex        =   46
         Top             =   3920
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
         Caption         =   " Created On"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild07.frx":1FCE
         Picture         =   "BookPOChild07.frx":1FEA
      End
      Begin TDBDate6Ctl.TDBDate MhDateInput311 
         Height          =   330
         Left            =   1680
         TabIndex        =   47
         Top             =   3915
         Width           =   5775
         _Version        =   65536
         _ExtentX        =   10186
         _ExtentY        =   582
         Calendar        =   "BookPOChild07.frx":2006
         Caption         =   "BookPOChild07.frx":211E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild07.frx":218A
         Keys            =   "BookPOChild07.frx":21A8
         Spin            =   "BookPOChild07.frx":2206
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
      Begin VB.Line Line4 
         X1              =   0
         X2              =   8300
         Y1              =   3185
         Y2              =   3185
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   8300
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   8300
         Y1              =   2645
         Y2              =   2645
      End
   End
End
Attribute VB_Name = "FrmBookPOChild07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public rstBookPOChild07 As New ADODB.Recordset
Public LaminatorCode As String
Public TitlePrinterQuantity As Long
Dim LaminationTypeCode As String
Dim FormType As String
Dim SizeCode As String
Dim rstLaminationTypeList As New ADODB.Recordset
Dim rstLaminatorRates As New ADODB.Recordset
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    
    CenterForm Me
    BusySystemIndicator True
    DisableCloseButton Me
    FormType = FrmBookPrintOrder.rstBookList.Fields("FormType").Value
    SizeCode = FrmBookPrintOrder.rstBookList.Fields("SizeCode").Value
    Text5.Text = Trim(FrmBookPrintOrder.Text2.Text)
    Text4.Text = Trim(FrmBookPrintOrder.Text7.Text)
    Text2.Text = Trim(FrmBookPrintOrder.Text3.Text)
    ClearFields
    rstLaminationTypeList.Open "Select Name As Col0, Code From GeneralMaster Where Type = '7' Order By Name", CxnDatabase, adOpenKeyset, adLockReadOnly
    rstLaminationTypeList.ActiveConnection = Nothing
    If Val(CheckNull(rstBookPOChild07.Fields("ActualQuantity").Value)) = 0 Then
        LaminationTypeCode = FrmBookPrintOrder.rstBookList.Fields("LaminationType").Value
        If rstLaminationTypeList.RecordCount > 0 Then rstLaminationTypeList.MoveFirst
        rstLaminationTypeList.Find "[Code] = '" & LaminationTypeCode & "'"
        If Not rstLaminationTypeList.EOF Then
           Text3.Text = rstLaminationTypeList.Fields("Col0").Value
        End If
        MhDateInput1.Text = Format(GetDate(FrmBookPrintOrder.MhDateInput1.Text), "dd-MM-yyyy")
        MhDateInput3.Text = Format(DateAdd("d", 5, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
        MhRealInput1.Text = Format(TitlePrinterQuantity, "0")
        If rstLaminationTypeList.RecordCount > 0 Then
            rstLaminationTypeList.MoveFirst
            rstLaminationTypeList.Find "[Code] = '" & LaminationTypeCode & "'"
            If Not rstLaminationTypeList.EOF Then
                Text3.Text = rstLaminationTypeList.Fields("Col0").Value
            End If
        End If
    Else
        LoadFields
    End If
    BusySystemIndicator False
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    CloseForm Me
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
        CloseForm Me
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CloseRecordset(rstLaminationTypeList)
    Call CloseRecordset(rstLaminatorRates)
End Sub
Private Sub ClearFields()
    MhDateInput1.Text = Format(Date, "dd-MM-yyyy")
    MhDateInput2.Text = "  -  -    "
    MhDateInput311.Text = Format(Date, "dd-MM-yyyy")
    MhDateInput3.Text = Format(DateAdd("d", 15, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
    Text3.Text = ""
    Text6.Text = ""
    Text8.Text = ""
    MhRealInput1.Text = "0"
    MhRealInput2.Text = "0"
    MhRealInput3.Text = "0"
    MhRealInput4.Text = "0"
    MhRealInput5.Text = "0.00"
    MhRealInput6.Text = "0.00"
    MhRealInput9.Text = "0.00"
    MhRealInput10.Text = "0.00"
    MhRealInput15.Text = "0.00"
    MhRealInput16.Text = "0.00"
    MhRealInput17.Text = "0.00"
    TxtAdNar.Text = ""
End Sub
Private Sub LoadFields()
    If rstBookPOChild07.RecordCount = 0 Then Exit Sub
    MhDateInput1.Text = Format(rstBookPOChild07.Fields("OrderDate").Value, "dd-MM-yyyy")
    MhDateInput3.Text = Format(rstBookPOChild07.Fields("TargetDate").Value, "dd-MM-yyyy")
    
    
    If Not IsNull(rstBookPOChild07.Fields("ExtendDate").Value) Then
        MhDateInput31.Text = Format(rstBookPOChild07.Fields("ExtendDate").Value, "dd-MM-yyyy")
    Else
        MhDateInput31.Text = "  -  -    "
    End If
 
    If Not IsNull(rstBookPOChild07.Fields("CreatedOn").Value) Then
        MhDateInput311.Text = Format(rstBookPOChild07.Fields("CreatedOn").Value, "dd-MM-yyyy")
    End If
    
    
    MhRealInput1.Text = Format(Val(rstBookPOChild07.Fields("ActualQuantity").Value), "0")
    MhRealInput2.Text = Format(Val(rstBookPOChild07.Fields("BillingQuantity").Value), "0")
    MhRealInput3.Text = Format(Val(rstBookPOChild07.Fields("QuantityToBinder").Value), "0")
    MhRealInput4.Text = Format(Val(rstBookPOChild07.Fields("QuantityToOffice").Value), "0")
    MhRealInput5.Text = Format(Val(rstBookPOChild07.Fields("LaminationRate").Value), "0.00")
    MhRealInput6.Text = Format(Val(rstBookPOChild07.Fields("LaminationAmount").Value), "0.00")
    MhRealInput9.Text = Format(Val(rstBookPOChild07.Fields("Adjustment").Value), "0.00")
    MhRealInput10.Text = Format(Val(rstBookPOChild07.Fields("BillAmount").Value), "0.00")
    LaminationTypeCode = rstBookPOChild07.Fields("LaminationType").Value
    If rstLaminationTypeList.RecordCount > 0 Then rstLaminationTypeList.MoveFirst
    rstLaminationTypeList.Find "[Code] = '" & LaminationTypeCode & "'"
    If Not rstLaminationTypeList.EOF Then
       Text3.Text = rstLaminationTypeList.Fields("Col0").Value
    End If
    Text8.Text = rstBookPOChild07.Fields("BillNo").Value
    If Not IsNull(rstBookPOChild07.Fields("BillDate").Value) Then
        MhDateInput2.Text = Format(rstBookPOChild07.Fields("BillDate").Value, "dd-MM-yyyy")
    End If
    MhRealInput15.Text = Format(Val(rstBookPOChild07.Fields("VAT%").Value), "0.00")
    MhRealInput17.Text = Format(Val(rstBookPOChild07.Fields("VAT").Value), "0.00")
    MhRealInput16.Text = Format(Val(rstBookPOChild07.Fields("PaidAmount").Value), "0.00")
    Text6.Text = rstBookPOChild07.Fields("Remarks").Value
    TxtAdNar.Text = rstBookPOChild07.Fields("AdjustmentRemarks").Value
End Sub
Private Sub SaveFields()
    rstBookPOChild07.Fields("OrderDate").Value = GetDate(MhDateInput1.Text)
    rstBookPOChild07.Fields("TargetDate").Value = GetDate(MhDateInput3.Text)
    
    If Not IsDate(MhDateInput31.Text) Then
         rstBookPOChild07.Fields("ExtendDate").Value = Null
    Else
         rstBookPOChild07.Fields("ExtendDate").Value = GetDate(MhDateInput31.Text)
    End If
    
        
    If Not IsDate(MhDateInput311.Text) Then
         rstBookPOChild07.Fields("CreatedOn").Value = Null
    Else
         rstBookPOChild07.Fields("CreatedOn").Value = GetDate(MhDateInput311.Text)
    End If
    
    rstBookPOChild07.Fields("LaminationType").Value = LaminationTypeCode
    rstBookPOChild07.Fields("ActualQuantity").Value = Val(MhRealInput1.Text)
    rstBookPOChild07.Fields("BillingQuantity").Value = Val(MhRealInput2.Text)
    rstBookPOChild07.Fields("QuantityToBinder").Value = Val(MhRealInput3.Text)
    rstBookPOChild07.Fields("QuantityToOffice").Value = Val(MhRealInput4.Text)
    rstBookPOChild07.Fields("LaminationRate").Value = Val(MhRealInput5.Text)
    rstBookPOChild07.Fields("LaminationAmount").Value = Val(MhRealInput6.Text)
    rstBookPOChild07.Fields("Adjustment").Value = Val(MhRealInput9.Text)
    rstBookPOChild07.Fields("BillAmount").Value = Val(MhRealInput10.Text)
    rstBookPOChild07.Fields("BillNo").Value = Text8.Text
    If Not IsDate(MhDateInput2.Text) Then
         rstBookPOChild07.Fields("BillDate").Value = Null
    Else
         rstBookPOChild07.Fields("BillDate").Value = GetDate(MhDateInput2.Text)
    End If
    rstBookPOChild07.Fields("VAT%").Value = Val(MhRealInput15.Text)
    rstBookPOChild07.Fields("VAT").Value = Val(MhRealInput17.Text)
    rstBookPOChild07.Fields("PaidAmount").Value = Val(MhRealInput16.Text)
    rstBookPOChild07.Fields("Remarks").Value = Text6.Text
    rstBookPOChild07.Fields("AdjustmentRemarks").Value = IIf(Val(MhRealInput9.Text) <> 0, TxtAdNar.Text, "")
    If Not CheckEmpty(Text8.Text, False) Then If IsNull(rstBookPOChild07.Fields("BillFeedDate").Value) Then rstBookPOChild07.Fields("BillFeedDate").Value = Now()
    Dim lpBuff As String * 1024
    GetComputerName lpBuff, Len(lpBuff)
    If Not CheckEmpty(Text8.Text, False) Then If IsNull(rstBookPOChild07.Fields("ComputerName").Value) Then rstBookPOChild07.Fields("ComputerName").Value = Left(lpBuff, (InStr(1, lpBuff, vbNullChar)) - 1)
End Sub
Private Sub MhDateInput1_Validate(Cancel As Boolean)
    If Not IsDate(GetDate(MhDateInput1.Text)) Then
        Cancel = True
    ElseIf Format(GetDate(MhDateInput1.Text), "yyyymmdd") < Format(FinancialYearFrom, "yyyymmdd") Or Format(GetDate(MhDateInput1.Text), "yyyymmdd") > Format(FinancialYearTo, "yyyymmdd") Then
        Cancel = True
    ElseIf Val(CheckNull(rstBookPOChild07.Fields("ActualQuantity").Value)) = 0 Then
        'MhDateInput3.Text = Format(DateAdd("d", 2, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
        MhDateInput3.Text = Format(DateAdd("d", 15, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
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
Private Sub MhDateInput3_Validate(Cancel As Boolean)
    If Not IsDate(GetDate(MhDateInput3.Text)) Then
        Cancel = True
    ElseIf Format(GetDate(MhDateInput3.Text), "yyyymmdd") <= Format(GetDate(MhDateInput1.Text), "yyyymmdd") Then
        DisplayError ("Target Date cann't be prior to Order Date")
        MhDateInput3.SetFocus
        Cancel = True
    End If
End Sub
Private Sub MhRealInput1_Validate(Cancel As Boolean)
    If Val(MhRealInput2.Text) = 0 Then
        MhRealInput2.Text = Format(Val(MhRealInput1.Text), "0")
    End If
    CalculateAmount
    MhRealInput3.Text = Format(Val(MhRealInput1.Text) - Val(MhRealInput4.Text), "0")
End Sub
Private Sub MhRealInput3_Validate(Cancel As Boolean)
    MhRealInput4.Text = Format(Val(MhRealInput1.Text) - Val(MhRealInput3.Text), "0")
End Sub
Private Sub MhRealInput2_Validate(Cancel As Boolean)
    If Val(MhRealInput2.Text) > Val(MhRealInput1.Text) Then
        If MsgBox("Billing Quantity is greater than Actual Quantity !" & vbCrLf & "         Would you like to continue ?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Proceed !") = vbNo Then
            MhRealInput2.SetFocus
            Cancel = True
            Exit Sub
        End If
    End If
    CalculateAmount
End Sub
Private Sub MhRealInput5_Validate(Cancel As Boolean)
    CalculateAmount
End Sub
Private Sub MhRealInput15_Validate(Cancel As Boolean)
    CalculateTotalAmount
End Sub
Private Sub MhRealInput9_Validate(Cancel As Boolean)
    CalculateTotalAmount
End Sub
Private Sub CalculateAmount()   'Calculate Amount
    MhRealInput6.Text = Format(Val(MhRealInput5.Text) * Val(MhRealInput2.Text), "0.00")
    CalculateTotalAmount
End Sub
Private Sub CalculateTotalAmount()  'Calculate Total Amount
    'MhRealInput17.Text = Format((Val(MhRealInput6.Text) * 80 / 100) * Val(MhRealInput15.Text) / 100, "0.00")
    
    MhRealInput17.Text = Format((Val(MhRealInput6.Text) * 100 / 100) * Val(MhRealInput15.Text) / 100, "0.00")
    MhRealInput10.Text = Format(Val(MhRealInput6.Text) + Val(MhRealInput17.Text) + Val(MhRealInput9.Text), "0.00")
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
    If rstLaminationTypeList.RecordCount = 0 Then
        DisplayError ("No Record in Lamination Type Master")
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
        Call DisplaySelectionList(Text3, LaminationTypeCode)
        Call CloseForm(FrmSelectionList)
        If CheckEmpty(Text3.Text, False) Then
            Text3.Text = "?"
        End If
        If RTrim(LaminationTypeCode) <> "" Then
            SendKeys "{TAB}"
        End If
        Cancel = True
    Else
        LaminationTypeCode = rstLaminationTypeList.Fields("Code").Value
        Call GetLaminatorRates: CalculateAmount
    End If
End Sub
Private Sub GetLaminatorRates()
    Dim LaminationRate As Double
    
    On Error GoTo ErrorHandler
    If rstLaminatorRates.State = adStateOpen Then rstLaminatorRates.Close
    rstLaminatorRates.Open "Select * From AccountChild07 Where Code = '" & LaminatorCode & "' And [Size] = '" & SizeCode & "' And LaminationType = '" & LaminationTypeCode & "'", CxnDatabase, adOpenKeyset, adLockReadOnly
    If rstLaminatorRates.RecordCount = 0 Then
        If rstLaminatorRates.State = adStateOpen Then rstLaminatorRates.Close
        rstLaminatorRates.Open "Select * From AccountMaster,AccountChild07 Where AccountMaster.Code = AccountChild07.Code And [Name] Like '%Rate%' And [Size] = '" & SizeCode & "' And LaminationType = '" & LaminationTypeCode & "'", CxnDatabase, adOpenKeyset, adLockReadOnly
    End If
    If rstLaminatorRates.RecordCount > 0 Then
        LaminationRate = rstLaminatorRates.Fields("Rate" & IIf(FormType = "1", "08", IIf(FormType = "2", "16", IIf(FormType = "3", "04", IIf(FormType = "4", "12", IIf(FormType = "5", "24", IIf(FormType = "6", "32", "64"))))))).Value
    End If
    If Val(MhRealInput5.Text) <> LaminationRate And Val(MhRealInput5.Text) <> 0 Then
        If MsgBox("Lamination Rate is different from that in Master ! Change Rate?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then
            MhRealInput5.Text = Format(LaminationRate, "0.00")
        End If
    Else
        MhRealInput5.Text = Format(LaminationRate, "0.00")
    End If
    Exit Sub
ErrorHandler:
    DisplayError ("Failed to Fetch Laminator Rates")
End Sub
Private Function CheckMandatoryFields() As Boolean
    If CheckEmpty(Text3.Text, False) Then Text3.SetFocus: CheckMandatoryFields = True: Exit Function
    If Not CheckExists(Text3, "Col0", rstLaminationTypeList, LaminationTypeCode) Then Text3.SetFocus: CheckMandatoryFields = True: Exit Function
    If Val(MhRealInput16.Text) <> 0 Then If Val(MhRealInput16.Text) <> Val(MhRealInput10.Text) Then MhRealInput9.SetFocus: CheckMandatoryFields = True: Exit Function
    If Val(MhRealInput9.Text) <> 0 Then If CheckEmpty(TxtAdNar.Text, False) Then TxtAdNar.SetFocus: CheckMandatoryFields = True: Exit Function
End Function
Private Sub cmdProceed_Click()
    If CheckMandatoryFields Then Exit Sub
    SaveFields
    rstBookPOChild07.Update
    Call CloseForm(Me)
End Sub
Private Sub cmdCancel_Click()
    rstBookPOChild07.CancelUpdate
    Call CloseForm(Me)
End Sub
