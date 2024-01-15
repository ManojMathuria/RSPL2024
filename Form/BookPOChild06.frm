VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmBookPOChild06 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Title Printing Order Details"
   ClientHeight    =   6510
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
   Icon            =   "BookPOChild06.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "FrmLogin"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   8280
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Height          =   375
      Left            =   7815
      Picture         =   "BookPOChild06.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Cancel"
      Top             =   465
      Width           =   375
   End
   Begin VB.CommandButton cmdProceed 
      Height          =   375
      Left            =   7815
      Picture         =   "BookPOChild06.frx":0544
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Save"
      Top             =   105
      Width           =   375
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
      Height          =   6195
      Left            =   120
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   105
      Width           =   7575
      _Version        =   65536
      _ExtentX        =   13361
      _ExtentY        =   10927
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
      Picture         =   "BookPOChild06.frx":0646
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
         Left            =   1320
         MaxLength       =   139
         TabIndex        =   24
         Top             =   5130
         Width           =   6140
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
         Left            =   1320
         MaxLength       =   139
         TabIndex        =   23
         Top             =   4815
         Width           =   6140
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
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   55
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
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   960
         Width           =   3735
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
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   20
         Top             =   4290
         Width           =   1095
      End
      Begin VB.TextBox Text1 
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
         TabIndex        =   17
         Top             =   3435
         Width           =   6140
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
         Locked          =   -1  'True
         MaxLength       =   60
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   640
         Width           =   3735
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
         TabIndex        =   3
         Top             =   1280
         Width           =   1095
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
         Height          =   330
         Left            =   2400
         TabIndex        =   29
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
         Caption         =   " Order Date"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild06.frx":0662
         Picture         =   "BookPOChild06.frx":067E
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
         Height          =   330
         Left            =   120
         TabIndex        =   30
         Top             =   960
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
         Caption         =   " Printer Name"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild06.frx":069A
         Picture         =   "BookPOChild06.frx":06B6
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
         Height          =   330
         Left            =   2400
         TabIndex        =   31
         Top             =   1280
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
         Caption         =   " Actual Quantity"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild06.frx":06D2
         Picture         =   "BookPOChild06.frx":06EE
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
         Height          =   330
         Left            =   5040
         TabIndex        =   32
         Top             =   1275
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
         Caption         =   " Billing Quantity"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild06.frx":070A
         Picture         =   "BookPOChild06.frx":0726
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel9 
         Height          =   330
         Left            =   120
         TabIndex        =   33
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
         Caption         =   " Printing Type"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild06.frx":0742
         Picture         =   "BookPOChild06.frx":075E
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
         Height          =   330
         Left            =   120
         TabIndex        =   34
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
         Caption         =   " Total Plates"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild06.frx":077A
         Picture         =   "BookPOChild06.frx":0796
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
         Height          =   330
         Left            =   2400
         TabIndex        =   35
         Top             =   2225
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
         Caption         =   " Plate Rate"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild06.frx":07B2
         Picture         =   "BookPOChild06.frx":07CE
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel7 
         Height          =   330
         Left            =   2400
         TabIndex        =   36
         Top             =   2540
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
         Caption         =   " Print Rate"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild06.frx":07EA
         Picture         =   "BookPOChild06.frx":0806
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
         Height          =   330
         Left            =   2400
         TabIndex        =   37
         Top             =   2855
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
         Caption         =   " Adjmnt"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild06.frx":0822
         Picture         =   "BookPOChild06.frx":083E
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel10 
         Height          =   330
         Left            =   120
         TabIndex        =   38
         Top             =   1275
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
         Caption         =   " Ref.No."
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild06.frx":085A
         Picture         =   "BookPOChild06.frx":0876
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
         Height          =   330
         Index           =   0
         Left            =   2400
         TabIndex        =   39
         Top             =   1590
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
         Caption         =   " Plate Type"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild06.frx":0892
         Picture         =   "BookPOChild06.frx":08AE
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel12 
         Height          =   330
         Left            =   120
         TabIndex        =   40
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
         Caption         =   " Total Forms"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild06.frx":08CA
         Picture         =   "BookPOChild06.frx":08E6
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel13 
         Height          =   330
         Left            =   5040
         TabIndex        =   41
         Top             =   2225
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
         Caption         =   " Plate Amount"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild06.frx":0902
         Picture         =   "BookPOChild06.frx":091E
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel14 
         Height          =   330
         Left            =   5040
         TabIndex        =   42
         Top             =   2540
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
         Caption         =   " Print Amount"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild06.frx":093A
         Picture         =   "BookPOChild06.frx":0956
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel15 
         Height          =   330
         Left            =   5040
         TabIndex        =   43
         Top             =   2855
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
         Caption         =   " Total Amount"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild06.frx":0972
         Picture         =   "BookPOChild06.frx":098E
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel16 
         Height          =   330
         Left            =   120
         TabIndex        =   44
         Top             =   3435
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
         Caption         =   " Paper Name"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild06.frx":09AA
         Picture         =   "BookPOChild06.frx":09C6
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel17 
         Height          =   330
         Left            =   120
         TabIndex        =   45
         Top             =   3750
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
         Caption         =   " Wastage (%)"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild06.frx":09E2
         Picture         =   "BookPOChild06.frx":09FE
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel18 
         Height          =   330
         Left            =   2400
         TabIndex        =   46
         Top             =   3750
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
         Caption         =   " Ups/Sheet"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild06.frx":0A1A
         Picture         =   "BookPOChild06.frx":0A36
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel21 
         Height          =   330
         Left            =   5040
         TabIndex        =   47
         Top             =   3750
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
         Caption         =   " Consumption"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild06.frx":0A52
         Picture         =   "BookPOChild06.frx":0A6E
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel19 
         Height          =   330
         Left            =   120
         TabIndex        =   48
         Top             =   4290
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
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
         Picture         =   "BookPOChild06.frx":0A8A
         Picture         =   "BookPOChild06.frx":0AA6
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel20 
         Height          =   330
         Left            =   5040
         TabIndex        =   49
         Top             =   4290
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
         Caption         =   " Paid Amount"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild06.frx":0AC2
         Picture         =   "BookPOChild06.frx":0ADE
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel23 
         Height          =   330
         Left            =   2400
         TabIndex        =   50
         Top             =   4290
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
         Caption         =   " Bill Date"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild06.frx":0AFA
         Picture         =   "BookPOChild06.frx":0B16
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel24 
         Height          =   330
         Left            =   5040
         TabIndex        =   51
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
         Caption         =   " Target Date"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild06.frx":0B32
         Picture         =   "BookPOChild06.frx":0B4E
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel25 
         Height          =   330
         Left            =   120
         TabIndex        =   53
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
         Caption         =   " Book Name"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild06.frx":0B6A
         Picture         =   "BookPOChild06.frx":0B86
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel26 
         Height          =   330
         Left            =   5040
         TabIndex        =   54
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
         Caption         =   " Ups/Plate"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild06.frx":0BA2
         Picture         =   "BookPOChild06.frx":0BBE
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel27 
         Height          =   330
         Left            =   120
         TabIndex        =   56
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
         Caption         =   " Order No."
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild06.frx":0BDA
         Picture         =   "BookPOChild06.frx":0BF6
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel28 
         Height          =   330
         Left            =   120
         TabIndex        =   57
         Top             =   4815
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
         Caption         =   " Remarks"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild06.frx":0C12
         Picture         =   "BookPOChild06.frx":0C2E
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel22 
         Height          =   330
         Left            =   120
         TabIndex        =   58
         Top             =   2535
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
         Picture         =   "BookPOChild06.frx":0C4A
         Picture         =   "BookPOChild06.frx":0C66
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput6 
         Height          =   330
         Left            =   1320
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   2220
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild06.frx":0C82
         Caption         =   "BookPOChild06.frx":0CA2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild06.frx":0D0E
         Keys            =   "BookPOChild06.frx":0D2C
         Spin            =   "BookPOChild06.frx":0D76
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
         ValueVT         =   1972502533
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput5 
         Height          =   330
         Left            =   3840
         TabIndex        =   14
         Top             =   2540
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   582
         Calculator      =   "BookPOChild06.frx":0D9E
         Caption         =   "BookPOChild06.frx":0DBE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild06.frx":0E2A
         Keys            =   "BookPOChild06.frx":0E48
         Spin            =   "BookPOChild06.frx":0E92
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
         ValueVT         =   1973878789
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput4 
         Height          =   330
         Left            =   3840
         TabIndex        =   11
         Top             =   2225
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   582
         Calculator      =   "BookPOChild06.frx":0EBA
         Caption         =   "BookPOChild06.frx":0EDA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild06.frx":0F46
         Keys            =   "BookPOChild06.frx":0F64
         Spin            =   "BookPOChild06.frx":0FAE
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
         ValueVT         =   1973878789
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput9 
         Height          =   330
         Left            =   3840
         TabIndex        =   16
         Top             =   2855
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   582
         Calculator      =   "BookPOChild06.frx":0FD6
         Caption         =   "BookPOChild06.frx":0FF6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild06.frx":1062
         Keys            =   "BookPOChild06.frx":1080
         Spin            =   "BookPOChild06.frx":10CA
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
         ValueVT         =   1972633605
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput7 
         Height          =   330
         Left            =   6360
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   2225
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild06.frx":10F2
         Caption         =   "BookPOChild06.frx":1112
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild06.frx":117E
         Keys            =   "BookPOChild06.frx":119C
         Spin            =   "BookPOChild06.frx":11E6
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
         ValueVT         =   1970405381
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput8 
         Height          =   330
         Left            =   6360
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   2540
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild06.frx":120E
         Caption         =   "BookPOChild06.frx":122E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild06.frx":129A
         Keys            =   "BookPOChild06.frx":12B8
         Spin            =   "BookPOChild06.frx":1302
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
         ValueVT         =   1972633605
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput10 
         Height          =   330
         Left            =   6360
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   2855
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild06.frx":132A
         Caption         =   "BookPOChild06.frx":134A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild06.frx":13B6
         Keys            =   "BookPOChild06.frx":13D4
         Spin            =   "BookPOChild06.frx":141E
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
         ValueVT         =   1972633605
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput16 
         Height          =   330
         Left            =   6360
         TabIndex        =   22
         Top             =   4290
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild06.frx":1446
         Caption         =   "BookPOChild06.frx":1466
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild06.frx":14D2
         Keys            =   "BookPOChild06.frx":14F0
         Spin            =   "BookPOChild06.frx":153A
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
         ValueVT         =   1970405381
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput1 
         Height          =   330
         Left            =   3840
         TabIndex        =   4
         Top             =   1275
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   582
         Calculator      =   "BookPOChild06.frx":1562
         Caption         =   "BookPOChild06.frx":1582
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild06.frx":15EE
         Keys            =   "BookPOChild06.frx":160C
         Spin            =   "BookPOChild06.frx":1656
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
         ValueVT         =   1972502533
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput2 
         Height          =   330
         Left            =   6360
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1280
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild06.frx":167E
         Caption         =   "BookPOChild06.frx":169E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild06.frx":170A
         Keys            =   "BookPOChild06.frx":1728
         Spin            =   "BookPOChild06.frx":1772
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
         ReadOnly        =   1
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   1970405381
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput15 
         Height          =   330
         Left            =   6360
         TabIndex        =   9
         Top             =   1590
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild06.frx":179A
         Caption         =   "BookPOChild06.frx":17BA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild06.frx":1826
         Keys            =   "BookPOChild06.frx":1844
         Spin            =   "BookPOChild06.frx":188E
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "#0"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "#0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   1972633605
         Value           =   2
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput3 
         Height          =   330
         Left            =   1320
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1905
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild06.frx":18B6
         Caption         =   "BookPOChild06.frx":18D6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild06.frx":1942
         Keys            =   "BookPOChild06.frx":1960
         Spin            =   "BookPOChild06.frx":19AA
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
         ReadOnly        =   1
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   1972633605
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput12 
         Height          =   330
         Left            =   3840
         TabIndex        =   19
         Top             =   3750
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   582
         Calculator      =   "BookPOChild06.frx":19D2
         Caption         =   "BookPOChild06.frx":19F2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild06.frx":1A5E
         Keys            =   "BookPOChild06.frx":1A7C
         Spin            =   "BookPOChild06.frx":1AC6
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "#0"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "#0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   1970405381
         Value           =   4
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput11 
         Height          =   330
         Left            =   1320
         TabIndex        =   18
         Top             =   3750
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild06.frx":1AEE
         Caption         =   "BookPOChild06.frx":1B0E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild06.frx":1B7A
         Keys            =   "BookPOChild06.frx":1B98
         Spin            =   "BookPOChild06.frx":1BE2
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
         ValueVT         =   1970405381
         Value           =   4
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput17 
         Height          =   330
         Left            =   1800
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   2535
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   582
         Calculator      =   "BookPOChild06.frx":1C0A
         Caption         =   "BookPOChild06.frx":1C2A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild06.frx":1C96
         Keys            =   "BookPOChild06.frx":1CB4
         Spin            =   "BookPOChild06.frx":1CFE
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
         ValueVT         =   1970405381
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput18 
         Height          =   330
         Left            =   1320
         TabIndex        =   15
         Top             =   2535
         Width           =   495
         _Version        =   65536
         _ExtentX        =   873
         _ExtentY        =   582
         Calculator      =   "BookPOChild06.frx":1D26
         Caption         =   "BookPOChild06.frx":1D46
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild06.frx":1DB2
         Keys            =   "BookPOChild06.frx":1DD0
         Spin            =   "BookPOChild06.frx":1E1A
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
         ValueVT         =   1970405381
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput13 
         Height          =   330
         Left            =   6360
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   3750
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild06.frx":1E42
         Caption         =   "BookPOChild06.frx":1E62
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild06.frx":1ECE
         Keys            =   "BookPOChild06.frx":1EEC
         Spin            =   "BookPOChild06.frx":1F36
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "########0.000"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "########0.000"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999999.999
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   1
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   1970405381
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBDate6Ctl.TDBDate MhDateInput1 
         Height          =   330
         Left            =   3840
         TabIndex        =   0
         Top             =   105
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   582
         Calendar        =   "BookPOChild06.frx":1F5E
         Caption         =   "BookPOChild06.frx":2076
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild06.frx":20E2
         Keys            =   "BookPOChild06.frx":2100
         Spin            =   "BookPOChild06.frx":215E
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
         Calendar        =   "BookPOChild06.frx":2186
         Caption         =   "BookPOChild06.frx":229E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild06.frx":230A
         Keys            =   "BookPOChild06.frx":2328
         Spin            =   "BookPOChild06.frx":2386
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
         Left            =   3840
         TabIndex        =   21
         Top             =   4290
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   582
         Calendar        =   "BookPOChild06.frx":23AE
         Caption         =   "BookPOChild06.frx":24C6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild06.frx":2532
         Keys            =   "BookPOChild06.frx":2550
         Spin            =   "BookPOChild06.frx":25AE
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput19 
         Height          =   330
         Left            =   1320
         TabIndex        =   6
         ToolTipText     =   "Front"
         Top             =   1590
         Width           =   495
         _Version        =   65536
         _ExtentX        =   873
         _ExtentY        =   582
         Calculator      =   "BookPOChild06.frx":25D6
         Caption         =   "BookPOChild06.frx":25F6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild06.frx":2662
         Keys            =   "BookPOChild06.frx":2680
         Spin            =   "BookPOChild06.frx":26CA
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
         MaxValue        =   8
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
         Value           =   4
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput20 
         Height          =   330
         Left            =   1800
         TabIndex        =   7
         ToolTipText     =   "Back"
         Top             =   1590
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   582
         Calculator      =   "BookPOChild06.frx":26F2
         Caption         =   "BookPOChild06.frx":2712
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild06.frx":277E
         Keys            =   "BookPOChild06.frx":279C
         Spin            =   "BookPOChild06.frx":27E6
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
         MaxValue        =   8
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
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel29 
         Height          =   330
         Left            =   5040
         TabIndex        =   64
         Top             =   960
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
         Caption         =   " Processing"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild06.frx":280E
         Picture         =   "BookPOChild06.frx":282A
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel32 
         Height          =   330
         Index           =   0
         Left            =   120
         TabIndex        =   65
         Top             =   5130
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
         Caption         =   " Adj.Remarks"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild06.frx":2846
         Picture         =   "BookPOChild06.frx":2862
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel321 
         Height          =   330
         Left            =   5040
         TabIndex        =   66
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
         Caption         =   " Extend Date"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild06.frx":287E
         Picture         =   "BookPOChild06.frx":289A
      End
      Begin TDBDate6Ctl.TDBDate MhDateInput31 
         Height          =   330
         Left            =   6360
         TabIndex        =   67
         Top             =   640
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calendar        =   "BookPOChild06.frx":28B6
         Caption         =   "BookPOChild06.frx":29CE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild06.frx":2A3A
         Keys            =   "BookPOChild06.frx":2A58
         Spin            =   "BookPOChild06.frx":2AB6
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
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel51 
         Height          =   330
         Left            =   120
         TabIndex        =   68
         Top             =   2855
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
         Caption         =   " Unit Cost"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild06.frx":2ADE
         Picture         =   "BookPOChild06.frx":2AFA
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput101 
         Height          =   330
         Left            =   1320
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   2855
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild06.frx":2B16
         Caption         =   "BookPOChild06.frx":2B36
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild06.frx":2BA2
         Keys            =   "BookPOChild06.frx":2BC0
         Spin            =   "BookPOChild06.frx":2C0A
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
         MaxValue        =   999999999.999
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   1
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   1972633605
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel32 
         Height          =   330
         Index           =   1
         Left            =   120
         TabIndex        =   70
         Top             =   5445
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
         Caption         =   " Created On"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild06.frx":2C32
         Picture         =   "BookPOChild06.frx":2C4E
      End
      Begin TDBDate6Ctl.TDBDate MhDateInput21 
         Height          =   330
         Left            =   1320
         TabIndex        =   71
         Top             =   5445
         Width           =   6135
         _Version        =   65536
         _ExtentX        =   10821
         _ExtentY        =   582
         Calendar        =   "BookPOChild06.frx":2C6A
         Caption         =   "BookPOChild06.frx":2D82
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild06.frx":2DEE
         Keys            =   "BookPOChild06.frx":2E0C
         Spin            =   "BookPOChild06.frx":2E6A
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
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
         Height          =   330
         Index           =   1
         Left            =   2400
         TabIndex        =   72
         Top             =   1910
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
         Caption         =   " Print Repeat"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild06.frx":2E92
         Picture         =   "BookPOChild06.frx":2EAE
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel32 
         Height          =   330
         Index           =   2
         Left            =   120
         TabIndex        =   73
         Top             =   5760
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
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
         Caption         =   " PDF Received Date"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild06.frx":2ECA
         Picture         =   "BookPOChild06.frx":2EE6
      End
      Begin TDBDate6Ctl.TDBDate MhDateInput22 
         Height          =   330
         Left            =   2280
         TabIndex        =   74
         Top             =   5760
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calendar        =   "BookPOChild06.frx":2F02
         Caption         =   "BookPOChild06.frx":301A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild06.frx":3086
         Keys            =   "BookPOChild06.frx":30A4
         Spin            =   "BookPOChild06.frx":3102
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
      Begin TDBDate6Ctl.TDBDate MhDateInput23 
         Height          =   330
         Left            =   5640
         TabIndex        =   75
         Top             =   5760
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   582
         Calendar        =   "BookPOChild06.frx":312A
         Caption         =   "BookPOChild06.frx":3242
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild06.frx":32AE
         Keys            =   "BookPOChild06.frx":32CC
         Spin            =   "BookPOChild06.frx":332A
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
         Index           =   3
         Left            =   3840
         TabIndex        =   76
         Top             =   5760
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
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
         Caption         =   " PDF Send Date"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild06.frx":3352
         Picture         =   "BookPOChild06.frx":336E
      End
      Begin MSForms.ComboBox Combo11 
         Height          =   330
         Left            =   3840
         TabIndex        =   10
         Top             =   1910
         Width           =   3615
         VariousPropertyBits=   545282075
         BackColor       =   16777215
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "6376;582"
         ListRows        =   3
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
         Left            =   6360
         TabIndex        =   2
         Top             =   960
         Width           =   1095
         VariousPropertyBits=   545282075
         BackColor       =   16777215
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "1931;582"
         ListRows        =   3
         MatchEntry      =   0
         ShowDropButtonWhen=   1
         SpecialEffect   =   0
         FontName        =   "Calibri"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Line Line4 
         X1              =   0
         X2              =   8300
         Y1              =   4710
         Y2              =   4710
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   8300
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   8300
         Y1              =   4185
         Y2              =   4185
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   330
         Left            =   3840
         TabIndex        =   8
         Top             =   1590
         Width           =   1215
         VariousPropertyBits=   545282075
         BackColor       =   16777215
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "2143;582"
         MatchEntry      =   0
         ShowDropButtonWhen=   1
         SpecialEffect   =   0
         FontName        =   "Calibri"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   8300
         Y1              =   3330
         Y2              =   3330
      End
   End
End
Attribute VB_Name = "FrmBookPOChild06"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public rstBookPOChild06 As New ADODB.Recordset
Dim rstPaperList As New ADODB.Recordset
Dim rstRefList As New ADODB.Recordset
Dim rstPrinterRates As New ADODB.Recordset
Public PrinterCode As String
Public TextPrinterQuantity As Long

Dim BookCode As String
Dim SizeCode As String
Dim RefCode As String
Dim PaperCode As String
Dim BillingQuantity As Double
Dim PaperBalance As Long

Private Sub Combo11_Validate(Cancel As Boolean)
    If Combo11.Text = "2nd Print" Or Combo11.Text = "3rd Print" Then
       MhRealInput4.Text = "0.00"
       MhRealInput7.Text = "0.00"
    End If
    CalculateTotalAmount
End Sub

Private Sub Form_Load()
    
    On Error GoTo ErrorHandler
    
    CenterForm Me
    BusySystemIndicator True
    DisableCloseButton Me
    
    AbortPO = False
    BookCode = FrmBookPrintOrder.rstBookList.Fields("Code").Value
    SizeCode = FrmBookPrintOrder.rstBookList.Fields("SizeCode").Value
    Text5.Text = Trim(FrmBookPrintOrder.Text2.Text)
    Text4.Text = Trim(FrmBookPrintOrder.Text6.Text)
    Text2.Text = Trim(FrmBookPrintOrder.Text3.Text)
    Combo1.AddItem "Old", 0
    Combo1.AddItem "New", 1
    Combo1.AddItem "Revised", 2
    Combo2.AddItem "Deepatch", 0
    Combo2.AddItem "PS", 1
    Combo2.AddItem "Wipeon", 2
    Combo2.AddItem "CTP", 3
    
    Combo11.AddItem "Ist Print", 0
    Combo11.AddItem "2nd Print", 1
    Combo11.AddItem "3rd Print", 2
    ClearFields
    rstPaperList.Open "Select Name As Col0, Code From PaperMaster Where PaperMaster.Type = '2' Order by Name", CxnDatabase, adOpenKeyset, adLockReadOnly
    rstPaperList.ActiveConnection = Nothing
    Call LoadRefList(BookCode, CheckNull(rstBookPOChild06.Fields("Code").Value))
    If Val(CheckNull(rstBookPOChild06.Fields("ActualQuantity").Value)) = 0 Then
        
        MhRealInput19.Text = Format(FrmBookPrintOrder.rstBookList.Fields("TitleFrontColor").Value, "0")
        MhRealInput20.Text = Format(FrmBookPrintOrder.rstBookList.Fields("TitleBackColor").Value, "0")
        MhRealInput3.Text = Format(Val(MhRealInput19.Text) + Val(MhRealInput20.Text), "0")
        Combo2.ListIndex = Val(FrmBookPrintOrder.rstBookList.Fields("TitlePlateType").Value) - 1
        MhDateInput1.Text = Format(GetDate(FrmBookPrintOrder.MhDateInput1.Text), "dd-MM-yyyy")
        'MhDateInput3.Text = Format(DateAdd("d", 15, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
        MhRealInput1.Text = Format(TextPrinterQuantity, "0")
        
        
    Else
        LoadFields
    End If
    
    BusySystemIndicator False
    
    Exit Sub
    
ErrorHandler:

    BusySystemIndicator False
    Call CloseForm(Me)

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If Shift = 0 And KeyCode = vbKeyReturn Then
       Sendkeys "{TAB}"
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
        Call CloseForm(Me)
    End If
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    
    TextPrinterQuantity = 0
    Call CloseRecordset(rstPaperList)
    Call CloseRecordset(rstRefList)
    Call CloseRecordset(rstPrinterRates)
  
End Sub
Private Sub ClearFields()
    MhDateInput1.Text = Format(Date, "dd-MM-yyyy")
    MhDateInput2.Text = "  -  -    "
    MhDateInput22.Text = "  -  -    "
    MhDateInput23.Text = "  -  -    "
    
    MhDateInput21.Text = Format(Date, "dd-MM-yyyy")
    MhDateInput3.Text = Format(DateAdd("d", 15, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
    MhRealInput19.Text = "0"
    MhRealInput20.Text = "0"
    Combo1.ListIndex = 0
    Combo2.ListIndex = 0
    Combo11.ListIndex = -1
    
    Text1.Text = ""
    Text3.Text = ""
    Text6.Text = ""
    Text8.Text = ""
    
    MhRealInput1.Text = "0"
    MhRealInput2.Text = "0"
    MhRealInput3.Text = "4"
    MhRealInput4.Text = "0.00"
    MhRealInput5.Text = "0.00"
    MhRealInput6.Text = "0.00"
    MhRealInput7.Text = "0.00"
    MhRealInput8.Text = "0.00"
    MhRealInput9.Text = "0.00"
    MhRealInput10.Text = "0.00"
    MhRealInput11.Text = "0.00"
    MhRealInput12.Text = "4"
    MhRealInput13.Text = "0.000"
    MhRealInput15.Text = "2"
    MhRealInput16.Text = "0.00"
    MhRealInput17.Text = "0.00"
    MhRealInput18.Text = "0.00"
    MhRealInput101.Text = "0.000"
    TxtAdNar.Text = ""
    RefCode = ""
    BillingQuantity = 0
    
End Sub
Private Sub LoadFields()
    If rstBookPOChild06.RecordCount = 0 Then Exit Sub
    
    MhDateInput1.Text = Format(rstBookPOChild06.Fields("OrderDate").Value, "dd-MM-yyyy")
    MhDateInput3.Text = Format(rstBookPOChild06.Fields("TargetDate").Value, "dd-MM-yyyy")
    
    
    If Not IsNull(rstBookPOChild06.Fields("CreatedOn").Value) Then
        MhDateInput21.Text = Format(rstBookPOChild06.Fields("CreatedOn").Value, "dd-MM-yyyy")
    Else
        MhDateInput21.Text = "  -  -    "
    End If
          
    If Not IsNull(rstBookPOChild06.Fields("ExtendDate").Value) Then
        MhDateInput31.Text = Format(rstBookPOChild06.Fields("ExtendDate").Value, "dd-MM-yyyy")
    End If
    
    If rstBookPOChild06.Fields("PlateMaking").Value <> "" Then
       Combo11.Text = Trim(rstBookPOChild06.Fields("PlateMaking").Value)
    End If
    
    Combo1.ListIndex = IIf(rstBookPOChild06.Fields("Processing").Value = "O", 0, IIf(rstBookPOChild06.Fields("Processing").Value = "N", 1, 2))
    RefCode = rstBookPOChild06.Fields("Ref").Value
    If rstRefList.RecordCount > 0 Then rstRefList.MoveFirst
    rstRefList.Find "[Code] = '" & RefCode & "'"
    If Not rstRefList.EOF Then
        Text3.Text = Trim(rstRefList.Fields("Name").Value)
    End If
    MhRealInput1.Text = Format(Val(rstBookPOChild06.Fields("ActualQuantity").Value), "0")
    MhRealInput2.Text = Format(Val(rstBookPOChild06.Fields("BillingQuantity").Value), "0")
    BillingQuantity = Val(rstBookPOChild06.Fields("BillingQuantity").Value)
    MhRealInput19.Text = Format(Val(rstBookPOChild06.Fields("FrontPrintingType").Value), "0")
    MhRealInput20.Text = Format(Val(rstBookPOChild06.Fields("BackPrintingType").Value), "0")
    Combo2.ListIndex = Val(rstBookPOChild06.Fields("PlateType").Value) - 1
    MhRealInput15.Text = Format(Val(rstBookPOChild06.Fields("Titles/Sheet1").Value), "0")
    MhRealInput3.Text = Format(Val(rstBookPOChild06.Fields("TotalPlates").Value), "0")
    MhRealInput6.Text = Format(Val(rstBookPOChild06.Fields("TotalForms").Value), "0.00")
    MhRealInput5.Text = Format(Val(rstBookPOChild06.Fields("PrintRate").Value), "0.00")
    
    MhRealInput4.Text = Format(Val(rstBookPOChild06.Fields("PlateRate").Value), "0.00")
    MhRealInput7.Text = Format(Val(rstBookPOChild06.Fields("PlateAmount").Value), "0.00")
    MhRealInput8.Text = Format(Val(rstBookPOChild06.Fields("PrintAmount").Value), "0.00")
    MhRealInput9.Text = Format(Val(rstBookPOChild06.Fields("Adjustment").Value), "0.00")
    MhRealInput10.Text = Format(Val(rstBookPOChild06.Fields("BillAmount").Value), "0.00")
    
    If IsNull(rstBookPOChild06.Fields("UnitCost").Value) Or rstBookPOChild06.Fields("UnitCost").Value = "0" Then
       MhRealInput101.Text = Format(Val(MhRealInput10.Text) / Val(MhRealInput1.Text), "0.000")   'Unit Cost
    Else
        MhRealInput101.Text = Format(Val(rstBookPOChild06.Fields("UnitCost").Value), "0.000")
    End If
    PaperCode = rstBookPOChild06.Fields("Paper").Value
    If rstPaperList.RecordCount > 0 Then rstPaperList.MoveFirst
    rstPaperList.Find "[Code] = '" & PaperCode & "'"
    If Not rstPaperList.EOF Then
       Text1.Text = rstPaperList.Fields("Col0").Value
    End If
    MhRealInput11.Text = Format(Val(rstBookPOChild06.Fields("PaperWastage%").Value), "0.00")
    MhRealInput12.Text = Format(Val(rstBookPOChild06.Fields("Titles/Sheet2").Value), "0")
    MhRealInput13.Text = Format(Val(rstBookPOChild06.Fields("PaperConsumptionOther").Value), "0.000")
    PaperBalance = CalculatePaperBalance(PrinterCode, PaperCode, CheckNull(rstBookPOChild06.Fields("Code").Value), "BPOT")
    Text8.Text = rstBookPOChild06.Fields("BillNo").Value
    
    If Not IsNull(rstBookPOChild06.Fields("BillDate").Value) Then
        MhDateInput2.Text = Format(rstBookPOChild06.Fields("BillDate").Value, "dd-MM-yyyy")
    End If
    
    If Not IsNull(rstBookPOChild06.Fields("PDFSendToProduction").Value) Then
        MhDateInput22.Text = Format(rstBookPOChild06.Fields("PDFSendToProduction").Value, "dd-MM-yyyy")
    End If
    
    If Not IsNull(rstBookPOChild06.Fields("PDFSendToPrinter").Value) Then
        MhDateInput23.Text = Format(rstBookPOChild06.Fields("PDFSendToPrinter").Value, "dd-MM-yyyy")
    End If
    
    
    MhRealInput18.Text = Format(Val(rstBookPOChild06.Fields("VAT%").Value), "0.00")
    MhRealInput17.Text = Format(Val(rstBookPOChild06.Fields("VAT").Value), "0.00")
    MhRealInput16.Text = Format(Val(rstBookPOChild06.Fields("PaidAmount").Value), "0.00")
    Text6.Text = rstBookPOChild06.Fields("Remarks").Value
    TxtAdNar.Text = rstBookPOChild06.Fields("AdjustmentRemarks").Value
End Sub
Private Sub SaveFields()
    rstBookPOChild06.Fields("OrderDate").Value = GetDate(MhDateInput1.Text)
    rstBookPOChild06.Fields("TargetDate").Value = GetDate(MhDateInput3.Text)
    
    If Not IsDate(MhDateInput21.Text) Then
        rstBookPOChild06.Fields("CreatedOn").Value = Null
    Else
        rstBookPOChild06.Fields("CreatedOn").Value = GetDate(MhDateInput21.Text)
    End If
    
    If Not IsDate(MhDateInput31.Text) Then
         rstBookPOChild06.Fields("ExtendDate").Value = Null
    Else
         rstBookPOChild06.Fields("ExtendDate").Value = GetDate(MhDateInput31.Text)
    End If
    
    rstBookPOChild06.Fields("PlateMaking").Value = Combo11.Text
    rstBookPOChild06.Fields("Processing").Value = IIf(Combo1.ListIndex = 0, "O", IIf(Combo1.ListIndex = 1, "N", "R"))
    rstBookPOChild06.Fields("Ref").Value = RefCode
    
    rstBookPOChild06.Fields("ActualQuantity").Value = Format(Val(MhRealInput1.Text), "0")
    rstBookPOChild06.Fields("BillingQuantity").Value = Format(Val(MhRealInput2.Text), "0")
    rstBookPOChild06.Fields("FrontPrintingType").Value = Format(Val(MhRealInput19.Text), "0")
    rstBookPOChild06.Fields("BackPrintingType").Value = Format(Val(MhRealInput20.Text), "0")
    rstBookPOChild06.Fields("PlateType").Value = Trim(str(Combo2.ListIndex + 1))
    rstBookPOChild06.Fields("Titles/Sheet1").Value = Format(Val(MhRealInput15.Text), "0")
    rstBookPOChild06.Fields("TotalPlates").Value = Format(Val(MhRealInput3.Text), "0")
    rstBookPOChild06.Fields("PlateRate").Value = Format(Val(MhRealInput4.Text), "0.00")
    rstBookPOChild06.Fields("PlateAmount").Value = Format(Val(MhRealInput7.Text), "0.00")
    rstBookPOChild06.Fields("TotalForms").Value = Format(Val(MhRealInput6.Text), "0.00")
    rstBookPOChild06.Fields("PrintRate").Value = Format(Val(MhRealInput5.Text), "0.00")
    rstBookPOChild06.Fields("PrintAmount").Value = Format(Val(MhRealInput8.Text), "0.00")
    rstBookPOChild06.Fields("Adjustment").Value = Format(Val(MhRealInput9.Text), "0.00")
    rstBookPOChild06.Fields("BillAmount").Value = Format(Val(MhRealInput10.Text), "0.00")
    rstBookPOChild06.Fields("Paper").Value = PaperCode
    rstBookPOChild06.Fields("PaperWastage%").Value = Format(Val(MhRealInput11.Text), "0.00")
    rstBookPOChild06.Fields("Titles/Sheet2").Value = Format(Val(MhRealInput12.Text), "0")
    rstBookPOChild06.Fields("PaperConsumptionOther").Value = Format(Val(MhRealInput13.Text), "0.000")
    rstBookPOChild06.Fields("PaperConsumptionSheets").Value = Format(CLng((Int(Val(MhRealInput13.Text)) * 500) + ((Val(MhRealInput13.Text) - Int(Val(MhRealInput13.Text))) * 1000)), "0")
    rstBookPOChild06.Fields("BillNo").Value = Text8.Text
    
    If Not IsDate(MhDateInput2.Text) Then
         rstBookPOChild06.Fields("BillDate").Value = Null
    Else
         rstBookPOChild06.Fields("BillDate").Value = GetDate(MhDateInput2.Text)
    End If
    
    If Not IsDate(MhDateInput22.Text) Then
         rstBookPOChild06.Fields("PDFSendToProduction").Value = Null
    Else
         rstBookPOChild06.Fields("PDFSendToProduction").Value = GetDate(MhDateInput22.Text)
    End If
    
    If Not IsDate(MhDateInput23.Text) Then
         rstBookPOChild06.Fields("PDFSendToPrinter").Value = Null
    Else
         rstBookPOChild06.Fields("PDFSendToPrinter").Value = GetDate(MhDateInput23.Text)
    End If
    
    rstBookPOChild06.Fields("VAT%").Value = Format(Val(MhRealInput18.Text), "0.00")
    rstBookPOChild06.Fields("VAT").Value = Format(Val(MhRealInput17.Text), "0.00")
    rstBookPOChild06.Fields("PaidAmount").Value = Format(Val(MhRealInput16.Text), "0.00")
    rstBookPOChild06.Fields("Remarks").Value = Text6.Text
    rstBookPOChild06.Fields("AdjustmentRemarks").Value = IIf(Val(MhRealInput9.Text) <> 0, TxtAdNar.Text, "")
    If Not CheckEmpty(Text8.Text, False) Then If IsNull(rstBookPOChild06.Fields("BillFeedDate").Value) Then rstBookPOChild06.Fields("BillFeedDate").Value = Now()
    
    Dim lpBuff As String * 1024
    GetComputerName lpBuff, Len(lpBuff)
    If Not CheckEmpty(Text8.Text, False) Then If IsNull(rstBookPOChild06.Fields("ComputerName").Value) Then rstBookPOChild06.Fields("ComputerName").Value = Left(lpBuff, (InStr(1, lpBuff, vbNullChar)) - 1)
    rstBookPOChild06.Fields("UnitCost").Value = Format(Val(MhRealInput101.Text), "0.000")

End Sub
Private Sub MhDateInput1_Validate(Cancel As Boolean)
    If Not IsDate(GetDate(MhDateInput1.Text)) Then
        Cancel = True
    ElseIf Format(GetDate(MhDateInput1.Text), "yyyymmdd") < Format(FinancialYearFrom, "yyyymmdd") Or Format(GetDate(MhDateInput1.Text), "yyyymmdd") > Format(FinancialYearTo, "yyyymmdd") Then
        Cancel = True
    ElseIf Val(CheckNull(rstBookPOChild06.Fields("ActualQuantity").Value)) = 0 Then
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



Private Sub MhRealInput6_Validate(Cancel As Boolean)
    If MhRealInput6.Value > 0 And MhRealInput1.Value > 0 Then
         'Set Target Date based on Quantity and Form
         If MhRealInput6.Value <= 120 And MhRealInput1.Value <= 3300 Then
            MhDateInput3.Text = Format(DateAdd("d", 10, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
         ElseIf MhRealInput6.Value <= 120 And MhRealInput1.Value > 3300 And MhRealInput1.Value <= 5500 Then
            MhDateInput3.Text = Format(DateAdd("d", 12, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
         ElseIf MhRealInput6.Value <= 120 And MhRealInput1.Value > 5500 And MhRealInput1.Value <= 10500 Then
            MhDateInput3.Text = Format(DateAdd("d", 15, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
         ElseIf MhRealInput6.Value <= 120 And MhRealInput1.Value > 10500 Then
            MhDateInput3.Text = Format(DateAdd("d", 20, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
         
         ElseIf MhRealInput6.Value > 121 And MhRealInput6.Value <= 200 And MhRealInput1.Value <= 3300 Then
            MhDateInput3.Text = Format(DateAdd("d", 10, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
         ElseIf MhRealInput6.Value > 121 And MhRealInput6.Value <= 200 And MhRealInput1.Value > 3300 And MhRealInput1.Value <= 5500 Then
            MhDateInput3.Text = Format(DateAdd("d", 12, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
         ElseIf MhRealInput6.Value > 121 And MhRealInput6.Value <= 200 And MhRealInput1.Value > 5500 And MhRealInput1.Value <= 10500 Then
            MhDateInput3.Text = Format(DateAdd("d", 18, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
         ElseIf MhRealInput6.Value > 121 And MhRealInput6.Value <= 200 And MhRealInput1.Value > 10500 Then
            MhDateInput3.Text = Format(DateAdd("d", 20, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
         ElseIf MhRealInput6.Value > 200 And MhRealInput1.Value <= 3300 Then
            MhDateInput3.Text = Format(DateAdd("d", 10, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
         ElseIf MhRealInput6.Value > 200 And MhRealInput1.Value > 3300 And MhRealInput1.Value < 5500 Then
            MhDateInput3.Text = Format(DateAdd("d", 15, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
         ElseIf MhRealInput6.Value > 200 And MhRealInput1.Value > 5500 And MhRealInput1.Value < 10500 Then
            MhDateInput3.Text = Format(DateAdd("d", 20, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
         ElseIf MhRealInput6.Value > 200 And MhRealInput1.Value > 10500 Then
            MhDateInput3.Text = Format(DateAdd("d", 25, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
         End If
     End If
End Sub

Private Sub Text3_Change()
    If Text3.Text = " " Then
        Text3.Text = "?"
        Sendkeys "{TAB}"
    ElseIf CheckEmpty(Text3, False) Then
        RefCode = ""
    End If
End Sub
Private Sub Text3_Validate(Cancel As Boolean)
    Dim SearchString As String
    
    If CheckEmpty(Text3, False) Then
        Exit Sub
    End If
    SearchString = FixQuote(Text3.Text)
    If rstRefList.RecordCount = 0 Then
        DisplayError ("No Pending Reference")
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
        Call LoadSelectionList(rstRefList, "List of References...", "Ref.No.")
        SearchOrder = 0
        Call DisplaySelectionList(Text3, RefCode)
        Call CloseForm(FrmSelectionList)
        If CheckEmpty(Text3.Text, False) Then
            Text3.Text = "?"
        End If
        If RTrim(RefCode) <> "" Then
            Sendkeys "{TAB}"
        End If
        Cancel = True
    Else
        RefCode = rstRefList.Fields("Code").Value
        Text3.Text = Trim(rstRefList.Fields("Name").Value)
        If Val(CheckNull(rstBookPOChild06.Fields("ActualQuantity").Value)) = 0 Then
            MhRealInput1.Text = Format(Val(rstRefList.Fields("BalanceQuantity").Value), "0")
            MhRealInput1_Validate False
        End If
    End If
End Sub
Private Sub Text1_Change()
    If Text1.Text = " " Then
        Text1.Text = "?"
        Sendkeys "{TAB}"
    End If
End Sub
Private Sub Text1_Validate(Cancel As Boolean)
    Dim SearchString As String
    
    SearchString = FixQuote(Text1.Text)
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
        Call DisplaySelectionList(Text1, PaperCode)
        Call CloseForm(FrmSelectionList)
        If CheckEmpty(Text1.Text, False) Then
            Text1.Text = "?"
        End If
        If RTrim(PaperCode) <> "" Then
            Sendkeys "{TAB}"
        End If
        Cancel = True
    Else
        PaperCode = rstPaperList.Fields("Code").Value
    End If
End Sub
Private Sub MhRealInput1_Validate(Cancel As Boolean)    'Actual Quantity
    If Val(MhRealInput1.Text) = 0 Then
        Cancel = True: Exit Sub
    End If
    MhRealInput2.Text = Format(Val(MhRealInput1.Text), "0")
    Call CalculateTotalForms
    Call GetPrinterRates("B"): Call CalculatePrintAmount: Call CalculatePlateAmount: Call CalculateConsumption
    If MhRealInput6.Value > 0 And MhRealInput1.Value > 0 Then
         'Set Target Date based on Quantity and Form
         If MhRealInput6.Value <= 120 And MhRealInput1.Value <= 3300 Then
            MhDateInput3.Text = Format(DateAdd("d", 10, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
         ElseIf MhRealInput6.Value <= 120 And MhRealInput1.Value > 3300 And MhRealInput1.Value <= 5500 Then
            MhDateInput3.Text = Format(DateAdd("d", 12, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
         ElseIf MhRealInput6.Value <= 120 And MhRealInput1.Value > 5500 And MhRealInput1.Value <= 10500 Then
            MhDateInput3.Text = Format(DateAdd("d", 15, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
         ElseIf MhRealInput6.Value <= 120 And MhRealInput1.Value > 10500 Then
            MhDateInput3.Text = Format(DateAdd("d", 20, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
         
         ElseIf MhRealInput6.Value > 121 And MhRealInput6.Value <= 200 And MhRealInput1.Value <= 3300 Then
            MhDateInput3.Text = Format(DateAdd("d", 10, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
         ElseIf MhRealInput6.Value > 121 And MhRealInput6.Value <= 200 And MhRealInput1.Value > 3300 And MhRealInput1.Value <= 5500 Then
            MhDateInput3.Text = Format(DateAdd("d", 12, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
         ElseIf MhRealInput6.Value > 121 And MhRealInput6.Value <= 200 And MhRealInput1.Value > 5500 And MhRealInput1.Value <= 10500 Then
            MhDateInput3.Text = Format(DateAdd("d", 18, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
         ElseIf MhRealInput6.Value > 121 And MhRealInput6.Value <= 200 And MhRealInput1.Value > 10500 Then
            MhDateInput3.Text = Format(DateAdd("d", 20, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
         ElseIf MhRealInput6.Value > 200 And MhRealInput1.Value <= 3300 Then
            MhDateInput3.Text = Format(DateAdd("d", 10, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
         ElseIf MhRealInput6.Value > 200 And MhRealInput1.Value > 3300 And MhRealInput1.Value < 5500 Then
            MhDateInput3.Text = Format(DateAdd("d", 15, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
         ElseIf MhRealInput6.Value > 200 And MhRealInput1.Value > 5500 And MhRealInput1.Value < 10500 Then
            MhDateInput3.Text = Format(DateAdd("d", 20, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
         ElseIf MhRealInput6.Value > 200 And MhRealInput1.Value > 10500 Then
            MhDateInput3.Text = Format(DateAdd("d", 25, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
         End If
     End If


End Sub
Private Sub MhRealInput19_Validate(Cancel As Boolean)   'Title's Front Side Color
    MhRealInput3.Text = Format(Val(MhRealInput19.Text) + Val(MhRealInput20.Text), "0"): Call GetPrinterRates("W")
    Call CalculatePlateAmount: Call CalculatePrintAmount
End Sub
Private Sub MhRealInput20_Validate(Cancel As Boolean)   'Title's Back Side Color
    MhRealInput3.Text = Format(Val(MhRealInput19.Text) + Val(MhRealInput20.Text), "0"): Call GetPrinterRates("W")
    Call CalculatePlateAmount: Call CalculatePrintAmount: CalculateTotalForms
End Sub
Private Sub MhRealInput15_Validate(Cancel As Boolean)   'Titles/Sheet For Calculating Total Forms
    Call CalculateTotalForms:     Call CalculatePrintAmount: Call CalculateConsumption
End Sub
Private Sub MhRealInput4_Validate(Cancel As Boolean)    'Plate Rate
    If Combo11.Text = "2nd Print" Or Combo11.Text = "3rd Print" Then
       MhRealInput4.Text = "0.00"
       MhRealInput7.Text = "0.00"
    Else
       CalculatePlateAmount
    End If
End Sub
Private Sub MhRealInput5_Validate(Cancel As Boolean)    'Print Rate
    CalculatePrintAmount
End Sub
Private Sub MhRealInput18_Validate(Cancel As Boolean)   'VAT%
    CalculateTotalAmount
End Sub
Private Sub MhRealInput9_Validate(Cancel As Boolean)    'Adjustment
    CalculateTotalAmount
End Sub
Private Sub MhRealInput11_Validate(Cancel As Boolean)   'Wastage Percentage
    CalculateConsumption
End Sub
Private Sub MhRealInput12_Validate(Cancel As Boolean)   'Titles/Sheet For Calculating Paper Consumption
    CalculateConsumption
End Sub
Private Sub cmdProceed_Click()
    Dim Stock As Double
    If CheckMandatoryFields Then Exit Sub
    If FrmBookPrintOrder.BookPOType <> "O" Then
        Stock = PaperBalance - CLng((Int(Val(MhRealInput13.Text)) * 500) + (Round(Val(MhRealInput13.Text) - Int(Val(MhRealInput13.Text)), 3) * 1000))
        If Stock < 0 Then
            Stock = Format(CLng(Int(Val(Abs(Stock)) / 500)) + ((Val(Abs(Stock)) Mod 500) / 1000), "0.000")
            If UserLevel <= 2 Then
                If MsgBox("Stock (" & Format(0 - Stock, "0.000") & ") of the Paper - " & Trim(Text1.Text) & vbCrLf & " is going negative ! Would you like to continue ?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Proceed !") = vbNo Then
                    Exit Sub
                End If
            Else
                Call DisplayError("Cann't Save ! Stock (" & Format(0 - Stock, "0.000") & ") of the Paper - " & Trim(Text1.Text) & " is going negative"): AbortPO = True: Exit Sub
            End If
        End If
    End If
    SaveFields
    rstBookPOChild06.Update
    Call CloseForm(Me)
End Sub
Private Sub cmdCancel_Click()
    rstBookPOChild06.CancelUpdate
    Call CloseForm(Me)
End Sub
Private Function CheckMandatoryFields() As Boolean
    If Combo1.ListIndex < 0 Then Combo1.SetFocus: CheckMandatoryFields = True: Exit Function
    If Combo2.ListIndex < 0 Then Combo2.SetFocus: CheckMandatoryFields = True: Exit Function
    If Combo11.ListIndex < 0 Then Combo11.SetFocus: CheckMandatoryFields = True: Exit Function
    If Val(MhRealInput15.Text) <= 0 Then MhRealInput15.SetFocus: CheckMandatoryFields = True: Exit Function
    If CheckEmpty(Text1.Text, False) Then Text1.SetFocus: CheckMandatoryFields = True: Exit Function
    If Not CheckExists(Text1, "Col0", rstPaperList, PaperCode) Then Text1.SetFocus: CheckMandatoryFields = True: Exit Function
    If Val(MhRealInput12.Text) <= 0 Then MhRealInput12.SetFocus: CheckMandatoryFields = True: Exit Function
    If Val(MhRealInput16.Text) <> 0 Then If Val(MhRealInput16.Text) <> Val(MhRealInput10.Text) Then MhRealInput9.SetFocus: CheckMandatoryFields = True: Exit Function
    If Val(MhRealInput9.Text) <> 0 Then If CheckEmpty(TxtAdNar.Text, False) Then TxtAdNar.SetFocus: CheckMandatoryFields = True: Exit Function
End Function
Private Sub GetPrinterRates(ByVal RateType As String)
    Dim PlateRate As Double, PrintRate As Double, PaperWastageRate As Double
    On Error GoTo ErrorHandler
    
    If rstPrinterRates.State = adStateOpen Then rstPrinterRates.Close
    rstPrinterRates.Open "Select Top 1 * From AccountChild06 Where Code = '" & PrinterCode & "' And [Size] = '" & SizeCode & "' And Range1 >= " & Val(MhRealInput2.Text) & " Order By Range1", CxnDatabase, adOpenKeyset, adLockReadOnly
    If rstPrinterRates.RecordCount = 0 Then
        If rstPrinterRates.State = adStateOpen Then rstPrinterRates.Close
        rstPrinterRates.Open "SELECT TOP 1 C.* FROM AccountMaster P INNER JOIN AccountChild06 C ON P.Code=C.Code WHERE Name LIKE '%Rate%' AND [Size]='" & SizeCode & "' AND Range1>= " & Val(MhRealInput2.Text) & " ORDER BY Range1", CxnDatabase, adOpenKeyset, adLockReadOnly
    End If
    If rstPrinterRates.RecordCount > 0 Then
        If RateType = "B" Then PrintRate = Val(rstPrinterRates.Fields("PrintRate1").Value): PlateRate = Val(rstPrinterRates.Fields(IIf(Combo2.ListIndex = 0, "DeepatchPlateRate1", IIf(Combo2.ListIndex = 1, "PSPlateRate1", IIf(Combo2.ListIndex = 2, "WipeonPlateRate1", "CTPPlateRate1")))).Value)
        If Val(MhRealInput19.Text) > 0 Then PaperWastageRate = PaperWastageRate + Val(rstPrinterRates.Fields("PaperWastageRate" & Trim(MhRealInput19.Text)).Value)
        If Val(MhRealInput20.Text) > 0 Then PaperWastageRate = PaperWastageRate + Val(rstPrinterRates.Fields("PaperWastageRate" & Trim(MhRealInput20.Text)).Value)
    End If
    If RateType = "B" Then
        If Val(MhRealInput4.Text) <> PlateRate Then
            If Val(CheckNull(rstBookPOChild06.Fields("ActualQuantity").Value)) <> 0 Then
                If MsgBox("Plate Rate is different from that in Master ! Change rate?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then
                    MhRealInput4.Text = Format(PlateRate, "0.00")
                End If
            Else
                MhRealInput4.Text = Format(PlateRate, "0.00")
            End If
        End If
        If Val(MhRealInput5.Text) <> PrintRate And Val(MhRealInput5.Text) <> 0 Then
            If MsgBox("Print Rate is different from that in Master ! Change rate?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then
                MhRealInput5.Text = Format(PrintRate, "0.00")
            End If
        Else
            MhRealInput5.Text = Format(PrintRate, "0.00")
        End If
    End If
    If Val(MhRealInput11.Text) <> PaperWastageRate And Val(MhRealInput11.Text) <> 0 Then
        If MsgBox("Paper Wastage is different from that in Master ! Change rate?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then
            MhRealInput11.Text = Format(PaperWastageRate, "0.00")
        End If
    Else
        MhRealInput11.Text = Format(PaperWastageRate, "0.00")
    End If
    If RateType = "B" Then Call CalculatePrintAmount: Call CalculatePlateAmount
    Call CalculateConsumption
Exit Sub
ErrorHandler:
    DisplayError ("Failed to Fetch Printer Rates")
End Sub
Private Sub LoadRefList(ByVal strBookCode As String, ByVal strOrderCode As String)
    Dim BalanceQuantity As Long
    On Error GoTo ErrorHandler
    If rstRefList.State = adStateOpen Then
        rstRefList.Close
    End If
    rstRefList.Open "Select P.Name,Quantity As PlannedQuantity,Format((Select Sum(ActualQuantity) From BookPOChild06,BookPOParent Where BookPOChild06.Ref=P.Code And BookPOParent.Code=BookPOChild06.Code And BookPOParent.Book=C.Book And BookPOChild06.Code<>'" & strOrderCode & "'),0) As PrintedQuantity,Quantity As BalanceQuantity,Remarks As Col0,[PaperWastage%],P.Code From PrintPVParent P,PrintPVChild C Where P.Code=C.Code And P.PlanningType ='2' And C.Book='" & strBookCode & "' Order By P.Name", CxnDatabase, adOpenKeyset, adLockOptimistic
    rstRefList.ActiveConnection = Nothing
    Do While Not rstRefList.EOF
        BalanceQuantity = (Val(CheckNull(rstRefList.Fields("PlannedQuantity").Value)) - Val(CheckNull(rstRefList.Fields("PrintedQuantity").Value)))
        If BalanceQuantity <> 0 Then
            rstRefList.Fields("Col0").Value = Trim(rstRefList.Fields("Name").Value) + " Quantity : " + Format(str(BalanceQuantity), "0")
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
Private Sub CalculateTotalForms()
    If Val(MhRealInput15.Text) > 0 Then
        MhRealInput6.Text = Format(Val(MhRealInput2.Text) / Val(MhRealInput15.Text), "0.00")
        CalculatePrintAmount
    End If
End Sub
Private Sub CalculatePlateAmount()
    MhRealInput7.Text = Format(Val(MhRealInput3.Text) * Val(MhRealInput4.Text), "0.00")
    CalculateTotalAmount
End Sub
Private Sub CalculatePrintAmount()
    MhRealInput8.Text = Format((Val(MhRealInput19.Text) + Val(MhRealInput20.Text)) * IIf(Val(MhRealInput6.Text) < 1000, 1, Val(MhRealInput6.Text) / 1000) * Val(MhRealInput5.Text), "0.00")
'    Dim aa As Double
'    Dim bb As Double
'    Dim cc As Double
'    Dim dd As Double
'    aa = Val(MhRealInput19.Text)
'    bb = Val(MhRealInput20.Text)
'    dd = Val(MhRealInput6.Text) / 1000
'    cc = Val(MhRealInput5.Text)
'    Dim rsl As Double
'    rsl = (Val(MhRealInput19.Text) + Val(MhRealInput20.Text)) * (Val(MhRealInput6.Text) / 1000) * Val(MhRealInput5.Text)
    CalculateTotalAmount
    
End Sub
Private Sub CalculateTotalAmount()
    'MhRealInput17.Text = Format(((Val(MhRealInput7.Text) + Val(MhRealInput8.Text)) * 80 / 100) * Val(MhRealInput18.Text) / 100, "0.00")
    MhRealInput17.Text = Format(((Val(MhRealInput7.Text) + Val(MhRealInput8.Text)) * 100 / 100) * Val(MhRealInput18.Text) / 100, "0.00")
    MhRealInput10.Text = Format(Val(MhRealInput7.Text) + Val(MhRealInput8.Text) + Val(MhRealInput17.Text) + Val(MhRealInput9.Text), "0.00")
    
    If Val(MhRealInput10.Text) > 0 Then
       MhRealInput101.Text = Format(Val(MhRealInput10.Text) / Val(MhRealInput1.Text), "0.000") 'Unit Cost
    End If
End Sub

Private Sub CalculateConsumption()
    Dim N As Long, W As Long
    
    If Val(MhRealInput12.Text) > 0 Then
        W = CLng(Val(MhRealInput6.Text) * Val(MhRealInput11.Text) / 100)
        If W < 80 Then
            W = 80 * Val(MhRealInput15.Text) / Val(MhRealInput12.Text)
        Else
            W = CLng((Val(MhRealInput2.Text) / Val(MhRealInput12.Text)) * Val(MhRealInput11.Text) / 100)
        End If
        N = CLng((Val(MhRealInput1.Text) / Val(MhRealInput12.Text)) + W)
        MhRealInput13.Text = Format(CLng(Int(Val(N) / 500)) + ((Val(N) Mod 500) / 1000), "0.000")
        PaperBalance = CalculatePaperBalance(PrinterCode, PaperCode, CheckNull(rstBookPOChild06.Fields("Code").Value), "BPOT")
    End If
    
End Sub
Private Sub Combo2_Validate(Cancel As Boolean)
    
    If Combo2.ListIndex = 1 Or Combo2.ListIndex = 3 Then    'PS/CTP Plate Details
        On Error Resume Next
        FrmPSPlateRegister.BookCode = BookCode
        FrmPSPlateRegister.BookName = Trim(Text2.Text)
        FrmPSPlateRegister.OrderCode = IIf(CheckNull(rstBookPOChild06.Fields("Code").Value) = "", "999999", rstBookPOChild06.Fields("Code").Value)
        FrmPSPlateRegister.OrderDate = GetDate(MhDateInput1.Text)
        FrmPSPlateRegister.OrderType = "06"
        FrmPSPlateRegister.PlateType = ""
        Load FrmPSPlateRegister
        If Err.Number <> 364 Then FrmPSPlateRegister.Show vbModal
        On Error GoTo 0
    End If
    
End Sub
