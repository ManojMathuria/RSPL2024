VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmBookPOChild05 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Book Printing Order Details"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8760
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "BookPOChild05.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "FrmLogin"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   8760
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Height          =   375
      Left            =   8295
      Picture         =   "BookPOChild05.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Cancel"
      Top             =   465
      Width           =   375
   End
   Begin VB.CommandButton cmdProceed 
      Height          =   375
      Left            =   8295
      Picture         =   "BookPOChild05.frx":0544
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Save"
      Top             =   105
      Width           =   375
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
      Height          =   8070
      Left            =   120
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   105
      Width           =   8055
      _Version        =   65536
      _ExtentX        =   14208
      _ExtentY        =   14235
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
      Picture         =   "BookPOChild05.frx":0646
      Begin VB.TextBox TxtPOrder 
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
         TabIndex        =   34
         Top             =   7305
         Width           =   6615
      End
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
         TabIndex        =   33
         Top             =   7005
         Width           =   6615
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
         TabIndex        =   32
         Top             =   6690
         Width           =   6615
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
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   105
         Width           =   1575
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
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   960
         Width           =   4215
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
         TabIndex        =   29
         Top             =   6150
         Width           =   1575
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
         TabIndex        =   25
         Top             =   3750
         Width           =   4215
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
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   645
         Width           =   4215
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
         Width           =   1575
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
         Height          =   330
         Left            =   2880
         TabIndex        =   40
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
         Picture         =   "BookPOChild05.frx":0662
         Picture         =   "BookPOChild05.frx":067E
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
         Height          =   330
         Left            =   120
         TabIndex        =   41
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
         Picture         =   "BookPOChild05.frx":069A
         Picture         =   "BookPOChild05.frx":06B6
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
         Height          =   330
         Left            =   2880
         TabIndex        =   42
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
         Picture         =   "BookPOChild05.frx":06D2
         Picture         =   "BookPOChild05.frx":06EE
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
         Height          =   645
         Left            =   5520
         TabIndex        =   43
         Top             =   1275
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   1138
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
         Picture         =   "BookPOChild05.frx":070A
         Picture         =   "BookPOChild05.frx":0726
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel9 
         Height          =   330
         Left            =   120
         TabIndex        =   44
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
         Picture         =   "BookPOChild05.frx":0742
         Picture         =   "BookPOChild05.frx":075E
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
         Height          =   330
         Left            =   120
         TabIndex        =   45
         Top             =   2225
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
         Picture         =   "BookPOChild05.frx":077A
         Picture         =   "BookPOChild05.frx":0796
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
         Height          =   330
         Left            =   2880
         TabIndex        =   46
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
         Caption         =   " Plate Rate"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild05.frx":07B2
         Picture         =   "BookPOChild05.frx":07CE
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel7 
         Height          =   330
         Left            =   2880
         TabIndex        =   47
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
         Caption         =   " Print Rate"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild05.frx":07EA
         Picture         =   "BookPOChild05.frx":0806
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
         Height          =   330
         Left            =   2880
         TabIndex        =   48
         Top             =   3165
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
         Picture         =   "BookPOChild05.frx":0822
         Picture         =   "BookPOChild05.frx":083E
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel10 
         Height          =   330
         Left            =   120
         TabIndex        =   49
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
         Picture         =   "BookPOChild05.frx":085A
         Picture         =   "BookPOChild05.frx":0876
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
         Height          =   330
         Left            =   120
         TabIndex        =   50
         Top             =   1910
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
         Caption         =   " Pages/Forms"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild05.frx":0892
         Picture         =   "BookPOChild05.frx":08AE
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel12 
         Height          =   330
         Left            =   120
         TabIndex        =   51
         Top             =   2540
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
         Picture         =   "BookPOChild05.frx":08CA
         Picture         =   "BookPOChild05.frx":08E6
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel13 
         Height          =   330
         Left            =   5520
         TabIndex        =   52
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
         Caption         =   " Plate Amount"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild05.frx":0902
         Picture         =   "BookPOChild05.frx":091E
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel14 
         Height          =   330
         Left            =   5520
         TabIndex        =   53
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
         Caption         =   " Print Amount"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild05.frx":093A
         Picture         =   "BookPOChild05.frx":0956
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel15 
         Height          =   330
         Left            =   5520
         TabIndex        =   54
         Top             =   3165
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
         Picture         =   "BookPOChild05.frx":0972
         Picture         =   "BookPOChild05.frx":098E
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel16 
         Height          =   330
         Left            =   120
         TabIndex        =   55
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
         Caption         =   " Paper Name"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild05.frx":09AA
         Picture         =   "BookPOChild05.frx":09C6
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel17 
         Height          =   330
         Index           =   0
         Left            =   120
         TabIndex        =   56
         Top             =   4065
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
         Picture         =   "BookPOChild05.frx":09E2
         Picture         =   "BookPOChild05.frx":09FE
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel18 
         Height          =   330
         Left            =   2880
         TabIndex        =   57
         Top             =   4065
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
         Caption         =   " Consumption"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild05.frx":0A1A
         Picture         =   "BookPOChild05.frx":0A36
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel21 
         Height          =   330
         Left            =   5520
         TabIndex        =   58
         Top             =   4065
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
         Caption         =   " Total Consmptn"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild05.frx":0A52
         Picture         =   "BookPOChild05.frx":0A6E
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel19 
         Height          =   330
         Left            =   120
         TabIndex        =   59
         Top             =   6150
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
         Picture         =   "BookPOChild05.frx":0A8A
         Picture         =   "BookPOChild05.frx":0AA6
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel20 
         Height          =   330
         Left            =   5520
         TabIndex        =   60
         Top             =   6150
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
         Picture         =   "BookPOChild05.frx":0AC2
         Picture         =   "BookPOChild05.frx":0ADE
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel23 
         Height          =   330
         Left            =   2880
         TabIndex        =   61
         Top             =   6150
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
         Picture         =   "BookPOChild05.frx":0AFA
         Picture         =   "BookPOChild05.frx":0B16
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel24 
         Height          =   330
         Left            =   5520
         TabIndex        =   62
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
         Picture         =   "BookPOChild05.frx":0B32
         Picture         =   "BookPOChild05.frx":0B4E
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel25 
         Height          =   330
         Left            =   120
         TabIndex        =   64
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
         Picture         =   "BookPOChild05.frx":0B6A
         Picture         =   "BookPOChild05.frx":0B86
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel26 
         Height          =   330
         Index           =   0
         Left            =   2880
         TabIndex        =   65
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
         Caption         =   " Plate Type"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild05.frx":0BA2
         Picture         =   "BookPOChild05.frx":0BBE
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel27 
         Height          =   330
         Left            =   120
         TabIndex        =   67
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
         Picture         =   "BookPOChild05.frx":0BDA
         Picture         =   "BookPOChild05.frx":0BF6
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel28 
         Height          =   330
         Left            =   120
         TabIndex        =   68
         Top             =   6690
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
         Picture         =   "BookPOChild05.frx":0C12
         Picture         =   "BookPOChild05.frx":0C2E
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel22 
         Height          =   330
         Left            =   120
         TabIndex        =   69
         Top             =   2850
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
         Picture         =   "BookPOChild05.frx":0C4A
         Picture         =   "BookPOChild05.frx":0C66
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel29 
         Height          =   330
         Left            =   5520
         TabIndex        =   70
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
         Picture         =   "BookPOChild05.frx":0C82
         Picture         =   "BookPOChild05.frx":0C9E
      End
      Begin TDBDate6Ctl.TDBDate MhDateInput1 
         Height          =   330
         Left            =   4320
         TabIndex        =   0
         Top             =   105
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   582
         Calendar        =   "BookPOChild05.frx":0CBA
         Caption         =   "BookPOChild05.frx":0DD2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":0E3E
         Keys            =   "BookPOChild05.frx":0E5C
         Spin            =   "BookPOChild05.frx":0EBA
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
         Left            =   6840
         TabIndex        =   1
         Top             =   105
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calendar        =   "BookPOChild05.frx":0EE2
         Caption         =   "BookPOChild05.frx":0FFA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":1066
         Keys            =   "BookPOChild05.frx":1084
         Spin            =   "BookPOChild05.frx":10E2
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
         Left            =   4320
         TabIndex        =   30
         Top             =   6150
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   582
         Calendar        =   "BookPOChild05.frx":110A
         Caption         =   "BookPOChild05.frx":1222
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":128E
         Keys            =   "BookPOChild05.frx":12AC
         Spin            =   "BookPOChild05.frx":130A
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
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel30 
         Height          =   330
         Left            =   5520
         TabIndex        =   71
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
         Caption         =   " Forms/Sheet"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild05.frx":1332
         Picture         =   "BookPOChild05.frx":134E
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel31 
         Height          =   330
         Left            =   5520
         TabIndex        =   72
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
         Caption         =   " Forms/Sheet"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild05.frx":136A
         Picture         =   "BookPOChild05.frx":1386
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput1 
         Height          =   330
         Left            =   4320
         TabIndex        =   4
         Top             =   1275
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":13A2
         Caption         =   "BookPOChild05.frx":13C2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":142E
         Keys            =   "BookPOChild05.frx":144C
         Spin            =   "BookPOChild05.frx":1496
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
         Left            =   6840
         TabIndex        =   5
         ToolTipText     =   "One Color"
         Top             =   1280
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":14BE
         Caption         =   "BookPOChild05.frx":14DE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":154A
         Keys            =   "BookPOChild05.frx":1568
         Spin            =   "BookPOChild05.frx":15B2
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput19 
         Height          =   330
         Left            =   6840
         TabIndex        =   6
         ToolTipText     =   "Double & Four Color"
         Top             =   1590
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":15DA
         Caption         =   "BookPOChild05.frx":15FA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":1666
         Keys            =   "BookPOChild05.frx":1684
         Spin            =   "BookPOChild05.frx":16CE
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput15 
         Height          =   330
         Left            =   1320
         TabIndex        =   8
         ToolTipText     =   "Two & Four Color"
         Top             =   1905
         Width           =   495
         _Version        =   65536
         _ExtentX        =   873
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":16F6
         Caption         =   "BookPOChild05.frx":1716
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":1782
         Keys            =   "BookPOChild05.frx":17A0
         Spin            =   "BookPOChild05.frx":17EA
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
         ValueVT         =   1972502533
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput17 
         Height          =   330
         Left            =   1800
         TabIndex        =   9
         ToolTipText     =   "¼ Form"
         Top             =   1905
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":1812
         Caption         =   "BookPOChild05.frx":1832
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":189E
         Keys            =   "BookPOChild05.frx":18BC
         Spin            =   "BookPOChild05.frx":1906
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput20 
         Height          =   330
         Left            =   2160
         TabIndex        =   10
         ToolTipText     =   "½ Form"
         Top             =   1905
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":192E
         Caption         =   "BookPOChild05.frx":194E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":19BA
         Keys            =   "BookPOChild05.frx":19D8
         Spin            =   "BookPOChild05.frx":1A22
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput21 
         Height          =   330
         Left            =   2520
         TabIndex        =   11
         ToolTipText     =   "1 Form"
         Top             =   1905
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":1A4A
         Caption         =   "BookPOChild05.frx":1A6A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":1AD6
         Keys            =   "BookPOChild05.frx":1AF4
         Spin            =   "BookPOChild05.frx":1B3E
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput22 
         Height          =   330
         Left            =   6840
         TabIndex        =   13
         Top             =   1910
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":1B66
         Caption         =   "BookPOChild05.frx":1B86
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":1BF2
         Keys            =   "BookPOChild05.frx":1C10
         Spin            =   "BookPOChild05.frx":1C5A
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "0.00"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "0.00"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   2
         MinValue        =   0.5
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   1974075397
         Value           =   1
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput3 
         Height          =   330
         Left            =   1320
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "¼ Form"
         Top             =   2225
         Width           =   540
         _Version        =   65536
         _ExtentX        =   952
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":1C82
         Caption         =   "BookPOChild05.frx":1CA2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":1D0E
         Keys            =   "BookPOChild05.frx":1D2C
         Spin            =   "BookPOChild05.frx":1D76
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
         ValueVT         =   1974075397
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput23 
         Height          =   330
         Left            =   1850
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "½ Form"
         Top             =   2225
         Width           =   540
         _Version        =   65536
         _ExtentX        =   952
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":1D9E
         Caption         =   "BookPOChild05.frx":1DBE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":1E2A
         Keys            =   "BookPOChild05.frx":1E48
         Spin            =   "BookPOChild05.frx":1E92
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
         ValueVT         =   1974075397
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput24 
         Height          =   330
         Left            =   2375
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "1 Form"
         Top             =   2220
         Width           =   520
         _Version        =   65536
         _ExtentX        =   917
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":1EBA
         Caption         =   "BookPOChild05.frx":1EDA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":1F46
         Keys            =   "BookPOChild05.frx":1F64
         Spin            =   "BookPOChild05.frx":1FAE
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
         ValueVT         =   1974075397
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput4 
         Height          =   330
         Left            =   4320
         TabIndex        =   18
         Top             =   2540
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":1FD6
         Caption         =   "BookPOChild05.frx":1FF6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":2062
         Keys            =   "BookPOChild05.frx":2080
         Spin            =   "BookPOChild05.frx":20CA
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###########0.00"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "###########0.00"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999999999.99
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput7 
         Height          =   330
         Left            =   6840
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   2540
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":20F2
         Caption         =   "BookPOChild05.frx":2112
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":217E
         Keys            =   "BookPOChild05.frx":219C
         Spin            =   "BookPOChild05.frx":21E6
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###########0.00"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "###########0.00"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999999999.99
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   -1
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   1974599685
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput6 
         Height          =   330
         Left            =   1320
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "¼ Form"
         Top             =   2540
         Width           =   540
         _Version        =   65536
         _ExtentX        =   952
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":220E
         Caption         =   "BookPOChild05.frx":222E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":229A
         Keys            =   "BookPOChild05.frx":22B8
         Spin            =   "BookPOChild05.frx":2302
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
         ValueVT         =   1974075397
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput25 
         Height          =   330
         Left            =   1850
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "½ Form"
         Top             =   2540
         Width           =   540
         _Version        =   65536
         _ExtentX        =   952
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":232A
         Caption         =   "BookPOChild05.frx":234A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":23B6
         Keys            =   "BookPOChild05.frx":23D4
         Spin            =   "BookPOChild05.frx":241E
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
         ValueVT         =   1974075397
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput26 
         Height          =   330
         Left            =   2375
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "1 Form"
         Top             =   2540
         Width           =   520
         _Version        =   65536
         _ExtentX        =   917
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":2446
         Caption         =   "BookPOChild05.frx":2466
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":24D2
         Keys            =   "BookPOChild05.frx":24F0
         Spin            =   "BookPOChild05.frx":253A
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
         ValueVT         =   1974075397
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput5 
         Height          =   330
         Left            =   4320
         TabIndex        =   22
         Top             =   2855
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":2562
         Caption         =   "BookPOChild05.frx":2582
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":25EE
         Keys            =   "BookPOChild05.frx":260C
         Spin            =   "BookPOChild05.frx":2656
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###########0.00"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "###########0.00"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999999999.99
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput8 
         Height          =   330
         Left            =   6840
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   2855
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":267E
         Caption         =   "BookPOChild05.frx":269E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":270A
         Keys            =   "BookPOChild05.frx":2728
         Spin            =   "BookPOChild05.frx":2772
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###########0.00"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "###########0.00"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999999999.99
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   -1
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   1973878789
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput14 
         Height          =   330
         Left            =   1320
         TabIndex        =   23
         Top             =   2850
         Width           =   810
         _Version        =   65536
         _ExtentX        =   1429
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":279A
         Caption         =   "BookPOChild05.frx":27BA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":2826
         Keys            =   "BookPOChild05.frx":2844
         Spin            =   "BookPOChild05.frx":288E
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
         ValueVT         =   1974075397
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput18 
         Height          =   330
         Left            =   2115
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   2850
         Width           =   780
         _Version        =   65536
         _ExtentX        =   1376
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":28B6
         Caption         =   "BookPOChild05.frx":28D6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":2942
         Keys            =   "BookPOChild05.frx":2960
         Spin            =   "BookPOChild05.frx":29AA
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###########0.00"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "###########0.00"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999999999.99
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   -1
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   1974075397
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput9 
         Height          =   330
         Left            =   4320
         TabIndex        =   24
         Top             =   3165
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":29D2
         Caption         =   "BookPOChild05.frx":29F2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":2A5E
         Keys            =   "BookPOChild05.frx":2A7C
         Spin            =   "BookPOChild05.frx":2AC6
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###########0.00"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "###########0.00"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999999999.99
         MinValue        =   -999999999999.99
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   1974075397
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput10 
         Height          =   330
         Left            =   6840
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   3165
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":2AEE
         Caption         =   "BookPOChild05.frx":2B0E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":2B7A
         Keys            =   "BookPOChild05.frx":2B98
         Spin            =   "BookPOChild05.frx":2BE2
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###########0.00"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "###########0.00"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999999999.99
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   -1
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   1974075397
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput27 
         Height          =   330
         Left            =   6840
         TabIndex        =   26
         Top             =   3750
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":2C0A
         Caption         =   "BookPOChild05.frx":2C2A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":2C96
         Keys            =   "BookPOChild05.frx":2CB4
         Spin            =   "BookPOChild05.frx":2CFE
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "0.00"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "0.00"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   2
         MinValue        =   0.5
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   5
         Value           =   1
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput11 
         Height          =   330
         Left            =   1320
         TabIndex        =   27
         Top             =   4065
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":2D26
         Caption         =   "BookPOChild05.frx":2D46
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":2DB2
         Keys            =   "BookPOChild05.frx":2DD0
         Spin            =   "BookPOChild05.frx":2E1A
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput12 
         Height          =   330
         Left            =   4320
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   4065
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":2E42
         Caption         =   "BookPOChild05.frx":2E62
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":2ECE
         Keys            =   "BookPOChild05.frx":2EEC
         Spin            =   "BookPOChild05.frx":2F36
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
         ValueVT         =   1964507141
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput13 
         Height          =   330
         Left            =   6840
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   4065
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":2F5E
         Caption         =   "BookPOChild05.frx":2F7E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":2FEA
         Keys            =   "BookPOChild05.frx":3008
         Spin            =   "BookPOChild05.frx":3052
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput16 
         Height          =   330
         Left            =   6840
         TabIndex        =   31
         Top             =   6150
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":307A
         Caption         =   "BookPOChild05.frx":309A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":3106
         Keys            =   "BookPOChild05.frx":3124
         Spin            =   "BookPOChild05.frx":316E
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###########0.00"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "###########0.00"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999999999.99
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   1972764677
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin FPSpreadADO.fpSpread fpSpread1 
         Height          =   1335
         Left            =   120
         TabIndex        =   28
         Top             =   4605
         Width           =   7815
         _Version        =   524288
         _ExtentX        =   13785
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
         SpreadDesigner  =   "BookPOChild05.frx":3196
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel32 
         Height          =   330
         Index           =   0
         Left            =   120
         TabIndex        =   79
         Top             =   7005
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
         Picture         =   "BookPOChild05.frx":3EA4
         Picture         =   "BookPOChild05.frx":3EC0
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel32 
         Height          =   330
         Index           =   1
         Left            =   5520
         TabIndex        =   80
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
         Picture         =   "BookPOChild05.frx":3EDC
         Picture         =   "BookPOChild05.frx":3EF8
      End
      Begin TDBDate6Ctl.TDBDate MhDateInput31 
         Height          =   330
         Left            =   6840
         TabIndex        =   81
         Top             =   645
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calendar        =   "BookPOChild05.frx":3F14
         Caption         =   "BookPOChild05.frx":402C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":4098
         Keys            =   "BookPOChild05.frx":40B6
         Spin            =   "BookPOChild05.frx":4114
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
         TabIndex        =   82
         Top             =   3165
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
         Picture         =   "BookPOChild05.frx":413C
         Picture         =   "BookPOChild05.frx":4158
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput101 
         Height          =   330
         Left            =   1320
         TabIndex        =   83
         TabStop         =   0   'False
         Top             =   3165
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         Calculator      =   "BookPOChild05.frx":4174
         Caption         =   "BookPOChild05.frx":4194
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":4200
         Keys            =   "BookPOChild05.frx":421E
         Spin            =   "BookPOChild05.frx":4268
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
         ValueVT         =   1974075397
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel17 
         Height          =   330
         Index           =   1
         Left            =   2880
         TabIndex        =   84
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
         Caption         =   " Book Status"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild05.frx":4290
         Picture         =   "BookPOChild05.frx":42AC
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel32 
         Height          =   330
         Index           =   2
         Left            =   120
         TabIndex        =   85
         Top             =   7305
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
         Caption         =   " PO No."
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild05.frx":42C8
         Picture         =   "BookPOChild05.frx":42E4
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel26 
         Height          =   330
         Index           =   1
         Left            =   2880
         TabIndex        =   86
         Top             =   2220
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
         Caption         =   "Print Repeat"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild05.frx":4300
         Picture         =   "BookPOChild05.frx":431C
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel32 
         Height          =   330
         Index           =   3
         Left            =   120
         TabIndex        =   87
         Top             =   7620
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
         Picture         =   "BookPOChild05.frx":4338
         Picture         =   "BookPOChild05.frx":4354
      End
      Begin TDBDate6Ctl.TDBDate MhDateInput17 
         Height          =   330
         Left            =   2280
         TabIndex        =   88
         Top             =   7620
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   582
         Calendar        =   "BookPOChild05.frx":4370
         Caption         =   "BookPOChild05.frx":4488
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":44F4
         Keys            =   "BookPOChild05.frx":4512
         Spin            =   "BookPOChild05.frx":4570
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
      Begin TDBDate6Ctl.TDBDate MhDateInput18 
         Height          =   330
         Left            =   6240
         TabIndex        =   89
         Top             =   7620
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   582
         Calendar        =   "BookPOChild05.frx":4598
         Caption         =   "BookPOChild05.frx":46B0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild05.frx":471C
         Keys            =   "BookPOChild05.frx":473A
         Spin            =   "BookPOChild05.frx":4798
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
         Index           =   4
         Left            =   4080
         TabIndex        =   90
         Top             =   7620
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
         Caption         =   " PDF Send Date"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild05.frx":47C0
         Picture         =   "BookPOChild05.frx":47DC
      End
      Begin MSForms.ComboBox Combo31 
         Height          =   330
         Left            =   4320
         TabIndex        =   17
         Top             =   2220
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
      Begin MSForms.ComboBox Combo21 
         Height          =   330
         Left            =   4320
         TabIndex        =   7
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
      Begin MSForms.ComboBox Combo3 
         Height          =   330
         Left            =   6840
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
      Begin VB.Line Line5 
         X1              =   0
         X2              =   8055
         Y1              =   6045
         Y2              =   6045
      End
      Begin VB.Line Line4 
         X1              =   0
         X2              =   8055
         Y1              =   6570
         Y2              =   6570
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   8055
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   8055
         Y1              =   4500
         Y2              =   4500
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   330
         Left            =   4320
         TabIndex        =   12
         Top             =   1905
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
         X2              =   8055
         Y1              =   3645
         Y2              =   3645
      End
      Begin MSForms.ComboBox Combo1 
         Height          =   330
         Left            =   1320
         TabIndex        =   37
         Top             =   1590
         Width           =   1575
         VariousPropertyBits=   545282073
         BackColor       =   16777215
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "2778;582"
         MatchEntry      =   0
         ShowDropButtonWhen=   1
         SpecialEffect   =   0
         FontName        =   "Calibri"
         FontEffects     =   1073750016
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
End
Attribute VB_Name = "FrmBookPOChild05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public rstBookPOChild05 As New ADODB.Recordset
Dim rstPaperList As New ADODB.Recordset
Dim rstRefList As New ADODB.Recordset
Dim rstPrinterRates As New ADODB.Recordset
Dim rstPurchaseOrder As New ADODB.Recordset
Public PrinterCode As String
Dim BookCode As String
Dim SizeCode As String
Dim RefCode As String
Dim PaperCode As String
Dim PurchaseOrder As String
Private Sub Combo31_Validate(Cancel As Boolean)
If Combo1.Text = "2 Color" Or Combo1.Text = "4 Color" Then
    If Combo31.Text = "2nd Print" Or Combo31.Text = "3rd Print" Then
       MhRealInput4.Text = "0.00"
       MhRealInput7.Text = "0.00"
    End If
    fpSpread1.SetText 14, fpSpread1.ActiveRow, Val(MhRealInput4.Text)
    CalculateAmount
End If
End Sub
Private Sub Form_Load()
    Dim Cnt As Integer, Pages As Variant
    On Error GoTo ErrorHandler
    CenterForm Me
    BusySystemIndicator True
    DisableCloseButton Me
    For Cnt = 11 To 24
        fpSpread1.Col = Cnt
        fpSpread1.ColHidden = True
    Next
    
    AbortPO = False
    BookCode = FrmBookPrintOrder.rstBookList.Fields("Code").Value
    SizeCode = FrmBookPrintOrder.rstBookList.Fields("SizeCode").Value
        
    Text5.Text = Trim(FrmBookPrintOrder.Text2.Text)
    Text4.Text = Trim(FrmBookPrintOrder.Text5.Text)
    Text2.Text = Trim(FrmBookPrintOrder.Text3.Text)
    
    Combo1.AddItem "1 Color", 0
    Combo1.AddItem "2 Color", 1
    Combo1.AddItem "4 Color", 2
    
    Combo2.AddItem "Deepatch", 0
    Combo2.AddItem "PS", 1
    Combo2.AddItem "Wipeon", 2
    Combo2.AddItem "CTP", 3
    
    Combo3.AddItem "Old", 0
    Combo3.AddItem "New", 1
    Combo3.AddItem "Revised", 2
    
    Combo31.AddItem "Ist Print", 0
    Combo31.AddItem "2nd Print", 1
    Combo31.AddItem "3rd Print", 2
    
    Combo21.AddItem "New", 0
    Combo21.AddItem "L/C", 1
    Combo21.AddItem "N/C", 2
    Combo21.AddItem "Pending", 3
    
    ClearFields
    rstPaperList.Open "Select Name As Col0, Code From PaperMaster Where PaperMaster.Type = '1' Order By Name", CxnDatabase, adOpenKeyset, adLockReadOnly
    rstPaperList.ActiveConnection = Nothing
    
    
'    Dim aa As String
'    aa = "Select TRIM(Name) As Col0, Code From PaperPOParent Where PaperPOParent.OrderType = '1' Order By Code"
'
    rstPurchaseOrder.Open "Select TRIM(Name) As Col0, Code From PaperPOParent Where PaperPOParent.OrderType = '1' Order By Code", CxnDatabase, adOpenKeyset, adLockReadOnly
    rstPurchaseOrder.ActiveConnection = Nothing
    
    
    Call LoadRefList(BookCode, CheckNull(rstBookPOChild05.Fields("Code").Value))
    If Val(CheckNull(rstBookPOChild05.Fields("ActualQuantity").Value)) = 0 Then
        For Cnt = 1 To fpSpread1.MaxRows
            fpSpread1.SetText 1, Cnt, Val(FrmBookPrintOrder.rstBookList.Fields(IIf(Cnt = 1, "One", IIf(Cnt = 2, "Two", "Four")) & "ColorPages").Value)
            fpSpread1.SetText 2, Cnt, Val(FrmBookPrintOrder.rstBookList.Fields(IIf(Cnt = 1, "One", IIf(Cnt = 2, "Two", "Four")) & "ColorForms").Value)
            fpSpread1.SetText 3, Cnt, Val(FrmBookPrintOrder.rstBookList.Fields(IIf(Cnt = 1, "One", IIf(Cnt = 2, "Two", "Four")) & "Color¼Forms").Value)
            fpSpread1.SetText 4, Cnt, Val(FrmBookPrintOrder.rstBookList.Fields(IIf(Cnt = 1, "One", IIf(Cnt = 2, "Two", "Four")) & "Color½Forms").Value)
            fpSpread1.SetText 5, Cnt, Val(FrmBookPrintOrder.rstBookList.Fields(IIf(Cnt = 1, "One", IIf(Cnt = 2, "Two", "Four")) & "Color1F/BForms").Value) + Val(FrmBookPrintOrder.rstBookList.Fields(IIf(Cnt = 1, "One", IIf(Cnt = 2, "Two", "Four")) & "Color1W/TForms").Value)
            fpSpread1.SetText 6, Cnt, IIf(FrmBookPrintOrder.rstBookList.Fields(IIf(Cnt = 1, "One", IIf(Cnt = 2, "Two", "Four")) & "ColorPlateType").Value = "1", "Deepatch", IIf(FrmBookPrintOrder.rstBookList.Fields(IIf(Cnt = 1, "One", IIf(Cnt = 2, "Two", "Four")) & "ColorPlateType").Value = "2", "PS", IIf(FrmBookPrintOrder.rstBookList.Fields(IIf(Cnt = 1, "One", IIf(Cnt = 2, "Two", "Four")) & "ColorPlateType").Value = "3", "Wipeon", "CTP")))
          
            fpSpread1.SetText 7, Cnt, 0#
            fpSpread1.SetText 8, Cnt, 0#
            fpSpread1.SetText 9, Cnt, 0#
            fpSpread1.SetText 10, Cnt, 0#
        Next
        MhDateInput1.Text = Format(GetDate(FrmBookPrintOrder.MhDateInput1.Text), "dd-MM-yyyy")
        
        'MhDateInput3.Text = Format(DateAdd("d", 15, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
        
    Else
        LoadFields
    End If
    
    For Cnt = 1 To fpSpread1.MaxRows
        fpSpread1.GetText 1, Cnt, Pages
        If Val(Pages) > 0 Then
            fpSpread1.SetActiveCell 1, Cnt
            fpSpread1_DblClick 1, Cnt
            Exit For
        End If
    Next
    
       
    
    BusySystemIndicator False
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    Call CloseForm(Me)
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
       Call CloseForm(Me)
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CloseRecordset(rstPaperList)
    Call CloseRecordset(rstRefList)
    Call CloseRecordset(rstPrinterRates)
    Call CloseRecordset(rstPurchaseOrder)
    
End Sub

Private Sub Combo21_Validate(Cancel As Boolean)
    If CheckEmpty(Combo21, True) Then
        Cancel = True
    End If
End Sub
Private Sub ClearFields()
    MhDateInput1.Text = Format(Date, "dd-MM-yyyy")  'Order Date
    MhDateInput3.Text = Format(DateAdd("d", 15, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")    'Target Date
    Combo3.ListIndex = 0                 'Processing
    Combo3.ListIndex = 0
    Combo31.ListIndex = -1
    Text3.Text = ""                            'Ref.No.
    MhRealInput1.Text = "0"            'Actual Quantity
    MhRealInput2.Text = "0"            'Billing Quantity (Single Color)
    MhRealInput19.Text = "0"          'Billing Quantity (Double & Four)
    Combo1.ListIndex = 0                 'Printing Type
    MhRealInput15.Text = "0"          'Pages
    MhRealInput17.Text = "0"          'Qtr Form
    MhRealInput20.Text = "0"          'Half Form
    MhRealInput21.Text = "0"          'Full Form
    Combo2.ListIndex = 0                 'Plate Type
    MhRealInput22.Text = "1.00"     'Forms/Sheet For Printing Purpose
    MhRealInput3.Text = "0"            'Total Plates-¼F
    MhRealInput23.Text = "0"          'Total Plates-½F
    MhRealInput24.Text = "0"          'Total Plates-1F
    MhRealInput4.Text = "0.00"       'Plate Rate
    MhRealInput7.Text = "0.00"       'Plate Amount
    MhRealInput6.Text = "0"            'Total Forms-¼F
    MhRealInput25.Text = "0"          'Total Forms-½F
    MhRealInput26.Text = "0"          'Total Forms-1F
    MhRealInput5.Text = "0.00"       'Print Rate
    MhRealInput8.Text = "0.00"       'Print Amount
    MhRealInput14.Text = "0.00"     'VAT %
    MhRealInput18.Text = "0.00"     'VAT Amount
    MhRealInput9.Text = "0.00"       'Adjustment
    MhRealInput10.Text = "0.00"     'Total Amount
    Text1.Text = ""                            'Paper Name
    MhRealInput27.Text = "1.00"     'Forms/Sheet For Paper Purpose
    MhRealInput11.Text = "0.00"     'Paper Wastage (in %)
    MhRealInput12.Text = "0.000"   'Paper Consumption
    MhRealInput13.Text = "0.000"   'Total Paper Consumption
    Text8.Text = ""                            'Bill No.
    MhDateInput2.Text = "  -  -    "    'Bill Date
    MhRealInput16.Text = "0.00"      'Bill Amount
    MhRealInput101.Text = "0.000"      'Unit Cost
    Text6.Text = ""                             'Remarks
    TxtAdNar.Text = ""
    TxtPOrder.Text = ""
    
    MhDateInput17.Text = "  -  -    "    'PDF Send To Production Date
    MhDateInput18.Text = "  -  -    "    'PDF Send To Production Date
    
    RefCode = ""
    PurchaseOrder = ""
    fpSpread1.ClearRange 1, 1, fpSpread1.MaxCols, fpSpread1.MaxRows, True
End Sub
Private Sub LoadFields()
    Dim Cnt As Integer
    If rstBookPOChild05.RecordCount = 0 Then Exit Sub
    MhDateInput1.Text = Format(rstBookPOChild05.Fields("OrderDate").Value, "dd-MM-yyyy")
    MhDateInput3.Text = Format(rstBookPOChild05.Fields("TargetDate").Value, "dd-MM-yyyy")
    If Not IsNull(rstBookPOChild05.Fields("ExtendDate").Value) Then
        MhDateInput31.Text = Format(rstBookPOChild05.Fields("ExtendDate").Value, "dd-MM-yyyy")
    End If
    
    If rstBookPOChild05.Fields("PlateMaking").Value <> "" Then
       Combo31.Text = Trim(rstBookPOChild05.Fields("PlateMaking").Value)
    End If
    
    Combo3.ListIndex = IIf(rstBookPOChild05.Fields("Processing").Value = "O", 0, IIf(rstBookPOChild05.Fields("Processing").Value = "N", 1, 2))
    RefCode = rstBookPOChild05.Fields("Ref").Value
    
    If rstRefList.RecordCount > 0 Then rstRefList.MoveFirst
    
    rstRefList.Find "[Code] = '" & RefCode & "'"
    If Not rstRefList.EOF Then
        Text3.Text = Trim(rstRefList.Fields("Name").Value)
    End If
    
    If Not IsNull(rstBookPOChild05.Fields("POCode").Value) Then
       PurchaseOrder = rstBookPOChild05.Fields("POCode").Value
        If rstPurchaseOrder.RecordCount > 0 Then rstPurchaseOrder.MoveFirst
        rstPurchaseOrder.Find "[Code] = '" & PurchaseOrder & "'"
        If Not rstPurchaseOrder.EOF Then
            TxtPOrder.Text = Trim(rstPurchaseOrder.Fields("Col0").Value)
        End If
    End If
     
    MhRealInput1.Text = Format(Val(rstBookPOChild05.Fields("ActualQuantity").Value), "0")
    MhRealInput2.Text = Format(Val(rstBookPOChild05.Fields("BillingQuantity01").Value), "0")
    MhRealInput19.Text = Format(Val(rstBookPOChild05.Fields("BillingQuantity02").Value), "0")
    
    For Cnt = 1 To fpSpread1.MaxRows
        fpSpread1.SetText 1, Cnt, Val(rstBookPOChild05.Fields("Pages" & IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value)
        fpSpread1.SetText 2, Cnt, Val(rstBookPOChild05.Fields("Forms" & IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value)
        fpSpread1.SetText 3, Cnt, Val(rstBookPOChild05.Fields("Forms" & IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4")) & "-¼").Value)
        fpSpread1.SetText 4, Cnt, Val(rstBookPOChild05.Fields("Forms" & IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4")) & "-½").Value)
        fpSpread1.SetText 5, Cnt, Val(rstBookPOChild05.Fields("Forms" & IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4")) & "-1").Value)
        fpSpread1.SetText 6, Cnt, IIf(rstBookPOChild05.Fields("PlateType" & IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value = "1", "Deepatch", IIf(rstBookPOChild05.Fields("PlateType" & IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value = "2", "PS", IIf(rstBookPOChild05.Fields("PlateType" & IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value = "3", "Wipeon", "CTP")))
        fpSpread1.SetText 7, Cnt, Val(rstBookPOChild05.Fields("PlateAmount" & IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value)
        fpSpread1.SetText 8, Cnt, Val(rstBookPOChild05.Fields("PrintAmount" & IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value)
        fpSpread1.SetText 9, Cnt, Val(rstBookPOChild05.Fields("PaperWastage" & IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4")) & "%").Value)
        fpSpread1.SetText 10, Cnt, Val(rstBookPOChild05.Fields("PaperConsumptionOther" & IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value)
        fpSpread1.SetText 11, Cnt, Val(rstBookPOChild05.Fields("TotalPlates" & IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4")) & "-¼").Value)
        fpSpread1.SetText 12, Cnt, Val(rstBookPOChild05.Fields("TotalPlates" & IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4")) & "-½").Value)
        fpSpread1.SetText 13, Cnt, Val(rstBookPOChild05.Fields("TotalPlates" & IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4")) & "-1").Value)
        fpSpread1.SetText 14, Cnt, Val(rstBookPOChild05.Fields("PlateRate" & IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value)
        fpSpread1.SetText 15, Cnt, Val(rstBookPOChild05.Fields("TotalForms" & IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4")) & "-¼").Value)
        fpSpread1.SetText 16, Cnt, Val(rstBookPOChild05.Fields("TotalForms" & IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4")) & "-½").Value)
        fpSpread1.SetText 17, Cnt, Val(rstBookPOChild05.Fields("TotalForms" & IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4")) & "-1").Value)
        fpSpread1.SetText 18, Cnt, Val(rstBookPOChild05.Fields("PrintRate" & IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value)
        If rstPaperList.RecordCount > 0 Then rstPaperList.MoveFirst
        rstPaperList.Find "[Code] = '" & rstBookPOChild05.Fields("Paper" & IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value & "'"
        If Not rstPaperList.EOF Then
            fpSpread1.SetText 19, Cnt, rstPaperList.Fields("Col0").Value
        End If
        fpSpread1.SetText 20, Cnt, rstBookPOChild05.Fields("Paper" & IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value
        fpSpread1.SetText 21, Cnt, Val(rstBookPOChild05.Fields("Forms/Sheet1-" & IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value)
        fpSpread1.SetText 22, Cnt, Val(rstBookPOChild05.Fields("Forms/Sheet2-" & IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value)
    Next
    MhRealInput14.Text = Format(Val(rstBookPOChild05.Fields("VAT%").Value), "0.00")
    MhRealInput18.Text = Format(Val(rstBookPOChild05.Fields("VAT").Value), "0.00")
    MhRealInput9.Text = Format(Val(rstBookPOChild05.Fields("Adjustment").Value), "0.00")
    MhRealInput10.Text = Format(Val(rstBookPOChild05.Fields("BillAmount").Value), "0.00")
    MhRealInput13.Text = Format(Val(rstBookPOChild05.Fields("TotalPaperConsumption").Value), "0.000")
    Text8.Text = rstBookPOChild05.Fields("BillNo").Value
    
    If IsNull(rstBookPOChild05.Fields("UnitCost").Value) Or rstBookPOChild05.Fields("UnitCost").Value = "0" Then
       MhRealInput101.Text = Format(Val(MhRealInput10.Text) / Val(MhRealInput1.Text), "0.000")   'Unit Cost
    Else
        MhRealInput101.Text = Format(Val(rstBookPOChild05.Fields("UnitCost").Value), "0.000")
    End If
    
    If rstBookPOChild05.Fields("BookStatus").Value <> "" Then
        Combo21.Text = rstBookPOChild05.Fields("BookStatus").Value
    End If
    
    If Not IsNull(rstBookPOChild05.Fields("BillDate").Value) Then
        MhDateInput2.Text = Format(rstBookPOChild05.Fields("BillDate").Value, "dd-MM-yyyy")
    End If
    MhRealInput16.Text = Format(Val(rstBookPOChild05.Fields("PaidAmount").Value), "0.00")
    Text6.Text = rstBookPOChild05.Fields("Remarks").Value
    TxtAdNar.Text = rstBookPOChild05.Fields("AdjustmentRemarks").Value
    
    If Not IsNull(rstBookPOChild05.Fields("PDFSendToProduction").Value) Then
        MhDateInput17.Text = Format(rstBookPOChild05.Fields("PDFSendToProduction").Value, "dd-MM-yyyy")
    End If
    
    If Not IsNull(rstBookPOChild05.Fields("PDFSendToPrinter").Value) Then
        MhDateInput18.Text = Format(rstBookPOChild05.Fields("PDFSendToPrinter").Value, "dd-MM-yyyy")
    End If
End Sub
Private Sub SaveFields()
    
    Dim Cnt As Integer, Content As Variant
    rstBookPOChild05.Fields("OrderDate").Value = GetDate(MhDateInput1.Text)
    rstBookPOChild05.Fields("TargetDate").Value = GetDate(MhDateInput3.Text)
     
    If Not IsDate(MhDateInput31.Text) Then
         rstBookPOChild05.Fields("ExtendDate").Value = Null
    Else
         rstBookPOChild05.Fields("ExtendDate").Value = GetDate(MhDateInput31.Text)
    End If
    
    rstBookPOChild05.Fields("Processing").Value = IIf(Combo3.ListIndex = 0, "O", IIf(Combo3.ListIndex = 1, "N", "R"))
    rstBookPOChild05.Fields("BookStatus").Value = Combo21.Text
    rstBookPOChild05.Fields("PlateMaking").Value = Combo31.Text
    
    
    rstBookPOChild05.Fields("Ref").Value = RefCode
    rstBookPOChild05.Fields("POCode").Value = PurchaseOrder
    
    rstBookPOChild05.Fields("ActualQuantity").Value = Val(MhRealInput1.Text)
    rstBookPOChild05.Fields("BillingQuantity01").Value = Val(MhRealInput2.Text)
    rstBookPOChild05.Fields("BillingQuantity02").Value = Val(MhRealInput19.Text)
    For Cnt = 1 To fpSpread1.MaxRows
        fpSpread1.GetText 1, Cnt, Content
        rstBookPOChild05.Fields("Pages" + IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value = Val(Content)
        fpSpread1.GetText 2, Cnt, Content
        rstBookPOChild05.Fields("Forms" + IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value = Val(Content)
        fpSpread1.GetText 3, Cnt, Content
        rstBookPOChild05.Fields("Forms" + IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4")) & "-¼").Value = Val(Content)
        fpSpread1.GetText 4, Cnt, Content
        rstBookPOChild05.Fields("Forms" + IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4")) & "-½").Value = Val(Content)
        fpSpread1.GetText 5, Cnt, Content
        rstBookPOChild05.Fields("Forms" + IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4")) & "-1").Value = Val(Content)
        fpSpread1.GetText 6, Cnt, Content
        rstBookPOChild05.Fields("PlateType" + IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value = IIf(Content = "Deepatch", "1", IIf(Content = "PS", "2", IIf(Content = "Wipeon", "3", "4")))
        fpSpread1.GetText 7, Cnt, Content
        
        If Combo1.Text = "2 Color" Or Combo1.Text = "4 Color" Then
            If Combo31.Text = "2nd Print" Or Combo31.Text = "3rd Print" Then
                rstBookPOChild05.Fields("PlateAmount" + IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value = Val(Content) ' 0#
            Else
              rstBookPOChild05.Fields("PlateAmount" + IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value = Val(Content)
            End If
        Else
            rstBookPOChild05.Fields("PlateAmount" + IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value = Val(Content)
        End If
        
        fpSpread1.GetText 8, Cnt, Content
        rstBookPOChild05.Fields("PrintAmount" + IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value = Val(Content)
        fpSpread1.GetText 9, Cnt, Content
        rstBookPOChild05.Fields("PaperWastage" + IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4")) & "%").Value = Val(Content)
        fpSpread1.GetText 10, Cnt, Content
        rstBookPOChild05.Fields("PaperConsumptionOther" + IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value = Val(Content)
        rstBookPOChild05.Fields("PaperConsumptionSheets" + IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value = Int(Val(Content)) * 500 + (Val(Content) - Int(Val(Content))) * 1000
        fpSpread1.GetText 11, Cnt, Content
        rstBookPOChild05.Fields("TotalPlates" + IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4")) & "-¼").Value = Val(Content)
        fpSpread1.GetText 12, Cnt, Content
        rstBookPOChild05.Fields("TotalPlates" + IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4")) & "-½").Value = Val(Content)
        fpSpread1.GetText 13, Cnt, Content
        rstBookPOChild05.Fields("TotalPlates" + IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4")) & "-1").Value = Val(Content)
        fpSpread1.GetText 14, Cnt, Content
        
        If Combo1.Text = "2 Color" Or Combo1.Text = "4 Color" Then
            If Combo31.Text = "2nd Print" Or Combo31.Text = "3rd Print" Then
                rstBookPOChild05.Fields("PlateRate" + IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value = Val(Content) ' 0#
            Else
              rstBookPOChild05.Fields("PlateRate" + IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value = Val(Content)
            End If
        Else
          rstBookPOChild05.Fields("PlateRate" + IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value = Val(Content)
        End If
               
        fpSpread1.GetText 15, Cnt, Content
        rstBookPOChild05.Fields("TotalForms" + IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4")) & "-¼").Value = Val(Content)
        fpSpread1.GetText 16, Cnt, Content
        rstBookPOChild05.Fields("TotalForms" + IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4")) & "-½").Value = Val(Content)
        fpSpread1.GetText 17, Cnt, Content
        rstBookPOChild05.Fields("TotalForms" + IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4")) & "-1").Value = Val(Content)
        fpSpread1.GetText 18, Cnt, Content
        rstBookPOChild05.Fields("PrintRate" + IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value = Val(Content)
        fpSpread1.GetText 20, Cnt, Content
        rstBookPOChild05.Fields("Paper" + IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value = Content
        fpSpread1.GetText 21, Cnt, Content
        rstBookPOChild05.Fields("Forms/Sheet1-" + IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value = Val(Content)
        fpSpread1.GetText 22, Cnt, Content
        rstBookPOChild05.Fields("Forms/Sheet2-" + IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4"))).Value = Val(Content)
    Next
    rstBookPOChild05.Fields("VAT%").Value = Format(Val(MhRealInput14.Text), "0.00")
    rstBookPOChild05.Fields("VAT").Value = Format(Val(MhRealInput18.Text), "0.00")
    rstBookPOChild05.Fields("Adjustment").Value = Format(Val(MhRealInput9.Text), "0.00")
    rstBookPOChild05.Fields("BillAmount").Value = Format(Val(MhRealInput10.Text), "0.00")
    rstBookPOChild05.Fields("TotalPaperConsumption").Value = Format(Val(MhRealInput13.Text), "0.000")
    rstBookPOChild05.Fields("BillNo").Value = Text8.Text
    
    If Not IsDate(MhDateInput2.Text) Then
         rstBookPOChild05.Fields("BillDate").Value = Null
    Else
         rstBookPOChild05.Fields("BillDate").Value = GetDate(MhDateInput2.Text)
    End If
    
    If Not IsDate(MhDateInput17.Text) Then
         rstBookPOChild05.Fields("PDFSendToProduction").Value = Null
    Else
         rstBookPOChild05.Fields("PDFSendToProduction").Value = GetDate(MhDateInput17.Text)
    End If
    
    If Not IsDate(MhDateInput18.Text) Then
         rstBookPOChild05.Fields("PDFSendToPrinter").Value = Null
    Else
         rstBookPOChild05.Fields("PDFSendToPrinter").Value = GetDate(MhDateInput18.Text)
    End If
    
    rstBookPOChild05.Fields("PaidAmount").Value = Format(Val(MhRealInput16.Text), "0.00")
    rstBookPOChild05.Fields("Remarks").Value = Text6.Text
    rstBookPOChild05.Fields("AdjustmentRemarks").Value = IIf(Val(MhRealInput9.Text) <> 0, TxtAdNar.Text, "")
    If Not CheckEmpty(Text8.Text, False) Then If IsNull(rstBookPOChild05.Fields("BillFeedDate").Value) Then rstBookPOChild05.Fields("BillFeedDate").Value = Now()
    Dim lpBuff As String * 1024
    GetComputerName lpBuff, Len(lpBuff)
    If Not CheckEmpty(Text8.Text, False) Then If IsNull(rstBookPOChild05.Fields("ComputerName").Value) Then rstBookPOChild05.Fields("ComputerName").Value = Left(lpBuff, (InStr(1, lpBuff, vbNullChar)) - 1)
    rstBookPOChild05.Fields("UnitCost").Value = Format(Val(MhRealInput101.Text), "0.000")
    
    
End Sub
Private Sub MhDateInput1_Validate(Cancel As Boolean)
    If Not IsDate(GetDate(MhDateInput1.Text)) Then
        Cancel = True
    ElseIf Format(GetDate(MhDateInput1.Text), "yyyymmdd") < Format(FinancialYearFrom, "yyyymmdd") Or Format(GetDate(MhDateInput1.Text), "yyyymmdd") > Format(FinancialYearTo, "yyyymmdd") Then
        Cancel = True
    ElseIf Val(CheckNull(rstBookPOChild05.Fields("ActualQuantity").Value)) = 0 Then
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

Private Sub Text3_Change()
    If Text3.Text = " " Then
        Text3.Text = "?"
        SendKeys "{TAB}"
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
            SendKeys "{TAB}"
        End If
        Cancel = True
    Else
        RefCode = rstRefList.Fields("Code").Value
        Text3.Text = Trim(rstRefList.Fields("Name").Value)
        If Val(CheckNull(rstBookPOChild05.Fields("ActualQuantity").Value)) = 0 Then
            MhRealInput1.Text = Abs(Format(Val(rstRefList.Fields("BalanceQuantity").Value), "0"))
            CalculateAQD
        End If
    End If
End Sub
Private Sub MhRealInput1_Validate(Cancel As Boolean)
    If Val(MhRealInput1.Text) <= 0 Then
        Cancel = True
    Else
        
        CalculateAQD
        If MhRealInput15.Value > 0 And MhRealInput1.Value > 0 Then
            'Set Target Date based on Quantity and Form
            If MhRealInput15.Value <= 120 And MhRealInput1.Value <= 3300 Then
               MhDateInput3.Text = Format(DateAdd("d", 10, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
            ElseIf MhRealInput15.Value <= 120 And MhRealInput1.Value > 3300 And MhRealInput1.Value <= 5500 Then
               MhDateInput3.Text = Format(DateAdd("d", 12, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
            ElseIf MhRealInput15.Value <= 120 And MhRealInput1.Value > 5500 And MhRealInput1.Value <= 10500 Then
               MhDateInput3.Text = Format(DateAdd("d", 15, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
            ElseIf MhRealInput15.Value <= 120 And MhRealInput1.Value > 10500 Then
               MhDateInput3.Text = Format(DateAdd("d", 20, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
            
            ElseIf MhRealInput15.Value > 121 And MhRealInput15.Value <= 200 And MhRealInput1.Value <= 3300 Then
               MhDateInput3.Text = Format(DateAdd("d", 10, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
            ElseIf MhRealInput15.Value > 121 And MhRealInput15.Value <= 200 And MhRealInput1.Value > 3300 And MhRealInput1.Value <= 5500 Then
               MhDateInput3.Text = Format(DateAdd("d", 12, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
            ElseIf MhRealInput15.Value > 121 And MhRealInput15.Value <= 200 And MhRealInput1.Value > 5500 And MhRealInput1.Value <= 10500 Then
               MhDateInput3.Text = Format(DateAdd("d", 18, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
            ElseIf MhRealInput15.Value > 121 And MhRealInput15.Value <= 200 And MhRealInput1.Value > 10500 Then
               MhDateInput3.Text = Format(DateAdd("d", 20, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
            ElseIf MhRealInput15.Value > 200 And MhRealInput1.Value <= 3300 Then
               MhDateInput3.Text = Format(DateAdd("d", 10, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
            ElseIf MhRealInput15.Value > 200 And MhRealInput1.Value > 3300 And MhRealInput1.Value < 5500 Then
               MhDateInput3.Text = Format(DateAdd("d", 15, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
            ElseIf MhRealInput15.Value > 200 And MhRealInput1.Value > 5500 And MhRealInput1.Value < 10500 Then
               MhDateInput3.Text = Format(DateAdd("d", 20, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
            ElseIf MhRealInput15.Value > 200 And MhRealInput1.Value > 10500 Then
               MhDateInput3.Text = Format(DateAdd("d", 25, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
            End If
        End If
    End If
End Sub
Private Sub MhRealInput2_Validate(Cancel As Boolean)
    If Val(MhRealInput2.Text) <= 0 Then
        Cancel = True
    ElseIf Val(MhRealInput2.Text) Mod 1000 <> 0 Then
        MhRealInput2.SetFocus
        Cancel = True
    Else
        CalculateBQD ("S")
    End If
End Sub
Private Sub MhRealInput19_Validate(Cancel As Boolean)
    If Val(MhRealInput2.Text) <= 0 Then
        Cancel = True
    ElseIf Val(MhRealInput19.Text) Mod 1000 <> 0 Then
        MhRealInput19.SetFocus
        Cancel = True
    Else
        CalculateBQD ("O")
    End If
End Sub
Private Sub MhRealInput15_Validate(Cancel As Boolean)
    fpSpread1.SetText 1, fpSpread1.ActiveRow, Val(MhRealInput15.Text)
    fpSpread1.SetText 2, fpSpread1.ActiveRow, Val(MhRealInput15.Text) / IIf(FrmBookPrintOrder.rstBookList.Fields("FormType").Value = "1", "08", IIf(FrmBookPrintOrder.rstBookList.Fields("FormType").Value = "2", "16", IIf(FrmBookPrintOrder.rstBookList.Fields("FormType").Value = "3", "04", IIf(FrmBookPrintOrder.rstBookList.Fields("FormType").Value = "4", "12", IIf(FrmBookPrintOrder.rstBookList.Fields("FormType").Value = "5", "24", IIf(FrmBookPrintOrder.rstBookList.Fields("FormType").Value = "6", "32", IIf(FrmBookPrintOrder.rstBookList.Fields("FormType").Value = "7", "64", "06")))))))
    
    If MhRealInput15.Value > 0 And MhRealInput1.Value > 0 Then
        'Set Target Date based on Quantity and Form
        If MhRealInput15.Value <= 120 And MhRealInput1.Value <= 3300 Then
           MhDateInput3.Text = Format(DateAdd("d", 10, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
        ElseIf MhRealInput15.Value <= 120 And MhRealInput1.Value > 3300 And MhRealInput1.Value <= 5500 Then
           MhDateInput3.Text = Format(DateAdd("d", 12, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
        ElseIf MhRealInput15.Value <= 120 And MhRealInput1.Value > 5500 And MhRealInput1.Value <= 10500 Then
           MhDateInput3.Text = Format(DateAdd("d", 15, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
        ElseIf MhRealInput15.Value <= 120 And MhRealInput1.Value > 10500 Then
           MhDateInput3.Text = Format(DateAdd("d", 20, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
        
        ElseIf MhRealInput15.Value > 121 And MhRealInput15.Value <= 200 And MhRealInput1.Value <= 3300 Then
           MhDateInput3.Text = Format(DateAdd("d", 10, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
        ElseIf MhRealInput15.Value > 121 And MhRealInput15.Value <= 200 And MhRealInput1.Value > 3300 And MhRealInput1.Value <= 5500 Then
           MhDateInput3.Text = Format(DateAdd("d", 12, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
        ElseIf MhRealInput15.Value > 121 And MhRealInput15.Value <= 200 And MhRealInput1.Value > 5500 And MhRealInput1.Value <= 10500 Then
           MhDateInput3.Text = Format(DateAdd("d", 18, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
        ElseIf MhRealInput15.Value > 121 And MhRealInput15.Value <= 200 And MhRealInput1.Value > 10500 Then
           MhDateInput3.Text = Format(DateAdd("d", 20, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
        ElseIf MhRealInput15.Value > 200 And MhRealInput1.Value <= 3300 Then
           MhDateInput3.Text = Format(DateAdd("d", 10, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
        ElseIf MhRealInput15.Value > 200 And MhRealInput1.Value > 3300 And MhRealInput1.Value < 5500 Then
           MhDateInput3.Text = Format(DateAdd("d", 15, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
        ElseIf MhRealInput15.Value > 200 And MhRealInput1.Value > 5500 And MhRealInput1.Value < 10500 Then
           MhDateInput3.Text = Format(DateAdd("d", 20, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
        ElseIf MhRealInput15.Value > 200 And MhRealInput1.Value > 10500 Then
           MhDateInput3.Text = Format(DateAdd("d", 25, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
        End If
    End If
End Sub
Private Sub MhRealInput17_Validate(Cancel As Boolean)   '¼ Forms
    Dim Forms¼ As Variant, Forms½ As Variant, Forms1 As Variant

    fpSpread1.SetText 3, fpSpread1.ActiveRow, Val(MhRealInput17.Text)
    Call CalculateTotalPlates(IIf(fpSpread1.ActiveRow = 1, "1", IIf(fpSpread1.ActiveRow = 2, "2", "4")), Val(MhRealInput17.Text), "¼")
    Call CalculateTotalForms(IIf(fpSpread1.ActiveRow = 1, "1", IIf(fpSpread1.ActiveRow = 2, "2", "4")), Val(MhRealInput17.Text), "¼")
    CalculateAmount
    Call CalculateConsumption(IIf(fpSpread1.ActiveRow = 1, "1", IIf(fpSpread1.ActiveRow = 2, "2", "4")))
    fpSpread1.GetText 3, fpSpread1.ActiveRow, Forms¼
    fpSpread1.GetText 4, fpSpread1.ActiveRow, Forms½
    fpSpread1.GetText 5, fpSpread1.ActiveRow, Forms1
    fpSpread1.SetText 2, fpSpread1.ActiveRow, Val(Forms¼) * 0.25 + Val(Forms½) * 0.5 + Val(Forms1) * 1
End Sub
Private Sub MhRealInput20_Validate(Cancel As Boolean)   '½ Forms
    Dim Forms¼ As Variant, Forms½ As Variant, Forms1 As Variant
    fpSpread1.SetText 4, fpSpread1.ActiveRow, Val(MhRealInput20.Text)
    Call CalculateTotalPlates(IIf(fpSpread1.ActiveRow = 1, "1", IIf(fpSpread1.ActiveRow = 2, "2", "4")), Val(MhRealInput20.Text), "½")
    Call CalculateTotalForms(IIf(fpSpread1.ActiveRow = 1, "1", IIf(fpSpread1.ActiveRow = 2, "2", "4")), Val(MhRealInput20.Text), "½")
    CalculateAmount
    Call CalculateConsumption(IIf(fpSpread1.ActiveRow = 1, "1", IIf(fpSpread1.ActiveRow = 2, "2", "4")))
    fpSpread1.GetText 3, fpSpread1.ActiveRow, Forms¼
    fpSpread1.GetText 4, fpSpread1.ActiveRow, Forms½
    fpSpread1.GetText 5, fpSpread1.ActiveRow, Forms1
    fpSpread1.SetText 2, fpSpread1.ActiveRow, Val(Forms¼) * 0.25 + Val(Forms½) * 0.5 + Val(Forms1) * 1
End Sub
Private Sub MhRealInput21_Validate(Cancel As Boolean)   '1 Forms
    Dim Forms¼ As Variant, Forms½ As Variant, Forms1 As Variant

    fpSpread1.SetText 5, fpSpread1.ActiveRow, Val(MhRealInput21.Text)
    Call CalculateTotalPlates(IIf(fpSpread1.ActiveRow = 1, "1", IIf(fpSpread1.ActiveRow = 2, "2", "4")), Val(MhRealInput21.Text), "1")
    Call CalculateTotalForms(IIf(fpSpread1.ActiveRow = 1, "1", IIf(fpSpread1.ActiveRow = 2, "2", "4")), Val(MhRealInput21.Text), "1")
    CalculateAmount
    Call CalculateConsumption(IIf(fpSpread1.ActiveRow = 1, "1", IIf(fpSpread1.ActiveRow = 2, "2", "4")))
    fpSpread1.GetText 3, fpSpread1.ActiveRow, Forms¼
    fpSpread1.GetText 4, fpSpread1.ActiveRow, Forms½
    fpSpread1.GetText 5, fpSpread1.ActiveRow, Forms1
    fpSpread1.SetText 2, fpSpread1.ActiveRow, Val(Forms¼) * 0.25 + Val(Forms½) * 0.5 + Val(Forms1) * 1
End Sub
Private Sub Combo2_Validate(Cancel As Boolean)  'Plate Type
    fpSpread1.SetText 6, fpSpread1.ActiveRow, IIf(Combo2.ListIndex = 0, "Deepatch", IIf(Combo2.ListIndex = 1, "PS", IIf(Combo2.ListIndex = 2, "Wipeon", "CTP")))
    GetPrinterRates IIf(fpSpread1.ActiveRow = 1, "1", IIf(fpSpread1.ActiveRow = 2, "2", "4")), "L"  'Get Plate Rates
    CalculateAmount
    If Combo2.ListIndex = 1 Or Combo2.ListIndex = 3 Then    'PS/CTP Plate Details
        On Error Resume Next
        FrmPSPlateRegister.BookCode = BookCode
        FrmPSPlateRegister.BookName = Trim(Text2.Text)
        FrmPSPlateRegister.OrderCode = IIf(CheckNull(rstBookPOChild05.Fields("Code").Value) = "", "999999", rstBookPOChild05.Fields("Code").Value)
        FrmPSPlateRegister.OrderDate = GetDate(MhDateInput1.Text)
        FrmPSPlateRegister.OrderType = "05"
        FrmPSPlateRegister.PlateType = IIf(fpSpread1.ActiveRow = 1, "1", IIf(fpSpread1.ActiveRow = 2, "2", "4"))
        Load FrmPSPlateRegister
        If Err.Number <> 364 Then FrmPSPlateRegister.Show vbModal
        On Error GoTo 0
    End If
End Sub
Private Sub MhRealInput22_Validate(Cancel As Boolean)   'Forms/Sheet For Printing Purpose
    Dim Forms As Variant
    If Val(MhRealInput22.Text) <> 0.5 And Val(MhRealInput22.Text) <> 1 And Val(MhRealInput22.Text) <> 2 Then
        Cancel = True
    Else
        fpSpread1.SetText 21, fpSpread1.ActiveRow, Val(MhRealInput22.Text)
        fpSpread1.GetText 3, fpSpread1.ActiveRow, Forms   '¼ Forms
        Call CalculateTotalPlates(IIf(fpSpread1.ActiveRow = 1, "1", IIf(fpSpread1.ActiveRow = 2, "2", "4")), Val(Forms), "¼")
        Call CalculateTotalForms(IIf(fpSpread1.ActiveRow = 1, "1", IIf(fpSpread1.ActiveRow = 2, "2", "4")), Val(Forms), "¼")
        fpSpread1.GetText 4, fpSpread1.ActiveRow, Forms   '½ Forms
        Call CalculateTotalPlates(IIf(fpSpread1.ActiveRow = 1, "1", IIf(fpSpread1.ActiveRow = 2, "2", "4")), Val(Forms), "½")
        Call CalculateTotalForms(IIf(fpSpread1.ActiveRow = 1, "1", IIf(fpSpread1.ActiveRow = 2, "2", "4")), Val(Forms), "½")
        fpSpread1.GetText 5, fpSpread1.ActiveRow, Forms   '1 Forms
        Call CalculateTotalPlates(IIf(fpSpread1.ActiveRow = 1, "1", IIf(fpSpread1.ActiveRow = 2, "2", "4")), Val(Forms), "1")
        Call CalculateTotalForms(IIf(fpSpread1.ActiveRow = 1, "1", IIf(fpSpread1.ActiveRow = 2, "2", "4")), Val(Forms), "1")
        CalculateAmount
    End If
End Sub
Private Sub MhRealInput4_Validate(Cancel As Boolean)    'Plate Rate
    If Combo1.Text = "2 Color" Or Combo1.Text = "4 Color" Then
        If Combo31.Text = "2nd Print" Or Combo31.Text = "3rd Print" Then
           MhRealInput4.Text = "0.00"
           MhRealInput7.Text = "0.00"
        Else
           fpSpread1.SetText 14, fpSpread1.ActiveRow, Val(MhRealInput4.Text)
           CalculateAmount
        End If
    Else
      fpSpread1.SetText 14, fpSpread1.ActiveRow, Val(MhRealInput4.Text)
      CalculateAmount
    End If
    
End Sub
Private Sub MhRealInput5_Validate(Cancel As Boolean)    'Print Rate
    fpSpread1.SetText 18, fpSpread1.ActiveRow, Val(MhRealInput5.Text)
    CalculateAmount
End Sub
Private Sub MhRealInput14_Validate(Cancel As Boolean)   'VAT
    CalculateTotalAmount
End Sub
Private Sub MhRealInput9_Validate(Cancel As Boolean)    'Adjustment
    CalculateTotalAmount
End Sub
Private Sub Text1_Change()  'Paper
    If Text1.Text = " " Then
        Text1.Text = "?"
        SendKeys "{TAB}"
    ElseIf CheckEmpty(Text1, False) Then
        PaperCode = ""
        fpSpread1.SetText 19, fpSpread1.ActiveRow, ""
        fpSpread1.SetText 20, fpSpread1.ActiveRow, ""
    End If
End Sub
Private Sub Text1_Validate(Cancel As Boolean)
    Dim SearchString As String
    
    If Val(MhRealInput8.Text) = 0 Then
        If CheckEmpty(Text1, False) Then Exit Sub
    End If
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
            SendKeys "{TAB}"
        End If
        Cancel = True
    Else
        PaperCode = rstPaperList.Fields("Code").Value
        fpSpread1.SetText 19, fpSpread1.ActiveRow, Trim(Text1.Text)
        fpSpread1.SetText 20, fpSpread1.ActiveRow, PaperCode
    End If
End Sub
Private Sub MhRealInput27_Validate(Cancel As Boolean)   'Forms/Sheet For Paper Purpose
    If Val(MhRealInput27.Text) <> 0.5 And Val(MhRealInput27.Text) <> 1 And Val(MhRealInput27.Text) <> 2 Then
        Cancel = True
    Else
        fpSpread1.SetText 22, fpSpread1.ActiveRow, Val(MhRealInput27.Text)
        Call CalculateConsumption(IIf(fpSpread1.ActiveRow = 1, "1", IIf(fpSpread1.ActiveRow = 2, "2", "4")))
    End If
End Sub
Private Sub MhRealInput11_Validate(Cancel As Boolean)   'Paper Wastage Rate
    fpSpread1.SetText 9, fpSpread1.ActiveRow, Val(MhRealInput11.Text)
    Call CalculateConsumption(IIf(fpSpread1.ActiveRow = 1, "1", IIf(fpSpread1.ActiveRow = 2, "2", "4")))
End Sub
Private Sub fpSpread1_LeaveRow(ByVal Row As Long, ByVal RowWasLast As Boolean, ByVal RowChanged As Boolean, ByVal AllCellsHaveData As Boolean, ByVal NewRow As Long, ByVal NewRowIsLast As Long, Cancel As Boolean)
    fpSpread1.SetActiveCell 1, NewRow
    fpSpread1_DblClick 1, NewRow
End Sub
Private Sub fpSpread1_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim Content As Variant
    Combo1.ListIndex = IIf(Row = 1, 0, IIf(Row = 2, 1, 2))  'Printing Type
    fpSpread1.GetText 1, Row, Content   'Pages
    MhRealInput15.Text = Format(Val(Content), "0")
    fpSpread1.GetText 3, Row, Content   '¼ F
    MhRealInput17.Text = Format(Val(Content), "0")
    fpSpread1.GetText 4, Row, Content   '½ F
    MhRealInput20.Text = Format(Val(Content), "0")
    fpSpread1.GetText 5, Row, Content   '1 F
    MhRealInput21.Text = Format(Val(Content), "0")
    fpSpread1.GetText 6, Row, Content   'Plate Type
    Combo2.ListIndex = IIf(Content = "Deepatch", 0, IIf(Content = "PS", 1, IIf(Content = "Wipeon", 2, 3)))
    fpSpread1.GetText 7, Row, Content   'Plate Amount
    MhRealInput7.Text = Format(Val(Content), "0.00")
    fpSpread1.GetText 8, Row, Content   'Print Amount
    MhRealInput8.Text = Format(Val(Content), "0.00")
    fpSpread1.GetText 9, Row, Content   'Paper Wastage
    MhRealInput11.Text = Format(Val(Content), "0.00")
    fpSpread1.GetText 10, Row, Content   'Paper Consumption (Reams)
    MhRealInput12.Text = Format(Val(Content), "0.000")
    fpSpread1.GetText 11, Row, Content   'Total Plates - ¼F
    MhRealInput3.Text = Format(Val(Content), "0")
    fpSpread1.GetText 12, Row, Content   'Total Plates - ½F
    MhRealInput23.Text = Format(Val(Content), "0")
    fpSpread1.GetText 13, Row, Content   'Total Plates - 1F
    MhRealInput24.Text = Format(Val(Content), "0")
    fpSpread1.GetText 14, Row, Content   'Plate Rate
    MhRealInput4.Text = Format(Val(Content), "0.00")
    fpSpread1.GetText 15, Row, Content   'Total Forms - ¼F
    MhRealInput6.Text = Format(Val(Content), "0")
    fpSpread1.GetText 16, Row, Content   'Total Forms - ½F
    MhRealInput25.Text = Format(Val(Content), "0")
    fpSpread1.GetText 17, Row, Content   'Total Forms - 1F
    MhRealInput26.Text = Format(Val(Content), "0")
    fpSpread1.GetText 18, Row, Content   'Print Rate
    MhRealInput5.Text = Format(Val(Content), "0.00")
    fpSpread1.GetText 19, Row, Content   'Paper Name
    Text1.Text = Content
    fpSpread1.GetText 21, Row, Content   'Forms/Sheet - For Printing Purpose
    MhRealInput22.Text = Format(IIf(Val(Content) = 0, 1, Val(Content)), "0.00")
    fpSpread1.GetText 22, Row, Content   'Forms/Sheet - For Paper Purpose
    MhRealInput27.Text = Format(IIf(Val(Content) = 0, 1, Val(Content)), "0.00")
End Sub
Private Sub LoadRefList(ByVal strBookCode As String, ByVal strOrderCode As String)
    Dim BalanceQuantity As Long
    On Error GoTo ErrorHandler
    If rstRefList.State = adStateOpen Then
        rstRefList.Close
    End If
    
    'And LEFT(BookPOParent.Code,1)<>'*'
    Dim sqlqey As String
    sqlqey = "Select P.Name,Quantity As PlannedQuantity,Format((Select Sum(ActualQuantity) From BookPOChild05,BookPOParent Where BookPOChild05.Ref=P.Code And BookPOParent.Code=BookPOChild05.Code And BookPOParent.Book=C.Book And BookPOChild05.Code<>'" & strOrderCode & "' And LEFT(BookPOParent.Code,1)<>'*'),0) As PrintedQuantity,Quantity As BalanceQuantity,Remarks As Col0,[PaperWastage%],P.Code From PrintPVParent P,PrintPVChild C Where P.Code=C.Code And P.PlanningType ='1' And C.Book='" & strBookCode & "' Order By P.Name"
    rstRefList.Open sqlqey, CxnDatabase, adOpenKeyset, adLockOptimistic
        
    rstRefList.ActiveConnection = Nothing
    
    Do While Not rstRefList.EOF
        BalanceQuantity = (Val(CheckNull(rstRefList.Fields("PlannedQuantity").Value)) - Val(CheckNull(rstRefList.Fields("PrintedQuantity").Value)))
        If BalanceQuantity <> 0 Then
            rstRefList.Fields("Col0").Value = Trim(rstRefList.Fields("Name").Value) + " Quantity : " + Format(str(BalanceQuantity), "#0")
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

Private Sub GetPrinterRates(ByVal xPrintingType As String, ByVal xRateType As String)   'xRateType : 'B'-Both Plate & Print Rate 'L'-Only Plate Rate
    Dim PrintRate As Double, PlateRate As Double, PaperWastageRate As Double, CurrentRate As Variant, PlateType As Variant
    On Error GoTo ErrorHandler

    If rstPrinterRates.State = adStateOpen Then rstPrinterRates.Close
    rstPrinterRates.Open "Select Top 1 * From AccountChild05 Where Code = '" & PrinterCode & "' And [Size] = '" & SizeCode & "' And Range" & Trim(xPrintingType) & " >= " & IIf(xPrintingType = "1", Val(MhRealInput2.Text), Val(MhRealInput19.Text)) & " Order By Range" & Trim(xPrintingType), CxnDatabase, adOpenKeyset, adLockReadOnly
    If rstPrinterRates.RecordCount = 0 Then
        If rstPrinterRates.State = adStateOpen Then rstPrinterRates.Close
        rstPrinterRates.Open "Select Top 1 * From AccountMaster,AccountChild05 Where AccountMaster.Code = AccountChild05.Code And [Name] Like '%Rate%' And [Size] = '" & SizeCode & "' And Range" & Trim(xPrintingType) & " >= " & IIf(xPrintingType = "1", Val(MhRealInput2.Text), Val(MhRealInput19.Text)) & " Order By Range" & Trim(xPrintingType), CxnDatabase, adOpenKeyset, adLockReadOnly
    End If
    If rstPrinterRates.RecordCount > 0 Then
        fpSpread1.GetText 6, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), PlateType
        PlateRate = rstPrinterRates.Fields(PlateType & "PlateRate" & Trim(xPrintingType)).Value
        PrintRate = rstPrinterRates.Fields("PrintRate" & Trim(xPrintingType)).Value
        PrintRate = PrintRate + IIf(PrintRate > 0, Val(FrmBookPrintOrder.rstBookList.Fields("AddOnRate01").Value), 0)
        PaperWastageRate = Val(rstPrinterRates.Fields("PaperWastageRate" & Trim(xPrintingType)))
    End If
    fpSpread1.GetText 14, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), CurrentRate  'Plate Rate
    If CurrentRate <> PlateRate Then
        If Val(CheckNull(rstBookPOChild05.Fields("ActualQuantity").Value)) <> 0 Then
            If MsgBox(IIf(xPrintingType = "1", "Single", IIf(xPrintingType = "2", "Double", "Four")) + " Color(s) Plate rate is different from that in Master ! Change rate?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then
                fpSpread1.SetText 14, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), PlateRate
            End If
        Else
            fpSpread1.SetText 14, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), PlateRate
        End If
    End If
    If xRateType = "B" Then
        fpSpread1.GetText 18, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), CurrentRate  'Print Rate
        If CurrentRate <> PrintRate And CurrentRate <> 0 Then
            If MsgBox(IIf(xPrintingType = "1", "Single", IIf(xPrintingType = "2", "Double", "Four")) + " Color(s) Printing Rate is different from that in Master ! Change Rate?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then
                fpSpread1.SetText 18, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), PrintRate
            End If
        Else
            fpSpread1.SetText 18, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), PrintRate
        End If
        fpSpread1.GetText 9, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), CurrentRate   'Paper Wastage Rate
        If CurrentRate <> PaperWastageRate Then
            If Val(CheckNull(rstBookPOChild05.Fields("ActualQuantity").Value)) <> 0 Then
                If MsgBox(IIf(xPrintingType = "1", "Single", IIf(xPrintingType = "2", "Double", "Four")) + " Color(s) Paper Wastage is different from that in Master ! Change Wastage?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then
                    fpSpread1.SetText 9, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), PaperWastageRate
                End If
            Else
                fpSpread1.SetText 9, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), PaperWastageRate
            End If
        End If
    End If
    If fpSpread1.ActiveRow = IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)) Then
        fpSpread1.GetText 14, fpSpread1.ActiveRow, CurrentRate  'Plate Rate
        MhRealInput4.Text = Format(CurrentRate, "0.00")
        fpSpread1.GetText 18, fpSpread1.ActiveRow, CurrentRate  'Print Rate
        MhRealInput5.Text = Format(CurrentRate, "0.00")
        fpSpread1.GetText 9, fpSpread1.ActiveRow, CurrentRate   'Paper Wastage Rate
        MhRealInput11.Text = Format(CurrentRate, "0.00")
    End If
    Exit Sub
ErrorHandler:
    DisplayError ("Failed to Fetch Printer Rates")
End Sub

Private Sub CalculateAQD()   'Calculate Actual Quantity Dependents
    
    Dim Q1 As Double, Q24 As Double
    'For Single Color : Actual Quantity = Billing Quantity + 10 % of Billing Quantity + 99
    Q1 = Val(MhRealInput1.Text) * 100 / (10 + 100) Mod 1000
    Q1 = IIf(Val(MhRealInput1.Text) > 99 And Q1 > 0 And Int(Q1) <= 90, Val(MhRealInput1.Text) - 99, Val(MhRealInput1.Text))  'New Actual Quantity
    Q1 = Int(Q1 * 100 / (10 + 100) / 1000) * 1000 + IIf(Q1 * 100 / (10 + 100) Mod 1000 = 0, 0, 1000)
    'For Double/Four Color : Actual Quantity = Billing Quantity - 200 + 99 OR Actual Quantity = Billing Quantity - 500 + 99
    Q24 = IIf(Int(Val(MhRealInput1.Text) / 1000) = 0, 1000, Int(Val(MhRealInput1.Text) / 1000) * 1000) + IIf(Val(MhRealInput1.Text) Mod 1000 <= IIf(Val(MhRealInput1.Text) <= 10000, 299, 599), 0, 1000)
    If Val(MhRealInput2.Text) = 0 Then
        MhRealInput2.Text = Format(Q1, "0")
    ElseIf Val(MhRealInput2.Text) <> Q1 Then
        If MsgBox("Variation (Single Color) between Billing Quantity (" & MhRealInput2.Text & ") Vs Calculated Billing Quantity (" & Trim(str(Q1)) & ") ! Change Quantity ?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then
            MhRealInput2.Text = Format(Q1, "0")
        End If
    End If
    If Val(MhRealInput19.Text) = 0 Then
        MhRealInput19.Text = Q24
    ElseIf Val(MhRealInput19.Text) <> Q24 Then
        If MsgBox("Variation (Double & Four Color) between Billing Quantity (" & MhRealInput19.Text & ") Vs Calculated Billing Quantity (" & Trim(str(Q24)) & ") ! Change Quantity ?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then
            MhRealInput19.Text = Format(Q24, "0")
        End If
    End If
    CalculateBQD ("S")
    CalculateBQD ("O")
    Call CalculateConsumption("1"): Call CalculateConsumption("2"): Call CalculateConsumption("4")
End Sub

Private Sub CalculateBQD(ByVal xPrintingType As String)    'Calculate Billing Quantity Dependents
    Dim Cnt As Integer, Content As Variant, Forms As Variant
    
    For Cnt = IIf(xPrintingType = "S", 1, 2) To IIf(xPrintingType = "S", 1, fpSpread1.MaxRows)
        fpSpread1.GetText 1, Cnt, Content   'Pages
        If Val(Content) <> 0 Then
            GetPrinterRates IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4")), "B"  'Get Both Plate & Printing Rates
        End If
        fpSpread1.GetText 3, Cnt, Forms
        Call CalculateTotalForms(IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4")), Val(Forms), "¼")
        fpSpread1.GetText 4, Cnt, Forms
        Call CalculateTotalForms(IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4")), Val(Forms), "½")
        fpSpread1.GetText 5, Cnt, Forms
        Call CalculateTotalForms(IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4")), Val(Forms), "1")
    Next
    CalculateAmount
End Sub
Private Function CalculateConsumption(ByVal xPrintingType As String) As Double
    Dim Forms¼ As Variant, Forms½ As Variant, Forms1 As Variant, WastageRate As Variant, CurrentPaperConsumption As Variant, Cnt As Integer, FS As Variant
    
    fpSpread1.GetText 3, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), Forms¼
    fpSpread1.GetText 4, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), Forms½
    fpSpread1.GetText 5, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), Forms1
    fpSpread1.GetText 9, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), WastageRate

    CalculateConsumption = CLng(Val(MhRealInput1.Text) * (Val(Forms¼) * 0.25 + Val(Forms½) * 0.5 + Val(Forms1) * 1) * ((100 + Val(WastageRate)) / 100))
    
    CalculateConsumption = CLng(Val(CalculateConsumption) / 2)
    
    fpSpread1.GetText 22, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), FS
    CalculateConsumption = IIf(Val(FS) = 0.5, 2, IIf(Val(FS) = 2, 0.5, 1)) * CalculateConsumption
    CalculateConsumption = Format(CLng(Int(Val(CalculateConsumption) / 500)) + ((Val(CalculateConsumption) Mod 500) / 1000), "0.000")
    fpSpread1.SetText 10, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), CalculateConsumption
    If fpSpread1.ActiveRow = IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)) Then
        MhRealInput12.Text = Format(Val(CalculateConsumption), "0.000")
    End If
    
    'Total Consumption Calculation
    For Cnt = 1 To fpSpread1.MaxRows
        fpSpread1.GetText 10, Cnt, CurrentPaperConsumption
        MhRealInput13.Text = Format(IIf(Cnt = 1, 0, Val(MhRealInput13.Text)) + CLng((Int(Val(CurrentPaperConsumption)) * 500) + ((Val(CurrentPaperConsumption) - Int(Val(CurrentPaperConsumption))) * 1000)), "0.000")
    Next
    MhRealInput13.Text = Format(CLng(Int(Val(MhRealInput13.Text) / 500)) + ((Val(MhRealInput13.Text) Mod 500) / 1000), "0.000")

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
            MhRealInput7.Text = Format(IIf(Cnt = 1, 1, IIf(Cnt = 2, 2, 4)) * (Val(TotalPlates¼) + Val(TotalPlates½) + Val(TotalPlates1)) * Val(PlateRate), "0.00")
            MhRealInput8.Text = Format(IIf(Cnt = 1, 1, IIf(Cnt = 2, 2, 4)) * (Val(TotalForms¼) + Val(TotalForms½) + Val(TotalForms1)) * Val(PrintRate), "0.00")
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
    'MhRealInput18.Text = Format(((TotalAmount) * 80 / 100) * Val(MhRealInput14.Text) / 100, "0.00")   'VAT
    MhRealInput18.Text = Format(((TotalAmount) * 100 / 100) * Val(MhRealInput14.Text) / 100, "0.00")   'VAT
    MhRealInput10.Text = Format(TotalAmount + Val(MhRealInput18.Text) + Val(MhRealInput9.Text), "0.00")   'Total Amount
    If Val(MhRealInput10.Text) > 0 Then
       MhRealInput101.Text = Format(Val(MhRealInput10.Text) / Val(MhRealInput1.Text), "0.000")   'Unit Cost
    End If
    
End Function
Private Function CalculateTotalForms(ByVal xPrintingType As String, ByVal Forms As Double, ByVal FormType As String) As Double
    Dim FS As Variant
    
    fpSpread1.GetText 21, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), FS
    CalculateTotalForms = (Int(IIf(xPrintingType = "1", Val(MhRealInput2.Text), Val(MhRealInput19.Text)) * IIf(FormType = "¼", 0.25, IIf(FormType = "½", 0.5, 1)) / 1000) + IIf(IIf(xPrintingType = "1", Val(MhRealInput2.Text), Val(MhRealInput19.Text)) * IIf(FormType = "¼", 0.25, IIf(FormType = "½", 0.5, 1)) Mod 1000 = 0, 0, 1)) * Forms
    CalculateTotalForms = IIf(Val(FS) = 0.5, 2, IIf(Val(FS) = 2, 0.5, 1)) * Val(CalculateTotalForms)
    If FrmBookPrintOrder.rstBookList.Fields("DuplexPrinting").Value = "N" Then CalculateTotalForms = 0.5 * CalculateTotalForms
    CalculateTotalForms = Int(Val(CalculateTotalForms)) + IIf(Val(CalculateTotalForms) - Int(Val(CalculateTotalForms)) = 0, 0, 1)
 
    If FormType = "¼" Then
        
        fpSpread1.SetText 15, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), CalculateTotalForms
        If fpSpread1.ActiveRow = IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)) Then
            MhRealInput6.Text = Format(CalculateTotalForms, "0")
        End If
    ElseIf FormType = "½" Then
    
        fpSpread1.SetText 16, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), CalculateTotalForms
        
        If fpSpread1.ActiveRow = IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)) Then
            MhRealInput25.Text = Format(CalculateTotalForms, "0")
        End If
    Else
        fpSpread1.SetText 17, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), CalculateTotalForms
        If fpSpread1.ActiveRow = IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)) Then
            MhRealInput26.Text = Format(CalculateTotalForms, "0")
        End If
        
    End If
    
End Function
Private Function CalculateTotalPlates(ByVal xPrintingType As String, ByVal Forms As Double, ByVal FormType As String) As Double
    Dim FS As Variant
    fpSpread1.GetText 21, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), FS
    CalculateTotalPlates = Forms
    CalculateTotalPlates = IIf(Val(FS) = 0.5, 2, IIf(Val(FS) = 2, 0.5, 1)) * Val(CalculateTotalPlates)
    If FrmBookPrintOrder.rstBookList.Fields("DuplexPrinting").Value = "N" Then CalculateTotalPlates = 0.5 * CalculateTotalPlates
    CalculateTotalPlates = Int(Val(CalculateTotalPlates)) + IIf(Val(CalculateTotalPlates) - Int(Val(CalculateTotalPlates)) = 0.5, 1, 0)
    If FormType = "¼" Then
        fpSpread1.SetText 11, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), CalculateTotalPlates
        If fpSpread1.ActiveRow = IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)) Then
            MhRealInput3.Text = Format(CalculateTotalPlates, "0")
        End If
    ElseIf FormType = "½" Then
        fpSpread1.SetText 12, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), CalculateTotalPlates
        If fpSpread1.ActiveRow = IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)) Then
            MhRealInput23.Text = Format(CalculateTotalPlates, "0")
        End If
    Else
        fpSpread1.SetText 13, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), CalculateTotalPlates
        If fpSpread1.ActiveRow = IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)) Then
            MhRealInput24.Text = Format(CalculateTotalPlates, "0")
        End If
    End If
End Function
Private Sub cmdProceed_Click()
    Dim Cnt As Integer, PaperBalance As Double, PaperCode As Variant, PaperName As Variant, PaperStock As Variant, PaperConsumption As Variant
    If CheckMandatoryFields Then Exit Sub
    If FrmBookPrintOrder.BookPOType <> "O" Then
       For Cnt = 1 To fpSpread1.MaxRows
            fpSpread1.SetActiveCell 1, Cnt
            fpSpread1_DblClick 1, Cnt
            fpSpread1.GetText 20, Cnt, PaperCode
            PaperStock = CalculatePaperBalance(PrinterCode, PaperCode, CheckNull(rstBookPOChild05.Fields("Code").Value), "BPOB")
            fpSpread1.GetText 10, Cnt, PaperConsumption
            If Not CheckEmpty(PaperCode, False) Then
                PaperBalance = Val(PaperStock) - Int(Val(PaperConsumption)) * 500 - Round((Val(PaperConsumption) - Int(Val(PaperConsumption))), 3) * 1000
                If PaperBalance < 0 Then
                    PaperBalance = Format(CLng(Int(Val(Abs(PaperBalance)) / 500)) + ((Val(Abs(PaperBalance)) Mod 500) / 1000), "0.000")
                    fpSpread1.GetText 19, Cnt, PaperName
                    If UserLevel <= 2 Then
                        If MsgBox("Stock (" & Format(0 - PaperBalance, "0.000") & ") of the Paper - " & Trim(PaperName) & vbCrLf & " is going negative ! Would you like to continue ?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Proceed !") = vbNo Then
                            Exit Sub
                        End If
                    Else
                        Call DisplayError("Cann't Save ! Stock (" & Format(0 - PaperBalance, "0.000") & ") of the Paper - " & Trim(PaperName) & " is going negative"): AbortPO = True: Exit Sub
                    End If
                End If
            End If
       Next
    End If
    SaveFields
    rstBookPOChild05.Update
    Call CloseForm(Me)
End Sub
Private Function CheckMandatoryFields() As Boolean
    Dim Cnt As Integer, Pages As Variant, Paper As Variant, Forms¼ As Variant, Forms½ As Variant, Forms1 As Variant, TotalForms As Variant
    If Combo2.ListIndex < 0 Then Combo2.SetFocus: CheckMandatoryFields = True: Exit Function
    If Combo3.ListIndex < 0 Then Combo3.SetFocus: CheckMandatoryFields = True: Exit Function
    If Combo31.ListIndex < 0 Then Combo31.SetFocus: CheckMandatoryFields = True: Exit Function
    If Val(MhRealInput16.Text) <> 0 Then If Val(MhRealInput16.Text) <> Val(MhRealInput10.Text) Then MhRealInput9.SetFocus: CheckMandatoryFields = True: Exit Function
    If Val(MhRealInput9.Text) <> 0 Then If CheckEmpty(TxtAdNar.Text, False) Then TxtAdNar.SetFocus: CheckMandatoryFields = True: Exit Function
    
    For Cnt = 1 To fpSpread1.MaxRows
        
        fpSpread1.SetActiveCell 1, Cnt
        fpSpread1_DblClick 1, Cnt
        fpSpread1.GetText 1, Cnt, Pages
        fpSpread1.GetText 20, Cnt, Paper
        
        If Pages <> 0 Then
            If CheckNull(Paper) = "" Then
                MhRealInput15.SetFocus
                CheckMandatoryFields = True
                Exit For
            End If
            
        End If
        
        
        fpSpread1.GetText 2, Cnt, TotalForms
        fpSpread1.GetText 3, Cnt, Forms¼
        fpSpread1.GetText 4, Cnt, Forms½
        fpSpread1.GetText 5, Cnt, Forms1
        
        If Val(Forms¼) * 0.25 + Val(Forms½) * 0.5 + Val(Forms1) * 1 <> TotalForms Then
            DisplayError ("Variation between Total Forms Vs Bifurcated Forms")
            MhRealInput17.SetFocus
            CheckMandatoryFields = True
            Exit For
        End If
    Next
    
End Function
Private Sub cmdCancel_Click()
    rstBookPOChild05.CancelUpdate
    Call CloseForm(Me)
End Sub

Private Sub TxtPOrder_Change()
    
'    If TxtPOrder.Text = " " Then
'        TxtPOrder.Text = "?"
'        SendKeys "{TAB}"
'    ElseIf CheckEmpty(TxtPOrder, False) Then
'        PurchaseOrder = ""
'    End If

 If TxtPOrder.Text = " " Then
        TxtPOrder.Text = "?"
        SendKeys "{TAB}"
End If
    
End Sub

Private Sub TxtPOrder_Validate(Cancel As Boolean)
    Dim SearchString As String
    If CheckEmpty(TxtPOrder, False) Then Exit Sub
    SearchString = FixQuote(TxtPOrder.Text)
    If rstPurchaseOrder.RecordCount = 0 Then
        DisplayError ("No Record in Paper Purchase")
        Cancel = True
        Exit Sub
    Else
        rstPurchaseOrder.MoveFirst
    End If
    rstPurchaseOrder.Find "[Col0] = '" & RTrim(SearchString) & "'"
    If rstPurchaseOrder.EOF Then
        SelectionType = "S"
        PurchaseOrder = ""
        Call LoadSelectionList(rstPurchaseOrder, "List of Paper Purchase...", "PO.No.")
        SearchOrder = 0
        Call DisplaySelectionList(TxtPOrder, PurchaseOrder)
        Call CloseForm(FrmSelectionList)
        If CheckEmpty(TxtPOrder.Text, False) Then TxtPOrder.Text = "?"
        If RTrim(PurchaseOrder) <> "" Then
            SendKeys "{TAB}"
        End If
        Cancel = True
    Else
        PurchaseOrder = rstPurchaseOrder.Fields("Code").Value
        TxtPOrder.Text = rstPurchaseOrder.Fields("Col0").Value
        
        
    End If

End Sub
