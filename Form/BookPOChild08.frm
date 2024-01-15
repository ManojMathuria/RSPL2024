VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmBookPOChild08 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Book Binding Order Details"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8280
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "BookPOChild08.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "FrmLogin"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   8280
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7815
      Picture         =   "BookPOChild08.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Cancel"
      Top             =   465
      Width           =   375
   End
   Begin VB.CommandButton cmdProceed 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7815
      Picture         =   "BookPOChild08.frx":0544
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Save"
      Top             =   105
      Width           =   375
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
      Height          =   7155
      Left            =   120
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   105
      Width           =   7575
      _Version        =   65536
      _ExtentX        =   13361
      _ExtentY        =   12621
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
      Picture         =   "BookPOChild08.frx":0646
      Begin VB.TextBox Text211 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc1"
         Height          =   370
         Left            =   1560
         MaxLength       =   139
         TabIndex        =   2
         Top             =   1280
         Width           =   1095
      End
      Begin VB.TextBox TxtAdNar 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc1"
         Height          =   330
         Left            =   1560
         MaxLength       =   139
         TabIndex        =   30
         Top             =   6330
         Width           =   5895
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc1"
         Height          =   330
         Left            =   1560
         MaxLength       =   139
         TabIndex        =   29
         Top             =   6015
         Width           =   3615
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc1"
         Height          =   330
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   105
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc1"
         Height          =   330
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   640
         Width           =   3615
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc1"
         Height          =   330
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   26
         Top             =   5370
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc1"
         Height          =   330
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   60
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   965
         Width           =   5895
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc1"
         Height          =   330
         Left            =   4080
         MaxLength       =   40
         TabIndex        =   7
         Top             =   1635
         Width           =   3375
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
         Height          =   330
         Left            =   2640
         TabIndex        =   37
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
         Picture         =   "BookPOChild08.frx":0662
         Picture         =   "BookPOChild08.frx":067E
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
         Height          =   330
         Left            =   120
         TabIndex        =   38
         Top             =   960
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
         Caption         =   " Book Name"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild08.frx":069A
         Picture         =   "BookPOChild08.frx":06B6
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel9 
         Height          =   330
         Left            =   120
         TabIndex        =   39
         Top             =   1950
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
         Picture         =   "BookPOChild08.frx":06D2
         Picture         =   "BookPOChild08.frx":06EE
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
         Height          =   330
         Left            =   120
         TabIndex        =   40
         Top             =   2265
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
         Caption         =   " Folding Rate"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild08.frx":070A
         Picture         =   "BookPOChild08.frx":0726
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
         Height          =   330
         Left            =   2640
         TabIndex        =   41
         Top             =   2265
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
         Caption         =   " Stitching Rate"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild08.frx":0742
         Picture         =   "BookPOChild08.frx":075E
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
         Height          =   330
         Left            =   120
         TabIndex        =   42
         Top             =   2895
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
         Caption         =   " Rate/Book"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild08.frx":077A
         Picture         =   "BookPOChild08.frx":0796
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel10 
         Height          =   330
         Index           =   0
         Left            =   2640
         TabIndex        =   43
         Top             =   1635
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
         Caption         =   " Binding Type"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild08.frx":07B2
         Picture         =   "BookPOChild08.frx":07CE
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
         Height          =   330
         Left            =   2640
         TabIndex        =   44
         Top             =   1950
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
         Caption         =   " Billing Quantity"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild08.frx":07EA
         Picture         =   "BookPOChild08.frx":0806
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel13 
         Height          =   330
         Left            =   5160
         TabIndex        =   45
         Top             =   2265
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
         Caption         =   " Pasting Rate"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild08.frx":0822
         Picture         =   "BookPOChild08.frx":083E
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel15 
         Height          =   330
         Left            =   2640
         TabIndex        =   46
         Top             =   2895
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
         Picture         =   "BookPOChild08.frx":085A
         Picture         =   "BookPOChild08.frx":0876
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel19 
         Height          =   330
         Left            =   120
         TabIndex        =   47
         Top             =   5370
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
         Picture         =   "BookPOChild08.frx":0892
         Picture         =   "BookPOChild08.frx":08AE
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel20 
         Height          =   330
         Left            =   5160
         TabIndex        =   48
         Top             =   5370
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
         Picture         =   "BookPOChild08.frx":08CA
         Picture         =   "BookPOChild08.frx":08E6
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel23 
         Height          =   330
         Left            =   2640
         TabIndex        =   49
         Top             =   5370
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
         Picture         =   "BookPOChild08.frx":0902
         Picture         =   "BookPOChild08.frx":091E
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel24 
         Height          =   330
         Left            =   5160
         TabIndex        =   50
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
         Picture         =   "BookPOChild08.frx":093A
         Picture         =   "BookPOChild08.frx":0956
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel25 
         Height          =   330
         Left            =   120
         TabIndex        =   52
         Top             =   645
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
         Caption         =   " Binder Name"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild08.frx":0972
         Picture         =   "BookPOChild08.frx":098E
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel26 
         Height          =   330
         Left            =   5160
         TabIndex        =   53
         Top             =   1950
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
         Caption         =   " Adj.Quantity"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild08.frx":09AA
         Picture         =   "BookPOChild08.frx":09C6
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel27 
         Height          =   330
         Left            =   120
         TabIndex        =   55
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
         Picture         =   "BookPOChild08.frx":09E2
         Picture         =   "BookPOChild08.frx":09FE
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel28 
         Height          =   330
         Left            =   120
         TabIndex        =   56
         Top             =   6015
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
         Picture         =   "BookPOChild08.frx":0A1A
         Picture         =   "BookPOChild08.frx":0A36
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
         Height          =   330
         Left            =   120
         TabIndex        =   57
         Top             =   4155
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
         Caption         =   " Cartage/Box"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild08.frx":0A52
         Picture         =   "BookPOChild08.frx":0A6E
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
         Height          =   330
         Index           =   0
         Left            =   5160
         TabIndex        =   58
         Top             =   4470
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
         Picture         =   "BookPOChild08.frx":0A8A
         Picture         =   "BookPOChild08.frx":0AA6
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel12 
         Height          =   330
         Left            =   120
         TabIndex        =   59
         Top             =   1635
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
         Caption         =   " Binding Form"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild08.frx":0AC2
         Picture         =   "BookPOChild08.frx":0ADE
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel14 
         Height          =   330
         Left            =   2640
         TabIndex        =   60
         Top             =   4155
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
         Caption         =   " Cartage Amt"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild08.frx":0AFA
         Picture         =   "BookPOChild08.frx":0B16
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel16 
         Height          =   330
         Left            =   5160
         TabIndex        =   61
         Top             =   4155
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
         Picture         =   "BookPOChild08.frx":0B32
         Picture         =   "BookPOChild08.frx":0B4E
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel17 
         Height          =   330
         Left            =   5160
         TabIndex        =   62
         Top             =   2895
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
         Caption         =   " Pkt/Box"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild08.frx":0B6A
         Picture         =   "BookPOChild08.frx":0B86
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel7 
         Height          =   330
         Index           =   0
         Left            =   120
         TabIndex        =   63
         Top             =   4470
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
         Picture         =   "BookPOChild08.frx":0BA2
         Picture         =   "BookPOChild08.frx":0BBE
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel18 
         Height          =   330
         Index           =   0
         Left            =   5160
         TabIndex        =   64
         Top             =   6015
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
         Caption         =   " Received Qty"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild08.frx":0BDA
         Picture         =   "BookPOChild08.frx":0BF6
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
         Calendar        =   "BookPOChild08.frx":0C12
         Caption         =   "BookPOChild08.frx":0D2A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":0D96
         Keys            =   "BookPOChild08.frx":0DB4
         Spin            =   "BookPOChild08.frx":0E12
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
         Calendar        =   "BookPOChild08.frx":0E3A
         Caption         =   "BookPOChild08.frx":0F52
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":0FBE
         Keys            =   "BookPOChild08.frx":0FDC
         Spin            =   "BookPOChild08.frx":103A
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
         TabIndex        =   27
         Top             =   5370
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calendar        =   "BookPOChild08.frx":1062
         Caption         =   "BookPOChild08.frx":117A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":11E6
         Keys            =   "BookPOChild08.frx":1204
         Spin            =   "BookPOChild08.frx":1262
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
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel21 
         Height          =   330
         Left            =   120
         TabIndex        =   65
         Top             =   2580
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
         Caption         =   " Folding Amount"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild08.frx":128A
         Picture         =   "BookPOChild08.frx":12A6
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel22 
         Height          =   330
         Left            =   2640
         TabIndex        =   66
         Top             =   2580
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
         Caption         =   " Stitching Amt"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild08.frx":12C2
         Picture         =   "BookPOChild08.frx":12DE
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel29 
         Height          =   330
         Left            =   5160
         TabIndex        =   67
         Top             =   2580
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
         Caption         =   " Pasting Amt"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild08.frx":12FA
         Picture         =   "BookPOChild08.frx":1316
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput7 
         Height          =   330
         Left            =   1560
         TabIndex        =   5
         Top             =   1635
         Width           =   600
         _Version        =   65536
         _ExtentX        =   1050
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":1332
         Caption         =   "BookPOChild08.frx":1352
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":13BE
         Keys            =   "BookPOChild08.frx":13DC
         Spin            =   "BookPOChild08.frx":1426
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "####0"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "####0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99999
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput18 
         Height          =   330
         Left            =   2145
         TabIndex        =   6
         Top             =   1635
         Width           =   510
         _Version        =   65536
         _ExtentX        =   900
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":144E
         Caption         =   "BookPOChild08.frx":146E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":14DA
         Keys            =   "BookPOChild08.frx":14F8
         Spin            =   "BookPOChild08.frx":1542
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "####0"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "####0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99999
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput1 
         Height          =   330
         Left            =   1560
         TabIndex        =   8
         Top             =   1950
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":156A
         Caption         =   "BookPOChild08.frx":158A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":15F6
         Keys            =   "BookPOChild08.frx":1614
         Spin            =   "BookPOChild08.frx":165E
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput3 
         Height          =   330
         Left            =   4080
         TabIndex        =   9
         Top             =   1950
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":1686
         Caption         =   "BookPOChild08.frx":16A6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":1712
         Keys            =   "BookPOChild08.frx":1730
         Spin            =   "BookPOChild08.frx":177A
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput4 
         Height          =   330
         Left            =   6360
         TabIndex        =   10
         Top             =   1950
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":17A2
         Caption         =   "BookPOChild08.frx":17C2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":182E
         Keys            =   "BookPOChild08.frx":184C
         Spin            =   "BookPOChild08.frx":1896
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
         ValueVT         =   1968898053
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput11 
         Height          =   330
         Left            =   6360
         TabIndex        =   16
         Top             =   2895
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":18BE
         Caption         =   "BookPOChild08.frx":18DE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":194A
         Keys            =   "BookPOChild08.frx":1968
         Spin            =   "BookPOChild08.frx":19B2
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
         ValueVT         =   1968898053
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput12 
         Height          =   330
         Left            =   1560
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   3840
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":19DA
         Caption         =   "BookPOChild08.frx":19FA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":1A66
         Keys            =   "BookPOChild08.frx":1A84
         Spin            =   "BookPOChild08.frx":1ACE
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
         ReadOnly        =   -1
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   1
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput2 
         Height          =   330
         Left            =   1560
         TabIndex        =   11
         Top             =   2265
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":1AF6
         Caption         =   "BookPOChild08.frx":1B16
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":1B82
         Keys            =   "BookPOChild08.frx":1BA0
         Spin            =   "BookPOChild08.frx":1BEA
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
         ValueVT         =   1968898053
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput5 
         Height          =   330
         Left            =   4080
         TabIndex        =   12
         Top             =   2265
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":1C12
         Caption         =   "BookPOChild08.frx":1C32
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":1C9E
         Keys            =   "BookPOChild08.frx":1CBC
         Spin            =   "BookPOChild08.frx":1D06
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
         ValueVT         =   1968898053
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput6 
         Height          =   330
         Left            =   6360
         TabIndex        =   13
         Top             =   2265
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":1D2E
         Caption         =   "BookPOChild08.frx":1D4E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":1DBA
         Keys            =   "BookPOChild08.frx":1DD8
         Spin            =   "BookPOChild08.frx":1E22
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
         ValueVT         =   1968898053
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput20 
         Height          =   330
         Left            =   1560
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   2580
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":1E4A
         Caption         =   "BookPOChild08.frx":1E6A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":1ED6
         Keys            =   "BookPOChild08.frx":1EF4
         Spin            =   "BookPOChild08.frx":1F3E
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
         ValueVT         =   1968898053
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput21 
         Height          =   330
         Left            =   4080
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   2580
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":1F66
         Caption         =   "BookPOChild08.frx":1F86
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":1FF2
         Keys            =   "BookPOChild08.frx":2010
         Spin            =   "BookPOChild08.frx":205A
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
         ValueVT         =   1968898053
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput22 
         Height          =   330
         Left            =   6360
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   2580
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":2082
         Caption         =   "BookPOChild08.frx":20A2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":210E
         Keys            =   "BookPOChild08.frx":212C
         Spin            =   "BookPOChild08.frx":2176
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
         ValueVT         =   1968898053
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput8 
         Height          =   330
         Left            =   1560
         TabIndex        =   14
         Top             =   2895
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":219E
         Caption         =   "BookPOChild08.frx":21BE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":222A
         Keys            =   "BookPOChild08.frx":2248
         Spin            =   "BookPOChild08.frx":2292
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
         ValueVT         =   1968898053
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput13 
         Height          =   330
         Left            =   1560
         TabIndex        =   23
         Top             =   4155
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":22BA
         Caption         =   "BookPOChild08.frx":22DA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":2346
         Keys            =   "BookPOChild08.frx":2364
         Spin            =   "BookPOChild08.frx":23AE
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
         ValueVT         =   1968898053
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput14 
         Height          =   330
         Left            =   4080
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   4155
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":23D6
         Caption         =   "BookPOChild08.frx":23F6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":2462
         Keys            =   "BookPOChild08.frx":2480
         Spin            =   "BookPOChild08.frx":24CA
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
         ValueVT         =   1968898053
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput15 
         Height          =   330
         Left            =   6360
         TabIndex        =   24
         Top             =   4155
         Width           =   495
         _Version        =   65536
         _ExtentX        =   873
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":24F2
         Caption         =   "BookPOChild08.frx":2512
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":257E
         Keys            =   "BookPOChild08.frx":259C
         Spin            =   "BookPOChild08.frx":25E6
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInput17 
         Height          =   330
         Left            =   6840
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   4155
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":260E
         Caption         =   "BookPOChild08.frx":262E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":269A
         Keys            =   "BookPOChild08.frx":26B8
         Spin            =   "BookPOChild08.frx":2702
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
         ValueVT         =   1968898053
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput9 
         Height          =   330
         Left            =   1560
         TabIndex        =   25
         Top             =   4470
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":272A
         Caption         =   "BookPOChild08.frx":274A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":27B6
         Keys            =   "BookPOChild08.frx":27D4
         Spin            =   "BookPOChild08.frx":281E
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
         ValueVT         =   1968898053
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput10 
         Height          =   330
         Left            =   6360
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   4470
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":2846
         Caption         =   "BookPOChild08.frx":2866
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":28D2
         Keys            =   "BookPOChild08.frx":28F0
         Spin            =   "BookPOChild08.frx":293A
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
         ValueVT         =   1968898053
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput16 
         Height          =   330
         Left            =   6360
         TabIndex        =   28
         Top             =   5370
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":2962
         Caption         =   "BookPOChild08.frx":2982
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":29EE
         Keys            =   "BookPOChild08.frx":2A0C
         Spin            =   "BookPOChild08.frx":2A56
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput19 
         Height          =   330
         Left            =   6360
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   6015
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":2A7E
         Caption         =   "BookPOChild08.frx":2A9E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":2B0A
         Keys            =   "BookPOChild08.frx":2B28
         Spin            =   "BookPOChild08.frx":2B72
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel32 
         Height          =   330
         Left            =   120
         TabIndex        =   73
         Top             =   6330
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
         Picture         =   "BookPOChild08.frx":2B9A
         Picture         =   "BookPOChild08.frx":2BB6
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel30 
         Height          =   330
         Left            =   120
         TabIndex        =   74
         Top             =   3525
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
         Caption         =   " Total Packets"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild08.frx":2BD2
         Picture         =   "BookPOChild08.frx":2BEE
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel31 
         Height          =   330
         Left            =   120
         TabIndex        =   75
         Top             =   3840
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
         Caption         =   " Total Boxes"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild08.frx":2C0A
         Picture         =   "BookPOChild08.frx":2C26
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel33 
         Height          =   330
         Left            =   2640
         TabIndex        =   76
         Top             =   3525
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
         Caption         =   " Pkt Packing Rate"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild08.frx":2C42
         Picture         =   "BookPOChild08.frx":2C5E
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel34 
         Height          =   330
         Left            =   2640
         TabIndex        =   77
         Top             =   3840
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
         Caption         =   " Box Packing Rate"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild08.frx":2C7A
         Picture         =   "BookPOChild08.frx":2C96
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel35 
         Height          =   330
         Left            =   5160
         TabIndex        =   78
         Top             =   3525
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
         Caption         =   " Pkt Pack Amt"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild08.frx":2CB2
         Picture         =   "BookPOChild08.frx":2CCE
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel36 
         Height          =   330
         Left            =   5160
         TabIndex        =   79
         Top             =   3840
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
         Caption         =   " Box Pack Amt"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild08.frx":2CEA
         Picture         =   "BookPOChild08.frx":2D06
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput23 
         Height          =   330
         Left            =   4080
         TabIndex        =   15
         Top             =   2895
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":2D22
         Caption         =   "BookPOChild08.frx":2D42
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":2DAE
         Keys            =   "BookPOChild08.frx":2DCC
         Spin            =   "BookPOChild08.frx":2E16
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
         ValueVT         =   1968898053
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput24 
         Height          =   330
         Left            =   1560
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   3525
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":2E3E
         Caption         =   "BookPOChild08.frx":2E5E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":2ECA
         Keys            =   "BookPOChild08.frx":2EE8
         Spin            =   "BookPOChild08.frx":2F32
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
         ReadOnly        =   -1
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   1
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput25 
         Height          =   330
         Left            =   4080
         TabIndex        =   20
         Top             =   3525
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":2F5A
         Caption         =   "BookPOChild08.frx":2F7A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":2FE6
         Keys            =   "BookPOChild08.frx":3004
         Spin            =   "BookPOChild08.frx":304E
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
         ValueVT         =   1966145541
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput26 
         Height          =   330
         Left            =   4080
         TabIndex        =   22
         Top             =   3840
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":3076
         Caption         =   "BookPOChild08.frx":3096
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":3102
         Keys            =   "BookPOChild08.frx":3120
         Spin            =   "BookPOChild08.frx":316A
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
         ValueVT         =   1968898053
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput27 
         Height          =   330
         Left            =   6360
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   3525
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":3192
         Caption         =   "BookPOChild08.frx":31B2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":321E
         Keys            =   "BookPOChild08.frx":323C
         Spin            =   "BookPOChild08.frx":3286
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
         ValueVT         =   1968898053
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput28 
         Height          =   330
         Left            =   6360
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   3840
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":32AE
         Caption         =   "BookPOChild08.frx":32CE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":333A
         Keys            =   "BookPOChild08.frx":3358
         Spin            =   "BookPOChild08.frx":33A2
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
         ValueVT         =   1968898053
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel37 
         Height          =   330
         Left            =   120
         TabIndex        =   82
         Top             =   3210
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
         Caption         =   " Loose Qty/Box"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild08.frx":33CA
         Picture         =   "BookPOChild08.frx":33E6
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel38 
         Height          =   330
         Left            =   2640
         TabIndex        =   83
         Top             =   3210
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
         Caption         =   " Extra Loose Qty"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild08.frx":3402
         Picture         =   "BookPOChild08.frx":341E
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel39 
         Height          =   330
         Left            =   5160
         TabIndex        =   84
         Top             =   3210
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
         Caption         =   " Tot.Loose Qty"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild08.frx":343A
         Picture         =   "BookPOChild08.frx":3456
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput29 
         Height          =   330
         Left            =   1560
         TabIndex        =   17
         Top             =   3210
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":3472
         Caption         =   "BookPOChild08.frx":3492
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":34FE
         Keys            =   "BookPOChild08.frx":351C
         Spin            =   "BookPOChild08.frx":3566
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
         ValueVT         =   1968898053
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput30 
         Height          =   330
         Left            =   4080
         TabIndex        =   18
         Top             =   3210
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":358E
         Caption         =   "BookPOChild08.frx":35AE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":361A
         Keys            =   "BookPOChild08.frx":3638
         Spin            =   "BookPOChild08.frx":3682
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
         ValueVT         =   282853377
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput31 
         Height          =   330
         Left            =   6360
         TabIndex        =   85
         TabStop         =   0   'False
         Top             =   3210
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":36AA
         Caption         =   "BookPOChild08.frx":36CA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":3736
         Keys            =   "BookPOChild08.frx":3754
         Spin            =   "BookPOChild08.frx":379E
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
         ReadOnly        =   -1
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   1968898053
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel321 
         Height          =   330
         Left            =   5160
         TabIndex        =   86
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
         Picture         =   "BookPOChild08.frx":37C6
         Picture         =   "BookPOChild08.frx":37E2
      End
      Begin TDBDate6Ctl.TDBDate MhDateInput31 
         Height          =   330
         Left            =   6360
         TabIndex        =   87
         Top             =   640
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calendar        =   "BookPOChild08.frx":37FE
         Caption         =   "BookPOChild08.frx":3916
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":3982
         Keys            =   "BookPOChild08.frx":39A0
         Spin            =   "BookPOChild08.frx":39FE
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
         Index           =   0
         Left            =   2640
         TabIndex        =   88
         Top             =   4470
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
         Caption         =   " Unit Cost"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild08.frx":3A26
         Picture         =   "BookPOChild08.frx":3A42
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput101 
         Height          =   330
         Left            =   4080
         TabIndex        =   89
         TabStop         =   0   'False
         Top             =   4470
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":3A5E
         Caption         =   "BookPOChild08.frx":3A7E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":3AEA
         Keys            =   "BookPOChild08.frx":3B08
         Spin            =   "BookPOChild08.frx":3B52
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
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel281 
         Height          =   370
         Left            =   120
         TabIndex        =   90
         Top             =   1280
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   653
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
         Caption         =   " Edition"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild08.frx":3B7A
         Picture         =   "BookPOChild08.frx":3B96
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel18 
         Height          =   370
         Index           =   1
         Left            =   5160
         TabIndex        =   91
         Top             =   1280
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   653
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
         Caption         =   " A.Recvd Date"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild08.frx":3BB2
         Picture         =   "BookPOChild08.frx":3BCE
      End
      Begin TDBDate6Ctl.TDBDate MhDateInput311 
         Height          =   370
         Left            =   6360
         TabIndex        =   4
         Top             =   1280
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   653
         Calendar        =   "BookPOChild08.frx":3BEA
         Caption         =   "BookPOChild08.frx":3D02
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":3D6E
         Keys            =   "BookPOChild08.frx":3D8C
         Spin            =   "BookPOChild08.frx":3DEA
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
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel7 
         Height          =   330
         Index           =   1
         Left            =   120
         TabIndex        =   92
         Top             =   4790
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
         Caption         =   " Noida"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild08.frx":3E12
         Picture         =   "BookPOChild08.frx":3E2E
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel51 
         Height          =   330
         Index           =   1
         Left            =   2640
         TabIndex        =   93
         Top             =   4790
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
         Caption         =   " D.Daryaganj"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild08.frx":3E4A
         Picture         =   "BookPOChild08.frx":3E66
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
         Height          =   330
         Index           =   1
         Left            =   5160
         TabIndex        =   94
         Top             =   4790
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
         Caption         =   " 8 No"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild08.frx":3E82
         Picture         =   "BookPOChild08.frx":3E9E
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInputD 
         Height          =   330
         Left            =   4080
         TabIndex        =   95
         Top             =   4785
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":3EBA
         Caption         =   "BookPOChild08.frx":3EDA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":3F46
         Keys            =   "BookPOChild08.frx":3F64
         Spin            =   "BookPOChild08.frx":3FAE
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
      Begin TDBNumber6Ctl.TDBNumber MhRealInputN 
         Height          =   330
         Left            =   1560
         TabIndex        =   96
         Top             =   4785
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":3FD6
         Caption         =   "BookPOChild08.frx":3FF6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":4062
         Keys            =   "BookPOChild08.frx":4080
         Spin            =   "BookPOChild08.frx":40CA
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
         ValueVT         =   1965359109
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInputE 
         Height          =   330
         Left            =   6360
         TabIndex        =   97
         Top             =   4790
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calculator      =   "BookPOChild08.frx":40F2
         Caption         =   "BookPOChild08.frx":4112
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BookPOChild08.frx":417E
         Keys            =   "BookPOChild08.frx":419C
         Spin            =   "BookPOChild08.frx":41E6
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
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel10 
         Height          =   370
         Index           =   1
         Left            =   2640
         TabIndex        =   98
         Top             =   1275
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   653
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
         Caption         =   " Adv. Copy Req"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild08.frx":420E
         Picture         =   "BookPOChild08.frx":422A
      End
      Begin MSForms.ComboBox Combo1 
         Height          =   370
         Left            =   4080
         TabIndex        =   3
         Top             =   1275
         Width           =   1095
         VariousPropertyBits=   545282075
         BackColor       =   16777215
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "1931;653"
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
         Y1              =   5910
         Y2              =   5910
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
         Y1              =   5265
         Y2              =   5265
      End
   End
End
Attribute VB_Name = "FrmBookPOChild08"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public rstBookPOChild08 As New ADODB.Recordset
Dim rstBindingTypeList As New ADODB.Recordset
Dim rstBinderRates As New ADODB.Recordset
Public rstRefList As New ADODB.Recordset

Public BinderCode As String
Public BookPrinterQuantity As Long
Dim FormType As String
Dim SizeCode As String
Dim BindingTypeCode As String
Private Sub Combo1_Validate(Cancel As Boolean)
    If CheckEmpty(Combo1, True) Then
        Cancel = True
    End If
End Sub
Private Sub Form_Load()
   
    On Error GoTo ErrorHandler
    CenterForm Me
    BusySystemIndicator True
    DisableCloseButton Me
    
    FormType = FrmBookPrintOrder.rstBookList.Fields("FormType").Value
    SizeCode = FrmBookPrintOrder.rstBookList.Fields("SizeCode").Value
    Text5.Text = Trim(FrmBookPrintOrder.Text2.Text)
    Text4.Text = Trim(FrmBookPrintOrder.Text8.Text)
    
    Text2.Text = Trim(FrmBookPrintOrder.Text3.Text)
    
    Combo1.AddItem "Yes", 0
    Combo1.AddItem "No", 1
    ClearFields
    Call LoadRefList(Text5.Text)
    
    rstBindingTypeList.Open "Select Name As Col0, Code From GeneralMaster Where Type = '6' Order By Name", CxnDatabase, adOpenKeyset, adLockReadOnly
    rstBindingTypeList.ActiveConnection = Nothing

    If Val(CheckNull(rstBookPOChild08.Fields("ActualQuantity").Value)) = 0 Then
        
        MhRealInput1.Text = Format(BookPrinterQuantity, "0")
        MhRealInput7.Text = Format(Val(FrmBookPrintOrder.rstBookList.Fields("BindingForms01").Value), "0")
        MhRealInput18.Text = Format(Val(FrmBookPrintOrder.rstBookList.Fields("BindingForms02").Value), "0")
        MhRealInput23.Text = Format(Val(FrmBookPrintOrder.rstBookList.Fields("Qty/Pkt").Value), "0")
        MhRealInput11.Text = Format(Val(FrmBookPrintOrder.rstBookList.Fields("Pkt/Box").Value), "0")
        MhRealInput29.Text = Format(Val(FrmBookPrintOrder.rstBookList.Fields("LooseQty/Box").Value), "0")
        BindingTypeCode = FrmBookPrintOrder.rstBookList.Fields("BindingType").Value
        If rstBindingTypeList.RecordCount > 0 Then rstBindingTypeList.MoveFirst
        rstBindingTypeList.Find "[Code] = '" & BindingTypeCode & "'"
        If Not rstBindingTypeList.EOF Then Text3.Text = rstBindingTypeList.Fields("Col0").Value
        MhDateInput1.Text = Format(GetDate(FrmBookPrintOrder.MhDateInput1.Text), "dd-MM-yyyy")
        'MhDateInput3.Text = Format(DateAdd("d", 15, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
          
    Else
        LoadFields
    End If
    
    BusySystemIndicator False
    
    Exit Sub
    
ErrorHandler:

    BusySystemIndicator False
    Call CloseForm(Me)
    
End Sub

Private Sub LoadRefList(ByVal strOrderCode As String)
    On Error GoTo ErrorHandler
    If rstRefList.State = adStateOpen Then
        rstRefList.Close
    End If
    rstRefList.Open "Select Warehouse1,Warehouse2,Warehouse3  FROM PrintPVChild Where Code in(Select Ref From BookPOChild05 WHERE Code In(Select Code From BookPOParent Where TRIM(Name)='" & strOrderCode & "')) And Book in(Select  Book From BookPOParent Where TRIM(Name)='" & strOrderCode & "')", CxnDatabase, adOpenKeyset, adLockOptimistic
       
    rstRefList.ActiveConnection = Nothing
    Exit Sub
ErrorHandler:
    DisplayError ("Failed to Load Ref List")
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
    BookPrinterQuantity = 0
    Call CloseRecordset(rstBindingTypeList)
    Call CloseRecordset(rstBinderRates)
End Sub
Private Sub ClearFields()
    MhDateInput1.Text = Format(Date, "dd-MM-yyyy")
    MhDateInput2.Text = "  -  -    "
    MhDateInput311.Text = "  -  -    "
    MhDateInput3.Text = Format(DateAdd("d", 15, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
    Text3.Text = ""
    Text211.Text = ""
    Text6.Text = ""
    Text8.Text = ""
    MhRealInput1.Text = "0"
    MhRealInput2.Text = "0.00"
    MhRealInput3.Text = "0"
    MhRealInput4.Text = "0"
    MhRealInput5.Text = "0.00"
    MhRealInput6.Text = "0.00"
    MhRealInput7.Text = "0"
    MhRealInput8.Text = "0.00"
    MhRealInput9.Text = "0.00"
    MhRealInput10.Text = "0.00"
    MhRealInput11.Text = "0"
    MhRealInput12.Text = "0"
    MhRealInput13.Text = "0.00"
    MhRealInput14.Text = "0.00"
    MhRealInput15.Text = "0.00"
    MhRealInput16.Text = "0.00"
    MhRealInput17.Text = "0.00"
    MhRealInput18.Text = "0"
    MhRealInput19.Text = "0"
    MhRealInput20.Text = "0.00"
    MhRealInput21.Text = "0.00"
    MhRealInput22.Text = "0.00"
    MhRealInput23.Text = "0"
    MhRealInput24.Text = "0"
    MhRealInput25.Text = "0.00"
    MhRealInput26.Text = "0.00"
    MhRealInput27.Text = "0.00"
    MhRealInput28.Text = "0.00"
    MhRealInput101.Text = "0.000"
    MhRealInput29.Text = "0"
    MhRealInput30.Text = "0"
    MhRealInput31.Text = "0"
    TxtAdNar.Text = ""
    
End Sub
Private Sub LoadFields()
    
    If rstBookPOChild08.RecordCount = 0 Then Exit Sub
    MhDateInput1.Text = Format(rstBookPOChild08.Fields("OrderDate").Value, "dd-MM-yyyy")
    MhDateInput3.Text = Format(rstBookPOChild08.Fields("TargetDate").Value, "dd-MM-yyyy")
    
    If Not IsNull(rstBookPOChild08.Fields("ExtendDate").Value) Then MhDateInput31.Text = Format(rstBookPOChild08.Fields("ExtendDate").Value, "dd-MM-yyyy")
    
    If Not IsNull(rstBookPOChild08.Fields("AdvanceRecvdDate").Value) Then MhDateInput311.Text = Format(rstBookPOChild08.Fields("AdvanceRecvdDate").Value, "dd-MM-yyyy")
    
    MhRealInput7.Text = Format(Val(rstBookPOChild08.Fields("BindingForms").Value), "0")
    MhRealInput18.Text = Format(Val(rstBookPOChild08.Fields("ExtraForms").Value), "0")
    MhRealInput1.Text = Format(Val(rstBookPOChild08.Fields("ActualQuantity").Value), "0")
    MhRealInput3.Text = Format(Val(rstBookPOChild08.Fields("BillingQuantity").Value), "0")
    MhRealInput4.Text = Format(Val(rstBookPOChild08.Fields("AdjustQuantity").Value), "0")
    MhRealInput2.Text = Format(Val(rstBookPOChild08.Fields("FormFoldRate").Value), "0.00")
    MhRealInput5.Text = Format(Val(rstBookPOChild08.Fields("FormStitchRate").Value), "0.00")
    MhRealInput6.Text = Format(Val(rstBookPOChild08.Fields("FormPasteRate").Value), "0.00")
    MhRealInput8.Text = Format(Val(rstBookPOChild08.Fields("Rate/Book").Value), "0.00")
    MhRealInput29.Text = Format(Val(rstBookPOChild08.Fields("LooseQty/Box").Value), "0")
    MhRealInput30.Text = Format(Val(rstBookPOChild08.Fields("ExtraLooseQty").Value), "0")
    MhRealInput31.Text = Format(Val(rstBookPOChild08.Fields("TotalLooseQty").Value), "0")
    MhRealInput23.Text = Format(Val(rstBookPOChild08.Fields("Qty/Pkt").Value), "0")
    MhRealInput24.Text = Format(Val(rstBookPOChild08.Fields("TotalPkts").Value), "0")
    MhRealInput11.Text = Format(Val(rstBookPOChild08.Fields("Pkt/Box").Value), "0")
    MhRealInput12.Text = Format(Val(rstBookPOChild08.Fields("TotalBoxes").Value), "0")
    MhRealInput25.Text = Format(Val(rstBookPOChild08.Fields("PktPackRate").Value), "0.00")
    MhRealInput26.Text = Format(Val(rstBookPOChild08.Fields("BoxPackRate").Value), "0.00")
    MhRealInput13.Text = Format(Val(rstBookPOChild08.Fields("CartageRate").Value), "0.00")
    MhRealInput9.Text = Format(Val(rstBookPOChild08.Fields("Adjustment").Value), "0.00")
    MhRealInput10.Text = Format(Val(rstBookPOChild08.Fields("BillAmount").Value), "0.00")
          
    If IsNull(rstBookPOChild08.Fields("UnitCost").Value) Or rstBookPOChild08.Fields("UnitCost").Value = "0" Then
       'MhRealInput101.Text = Format(Val(MhRealInput10.Text) / Val(MhRealInput3.Text), "0.000")   'Comment On 17 Dec 15
    Else
        MhRealInput101.Text = Format(Val(rstBookPOChild08.Fields("UnitCost").Value), "0.000") 'Unit Cost
    End If
    
    If rstBookPOChild08.Fields("BookEdition").Value <> "" Then Text211.Text = rstBookPOChild08.Fields("BookEdition").Value
    
    BindingTypeCode = rstBookPOChild08.Fields("BindingType").Value
    
    
    If rstBindingTypeList.RecordCount > 0 Then rstBindingTypeList.MoveFirst
    rstBindingTypeList.Find "[Code] = '" & BindingTypeCode & "'"
    If Not rstBindingTypeList.EOF Then
       Text3.Text = rstBindingTypeList.Fields("Col0").Value
    End If
    Text8.Text = rstBookPOChild08.Fields("BillNo").Value
    If Not IsNull(rstBookPOChild08.Fields("BillDate").Value) Then MhDateInput2.Text = Format(rstBookPOChild08.Fields("BillDate").Value, "dd-MM-yyyy")
    
    MhRealInput15.Text = Format(Val(rstBookPOChild08.Fields("VAT%").Value), "0.00")
    MhRealInput17.Text = Format(Val(rstBookPOChild08.Fields("VAT").Value), "0.00")
    MhRealInput16.Text = Format(Val(rstBookPOChild08.Fields("PaidAmount").Value), "0.00")
    Text6.Text = rstBookPOChild08.Fields("Remarks").Value
    TxtAdNar.Text = rstBookPOChild08.Fields("AdjustmentRemarks").Value
    
   
    If rstRefList.RecordCount > 0 Then
        
        If Not IsNull(rstRefList.Fields("Warehouse1").Value) Then
                MhRealInputN.Text = rstRefList.Fields("Warehouse1").Value
        End If
        
        If Not IsNull(rstRefList.Fields("Warehouse2").Value) Then
            MhRealInputD.Text = rstRefList.Fields("Warehouse2").Value
        End If
        If Not IsNull(rstRefList.Fields("Warehouse3").Value) Then
            MhRealInputE.Text = rstRefList.Fields("Warehouse3").Value
        End If
        
   End If
    
    If rstBookPOChild08.Fields("AdvanceCopyRequired").Value <> "" Then
        Combo1.Text = rstBookPOChild08.Fields("AdvanceCopyRequired").Value
    End If
    
    CalculateTotalAmount
    
End Sub
Private Sub SaveFields()
    
    rstBookPOChild08.Fields("OrderDate").Value = GetDate(MhDateInput1.Text)
    rstBookPOChild08.Fields("TargetDate").Value = GetDate(MhDateInput3.Text)
    
    If Not IsDate(MhDateInput31.Text) Then rstBookPOChild08.Fields("ExtendDate").Value = Null Else rstBookPOChild08.Fields("ExtendDate").Value = GetDate(MhDateInput31.Text)
    
    If Not IsDate(MhDateInput311.Text) Then rstBookPOChild08.Fields("AdvanceRecvdDate").Value = Null Else rstBookPOChild08.Fields("AdvanceRecvdDate").Value = GetDate(MhDateInput311.Text)
    rstBookPOChild08.Fields("AdvanceCopyRequired").Value = Combo1.Text
    
    
    rstBookPOChild08.Fields("BindingType").Value = BindingTypeCode
    rstBookPOChild08.Fields("BindingForms").Value = Format(Val(MhRealInput7.Text), "0")
    rstBookPOChild08.Fields("ExtraForms").Value = Format(Val(MhRealInput18.Text), "0")
    rstBookPOChild08.Fields("ActualQuantity").Value = Format(Val(MhRealInput1.Text), "0")
    rstBookPOChild08.Fields("BillingQuantity").Value = Format(Val(MhRealInput3.Text), "0")
    rstBookPOChild08.Fields("AdjustQuantity").Value = Format(Val(MhRealInput4.Text), "0")
    rstBookPOChild08.Fields("FormFoldRate").Value = Format(Val(MhRealInput2.Text), "0.00")
    rstBookPOChild08.Fields("FormStitchRate").Value = Format(Val(MhRealInput5.Text), "0.00")
    rstBookPOChild08.Fields("FormPasteRate").Value = Format(Val(MhRealInput6.Text), "0.00")
    rstBookPOChild08.Fields("Rate/Book").Value = Format(Val(MhRealInput8.Text), "0.00")
    rstBookPOChild08.Fields("LooseQty/Box").Value = Format(Val(MhRealInput29.Text), "0")
    rstBookPOChild08.Fields("ExtraLooseQty").Value = Format(Val(MhRealInput30.Text), "0")
    rstBookPOChild08.Fields("TotalLooseQty").Value = Format(Val(MhRealInput31.Text), "0")
    rstBookPOChild08.Fields("Qty/Pkt").Value = Format(Val(MhRealInput23.Text), "0")
    rstBookPOChild08.Fields("TotalPkts").Value = Format(Val(MhRealInput24.Text), "0")
    rstBookPOChild08.Fields("Pkt/Box").Value = Format(Val(MhRealInput11.Text), "0")
    rstBookPOChild08.Fields("TotalBoxes").Value = Format(Val(MhRealInput12.Text), "0")
    rstBookPOChild08.Fields("PktPackRate").Value = Format(Val(MhRealInput25.Text), "0.00")
    rstBookPOChild08.Fields("BoxPackRate").Value = Format(Val(MhRealInput26.Text), "0.00")
    rstBookPOChild08.Fields("CartageRate").Value = Format(Val(MhRealInput13.Text), "0.00")
    rstBookPOChild08.Fields("Adjustment").Value = Format(Val(MhRealInput9.Text), "0.00")
    rstBookPOChild08.Fields("BillAmount").Value = Format(Val(MhRealInput10.Text), "0.00")
    rstBookPOChild08.Fields("BillNo").Value = Text8.Text
      
    If Not IsDate(MhDateInput2.Text) Then rstBookPOChild08.Fields("BillDate").Value = Null Else rstBookPOChild08.Fields("BillDate").Value = GetDate(MhDateInput2.Text)
        
    rstBookPOChild08.Fields("VAT%").Value = Format(Val(MhRealInput15.Text), "0.00")
    rstBookPOChild08.Fields("VAT").Value = Format(Val(MhRealInput17.Text), "0.00")
    rstBookPOChild08.Fields("PaidAmount").Value = Format(Val(MhRealInput16.Text), "0.00")
    rstBookPOChild08.Fields("Remarks").Value = Text6.Text
    rstBookPOChild08.Fields("AdjustmentRemarks").Value = IIf(Val(MhRealInput9.Text) <> 0, TxtAdNar.Text, "")
    If Not CheckEmpty(Text8.Text, False) Then If IsNull(rstBookPOChild08.Fields("BillFeedDate").Value) Then rstBookPOChild08.Fields("BillFeedDate").Value = Now()
    Dim lpBuff As String * 1024
    GetComputerName lpBuff, Len(lpBuff)
    If Not CheckEmpty(Text8.Text, False) Then If IsNull(rstBookPOChild08.Fields("ComputerName").Value) Then rstBookPOChild08.Fields("ComputerName").Value = Left(lpBuff, (InStr(1, lpBuff, vbNullChar)) - 1)
    rstBookPOChild08.Fields("UnitCost").Value = Format(Val(MhRealInput101.Text), "0.000")
    rstBookPOChild08.Fields("BookEdition").Value = Text211.Text
      
    
End Sub
Private Sub MhDateInput1_Validate(Cancel As Boolean)
    
    If Not IsDate(GetDate(MhDateInput1.Text)) Then
        Cancel = True
    ElseIf Format(GetDate(MhDateInput1.Text), "yyyymmdd") < Format(FinancialYearFrom, "yyyymmdd") Or Format(GetDate(MhDateInput1.Text), "yyyymmdd") > Format(FinancialYearTo, "yyyymmdd") Then
        Cancel = True
    ElseIf Val(CheckNull(rstBookPOChild08.Fields("ActualQuantity").Value)) = 0 Then
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

Private Sub Text211_Validate(Cancel As Boolean)
    If CheckEmpty(Text211, True) Then
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
        Call DisplaySelectionList(Text3, BindingTypeCode)
        Call CloseForm(FrmSelectionList)
        If CheckEmpty(Text3.Text, False) Then
            Text3.Text = "?"
        End If
        If RTrim(BindingTypeCode) <> "" Then
            SendKeys "{TAB}"
        End If
        Cancel = True
    Else
        BindingTypeCode = rstBindingTypeList.Fields("Code").Value
        Call GetBinderRates: CalculateTotalAmount
    End If
End Sub
Private Sub MhRealInput7_Validate(Cancel As Boolean)
    CalculateTotalAmount
    
    
   If MhRealInput7.Value > 0 And MhRealInput1.Value > 0 Then
         'Set Target Date based on Quantity and Form
         If MhRealInput7.Value <= 120 And MhRealInput1.Value <= 3300 Then
            MhDateInput3.Text = Format(DateAdd("d", 10, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
         ElseIf MhRealInput7.Value <= 120 And MhRealInput1.Value > 3300 And MhRealInput1.Value <= 5500 Then
            MhDateInput3.Text = Format(DateAdd("d", 12, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
         ElseIf MhRealInput7.Value <= 120 And MhRealInput1.Value > 5500 And MhRealInput1.Value <= 10500 Then
            MhDateInput3.Text = Format(DateAdd("d", 15, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
         ElseIf MhRealInput7.Value <= 120 And MhRealInput1.Value > 10500 Then
            MhDateInput3.Text = Format(DateAdd("d", 20, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
         
         ElseIf MhRealInput7.Value > 121 And MhRealInput7.Value <= 200 And MhRealInput1.Value <= 3300 Then
            MhDateInput3.Text = Format(DateAdd("d", 10, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
         ElseIf MhRealInput7.Value > 121 And MhRealInput7.Value <= 200 And MhRealInput1.Value > 3300 And MhRealInput1.Value <= 5500 Then
            MhDateInput3.Text = Format(DateAdd("d", 12, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
         ElseIf MhRealInput7.Value > 121 And MhRealInput7.Value <= 200 And MhRealInput1.Value > 5500 And MhRealInput1.Value <= 10500 Then
            MhDateInput3.Text = Format(DateAdd("d", 18, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
         ElseIf MhRealInput7.Value > 121 And MhRealInput7.Value <= 200 And MhRealInput1.Value > 10500 Then
            MhDateInput3.Text = Format(DateAdd("d", 20, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
         ElseIf MhRealInput7.Value > 200 And MhRealInput1.Value <= 3300 Then
            MhDateInput3.Text = Format(DateAdd("d", 10, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
         ElseIf MhRealInput7.Value > 200 And MhRealInput1.Value > 3300 And MhRealInput1.Value < 5500 Then
            MhDateInput3.Text = Format(DateAdd("d", 15, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
         ElseIf MhRealInput7.Value > 200 And MhRealInput1.Value > 5500 And MhRealInput1.Value < 10500 Then
            MhDateInput3.Text = Format(DateAdd("d", 20, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
         ElseIf MhRealInput7.Value > 200 And MhRealInput1.Value > 10500 Then
            MhDateInput3.Text = Format(DateAdd("d", 25, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
         End If
     End If
    
End Sub
Private Sub MhRealInput1_Validate(Cancel As Boolean)
    If Val(MhRealInput3.Text) = 0 Then MhRealInput3.Text = Format(Val(MhRealInput1.Text), "0"): Exit Sub
    If Val(MhRealInput3.Text) <> Val(MhRealInput1.Text) Then If MsgBox("Alter billing quantity?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Alter !") = vbYes Then MhRealInput3.Text = Format(Val(MhRealInput1.Text), "0")
    
    If MhRealInput7.Value > 0 And MhRealInput1.Value > 0 Then
         'Set Target Date based on Quantity and Form
         If MhRealInput7.Value <= 120 And MhRealInput1.Value <= 3300 Then
            MhDateInput3.Text = Format(DateAdd("d", 10, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
         ElseIf MhRealInput7.Value <= 120 And MhRealInput1.Value > 3300 And MhRealInput1.Value <= 5500 Then
            MhDateInput3.Text = Format(DateAdd("d", 12, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
         ElseIf MhRealInput7.Value <= 120 And MhRealInput1.Value > 5500 And MhRealInput1.Value <= 10500 Then
            MhDateInput3.Text = Format(DateAdd("d", 15, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
         ElseIf MhRealInput7.Value <= 120 And MhRealInput1.Value > 10500 Then
            MhDateInput3.Text = Format(DateAdd("d", 20, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
         ElseIf MhRealInput7.Value > 121 And MhRealInput7.Value <= 200 And MhRealInput1.Value <= 3300 Then
            MhDateInput3.Text = Format(DateAdd("d", 10, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
         ElseIf MhRealInput7.Value > 121 And MhRealInput7.Value <= 200 And MhRealInput1.Value > 3300 And MhRealInput1.Value <= 5500 Then
            MhDateInput3.Text = Format(DateAdd("d", 12, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
         ElseIf MhRealInput7.Value > 121 And MhRealInput7.Value <= 200 And MhRealInput1.Value > 5500 And MhRealInput1.Value <= 10500 Then
            MhDateInput3.Text = Format(DateAdd("d", 18, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
         ElseIf MhRealInput7.Value > 121 And MhRealInput7.Value <= 200 And MhRealInput1.Value > 10500 Then
            MhDateInput3.Text = Format(DateAdd("d", 20, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
         ElseIf MhRealInput7.Value > 200 And MhRealInput1.Value <= 3300 Then
            MhDateInput3.Text = Format(DateAdd("d", 10, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
         ElseIf MhRealInput7.Value > 200 And MhRealInput1.Value > 3300 And MhRealInput1.Value < 5500 Then
            MhDateInput3.Text = Format(DateAdd("d", 15, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
         ElseIf MhRealInput7.Value > 200 And MhRealInput1.Value > 5500 And MhRealInput1.Value < 10500 Then
            MhDateInput3.Text = Format(DateAdd("d", 20, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
         ElseIf MhRealInput7.Value > 200 And MhRealInput1.Value > 10500 Then
            MhDateInput3.Text = Format(DateAdd("d", 25, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
         End If
    End If
End Sub
Private Sub MhRealInput3_Validate(Cancel As Boolean)
    CalculateBundle
    CalculateTotalAmount
End Sub
Private Sub MhRealInput29_Validate(Cancel As Boolean)   'Loose/Box
    CalculateBundle
End Sub
Private Sub MhRealInput30_Validate(Cancel As Boolean)   'Extra Loose
    CalculateBundle
End Sub
Private Sub MhRealInput23_Validate(Cancel As Boolean)   'Qty/Pkt
    CalculateBundle
End Sub
Private Sub MhRealInput11_Validate(Cancel As Boolean)   'Pkt/Box
    CalculateBundle
End Sub
Private Sub MhRealInput2_Validate(Cancel As Boolean)
    CalculateTotalAmount
End Sub
Private Sub MhRealInput5_Validate(Cancel As Boolean)
    CalculateTotalAmount
End Sub
Private Sub MhRealInput6_Validate(Cancel As Boolean)
    CalculateTotalAmount
End Sub
Private Sub MhRealInput8_Validate(Cancel As Boolean)
    CalculateTotalAmount
End Sub
Private Sub MhRealInput25_Validate(Cancel As Boolean)
    CalculateTotalAmount
End Sub
Private Sub MhRealInput26_Validate(Cancel As Boolean)
    CalculateTotalAmount
End Sub
Private Sub MhRealInput13_Validate(Cancel As Boolean)
    CalculateTotalAmount
End Sub
Private Sub MhRealInput14_Validate(Cancel As Boolean)
    CalculateTotalAmount
End Sub
Private Sub MhRealInput15_Validate(Cancel As Boolean)
    CalculateTotalAmount
End Sub
Private Sub MhRealInput9_Validate(Cancel As Boolean)
    CalculateTotalAmount
End Sub
Private Sub CalculateBundle()
    
    Dim CalcPkt As Long, CalcBox As Long, CalcLoose As Long, TotalBox As Long
    'Total box Calculation
    If Val(MhRealInput23.Text) * Val(MhRealInput11.Text) + Val(MhRealInput29.Text) > 0 Then TotalBox = Int((Val(MhRealInput3.Text) - Val(MhRealInput30.Text)) / (Val(MhRealInput23.Text) * Val(MhRealInput11.Text) + Val(MhRealInput29.Text)))   'int((billing quantity - extra loose quantity) / quantity per box)
    MhRealInput12.Text = Format(TotalBox, "0")
    'Total Loose Calculation
    If Val(MhRealInput23.Text) > 0 Then 'qty per packet
        CalcLoose = Val(MhRealInput30.Text) Mod Val(MhRealInput23.Text) 'Loose qty remaining from extra loose qty after packet making
        CalcLoose = CalcLoose + (Val(MhRealInput3.Text) - Val(MhRealInput30.Text) - (TotalBox * Val(MhRealInput29.Text))) Mod Val(MhRealInput23.Text)
    End If
    CalcLoose = CalcLoose + TotalBox * Val(MhRealInput29.Text)
    MhRealInput31.Text = Format(CalcLoose, "0")
    'Total Packet Calculation
    If Val(MhRealInput23.Text) > 0 Then 'qty per packet
        CalcPkt = Int(Val(MhRealInput30.Text) / Val(MhRealInput23.Text))
        CalcPkt = CalcPkt + Int((Val(MhRealInput3.Text) - Val(MhRealInput30.Text) - (TotalBox * Val(MhRealInput29.Text))) / Val(MhRealInput23.Text))
    End If
    MhRealInput24.Text = Format(CalcPkt, "0")
    
    CalculateTotalAmount
    
End Sub
Private Sub CalculateTotalAmount()
    
    MhRealInput20.Text = Format((Val(MhRealInput2.Text) * Val(MhRealInput3.Text) * (Val(MhRealInput7.Text) + Val(MhRealInput18.Text))) / 1000, "0.00") 'Folding Amount
    MhRealInput21.Text = Format((Val(MhRealInput5.Text) * Val(MhRealInput3.Text) * (Val(MhRealInput7.Text) + Val(MhRealInput18.Text))) / 1000, "0.00") 'Stitching Amount
    MhRealInput22.Text = Format((Val(MhRealInput6.Text) * Val(MhRealInput3.Text)) / 1000, "0.00") 'Pasting Amount
    MhRealInput27.Text = Format(Val(MhRealInput24.Text) * Val(MhRealInput25.Text), "0.00")  'Pkt Packing Amount
    MhRealInput28.Text = Format(Val(MhRealInput12.Text) * Val(MhRealInput26.Text), "0.00")  'Box Packing Amount
    MhRealInput14.Text = Format(Val(MhRealInput12.Text) * Val(MhRealInput13.Text), "0.00")  'Cartage
    'MhRealInput17.Text = Format(((Val(MhRealInput20.Text) + Val(MhRealInput21.Text) + Val(MhRealInput22.Text) + Val(MhRealInput27.Text) + Val(MhRealInput28.Text) + (Val(MhRealInput8.Text) * Val(MhRealInput3.Text))) * 80 / 100) * Val(MhRealInput15.Text) / 100, "0.00") 'VAT
    MhRealInput17.Text = Format(((Val(MhRealInput20.Text) + Val(MhRealInput21.Text) + Val(MhRealInput22.Text) + Val(MhRealInput27.Text) + Val(MhRealInput28.Text) + (Val(MhRealInput8.Text) * Val(MhRealInput3.Text))) * 100 / 100) * Val(MhRealInput15.Text) / 100, "0.00") 'VAT
    MhRealInput10.Text = Format(Val(MhRealInput20.Text) + Val(MhRealInput21.Text) + Val(MhRealInput22.Text) + Val(MhRealInput27.Text) + Val(MhRealInput28.Text) + (Val(MhRealInput8.Text) * Val(MhRealInput3.Text)) + Val(MhRealInput14.Text) + Val(MhRealInput17.Text) + Val(MhRealInput9.Text), "0.00") 'Total Amount
    
    If Val(MhRealInput10.Text) > 0 Then
      MhRealInput101.Text = Format(Val(MhRealInput10.Text) / Val(MhRealInput3.Text), "0.000") 'Unit Cost
    End If
    
End Sub
Private Sub cmdProceed_Click()
    
    If CheckMandatoryFields Then Exit Sub
    SaveFields
    
    rstBookPOChild08.Update
    FrmBookPOChild0801.BookQuantity = Val(MhRealInput1.Text)
    Load FrmBookPOChild0801
    FrmBookPOChild0801.Show vbModal
    Call CloseForm(Me)
    
End Sub
Private Sub cmdCancel_Click()
    rstBookPOChild08.CancelUpdate
    Call CloseForm(Me)
End Sub
Private Function CheckMandatoryFields() As Boolean
    
    If CheckEmpty(Text3.Text, False) Then Text3.SetFocus: CheckMandatoryFields = True: Exit Function
    If Not CheckExists(Text3, "Col0", rstBindingTypeList, BindingTypeCode) Then Text3.SetFocus: CheckMandatoryFields = True: Exit Function
    If Val(MhRealInput16.Text) <> 0 Then If Val(MhRealInput16.Text) <> Val(MhRealInput10.Text) Then MhRealInput9.SetFocus: CheckMandatoryFields = True: Exit Function
    If Val(MhRealInput9.Text) <> 0 Then If CheckEmpty(TxtAdNar.Text, False) Then TxtAdNar.SetFocus: CheckMandatoryFields = True: Exit Function

End Function
Private Sub GetBinderRates()
    
    Dim FoldingRate As Double, StitchingRate As Double, PastingRate As Double, RPB As Double, PktPackRate As Double, BoxPackRate As Double
    On Error GoTo ErrorHandler
    If rstBinderRates.State = adStateOpen Then rstBinderRates.Close
    rstBinderRates.Open "Select Top 1 * From AccountChild08 Where Code = '" & BinderCode & "' And [Size] = '" & SizeCode & "' And BindingType = '" & BindingTypeCode & "' And Range" & IIf(FormType = "1", "08", IIf(FormType = "2", "16", IIf(FormType = "3", "04", IIf(FormType = "4", "12", IIf(FormType = "5", "24", IIf(FormType = "6", "32", "64")))))) & " >= " & Val(MhRealInput7.Text) + Val(MhRealInput18.Text) & " Order By Range" & IIf(FormType = "1", "08", IIf(FormType = "2", "16", IIf(FormType = "3", "04", IIf(FormType = "4", "12", IIf(FormType = "5", "24", IIf(FormType = "6", "32", "64")))))), CxnDatabase, adOpenKeyset, adLockReadOnly
    If rstBinderRates.RecordCount = 0 Then
        If rstBinderRates.State = adStateOpen Then rstBinderRates.Close
        rstBinderRates.Open "Select Top 1 * From AccountMaster,AccountChild08 Where AccountMaster.Code = AccountChild08.Code And [Name] Like '%Rate%' And [Size] = '" & SizeCode & "' And BindingType = '" & BindingTypeCode & "' And Range" & IIf(FormType = "1", "08", IIf(FormType = "2", "16", IIf(FormType = "3", "04", IIf(FormType = "4", "12", IIf(FormType = "5", "24", IIf(FormType = "6", "32", "64")))))) & " >= " & Val(MhRealInput7.Text) + Val(MhRealInput18.Text) & " Order By Range" & IIf(FormType = "1", "08", IIf(FormType = "2", "16", IIf(FormType = "3", "04", IIf(FormType = "4", "12", IIf(FormType = "5", "24", IIf(FormType = "6", "32", "64")))))), CxnDatabase, adOpenKeyset, adLockReadOnly
    End If
    If rstBinderRates.RecordCount > 0 Then
        FoldingRate = rstBinderRates.Fields("FormFoldRate" & IIf(FormType = "1", "08", IIf(FormType = "2", "16", IIf(FormType = "3", "04", IIf(FormType = "4", "12", IIf(FormType = "5", "24", IIf(FormType = "6", "32", "64"))))))).Value
        StitchingRate = rstBinderRates.Fields("FormStitchRate" & IIf(FormType = "1", "08", IIf(FormType = "2", "16", IIf(FormType = "3", "04", IIf(FormType = "4", "12", IIf(FormType = "5", "24", IIf(FormType = "6", "32", "64"))))))).Value
        PastingRate = rstBinderRates.Fields("FormPasteRate" & IIf(FormType = "1", "08", IIf(FormType = "2", "16", IIf(FormType = "3", "04", IIf(FormType = "4", "12", IIf(FormType = "5", "24", IIf(FormType = "6", "32", "64"))))))).Value
        If Val(MhRealInput7.Text) + Val(MhRealInput18.Text) > 25 Then PastingRate = PastingRate + IIf(PastingRate > 0, (Val(MhRealInput7.Text) + Val(MhRealInput18.Text) - 25) * Val(FrmBookPrintOrder.rstBookList.Fields("AddOnRate02").Value) * 1000, 0)
        RPB = rstBinderRates.Fields("Rate/Book" & IIf(FormType = "1", "08", IIf(FormType = "2", "16", IIf(FormType = "3", "04", IIf(FormType = "4", "12", IIf(FormType = "5", "24", IIf(FormType = "6", "32", "64"))))))).Value
        PktPackRate = rstBinderRates.Fields("PktPackRate").Value
        BoxPackRate = rstBinderRates.Fields("BoxPackRate").Value
    End If
    If Val(MhRealInput2.Text) <> FoldingRate And Val(MhRealInput2.Text) <> 0 Then
        If MsgBox("Folding Rate is different from that in Master ! Change Rate?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput2.Text = Format(FoldingRate, "0.00")
    Else
        MhRealInput2.Text = Format(FoldingRate, "0.00")
    End If
    If Val(MhRealInput5.Text) <> StitchingRate And Val(MhRealInput5.Text) <> 0 Then
        If MsgBox("Stitching Rate is different from that in Master ! Change Rate?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput5.Text = Format(StitchingRate, "0.00")
    Else
        MhRealInput5.Text = Format(StitchingRate, "0.00")
    End If
    If Val(MhRealInput6.Text) <> PastingRate And Val(MhRealInput6.Text) <> 0 Then
        If MsgBox("Pasting Rate is different from that in Master ! Change Rate?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput6.Text = Format(PastingRate, "0.00")
    Else
        MhRealInput6.Text = Format(PastingRate, "0.00")
    End If
    If Val(MhRealInput8.Text) <> RPB And Val(MhRealInput8.Text) <> 0 Then
        If MsgBox("Rate/Book is different from that in Master ! Change Rate?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput8.Text = Format(RPB, "0.00")
    Else
        MhRealInput8.Text = Format(RPB, "0.00")
    End If
    If Val(MhRealInput25.Text) <> PktPackRate And Val(MhRealInput25.Text) <> 0 Then
        If MsgBox("Pkt Packing Rate is different from that in Master ! Change Rate?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput25.Text = Format(PktPackRate, "0.00")
    Else
        MhRealInput25.Text = Format(PktPackRate, "0.00")
    End If
    If Val(MhRealInput26.Text) <> BoxPackRate And Val(MhRealInput26.Text) <> 0 Then
        If MsgBox("Box Packing Rate is different from that in Master ! Change Rate?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then MhRealInput26.Text = Format(BoxPackRate, "0.00")
    Else
        MhRealInput26.Text = Format(BoxPackRate, "0.00")
    End If
    Exit Sub
ErrorHandler:
    DisplayError ("Failed to Fetch Binder Rates")
End Sub

