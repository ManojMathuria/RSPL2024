VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmChallan 
   Caption         =   "Delivery Challan"
   ClientHeight    =   8640
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11235
   Icon            =   "FrmChallan.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8640
   ScaleWidth      =   11235
   StartUpPosition =   2  'CenterScreen
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   8610
      Left            =   0
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   0
      Width           =   11190
      _Version        =   65536
      _ExtentX        =   19738
      _ExtentY        =   15187
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
      Picture         =   "FrmChallan.frx":0442
      Begin TabDlg.SSTab SSTab1 
         Height          =   8430
         Left            =   120
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   120
         Width           =   10965
         _ExtentX        =   19341
         _ExtentY        =   14870
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
         TabPicture(0)   =   "FrmChallan.frx":045E
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Text1"
         Tab(0).Control(1)=   "DataGrid1"
         Tab(0).Control(2)=   "Label1(1)"
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "&Details"
         TabPicture(1)   =   "FrmChallan.frx":047A
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
            TabIndex        =   20
            Top             =   7950
            Width           =   10260
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   7425
            Left            =   -74880
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   450
            Width           =   10740
            _ExtentX        =   18944
            _ExtentY        =   13097
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            BackColor       =   16776960
            HeadLines       =   1
            RowHeight       =   18
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
               DataField       =   ""
               Caption         =   ""
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
               DataField       =   ""
               Caption         =   ""
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
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   7980
            Left            =   120
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   360
            Width           =   10740
            _Version        =   65536
            _ExtentX        =   18944
            _ExtentY        =   14076
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
            Picture         =   "FrmChallan.frx":0496
            Begin VB.TextBox txtRefNo 
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
               Left            =   8300
               MaxLength       =   139
               TabIndex        =   17
               Top             =   4155
               Width           =   2345
            End
            Begin VB.TextBox txtTransitFormNumber 
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
               Left            =   3120
               MaxLength       =   139
               TabIndex        =   16
               Top             =   4155
               Width           =   2805
            End
            Begin VB.TextBox txtPONumber 
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
               Left            =   9360
               MaxLength       =   139
               TabIndex        =   2
               Top             =   45
               Width           =   1275
            End
            Begin VB.TextBox txtEWayBillNo 
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
               Left            =   3120
               MaxLength       =   139
               TabIndex        =   14
               Top             =   3840
               Width           =   2805
            End
            Begin VB.TextBox txtDestinationAddress 
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
               Left            =   3120
               MaxLength       =   139
               TabIndex        =   13
               Top             =   3525
               Width           =   7515
            End
            Begin VB.TextBox txtDestinationGoods 
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
               Left            =   8300
               MaxLength       =   139
               TabIndex        =   12
               Top             =   3210
               Width           =   2345
            End
            Begin VB.TextBox txtVehicleNo 
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
               Left            =   3120
               MaxLength       =   139
               TabIndex        =   11
               Top             =   3210
               Width           =   2785
            End
            Begin VB.TextBox txtCarrierAddress 
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
               Left            =   3120
               MaxLength       =   139
               TabIndex        =   10
               Top             =   2895
               Width           =   7515
            End
            Begin VB.TextBox txtCarrierName 
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
               Left            =   3120
               MaxLength       =   139
               TabIndex        =   9
               Top             =   2580
               Width           =   7515
            End
            Begin VB.TextBox txtConsigneeRegNo 
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
               Left            =   3120
               MaxLength       =   139
               TabIndex        =   8
               Top             =   2145
               Width           =   7515
            End
            Begin VB.TextBox txtConsigneeAddress 
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
               Left            =   3120
               MaxLength       =   139
               TabIndex        =   7
               Top             =   1830
               Width           =   7515
            End
            Begin VB.TextBox txtConsigneeName 
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
               Left            =   3120
               MaxLength       =   139
               TabIndex        =   6
               Text            =   "Supplier Name will come against P.O No."
               Top             =   1515
               Width           =   7515
            End
            Begin VB.TextBox txtConsignerRegNo 
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
               Left            =   3120
               MaxLength       =   139
               TabIndex        =   5
               Top             =   1080
               Width           =   7515
            End
            Begin VB.TextBox txtConsignerName 
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
               Left            =   3120
               MaxLength       =   40
               TabIndex        =   3
               Text            =   "Rachna Sagar Pvt. Ltd."
               Top             =   450
               Width           =   7515
            End
            Begin VB.TextBox txtConsignerAddress 
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
               Left            =   3120
               MaxLength       =   139
               TabIndex        =   4
               Top             =   770
               Width           =   7515
            End
            Begin VB.TextBox txtChallan 
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
               Top             =   50
               Width           =   1580
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
               Height          =   330
               Left            =   120
               TabIndex        =   23
               Top             =   45
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
               Caption         =   " Challan No."
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "FrmChallan.frx":04B2
               Picture         =   "FrmChallan.frx":04CE
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Left            =   5895
               TabIndex        =   24
               Top             =   45
               Width           =   1140
               _Version        =   65536
               _ExtentX        =   2011
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
               Caption         =   " Challan Date"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "FrmChallan.frx":04EA
               Picture         =   "FrmChallan.frx":0506
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   330
               Left            =   120
               TabIndex        =   25
               Top             =   450
               Width           =   3015
               _Version        =   65536
               _ExtentX        =   5318
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
               Caption         =   " Consignor's / Owner Name"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "FrmChallan.frx":0522
               Picture         =   "FrmChallan.frx":053E
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
               Height          =   330
               Index           =   0
               Left            =   120
               TabIndex        =   26
               Top             =   770
               Width           =   3015
               _Version        =   65536
               _ExtentX        =   5318
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
               Caption         =   " Consignor's Address"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "FrmChallan.frx":055A
               Picture         =   "FrmChallan.frx":0576
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
               Height          =   330
               Left            =   7785
               TabIndex        =   27
               Top             =   6645
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
               Caption         =   " Sub Total"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "FrmChallan.frx":0592
               Picture         =   "FrmChallan.frx":05AE
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
               Height          =   330
               Index           =   0
               Left            =   7785
               TabIndex        =   28
               Top             =   7275
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
               Caption         =   " SGST @ 6.00%"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "FrmChallan.frx":05CA
               Picture         =   "FrmChallan.frx":05E6
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel15 
               Height          =   330
               Left            =   7785
               TabIndex        =   29
               Top             =   6960
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
               Caption         =   " CGST @ 6.00%"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "FrmChallan.frx":0602
               Picture         =   "FrmChallan.frx":061E
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel16 
               Height          =   330
               Left            =   7785
               TabIndex        =   30
               Top             =   6330
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
               Caption         =   " Cartage"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "FrmChallan.frx":063A
               Picture         =   "FrmChallan.frx":0656
            End
            Begin TDBDate6Ctl.TDBDate MhDateInput1 
               Height          =   330
               Left            =   7020
               TabIndex        =   1
               Top             =   45
               Width           =   975
               _Version        =   65536
               _ExtentX        =   1720
               _ExtentY        =   582
               Calendar        =   "FrmChallan.frx":0672
               Caption         =   "FrmChallan.frx":078A
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "FrmChallan.frx":07F6
               Keys            =   "FrmChallan.frx":0814
               Spin            =   "FrmChallan.frx":0872
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
               TabIndex        =   31
               Top             =   4880
               Width           =   10515
               _Version        =   524288
               _ExtentX        =   18547
               _ExtentY        =   1940
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
               MaxRows         =   1000
               ScrollBars      =   2
               SpreadDesigner  =   "FrmChallan.frx":089A
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput5 
               Height          =   330
               Left            =   9345
               TabIndex        =   32
               Top             =   6645
               Width           =   1290
               _Version        =   65536
               _ExtentX        =   2275
               _ExtentY        =   582
               Calculator      =   "FrmChallan.frx":11A7
               Caption         =   "FrmChallan.frx":11C7
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "FrmChallan.frx":1233
               Keys            =   "FrmChallan.frx":1251
               Spin            =   "FrmChallan.frx":129B
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
               ValueVT         =   1966145541
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput4 
               Height          =   330
               Left            =   9345
               TabIndex        =   33
               TabStop         =   0   'False
               Top             =   6330
               Width           =   1290
               _Version        =   65536
               _ExtentX        =   2275
               _ExtentY        =   582
               Calculator      =   "FrmChallan.frx":12C3
               Caption         =   "FrmChallan.frx":12E3
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "FrmChallan.frx":134F
               Keys            =   "FrmChallan.frx":136D
               Spin            =   "FrmChallan.frx":13B7
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
               ForeColor       =   255
               Format          =   "######0.00"
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
               ValueVT         =   1966145541
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput6 
               Height          =   330
               Left            =   9345
               TabIndex        =   34
               Top             =   6960
               Width           =   1290
               _Version        =   65536
               _ExtentX        =   2275
               _ExtentY        =   582
               Calculator      =   "FrmChallan.frx":13DF
               Caption         =   "FrmChallan.frx":13FF
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "FrmChallan.frx":146B
               Keys            =   "FrmChallan.frx":1489
               Spin            =   "FrmChallan.frx":14D3
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
            Begin TDBNumber6Ctl.TDBNumber MhRealInput7 
               Height          =   330
               Left            =   9345
               TabIndex        =   35
               TabStop         =   0   'False
               Top             =   7275
               Width           =   1290
               _Version        =   65536
               _ExtentX        =   2275
               _ExtentY        =   582
               Calculator      =   "FrmChallan.frx":14FB
               Caption         =   "FrmChallan.frx":151B
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "FrmChallan.frx":1587
               Keys            =   "FrmChallan.frx":15A5
               Spin            =   "FrmChallan.frx":15EF
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
               ForeColor       =   255
               Format          =   "######0.00"
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
               ValueVT         =   1966145541
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
               Height          =   285
               Left            =   120
               TabIndex        =   36
               Top             =   5950
               Width           =   10515
               _Version        =   65536
               _ExtentX        =   18547
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
               Picture         =   "FrmChallan.frx":1617
               Picture         =   "FrmChallan.frx":1633
               Begin TDBNumber6Ctl.TDBNumber MhRealInput1 
                  Height          =   285
                  Left            =   4680
                  TabIndex        =   37
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   1440
                  _Version        =   65536
                  _ExtentX        =   2549
                  _ExtentY        =   503
                  Calculator      =   "FrmChallan.frx":164F
                  Caption         =   "FrmChallan.frx":166F
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "FrmChallan.frx":16DB
                  Keys            =   "FrmChallan.frx":16F9
                  Spin            =   "FrmChallan.frx":1743
                  AlignHorizontal =   1
                  AlignVertical   =   0
                  Appearance      =   0
                  BackColor       =   16777215
                  BorderStyle     =   1
                  BtnPositioning  =   0
                  ClipMode        =   0
                  ClearAction     =   0
                  DecimalPoint    =   "."
                  DisplayFormat   =   "#####0.000"
                  EditMode        =   1
                  Enabled         =   -1
                  ErrorBeep       =   0
                  ForeColor       =   255
                  Format          =   "#####0.000"
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
                  ValueVT         =   5
                  Value           =   0
                  MaxValueVT      =   5
                  MinValueVT      =   5
               End
               Begin TDBNumber6Ctl.TDBNumber MhRealInput3 
                  Height          =   285
                  Left            =   9230
                  TabIndex        =   38
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   1005
                  _Version        =   65536
                  _ExtentX        =   1764
                  _ExtentY        =   503
                  Calculator      =   "FrmChallan.frx":176B
                  Caption         =   "FrmChallan.frx":178B
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "FrmChallan.frx":17F7
                  Keys            =   "FrmChallan.frx":1815
                  Spin            =   "FrmChallan.frx":185F
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
                  ForeColor       =   255
                  Format          =   "######0.00"
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
                  ValueVT         =   5
                  Value           =   0
                  MaxValueVT      =   5
                  MinValueVT      =   5
               End
               Begin TDBNumber6Ctl.TDBNumber MhRealInput2 
                  Height          =   285
                  Left            =   7345
                  TabIndex        =   39
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   1185
                  _Version        =   65536
                  _ExtentX        =   2081
                  _ExtentY        =   503
                  Calculator      =   "FrmChallan.frx":1887
                  Caption         =   "FrmChallan.frx":18A7
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  DropDown        =   "FrmChallan.frx":1913
                  Keys            =   "FrmChallan.frx":1931
                  Spin            =   "FrmChallan.frx":197B
                  AlignHorizontal =   1
                  AlignVertical   =   0
                  Appearance      =   0
                  BackColor       =   16777215
                  BorderStyle     =   1
                  BtnPositioning  =   0
                  ClipMode        =   0
                  ClearAction     =   0
                  DecimalPoint    =   "."
                  DisplayFormat   =   "#####0.000"
                  EditMode        =   1
                  Enabled         =   -1
                  ErrorBeep       =   0
                  ForeColor       =   255
                  Format          =   "#####0.000"
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
                  ValueVT         =   5
                  Value           =   0
                  MaxValueVT      =   5
                  MinValueVT      =   5
               End
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel8 
               Height          =   330
               Index           =   1
               Left            =   7785
               TabIndex        =   41
               Top             =   7590
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
               Caption         =   " Total (Rounded) :"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "FrmChallan.frx":19A3
               Picture         =   "FrmChallan.frx":19BF
            End
            Begin TDBNumber6Ctl.TDBNumber MhRealInput8 
               Height          =   330
               Left            =   9345
               TabIndex        =   42
               TabStop         =   0   'False
               Top             =   7590
               Width           =   1290
               _Version        =   65536
               _ExtentX        =   2275
               _ExtentY        =   582
               Calculator      =   "FrmChallan.frx":19DB
               Caption         =   "FrmChallan.frx":19FB
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "FrmChallan.frx":1A67
               Keys            =   "FrmChallan.frx":1A85
               Spin            =   "FrmChallan.frx":1ACF
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
               ForeColor       =   255
               Format          =   "######0.00"
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
               ValueVT         =   1966145541
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
               Height          =   330
               Index           =   1
               Left            =   120
               TabIndex        =   43
               Top             =   1080
               Width           =   3015
               _Version        =   65536
               _ExtentX        =   5318
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
               Caption         =   " Consignor's Registration No."
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "FrmChallan.frx":1AF7
               Picture         =   "FrmChallan.frx":1B13
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
               Height          =   330
               Index           =   2
               Left            =   120
               TabIndex        =   44
               Top             =   1515
               Width           =   3015
               _Version        =   65536
               _ExtentX        =   5318
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
               Caption         =   " Consignee Name"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "FrmChallan.frx":1B2F
               Picture         =   "FrmChallan.frx":1B4B
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
               Height          =   330
               Index           =   3
               Left            =   120
               TabIndex        =   45
               Top             =   1830
               Width           =   3015
               _Version        =   65536
               _ExtentX        =   5318
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
               Caption         =   " Consignor's Address"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "FrmChallan.frx":1B67
               Picture         =   "FrmChallan.frx":1B83
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
               Height          =   330
               Index           =   4
               Left            =   120
               TabIndex        =   46
               Top             =   2145
               Width           =   3015
               _Version        =   65536
               _ExtentX        =   5318
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
               Caption         =   " Consignee GSTIN Registration No."
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "FrmChallan.frx":1B9F
               Picture         =   "FrmChallan.frx":1BBB
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
               Height          =   330
               Index           =   5
               Left            =   120
               TabIndex        =   47
               Top             =   2580
               Width           =   3015
               _Version        =   65536
               _ExtentX        =   5318
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
               Caption         =   " Carrier Name"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "FrmChallan.frx":1BD7
               Picture         =   "FrmChallan.frx":1BF3
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
               Height          =   330
               Index           =   6
               Left            =   120
               TabIndex        =   48
               Top             =   2895
               Width           =   3015
               _Version        =   65536
               _ExtentX        =   5318
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
               Caption         =   " Carrier Address"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "FrmChallan.frx":1C0F
               Picture         =   "FrmChallan.frx":1C2B
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
               Height          =   330
               Index           =   7
               Left            =   120
               TabIndex        =   49
               Top             =   3210
               Width           =   3015
               _Version        =   65536
               _ExtentX        =   5318
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
               Caption         =   " Vehicle No."
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "FrmChallan.frx":1C47
               Picture         =   "FrmChallan.frx":1C63
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
               Height          =   330
               Index           =   8
               Left            =   5895
               TabIndex        =   50
               Top             =   3210
               Width           =   2415
               _Version        =   65536
               _ExtentX        =   4260
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
               Caption         =   " Destination of goods."
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "FrmChallan.frx":1C7F
               Picture         =   "FrmChallan.frx":1C9B
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
               Height          =   330
               Index           =   9
               Left            =   120
               TabIndex        =   51
               Top             =   3525
               Width           =   3015
               _Version        =   65536
               _ExtentX        =   5318
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
               Caption         =   " Destination Address."
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "FrmChallan.frx":1CB7
               Picture         =   "FrmChallan.frx":1CD3
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
               Height          =   330
               Index           =   10
               Left            =   120
               TabIndex        =   52
               Top             =   3840
               Width           =   3015
               _Version        =   65536
               _ExtentX        =   5318
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
               Caption         =   " E-Way Bill No."
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "FrmChallan.frx":1CEF
               Picture         =   "FrmChallan.frx":1D0B
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
               Height          =   330
               Index           =   11
               Left            =   5895
               TabIndex        =   53
               Top             =   3840
               Width           =   2415
               _Version        =   65536
               _ExtentX        =   4260
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
               Caption         =   " E-Way Bill Date."
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "FrmChallan.frx":1D27
               Picture         =   "FrmChallan.frx":1D43
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
               Height          =   330
               Index           =   12
               Left            =   8160
               TabIndex        =   54
               Top             =   45
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
               Caption         =   " P.O Number"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "FrmChallan.frx":1D5F
               Picture         =   "FrmChallan.frx":1D7B
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
               Height          =   330
               Index           =   13
               Left            =   120
               TabIndex        =   55
               Top             =   4155
               Width           =   3015
               _Version        =   65536
               _ExtentX        =   5318
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
               Caption         =   " Transit Form Number"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "FrmChallan.frx":1D97
               Picture         =   "FrmChallan.frx":1DB3
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
               Height          =   330
               Index           =   14
               Left            =   5895
               TabIndex        =   56
               Top             =   4155
               Width           =   2415
               _Version        =   65536
               _ExtentX        =   4260
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
               Caption         =   " GSTIN Invoice Reference No."
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "FrmChallan.frx":1DCF
               Picture         =   "FrmChallan.frx":1DEB
            End
            Begin TDBDate6Ctl.TDBDate MhDateInput2 
               Height          =   330
               Left            =   8295
               TabIndex        =   15
               Top             =   3840
               Width           =   2340
               _Version        =   65536
               _ExtentX        =   4119
               _ExtentY        =   582
               Calendar        =   "FrmChallan.frx":1E07
               Caption         =   "FrmChallan.frx":1F1F
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "FrmChallan.frx":1F8B
               Keys            =   "FrmChallan.frx":1FA9
               Spin            =   "FrmChallan.frx":2007
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
            Begin VB.Line Line1 
               Index           =   3
               X1              =   0
               X2              =   10800
               Y1              =   1460
               Y2              =   1460
            End
            Begin VB.Line Line1 
               Index           =   2
               X1              =   0
               X2              =   10800
               Y1              =   4725
               Y2              =   4725
            End
            Begin VB.Line Line1 
               Index           =   1
               X1              =   0
               X2              =   10800
               Y1              =   2500
               Y2              =   2500
            End
            Begin VB.Line Line4 
               X1              =   0
               X2              =   10920
               Y1              =   6280
               Y2              =   6280
            End
            Begin VB.Line Line5 
               X1              =   0
               X2              =   13240
               Y1              =   0
               Y2              =   0
            End
            Begin VB.Line Line1 
               Index           =   0
               X1              =   0
               X2              =   10800
               Y1              =   405
               Y2              =   405
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   330
            Index           =   1
            Left            =   -74880
            TabIndex        =   57
            Top             =   7950
            Width           =   495
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
            Index           =   0
            Left            =   -74880
            TabIndex        =   40
            Top             =   8430
            Width           =   495
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   58
      Top             =   0
      Width           =   11235
      _ExtentX        =   19817
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
End
Attribute VB_Name = "FrmChallan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CxnPaperPurchaseOrder As New ADODB.Connection

Dim PaperPurchaseList As New ADODB.Recordset
Dim rstCompanyMasterList As New ADODB.Recordset
Dim rstDeliveryChallanParent As New ADODB.Recordset
Dim rstDeliveryChallanChild As New ADODB.Recordset


Dim rstPaperPOList As New ADODB.Recordset
Dim rstPaperPOParent As New ADODB.Recordset
Dim rstPaperPOChild As New ADODB.Recordset

Dim SupplierCode As String, AccountCode As String, PaperCode As String, BookCode As String, SizeCode As String, RefCode As String

Dim SortOrder, PrevStr
Dim dblBookMark As Double
Dim blnRecordExist As Boolean
Dim EditMode As Boolean
Dim oOutlook As New Outlook.Application
Dim EMailID As String
Dim Attachment As String
Dim Message As String

Public OrderType
Public PrinterCode As String

Dim CartridgeVat As Double

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
           If Me.ActiveControl.Name <> "fpSpread1" And Me.ActiveControl.Name <> "fpSpread2" And Me.ActiveControl.Name <> "fpSpread3" Then SendKeys "{TAB}"
        End If
        If Me.ActiveControl.Name <> "fpSpread1" And Me.ActiveControl.Name <> "fpSpread2" And Me.ActiveControl.Name <> "fpSpread3" Then KeyCode = 0

    End If
End Sub

Private Sub Form_Load()
    'On Error GoTo ErrorHandler

    CenterForm Me
    WheelHook DataGrid1
    BusySystemIndicator True

    Dim Cn As Integer

    Me.Caption = "Delivery Challan"
    CxnPaperPurchaseOrder.CursorLocation = adUseClient
    CxnPaperPurchaseOrder.Open CxnDatabase.ConnectionString


    If rstCompanyMasterList.State = adStateClosed Then rstCompanyMasterList.Open "Select Code,PrintName, Address1, Address2, Address3, Address4, Phone, Fax, EMail, Website,GSTIN From CompanyMaster", CxnDatabase, adOpenKeyset, adLockReadOnly


    PaperPurchaseList.Open "SELECT T.Code,T.Name + '-' + M.Name As Name,M.Address1, M.Address2, M.Address3, M.Address4,M.TIN,M.Phone, M.Mobile, M.EMail FROM PaperPOParent T INNER JOIN AccountMaster M ON T.Supplier=M.Code WHERE OrderType='1' ORDER BY T.Name", CxnPaperPurchaseOrder, adOpenKeyset, adLockOptimistic

    rstPaperPOList.Open "SELECT T.Code,T.Name,Date,M.Name As ConsigneeName,BillAmount FROM DeliveryChallanParent T INNER JOIN AccountMaster M ON T.Consignee=M.Code WHERE OrderType='1' ORDER BY T.Name", CxnPaperPurchaseOrder, adOpenKeyset, adLockOptimistic



    rstDeliveryChallanParent.CursorLocation = adUseClient

    PaperPurchaseList.Filter = adFilterNone

'    rstPaperPOParent.CursorLocation = adUseClient
'    rstPaperPOList.Filter = adFilterNone


    If rstPaperPOList.RecordCount > 0 Then rstPaperPOList.MoveLast
    Set DataGrid1.DataSource = rstPaperPOList
    BusySystemIndicator False
    SSTab1.Tab = 0
    SortOrder = "Name"
    If Not (rstPaperPOList.EOF Or rstPaperPOList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If

    rstPaperPOList.ActiveConnection = Nothing

    SetButtonsForNoRecord
    Exit Sub
'ErrorHandler:
'    BusySystemIndicator False
'    Unload Me
End Sub

Private Sub Form_Activate()
    EnableChildMenu True, True
    Text1.SetFocus
End Sub
Private Sub Form_Deactivate()
    DisableChildMenu
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Toolbar1.Buttons.Item(4).Enabled Then Call Form_KeyDown(vbKeyEscape, 0): Cancel = 1
End Sub
Private Sub Form_Unload(Cancel As Integer)

    WheelUnHook
    Call CloseRecordset(rstPaperPOList)
    Call CloseRecordset(rstPaperPOParent)
    Call CloseRecordset(rstPaperPOChild)
    Call CloseConnection(CxnPaperPurchaseOrder)
    ShowProgressInStatusBar False
    DisableChildMenu

End Sub
Private Sub MhDateInput51_Validate(Cancel As Boolean)
'If Not IsDate(GetDate(MhDateInput51.Text)) Then Cancel = True
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
        txtPONumber.SetFocus
    End If
End Sub

Public Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Dim HiLiteRecord As Boolean

    Dim UpdateFlag As Integer

    Dim CellVal01 As Variant, CellVal02 As Variant, CellVal03 As Variant, CellVal04 As Variant, i As Integer

    If Button.Index = 1 Then

        If rstPaperPOParent.State = adStateOpen Then rstPaperPOParent.Close
        rstPaperPOParent.Open "SELECT * FROM PaperPOParent WHERE Code=''", CxnPaperPurchaseOrder, adOpenKeyset, adLockOptimistic
        ClearFields

        Call LoadPaperList("")


        If AddRecord(rstPaperPOParent) Then
            Text2.Text = GenerateCode(CxnPaperPurchaseOrder, "SELECT MAX(VAL(Name)) FROM PaperPOParent WHERE OrderType='" & OrderType & "'", 10, Space(1))
            MhDateInput1.Text = Format(Date, "dd-MM-yyyy")
            Call SetButtons(False)
            SSTab1.Tab = 1
            Text3.SetFocus
            blnRecordExist = False
            CxnPaperPurchaseOrder.BeginTrans
        End If
    ElseIf Button.Index = 2 Then
        If rstPaperPOList.RecordCount = 0 Then Exit Sub
        SSTab1.Tab = 1
        EditRecord
    ElseIf Button.Index = 3 Then
        If rstPaperPOList.RecordCount = 0 Then Exit Sub
        If AllowTransactionsDeletion = 0 Then Call DisplayError("You don't have the rights to Delete this Voucher"): Exit Sub
        SSTab1.Tab = 1
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            On Error Resume Next
            MdiMainMenu.MousePointer = vbHourglass
            CxnPaperPurchaseOrder.Execute "DELETE FROM PaperPOParent WHERE Code='" & rstPaperPOList.Fields("Code").Value & "'"
'            CxnPaperPurchaseOrder.Execute "DELETE FROM PaperPOChild WHERE Code='" & rstPaperPOList.Fields("Code").Value & "'"
'            CxnPaperPurchaseOrder.Execute "DELETE FROM PaperIOChild WHERE Code='" & rstPaperPOList.Fields("Code").Value & "'"
            CxnPaperPurchaseOrder.Execute "DELETE FROM PaperPOChildRef WHERE Code='" & rstPaperPOList.Fields("Code").Value & "'"

            MdiMainMenu.MousePointer = vbNormal
            If Err.Number = 0 Then
                rstPaperPOList.Delete
                rstPaperPOList.MoveNext
                If rstPaperPOList.RecordCount > 0 And rstPaperPOList.EOF Then rstPaperPOList.MoveLast
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
        SaveFields
        UpdateFlag = 0
        If UpdateRecord(rstPaperPOParent) Then
            If UpdatePaperList("D") Then
                UpdateFlag = 1
                With fpSpread1
                    For i = 1 To .DataRowCnt
                        .SetActiveCell 6, i
                        .GetText 6, i, CellVal01
                        .GetText 7, i, CellVal02
                        If Val(CellVal01) <> 0 And CellVal02 <> "" Then
                            If Not UpdatePaperList("I1") Then UpdateFlag = 0: Exit For
                        End If
                    Next
                End With

                If UpdateFlag = 1 Then
                    With fpSpread2
                        For i = 1 To .DataRowCnt
                            .SetActiveCell 3, i
                            .GetText 3, i, CellVal01
                            .GetText 5, i, CellVal02
                            .GetText 6, i, CellVal03
                            If Val(CellVal01) <> 0 And CellVal02 <> "" And CellVal03 <> "" Then
                                If Not UpdatePaperList("I2") Then UpdateFlag = 0: Exit For
                            End If
                        Next
                    End With
                End If
                If UpdateFlag = 1 Then
                    With fpSpread3
                        For i = 1 To .DataRowCnt
                            .SetActiveCell 3, i
                            .GetText 3, i, CellVal01
                            .GetText 4, i, CellVal02
                            .GetText 5, i, CellVal03
                            .GetText 6, i, CellVal04
                            If Val(CellVal01) <> 0 And CellVal02 <> "" And CellVal03 <> "" And CellVal04 <> "" Then
                                If Not UpdatePaperList("I3") Then UpdateFlag = 0: Exit For
                            End If
                        Next
                    End With
              End If

            End If
        End If
        If UpdateFlag Then
            AddToList
            CxnPaperPurchaseOrder.CommitTrans
            If rstPaperPOParent.State = adStateOpen Then rstPaperPOParent.Close
            rstPaperPOParent.CursorLocation = adUseClient
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
        If CancelRecordUpdate(rstPaperPOParent) Then
            CxnPaperPurchaseOrder.RollbackTrans
            If rstPaperPOParent.State = adStateOpen Then rstPaperPOParent.Close
            rstPaperPOParent.CursorLocation = adUseClient
            Call SetButtons(True)
            SetButtonsForNoRecord
            SSTab1.Tab = 0
            LockFields (False)
        End If
    ElseIf Button.Index = 6 Then
        SSTab1.Tab = 0
        Set DataGrid1.DataSource = Nothing
        rstPaperPOList.ActiveConnection = CxnPaperPurchaseOrder
        Do While Not RefreshRecord(rstPaperPOList)
        Loop
        Set DataGrid1.DataSource = rstPaperPOList
        rstPaperPOList.ActiveConnection = Nothing
        If rstPaperPOList.RecordCount > 0 Then rstPaperPOList.MoveLast
        rstSupplierList.ActiveConnection = CxnPaperPurchaseOrder

        Do While Not RefreshRecord(rstSupplierList)
        Loop

        rstSupplierList.ActiveConnection = Nothing
        rstPaperList.ActiveConnection = CxnPaperPurchaseOrder

        Do While Not RefreshRecord(rstPaperList)
        Loop

        rstPaperList.ActiveConnection = Nothing
        rstAccountList.ActiveConnection = CxnPaperPurchaseOrder

        Do While Not RefreshRecord(rstAccountList)
        Loop

        rstAccountList.ActiveConnection = Nothing
        rstBookList.ActiveConnection = CxnPaperPurchaseOrder
        Do While Not RefreshRecord(rstBookList)
        Loop

        rstBookList.ActiveConnection = Nothing
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
        If rstPaperPOList.RecordCount = 0 Then Exit Sub
        Call DisplayMenu("P")
        HiLiteRecord = True
    ElseIf Button.Index = 10 Then
        If rstPaperPOList.RecordCount = 0 Then Exit Sub
        Call DisplayMenu("S")
        HiLiteRecord = True
    ElseIf Button.Index = 11 Then
        If rstPaperPOList.RecordCount = 0 Then Exit Sub
        Call DisplayMenu("M")
        HiLiteRecord = True
    ElseIf Button.Index = 13 Then
        If rstPaperPOList.RecordCount > 0 Then rstPaperPOList.MoveFirst
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 14 Then
        If rstPaperPOList.RecordCount > 0 Then
            rstPaperPOList.MovePrevious
            If rstPaperPOList.BOF Then rstPaperPOList.MoveNext
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 15 Then
        If rstPaperPOList.RecordCount > 0 Then
            rstPaperPOList.MoveNext
            If rstPaperPOList.EOF Then rstPaperPOList.MovePrevious
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 16 Then
        If rstPaperPOList.RecordCount > 0 Then rstPaperPOList.MoveLast
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 18 Then
        Unload Me
        HiLiteRecord = False
    End If

    If HiLiteRecord Then
        If Not (rstPaperPOList.EOF Or rstPaperPOList.BOF) Then
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
    If ColIndex = 0 Or ColIndex = 2 Then
        SortOrder = DataGrid1.Columns(ColIndex).DataField
        rstPaperPOList.Sort = "[" + SortOrder & "] Asc"
    End If
    DataGrid1.ClearSelCols
    If Not (rstPaperPOList.EOF Or rstPaperPOList.BOF) Then
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
    If rstPaperPOList.RecordCount = 0 Then
        'Toolbar1.Buttons.Item(2).Enabled = False
'        Toolbar1.Buttons.Item(3).Enabled = False
'        Toolbar1.Buttons.Item(9).Enabled = False
'        Toolbar1.Buttons.Item(10).Enabled = False
'        Toolbar1.Buttons.Item(11).Enabled = False
'        Toolbar1.Buttons.Item(13).Enabled = False
'        Toolbar1.Buttons.Item(14).Enabled = False
'        Toolbar1.Buttons.Item(15).Enabled = False
'        Toolbar1.Buttons.Item(16).Enabled = False
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
        'MhDateInput3.Text = Format(DateAdd("d", 2, CDate(GetDate(MhDateInput1.Text))), "dd-MM-yyyy")
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
    MhRealInput12.Value = Val(MhRealInput19.Text) * Val(MhRealInput11.Text) / 100  'VAT
    Call CalculateTotal("N")    'VAT Changed

End Sub
Private Sub MhRealInput13_Validate(Cancel As Boolean)   'Cartage

    CartridgeVat = 0
    CartridgeVat = Val(MhRealInput13.Value) * Val(MhRealInput11.Text) / 100 'VAT
    MhRealInput12.Value = MhRealInput12.Value + CartridgeVat

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
    CartridgeVat = 0

    txtChallan.Text = ""
    txtConsignerName.Text = ""
    txtConsignerAddress.Text = ""
    txtConsignerRegNo.Text = ""

    txtConsigneeName.Text = ""
    txtConsigneeAddress.Text = ""
    txtConsigneeRegNo.Text = ""

    txtCarrierName.Text = ""
    txtCarrierAddress.Text = ""
    txtVehicleNo.Text = ""
    txtDestinationGoods.Text = ""
    txtDestinationAddress.Text = ""
    txtEWayBillNo.Text = ""
    txtTransitFormNumber.Text = ""
    txtRefNo.Text = ""
    txtPONumber.Text = ""


    MhDateInput1.Text = Format(Date, "dd-MM-yyyy")
    MhDateInput2.Text = "  -  -    "    'Bill Date

    MhRealInput1.Value = 0 'Total Quantity (Ream)
    MhRealInput2.Value = 0 'Total Quantity (Kg)
    MhRealInput3.Value = 0 'Total Gross Amount
    MhRealInput4.Value = 0  'Reams/bundle
    MhRealInput5.Value = 0  'Total bundles
    MhRealInput6.Value = 0   'Cartage/Kg
    MhRealInput7.Value = 0 'GST (%)
    MhRealInput8.Value = 0 'GST


    fpSpread1.ClearRange 1, 1, fpSpread1.MaxCols, fpSpread1.MaxRows, True: fpSpread1.SetActiveCell 1, 1

End Sub
Private Sub LoadFields()
    If rstPaperPOParent.EOF Or rstPaperPOParent.BOF Then Exit Sub
    Text2.Text = rstPaperPOParent.Fields("Name").Value
    MhDateInput1.Text = Format(rstPaperPOParent.Fields("Date").Value, "dd-MM-yyyy")

    MhDateInput3.Text = Format(rstPaperPOParent.Fields("DeliveryDate").Value, "dd-MM-yyyy")
    If Not IsNull(rstPaperPOParent.Fields("ExtendDate").Value) Then MhDateInput51.Text = Format(rstPaperPOParent.Fields("ExtendDate").Value, "dd-MM-yyyy")
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
    If Not IsDate(MhDateInput51.Text) Then rstPaperPOParent.Fields("ExtendDate").Value = Null Else rstPaperPOParent.Fields("ExtendDate").Value = GetDate(MhDateInput51.Text)
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
    ''****************For Book Ref********************
    If rstBookRef.State = adStateOpen Then rstBookRef.Close

    rstBookRef.Open "SELECT Paper As PaperCode,M1.Name As PaperName,Book As AccountCode,M2.Name As AccountName,Form,Quantity,Wastage,Consumption,TotalConsumption,PrinterCode,SizeCode,T.Pages,Color,BillingQty1,BillingQty2,SendQuantity,BalanceQuantity,[1F] FROM (PaperPOChildRef T INNER JOIN PaperMaster M1 ON T.Paper=M1.Code) INNER JOIN BookMaster M2 ON T.Book=M2.Code WHERE T.Code='" & strOrderCode & "' ORDER BY M1.Name,M2.Name", CxnPaperPurchaseOrder, adOpenKeyset, adLockOptimistic




    rstBookRef.ActiveConnection = Nothing
    If rstBookRef.RecordCount > 0 Then rstBookRef.MoveFirst
    i = 0
    Do While Not rstBookRef.EOF
        i = i + 1
        With fpSpread3

            .SetText 1, i, rstBookRef.Fields("PaperName").Value
            .SetText 2, i, rstBookRef.Fields("AccountName").Value
            .SetText 3, i, Val(rstBookRef.Fields("Form").Value)
            .SetText 4, i, Val(rstBookRef.Fields("Quantity").Value)
            .SetText 5, i, rstBookRef.Fields("AccountCode").Value
            .SetText 6, i, rstBookRef.Fields("PaperCode").Value

            If Not IsNull(rstBookRef.Fields("Wastage").Value) Then
                    .SetText 7, i, Val(rstBookRef.Fields("Wastage").Value)
            End If

            If Not IsNull(rstBookRef.Fields("Consumption").Value) Then
                    .SetText 8, i, Val(rstBookRef.Fields("Consumption").Value)
            End If

            If Not IsNull(rstBookRef.Fields("Consumption").Value) Then
                    .SetText 9, i, Val(rstBookRef.Fields("TotalConsumption").Value)
            End If
             .SetText 10, i, rstBookRef.Fields("PrinterCode").Value
             .SetText 11, i, rstBookRef.Fields("SizeCode").Value

            If Not IsNull(rstBookRef.Fields("Pages").Value) Then
                .SetText 12, i, Val(rstBookRef.Fields("Pages").Value)
            End If
             .SetText 13, i, rstBookRef.Fields("Color").Value

            If Not IsNull(rstBookRef.Fields("BillingQty1").Value) Then
                .SetText 15, i, Val(rstBookRef.Fields("BillingQty1").Value)
            End If

            If Not IsNull(rstBookRef.Fields("BillingQty2").Value) Then
                 .SetText 16, i, Val(rstBookRef.Fields("BillingQty2").Value)
            End If
            If Not IsNull(rstBookRef.Fields("SendQuantity").Value) Then
               .SetText 17, i, Val(rstBookRef.Fields("SendQuantity").Value)
            Else
               .SetText 17, i, 0
            End If
            If Not IsNull(rstBookRef.Fields("BalanceQuantity").Value) Then
                 .SetText 18, i, Val(rstBookRef.Fields("BalanceQuantity").Value)
            Else
               .SetText 18, i, 0
            End If
            If Not IsNull(rstBookRef.Fields("1F").Value) Then
             .SetText 19, i, Val(rstBookRef.Fields("1F").Value)
            End If
        End With
        rstBookRef.MoveNext
    Loop
    Set rstBookRef = Nothing
    '*****************End ****************************
    Exit Sub
ErrorHandler:
    DisplayError ("Failed to Load Paper List")
End Sub
Private Sub CalculateCartage()

    If Val(MhRealInput10.Text) <> 0 Then
        MhRealInput13.Value = Round(Val(MhRealInput18.Text) * Val(MhRealInput10.Text), 0) 'Total Cartage
        If Not blnRecordExist Then MhRealInput22.Value = MhRealInput13.Value
        CalculateTotal ("N")

    End If
End Sub
Private Sub CalculateTotal(ByVal strType As String)
    Dim Qty01 As Variant, Qty02 As Variant, Amt As Variant, TCAmt As Variant
    Dim i As Integer
    Dim Qty As Long
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

        Qty = 0
        TCAmt = 0

        With fpSpread3
            For i = 1 To .DataRowCnt
                .GetText 4, i, Qty01
                Qty = Qty + Int(Val(Qty01))
                MhRealInput221.Value = Qty
                .GetText 9, i, Qty02
                TCAmt = TCAmt + (Val(Qty02))
                MhRealInput2211.Value = TCAmt
             Next
        End With

        MhRealInput22111.Value = Abs(MhRealInput17.Value - MhRealInput2211.Value)
        MhRealInput8_Validate False 'Calculate Total bundles
        MhRealInput12.Value = Val(MhRealInput19.Text) * Val(MhRealInput11.Text) / 100 'VAT

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
    Dim CellVal(1 To 17) As Variant, Sheets As Long
    On Error GoTo ErrorHandler
    UpdatePaperList = True
    If ActionType = "D" And (Not blnRecordExist) Then Exit Function
    If ActionType = "D" Then
        CxnPaperPurchaseOrder.Execute "DELETE FROM PaperPOChild WHERE Code='" & rstPaperPOParent.Fields("Code").Value & "'"
        CxnPaperPurchaseOrder.Execute "DELETE FROM PaperIOChild WHERE Code='" & rstPaperPOParent.Fields("Code").Value & "'"
        CxnPaperPurchaseOrder.Execute "DELETE FROM PaperPOChildRef WHERE Code='" & rstPaperPOParent.Fields("Code").Value & "'"
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
    ElseIf ActionType = "I3" Then
        With fpSpread3
            .GetText 3, .ActiveRow, CellVal(1)  'Form
            .GetText 4, .ActiveRow, CellVal(2)  'Quantity
            .GetText 5, .ActiveRow, CellVal(3)  'Book Code
            .GetText 6, .ActiveRow, CellVal(4)  'Paper

            .GetText 7, .ActiveRow, CellVal(6)  'Wastage
            .GetText 8, .ActiveRow, CellVal(7)  'Consumption
            .GetText 9, .ActiveRow, CellVal(8)  'T Consptn
            .GetText 10, .ActiveRow, CellVal(9)  'PrinterCode
            .GetText 11, .ActiveRow, CellVal(10)  'SizeCode
            .GetText 12, .ActiveRow, CellVal(11)  'Pages
            .GetText 13, .ActiveRow, CellVal(12)  'Color
            .GetText 15, .ActiveRow, CellVal(13)  'BiilingQty1
            .GetText 16, .ActiveRow, CellVal(14)  'BillingQty2
            .GetText 17, .ActiveRow, CellVal(15)  'Paper Send
            .GetText 18, .ActiveRow, CellVal(16)  'Paper Balance
            .GetText 19, .ActiveRow, CellVal(17)  '1F
        End With

        Dim aaa As String
        aaa = "INSERT INTO PaperPOChildRef VALUES ('" & rstPaperPOParent.Fields("Code").Value & "','" & CellVal(4) & "','" & CellVal(3) & "'," & Val(CellVal(1)) & "," & Val(CellVal(2)) & "," & Val(CellVal(6)) & "," & Val(CellVal(7)) & "," & Val(CellVal(8)) & ",'" & CellVal(9) & "','" & CellVal(10) & "'," & Val(CellVal(11)) & ",'" & CellVal(12) & "'," & Val(CellVal(13)) & "," & Val(CellVal(14)) & "," & Val(CellVal(15)) & "," & Val(CellVal(16)) & "," & Val(CellVal(17)) & ")"

        CxnPaperPurchaseOrder.Execute "INSERT INTO PaperPOChildRef VALUES ('" & rstPaperPOParent.Fields("Code").Value & "','" & CellVal(4) & "','" & CellVal(3) & "'," & Val(CellVal(1)) & "," & Val(CellVal(2)) & "," & Val(CellVal(6)) & "," & Val(CellVal(7)) & "," & Val(CellVal(8)) & ",'" & CellVal(9) & "','" & CellVal(10) & "'," & Val(CellVal(11)) & ",'" & CellVal(12) & "'," & Val(CellVal(13)) & "," & Val(CellVal(14)) & "," & Val(CellVal(15)) & "," & Val(CellVal(16)) & "," & Val(CellVal(17)) & ")"
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

    Dim rstCompanyMaster As New ADODB.Recordset, rstPurchaseOrder As New ADODB.Recordset, rstPurchaseOrderChild As New ADODB.Recordset, rstPurchaseOrderRef As New ADODB.Recordset, Prefix As String
    Dim FileName As String
    Dim oOutlookMsg As Outlook.MailItem, RecordAffected As Integer
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    Prefix = IIf(OrderType = "1", "PB", "PT") & "/" & Right(Year(FinancialYearFrom), 2) + "-" + Right(Year(FinancialYearTo), 2) & "/"
    rstCompanyMaster.Open "SELECT PrintName,Address1,Address2,Address3,Address4,Phone,Fax,eMail FROM CompanyMaster", CxnDatabase, adOpenKeyset, adLockReadOnly

    rstPurchaseOrder.Open "SELECT '" & Prefix & "'+TRIM(P.Name) As OrderNo,[Date] As OrderDate,DeliveryDate,TRIM(M1.PrintName) As SupplierName,[VAT%],VAT,P.Cartage,Adjustment,BillAmount,Remarks,TRIM(M2.PrintName) As PaperName,'',QuantityOther,[Weight/Ream],QuantityKg,[Rate/Kg],(SELECT TOP 1 '" & Prefix & "'+TRIM(P1.Name)+'/'+FORMAT(P1.Date,'dd-MM-yyyy')+'/'+FORMAT([Rate/Kg],'0.00') FROM PaperPOParent P1 INNER JOIN PaperPOChild C1 ON P1.Code=C1.Code WHERE C1.Paper=C.Paper AND P1.Code<P.Code ORDER BY P1.Name DESC) As LastPurchaseRate,Amount,BillNo,BillDate,TRIM(eMail) As SupplierMail FROM ((PaperPOParent P LEFT JOIN PaperPOChild C ON P.Code=C.Code) LEFT JOIN AccountMaster M1 ON M1.Code=P.Supplier) LEFT JOIN PaperMaster M2 ON M2.Code=C.Paper WHERE P.Code='" & OrderCode & "' ORDER BY M2.PrintName", CxnDatabase, adOpenKeyset, adLockOptimistic

    rstPurchaseOrderChild.Open "SELECT '" & Prefix & "'+TRIM(P.Name) As OrderNo,[Date] As OrderDate,TRIM(M3.PrintName) As Godown,TRIM(M2.PrintName) As PaperName,TRIM(M1.PrintName) As PrinterName,'' As RefNo,QuantityOther As Quantity,Tat,'' As Remarks,M1.Address1 As PrinterAdd1,M1.Address2 As PrinterAdd2,M1.Address3 As PrinterAdd3,M1.Address4 As PrinterAdd4,TRIM(M1.eMail) As PrinterMail FROM (((PaperPOParent P INNER JOIN PaperIOChild C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON C.Account=M1.Code) INNER JOIN PaperMaster M2 ON C.Paper=M2.Code) INNER JOIN AccountMaster M3 ON P.Supplier=M3.Code WHERE P.Code='" & OrderCode & "' ORDER BY M2.PrintName", CxnDatabase, adOpenKeyset, adLockOptimistic
    rstPurchaseOrderRef.Open "SELECT Paper As PaperCode,M1.Name As PaperName,Book As BookCode,M2.Name As BookName,Form,Quantity,Wastage,Consumption,TotalConsumption,SendQuantity,BalanceQuantity FROM (PaperPOChildRef T INNER JOIN PaperMaster M1 ON T.Paper=M1.Code) INNER JOIN BookMaster M2 ON T.Book=M2.Code WHERE T.Code='" & OrderCode & "' ORDER BY M1.Name,M2.Name", CxnDatabase, adOpenKeyset, adLockOptimistic
    Screen.MousePointer = vbNormal
    rstPurchaseOrder.ActiveConnection = Nothing: rstPurchaseOrderChild.ActiveConnection = Nothing: rstPurchaseOrderRef.ActiveConnection = Nothing
    If VchType = 1 Then
        rptPaperPurchaseOrder.Text1.SetText IIf(OrderType = "1", "Book", "Title") & " Paper Purchase Order"
        rptPaperPurchaseOrder.Text2.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
        rptPaperPurchaseOrder.Text3.SetText Trim(rstCompanyMaster.Fields("Address1").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address2").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address3").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address4").Value)
        rptPaperPurchaseOrder.Text24.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & Space(1) & "Fax : " & Trim(rstCompanyMaster.Fields("Fax").Value) & Space(1) & "e-Mail : " & Trim(rstCompanyMaster.Fields("eMail").Value)

        rptPaperPurchaseOrder.Text20.SetText "Add : GST @" + Format(rstPurchaseOrder.Fields("VAT%").Value, "0.00") + "%"

        rptPaperPurchaseOrder.Text28.SetText " (" & Trim(NumberToWords(rstPurchaseOrder.Fields("BillAmount").Value, True)) & ")"

        rptPaperPurchaseOrder.Text27.SetText "for " & Trim(rstPurchaseOrder.Fields("SupplierName").Value)
        rptPaperPurchaseOrder.Text9.SetText "for " & Trim(rstCompanyMaster.Fields("PrintName").Value)
    '   **************  By Shamshad Alam **********************************************
        rptPaperPurchaseOrder.Text8.SetText Trim(COMPANY_CIN) 'Add here company cin no

        If rstPurchaseOrderRef.RecordCount = 0 Then
           rptPaperPurchaseOrder.Section21.Suppress = True
        End If
        rptPaperPurchaseOrder.Database.SetDataSource rstPurchaseOrder, 3, 1
        rptPaperPurchaseOrder.Subreport1.OpenSubreport.Database.SetDataSource rstPurchaseOrderChild, 3, 1
        rptPaperPurchaseOrder.Subreport2.OpenSubreport.Database.SetDataSource rstPurchaseOrderRef, 3, 1
        'sString = Replace(sString, Chr(34), Chr(39))'Replace double qote with single qoate
         EMailID = ""
         EMailID = Replace(rstPurchaseOrder.Fields("SupplierMail").Value, Chr(39), "") 'Replace single qote with space
         'EMailID = "ms.alam@rachnasagar.in"
        Attachment = Trim(rstPurchaseOrder.Fields("OrderNo").Value)
        Attachment = Mid(Attachment, InStr(4, Attachment, "/") + 1)
        Message = "Dear Sir,<Br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Please find attached herewith PO #" & Trim(rstPurchaseOrder.Fields("OrderNo").Value) & " for doing the needful at your end. An early execution of the order will be highly appreciated.<Br>Kindly acknowledge the receipt of mail and confirm the date of execution of order.<Br><Br>" & IIf(Note = "", "", "<b><u>Note : " & Note & "</b></u><Br><Br>") & "Thanks & Regards<Br>Production Department<Br>" & Trim(rstCompanyMaster.Fields("PrintName").Value) & "<Br>Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & "<Br>E-Mail : <a HRef='mailto:" & Trim(rstCompanyMaster.Fields("EMail").Value) & "'>" & Trim(rstCompanyMaster.Fields("EMail").Value) & "</a>"

        If OutputType = "S" Then
            FrmReportViewer.EMailID = Trim(EMailID)
            FrmReportViewer.Subject = IIf(OrderType = "1", "Book", "Title") & " Paper Purchase Order #" & Trim(rstPurchaseOrder.Fields("OrderNo").Value)
            FrmReportViewer.Attachment = Attachment
            FrmReportViewer.Message = Message
            Set FrmReportViewer.Report = rptPaperPurchaseOrder
            FrmReportViewer.Show vbModal

         ElseIf OutputType = "P" Then
            rptPaperPurchaseOrder.PrintOut False    'Print Report Without Prompt
         Else
            rptPaperPurchaseOrder.ExportOptions.FormatType = crEFTPortableDocFormat    ' Set the Export Format As .Pdf
            rptPaperPurchaseOrder.ExportOptions.DestinationType = crEDTDiskFile
            rptPaperPurchaseOrder.ExportOptions.DiskFileName = App.Path & "\Report\" & Attachment & ".Pdf"
            rptPaperPurchaseOrder.Export False
            rstPurchaseOrder.MoveFirst

            Set oOutlookMsg = oOutlook.CreateItem(olMailItem)
            With oOutlookMsg
                .To = Trim(EMailID)
                .Subject = IIf(OrderType = "1", "Book", "Title") & " Paper Purchase Order #" & Trim(rstPurchaseOrder.Fields("OrderNo").Value)
                .HTMLBody = "<Font Face='Calibri' Size='3'>" & Message & "</a>" & "</Font>"
                .Attachments.Add (App.Path & "\Report\" & Attachment & ".Pdf")
                .Importance = olImportanceHigh
                .ReadReceiptRequested = True
                If CheckEmpty(.To, False) Then .Display Else .Send
            End With
            Set oOutlookMsg = Nothing
         End If
         Set rptPaperPurchaseOrder = Nothing
    Else

        Dim PrinterMail As String
        PrinterMail = ""
        Do While Not rstPurchaseOrderChild.EOF
            If Trim(rstPurchaseOrderChild.Fields("PrinterMail").Value) <> "" Then PrinterMail = PrinterMail + IIf(PrinterMail = "", "", ";") & Trim(Replace(rstPurchaseOrderChild.Fields("PrinterMail").Value, Chr(39), ""))
            rstPurchaseOrderChild.MoveNext
        Loop
        rstPurchaseOrderChild.MoveFirst
        rptPaperIssueOrder.Text1.SetText IIf(OrderType = "1", "Book", "Title") & " Paper Issue Voucher"
        rptPaperIssueOrder.Text2.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
        rptPaperIssueOrder.Text3.SetText Trim(rstCompanyMaster.Fields("Address1").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address2").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address3").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address4").Value)
        rptPaperIssueOrder.Text24.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & Space(1) & "Fax : " & Trim(rstCompanyMaster.Fields("Fax").Value) & Space(1) & "e-Mail : " & Trim(rstCompanyMaster.Fields("eMail").Value)
        rptPaperIssueOrder.Text9.SetText "for " & Trim(rstCompanyMaster.Fields("PrintName").Value)
        rptPaperIssueOrder.Database.SetDataSource rstPurchaseOrderChild, 3, 1

        'PrinterMail = "ms.alam@rachnasagar.in"

        Attachment = Trim(rstPurchaseOrderChild.Fields("OrderNo").Value)
        Attachment = Mid(Attachment, InStr(4, Attachment, "/") + 1)
        Message = "Dear Sir,<Br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Please find attached herewith PO #" & Trim(rstPurchaseOrderChild.Fields("OrderNo").Value) & " for doing the needful at your end. An early execution of the order will be highly appreciated.<Br>Kindly acknowledge the receipt of mail and confirm the date of execution of order.<Br><Br>" & IIf(Note = "", "", "<b><u>Note : " & Note & "</b></u><Br><Br>") & "Thanks & Regards<Br>Production Department<Br>" & Trim(rstCompanyMaster.Fields("PrintName").Value) & "<Br>Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & "<Br>E-Mail : <a HRef='mailto:" & Trim(rstCompanyMaster.Fields("EMail").Value) & "'>" & Trim(rstCompanyMaster.Fields("EMail").Value) & "</a>"
        If OutputType = "S" Then
            FrmReportViewer.EMailID = Trim(PrinterMail)
            FrmReportViewer.Subject = IIf(OrderType = "1", "Book", "Title") & " Paper Issue Voucher #" & Trim(rstPurchaseOrderChild.Fields("OrderNo").Value)
            FrmReportViewer.Attachment = Attachment
            FrmReportViewer.Message = Message
            Set FrmReportViewer.Report = rptPaperIssueOrder
            FrmReportViewer.Show vbModal
        ElseIf OutputType = "P" Then
            rptPaperIssueOrder.PrintOut False    'Print Report Without Prompt
        Else
            rptPaperIssueOrder.ExportOptions.FormatType = crEFTPortableDocFormat    ' Set the Export Format As .Pdf
            rptPaperIssueOrder.ExportOptions.DestinationType = crEDTDiskFile
            rptPaperIssueOrder.ExportOptions.DiskFileName = App.Path & "\Report\" & Attachment & ".Pdf"
            rptPaperIssueOrder.Export False
            rstPurchaseOrderChild.MoveFirst

            Set oOutlookMsg = oOutlook.CreateItem(olMailItem)
            With oOutlookMsg
                .To = Trim(PrinterMail)
                .Subject = IIf(OrderType = "1", "Book", "Title") & " Paper Issue Voucher #" & Trim(rstPurchaseOrderChild.Fields("OrderNo").Value)
                 .HTMLBody = "<Font Face='Calibri' Size='3'>" & Message & "</a>" & "</Font>"
                .Attachments.Add (App.Path & "\Report\" & Attachment & ".Pdf")
                .Importance = olImportanceHigh
                .ReadReceiptRequested = True
                If CheckEmpty(.To, False) Then .Display Else .Send
            End With
            Set oOutlookMsg = Nothing
        End If
        Set rptPaperIssueOrder = Nothing
    End If
    Call CloseRecordset(rstPurchaseOrder): Call CloseRecordset(rstCompanyMaster): Call CloseRecordset(rstPurchaseOrderChild): Call CloseRecordset(rstPurchaseOrderRef)
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
                    If Not blnRecordExist Then MhRealInput8.Value = Val(rstPaperList.Fields("Reams/Bundle").Value)
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
'        MhRealInput2211.Value = Val(MhRealInput17.Text) - Val(MhRealInput2211.Text)
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
Private Sub fpSpread3_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    EditMode = IIf(Mode = 1, True, False)

'    If Mode = 1 Then
'       EditMode = True
'    Else
'        EditMode = False
'    End If

End Sub
Private Function ChkPaper() As Boolean

    Dim i As Integer, K As Integer, Paper01 As Variant, Qty01 As Variant, Paper02 As Variant, Qty02 As Variant, Qty As Long, Price As Variant
    ChkPaper = True
    For i = 1 To fpSpread1.DataRowCnt
        fpSpread1.GetText 1, i, Paper01
        fpSpread1.GetText 2, i, Qty01
        fpSpread1.GetText 5, i, Price
        If Val(Price) = 0 Then DisplayError ("Price of Paper at row #" & Trim(str(i)) & " is zero"): ChkPaper = False: Exit Function
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
Private Sub Export2Excel()
On Error GoTo er
Dim oExcel As Object
Set oExcel = CreateObject("Excel.Application")
Dim oWorkBook As Object
Dim oWorkSheet As Object
Dim i As Integer, K As Integer, M As Integer, j As Integer
Dim lRow As Long
Dim LastRow As Long
Dim LastCol As Long
oExcel.Visible = False

If rstPaperPOList.RecordCount = 0 Then On Error GoTo 0: Screen.MousePointer = vbNormal: Exit Sub
If Not FileExist(App.Path & "\Report\Book1.xlsx") Then DisplayError ("Template File Missing"): Exit Sub
Set oExcel = CreateObject("Excel.Application")
oExcel.Workbooks.Open App.Path & "\Template\Blank.xlsx"
Set oWorkSheet = oExcel.Workbooks("Blank.xlsx").Sheets("Sheet1")
 oWorkSheet.Cells(1, 1).Value = "Code"
 oWorkSheet.Cells(1, 2).Value = "Order No"
 oWorkSheet.Cells(1, 3).Value = "Order Date"
 oWorkSheet.Cells(1, 4).Value = "Supplier Name"
 oWorkSheet.Cells(1, 5).Value = "Order Amount"
i = 2
rstPaperPOList.MoveFirst
Do While Not rstPaperPOList.EOF
    oWorkSheet.Cells(i, "A").Value = rstPaperPOList.Fields("Code").Value
    oWorkSheet.Cells(i, "B").Value = rstPaperPOList.Fields("Name").Value
    oWorkSheet.Cells(i, "C").Value = Format(rstPaperPOList.Fields("Date").Value, "dd/MM/yyyy")
    oWorkSheet.Cells(i, "D").Value = rstPaperPOList.Fields("SupplierName").Value
    oWorkSheet.Cells(i, "E").Value = rstPaperPOList.Fields("BillAmount").Value
 i = i + 1
    rstPaperPOList.MoveNext
Loop

oExcel.Range("A:A").EntireColumn.Hidden = True

Screen.MousePointer = vbNormal
oExcel.Sheets("Sheet1").Activate
oExcel.Columns("A:L").EntireColumn.AutoFit
oExcel.Workbooks.Item(1).Save
oExcel.Visible = True
Set oExcel = Nothing
er:
If Err.Number = 1004 Then
Exit Sub
End If
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

Private Function CalculateConsumption(ByVal xPrintingType As String, ByVal MhRealInput1 As Variant) As Double

    Dim Forms¼ As Variant, Forms½ As Variant, Forms1 As Variant, WastageRate As Variant, CurrentPaperConsumption As Variant, Cnt As Integer, FS As Variant

    fpSpread4.GetText 3, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), Forms¼
    fpSpread4.GetText 4, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), Forms½
    fpSpread4.GetText 5, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), Forms1
    fpSpread4.GetText 9, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), WastageRate

    CalculateConsumption = CLng(Val(MhRealInput1) * (Val(Forms¼) * 0.25 + Val(Forms½) * 0.5 + Val(Forms1) * 1) * ((100 + Val(WastageRate)) / 100))
    CalculateConsumption = CLng(Val(CalculateConsumption) / 2)
    fpSpread4.GetText 22, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), FS
    CalculateConsumption = IIf(Val(FS) = 0.5, 2, IIf(Val(FS) = 2, 0.5, 1)) * CalculateConsumption
    CalculateConsumption = Format(CLng(Int(Val(CalculateConsumption) / 500)) + ((Val(CalculateConsumption) Mod 500) / 1000), "0.000")
    fpSpread4.SetText 10, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), CalculateConsumption
    If fpSpread4.ActiveRow = IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)) Then
        MhRealInput131.Text = Format(Val(CalculateConsumption), "0.000")
    End If
    For Cnt = 1 To fpSpread4.MaxRows
        fpSpread4.GetText 10, Cnt, CurrentPaperConsumption
        MhRealInput1310.Text = Format(IIf(Cnt = 1, 0, Val(MhRealInput1310.Text)) + CLng((Int(Val(CurrentPaperConsumption)) * 500) + ((Val(CurrentPaperConsumption) - Int(Val(CurrentPaperConsumption))) * 1000)), "0.000")
    Next
    MhRealInput1310.Text = Format(CLng(Int(Val(MhRealInput1310.Text) / 500)) + ((Val(MhRealInput1310.Text) Mod 500) / 1000), "0.000")

End Function

Private Sub CalculateAmount()
    Dim Cnt As Integer, TotalPlates¼ As Variant, TotalPlates½ As Variant, TotalPlates1 As Variant, PlateRate As Variant, TotalForms¼ As Variant, TotalForms½ As Variant, TotalForms1 As Variant, PrintRate As Variant

    For Cnt = 1 To fpSpread4.MaxRows
        fpSpread4.GetText 11, Cnt, TotalPlates¼
        fpSpread4.GetText 12, Cnt, TotalPlates½
        fpSpread4.GetText 13, Cnt, TotalPlates1
        fpSpread4.GetText 14, Cnt, PlateRate
        fpSpread4.GetText 15, Cnt, TotalForms¼
        fpSpread4.GetText 16, Cnt, TotalForms½
        fpSpread4.GetText 17, Cnt, TotalForms1
        fpSpread4.GetText 18, Cnt, PrintRate
        fpSpread4.SetText 7, Cnt, IIf(Cnt = 1, 1, IIf(Cnt = 2, 2, 4)) * (Val(TotalPlates¼) + Val(TotalPlates½) + Val(TotalPlates1)) * Val(PlateRate)
        fpSpread4.SetText 8, Cnt, IIf(Cnt = 1, 1, IIf(Cnt = 2, 2, 4)) * (Val(TotalForms¼) + Val(TotalForms½) + Val(TotalForms1)) * Val(PrintRate)
        If fpSpread4.ActiveRow = Cnt Then
            'MhRealInput7.Text = Format(IIf(Cnt = 1, 1, IIf(Cnt = 2, 2, 4)) * (Val(TotalPlates¼) + Val(TotalPlates½) + Val(TotalPlates1)) * Val(PlateRate), "0.00")
            'MhRealInput8.Text = Format(IIf(Cnt = 1, 1, IIf(Cnt = 2, 2, 4)) * (Val(TotalForms¼) + Val(TotalForms½) + Val(TotalForms1)) * Val(PrintRate), "0.00")
        End If
    Next
    CalculateTotalAmount
End Sub
Private Function CalculateTotalAmount() As Double
    Dim Cnt As Integer, PlateAmount As Variant, PrintAmount As Variant, TotalAmount As Double
    For Cnt = 1 To fpSpread4.MaxRows
        fpSpread4.GetText 7, Cnt, PlateAmount
        fpSpread4.GetText 8, Cnt, PrintAmount
        TotalAmount = TotalAmount + PlateAmount + PrintAmount
    Next
End Function
Private Function CalculateTotalForms(ByVal xPrintingType As String, ByVal Forms As Double, ByVal FormType As String, ByVal MhRealInput2 As Variant, ByVal MhRealInput19 As Variant) As Double
    Dim FS As Variant
    fpSpread4.GetText 21, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), FS
    CalculateTotalForms = (Int(IIf(xPrintingType = "1", Val(MhRealInput2), Val(MhRealInput19)) * IIf(FormType = "¼", 0.25, IIf(FormType = "½", 0.5, 1)) / 1000) + IIf(IIf(xPrintingType = "1", Val(MhRealInput2), Val(MhRealInput19)) * IIf(FormType = "¼", 0.25, IIf(FormType = "½", 0.5, 1)) Mod 1000 = 0, 0, 1)) * Forms
    CalculateTotalForms = IIf(Val(FS) = 0.5, 2, IIf(Val(FS) = 2, 0.5, 1)) * Val(CalculateTotalForms)
    If rstBookList.Fields("DuplexPrinting").Value = "N" Then CalculateTotalForms = 0.5 * CalculateTotalForms
    CalculateTotalForms = Int(Val(CalculateTotalForms)) + IIf(Val(CalculateTotalForms) - Int(Val(CalculateTotalForms)) = 0, 0, 1)
    If FormType = "¼" Then
        fpSpread4.SetText 15, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), CalculateTotalForms
        If fpSpread4.ActiveRow = IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)) Then
        End If
    ElseIf FormType = "½" Then
        fpSpread4.SetText 16, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), CalculateTotalForms
        If fpSpread4.ActiveRow = IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)) Then
        End If
    Else
        fpSpread4.SetText 17, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), CalculateTotalForms
        If fpSpread4.ActiveRow = IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)) Then
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
        MhRealInputBillingQty1.Text = Format(Q1, "0")
    ElseIf Val(xMhRealInput2) <> Q1 Then
        If MsgBox("Variation (Single Color) between Billing Quantity (" & xMhRealInput2 & ") Vs Calculated Billing Quantity (" & Trim(str(Q1)) & ") ! Change Quantity ?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then
            MhRealInputBillingQty1.Text = Format(Q1, "0")
        End If
    End If
    If Val(xMhRealInput19) = 0 Then
        MhRealInputBillingQty2.Text = Format(Q24, "0")
    ElseIf Val(xMhRealInput19) <> Q24 Then
        If MsgBox("Variation (Double & Four Color) between Billing Quantity (" & xMhRealInput19 & ") Vs Calculated Billing Quantity (" & Trim(str(Q24)) & ") ! Change Quantity ?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then
            MhRealInputBillingQty2.Text = Format(Q24, "0")
        End If
    End If
    Call CalculateBQD("S", Format(Q1, "0"), xMhRealInput1, Size_Code, Ac_Code)
    Call CalculateBQD("O", Format(Q1, "0"), xMhRealInput1, Size_Code, Ac_Code)
    Call CalculateConsumption("1", xMhRealInput1): Call CalculateConsumption("2", xMhRealInput1): Call CalculateConsumption("4", xMhRealInput1)
End Sub

Private Sub CalculateBQD(ByVal xPrintingType As String, ByVal BillingQty As Variant, ByVal ActualQty As Variant, ByVal Size_Code As String, ByVal Acc_Code As String) 'Calculate Billing Quantity Dependents
    Dim Cnt As Integer, Content As Variant, Forms As Variant

    For Cnt = IIf(xPrintingType = "S", 1, 2) To IIf(xPrintingType = "S", 1, fpSpread4.MaxRows)
        fpSpread4.GetText 1, Cnt, Content   'Pages
        If Val(Content) <> 0 Then
            GetPrinterRates IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4")), "B", BillingQty, ActualQty, Size_Code, Acc_Code 'Get Both Plate & Printing Rates
        End If
        fpSpread4.GetText 3, Cnt, Forms
        Call CalculateTotalForms(IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4")), Val(Forms), "¼", BillingQty, ActualQty)
        fpSpread4.GetText 4, Cnt, Forms
        Call CalculateTotalForms(IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4")), Val(Forms), "½", BillingQty, ActualQty)
        fpSpread4.GetText 5, Cnt, Forms
        Call CalculateTotalForms(IIf(Cnt = 1, "1", IIf(Cnt = 2, "2", "4")), Val(Forms), "1", BillingQty, ActualQty)
    Next
    CalculateAmount
End Sub

Private Sub fpSpread3_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim Paper As Variant, Qty As Variant, Ac_Code As Variant, Size_Code As Variant, color As Variant, BillingQty1 As Variant, BillingQty2 As Variant, page As Variant, SendQty As Variant, balQty As Variant, TConsumption As Variant
    With fpSpread3
         If Col = 4 Then 'And OrderType = "1"
            .GetText 1, Row, Paper
            .GetText 4, Row, Qty
            .GetText 10, Row, Ac_Code
            .GetText 11, Row, Size_Code
            .GetText 12, Row, page
            .GetText 13, Row, color
            .GetText 15, Row, BillingQty1
            .GetText 16, Row, BillingQty2
            If Paper = "" Then .SetText 4, Row, "" Else .SetText 4, Row, Int(Qty)

            If Qty > 0 And RefCode = "" Then
                 Call CalculateAQD(Qty, Format(BillingQty1, "0"), Format(BillingQty2, "0"), Size_Code, Ac_Code)
                .SetText 7, Row, Format(Val(MhRealInputWastage.Value), "0.000")
                If MhRealInput131.Value > 0 Then
                    .SetText 8, Row, Format(Val(MhRealInput131.Value), "0.000")
                    .SetText 9, Row, Format(Val(MhRealInput1310.Value), "0.000")
                Else
                    .SetText 8, Row, Format(Val(MhRealInput1310.Value), "0.000")
                    .SetText 9, Row, Format(Val(MhRealInput1310.Value), "0.000")
                End If
            End If
            CalculateTotal ("G")
            RefCode = ""
        ElseIf Col = 17 Then
               .GetText 17, Row, SendQty
               .GetText 9, Row, TConsumption

            If SendQty > 0 And SendQty <= Val(MhRealInput20.Value) Then
              '.GetText 18, Row, balQty
              .SetText 18, Row, Format(Val(TConsumption) - SendQty, "0.000")
            Else
              .SetText 17, Row, 0
              .SetText 18, Row, 0

            End If
         End If

    End With
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
        fpSpread4.GetText 6, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), PlateType
        PlateRate = rstPrinterRates.Fields(PlateType & "PlateRate" & Trim(xPrintingType)).Value
        PrintRate = rstPrinterRates.Fields("PrintRate" & Trim(xPrintingType)).Value
        PrintRate = PrintRate + IIf(PrintRate > 0, Val(rstBookList.Fields("AddOnRate01").Value), 0)
        PaperWastageRate = Val(rstPrinterRates.Fields("PaperWastageRate" & Trim(xPrintingType)))
        MhRealInputWastage.Text = Format(PaperWastageRate, "0.000")
    End If
    fpSpread4.GetText 14, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), CurrentRate  'Plate Rate
    If CurrentRate <> PlateRate Then
'        If Val(CheckNull(MhRealInput19)) <> 0 Then
'            If MsgBox(IIf(xPrintingType = "1", "Single", IIf(xPrintingType = "2", "Double", "Four")) + " Color(s) Plate rate is different from that in Master ! Change rate?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then
'                fpSpread4.SetText 14, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), PlateRate
'            End If
'        Else
            fpSpread4.SetText 14, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), PlateRate
'        End If
    End If
    If xRateType = "B" Then
        fpSpread4.GetText 18, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), CurrentRate  'Print Rate
        If CurrentRate <> PrintRate And CurrentRate <> 0 Then
'            If MsgBox(IIf(xPrintingType = "1", "Single", IIf(xPrintingType = "2", "Double", "Four")) + " Color(s) Printing Rate is different from that in Master ! Change Rate?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then
'                fpSpread4.SetText 18, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), PrintRate
'            End If
        Else
            fpSpread4.SetText 18, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), PrintRate
        End If
        fpSpread4.GetText 9, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), CurrentRate   'Paper Wastage Rate
        If CurrentRate <> PaperWastageRate Then
'            If Val(CheckNull(rstBookList.Fields("ActualQuantity").Value)) <> 0 Then
'                If MsgBox(IIf(xPrintingType = "1", "Single", IIf(xPrintingType = "2", "Double", "Four")) + " Color(s) Paper Wastage is different from that in Master ! Change Wastage?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Change !") = vbYes Then
'                    fpSpread4.SetText 9, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), PaperWastageRate
'                End If
'            Else
                fpSpread4.SetText 9, IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)), PaperWastageRate
'            End If
        End If
    End If
    If fpSpread4.ActiveRow = IIf(xPrintingType = "1", 1, IIf(xPrintingType = "2", 2, 3)) Then
        fpSpread4.GetText 14, fpSpread4.ActiveRow, CurrentRate  'Plate Rate
        fpSpread4.GetText 18, fpSpread4.ActiveRow, CurrentRate  'Print Rate
        fpSpread4.GetText 9, fpSpread4.ActiveRow, CurrentRate   'Paper Wastage Rate
    End If
    Exit Sub
ErrorHandler:
    DisplayError ("Failed to Fetch Printer Rates")
End Sub

Private Sub fpSpread3_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask And KeyCode = vbKeyD Then
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            fpSpread3.DeleteRows fpSpread2.ActiveRow, 1: fpSpread3.SetFocus
            CalculateTotal ("G")
        End If

    ElseIf KeyCode = vbKeySpace Then
        Dim Paper As Variant, Book As Variant, Reference As Variant
        With fpSpread3
            .GetText 1, .ActiveRow, Paper
            If .ActiveCol = 1 Then
                If Paper = "" Then
                    fpSpread1.GetText 1, fpSpread1.ActiveRow, Paper
                    .SetText 1, .ActiveRow, Paper
                    fpSpread1.GetText 7, fpSpread1.ActiveRow, Paper
                    .SetText 6, .ActiveRow, Paper
                     'Added On 8 July 15 to add printer code
                    fpSpread2.GetText 5, fpSpread2.ActiveRow, Paper
                    .SetText 10, .ActiveRow, Paper
                    If Paper <> "" Then SendKeys "{ENTER}"
                End If
            ElseIf .ActiveCol = 2 Then
                If Paper <> "" Then
                    .GetText 2, .ActiveRow, Book
                    Text6.Text = FixQuote(Book)
                    If rstBookList.RecordCount = 0 Then DisplayError ("No Record in Book Master"): .SetActiveCell 1, .ActiveRow: Exit Sub Else rstBookList.MoveFirst
                    rstBookList.Find "[Col0] = '" & RTrim(Book) & "'"
                    SelectionType = "S"
                    BookCode = ""
                    Call LoadSelectionList(rstBookList, "List of Books...", "Name")
                    SearchOrder = 0
                    Call DisplaySelectionList(Text6, BookCode)
                    Call CloseForm(FrmSelectionList)
                    If BookCode = "" Then
                        .SetActiveCell 2, .ActiveRow
                    Else
                         rstBookList.MoveFirst: rstBookList.Find "[Code] ='" & BookCode & "'"
                        .SetText 2, .ActiveRow, Text6.Text
                        .SetText 3, .ActiveRow, Trim(rstBookList.Fields("Forms").Value)
                        Call LoadRefList(BookCode)

'                        If Not rstPlaningRef.EOF Then
'                          rstPlaningRef.MoveFirst: rstPlaningRef.Find "[BookCode] ='" & BookCode & "'"
'                        End If
'                        If Not rstPlaningRef.EOF Then
'                            'Get From Planing References
'                            .SetText 4, .ActiveRow, Trim(rstPlaningRef.Fields("Quantity").Value)
'                            .SetText 7, .ActiveRow, Trim(rstPlaningRef.Fields("PaperWastage").Value)
'                            .SetText 8, .ActiveRow, Trim(rstPlaningRef.Fields("PaperConsumption").Value)
'                            .SetText 9, .ActiveRow, Trim(rstPlaningRef.Fields("PaperConsumption").Value)
'                        End If

                        .SetText 5, .ActiveRow, BookCode
                        .SetText 11, .ActiveRow, Trim(rstBookList.Fields("SizeCode").Value)
                        .SetText 12, .ActiveRow, Trim(rstBookList.Fields("Pages").Value)


                        If Val(rstBookList.Fields("TwoColorPages").Value) = 0 And Val(rstBookList.Fields("FourColorPages").Value) = 0 Then
                          .SetText 13, .ActiveRow, "1 Color"
                        ElseIf Val(rstBookList.Fields("OneColorPages").Value) = 0 And Val(rstBookList.Fields("FourColorPages").Value) = 0 Then
                          .SetText 13, .ActiveRow, "2 Color"
                        ElseIf Val(rstBookList.Fields("OneColorPages").Value) = 0 And Val(rstBookList.Fields("TwoColorPages").Value) = 0 Then
                          .SetText 13, .ActiveRow, "4 Color"
                        ElseIf Val(rstBookList.Fields("OneColorPages").Value) = 0 And Val(rstBookList.Fields("TwoColorPages").Value) = 0 And Val(rstBookList.Fields("FourColorPages").Value) = 0 Then
                          .SetText 13, .ActiveRow, "6 Color"
                        Else
                          .SetText 13, .ActiveRow, "Multi Color"
                        End If

                        .SetText 15, .ActiveRow, 0
                        .SetText 16, .ActiveRow, 0

                        Dim Cnt As Integer
                            For Cnt = 1 To fpSpread4.MaxRows
                                  fpSpread4.SetText 1, Cnt, Val(rstBookList.Fields(IIf(Cnt = 1, "One", IIf(Cnt = 2, "Two", "Four")) & "ColorPages").Value)
                                  fpSpread4.SetText 2, Cnt, Val(rstBookList.Fields(IIf(Cnt = 1, "One", IIf(Cnt = 2, "Two", "Four")) & "ColorForms").Value)
                                  fpSpread4.SetText 3, Cnt, Val(rstBookList.Fields(IIf(Cnt = 1, "One", IIf(Cnt = 2, "Two", "Four")) & "Color¼Forms").Value)
                                  fpSpread4.SetText 4, Cnt, Val(rstBookList.Fields(IIf(Cnt = 1, "One", IIf(Cnt = 2, "Two", "Four")) & "Color½Forms").Value)
                                  fpSpread4.SetText 5, Cnt, Val(rstBookList.Fields(IIf(Cnt = 1, "One", IIf(Cnt = 2, "Two", "Four")) & "Color1F/BForms").Value) + Val(rstBookList.Fields(IIf(Cnt = 1, "One", IIf(Cnt = 2, "Two", "Four")) & "Color1W/TForms").Value)
                                  fpSpread4.SetText 6, Cnt, IIf(rstBookList.Fields(IIf(Cnt = 1, "One", IIf(Cnt = 2, "Two", "Four")) & "ColorPlateType").Value = "1", "Deepatch", IIf(rstBookList.Fields(IIf(Cnt = 1, "One", IIf(Cnt = 2, "Two", "Four")) & "ColorPlateType").Value = "2", "PS", IIf(rstBookList.Fields(IIf(Cnt = 1, "One", IIf(Cnt = 2, "Two", "Four")) & "ColorPlateType").Value = "3", "Wipeon", "CTP")))
                                  fpSpread4.SetText 7, Cnt, 0#
                                  fpSpread4.SetText 8, Cnt, 0#
                                  fpSpread4.SetText 9, Cnt, 0#
                                  fpSpread4.SetText 10, Cnt, 0#
                            Next

                      SendKeys "{ENTER}"
                     End If
                End If

            ElseIf .ActiveCol = 4 Then
                If BookCode <> "" Then

                    .GetText 4, .ActiveRow, Reference
                    Text6.Text = FixQuote(Reference)
                    If rstRefList.RecordCount = 0 Then DisplayError ("No Record in Planning"): .SetActiveCell 2, .ActiveRow: Exit Sub Else rstRefList.MoveFirst
                    rstRefList.Find "[Col0] = '" & RTrim(Reference) & "'"
                    SelectionType = "S"
                    RefCode = ""
                    Call LoadSelectionList(rstRefList, "List of Reference...", "Ref.No")
                    SearchOrder = 0
                    Call DisplaySelectionList(Text6, RefCode)
                    Call CloseForm(FrmSelectionList)

                    If RefCode = "" Then
                        .SetActiveCell 4, .ActiveRow
                    Else

                        If Not rstRefList.EOF Then
                          rstRefList.MoveFirst: rstRefList.Find "[Code] ='" & RefCode & "'"
                        End If
                        If Not rstRefList.EOF Then
                            .SetText 4, .ActiveRow, Trim(rstRefList.Fields("Quantity").Value) 'Text6.Text
                            .SetText 7, .ActiveRow, Trim(rstRefList.Fields("PaperWastage").Value)
                            .SetText 8, .ActiveRow, Trim(rstRefList.Fields("PaperConsumption").Value)
                            .SetText 9, .ActiveRow, Trim(rstRefList.Fields("PaperConsumption").Value)
                         End If

                        SendKeys "{ENTER}"
                    End If
                End If

            End If
        End With
    End If
End Sub

Private Sub LoadRefList(ByVal strBookCode As String)
    Dim BalanceQuantity As Long
    On Error GoTo ErrorHandler
    If rstRefList.State = adStateOpen Then
        rstRefList.Close
    End If

    Dim aa As String
    aa = "Select (Trim(PrintPVParent.Name)+' Quantity : ' + Cstr(Quantity)) As Col0,Trim(PrintPVParent.Name) As Code,PrintPVChild.Forms,Quantity, [PaperWastage%] As PaperWastage,PaperConsumption,Trim(PrintName) As BookName ,Book As BookCode From (PrintPVParent Inner Join PrintPVChild On (PrintPVParent.Code = PrintPVChild.Code)) Inner Join BookMaster On PrintPVChild.Book = BookMaster.Code Where PrintPVChild.Code Not in(Select Distinct Ref From BookPOChild05 Where Ref<>'')And PrintPVChild.Book='" & strBookCode & "' Order By BookMaster.PrintName"

    rstRefList.Open "Select (Trim(PrintPVParent.Name)+' Quantity : ' + Cstr(Quantity)) As Col0,Trim(PrintPVParent.Name) As Code,PrintPVChild.Forms,Quantity, [PaperWastage%] As PaperWastage,PaperConsumption,Trim(PrintName) As BookName ,Book As BookCode From (PrintPVParent Inner Join PrintPVChild On (PrintPVParent.Code = PrintPVChild.Code)) Inner Join BookMaster On PrintPVChild.Book = BookMaster.Code Where PrintPVChild.Code Not in (Select  Distinct M.Ref From BookPOChild05 M Inner Join BookPOParent P On M.Code=P.Code  Where M.Ref<>'' And P.Book='" & strBookCode & "')And PrintPVChild.Book='" & strBookCode & "' Order By BookMaster.PrintName", CxnDatabase, adOpenKeyset, adLockOptimistic
    rstRefList.ActiveConnection = Nothing


'    Do While Not rstRefList.EOF
'            rstRefList.Fields("Col0").Value = Trim(rstRefList.Fields("Name").Value) + " Quantity : " + Format(str(rstRefList.Fields("Quantity").Value), "#0")
'            rstRefList.Update
'            rstRefList.MoveNext
'    Loop
    Exit Sub
ErrorHandler:
    DisplayError ("Failed to Load Ref List")
End Sub



