VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0F1F1508-C40A-101B-AD04-00AA00575482}#1.0#0"; "mhrinp32.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmMaterialIssueOrder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Material Issue Order"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   8715
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MaterialIssueOrder.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4875
   ScaleWidth      =   8715
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   4870
      Left            =   15
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   8700
      _Version        =   65536
      _ExtentX        =   15346
      _ExtentY        =   8590
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
      Picture         =   "MaterialIssueOrder.frx":0442
      Begin TabDlg.SSTab SSTab1 
         Height          =   4630
         Left            =   120
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   120
         Width           =   8460
         _ExtentX        =   14923
         _ExtentY        =   8176
         _Version        =   393216
         Style           =   1
         Tabs            =   2
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
         TabPicture(0)   =   "MaterialIssueOrder.frx":045E
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "DataGrid1"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Text1"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "&Details"
         TabPicture(1)   =   "MaterialIssueOrder.frx":047A
         Tab(1).ControlEnabled=   0   'False
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
            Left            =   600
            TabIndex        =   9
            Top             =   4160
            Width           =   7760
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   3640
            Left            =   120
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   450
            Width           =   8235
            _ExtentX        =   14526
            _ExtentY        =   6429
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
               DataField       =   "SourceName"
               Caption         =   "Source Name"
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
                  ColumnWidth     =   5564.977
               EndProperty
            EndProperty
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   4000
            Left            =   -74880
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   480
            Width           =   8235
            _Version        =   65536
            _ExtentX        =   14526
            _ExtentY        =   7056
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
            Picture         =   "MaterialIssueOrder.frx":0496
            Begin FPSpreadADO.fpSpread fpSpread1 
               Height          =   2415
               Left            =   120
               TabIndex        =   4
               Top             =   1470
               Width           =   8010
               _Version        =   524288
               _ExtentX        =   14129
               _ExtentY        =   4260
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
               MaxCols         =   9
               MaxRows         =   100
               OperationMode   =   2
               SpreadDesigner  =   "MaterialIssueOrder.frx":04B2
            End
            Begin MhinrelLib.MhRealInput MhRealInput4 
               Height          =   255
               Left            =   6890
               TabIndex        =   17
               TabStop         =   0   'False
               Top             =   3625
               Width           =   1005
               _Version        =   65536
               _ExtentX        =   1773
               _ExtentY        =   450
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Comic Sans MS"
                  Size            =   8.25
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
               DecimalPlaces   =   0
               VAlignment      =   2
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel19 
               Height          =   255
               Left            =   120
               TabIndex        =   16
               Top             =   3625
               Width           =   8010
               _Version        =   65536
               _ExtentX        =   14129
               _ExtentY        =   450
               _StockProps     =   77
               BackColor       =   32896
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Comic Sans MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   0   'False
               TintColor       =   16711935
               Caption         =   ""
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "MaterialIssueOrder.frx":0AF8
               Picture         =   "MaterialIssueOrder.frx":0B14
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
               Left            =   1440
               MaxLength       =   10
               TabIndex        =   0
               Top             =   105
               Width           =   1650
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
               Left            =   1440
               MaxLength       =   139
               TabIndex        =   3
               Top             =   950
               Width           =   6690
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
               Left            =   1440
               MaxLength       =   40
               TabIndex        =   2
               Top             =   630
               Width           =   6690
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
               Height          =   330
               Left            =   120
               TabIndex        =   12
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
               Caption         =   " Order No."
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "MaterialIssueOrder.frx":0B30
               Picture         =   "MaterialIssueOrder.frx":0B4C
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Left            =   5955
               TabIndex        =   13
               Top             =   105
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
               _ExtentY        =   582
               _StockProps     =   77
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
               Picture         =   "MaterialIssueOrder.frx":0B68
               Picture         =   "MaterialIssueOrder.frx":0B84
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   330
               Left            =   120
               TabIndex        =   14
               Top             =   630
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
               Caption         =   " Source Name"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "MaterialIssueOrder.frx":0BA0
               Picture         =   "MaterialIssueOrder.frx":0BBC
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
               Height          =   330
               Left            =   120
               TabIndex        =   15
               Top             =   945
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
               Caption         =   " Remarks"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "MaterialIssueOrder.frx":0BD8
               Picture         =   "MaterialIssueOrder.frx":0BF4
            End
            Begin TDBDate6Ctl.TDBDate MhDateInput1 
               Height          =   330
               Left            =   7035
               TabIndex        =   1
               Top             =   105
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
               _ExtentY        =   582
               Calendar        =   "MaterialIssueOrder.frx":0C10
               Caption         =   "MaterialIssueOrder.frx":0D28
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "MaterialIssueOrder.frx":0D94
               Keys            =   "MaterialIssueOrder.frx":0DB2
               Spin            =   "MaterialIssueOrder.frx":0E10
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
            Begin VB.TextBox Text9 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   3840
               TabIndex        =   18
               Top             =   2520
               Width           =   2535
            End
            Begin VB.Line Line1 
               X1              =   0
               X2              =   8280
               Y1              =   525
               Y2              =   525
            End
            Begin VB.Line Line2 
               X1              =   0
               X2              =   8280
               Y1              =   1365
               Y2              =   1365
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
            ForeColor       =   &H8000000E&
            Height          =   330
            Left            =   120
            TabIndex        =   10
            Top             =   4160
            Width           =   495
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   8715
      _ExtentX        =   15372
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
            Enabled         =   0   'False
            Object.Visible         =   0   'False
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
Attribute VB_Name = "FrmMaterialIssueOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CxnMaterialIssueOrder As New ADODB.Connection
Dim rstCompanyMaster As New ADODB.Recordset
Dim rstMaterialIOList As New ADODB.Recordset
Dim rstMaterialIOParent As New ADODB.Recordset
Dim rstMaterialIOChild As New ADODB.Recordset
Dim rstGodownList As New ADODB.Recordset
Dim rstSourceList As New ADODB.Recordset
Dim rstRefList As New ADODB.Recordset
Dim rstFreshBookList As New ADODB.Recordset
Dim rstRepairBookList As New ADODB.Recordset
Dim SourceCode As String
Dim RefCode As String
Dim FreshBook As String
Dim RepairBook As String
Dim Title As String
Dim godown As String
Dim EditMode As Boolean
Dim T As Integer
Dim SortOrder As String
Dim PrevStr As String
Dim dblBookMark As Double
Dim blnRecordExist As Boolean
Dim OutputTo As String

Private Sub Form_Load()
    On Error GoTo ErrorHandler
    
    CenterForm Me
    BusySystemIndicator True
    CxnMaterialIssueOrder.CursorLocation = adUseClient
    CxnMaterialIssueOrder.Open CxnDatabase.ConnectionString
    rstCompanyMaster.Open "Select PrintName, Address1, Address2, Address3, Address4, Phone, Fax, EMail, Website From CompanyMaster", CxnMaterialIssueOrder, adOpenKeyset, adLockReadOnly
    rstSourceList.Open "Select Name As Col0, Code From AccountMaster Where Type In ('01','05','07','09') Order by Name", CxnMaterialIssueOrder, adOpenKeyset, adLockReadOnly
    rstGodownList.Open "Select Name As Col0,Code From AccountMaster Where Type In ('08','09') Order By Name", CxnMaterialIssueOrder, adOpenKeyset, adLockReadOnly
    rstFreshBookList.Open "Select Name,Board,'3'+Code As NCode From BookMaster Where Type='F' Order By Name", CxnMaterialIssueOrder, adOpenKeyset, adLockOptimistic
    rstRepairBookList.Open "Select Name,'4'+Code As NCode From BookMaster Where Type='R' Order By Name", CxnMaterialIssueOrder, adOpenKeyset, adLockOptimistic
    rstMaterialIOList.Open "Select T.Code,T.Name,T.Date,M.Name As SourceName From MaterialIOParent T, AccountMaster M Where T.Source = M.Code And T.Type='0' Order By T.Name", CxnMaterialIssueOrder, adOpenKeyset, adLockOptimistic
    rstMaterialIOParent.CursorLocation = adUseClient
    rstMaterialIOList.Filter = adFilterNone
    If rstMaterialIOList.RecordCount > 0 Then rstMaterialIOList.MoveLast
    Set DataGrid1.DataSource = rstMaterialIOList
    BusySystemIndicator False
    SSTab1.Tab = 0
    SortOrder = "Name"
    If Not (rstMaterialIOList.EOF Or rstMaterialIOList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    rstMaterialIOList.ActiveConnection = Nothing
    rstGodownList.ActiveConnection = Nothing
    rstSourceList.ActiveConnection = Nothing
    rstFreshBookList.ActiveConnection = Nothing
    rstRepairBookList.ActiveConnection = Nothing
    Call RefreshDropDownList("A")
    With fpSpread1
        .Col = 6
        .ColHidden = True
        .Col = 7
        .ColHidden = True
        .Col = 8
        .ColHidden = True
        .Col = 9
        .ColHidden = True
        .Col = 10
        .ColHidden = True
    End With
    SetButtonsForNoRecord
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    Unload Me
End Sub
Private Sub Form_Activate()
    EnableChildMenu True
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
            If Not EditMode Then
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
        If Not EditMode Then
            Toolbar1_ButtonClick Toolbar1.Buttons.Item(4)
        End If
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
           If Me.ActiveControl.Name <> "fpSpread1" Then
              SendKeys "{TAB}"
           End If
        End If
        If Me.ActiveControl.Name <> "fpSpread1" Then
            KeyCode = 0
        End If
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Toolbar1.Buttons.Item(4).Enabled Then
        Call Form_KeyDown(vbKeyEscape, 0)
        Cancel = 1
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CloseRecordset(rstCompanyMaster)
    Call CloseRecordset(rstMaterialIOList)
    Call CloseRecordset(rstMaterialIOParent)
    Call CloseRecordset(rstMaterialIOChild)
    Call CloseRecordset(rstGodownList)
    Call CloseRecordset(rstSourceList)
    Call CloseRecordset(rstRefList)
    Call CloseRecordset(rstFreshBookList)
    Call CloseRecordset(rstRepairBookList)
    Call CloseConnection(CxnMaterialIssueOrder)
    ShowProgressInStatusBar False
    DisableChildMenu
End Sub

Private Sub Text1_Change()
    If rstMaterialIOList.RecordCount = 0 Then Exit Sub
    rstMaterialIOList.MoveFirst
    If Text1.Text <> "" Then
        If SortOrder = "Name" Then
           rstMaterialIOList.Find "[" & SortOrder & "] Like '%" & FixQuote(Text1.Text) & "%'"
        Else
           rstMaterialIOList.Find "[" & SortOrder & "] Like '" & FixQuote(Text1.Text) & "%'"
        End If
        If rstMaterialIOList.EOF Then
            rstMaterialIOList.MoveFirst
            If PrevStr <> "" And Len(Text1.Text) > 1 Then
                If dblBookMark <> 0 Then
                    rstMaterialIOList.Bookmark = dblBookMark
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
    If Not (rstMaterialIOList.EOF Or rstMaterialIOList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim KeyProcessed As Boolean
    
    If rstMaterialIOList.RecordCount = 0 Then Exit Sub
    If Shift = 0 And KeyCode = vbKeyUp Then
        With rstMaterialIOList
            .MovePrevious
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyBack Then
        With rstMaterialIOList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyDown Then
        With rstMaterialIOList
            .MoveNext
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageUp Then
        With rstMaterialIOList
            .Move (-1) * (DataGrid1.VisibleRows - 1)
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageUp Then
        With rstMaterialIOList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageDown Then
        With rstMaterialIOList
            .Move DataGrid1.VisibleRows - 1
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageDown Then
        With rstMaterialIOList
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
            If Not (rstMaterialIOList.EOF Or rstMaterialIOList.BOF) Then
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
        Text2.SetFocus
    End If
End Sub
Public Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim HiLiteRecord As Boolean
    Dim UpdateFlag As Integer, i As Integer
    Dim CellVal As Variant
    
    If Button.Index = 1 Then
        If rstMaterialIOParent.State = adStateOpen Then
           rstMaterialIOParent.Close
        End If
        rstMaterialIOParent.Open "Select * From MaterialIOParent Where Code = ''", CxnMaterialIssueOrder, adOpenKeyset, adLockOptimistic
        ClearFields
        If AddRecord(rstMaterialIOParent) Then
            Text2.Text = GenerateCode(CxnMaterialIssueOrder, "Select Max(Val(Name)) From MaterialIOParent Where Type='0'", 10, Space(1))
            MhDateInput1.Text = Format(Date, "dd-MM-yyyy")
            Call SetButtons(False)
            SSTab1.Tab = 1
            Text2.SetFocus
            blnRecordExist = False
            CxnMaterialIssueOrder.BeginTrans
        End If
    ElseIf Button.Index = 2 Then
        If rstMaterialIOList.RecordCount = 0 Then Exit Sub
        SSTab1.Tab = 1
        EditRecord
    ElseIf Button.Index = 3 Then
        If rstMaterialIOList.RecordCount = 0 Then Exit Sub
        If AllowTransactionsDeletion = 0 Then
            Call DisplayError("You don't have the rights to Delete this Voucher")
            Exit Sub
        End If
        SSTab1.Tab = 1
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            On Error Resume Next
            MdiMainMenu.MousePointer = vbHourglass
            CxnMaterialIssueOrder.Execute "Delete From MaterialIOParent Where Code = '" & rstMaterialIOList.Fields("Code").Value & "'"
            MdiMainMenu.MousePointer = vbNormal
            If Err.Number = 0 Then
                rstMaterialIOList.Delete
                rstMaterialIOList.MoveNext
                If rstMaterialIOList.RecordCount > 0 And rstMaterialIOList.EOF Then
                    rstMaterialIOList.MoveLast
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
        If blnRecordExist And AllowTransactionsModification = 0 Then
            Call DisplayError("You don't have the rights to Edit this Voucher")
            Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
            Exit Sub
        End If
        SaveFields
        UpdateFlag = 0
        If UpdateRecord(rstMaterialIOParent) Then
            If UpdateMaterialList("D") Then
                UpdateFlag = 1
                For i = 1 To fpSpread1.DataRowCnt
                    fpSpread1.SetActiveCell 5, i
                    fpSpread1.GetText 5, i, CellVal
                    If Val(CellVal) <> 0 Then
                        If Not UpdateMaterialList("I") Then
                            UpdateFlag = 0
                            Exit For
                        End If
                    End If
                Next
            End If
        End If
        If UpdateFlag Then
            AddToList
            CxnMaterialIssueOrder.CommitTrans
            If rstMaterialIOParent.State = adStateOpen Then
                rstMaterialIOParent.Close
            End If
            rstMaterialIOParent.CursorLocation = adUseClient
            Call SetButtons(True)
            SSTab1.Tab = 0
            ShowProgressInStatusBar True
            Timer1.Enabled = True
        Else
            DisplayError ("Failed to save the record")
            Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
        End If
    ElseIf Button.Index = 5 Then
        If CancelRecordUpdate(rstMaterialIOParent) Then
            CxnMaterialIssueOrder.RollbackTrans
            If rstMaterialIOParent.State = adStateOpen Then
                rstMaterialIOParent.Close
            End If
            rstMaterialIOParent.CursorLocation = adUseClient
            Call SetButtons(True)
            SetButtonsForNoRecord
            SSTab1.Tab = 0
        End If
    ElseIf Button.Index = 6 Then
        SSTab1.Tab = 0
        Set DataGrid1.DataSource = Nothing
        rstMaterialIOList.ActiveConnection = CxnMaterialIssueOrder
        Do While Not RefreshRecord(rstMaterialIOList)
        Loop
        Set DataGrid1.DataSource = rstMaterialIOList
        rstMaterialIOList.ActiveConnection = Nothing
        If rstMaterialIOList.RecordCount > 0 Then rstMaterialIOList.MoveLast
        rstSourceList.ActiveConnection = CxnMaterialIssueOrder
        Do While Not RefreshRecord(rstSourceList)
        Loop
        rstSourceList.ActiveConnection = Nothing
        HiLiteRecord = True
    ElseIf Button.Index = 7 Then
        SSTab1.Tab = 0
        With FrmFilter
            .Combo1.AddItem "Source", 0
            .Combo1.ListIndex = 0
            Set .srcForm = Me
            .Show vbModal
        End With
        HiLiteRecord = True
    ElseIf Button.Index = 9 Then
        If rstMaterialIOList.RecordCount = 0 Then Exit Sub
        OutputTo = "P"
        PrintMaterialIssueOrder
        HiLiteRecord = True
    ElseIf Button.Index = 10 Then
        If rstMaterialIOList.RecordCount = 0 Then Exit Sub
        OutputTo = "S"
        PrintMaterialIssueOrder
        HiLiteRecord = True
    ElseIf Button.Index = 13 Then
        If rstMaterialIOList.RecordCount > 0 Then rstMaterialIOList.MoveFirst
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 14 Then
        If rstMaterialIOList.RecordCount > 0 Then
            rstMaterialIOList.MovePrevious
            If rstMaterialIOList.BOF Then
                rstMaterialIOList.MoveNext
            End If
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 15 Then
        If rstMaterialIOList.RecordCount > 0 Then
            rstMaterialIOList.MoveNext
            If rstMaterialIOList.EOF Then
                rstMaterialIOList.MovePrevious
            End If
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 16 Then
        If rstMaterialIOList.RecordCount > 0 Then rstMaterialIOList.MoveLast
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 18 Then
        Unload Me
        HiLiteRecord = False
    End If
    If HiLiteRecord Then
        If Not (rstMaterialIOList.EOF Or rstMaterialIOList.BOF) Then
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
Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
    If ColIndex = 0 Then
       If SortOrder <> "Name" Then
          SortOrder = "Name"
          rstMaterialIOList.Sort = "Name Asc"
       End If
    ElseIf ColIndex = 2 Then
       If SortOrder <> "SourceName" Then
          SortOrder = "SourceName"
          rstMaterialIOList.Sort = "SourceName Asc"
       End If
    End If
    DataGrid1.ClearSelCols
    If Not (rstMaterialIOList.EOF Or rstMaterialIOList.BOF) Then
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
    Toolbar1.Buttons.Item(13).Enabled = bVal
    Toolbar1.Buttons.Item(14).Enabled = bVal
    Toolbar1.Buttons.Item(15).Enabled = bVal
    Toolbar1.Buttons.Item(16).Enabled = bVal
    Toolbar1.Buttons.Item(18).Enabled = bVal
    Mh3dFrame2.Enabled = Not bVal
End Sub
Private Sub SetButtonsForNoRecord()
    If rstMaterialIOList.RecordCount = 0 Then
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
    If rstMaterialIOParent.EOF Or rstMaterialIOParent.BOF Then Exit Sub
    If CheckEmpty(Text2, True) Then
        Cancel = True
    ElseIf CheckDuplicate(CxnMaterialIssueOrder, "MaterialIOParent", "Code", "[Name]+[Type]", Trim(Text2.Text) & "0", rstMaterialIOParent.Fields("Code").Value, False) Then
        Cancel = True
    End If
End Sub
Private Sub MhDateInput1_Validate(Cancel As Boolean)
    If Not IsDate(GetDate(MhDateInput1.Text)) Then
        Cancel = True
    ElseIf Format(GetDate(MhDateInput1.Text), "yyyymmdd") < Format(FinancialYearFrom, "yyyymmdd") Or Format(GetDate(MhDateInput1.Text), "yyyymmdd") > Format(FinancialYearTo, "yyyymmdd") Then
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
    If rstSourceList.RecordCount = 0 Then
        DisplayError ("No Record in Source Master")
        Cancel = True
        Exit Sub
    Else
        rstSourceList.MoveFirst
    End If
    rstSourceList.Find "[Col0] = '" & RTrim(SearchString) & "'"
    If rstSourceList.EOF Then
        SelectionType = "S"
        SourceCode = ""
        Call LoadSelectionList(rstSourceList, "List of Sources...", "Name")
        SearchOrder = 0
        Call DisplaySelectionList(Text3, SourceCode)
        Call CloseForm(FrmSelectionList)
        If CheckEmpty(Text3.Text, False) Then
            Text3.Text = "?"
        End If
        If RTrim(SourceCode) <> "" Then
            SendKeys "{TAB}"
        End If
        Cancel = True
    Else
        SourceCode = rstSourceList.Fields("Code").Value
    End If
End Sub
Private Sub ViewRecord()
    ClearFields
    If rstMaterialIOList.EOF Then
        If rstMaterialIOChild.State = adStateOpen Then
            rstMaterialIOChild.Close
        End If
        Exit Sub
    End If
    FindRecord
    LoadFields
End Sub
Private Sub FindRecord()
    If rstMaterialIOParent.State = adStateOpen Then
       rstMaterialIOParent.Close
    End If
    rstMaterialIOParent.Open "Select * From MaterialIOParent Where Code = '" & FixQuote(rstMaterialIOList.Fields("Code").Value) & "'", CxnMaterialIssueOrder, adOpenKeyset, adLockOptimistic
    If rstMaterialIOParent.RecordCount = 0 Then
       Call DisplayError("This Record has been deleted by Another User ! Click Ok To Refresh the Recordset")
       Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
    End If
End Sub
Private Sub ClearFields()
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    MhDateInput1.Text = Format(Date, "dd-MM-yyyy")
    MhRealInput4.Text = 0#
    fpSpread1.ClearRange 1, 1, fpSpread1.MaxCols, fpSpread1.MaxRows, True
End Sub
Private Sub LoadFields()
    If rstMaterialIOParent.EOF Or rstMaterialIOParent.BOF Then Exit Sub
    Text2.Text = rstMaterialIOParent.Fields("Name").Value
    MhDateInput1.Text = Format(rstMaterialIOParent.Fields("Date").Value, "dd-MM-yyyy")
    SourceCode = rstMaterialIOParent.Fields("Source").Value
    If rstSourceList.RecordCount > 0 Then rstSourceList.MoveFirst
    rstSourceList.Find "[Code] = '" & SourceCode & "'"
    If Not rstSourceList.EOF Then
       Text3.Text = rstSourceList.Fields("Col0").Value
    End If
    Text4.Text = rstMaterialIOParent.Fields("Remarks").Value
    Call LoadMaterialList(rstMaterialIOParent.Fields("Code").Value)
End Sub
Private Sub EditRecord()
    On Error GoTo ErrorHandler
    
    If rstMaterialIOParent.RecordCount = 0 Then Exit Sub
    If rstMaterialIOChild.State = adStateClosed Then
        SSTab1.Tab = 0
        Exit Sub
    End If
    If rstMaterialIOParent.State = adStateOpen Then
       rstMaterialIOParent.Close
    End If
    rstMaterialIOParent.CursorLocation = adUseServer
    rstMaterialIOParent.Open "Select * From MaterialIOParent Where Code = '" & FixQuote(rstMaterialIOList.Fields("Code").Value) & "'", CxnMaterialIssueOrder, adOpenKeyset, adLockPessimistic
    MdiMainMenu.MousePointer = vbHourglass
    rstMaterialIOParent.Fields("Printstatus") = "N"
    MdiMainMenu.MousePointer = vbNormal
    AddToList
    Call SetButtons(False)
    SSTab1.TabEnabled(0) = False
    Text2.SetFocus
    blnRecordExist = True
    CxnMaterialIssueOrder.BeginTrans
    Exit Sub
ErrorHandler:
    If Err.Number = -2147467259 Then
       Call DisplayError("Failed to Edit the record")
    End If
    MdiMainMenu.MousePointer = vbNormal
    SSTab1.Tab = 0
End Sub
Private Sub SaveFields()
    If rstMaterialIOParent.EOF Or rstMaterialIOParent.BOF Then Exit Sub
    If Not blnRecordExist Then
        rstMaterialIOParent.Fields("Code").Value = GenerateCode(CxnMaterialIssueOrder, "Select Max(Code) From MaterialIOParent", 6, "0")
        rstMaterialIOParent.Fields("CreatedBy").Value = UserCode
        rstMaterialIOParent.Fields("CreatedOn").Value = Now()
        rstMaterialIOParent.Fields("Recordstatus").Value = "N"
    Else
        rstMaterialIOParent.Fields("ModifiedBy").Value = UserCode
        rstMaterialIOParent.Fields("ModifiedOn").Value = Now()
        rstMaterialIOParent.Fields("Recordstatus").Value = "M"
    End If
    rstMaterialIOParent.Fields("Name").Value = Pad(Trim(Text2.Text), Space(1), 10, "L")
    rstMaterialIOParent.Fields("Date").Value = GetDate(MhDateInput1.Text)
    rstMaterialIOParent.Fields("Source").Value = SourceCode
    rstMaterialIOParent.Fields("Type").Value = "0"
    rstMaterialIOParent.Fields("Remarks").Value = Trim(Text4.Text)
    rstMaterialIOParent.Fields("PrintStatus").Value = "N"
End Sub
Private Sub AddToList()
    On Error Resume Next
    
    rstMaterialIOList.MoveFirst
    rstMaterialIOList.Find "[Code] = '" & rstMaterialIOParent.Fields("Code").Value & "'"
    If rstMaterialIOList.EOF Then
       rstMaterialIOList.AddNew
       rstMaterialIOList.Fields("Code").Value = rstMaterialIOParent.Fields("Code").Value
    End If
    rstMaterialIOList.Fields("Name").Value = Pad(rstMaterialIOParent.Fields("Name").Value, Space(1), 10, "L")
    rstMaterialIOList.Fields("Date").Value = rstMaterialIOParent.Fields("Date").Value
    rstSourceList.MoveFirst
    rstSourceList.Find "[Code] = '" & rstMaterialIOParent.Fields("Source").Value & "'"
    rstMaterialIOList.Fields("SourceName").Value = Trim(rstSourceList.Fields("Col0").Value)
    rstMaterialIOList.Update
    rstMaterialIOList.Sort = SortOrder & " Asc"
    rstMaterialIOList.Find "[Code] = '" & rstMaterialIOParent.Fields("Code").Value & "'"
End Sub
Private Function CheckMandatoryFields() As Boolean
    If CheckEmpty(Text2.Text, False) Then
       DisplayError ("Order No. cannot be blank")
       Text2.SetFocus
       CheckMandatoryFields = True
    ElseIf CheckEmpty(Text3.Text, False) Then
       Text3.SetFocus
       CheckMandatoryFields = True
    ElseIf Not CheckExists(Text3, "Col0", rstSourceList, SourceCode) Then
        Text3.SetFocus
        CheckMandatoryFields = True
    ElseIf CheckDuplicate(CxnMaterialIssueOrder, "MaterialIOParent", "Code", "[Name]+[Type]", Trim(Text2.Text) & "0", rstMaterialIOParent.Fields("Code").Value, False) Then
        Text2.SetFocus
        CheckMandatoryFields = True
    ElseIf CheckItem() Then
       fpSpread1.SetFocus
        CheckMandatoryFields = True
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
    If SrchFor = "Source" Then
        rstMaterialIOList.Filter = "[SourceName] Like '%" & SrchText & "%'"
    End If
End Sub
Private Sub PrintMaterialIssueOrder()
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
Private Sub fpSpread1_EnterRow(ByVal Row As Long, ByVal RowIsLast As Long)
    T = 0
End Sub
Private Sub fpSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim ActiveCellVal As Variant, Category As Variant, Item As Variant, ref As Variant
    
    On Error Resume Next
    fpSpread1.GetText Col, Row, ActiveCellVal
    If ActiveCellVal = "" Then
        Cancel = True
        Exit Sub
    End If
    fpSpread1.GetText 1, Row, Category
    If Col = 1 Then
        T = 0
        fpSpread1.Col = 2
        fpSpread1.TypeComboBoxList = IIf(Category = "Repair Book", RepairBook, IIf(Category = "Fresh Book", FreshBook, Title))
    ElseIf Col = 2 Then
        T = 0
        If Category = "Repair Book" Then
           If rstRepairBookList.RecordCount > 0 Then rstRepairBookList.MoveFirst
           rstRepairBookList.Find "[Name]='" & FixQuote(ActiveCellVal) & "'"
           If Not rstRepairBookList.EOF Then
                fpSpread1.SetText 6, Row, rstRepairBookList.Fields("NCode").Value
           End If
        Else
           If rstFreshBookList.RecordCount > 0 Then rstFreshBookList.MoveFirst
           rstFreshBookList.Find "[Name]='" & FixQuote(ActiveCellVal) & "'"
           If Not rstFreshBookList.EOF Then
                fpSpread1.SetText 6, Row, rstFreshBookList.Fields("NCode").Value
           End If
        End If
    ElseIf Col = 3 Then
        T = T + 1
        If T = 1 Then
            fpSpread1.GetText 6, Row, Item
            Call LoadRefList(Category, Right(Item, 6), SourceCode, CheckNull(rstMaterialIOParent.Fields("Code").Value))
            fpSpread1.GetText 4, Row, ref
            Cancel = ShowRefList(ref)
            If Cancel Then T = 0
        End If
        If rstGodownList.RecordCount > 0 Then rstGodownList.MoveFirst
        rstGodownList.Find "[Col0]='" & FixQuote(ActiveCellVal) & "'"
        If Not rstGodownList.EOF Then
             fpSpread1.SetText 7, Row, rstGodownList.Fields("Code").Value
        End If
    ElseIf Col = 4 Then
        T = 0
    ElseIf Col = 5 Then
        T = 0
    End If
End Sub
Private Sub fpSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    EditMode = IIf(Mode = 1, True, False)
End Sub
Private Function CheckItem() As Boolean
    Dim i As Integer, Item As Variant, Category As Variant
    
    CheckItem = False
    For i = 1 To fpSpread1.DataRowCnt
        fpSpread1.SetActiveCell 1, i
        fpSpread1.GetText 6, i, Item
        fpSpread1.GetText 1, i, Category
        If Category = "Repair Book" Then
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
Private Sub LoadMaterialList(ByVal strOrderCode As String)
    Dim i As Integer
    On Error GoTo ErrorHandler
    
    If rstMaterialIOChild.State = adStateOpen Then
       rstMaterialIOChild.Close
    End If
    rstMaterialIOChild.Open "Select C.Category,C.Category+C.Item As ItemCode,(SELECT Name FROM BookMaster WHERE Code=C.Item) As ItemName,C.Godown,M.Name As GodownName,C.Ref,(SELECT Name FROM BookPOParent WHERE Code=C.Ref) AS RefNo,C.Quantity From MaterialIOChild C,AccountMaster M Where C.Godown = M.Code And C.Code = '" & strOrderCode & "' Order By Category", CxnMaterialIssueOrder, adOpenKeyset, adLockOptimistic
    rstMaterialIOChild.ActiveConnection = Nothing
    If rstMaterialIOChild.RecordCount > 0 Then rstMaterialIOChild.MoveFirst
    i = 0
    Do While Not rstMaterialIOChild.EOF
        i = i + 1
        With fpSpread1
            .SetText 1, i, IIf(rstMaterialIOChild.Fields("Category").Value = "3", "Fresh Book", IIf(rstMaterialIOChild.Fields("Category").Value = "4", "Repair Book", "Title"))
            .Col = 2
            .TypeComboBoxList = IIf(rstMaterialIOChild.Fields("Category").Value = "4", RepairBook, IIf(rstMaterialIOChild.Fields("Category").Value = "3", FreshBook, Title))
            .SetText 2, i, rstMaterialIOChild.Fields("ItemName").Value
            .SetText 3, i, rstMaterialIOChild.Fields("GodownName").Value
            .SetText 4, i, Trim(rstMaterialIOChild.Fields("RefNo").Value)
            .SetText 5, i, Val(rstMaterialIOChild.Fields("Quantity").Value)
            .SetText 6, i, rstMaterialIOChild.Fields("ItemCode").Value
            .SetText 7, i, rstMaterialIOChild.Fields("Godown").Value
            .SetText 8, i, rstMaterialIOChild.Fields("Ref").Value
        End With
        rstMaterialIOChild.MoveNext
    Loop
    Exit Sub
ErrorHandler:
    DisplayError ("Failed to Load Material List")
End Sub
Private Sub LoadRefList(ByVal strCategory As String, ByVal strItemCode As String, ByVal strSourceCode As String, ByVal strOrderCode As String)
    Dim BalanceQuantity As Long
    On Error GoTo ErrorHandler
    If rstRefList.State = adStateOpen Then rstRefList.Close
    If strCategory = "Title" Then
        rstRefList.Open "Select Name,ActualQuantity-(SELECT IIF(ISNULL(SUM(ActualQuantity)),0,SUM(ActualQuantity)) FROM BookPOChild08 WHERE Code=P.Code) As ReceivedQuantity,Format((Select Sum(Quantity) From MaterialIOChild Where Category='5' AND Ref=P.Code And Item=P.Book And Code<>'" & strOrderCode & "'),0) As IssuedQuantity,QuantityToBinder As BalanceQuantity,Remarks As Col0,P.Code From BookPOParent P Inner Join BookPOChild07 C On (P.Code=C.Code And Laminator='" & strSourceCode & "' And Book='" & strItemCode & "' And Type<>'O') Order BY Name", CxnMaterialIssueOrder, adOpenKeyset, adLockOptimistic
    Else
        rstRefList.Open "Select Name,Format(ActualQuantity,0) As ReceivedQuantity,Format((Select Sum(Quantity) From MaterialIOChild Where (Category='3' OR Category='4') AND Ref=P.Code And Item=P.Book And Code<>'" & strOrderCode & "'),0) As IssuedQuantity,ActualQuantity As BalanceQuantity,Remarks As Col0,P.Code From BookPOParent P Inner Join BookPOChild05 C On (P.Code=C.Code And BookPrinter='" & strSourceCode & "' And Book='" & strItemCode & "' And Type<>'O') Order BY Name", CxnMaterialIssueOrder, adOpenKeyset, adLockOptimistic
    End If
    rstRefList.ActiveConnection = Nothing
    If rstRefList.RecordCount > 0 Then rstRefList.MoveFirst
    Do While Not rstRefList.EOF
        BalanceQuantity = (Val(CheckNull(rstRefList.Fields("ReceivedQuantity").Value)) - Val(CheckNull(rstRefList.Fields("IssuedQuantity").Value))) - CalculateQuantityIssued(strCategory, strItemCode)
        If BalanceQuantity <> 0 Then
            rstRefList.Fields("Col0").Value = Trim(rstRefList.Fields("Name").Value) + " Quantity : " + Format(BalanceQuantity, "0")
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
Private Function CalculateQuantityIssued(ByVal strCategory As String, ByVal strItemCode As String) As Long
    Dim i As Integer, ref As Variant, Item As Variant, Category As Variant, Quantity As Variant

    For i = 1 To fpSpread1.DataRowCnt
        If fpSpread1.ActiveRow <> i Then
            fpSpread1.GetText 1, i, Category
            fpSpread1.GetText 4, i, ref
            fpSpread1.GetText 5, i, Quantity
            fpSpread1.GetText 6, i, Item
            If Trim(ref) = Trim(rstRefList.Fields("Name").Value) And Category = strCategory And Right(Item, 6) = strItemCode Then
                CalculateQuantityIssued = CalculateQuantityIssued + Val(Quantity)
            End If
        End If
    Next
End Function
Private Function ShowRefList(ByVal ref As String) As Boolean
    Dim SearchString As String, Qty As Variant
    
    Text9.Text = ref
    SearchString = FixQuote(Text9.Text)
    If rstRefList.RecordCount = 0 Then
        DisplayError ("No Pending Order")
        ShowRefList = True
        fpSpread1.SetFocus
        Exit Function
    Else
        rstRefList.MoveFirst
    End If
    rstRefList.Find "[Name] = '" & Pad(Trim(SearchString), Space(1), 10, "L") & "'"
    SelectionType = "S"
    RefCode = ""
    Call LoadSelectionList(rstRefList, "List of Pending Orders...", "Order No.")
    SearchOrder = 0
    Call DisplaySelectionList(Text9, RefCode)
    Call CloseForm(FrmSelectionList)
    If RefCode <> "" Then
        rstRefList.Find "[Code]='" & FixQuote(RefCode) & "'"
        With fpSpread1
            .SetText 4, .ActiveRow, rstRefList.Fields("Name").Value
            .SetText 8, .ActiveRow, RefCode
            .GetText 5, .ActiveRow, Qty
            .SetText 9, .ActiveRow, Val(Right(Trim(rstRefList.Fields("Col0").Value), InStr(1, StrReverse(Trim(rstRefList.Fields("Col0").Value)), ":") - 1))
            If Val(Qty) = 0 Then
                .SetText 5, .ActiveRow, Val(Right(Trim(rstRefList.Fields("Col0").Value), InStr(1, StrReverse(Trim(rstRefList.Fields("Col0").Value)), ":") - 1))
            End If
            .SetActiveCell 5, .ActiveRow
        End With
    Else
        ShowRefList = True
    End If
End Function
Private Function UpdateMaterialList(ByVal ActionType As String) As Boolean
    Dim CellVal(1 To 5) As Variant
    On Error GoTo ErrorHandler

    UpdateMaterialList = True
    If ActionType = "D" And (Not blnRecordExist) Then Exit Function
    If ActionType <> "I" Then
        CxnMaterialIssueOrder.Execute "Delete From MaterialIOChild WHERE Code = '" & rstMaterialIOParent.Fields("Code").Value & "'"
    Else
        With fpSpread1
            .GetText 1, .ActiveRow, CellVal(1)
            .GetText 5, .ActiveRow, CellVal(2)
            .GetText 6, .ActiveRow, CellVal(3)
            .GetText 7, .ActiveRow, CellVal(4)
            .GetText 8, .ActiveRow, CellVal(5)
        End With
        CxnMaterialIssueOrder.Execute "Insert Into MaterialIOChild Values ('" & rstMaterialIOParent.Fields("Code").Value & "','" & IIf(CellVal(1) = "Fresh Book", "3", IIf(CellVal(1) = "Repair Book", "4", "5")) & "','" & Right(CellVal(3), 6) & "','" & CellVal(4) & "','" & CellVal(5) & "'," & Val(CellVal(2)) & ")"
    End If
    Exit Function
ErrorHandler:
    UpdateMaterialList = False
End Function
Private Sub RefreshDropDownList(ByVal xType As String)
    If xType = "R" Then
        rstGodownList.ActiveConnection = CxnMaterialIssueOrder
        Do While Not RefreshRecord(rstGodownList)
        Loop
        rstGodownList.ActiveConnection = Nothing
        rstFreshBookList.ActiveConnection = CxnMaterialIssueOrder
        Do While Not RefreshRecord(rstFreshBookList)
        Loop
        rstFreshBookList.ActiveConnection = Nothing
        rstRepairBookList.ActiveConnection = CxnMaterialIssueOrder
        Do While Not RefreshRecord(rstRepairBookList)
        Loop
        rstRepairBookList.ActiveConnection = Nothing
        FreshBook = "": RepairBook = "": godown = "": Title = ""
    End If
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
    Do While Not rstGodownList.EOF
        If godown = "" Then
            godown = rstGodownList.Fields("Col0").Value
        Else
            godown = godown + Chr$(9) + rstGodownList.Fields("Col0").Value
        End If
        rstGodownList.MoveNext
    Loop
    fpSpread1.Col = 3
    fpSpread1.TypeComboBoxList = godown
End Sub
