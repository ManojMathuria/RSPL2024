VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmMaterialMovement 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Material Movement"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   330
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
   Icon            =   "MaterialMovement.frx":0000
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
      Picture         =   "MaterialMovement.frx":0442
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
         TabPicture(0)   =   "MaterialMovement.frx":045E
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "DataGrid1"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Text1"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "&Details"
         TabPicture(1)   =   "MaterialMovement.frx":047A
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Mh3dFrame2"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
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
            ColumnCount     =   4
            BeginProperty Column00 
               DataField       =   "Name"
               Caption         =   "Voucher No."
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
               Caption         =   "Vch Date"
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
               DataField       =   "GodownFromName"
               Caption         =   "Godown From"
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
               DataField       =   "GodownToName"
               Caption         =   "Godown To"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
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
                  ColumnWidth     =   1019.906
               EndProperty
               BeginProperty Column01 
                  Locked          =   -1  'True
                  ColumnWidth     =   1124.787
               EndProperty
               BeginProperty Column02 
                  Locked          =   -1  'True
                  ColumnWidth     =   2940.095
               EndProperty
               BeginProperty Column03 
                  Locked          =   -1  'True
                  ColumnWidth     =   2594.835
               EndProperty
            EndProperty
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   3790
            Left            =   -74880
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   480
            Width           =   8235
            _Version        =   65536
            _ExtentX        =   14526
            _ExtentY        =   6685
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
            Picture         =   "MaterialMovement.frx":0496
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
               Left            =   1440
               MaxLength       =   40
               TabIndex        =   3
               Top             =   950
               Width           =   6690
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
                  Weight          =   400
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
               TabIndex        =   4
               Top             =   1260
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
               Caption         =   " Voucher No."
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "MaterialMovement.frx":04B2
               Picture         =   "MaterialMovement.frx":04CE
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Left            =   5835
               TabIndex        =   13
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
               Caption         =   " Voucher Date"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "MaterialMovement.frx":04EA
               Picture         =   "MaterialMovement.frx":0506
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
               Caption         =   " Godown From"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "MaterialMovement.frx":0522
               Picture         =   "MaterialMovement.frx":053E
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
               Height          =   330
               Left            =   120
               TabIndex        =   15
               Top             =   1260
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
               Picture         =   "MaterialMovement.frx":055A
               Picture         =   "MaterialMovement.frx":0576
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel12 
               Height          =   330
               Left            =   120
               TabIndex        =   16
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
               Caption         =   " Godown To"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "MaterialMovement.frx":0592
               Picture         =   "MaterialMovement.frx":05AE
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
               Calendar        =   "MaterialMovement.frx":05CA
               Caption         =   "MaterialMovement.frx":06E2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "MaterialMovement.frx":074E
               Keys            =   "MaterialMovement.frx":076C
               Spin            =   "MaterialMovement.frx":07CA
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
               Height          =   1880
               Left            =   120
               TabIndex        =   17
               Top             =   1800
               Width           =   8010
               _Version        =   524288
               _ExtentX        =   14129
               _ExtentY        =   3316
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
               MaxCols         =   5
               MaxRows         =   100
               OperationMode   =   2
               SpreadDesigner  =   "MaterialMovement.frx":07F2
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
               Y1              =   1680
               Y2              =   1680
            End
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00808000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Find"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
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
Attribute VB_Name = "FrmMaterialMovement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CxnMaterialMovement As New ADODB.Connection
Dim rstCompanyMaster As New ADODB.Recordset
Dim rstMaterialMVList As New ADODB.Recordset
Dim rstMaterialMVParent As New ADODB.Recordset
Dim rstMaterialMVChild As New ADODB.Recordset
Dim rstAccountList As New ADODB.Recordset
Dim rstOutsourceItemList As New ADODB.Recordset
Dim rstFreshBookList As New ADODB.Recordset
Dim rstRepairBookList As New ADODB.Recordset
Dim AccountFromCode As String
Dim AccountToCode As String
Dim OutsourceItem As String
Dim FreshBook As String
Dim RepairBook As String
Dim Title As String
Dim EditMode As Boolean
Dim SortOrder As String
Dim PrevStr As String
Dim dblBookMark As Double
Dim blnRecordExist As Boolean
Dim OutputTo As String
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    CenterForm Me
    BusySystemIndicator True
    CxnMaterialMovement.CursorLocation = adUseClient
    CxnMaterialMovement.Open CxnDatabase.ConnectionString
    rstCompanyMaster.Open "Select PrintName, Address1, Address2, Address3, Address4, Phone, Fax, EMail, Website From CompanyMaster", CxnMaterialMovement, adOpenKeyset, adLockReadOnly
    rstAccountList.Open "Select Name As Col0, Code From AccountMaster Where Type In ('08','09') Order by Name", CxnMaterialMovement, adOpenKeyset, adLockReadOnly
    rstOutsourceItemList.Open "Select Name,'1'+Code As NCode From OutsourceItemMaster Order By Name", CxnMaterialMovement, adOpenKeyset, adLockOptimistic
    rstFreshBookList.Open "Select Name,Board,'3'+Code As NCode From BookMaster Where Type='F' Order By Name", CxnMaterialMovement, adOpenKeyset, adLockOptimistic
    rstRepairBookList.Open "Select Name,'4'+Code As NCode From BookMaster Where Type='R' Order By Name", CxnMaterialMovement, adOpenKeyset, adLockOptimistic
    rstMaterialMVList.Open "Select T.Code,T.Name,T.Date,(Select Name From AccountMaster Where Code=T.AccountFrom) As GodownFromName,(Select Name From AccountMaster Where Code=T.AccountTo) As GodownToName From MaterialMVParent T Order By T.Name", CxnMaterialMovement, adOpenKeyset, adLockOptimistic
    rstMaterialMVParent.CursorLocation = adUseClient
    rstMaterialMVList.Filter = adFilterNone
    If rstMaterialMVList.RecordCount > 0 Then rstMaterialMVList.MoveLast
    Set DataGrid1.DataSource = rstMaterialMVList
    BusySystemIndicator False
    SSTab1.Tab = 0
    SortOrder = "Name"
    If Not (rstMaterialMVList.EOF Or rstMaterialMVList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    rstMaterialMVList.ActiveConnection = Nothing
    rstAccountList.ActiveConnection = Nothing
    rstOutsourceItemList.ActiveConnection = Nothing
    rstFreshBookList.ActiveConnection = Nothing
    rstRepairBookList.ActiveConnection = Nothing
    Call RefreshDropDownList("A")
    fpSpread1.Col = 4
    fpSpread1.ColHidden = True
    fpSpread1.Col = 5
    fpSpread1.ColHidden = True
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
              Sendkeys "{TAB}"
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
    Call CloseRecordset(rstMaterialMVList)
    Call CloseRecordset(rstMaterialMVParent)
    Call CloseRecordset(rstMaterialMVChild)
    Call CloseRecordset(rstAccountList)
    Call CloseRecordset(rstOutsourceItemList)
    Call CloseRecordset(rstFreshBookList)
    Call CloseRecordset(rstRepairBookList)
    Call CloseConnection(CxnMaterialMovement)
    ShowProgressInStatusBar False
    DisableChildMenu
End Sub
Private Sub Text1_Change()
    If rstMaterialMVList.RecordCount = 0 Then Exit Sub
    rstMaterialMVList.MoveFirst
    If Text1.Text <> "" Then
        If SortOrder = "Name" Then
           rstMaterialMVList.Find "[" & SortOrder & "] Like '%" & FixQuote(Text1.Text) & "%'"
        Else
           rstMaterialMVList.Find "[" & SortOrder & "] Like '" & FixQuote(Text1.Text) & "%'"
        End If
        If rstMaterialMVList.EOF Then
            rstMaterialMVList.MoveFirst
            If PrevStr <> "" And Len(Text1.Text) > 1 Then
                If dblBookMark <> 0 Then
                    rstMaterialMVList.Bookmark = dblBookMark
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
    If Not (rstMaterialMVList.EOF Or rstMaterialMVList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim KeyProcessed As Boolean
    
    If rstMaterialMVList.RecordCount = 0 Then Exit Sub
    If Shift = 0 And KeyCode = vbKeyUp Then
        With rstMaterialMVList
            .MovePrevious
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyBack Then
        With rstMaterialMVList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyDown Then
        With rstMaterialMVList
            .MoveNext
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageUp Then
        With rstMaterialMVList
            .Move (-1) * (DataGrid1.VisibleRows - 1)
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageUp Then
        With rstMaterialMVList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageDown Then
        With rstMaterialMVList
            .Move DataGrid1.VisibleRows - 1
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageDown Then
        With rstMaterialMVList
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
            If Not (rstMaterialMVList.EOF Or rstMaterialMVList.BOF) Then
                With DataGrid1.SelBookmarks
                    If .Count <> 0 Then .Remove 0
                    .Add DataGrid1.Bookmark
                End With
            End If
            'Error Occurs On 12.09.14
            'Text1.SetFocus
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
        If rstMaterialMVParent.State = adStateOpen Then
           rstMaterialMVParent.Close
        End If
        rstMaterialMVParent.Open "Select * From MaterialMVParent Where Code = ''", CxnMaterialMovement, adOpenKeyset, adLockOptimistic
        ClearFields
        If AddRecord(rstMaterialMVParent) Then
            Text2.Text = GenerateCode(CxnMaterialMovement, "Select Max(Val(Name)) From MaterialMVParent", 10, Space(1))
            MhDateInput1.Text = Format(Date, "dd-MM-yyyy")
            Call SetButtons(False)
            SSTab1.Tab = 1
            Text2.SetFocus
            blnRecordExist = False
            CxnMaterialMovement.BeginTrans
        End If
    ElseIf Button.Index = 2 Then
        If rstMaterialMVList.RecordCount = 0 Then Exit Sub
        SSTab1.Tab = 1
        EditRecord
    ElseIf Button.Index = 3 Then
        If rstMaterialMVList.RecordCount = 0 Then Exit Sub
        If AllowTransactionsDeletion = 0 Then
            Call DisplayError("You don't have the rights to Delete this Voucher")
            Exit Sub
        End If
        SSTab1.Tab = 1
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            On Error Resume Next
            MdiMainMenu.MousePointer = vbHourglass
            CxnMaterialMovement.Execute "Delete From MaterialMVParent Where Code = '" & rstMaterialMVList.Fields("Code").Value & "'"
            MdiMainMenu.MousePointer = vbNormal
            If Err.Number = 0 Then
                rstMaterialMVList.Delete
                rstMaterialMVList.MoveNext
                If rstMaterialMVList.RecordCount > 0 And rstMaterialMVList.EOF Then
                    rstMaterialMVList.MoveLast
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
        If UpdateRecord(rstMaterialMVParent) Then
            If UpdateMaterialList("D") Then
                UpdateFlag = 1
                For i = 1 To fpSpread1.DataRowCnt
                    fpSpread1.SetActiveCell 3, i
                    fpSpread1.GetText 3, i, CellVal
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
            CxnMaterialMovement.CommitTrans
            If rstMaterialMVParent.State = adStateOpen Then
                rstMaterialMVParent.Close
            End If
            rstMaterialMVParent.CursorLocation = adUseClient
            Call SetButtons(True)
            SSTab1.Tab = 0
            ShowProgressInStatusBar True
            Timer1.Enabled = True
        Else
            DisplayError ("Failed to save the record")
            Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
        End If
    ElseIf Button.Index = 5 Then
        If CancelRecordUpdate(rstMaterialMVParent) Then
            CxnMaterialMovement.RollbackTrans
            If rstMaterialMVParent.State = adStateOpen Then
                rstMaterialMVParent.Close
            End If
            rstMaterialMVParent.CursorLocation = adUseClient
            Call SetButtons(True)
            SetButtonsForNoRecord
            SSTab1.Tab = 0
        End If
    ElseIf Button.Index = 6 Then
        SSTab1.Tab = 0
        Set DataGrid1.DataSource = Nothing
        rstMaterialMVList.ActiveConnection = CxnMaterialMovement
        Do While Not RefreshRecord(rstMaterialMVList)
        Loop
        Set DataGrid1.DataSource = rstMaterialMVList
        rstMaterialMVList.ActiveConnection = Nothing
        If rstMaterialMVList.RecordCount > 0 Then rstMaterialMVList.MoveLast
        rstAccountList.ActiveConnection = CxnMaterialMovement
        Do While Not RefreshRecord(rstAccountList)
        Loop
        rstAccountList.ActiveConnection = Nothing
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
        If rstMaterialMVList.RecordCount = 0 Then Exit Sub
        OutputTo = "P"
        PrintMaterialMovement
        HiLiteRecord = True
    ElseIf Button.Index = 10 Then
        If rstMaterialMVList.RecordCount = 0 Then Exit Sub
        OutputTo = "S"
        PrintMaterialMovement
        HiLiteRecord = True
    ElseIf Button.Index = 13 Then
        If rstMaterialMVList.RecordCount > 0 Then rstMaterialMVList.MoveFirst
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 14 Then
        If rstMaterialMVList.RecordCount > 0 Then
            rstMaterialMVList.MovePrevious
            If rstMaterialMVList.BOF Then
                rstMaterialMVList.MoveNext
            End If
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 15 Then
        If rstMaterialMVList.RecordCount > 0 Then
            rstMaterialMVList.MoveNext
            If rstMaterialMVList.EOF Then
                rstMaterialMVList.MovePrevious
            End If
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 16 Then
        If rstMaterialMVList.RecordCount > 0 Then rstMaterialMVList.MoveLast
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 18 Then
        Unload Me
        HiLiteRecord = False
    End If
    If HiLiteRecord Then
        If Not (rstMaterialMVList.EOF Or rstMaterialMVList.BOF) Then
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
          rstMaterialMVList.Sort = "Name Asc"
       End If
    ElseIf ColIndex = 2 Then
       If SortOrder <> "GodownFromName" Then
          SortOrder = "GodownFromName"
          rstMaterialMVList.Sort = "GodownFromName Asc"
       End If
    ElseIf ColIndex = 3 Then
       If SortOrder <> "GodownToName" Then
          SortOrder = "GodownToName"
          rstMaterialMVList.Sort = "GodownToName Asc"
       End If
    End If
    DataGrid1.ClearSelCols
    If Not (rstMaterialMVList.EOF Or rstMaterialMVList.BOF) Then
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
    If rstMaterialMVList.RecordCount = 0 Then
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
    If rstMaterialMVParent.EOF Or rstMaterialMVParent.BOF Then Exit Sub
    If CheckEmpty(Text2, True) Then
        Cancel = True
    ElseIf CheckDuplicate(CxnMaterialMovement, "MaterialMVParent", "Code", "[Name]", Trim(Text2.Text), rstMaterialMVParent.Fields("Code").Value, False) Then
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
        Sendkeys "{TAB}"
    End If
End Sub
Private Sub Text3_Validate(Cancel As Boolean)
    Dim SearchString As String

    SearchString = FixQuote(Text3.Text)
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
        AccountFromCode = ""
        Call LoadSelectionList(rstAccountList, "List of Godowns...", "Name")
        SearchOrder = 0
        Call DisplaySelectionList(Text3, AccountFromCode)
        Call CloseForm(FrmSelectionList)
        If CheckEmpty(Text3.Text, False) Then
            Text3.Text = "?"
        End If
        If RTrim(AccountFromCode) <> "" Then
            Sendkeys "{TAB}"
        End If
        Cancel = True
    Else
        AccountFromCode = rstAccountList.Fields("Code").Value
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
        AccountToCode = ""
        Call LoadSelectionList(rstAccountList, "List of Godowns...", "Name")
        SearchOrder = 0
        Call DisplaySelectionList(Text9, AccountToCode)
        Call CloseForm(FrmSelectionList)
        If CheckEmpty(Text9.Text, False) Then
            Text9.Text = "?"
        End If
        If RTrim(AccountToCode) <> "" Then
            Sendkeys "{TAB}"
        End If
        Cancel = True
        Exit Sub
    ElseIf Text3.Text = Text9.Text Then
        Call DisplayError("Both godowns cann't be same")
        Text9.SelStart = 0
        Text9.SelLength = Len(Text9.Text)
        Cancel = True
        Exit Sub
    End If
    AccountToCode = rstAccountList.Fields("Code").Value
End Sub
Private Sub ViewRecord()
    ClearFields
    If rstMaterialMVList.EOF Then
        If rstMaterialMVChild.State = adStateOpen Then
            rstMaterialMVChild.Close
        End If
        Exit Sub
    End If
    FindRecord
    LoadFields
End Sub
Private Sub FindRecord()
    If rstMaterialMVParent.State = adStateOpen Then
       rstMaterialMVParent.Close
    End If
    rstMaterialMVParent.Open "Select * From MaterialMVParent Where Code = '" & FixQuote(rstMaterialMVList.Fields("Code").Value) & "'", CxnMaterialMovement, adOpenKeyset, adLockOptimistic
    If rstMaterialMVParent.RecordCount = 0 Then
       Call DisplayError("This Record has been deleted by Another User ! Click Ok To Refresh the Recordset")
       Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
    End If
End Sub
Private Sub ClearFields()
    Text2.Text = ""
    Text3.Text = ""
    Text9.Text = ""
    Text4.Text = ""
    MhDateInput1.Text = Format(Date, "dd-MM-yyyy")
    fpSpread1.ClearRange 1, 1, fpSpread1.MaxCols, fpSpread1.MaxRows, True
End Sub
Private Sub LoadFields()
    If rstMaterialMVParent.EOF Or rstMaterialMVParent.BOF Then Exit Sub
    Text2.Text = rstMaterialMVParent.Fields("Name").Value
    MhDateInput1.Text = Format(rstMaterialMVParent.Fields("Date").Value, "dd-MM-yyyy")
    AccountFromCode = rstMaterialMVParent.Fields("AccountFrom").Value
    If rstAccountList.RecordCount > 0 Then rstAccountList.MoveFirst
    rstAccountList.Find "[Code] = '" & AccountFromCode & "'"
    If Not rstAccountList.EOF Then
       Text3.Text = rstAccountList.Fields("Col0").Value
    End If
    AccountToCode = rstMaterialMVParent.Fields("AccountTo").Value
    If rstAccountList.RecordCount > 0 Then rstAccountList.MoveFirst
    rstAccountList.Find "[Code] = '" & AccountToCode & "'"
    If Not rstAccountList.EOF Then
       Text9.Text = rstAccountList.Fields("Col0").Value
    End If
    Text4.Text = rstMaterialMVParent.Fields("Remarks").Value
    Call LoadMaterialList(rstMaterialMVParent.Fields("Code").Value)
End Sub
Private Sub EditRecord()
    On Error GoTo ErrorHandler
    
    If rstMaterialMVParent.RecordCount = 0 Then Exit Sub
    If rstMaterialMVChild.State = adStateClosed Then
        SSTab1.Tab = 0
        Exit Sub
    End If
    If rstMaterialMVParent.State = adStateOpen Then
       rstMaterialMVParent.Close
    End If
    rstMaterialMVParent.CursorLocation = adUseServer
    rstMaterialMVParent.Open "Select * From MaterialMVParent Where Code = '" & FixQuote(rstMaterialMVList.Fields("Code").Value) & "'", CxnMaterialMovement, adOpenKeyset, adLockPessimistic
    MdiMainMenu.MousePointer = vbHourglass
    rstMaterialMVParent.Fields("Printstatus") = "N"
    MdiMainMenu.MousePointer = vbNormal
    AddToList
    Call SetButtons(False)
    SSTab1.TabEnabled(0) = False
    Text2.SetFocus
    blnRecordExist = True
    CxnMaterialMovement.BeginTrans
    Exit Sub
ErrorHandler:
    If Err.Number = -2147467259 Then
       Call DisplayError("Failed to Edit the record")
    End If
    MdiMainMenu.MousePointer = vbNormal
    SSTab1.Tab = 0
End Sub
Private Sub SaveFields()
    If rstMaterialMVParent.EOF Or rstMaterialMVParent.BOF Then Exit Sub
    If Not blnRecordExist Then
        rstMaterialMVParent.Fields("Code").Value = GenerateCode(CxnMaterialMovement, "Select Max(Code) From MaterialMVParent", 6, "0")
        rstMaterialMVParent.Fields("CreatedBy").Value = UserCode
        rstMaterialMVParent.Fields("CreatedOn").Value = Now()
        rstMaterialMVParent.Fields("Recordstatus").Value = "N"
    Else
        rstMaterialMVParent.Fields("ModifiedBy").Value = UserCode
        rstMaterialMVParent.Fields("ModifiedOn").Value = Now()
        rstMaterialMVParent.Fields("Recordstatus").Value = "M"
    End If
    rstMaterialMVParent.Fields("Name").Value = Pad(Trim(Text2.Text), Space(1), 10, "L")
    rstMaterialMVParent.Fields("Date").Value = GetDate(MhDateInput1.Text)
    rstMaterialMVParent.Fields("AccountFrom").Value = AccountFromCode
    rstMaterialMVParent.Fields("AccountTo").Value = AccountToCode
    rstMaterialMVParent.Fields("Remarks").Value = Trim(Text4.Text)
    rstMaterialMVParent.Fields("PrintStatus").Value = "N"
End Sub
Private Sub AddToList()
    On Error Resume Next
    
    rstMaterialMVList.MoveFirst
    rstMaterialMVList.Find "[Code] = '" & rstMaterialMVParent.Fields("Code").Value & "'"
    If rstMaterialMVList.EOF Then
       rstMaterialMVList.AddNew
       rstMaterialMVList.Fields("Code").Value = rstMaterialMVParent.Fields("Code").Value
    End If
    rstMaterialMVList.Fields("Name").Value = Pad(rstMaterialMVParent.Fields("Name").Value, Space(1), 10, "L")
    rstMaterialMVList.Fields("Date").Value = rstMaterialMVParent.Fields("Date").Value
    rstAccountList.MoveFirst
    rstAccountList.Find "[Code] = '" & rstMaterialMVParent.Fields("AccountFrom").Value & "'"
    rstMaterialMVList.Fields("GodownFromName").Value = Trim(rstAccountList.Fields("Col0").Value)
    rstAccountList.MoveFirst
    rstAccountList.Find "[Code] = '" & rstMaterialMVParent.Fields("AccountTo").Value & "'"
    rstMaterialMVList.Fields("GodownToName").Value = Trim(rstAccountList.Fields("Col0").Value)
    rstMaterialMVList.Update
    rstMaterialMVList.Sort = SortOrder & " Asc"
    rstMaterialMVList.Find "[Code] = '" & rstMaterialMVParent.Fields("Code").Value & "'"
End Sub
Private Function CheckMandatoryFields() As Boolean
    If CheckEmpty(Text2.Text, False) Then
       DisplayError ("Order No. cannot be blank")
       Text2.SetFocus
       CheckMandatoryFields = True
    ElseIf Not CheckExists(Text3, "Col0", rstAccountList, AccountFromCode) Then
        Text3.SetFocus
        CheckMandatoryFields = True
    ElseIf CheckEmpty(Text9.Text, False) Then
       Text9.SetFocus
       CheckMandatoryFields = True
    ElseIf Not CheckExists(Text9, "Col0", rstAccountList, AccountToCode) Then
        Text9.SetFocus
        CheckMandatoryFields = True
    ElseIf Text3.Text = Text9.Text Then
       DisplayError ("Both godowns cann't be same")
       Text9.SetFocus
       CheckMandatoryFields = True
    ElseIf CheckDuplicate(CxnMaterialMovement, "MaterialMVParent", "Code", "[Name]", Trim(Text2.Text), rstMaterialMVParent.Fields("Code").Value, False) Then
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
        rstMaterialMVList.Filter = "[AccountName] Like '%" & SrchText & "%'"
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
Private Sub fpSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim ActiveCellVal As Variant, Category As Variant, Item As Variant, Qty As Variant, BalanceQuantity As Long
    
    On Error Resume Next
    fpSpread1.GetText Col, Row, ActiveCellVal
    If ActiveCellVal = "" Then
        Cancel = True
        Exit Sub
    End If
    fpSpread1.GetText 1, Row, Category
    If Col = 1 Then
        fpSpread1.Col = 2
        fpSpread1.TypeComboBoxList = IIf(Category = "Outsource Item", OutsourceItem, IIf(Category = "Repair Book", RepairBook, IIf(Category = "Fresh Book", FreshBook, Title)))
    ElseIf Col = 2 Then
        If Category = "Outsource Item" Then
           If rstOutsourceItemList.RecordCount > 0 Then rstOutsourceItemList.MoveFirst
           rstOutsourceItemList.Find "[Name]='" & FixQuote(ActiveCellVal) & "'"
           If Not rstOutsourceItemList.EOF Then
                fpSpread1.SetText 4, Row, rstOutsourceItemList.Fields("NCode").Value
           End If
        ElseIf Category = "Repair Book" Then
           If rstRepairBookList.RecordCount > 0 Then rstRepairBookList.MoveFirst
           rstRepairBookList.Find "[Name]='" & FixQuote(ActiveCellVal) & "'"
           If Not rstRepairBookList.EOF Then
                fpSpread1.SetText 4, Row, rstRepairBookList.Fields("NCode").Value
           End If
        Else
           If rstFreshBookList.RecordCount > 0 Then rstFreshBookList.MoveFirst
           rstFreshBookList.Find "[Name]='" & FixQuote(ActiveCellVal) & "'"
           If Not rstFreshBookList.EOF Then
                fpSpread1.SetText 4, Row, rstFreshBookList.Fields("NCode").Value
           End If
        End If
        With fpSpread1
            .GetText 4, Row, Item
            BalanceQuantity = CalculateMaterialBalance(AccountFromCode, Category, Right(Item, 6), CheckNull(rstMaterialMVParent.Fields("Code").Value), "MV")
            .SetText 5, .ActiveRow, Val(BalanceQuantity)
            .GetText 3, .ActiveRow, Qty
            If Val(Qty) = 0 Then
                .SetText 3, .ActiveRow, Val(BalanceQuantity)
            End If
        End With
    End If
End Sub
Private Sub fpSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    EditMode = IIf(Mode = 1, True, False)
End Sub
Private Function CheckItem() As Boolean
    Dim i As Integer, Item As Variant, Category As Variant, BalanceQuantity As Long, Qty As Variant
    
    CheckItem = False
    For i = 1 To fpSpread1.DataRowCnt
        fpSpread1.SetActiveCell 1, i
        fpSpread1.GetText 4, i, Item
        fpSpread1.GetText 1, i, Category
        If Category = "Outsource Item" Then
            If Left(Item, 1) <> "1" Then
                CheckItem = True
            End If
        ElseIf Category = "Repair Book" Then
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
        With fpSpread1
            BalanceQuantity = CalculateMaterialBalance(AccountFromCode, Category, Right(Item, 6), CheckNull(rstMaterialMVParent.Fields("Code").Value), "MV")
            .GetText 3, .ActiveRow, Qty
            If Val(Qty) > Val(BalanceQuantity) Then
                Call DisplayError("Quantity cann't be greater than " & Format(Val(BalanceQuantity), "0.000") & " in row #" & Trim(str(i)))
                CheckItem = True
                Exit For
            End If
        End With
    Next
End Function
Private Sub LoadMaterialList(ByVal strOrderCode As String)
    Dim i As Integer
    On Error GoTo ErrorHandler
    
    If rstMaterialMVChild.State = adStateOpen Then
       rstMaterialMVChild.Close
    End If
    rstMaterialMVChild.Open "Select C.Category,C.Category+C.Item As ItemCode,IIF(Category='1',(SELECT Name FROM OutsourceItemMaster WHERE Code=C.Item),(SELECT Name FROM BookMaster WHERE Code=C.Item)) As ItemName,C.Quantity From MaterialMVChild C Where C.Code = '" & strOrderCode & "' Order By Category", CxnMaterialMovement, adOpenKeyset, adLockOptimistic
    rstMaterialMVChild.ActiveConnection = Nothing
    If rstMaterialMVChild.RecordCount > 0 Then rstMaterialMVChild.MoveFirst
    i = 0
    Do While Not rstMaterialMVChild.EOF
        i = i + 1
        With fpSpread1
            .SetText 1, i, IIf(rstMaterialMVChild.Fields("Category").Value = "1", "Outsource Item", IIf(rstMaterialMVChild.Fields("Category").Value = "3", "Fresh Book", IIf(rstMaterialMVChild.Fields("Category").Value = "4", "Repair Book", "Title")))
            .Col = 2
            .TypeComboBoxList = IIf(rstMaterialMVChild.Fields("Category").Value = "1", OutsourceItem, IIf(rstMaterialMVChild.Fields("Category").Value = "4", RepairBook, IIf(rstMaterialMVChild.Fields("Category").Value = "3", FreshBook, Title)))
            .SetText 2, i, rstMaterialMVChild.Fields("ItemName").Value
            .SetText 3, i, Val(rstMaterialMVChild.Fields("Quantity").Value)
            .SetText 4, i, rstMaterialMVChild.Fields("ItemCode").Value
        End With
        rstMaterialMVChild.MoveNext
    Loop
    Exit Sub
ErrorHandler:
    DisplayError ("Failed to Load Material List")
End Sub
Private Function UpdateMaterialList(ByVal ActionType As String) As Boolean
    Dim CellVal(1 To 3) As Variant
    On Error GoTo ErrorHandler

    UpdateMaterialList = True
    If ActionType = "D" And (Not blnRecordExist) Then Exit Function
    If ActionType <> "I" Then
        CxnMaterialMovement.Execute "Delete From MaterialMVChild WHERE Code = '" & rstMaterialMVParent.Fields("Code").Value & "'"
    Else
        With fpSpread1
            .GetText 1, .ActiveRow, CellVal(1)
            .GetText 3, .ActiveRow, CellVal(2)
            .GetText 4, .ActiveRow, CellVal(3)
        End With
        CxnMaterialMovement.Execute "Insert Into MaterialMVChild Values ('" & rstMaterialMVParent.Fields("Code").Value & "','" & IIf(CellVal(1) = "Outsource Item", "1", IIf(CellVal(1) = "Fresh Book", "3", IIf(CellVal(1) = "Repair Book", "4", "5"))) & "','" & Right(CellVal(3), 6) & "'," & Val(CellVal(2)) & ")"
    End If
    Exit Function
ErrorHandler:
    UpdateMaterialList = False
End Function
Private Sub RefreshDropDownList(ByVal xType As String)
    If xType = "R" Then
        rstOutsourceItemList.ActiveConnection = CxnMaterialMovement
        Do While Not RefreshRecord(rstOutsourceItemList)
        Loop
        rstOutsourceItemList.ActiveConnection = Nothing
        rstFreshBookList.ActiveConnection = CxnMaterialMovement
        Do While Not RefreshRecord(rstFreshBookList)
        Loop
        rstFreshBookList.ActiveConnection = Nothing
        rstRepairBookList.ActiveConnection = CxnMaterialMovement
        Do While Not RefreshRecord(rstRepairBookList)
        Loop
        rstRepairBookList.ActiveConnection = Nothing
        OutsourceItem = "": FreshBook = "": RepairBook = "": Title = ""
    End If
    Do While Not rstOutsourceItemList.EOF
        If OutsourceItem = "" Then
            OutsourceItem = rstOutsourceItemList.Fields("Name").Value
        Else
            OutsourceItem = OutsourceItem + Chr$(9) + rstOutsourceItemList.Fields("Name").Value
        End If
        rstOutsourceItemList.MoveNext
    Loop
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
End Sub
Private Sub PrintMaterialMovement()
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    rptMaterialMovement.Text2.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptMaterialMovement.Text3.SetText Trim(rstCompanyMaster.Fields("Address1").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address2").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address3").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address4").Value)
    If (Not CheckEmpty(rstCompanyMaster.Fields("Phone").Value, False)) And (Not CheckEmpty(rstCompanyMaster.Fields("Fax").Value, False)) Then
        rptMaterialMovement.Text24.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & Space(1) & "Fax : " & Trim(rstCompanyMaster.Fields("Fax").Value)
    ElseIf Not CheckEmpty(rstCompanyMaster.Fields("Fax").Value, False) Then
        rptMaterialMovement.Text24.SetText "Fax : " & Trim(rstCompanyMaster.Fields("Fax").Value)
    ElseIf Not CheckEmpty(rstCompanyMaster.Fields("Phone").Value, False) Then
        rptMaterialMovement.Text24.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value)
    Else
        rptMaterialMovement.Section5.Suppress = True
    End If
    If rstMaterialMVChild.State = adStateOpen Then
        rstMaterialMVChild.Close
    End If
    rstMaterialMVChild.Open "Select Trim(P.Name) As VchNo,[Date] As VchDate,(Select Trim(PrintName) From AccountMaster Where Code = P.AccountFrom) As GodownFrom,(Select Trim(PrintName) From AccountMaster Where Code = P.AccountTo) As GodownTo,Category,IIF(Category='1',(SELECT Trim(PrintName) FROM OutsourceItemMaster WHERE Code=C.Item),(SELECT Trim(PrintName) FROM BookMaster WHERE Code=C.Item)) As MaterialName," & _
                                               "Quantity,Remarks From MaterialMVParent P Left Join MaterialMVChild C On (P.Code=C.Code And P.Code = '" & rstMaterialMVList.Fields("Code").Value & "') Order By Category", CxnMaterialMovement, adOpenKeyset, adLockOptimistic
    rstMaterialMVChild.Sort = "Category,MaterialName"
    rptMaterialMovement.Text27.SetText "for " & Trim(rstMaterialMVChild.Fields("GodownTo").Value)
    rptMaterialMovement.Text9.SetText "for " & Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptMaterialMovement.Database.SetDataSource rstMaterialMVChild, 3, 1
    Screen.MousePointer = vbNormal
    If OutputTo = "S" Then
        Set FrmReportViewer.Report = rptMaterialMovement
        FrmReportViewer.Show vbModal
    Else
        rptMaterialMovement.PrintOut
    End If
    Set rptMaterialMovement = Nothing
    On Error GoTo 0
End Sub
