VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmStockJournal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stock Journal"
   ClientHeight    =   7155
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
   Icon            =   "StockJournal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7155
   ScaleWidth      =   8715
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   7150
      Left            =   15
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   8700
      _Version        =   65536
      _ExtentX        =   15346
      _ExtentY        =   12612
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
      Picture         =   "StockJournal.frx":0442
      Begin TabDlg.SSTab SSTab1 
         Height          =   6910
         Left            =   120
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   120
         Width           =   8460
         _ExtentX        =   14923
         _ExtentY        =   12197
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
         TabPicture(0)   =   "StockJournal.frx":045E
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Text1"
         Tab(0).Control(1)=   "DataGrid1"
         Tab(0).Control(2)=   "Label1"
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "&Details"
         TabPicture(1)   =   "StockJournal.frx":047A
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
            Left            =   -74400
            TabIndex        =   10
            Top             =   6450
            Width           =   7760
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   5910
            Left            =   -74880
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   450
            Width           =   8235
            _ExtentX        =   14526
            _ExtentY        =   10425
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
               DataField       =   "AccountName"
               Caption         =   "Godown Name"
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
            Height          =   6070
            Left            =   120
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   480
            Width           =   8235
            _Version        =   65536
            _ExtentX        =   14526
            _ExtentY        =   10707
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
            Picture         =   "StockJournal.frx":0496
            Begin FPSpreadADO.fpSpread fpSpread1 
               Height          =   2145
               Left            =   120
               TabIndex        =   4
               Top             =   1470
               Width           =   8010
               _Version        =   524288
               _ExtentX        =   14129
               _ExtentY        =   3784
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
               MaxCols         =   4
               MaxRows         =   100
               OperationMode   =   2
               SpreadDesigner  =   "StockJournal.frx":04B2
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
               MaxLength       =   60
               TabIndex        =   2
               Top             =   630
               Width           =   6690
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
               Height          =   330
               Left            =   120
               TabIndex        =   13
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
               Picture         =   "StockJournal.frx":0A70
               Picture         =   "StockJournal.frx":0A8C
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Left            =   5955
               TabIndex        =   14
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
               Picture         =   "StockJournal.frx":0AA8
               Picture         =   "StockJournal.frx":0AC4
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   330
               Left            =   120
               TabIndex        =   15
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
               Caption         =   " Godown Name"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "StockJournal.frx":0AE0
               Picture         =   "StockJournal.frx":0AFC
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel11 
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
               Caption         =   " Remarks"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "StockJournal.frx":0B18
               Picture         =   "StockJournal.frx":0B34
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
               Calendar        =   "StockJournal.frx":0B50
               Caption         =   "StockJournal.frx":0C68
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "StockJournal.frx":0CD4
               Keys            =   "StockJournal.frx":0CF2
               Spin            =   "StockJournal.frx":0D50
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
               TabIndex        =   17
               Top             =   2520
               Width           =   2535
            End
            Begin FPSpreadADO.fpSpread fpSpread2 
               Height          =   2145
               Left            =   120
               TabIndex        =   5
               Top             =   3810
               Width           =   8010
               _Version        =   524288
               _ExtentX        =   14129
               _ExtentY        =   3784
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
               MaxCols         =   4
               MaxRows         =   100
               OperationMode   =   2
               SpreadDesigner  =   "StockJournal.frx":0D78
            End
            Begin VB.Line Line3 
               X1              =   0
               X2              =   8280
               Y1              =   3705
               Y2              =   3705
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
            Left            =   -74880
            TabIndex        =   11
            Top             =   6450
            Width           =   495
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   7
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
Attribute VB_Name = "FrmStockJournal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CxnMaterialStockAdjustment As New ADODB.Connection
Dim rstCompanyMaster As New ADODB.Recordset
Dim rstMaterialSVList As New ADODB.Recordset
Dim rstMaterialSVParent As New ADODB.Recordset
Dim rstMaterialSVChild As ADODB.Recordset
Dim rstAccountList As New ADODB.Recordset
Dim rstOutsourceItemList As New ADODB.Recordset
Dim rstPaperList As New ADODB.Recordset
Dim rstFreshBookList As New ADODB.Recordset
Dim rstRepairBookList As New ADODB.Recordset
Dim AccountCode As String
Dim OutsourceItem As String
Dim Paper As String
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
    CxnMaterialStockAdjustment.CursorLocation = adUseClient
    CxnMaterialStockAdjustment.Open CxnDatabase.ConnectionString
    rstCompanyMaster.Open "Select PrintName, Address1, Address2, Address3, Address4, Phone, Fax, EMail, Website From CompanyMaster", CxnMaterialStockAdjustment, adOpenKeyset, adLockReadOnly
    rstAccountList.Open "SELECT TRIM(Name)+' ('+CHOOSE(VAL(Type)-4,'Book Printer','Title Printer','','Binder','Godown')+')'  As Col0,Code FROM AccountMaster WHERE Type IN ('05','06','08','09') ORDER BY Name", CxnMaterialStockAdjustment, adOpenKeyset, adLockReadOnly
    rstOutsourceItemList.Open "Select Name,'1'+Code As NCode From OutsourceItemMaster Order By Name", CxnMaterialStockAdjustment, adOpenKeyset, adLockOptimistic
    rstPaperList.Open "Select Name,'2'+Code As NCode From PaperMaster Order By Name", CxnMaterialStockAdjustment, adOpenKeyset, adLockOptimistic
    rstFreshBookList.Open "Select Name,Board,'3'+Code As NCode From BookMaster Where Type='F' Order By Name", CxnMaterialStockAdjustment, adOpenKeyset, adLockOptimistic
    rstRepairBookList.Open "Select Name,'4'+Code As NCode From BookMaster Where Type='R' Order By Name", CxnMaterialStockAdjustment, adOpenKeyset, adLockOptimistic
    rstMaterialSVList.Open "Select T.Code,T.Name,T.Date,M.Name As AccountName From MaterialSVParent T, AccountMaster M Where T.Account = M.Code Order By T.Name", CxnMaterialStockAdjustment, adOpenKeyset, adLockOptimistic
    rstMaterialSVParent.CursorLocation = adUseClient
    Set rstMaterialSVChild = New ADODB.Recordset
    rstMaterialSVList.Filter = adFilterNone
    If rstMaterialSVList.RecordCount > 0 Then rstMaterialSVList.MoveLast
    Set DataGrid1.DataSource = rstMaterialSVList
    BusySystemIndicator False
    SSTab1.Tab = 0
    SortOrder = "Name"
    If Not (rstMaterialSVList.EOF Or rstMaterialSVList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    rstMaterialSVList.ActiveConnection = Nothing
    rstAccountList.ActiveConnection = Nothing
    rstOutsourceItemList.ActiveConnection = Nothing
    rstPaperList.ActiveConnection = Nothing
    rstFreshBookList.ActiveConnection = Nothing
    rstRepairBookList.ActiveConnection = Nothing
    Call RefreshDropDownList("A")
    fpSpread1.Col = 4
    fpSpread1.ColHidden = True
    fpSpread2.Col = 4
    fpSpread2.ColHidden = True
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
           If Me.ActiveControl.Name <> "fpSpread1" And Me.ActiveControl.Name <> "fpSpread2" Then
              SendKeys "{TAB}"
           End If
        End If
        If Me.ActiveControl.Name <> "fpSpread1" And Me.ActiveControl.Name <> "fpSpread2" Then
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
    Call CloseRecordset(rstMaterialSVList)
    Call CloseRecordset(rstMaterialSVParent)
    Call CloseRecordset(rstMaterialSVChild)
    Call CloseRecordset(rstAccountList)
    Call CloseRecordset(rstOutsourceItemList)
    Call CloseRecordset(rstPaperList)
    Call CloseRecordset(rstFreshBookList)
    Call CloseRecordset(rstRepairBookList)
    Call CloseConnection(CxnMaterialStockAdjustment)
    ShowProgressInStatusBar False
    DisableChildMenu
End Sub
Private Sub Text1_Change()
    If rstMaterialSVList.RecordCount = 0 Then Exit Sub
    rstMaterialSVList.MoveFirst
    If Text1.Text <> "" Then
        If SortOrder = "Name" Then
           rstMaterialSVList.Find "[" & SortOrder & "] Like '%" & FixQuote(Text1.Text) & "%'"
        Else
           rstMaterialSVList.Find "[" & SortOrder & "] Like '" & FixQuote(Text1.Text) & "%'"
        End If
        If rstMaterialSVList.EOF Then
            rstMaterialSVList.MoveFirst
            If PrevStr <> "" And Len(Text1.Text) > 1 Then
                If dblBookMark <> 0 Then
                    rstMaterialSVList.Bookmark = dblBookMark
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
    If Not (rstMaterialSVList.EOF Or rstMaterialSVList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim KeyProcessed As Boolean
    
    If rstMaterialSVList.RecordCount = 0 Then Exit Sub
    If Shift = 0 And KeyCode = vbKeyUp Then
        With rstMaterialSVList
            .MovePrevious
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyBack Then
        With rstMaterialSVList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyDown Then
        With rstMaterialSVList
            .MoveNext
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageUp Then
        With rstMaterialSVList
            .Move (-1) * (DataGrid1.VisibleRows - 1)
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageUp Then
        With rstMaterialSVList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageDown Then
        With rstMaterialSVList
            .Move DataGrid1.VisibleRows - 1
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageDown Then
        With rstMaterialSVList
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
            If Not (rstMaterialSVList.EOF Or rstMaterialSVList.BOF) Then
                With DataGrid1.SelBookmarks
                    If .Count <> 0 Then .Remove 0
                    .Add DataGrid1.Bookmark
                End With
            End If
            'Error Occurs on 12.09.14
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
        If rstMaterialSVParent.State = adStateOpen Then
           rstMaterialSVParent.Close
        End If
        rstMaterialSVParent.Open "Select * From MaterialSVParent Where Code = ''", CxnMaterialStockAdjustment, adOpenKeyset, adLockOptimistic
        ClearFields
        If AddRecord(rstMaterialSVParent) Then
            Text2.Text = GenerateCode(CxnMaterialStockAdjustment, "Select Max(Val(Name)) From MaterialSVParent", 10, Space(1))
            MhDateInput1.Text = Format(Date, "dd-MM-yyyy")
            Call SetButtons(False)
            SSTab1.Tab = 1
            Text2.SetFocus
            blnRecordExist = False
            CxnMaterialStockAdjustment.BeginTrans
        End If
    ElseIf Button.Index = 2 Then
        If rstMaterialSVList.RecordCount = 0 Then Exit Sub
        SSTab1.Tab = 1
        EditRecord
    ElseIf Button.Index = 3 Then
        If rstMaterialSVList.RecordCount = 0 Then Exit Sub
        If AllowTransactionsDeletion = 0 Then
            Call DisplayError("You don't have the rights to Delete this Voucher")
            Exit Sub
        End If
        SSTab1.Tab = 1
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            On Error Resume Next
            MdiMainMenu.MousePointer = vbHourglass
            CxnMaterialStockAdjustment.Execute "Delete From MaterialSVParent Where Code = '" & rstMaterialSVList.Fields("Code").Value & "'"
            MdiMainMenu.MousePointer = vbNormal
            If Err.Number = 0 Then
                rstMaterialSVList.Delete
                rstMaterialSVList.MoveNext
                If rstMaterialSVList.RecordCount > 0 And rstMaterialSVList.EOF Then
                    rstMaterialSVList.MoveLast
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
        If UpdateRecord(rstMaterialSVParent) Then
            If UpdateMaterialList("D") Then
                UpdateFlag = 1
                For i = 1 To fpSpread1.DataRowCnt
                    fpSpread1.SetActiveCell 3, i
                    fpSpread1.GetText 3, i, CellVal
                    If Val(CellVal) <> 0 Then
                        If Not UpdateMaterialList("I1") Then
                            UpdateFlag = 0
                            Exit For
                        End If
                    End If
                Next
                If UpdateFlag = 1 Then
                    For i = 1 To fpSpread2.DataRowCnt
                        fpSpread2.SetActiveCell 3, i
                        fpSpread2.GetText 3, i, CellVal
                        If Val(CellVal) <> 0 Then
                            If Not UpdateMaterialList("I2") Then
                                UpdateFlag = 0
                                Exit For
                            End If
                        End If
                    Next
                End If
            End If
        End If
        If UpdateFlag Then
            AddToList
            CxnMaterialStockAdjustment.CommitTrans
            If rstMaterialSVParent.State = adStateOpen Then
                rstMaterialSVParent.Close
            End If
            rstMaterialSVParent.CursorLocation = adUseClient
            Call SetButtons(True)
            SSTab1.Tab = 0
            ShowProgressInStatusBar True
            Timer1.Enabled = True
        Else
            DisplayError ("Failed to save the record")
            Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
        End If
    ElseIf Button.Index = 5 Then
        If CancelRecordUpdate(rstMaterialSVParent) Then
            CxnMaterialStockAdjustment.RollbackTrans
            If rstMaterialSVParent.State = adStateOpen Then
                rstMaterialSVParent.Close
            End If
            rstMaterialSVParent.CursorLocation = adUseClient
            Call SetButtons(True)
            SetButtonsForNoRecord
            SSTab1.Tab = 0
        End If
    ElseIf Button.Index = 6 Then
        SSTab1.Tab = 0
        Set DataGrid1.DataSource = Nothing
        rstMaterialSVList.ActiveConnection = CxnMaterialStockAdjustment
        Do While Not RefreshRecord(rstMaterialSVList)
        Loop
        Set DataGrid1.DataSource = rstMaterialSVList
        rstMaterialSVList.ActiveConnection = Nothing
        If rstMaterialSVList.RecordCount > 0 Then rstMaterialSVList.MoveLast
        rstAccountList.ActiveConnection = CxnMaterialStockAdjustment
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
        If rstMaterialSVList.RecordCount = 0 Then Exit Sub
        OutputTo = "P"
        PrintMaterialStockAdjustment
        HiLiteRecord = True
    ElseIf Button.Index = 10 Then
        If rstMaterialSVList.RecordCount = 0 Then Exit Sub
        OutputTo = "S"
        PrintMaterialStockAdjustment
        HiLiteRecord = True
    ElseIf Button.Index = 13 Then
        If rstMaterialSVList.RecordCount > 0 Then rstMaterialSVList.MoveFirst
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 14 Then
        If rstMaterialSVList.RecordCount > 0 Then
            rstMaterialSVList.MovePrevious
            If rstMaterialSVList.BOF Then
                rstMaterialSVList.MoveNext
            End If
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 15 Then
        If rstMaterialSVList.RecordCount > 0 Then
            rstMaterialSVList.MoveNext
            If rstMaterialSVList.EOF Then
                rstMaterialSVList.MovePrevious
            End If
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 16 Then
        If rstMaterialSVList.RecordCount > 0 Then rstMaterialSVList.MoveLast
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 18 Then
        Unload Me
        HiLiteRecord = False
    End If
    If HiLiteRecord Then
        If Not (rstMaterialSVList.EOF Or rstMaterialSVList.BOF) Then
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
          rstMaterialSVList.Sort = "Name Asc"
       End If
    ElseIf ColIndex = 2 Then
       If SortOrder <> "AccountName" Then
          SortOrder = "AccountName"
          rstMaterialSVList.Sort = "AccountName Asc"
       End If
    End If
    DataGrid1.ClearSelCols
    If Not (rstMaterialSVList.EOF Or rstMaterialSVList.BOF) Then
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
    If rstMaterialSVList.RecordCount = 0 Then
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
    If rstMaterialSVParent.EOF Or rstMaterialSVParent.BOF Then Exit Sub
    If CheckEmpty(Text2, True) Then
        Cancel = True
    ElseIf CheckDuplicate(CxnMaterialStockAdjustment, "MaterialSVParent", "Code", "[Name]", Trim(Text2.Text), rstMaterialSVParent.Fields("Code").Value, False) Then
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
    If rstAccountList.RecordCount = 0 Then
        DisplayError ("No Record in Source Master")
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
        Call DisplaySelectionList(Text3, AccountCode)
        Call CloseForm(FrmSelectionList)
        If CheckEmpty(Text3.Text, False) Then
            Text3.Text = "?"
        End If
        If RTrim(AccountCode) <> "" Then
            SendKeys "{TAB}"
        End If
        Cancel = True
    Else
        AccountCode = rstAccountList.Fields("Code").Value
    End If
End Sub
Private Sub ViewRecord()
    ClearFields
    If rstMaterialSVList.EOF Then
        If rstMaterialSVChild.State = adStateOpen Then
            rstMaterialSVChild.Close
        End If
        Exit Sub
    End If
    FindRecord
    LoadFields
End Sub
Private Sub FindRecord()
    If rstMaterialSVParent.State = adStateOpen Then
       rstMaterialSVParent.Close
    End If
    rstMaterialSVParent.Open "Select * From MaterialSVParent Where Code = '" & FixQuote(rstMaterialSVList.Fields("Code").Value) & "'", CxnMaterialStockAdjustment, adOpenKeyset, adLockOptimistic
    If rstMaterialSVParent.RecordCount = 0 Then
       Call DisplayError("This Record has been deleted by Another User ! Click Ok To Refresh the Recordset")
       Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
    End If
End Sub
Private Sub ClearFields()
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    MhDateInput1.Text = Format(Date, "dd-MM-yyyy")
    fpSpread1.ClearRange 1, 1, fpSpread1.MaxCols, fpSpread1.MaxRows, True
    fpSpread2.ClearRange 1, 1, fpSpread2.MaxCols, fpSpread2.MaxRows, True
End Sub
Private Sub LoadFields()
    If rstMaterialSVParent.EOF Or rstMaterialSVParent.BOF Then Exit Sub
    Text2.Text = rstMaterialSVParent.Fields("Name").Value
    MhDateInput1.Text = Format(rstMaterialSVParent.Fields("Date").Value, "dd-MM-yyyy")
    AccountCode = rstMaterialSVParent.Fields("Account").Value
    If rstAccountList.RecordCount > 0 Then rstAccountList.MoveFirst
    rstAccountList.Find "[Code] = '" & AccountCode & "'"
    If Not rstAccountList.EOF Then
       Text3.Text = rstAccountList.Fields("Col0").Value
    End If
    Text4.Text = rstMaterialSVParent.Fields("Remarks").Value
    Call LoadMaterialList(rstMaterialSVParent.Fields("Code").Value)
End Sub
Private Sub EditRecord()
    On Error GoTo ErrorHandler
    
    If rstMaterialSVParent.RecordCount = 0 Then Exit Sub
    If rstMaterialSVChild.State = adStateClosed Then
        SSTab1.Tab = 0
        Exit Sub
    End If
    If rstMaterialSVParent.State = adStateOpen Then
       rstMaterialSVParent.Close
    End If
    rstMaterialSVParent.CursorLocation = adUseServer
    rstMaterialSVParent.Open "Select * From MaterialSVParent Where Code = '" & FixQuote(rstMaterialSVList.Fields("Code").Value) & "'", CxnMaterialStockAdjustment, adOpenKeyset, adLockPessimistic
    MdiMainMenu.MousePointer = vbHourglass
    rstMaterialSVParent.Fields("Printstatus") = "N"
    MdiMainMenu.MousePointer = vbNormal
    AddToList
    Call SetButtons(False)
    SSTab1.TabEnabled(0) = False
    Text2.SetFocus
    blnRecordExist = True
    CxnMaterialStockAdjustment.BeginTrans
    Exit Sub
ErrorHandler:
    If Err.Number = -2147467259 Then
       Call DisplayError("Failed to Edit the record")
    End If
    MdiMainMenu.MousePointer = vbNormal
    SSTab1.Tab = 0
End Sub
Private Sub SaveFields()
    If rstMaterialSVParent.EOF Or rstMaterialSVParent.BOF Then Exit Sub
    If Not blnRecordExist Then
        rstMaterialSVParent.Fields("Code").Value = GenerateCode(CxnMaterialStockAdjustment, "Select Max(Code) From MaterialSVParent", 6, "0")
        rstMaterialSVParent.Fields("CreatedBy").Value = UserCode
        rstMaterialSVParent.Fields("CreatedOn").Value = Now()
        rstMaterialSVParent.Fields("Recordstatus").Value = "N"
    Else
        rstMaterialSVParent.Fields("ModifiedBy").Value = UserCode
        rstMaterialSVParent.Fields("ModifiedOn").Value = Now()
        rstMaterialSVParent.Fields("Recordstatus").Value = "M"
    End If
    rstMaterialSVParent.Fields("Name").Value = Pad(Trim(Text2.Text), Space(1), 10, "L")
    rstMaterialSVParent.Fields("Date").Value = GetDate(MhDateInput1.Text)
    rstMaterialSVParent.Fields("Account").Value = AccountCode
    rstMaterialSVParent.Fields("Remarks").Value = Trim(Text4.Text)
    rstMaterialSVParent.Fields("PrintStatus").Value = "N"
End Sub
Private Sub AddToList()
    On Error Resume Next
    
    rstMaterialSVList.MoveFirst
    rstMaterialSVList.Find "[Code] = '" & rstMaterialSVParent.Fields("Code").Value & "'"
    If rstMaterialSVList.EOF Then
       rstMaterialSVList.AddNew
       rstMaterialSVList.Fields("Code").Value = rstMaterialSVParent.Fields("Code").Value
    End If
    rstMaterialSVList.Fields("Name").Value = Pad(rstMaterialSVParent.Fields("Name").Value, Space(1), 10, "L")
    rstMaterialSVList.Fields("Date").Value = rstMaterialSVParent.Fields("Date").Value
    rstAccountList.MoveFirst
    rstAccountList.Find "[Code] = '" & rstMaterialSVParent.Fields("Account").Value & "'"
    rstMaterialSVList.Fields("AccountName").Value = Trim(rstAccountList.Fields("Col0").Value)
    rstMaterialSVList.Update
    rstMaterialSVList.Sort = SortOrder & " Asc"
    rstMaterialSVList.Find "[Code] = '" & rstMaterialSVParent.Fields("Code").Value & "'"
End Sub
Private Function CheckMandatoryFields() As Boolean
    If CheckEmpty(Text2.Text, False) Then
       DisplayError ("Order No. cannot be blank")
       Text2.SetFocus
       CheckMandatoryFields = True
    ElseIf CheckEmpty(Text3.Text, False) Then
       Text3.SetFocus
       CheckMandatoryFields = True
    ElseIf Not CheckExists(Text3, "Col0", rstAccountList, AccountCode) Then
        Text3.SetFocus
        CheckMandatoryFields = True
    ElseIf CheckDuplicate(CxnMaterialStockAdjustment, "MaterialSVParent", "Code", "[Name]", Trim(Text2.Text), rstMaterialSVParent.Fields("Code").Value, False) Then
        Text2.SetFocus
        CheckMandatoryFields = True
    ElseIf CheckItem("1") Then
       fpSpread1.SetFocus
        CheckMandatoryFields = True
    ElseIf CheckItem("2") Then
       fpSpread2.SetFocus
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
        rstMaterialSVList.Filter = "[AccountName] Like '%" & SrchText & "%'"
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
Private Sub fpSpread2_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask And KeyCode = vbKeyD Then
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            fpSpread2.DeleteRows fpSpread2.ActiveRow, 1
            fpSpread2.SetFocus
        End If
    ElseIf Shift = 0 And KeyCode = vbKeyF5 Then
        Call RefreshDropDownList("R")
    End If
End Sub
Private Sub fpSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim ActiveCellVal As Variant, Category As Variant, Item As Variant
    
    fpSpread1.GetText Col, Row, ActiveCellVal
    If ActiveCellVal = "" Then
        Cancel = True
        Exit Sub
    End If
    fpSpread1.GetText 1, Row, Category
    If Col = 1 Then
        fpSpread1.Col = 2
        fpSpread1.TypeComboBoxList = IIf(Category = "Outsource Item", OutsourceItem, IIf(Category = "Paper", Paper, IIf(Category = "Repair Book", RepairBook, IIf(Category = "Fresh Book", FreshBook, Title))))
    ElseIf Col = 2 Then
        If Category = "Outsource Item" Then
           If rstOutsourceItemList.RecordCount > 0 Then rstOutsourceItemList.MoveFirst
           rstOutsourceItemList.Find "[Name]='" & FixQuote(ActiveCellVal) & "'"
           If Not rstOutsourceItemList.EOF Then
                fpSpread1.SetText 4, Row, rstOutsourceItemList.Fields("NCode").Value
           End If
        ElseIf Category = "Paper" Then
           If rstPaperList.RecordCount > 0 Then rstPaperList.MoveFirst
           rstPaperList.Find "[Name]='" & FixQuote(ActiveCellVal) & "'"
           If Not rstPaperList.EOF Then
                fpSpread1.SetText 4, Row, rstPaperList.Fields("NCode").Value
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
    End If
End Sub
Private Sub fpSpread2_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim ActiveCellVal As Variant, Category As Variant, Item As Variant
    
    fpSpread2.GetText Col, Row, ActiveCellVal
    If ActiveCellVal = "" Then
        Cancel = True
        Exit Sub
    End If
    fpSpread2.GetText 1, Row, Category
    If Col = 1 Then
        fpSpread2.Col = 2
        fpSpread2.TypeComboBoxList = IIf(Category = "Outsource Item", OutsourceItem, IIf(Category = "Paper", Paper, IIf(Category = "Repair Book", RepairBook, IIf(Category = "Fresh Book", FreshBook, Title))))
    ElseIf Col = 2 Then
        If Category = "Outsource Item" Then
           If rstOutsourceItemList.RecordCount > 0 Then rstOutsourceItemList.MoveFirst
           rstOutsourceItemList.Find "[Name]='" & FixQuote(ActiveCellVal) & "'"
           If Not rstOutsourceItemList.EOF Then
                fpSpread2.SetText 4, Row, rstOutsourceItemList.Fields("NCode").Value
           End If
        ElseIf Category = "Paper" Then
           If rstPaperList.RecordCount > 0 Then rstPaperList.MoveFirst
           rstPaperList.Find "[Name]='" & FixQuote(ActiveCellVal) & "'"
           If Not rstPaperList.EOF Then
                fpSpread2.SetText 4, Row, rstPaperList.Fields("NCode").Value
           End If
        ElseIf Category = "Repair Book" Then
           If rstRepairBookList.RecordCount > 0 Then rstRepairBookList.MoveFirst
           rstRepairBookList.Find "[Name]='" & FixQuote(ActiveCellVal) & "'"
           If Not rstRepairBookList.EOF Then
                fpSpread2.SetText 4, Row, rstRepairBookList.Fields("NCode").Value
           End If
        Else
           If rstFreshBookList.RecordCount > 0 Then rstFreshBookList.MoveFirst
           rstFreshBookList.Find "[Name]='" & FixQuote(ActiveCellVal) & "'"
           If Not rstFreshBookList.EOF Then
                fpSpread2.SetText 4, Row, rstFreshBookList.Fields("NCode").Value
           End If
        End If
    End If
End Sub
Private Sub fpSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    EditMode = IIf(Mode = 1, True, False)
End Sub
Private Sub fpSpread2_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    EditMode = IIf(Mode = 1, True, False)
End Sub
Private Function CheckItem(ByVal xNumber As String) As Boolean
    Dim i As Integer, K As Integer, Item01 As Variant, Category01 As Variant, Item02 As Variant, Category02 As Variant
    
    CheckItem = False
    If xNumber = "1" Then
        For i = 1 To fpSpread1.DataRowCnt
            fpSpread1.SetActiveCell 1, i
            fpSpread1.GetText 4, i, Item01
            fpSpread1.GetText 1, i, Category01
            If Category01 = "Outsource Item" Then
                If Left(Item01, 1) <> "1" Then
                    CheckItem = True
                End If
            ElseIf Category01 = "Paper" Then
                If Left(Item01, 1) <> "2" Then
                    CheckItem = True
                End If
            ElseIf Category01 = "Repair Book" Then
                If Left(Item01, 1) <> "4" Then
                    CheckItem = True
                End If
            Else
                If Left(Item01, 1) <> "3" And Left(Item01, 1) <> "5" Then
                    CheckItem = True
                End If
            End If
            If CheckItem Then
                DisplayError "Data mismatch in row #" & Trim(str(i))
                Exit For
            End If
        Next
    Else
        For i = 1 To fpSpread2.DataRowCnt
            fpSpread2.SetActiveCell 1, i
            fpSpread2.GetText 4, i, Item01
            fpSpread2.GetText 1, i, Category01
            If Category01 = "Outsource Item" Then
                If Left(Item01, 1) <> "1" Then
                    CheckItem = True
                End If
            ElseIf Category01 = "Paper" Then
                If Left(Item01, 1) <> "2" Then
                    CheckItem = True
                End If
            ElseIf Category01 = "Repair Book" Then
                If Left(Item01, 1) <> "4" Then
                    CheckItem = True
                End If
            Else
                If Left(Item01, 1) <> "3" And Left(Item01, 1) <> "5" Then
                    CheckItem = True
                End If
            End If
            If CheckItem Then
                DisplayError "Data mismatch in row #" & Trim(str(i))
                Exit For
            End If
        Next
    End If
    If Not CheckItem Then
        For i = 1 To fpSpread1.DataRowCnt
            fpSpread1.SetActiveCell 1, i
            fpSpread1.GetText 1, i, Category01
            fpSpread1.GetText 4, i, Item01
            For K = 1 To fpSpread2.DataRowCnt
                fpSpread2.SetActiveCell 1, K
                fpSpread2.GetText 1, K, Category02
                fpSpread2.GetText 4, K, Item02
                If Category02 = Category01 And Item02 = Item01 Then
                    CheckItem = True
                    Exit For
                End If
            Next
            If CheckItem Then
                DisplayError "Same item cann't be generated (row #" & Trim(str(i)) & ") and consumed (row #" & Trim(str(K)) & ") simultaneously"
                Exit For
            End If
        Next
    End If
End Function
Private Sub LoadMaterialList(ByVal strOrderCode As String)
    Dim i As Integer
    On Error GoTo ErrorHandler
    
    If rstMaterialSVChild.State = adStateOpen Then
       rstMaterialSVChild.Close
    End If
    rstMaterialSVChild.Open "Select C.Category,C.Category+C.Item As ItemCode,IIF(Category='1',(SELECT Name FROM OutsourceItemMaster WHERE Code=C.Item),IIF(Category='2',(SELECT Name FROM PaperMaster WHERE Code=C.Item),(SELECT Name FROM BookMaster WHERE Code=C.Item))) As ItemName,ABS(C.Quantity) As Qty From MaterialSVChild C Where C.Code = '" & strOrderCode & "' And Quantity > 0 Order By Category", CxnMaterialStockAdjustment, adOpenKeyset, adLockOptimistic
    rstMaterialSVChild.ActiveConnection = Nothing
    If rstMaterialSVChild.RecordCount > 0 Then rstMaterialSVChild.MoveFirst
    i = 0
    Do While Not rstMaterialSVChild.EOF
        i = i + 1
        With fpSpread1
            .SetText 1, i, IIf(rstMaterialSVChild.Fields("Category").Value = "1", "Outsource Item", IIf(rstMaterialSVChild.Fields("Category").Value = "2", "Paper", IIf(rstMaterialSVChild.Fields("Category").Value = "3", "Fresh Book", IIf(rstMaterialSVChild.Fields("Category").Value = "4", "Repair Book", "Title"))))
            .Col = 2
            .TypeComboBoxList = IIf(rstMaterialSVChild.Fields("Category").Value = "1", OutsourceItem, IIf(rstMaterialSVChild.Fields("Category").Value = "2", Paper, IIf(rstMaterialSVChild.Fields("Category").Value = "4", RepairBook, IIf(rstMaterialSVChild.Fields("Category").Value = "3", FreshBook, Title))))
            .SetText 2, i, rstMaterialSVChild.Fields("ItemName").Value
            .SetText 3, i, Val(rstMaterialSVChild.Fields("Qty").Value)
            .SetText 4, i, rstMaterialSVChild.Fields("ItemCode").Value
        End With
        rstMaterialSVChild.MoveNext
    Loop
    If rstMaterialSVChild.State = adStateOpen Then
       rstMaterialSVChild.Close
    End If
    rstMaterialSVChild.Open "Select C.Category,C.Category+C.Item As ItemCode,IIF(Category='1',(SELECT Name FROM OutsourceItemMaster WHERE Code=C.Item),IIF(Category='2',(SELECT Name FROM PaperMaster WHERE Code=C.Item),(SELECT Name FROM BookMaster WHERE Code=C.Item))) As ItemName,ABS(C.Quantity) As Qty From MaterialSVChild C Where C.Code = '" & strOrderCode & "' And Quantity < 0 Order By Category", CxnMaterialStockAdjustment, adOpenKeyset, adLockOptimistic
    rstMaterialSVChild.ActiveConnection = Nothing
    If rstMaterialSVChild.RecordCount > 0 Then rstMaterialSVChild.MoveFirst
    i = 0
    Do While Not rstMaterialSVChild.EOF
        i = i + 1
        With fpSpread2
            .SetText 1, i, IIf(rstMaterialSVChild.Fields("Category").Value = "1", "Outsource Item", IIf(rstMaterialSVChild.Fields("Category").Value = "2", "Paper", IIf(rstMaterialSVChild.Fields("Category").Value = "3", "Fresh Book", IIf(rstMaterialSVChild.Fields("Category").Value = "4", "Repair Book", "Title"))))
            .Col = 2
            .TypeComboBoxList = IIf(rstMaterialSVChild.Fields("Category").Value = "1", OutsourceItem, IIf(rstMaterialSVChild.Fields("Category").Value = "2", Paper, IIf(rstMaterialSVChild.Fields("Category").Value = "4", RepairBook, IIf(rstMaterialSVChild.Fields("Category").Value = "3", FreshBook, Title))))
            .SetText 2, i, rstMaterialSVChild.Fields("ItemName").Value
            .SetText 3, i, Val(rstMaterialSVChild.Fields("Qty").Value)
            .SetText 4, i, rstMaterialSVChild.Fields("ItemCode").Value
        End With
        rstMaterialSVChild.MoveNext
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
    If ActionType = "D" Then
        CxnMaterialStockAdjustment.Execute "Delete From MaterialSVChild WHERE Code = '" & rstMaterialSVParent.Fields("Code").Value & "'"
    ElseIf ActionType = "I1" Then
        With fpSpread1
            .GetText 1, .ActiveRow, CellVal(1)
            .GetText 3, .ActiveRow, CellVal(2)
            .GetText 4, .ActiveRow, CellVal(3)
        End With
        CxnMaterialStockAdjustment.Execute "Insert Into MaterialSVChild Values ('" & rstMaterialSVParent.Fields("Code").Value & "','" & IIf(CellVal(1) = "Outsource Item", "1", IIf(CellVal(1) = "Paper", "2", IIf(CellVal(1) = "Fresh Book", "3", IIf(CellVal(1) = "Repair Book", "4", "5")))) & "','" & Right(CellVal(3), 6) & "'," & Val(CellVal(2)) & ")"
    Else
        With fpSpread2
            .GetText 1, .ActiveRow, CellVal(1)
            .GetText 3, .ActiveRow, CellVal(2)
            .GetText 4, .ActiveRow, CellVal(3)
        End With
        CxnMaterialStockAdjustment.Execute "Insert Into MaterialSVChild Values ('" & rstMaterialSVParent.Fields("Code").Value & "','" & IIf(CellVal(1) = "Outsource Item", "1", IIf(CellVal(1) = "Paper", "2", IIf(CellVal(1) = "Fresh Book", "3", IIf(CellVal(1) = "Repair Book", "4", "5")))) & "','" & Right(CellVal(3), 6) & "'," & 0 - Val(CellVal(2)) & ")"
    End If
    Exit Function
ErrorHandler:
    UpdateMaterialList = False
End Function
Private Sub RefreshDropDownList(ByVal xType As String)
    If xType = "R" Then
        rstOutsourceItemList.ActiveConnection = CxnMaterialStockAdjustment
        Do While Not RefreshRecord(rstOutsourceItemList)
        Loop
        rstOutsourceItemList.ActiveConnection = Nothing
        rstPaperList.ActiveConnection = CxnMaterialStockAdjustment
        Do While Not RefreshRecord(rstPaperList)
        Loop
        rstPaperList.ActiveConnection = Nothing
        rstFreshBookList.ActiveConnection = CxnMaterialStockAdjustment
        Do While Not RefreshRecord(rstFreshBookList)
        Loop
        rstFreshBookList.ActiveConnection = Nothing
        rstRepairBookList.ActiveConnection = CxnMaterialStockAdjustment
        Do While Not RefreshRecord(rstRepairBookList)
        Loop
        rstRepairBookList.ActiveConnection = Nothing
        OutsourceItem = "": Paper = "": FreshBook = "": RepairBook = "": Title = ""
    End If
    Do While Not rstOutsourceItemList.EOF
        If OutsourceItem = "" Then
            OutsourceItem = rstOutsourceItemList.Fields("Name").Value
        Else
            OutsourceItem = OutsourceItem + Chr$(9) + rstOutsourceItemList.Fields("Name").Value
        End If
        rstOutsourceItemList.MoveNext
    Loop
    Do While Not rstPaperList.EOF
        If Paper = "" Then
            Paper = rstPaperList.Fields("Name").Value
        Else
            Paper = Paper + Chr$(9) + rstPaperList.Fields("Name").Value
        End If
        rstPaperList.MoveNext
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
Private Sub PrintMaterialStockAdjustment()
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    rptStockJournal.Text1.SetText "Stock Journal Voucher"
    rptStockJournal.Text2.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptStockJournal.Text3.SetText Trim(rstCompanyMaster.Fields("Address1").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address2").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address3").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address4").Value)
    If (Not CheckEmpty(rstCompanyMaster.Fields("Phone").Value, False)) And (Not CheckEmpty(rstCompanyMaster.Fields("Fax").Value, False)) Then
        rptStockJournal.Text24.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & Space(1) & "Fax : " & Trim(rstCompanyMaster.Fields("Fax").Value)
    ElseIf Not CheckEmpty(rstCompanyMaster.Fields("Fax").Value, False) Then
        rptStockJournal.Text24.SetText "Fax : " & Trim(rstCompanyMaster.Fields("Fax").Value)
    ElseIf Not CheckEmpty(rstCompanyMaster.Fields("Phone").Value, False) Then
        rptStockJournal.Text24.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value)
    Else
        rptStockJournal.Section5.Suppress = True
    End If
    If rstMaterialSVChild.State = adStateOpen Then
        rstMaterialSVChild.Close
    End If
    rstMaterialSVChild.Open "Select Trim(Name) As VchNo,[Date] As VchDate,(Select Trim(PrintName) From AccountMaster Where Code=P.Account) As Godown,Category,IIF(Category='1',(Select Trim(PrintName) From OutsourceItemMaster Where Code=C.Item),IIF(Category='2',(Select Trim(PrintName) From PaperMaster Where Code=C.Item),(Select Trim(PrintName) From BookMaster Where Code=C.Item))) As ItemName,IIF(Quantity>=0,'Items Generated','Items Consumed') As ItemType,Quantity,Remarks From MaterialSVParent P Left Join MaterialSVChild C On (P.Code=C.Code And P.Code='" & rstMaterialSVList.Fields("Code").Value & "' )", CxnMaterialStockAdjustment, adOpenKeyset, adLockOptimistic
    rptStockJournal.Text27.SetText "for " & Trim(rstMaterialSVChild.Fields("Godown").Value)
    rptStockJournal.Text9.SetText "for " & Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptStockJournal.Database.SetDataSource rstMaterialSVChild, 3, 1
    Screen.MousePointer = vbNormal
    If OutputTo = "S" Then
        Set FrmReportViewer.Report = rptStockJournal
        FrmReportViewer.Show vbModal
    Else
        rptStockJournal.PrintOut
    End If
    Set rptStockJournal = Nothing
    On Error GoTo 0
End Sub
