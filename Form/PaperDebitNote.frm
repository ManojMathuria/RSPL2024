VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmPaperDebitNote 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Paper Debit Note"
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   13740
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "PaperDebitNote.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7725
   ScaleWidth      =   13740
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   7715
      Left            =   15
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   13715
      _Version        =   65536
      _ExtentX        =   24192
      _ExtentY        =   13608
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
      Picture         =   "PaperDebitNote.frx":0442
      Begin TabDlg.SSTab SSTab1 
         Height          =   7485
         Left            =   120
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   120
         Width           =   13485
         _ExtentX        =   23786
         _ExtentY        =   13203
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabHeight       =   520
         ShowFocusRect   =   0   'False
         Enabled         =   0   'False
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
         TabPicture(0)   =   "PaperDebitNote.frx":045E
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "DataGrid1"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Text1"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "&Details"
         TabPicture(1)   =   "PaperDebitNote.frx":047A
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
            Top             =   7020
            Width           =   12785
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   6495
            Left            =   120
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   450
            Width           =   13260
            _ExtentX        =   23389
            _ExtentY        =   11456
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
               Caption         =   "        Vch No."
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
               Caption         =   "     Vch Date"
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
               Caption         =   "Account Name"
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
               DataField       =   "Amount"
               Caption         =   "                 Amount"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
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
                  ColumnAllowSizing=   0   'False
                  Locked          =   -1  'True
                  ColumnWidth     =   1080
               EndProperty
               BeginProperty Column01 
                  ColumnAllowSizing=   0   'False
                  Locked          =   -1  'True
                  ColumnWidth     =   1019.906
               EndProperty
               BeginProperty Column02 
                  ColumnAllowSizing=   0   'False
                  Locked          =   -1  'True
                  ColumnWidth     =   9074.835
               EndProperty
               BeginProperty Column03 
                  Alignment       =   1
                  ColumnAllowSizing=   0   'False
                  Locked          =   -1  'True
               EndProperty
            EndProperty
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   6690
            Left            =   -74880
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   480
            Width           =   13260
            _Version        =   65536
            _ExtentX        =   23389
            _ExtentY        =   11800
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
            Picture         =   "PaperDebitNote.frx":0496
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
               Width           =   11715
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
               Width           =   11715
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
               Caption         =   " Vch No."
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "PaperDebitNote.frx":04B2
               Picture         =   "PaperDebitNote.frx":04CE
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Left            =   11215
               TabIndex        =   13
               Top             =   105
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
               Caption         =   " Vch Date"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "PaperDebitNote.frx":04EA
               Picture         =   "PaperDebitNote.frx":0506
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
               Caption         =   " Account Name"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "PaperDebitNote.frx":0522
               Picture         =   "PaperDebitNote.frx":053E
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
               Picture         =   "PaperDebitNote.frx":055A
               Picture         =   "PaperDebitNote.frx":0576
            End
            Begin TDBDate6Ctl.TDBDate MhDateInput1 
               Height          =   330
               Left            =   12060
               TabIndex        =   1
               Top             =   105
               Width           =   1095
               _Version        =   65536
               _ExtentX        =   1931
               _ExtentY        =   582
               Calendar        =   "PaperDebitNote.frx":0592
               Caption         =   "PaperDebitNote.frx":06AA
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "PaperDebitNote.frx":0716
               Keys            =   "PaperDebitNote.frx":0734
               Spin            =   "PaperDebitNote.frx":0792
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
               Height          =   5115
               Left            =   120
               TabIndex        =   4
               Top             =   1470
               Width           =   13035
               _Version        =   524288
               _ExtentX        =   22992
               _ExtentY        =   9022
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
               SpreadDesigner  =   "PaperDebitNote.frx":07BA
            End
            Begin VB.TextBox Text5 
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
               Left            =   840
               MaxLength       =   100
               TabIndex        =   16
               TabStop         =   0   'False
               Top             =   3060
               Width           =   11715
            End
            Begin VB.Line Line1 
               X1              =   0
               X2              =   13305
               Y1              =   525
               Y2              =   525
            End
            Begin VB.Line Line2 
               X1              =   0
               X2              =   13305
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
            Top             =   7020
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
      Width           =   13740
      _ExtentX        =   24236
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
Attribute VB_Name = "FrmPaperDebitNote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cnxPaperDebitNote As New ADODB.Connection
Dim rstPaperDNList As New ADODB.Recordset
Dim rstPaperDNParent As New ADODB.Recordset
Dim rstPaperDNChild As New ADODB.Recordset
Dim rstAccountList As New ADODB.Recordset
Dim rstPaperList As New ADODB.Recordset
Dim AccountCode As String
Dim PaperCode As String
Dim oOutlook As New Outlook.Application
Dim EditMode As Boolean
Dim SortOrder As String
Dim PrevStr As String
Dim dblBookMark As Double
Dim blnRecordExist As Boolean
Private Sub Form_Load()
    'On Error GoTo ErrorHandler
    CenterForm Me
    BusySystemIndicator True
    cnxPaperDebitNote.CursorLocation = adUseClient
    If cnxPaperDebitNote.State = adStateOpen Then cnxPaperDebitNote.Close
    cnxPaperDebitNote.Open CxnDatabase.ConnectionString
    rstAccountList.Open "SELECT TRIM(Name)+' ('+CHOOSE(VAL(Type)-4,'Book Printer','Title Printer','','Book Binder')+')' As Col0,Code FROM AccountMaster WHERE Type IN ('05','06','08') ORDER BY Name", cnxPaperDebitNote, adOpenKeyset, adLockReadOnly
    rstPaperList.Open "SELECT TRIM(Name)+' ('+CHOOSE(VAL(Type),'Book','Title')+' Paper)' As Col0,[Weight/Ream],Code FROM PaperMaster ORDER BY Name", cnxPaperDebitNote, adOpenKeyset, adLockReadOnly
    rstPaperDNList.Open "SELECT T.Code,T.Name,T.Date,M.Name As AccountName,Amount FROM PaperDNParent T INNER JOIN AccountMaster M ON T.Account=M.Code ORDER BY T.Name", cnxPaperDebitNote, adOpenKeyset, adLockOptimistic
    rstPaperDNParent.CursorLocation = adUseClient
    rstPaperDNList.Filter = adFilterNone
    If rstPaperDNList.RecordCount > 0 Then rstPaperDNList.MoveLast
    Set DataGrid1.DataSource = rstPaperDNList
    BusySystemIndicator False
    SSTab1.Tab = 0
    SortOrder = "Name"
    If Not (rstPaperDNList.EOF Or rstPaperDNList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    rstPaperDNList.ActiveConnection = Nothing
    rstAccountList.ActiveConnection = Nothing
    rstPaperList.ActiveConnection = Nothing
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
           If Me.ActiveControl.Name <> "fpSpread1" Then SendKeys "{TAB}"
        End If
        If Me.ActiveControl.Name <> "fpSpread1" Then KeyCode = 0
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Toolbar1.Buttons.Item(4).Enabled Then Call Form_KeyDown(vbKeyEscape, 0): Cancel = 1
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CloseRecordset(rstPaperDNList)
    Call CloseRecordset(rstPaperDNParent)
    Call CloseRecordset(rstPaperDNChild)
    Call CloseRecordset(rstAccountList)
    Call CloseRecordset(rstPaperList)
    Call CloseConnection(cnxPaperDebitNote)
    ShowProgressInStatusBar False
    DisableChildMenu
End Sub
Private Sub Text1_Change()
    If rstPaperDNList.RecordCount = 0 Then Exit Sub
    rstPaperDNList.MoveFirst
    If Text1.Text <> "" Then
        If SortOrder = "Name" Then rstPaperDNList.Find "[" & SortOrder & "] Like '%" & FixQuote(Text1.Text) & "%'" Else rstPaperDNList.Find "[" & SortOrder & "] Like '" & FixQuote(Text1.Text) & "%'"
        If rstPaperDNList.EOF Then
            rstPaperDNList.MoveFirst
            If PrevStr <> "" And Len(Text1.Text) > 1 Then
                If dblBookMark <> 0 Then rstPaperDNList.Bookmark = dblBookMark
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
    If Not (rstPaperDNList.EOF Or rstPaperDNList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim KeyProcessed As Boolean
    If rstPaperDNList.RecordCount = 0 Then Exit Sub
    If Shift = 0 And KeyCode = vbKeyUp Then
        With rstPaperDNList
            .MovePrevious
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyBack Then
        With rstPaperDNList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyDown Then
        With rstPaperDNList
            .MoveNext
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageUp Then
        With rstPaperDNList
            .Move (-1) * (DataGrid1.VisibleRows - 1)
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageUp Then
        With rstPaperDNList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageDown Then
        With rstPaperDNList
            .Move DataGrid1.VisibleRows - 1
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageDown Then
        With rstPaperDNList
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
            If Not (rstPaperDNList.EOF Or rstPaperDNList.BOF) Then
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
    Dim HiLiteRecord As Boolean
    Dim UpdateFlag As Integer, i As Integer
    Dim CellVal01 As Variant, CellVal02 As Variant
    If Button.Index = 1 Then
        If rstPaperDNParent.State = adStateOpen Then rstPaperDNParent.Close
        rstPaperDNParent.Open "SELECT * FROM PaperDNParent WHERE Code=''", cnxPaperDebitNote, adOpenKeyset, adLockOptimistic
        ClearFields
        If AddRecord(rstPaperDNParent) Then
            Text2.Text = GenerateCode(cnxPaperDebitNote, "SELECT MAX(VAL(Name)) FROM PaperDNParent", 10, Space(1))
            MhDateInput1.Text = Format(Date, "dd-MM-yyyy")
            Call SetButtons(False)
            SSTab1.Tab = 1
            Text3.SetFocus
            blnRecordExist = False
            cnxPaperDebitNote.BeginTrans
        End If
    ElseIf Button.Index = 2 Then
        If rstPaperDNList.RecordCount = 0 Then Exit Sub
        SSTab1.Tab = 1
        EditRecord
    ElseIf Button.Index = 3 Then
        If rstPaperDNList.RecordCount = 0 Then Exit Sub
        If AllowTransactionsDeletion = 0 Then Call DisplayError("You don't have the rights to delete Paper Debit Note"): Exit Sub
        SSTab1.Tab = 1
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            On Error Resume Next
            MdiMainMenu.MousePointer = vbHourglass
            cnxPaperDebitNote.Execute "DELETE FROM PaperDNParent WHERE Code='" & rstPaperDNList.Fields("Code").Value & "'"
            MdiMainMenu.MousePointer = vbNormal
            If Err.Number = 0 Then
                rstPaperDNList.Delete
                rstPaperDNList.MoveNext
                If rstPaperDNList.RecordCount > 0 And rstPaperDNList.EOF Then rstPaperDNList.MoveLast
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
        If UpdateRecord(rstPaperDNParent) Then
            If UpdatePaperList("D") Then
                UpdateFlag = 1
                For i = 1 To fpSpread1.DataRowCnt
                    fpSpread1.SetActiveCell 6, i
                    fpSpread1.GetText 6, i, CellVal01
                    fpSpread1.GetText 7, i, CellVal02
                    If Val(CellVal01) <> 0 And CellVal02 <> "" Then
                        If Not UpdatePaperList("I") Then UpdateFlag = 0: Exit For
                    End If
                Next
            End If
        End If
        If UpdateFlag Then
            AddToList
            cnxPaperDebitNote.CommitTrans
            If rstPaperDNParent.State = adStateOpen Then rstPaperDNParent.Close
            rstPaperDNParent.CursorLocation = adUseClient
            Call SetButtons(True)
            SSTab1.Tab = 0
            ShowProgressInStatusBar True
            Timer1.Enabled = True
        Else
            DisplayError ("Failed to save the record")
            Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
        End If
    ElseIf Button.Index = 5 Then    'Cancel
        If CancelRecordUpdate(rstPaperDNParent) Then
            cnxPaperDebitNote.RollbackTrans
            If rstPaperDNParent.State = adStateOpen Then rstPaperDNParent.Close
            rstPaperDNParent.CursorLocation = adUseClient
            Call SetButtons(True)
            SetButtonsForNoRecord
            SSTab1.Tab = 0
        End If
    ElseIf Button.Index = 6 Then    'Refresh
        SSTab1.Tab = 0
        Set DataGrid1.DataSource = Nothing
        rstPaperDNList.ActiveConnection = cnxPaperDebitNote
        Do While Not RefreshRecord(rstPaperDNList)
        Loop
        Set DataGrid1.DataSource = rstPaperDNList
        rstPaperDNList.ActiveConnection = Nothing
        If rstPaperDNList.RecordCount > 0 Then rstPaperDNList.MoveLast
        rstAccountList.ActiveConnection = cnxPaperDebitNote
        Do While Not RefreshRecord(rstAccountList)
        Loop
        rstAccountList.ActiveConnection = Nothing
        rstPaperList.ActiveConnection = cnxPaperDebitNote
        Do While Not RefreshRecord(rstPaperList)
        Loop
        rstPaperList.ActiveConnection = Nothing
        HiLiteRecord = True
    ElseIf Button.Index = 7 Then
        SSTab1.Tab = 0
        With FrmFilter
            .Combo1.AddItem "Account", 0
            .Combo1.ListIndex = 0
            Set .srcForm = Me
            .Show vbModal
        End With
        HiLiteRecord = True
    ElseIf Button.Index = 9 Then
        If rstPaperDNList.RecordCount = 0 Then Exit Sub
        Call PrintPaperDebitNote(rstPaperDNList.Fields("Code").Value, "P")
        HiLiteRecord = True
    ElseIf Button.Index = 10 Then
        If rstPaperDNList.RecordCount = 0 Then Exit Sub
        Call PrintPaperDebitNote(rstPaperDNList.Fields("Code").Value, "S")
        HiLiteRecord = True
    ElseIf Button.Index = 11 Then
        If rstPaperDNList.RecordCount = 0 Then Exit Sub
        Call PrintPaperDebitNote(rstPaperDNList.Fields("Code").Value, "M")
        HiLiteRecord = True
    ElseIf Button.Index = 13 Then
        If rstPaperDNList.RecordCount > 0 Then rstPaperDNList.MoveFirst
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 14 Then
        If rstPaperDNList.RecordCount > 0 Then
            rstPaperDNList.MovePrevious
            If rstPaperDNList.BOF Then rstPaperDNList.MoveNext
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 15 Then
        If rstPaperDNList.RecordCount > 0 Then
            rstPaperDNList.MoveNext
            If rstPaperDNList.EOF Then rstPaperDNList.MovePrevious
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 16 Then
        If rstPaperDNList.RecordCount > 0 Then rstPaperDNList.MoveLast
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 18 Then
        Unload Me
        HiLiteRecord = False
    End If
    If HiLiteRecord Then
        If Not (rstPaperDNList.EOF Or rstPaperDNList.BOF) Then
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
        rstPaperDNList.Sort = "[" + SortOrder & "] Asc"
    End If
    DataGrid1.ClearSelCols
    If Not (rstPaperDNList.EOF Or rstPaperDNList.BOF) Then
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
    If rstPaperDNList.RecordCount = 0 Then
        Toolbar1.Buttons.Item(2).Enabled = False
        Toolbar1.Buttons.Item(3).Enabled = False
        Toolbar1.Buttons.Item(9).Enabled = False
        Toolbar1.Buttons.Item(10).Enabled = False
        Toolbar1.Buttons.Item(11).Enabled = False
        Toolbar1.Buttons.Item(13).Enabled = False
        Toolbar1.Buttons.Item(14).Enabled = False
        Toolbar1.Buttons.Item(15).Enabled = False
        Toolbar1.Buttons.Item(16).Enabled = False
    End If
End Sub
Private Sub Text2_Validate(Cancel As Boolean)
    If rstPaperDNParent.EOF Or rstPaperDNParent.BOF Then Exit Sub
    If CheckEmpty(Text2, True) Then
        Cancel = True
    ElseIf CheckDuplicate(cnxPaperDebitNote, "PaperDNParent", "Code", "[Name]", Trim(Text2.Text), rstPaperDNParent.Fields("Code").Value, False) Then
        Cancel = True
    End If
End Sub
Private Sub MhDateInput1_Validate(Cancel As Boolean)
    If Not IsDate(GetDate(MhDateInput1.Text)) Then Cancel = True
End Sub
Private Sub Text3_Change()
    If Text3.Text = " " Then Text3.Text = "?": SendKeys "{TAB}"
End Sub
Private Sub Text3_Validate(Cancel As Boolean)
    Dim SearchString As String
    SearchString = FixQuote(Text3.Text)
    If rstAccountList.RecordCount = 0 Then DisplayError ("No Record in Account Master"): Cancel = True: Exit Sub Else rstAccountList.MoveFirst
    rstAccountList.Find "[Col0] = '" & RTrim(SearchString) & "'"
    If rstAccountList.EOF Then
        SelectionType = "S"
        AccountCode = ""
        Call LoadSelectionList(rstAccountList, "List of Accounts...", "Name")
        SearchOrder = 0
        Call DisplaySelectionList(Text3, AccountCode)
        Call CloseForm(FrmSelectionList)
        If CheckEmpty(Text3.Text, False) Then Text3.Text = "?"
        If RTrim(AccountCode) <> "" Then SendKeys "{TAB}"
        Cancel = True
    Else
        AccountCode = rstAccountList.Fields("Code").Value
    End If
End Sub
Private Sub ViewRecord()
    ClearFields
    If rstPaperDNList.EOF Then
        If rstPaperDNChild.State = adStateOpen Then rstPaperDNChild.Close
        Exit Sub
    End If
    FindRecord
    LoadFields
End Sub
Private Sub FindRecord()
    If rstPaperDNParent.State = adStateOpen Then rstPaperDNParent.Close
    rstPaperDNParent.Open "SELECT * FROM PaperDNParent WHERE Code='" & FixQuote(rstPaperDNList.Fields("Code").Value) & "'", cnxPaperDebitNote, adOpenKeyset, adLockOptimistic
    If rstPaperDNParent.RecordCount = 0 Then
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
End Sub
Private Sub LoadFields()
    If rstPaperDNParent.EOF Or rstPaperDNParent.BOF Then Exit Sub
    Text2.Text = rstPaperDNParent.Fields("Name").Value
    MhDateInput1.Text = Format(rstPaperDNParent.Fields("Date").Value, "dd-MM-yyyy")
    AccountCode = rstPaperDNParent.Fields("Account").Value
    If rstAccountList.RecordCount > 0 Then rstAccountList.MoveFirst
    rstAccountList.Find "[Code] = '" & AccountCode & "'"
    If Not rstAccountList.EOF Then Text3.Text = rstAccountList.Fields("Col0").Value
    Text4.Text = rstPaperDNParent.Fields("Remarks").Value
    Call LoadPaperList(rstPaperDNParent.Fields("Code").Value)
End Sub
Private Sub EditRecord()
    On Error GoTo ErrorHandler
    If rstPaperDNParent.RecordCount = 0 Then Exit Sub
    If rstPaperDNChild.State = adStateClosed Then SSTab1.Tab = 0: Exit Sub
    If rstPaperDNParent.State = adStateOpen Then rstPaperDNParent.Close
    rstPaperDNParent.CursorLocation = adUseServer
    rstPaperDNParent.Open "SELECT * FROM PaperDNParent WHERE Code='" & FixQuote(rstPaperDNList.Fields("Code").Value) & "'", cnxPaperDebitNote, adOpenKeyset, adLockPessimistic
    MdiMainMenu.MousePointer = vbHourglass
    rstPaperDNParent.Fields("LockStatus") = "N"
    MdiMainMenu.MousePointer = vbNormal
    AddToList
    Call SetButtons(False)
    SSTab1.TabEnabled(0) = False
    Text3.SetFocus
    blnRecordExist = True
    cnxPaperDebitNote.BeginTrans
    Exit Sub
ErrorHandler:
    If Err.Number = -2147467259 Then Call DisplayError("Failed to Edit the record")
    MdiMainMenu.MousePointer = vbNormal
    SSTab1.Tab = 0
End Sub
Private Sub SaveFields()
    If rstPaperDNParent.EOF Or rstPaperDNParent.BOF Then Exit Sub
    If Not blnRecordExist Then
        rstPaperDNParent.Fields("Code").Value = GenerateCode(cnxPaperDebitNote, "SELECT MAX(Code) FROM PaperDNParent", 6, "0")
        rstPaperDNParent.Fields("CreatedBy").Value = UserCode
        rstPaperDNParent.Fields("CreatedOn").Value = Now()
    Else
        rstPaperDNParent.Fields("ModifiedBy").Value = UserCode
        rstPaperDNParent.Fields("ModifiedOn").Value = Now()
    End If
    rstPaperDNParent.Fields("Name").Value = Pad(Trim(Text2.Text), Space(1), 10, "L")
    rstPaperDNParent.Fields("Date").Value = GetDate(MhDateInput1.Text)
    rstPaperDNParent.Fields("Account").Value = AccountCode
    rstPaperDNParent.Fields("Remarks").Value = Trim(Text4.Text)
    rstPaperDNParent.Fields("Amount").Value = CalculateTotalAmt
    rstPaperDNParent.Fields("LockStatus").Value = "N"
End Sub
Private Sub AddToList()
    On Error Resume Next
    rstPaperDNList.MoveFirst
    rstPaperDNList.Find "[Code] = '" & rstPaperDNParent.Fields("Code").Value & "'"
    If rstPaperDNList.EOF Then rstPaperDNList.AddNew
    rstPaperDNList.Fields("Code").Value = rstPaperDNParent.Fields("Code").Value
    rstPaperDNList.Fields("Name").Value = Pad(rstPaperDNParent.Fields("Name").Value, Space(1), 10, "L")
    rstPaperDNList.Fields("Date").Value = rstPaperDNParent.Fields("Date").Value
    rstAccountList.MoveFirst
    rstAccountList.Find "[Code] = '" & rstPaperDNParent.Fields("Account").Value & "'"
    rstPaperDNList.Fields("AccountName").Value = Left(Trim(rstAccountList.Fields("Col0").Value), 40)
    rstPaperDNList.Fields("Amount").Value = CalculateTotalAmt
    rstPaperDNList.Update
    rstPaperDNList.Sort = SortOrder & " Asc"
    rstPaperDNList.Find "[Code] = '" & rstPaperDNParent.Fields("Code").Value & "'"
End Sub
Private Function CheckMandatoryFields() As Boolean
    If CheckEmpty(Text2.Text, False) Then
       DisplayError ("Vch No. cannot be blank")
       Text2.SetFocus
       CheckMandatoryFields = True
    ElseIf CheckEmpty(Text3.Text, False) Then
       Text3.SetFocus
       CheckMandatoryFields = True
    ElseIf Not CheckExists(Text3, "Col0", rstAccountList, AccountCode) Then
        Text3.SetFocus
        CheckMandatoryFields = True
    ElseIf CheckDuplicate(cnxPaperDebitNote, "PaperDNParent", "Code", "[Name]", Trim(Text2.Text), rstPaperDNParent.Fields("Code").Value, False) Then
        Text2.SetFocus
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
    If SrchFor = "Account" Then rstPaperDNList.Filter = "[AccountName] Like '%" & SrchText & "%'"
End Sub
Private Sub fpSpread1_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask And KeyCode = vbKeyD Then
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then fpSpread1.DeleteRows fpSpread1.ActiveRow, 1: fpSpread1.SetFocus
    ElseIf KeyCode = vbKeySpace Then
        Dim Paper As Variant
        With fpSpread1
            If .ActiveCol = 1 Then
                .GetText .ActiveCol, .ActiveRow, Paper
                Text5.Text = FixQuote(Paper)
                If rstPaperList.RecordCount = 0 Then DisplayError ("No Record in Paper Master"): .SetActiveCell 1, .ActiveRow: Exit Sub Else rstPaperList.MoveFirst
                rstPaperList.Find "[Col0] = '" & RTrim(Paper) & "'"
                SelectionType = "S"
                PaperCode = ""
                Call LoadSelectionList(rstPaperList, "List of Papers...", "Name")
                SearchOrder = 0
                Call DisplaySelectionList(Text5, PaperCode)
                Call CloseForm(FrmSelectionList)
                If PaperCode = "" Then
                    .SetActiveCell 1, .ActiveRow
                Else
                    rstPaperList.MoveFirst: rstPaperList.Find "[Code] ='" & PaperCode & "'"
                    .SetText 1, .ActiveRow, Text5.Text
                    .SetText 3, .ActiveRow, Val(rstPaperList.Fields("Weight/Ream").Value)
                    .SetText 7, .ActiveRow, PaperCode
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
            If Paper = "" Then .SetText 4, Row, "": .SetText 6, Row, "" Else .SetText 4, Row, GrWt: .SetText 6, Row, GrWt * Rate
        End If
    End With
End Sub
Private Sub fpSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    EditMode = IIf(Mode = 1, True, False)
End Sub
Private Sub LoadPaperList(ByVal DNCode As String)
    Dim i As Integer
    On Error GoTo ErrorHandler
    If rstPaperDNChild.State = adStateOpen Then rstPaperDNChild.Close
    rstPaperDNChild.Open "SELECT Paper,TRIM(M.Name)+' ('+CHOOSE(VAL(Type),'Book','Title')+' Paper)' As PaperName,Quantity,M.[Weight/Ream],Weight,Rate,Amount FROM PaperDNChild T INNER JOIN PaperMaster M ON T.Paper=M.Code WHERE T.Code='" & DNCode & "' ORDER BY M.Name", cnxPaperDebitNote, adOpenKeyset, adLockOptimistic
    rstPaperDNChild.ActiveConnection = Nothing
    If rstPaperDNChild.RecordCount > 0 Then rstPaperDNChild.MoveFirst
    i = 0
    Do While Not rstPaperDNChild.EOF
        i = i + 1
        With fpSpread1
            .SetText 1, i, rstPaperDNChild.Fields("PaperName").Value
            .SetText 2, i, Val(rstPaperDNChild.Fields("Quantity").Value)
            .SetText 3, i, Val(rstPaperDNChild.Fields("Weight/Ream").Value)
            .SetText 4, i, Val(rstPaperDNChild.Fields("Weight").Value)
            .SetText 5, i, Val(rstPaperDNChild.Fields("Rate").Value)
            .SetText 6, i, Val(rstPaperDNChild.Fields("Amount").Value)
            .SetText 7, i, rstPaperDNChild.Fields("Paper").Value
        End With
        rstPaperDNChild.MoveNext
    Loop
    fpSpread1.SetActiveCell 1, 1
    Exit Sub
ErrorHandler:
    DisplayError ("Failed to Load Paper List")
End Sub
Private Function CalculateTotalAmt() As Long
    Dim i As Integer, Amount As Variant
    For i = 1 To fpSpread1.DataRowCnt
        fpSpread1.GetText 6, i, Amount
        CalculateTotalAmt = CalculateTotalAmt + Val(Amount)
    Next
End Function
Private Function UpdatePaperList(ByVal ActionType As String) As Boolean
    Dim CellVal(1 To 5) As Variant
    On Error GoTo ErrorHandler
    UpdatePaperList = True
    If ActionType = "D" And (Not blnRecordExist) Then Exit Function
    If ActionType = "D" Then
        cnxPaperDebitNote.Execute "DELETE FROM PaperDNChild WHERE Code='" & rstPaperDNParent.Fields("Code").Value & "'"
    Else
        With fpSpread1
            .GetText 2, .ActiveRow, CellVal(1)  'Quantity
            .GetText 4, .ActiveRow, CellVal(2)  'Weight
            .GetText 5, .ActiveRow, CellVal(3)  'Rate
            .GetText 6, .ActiveRow, CellVal(4)  'Amount
            .GetText 7, .ActiveRow, CellVal(5)  'Paper
        End With
        cnxPaperDebitNote.Execute "INSERT INTO PaperDNChild VALUES ('" & rstPaperDNParent.Fields("Code").Value & "','" & CellVal(5) & "'," & Val(CellVal(1)) & "," & Val(CellVal(2)) & "," & Val(CellVal(3)) & "," & Val(CellVal(4)) & ")"
    End If
    Exit Function
ErrorHandler:
    UpdatePaperList = False
End Function
Private Sub PrintPaperDebitNote(ByVal VchNo As String, ByVal OutputTo As String)
    Dim rstCompanyMaster As New ADODB.Recordset, rstPaperDebitNote As New ADODB.Recordset
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    If rstCompanyMaster.State = adStateOpen Then rstCompanyMaster.Close: If rstPaperDebitNote.State = adStateOpen Then rstPaperDebitNote.Close
    rstCompanyMaster.Open "SELECT PrintName,Address1,Address2,Address3,Address4,Phone,Fax,eMail FROM CompanyMaster", CxnDatabase, adOpenKeyset, adLockReadOnly
    rstPaperDebitNote.Open "SELECT 'DN/SP/" & Right(Year(FinancialYearFrom), 2) + "-" + Right(Year(FinancialYearTo), 2) & "/'+TRIM(P.Name) As VchNo,P.Date As VchDate,P.Remarks,M2.PrintName As AccountName,M1.PrintName As PaperName,Quantity,[Weight/Ream],Weight,Rate,C.Amount,P.Amount As DNAmt,eMail FROM ((PaperDNParent P INNER JOIN PaperDNChild C ON P.Code=C.Code) INNER JOIN PaperMaster M1 ON C.Paper=M1.Code) INNER JOIN AccountMaster M2 ON P.Account=M2.Code WHERE P.Code='" & VchNo & "' ORDER BY M1.PrintName", CxnDatabase, adOpenKeyset, adLockReadOnly
    rstPaperDebitNote.ActiveConnection = Nothing
    Screen.MousePointer = vbNormal
    rptPaperDebitNote.Text2.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptPaperDebitNote.Text3.SetText Trim(rstCompanyMaster.Fields("Address1").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address2").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address3").Value) & Space(1) & Trim(rstCompanyMaster.Fields("Address4").Value)
    rptPaperDebitNote.Text24.SetText "Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & Space(1) & "Fax : " & Trim(rstCompanyMaster.Fields("Fax").Value) & Space(1) & "e-Mail : " & Trim(rstCompanyMaster.Fields("EMail").Value)
    rptPaperDebitNote.Text18.SetText " (" & Trim(NumberToWords(rstPaperDebitNote.Fields("DNAmt").Value, True)) & ")"
    rptPaperDebitNote.Text9.SetText "for " & Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptPaperDebitNote.Text28.SetText "Total Amount Debited due to " & Chr(34) & "PAPER A/C ADJUSTMENT" + Chr(34)
    rptPaperDebitNote.DiscardSavedData
    rptPaperDebitNote.Database.SetDataSource rstPaperDebitNote, 3, 1
    If OutputTo = "S" Then
        Set FrmReportViewer.Report = rptPaperDebitNote
        FrmReportViewer.Show vbModal
    ElseIf OutputTo = "P" Then
        rptPaperDebitNote.PrintOut False    'Print Report Without Prompt
    Else
        Dim oOutlookMsg As Outlook.MailItem, FileName As String
        Set oOutlookMsg = oOutlook.CreateItem(olMailItem)
        With oOutlookMsg
            .To = rstPaperDebitNote.Fields("EMail").Value
            .Subject = "Debit Note #" & Trim(rstPaperDebitNote.Fields("VchNo").Value)
            .HTMLBody = "<Font Face='Calibri' Size='3'>Dear Sir,<Br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Please find attached herewith Debit Note #" & Trim(rstPaperDebitNote.Fields("VchNo").Value) & " against " + Chr(34) + "PAPER A/C ADJUSTMENT" + Chr(34) + " for doing the needful at your end.<Br><b>Kindly do acknowledge the receipt of the mail</b>.<Br><Br>Thanks & Regards<Br>Production Department<Br>" & Trim(rstCompanyMaster.Fields("PrintName").Value) & "<Br>Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & "<Br>E-Mail : <a HRef='mailto:" & Trim(rstCompanyMaster.Fields("EMail").Value) & "'>" & Trim(rstCompanyMaster.Fields("EMail").Value) & "</a></Font>"
            rptPaperDebitNote.ExportOptions.FormatType = crEFTPortableDocFormat    ' Set the Export Format As .Pdf
            rptPaperDebitNote.ExportOptions.DestinationType = crEDTDiskFile
            FileName = FixAPIString(GetTemporaryFileName): FileName = Mid(FileName, 1, Len(FileName) - 4) & ".Pdf"
            rptPaperDebitNote.ExportOptions.DiskFileName = FileName
            rptPaperDebitNote.Export False
            .Attachments.Add (FileName)
            .Importance = olImportanceHigh
            .ReadReceiptRequested = True
            If CheckEmpty(.To, False) Then .Display Else .Send
        End With
        Set oOutlookMsg = Nothing
    End If
    Set rptPaperDebitNote = Nothing
    Call CloseRecordset(rstPaperDebitNote): Call CloseRecordset(rstCompanyMaster)
    On Error GoTo 0
End Sub

