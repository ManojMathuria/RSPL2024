VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmUserMaster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Master"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6750
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "UserMaster.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4875
   ScaleWidth      =   6750
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   4870
      Left            =   15
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   0
      Width           =   6735
      _Version        =   65536
      _ExtentX        =   11880
      _ExtentY        =   8590
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
      Picture         =   "UserMaster.frx":0442
      Begin TabDlg.SSTab SSTab1 
         Height          =   4630
         Left            =   120
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   120
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   8176
         _Version        =   393216
         Style           =   1
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
         TabPicture(0)   =   "UserMaster.frx":045E
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "DataGrid1"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Text1"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "&Details"
         TabPicture(1)   =   "UserMaster.frx":047A
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Mh3dFrame2"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "&Privileges"
         TabPicture(2)   =   "UserMaster.frx":0496
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Mh3dFrame4"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
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
            TabIndex        =   18
            Top             =   4160
            Width           =   5775
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   3640
            Left            =   120
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   450
            Width           =   6255
            _ExtentX        =   11033
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
            ColumnCount     =   1
            BeginProperty Column00 
               DataField       =   "Name"
               Caption         =   "Name"
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
                  Locked          =   -1  'True
                  ColumnWidth     =   5940.284
               EndProperty
            EndProperty
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
            Height          =   1805
            Left            =   -74880
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   480
            Width           =   6255
            _Version        =   65536
            _ExtentX        =   11033
            _ExtentY        =   3184
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
            Picture         =   "UserMaster.frx":04B2
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
               IMEMode         =   3  'DISABLE
               Left            =   1680
               MaxLength       =   10
               PasswordChar    =   "*"
               TabIndex        =   3
               Top             =   1055
               Width           =   4455
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
               IMEMode         =   3  'DISABLE
               Left            =   1680
               MaxLength       =   10
               PasswordChar    =   "*"
               TabIndex        =   2
               Top             =   740
               Width           =   4455
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
               TabIndex        =   1
               Top             =   425
               Width           =   4455
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
               MaxLength       =   40
               TabIndex        =   0
               Top             =   100
               Width           =   4455
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
               Height          =   330
               Left            =   120
               TabIndex        =   13
               Top             =   425
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
               Caption         =   " Print Name"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "UserMaster.frx":04CE
               Picture         =   "UserMaster.frx":04EA
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
               Height          =   330
               Left            =   120
               TabIndex        =   12
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
               Caption         =   " Name"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "UserMaster.frx":0506
               Picture         =   "UserMaster.frx":0522
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
               Height          =   330
               Left            =   120
               TabIndex        =   21
               Top             =   740
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
               Caption         =   " Password"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "UserMaster.frx":053E
               Picture         =   "UserMaster.frx":055A
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
               Height          =   330
               Left            =   120
               TabIndex        =   22
               Top             =   1055
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
               Caption         =   " Confirm Password"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "UserMaster.frx":0576
               Picture         =   "UserMaster.frx":0592
            End
            Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
               Height          =   330
               Left            =   120
               TabIndex        =   23
               Top             =   1370
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
               Caption         =   " Level"
               Alignment       =   0
               FillColor       =   8421376
               TextColor       =   16777215
               Picture         =   "UserMaster.frx":05AE
               Picture         =   "UserMaster.frx":05CA
            End
            Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame3 
               Height          =   330
               Left            =   1680
               TabIndex        =   24
               TabStop         =   0   'False
               Top             =   1370
               Width           =   4455
               _Version        =   65536
               _ExtentX        =   7858
               _ExtentY        =   582
               _StockProps     =   77
               TintColor       =   16711935
               Alignment       =   0
               AutoSize        =   0   'False
               BevelSize       =   0
               BevelStyle      =   0
               BorderColor     =   -2147483642
               BorderStyle     =   1
               FillColor       =   16777215
               FontStyle       =   0
               FontTransparent =   0   'False
               LightColor      =   -2147483643
               ShadowColor     =   -2147483632
               TextColor       =   -2147483640
               WallPaper       =   0
               NoPrefix        =   0   'False
               FormatString    =   ""
               Caption         =   ""
               Picture         =   "UserMaster.frx":05E6
               Begin VB.OptionButton Option1 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Administrator"
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Left            =   120
                  TabIndex        =   4
                  Top             =   60
                  Width           =   1455
               End
               Begin VB.OptionButton Option2 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Manager"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Left            =   1880
                  TabIndex        =   5
                  Top             =   60
                  Width           =   1095
               End
               Begin VB.OptionButton Option3 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Operator"
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Left            =   3240
                  TabIndex        =   6
                  Top             =   60
                  Width           =   1095
               End
            End
         End
         Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame4 
            Height          =   4040
            Left            =   -74880
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   480
            Width           =   6255
            _Version        =   65536
            _ExtentX        =   11033
            _ExtentY        =   7126
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
            Picture         =   "UserMaster.frx":0602
            Begin VB.CheckBox Check4 
               Caption         =   "Allow Deletion of Transactions"
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
               Left            =   3360
               TabIndex        =   10
               Top             =   430
               Width           =   2880
            End
            Begin VB.CheckBox Check3 
               Caption         =   "Allow Modification of Transactions"
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
               Left            =   120
               TabIndex        =   9
               Top             =   430
               Width           =   3255
            End
            Begin VB.CheckBox Check2 
               Caption         =   "Allow Deletion of Masters"
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
               Left            =   3360
               TabIndex        =   8
               Top             =   100
               Width           =   2535
            End
            Begin VB.CheckBox Check1 
               Caption         =   "Allow Modification of Masters"
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
               Left            =   120
               TabIndex        =   7
               Top             =   100
               Width           =   2895
            End
            Begin MSComctlLib.TreeView TreeView1 
               Height          =   2970
               Left            =   120
               TabIndex        =   11
               Top             =   960
               Width           =   6015
               _ExtentX        =   10610
               _ExtentY        =   5239
               _Version        =   393217
               Indentation     =   0
               LabelEdit       =   1
               LineStyle       =   1
               Style           =   6
               Checkboxes      =   -1  'True
               BorderStyle     =   1
               Appearance      =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.Line Line1 
               X1              =   0
               X2              =   6240
               Y1              =   860
               Y2              =   860
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
            TabIndex        =   20
            Top             =   4160
            Width           =   495
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   6750
      _ExtentX        =   11906
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
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Print"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Print Preview"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
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
Attribute VB_Name = "FrmUserMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CxnUserMaster As New ADODB.Connection
Dim rstUserList As New ADODB.Recordset
Dim rstUserMaster As New ADODB.Recordset
Dim rstUserChild As New ADODB.Recordset
Dim oEncrypt As New clsBlowFish
Dim PrevStr As String
Dim dblBookMark As Double
Dim blnRecordExist As Boolean
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    
    CenterForm Me
    BusySystemIndicator True
    CxnUserMaster.CursorLocation = adUseClient
    CxnUserMaster.Open CxnDatabase.ConnectionString
    rstUserList.Open "select Name,Code From UserMaster Order By Name", CxnUserMaster, adOpenKeyset, adLockOptimistic
    rstUserMaster.CursorLocation = adUseClient
    rstUserList.Filter = adFilterNone
    Set DataGrid1.DataSource = rstUserList
    BusySystemIndicator False
    SSTab1.Tab = 0
    If Not (rstUserList.EOF Or rstUserList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
    rstUserList.ActiveConnection = Nothing
    SetButtonsForNoRecord
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    Call CloseForm(FrmUserMaster)
End Sub
Private Sub Form_Activate()
    EnableChildMenu
    Text1.SetFocus
End Sub
Private Sub Form_Deactivate()
    DisableChildMenu
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyEscape Then
       If SSTab1.Tab = 0 Then
          Call CloseForm(FrmUserMaster)
       Else
           If Toolbar1.Buttons.Item(1).Enabled Then
              SSTab1.Tab = 0
           Else
              If MsgBox("Are you sure to Quit?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Quit !") <> vbYes Then
                   Me.ActiveControl.SetFocus
              Else
                 Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
              End If
           End If
       End If
       KeyCode = 0
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
       Toolbar1_ButtonClick Toolbar1.Buttons.Item(4)
       KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyF5 And Toolbar1.Buttons.Item(6).Enabled Then
       Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
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
           Sendkeys "{TAB}"
        End If
       KeyCode = 0
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Toolbar1.Buttons.Item(4).Enabled Then
        Call Form_KeyDown(vbKeyEscape, 0)
        Cancel = 1
    Else
        Call CloseForm(FrmUserMaster)
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CloseRecordset(rstUserList)
    Call CloseRecordset(rstUserMaster)
    Call CloseConnection(CxnUserMaster)
    Call CloseRecordset(rstUserChild)
    Set oEncrypt = Nothing
    ShowProgressInStatusBar False
    DisableChildMenu
End Sub
Private Sub Text1_Change()
    If rstUserList.RecordCount = 0 Then Exit Sub
    rstUserList.MoveFirst
    If Text1.Text <> "" Then
        rstUserList.Find "[Name] Like '" & FixQuote(Text1.Text) & "%'"
        If rstUserList.EOF Then
            rstUserList.MoveFirst
            If PrevStr <> "" And Len(Text1.Text) > 1 Then
                If dblBookMark <> 0 Then
                    rstUserList.Bookmark = dblBookMark
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
    If Not (rstUserList.EOF Or rstUserList.BOF) Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
            .Add DataGrid1.Bookmark
        End With
    End If
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim KeyProcessed As Boolean
    
    If rstUserList.RecordCount = 0 Then Exit Sub
    If Shift = 0 And KeyCode = vbKeyUp Then
        With rstUserList
            .MovePrevious
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyBack Then
        With rstUserList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyDown Then
        With rstUserList
            .MoveNext
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageUp Then
        With rstUserList
            .Move (-1) * (DataGrid1.VisibleRows - 1)
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageUp Then
        With rstUserList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageDown Then
        With rstUserList
            .Move DataGrid1.VisibleRows - 1
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageDown Then
        With rstUserList
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
    On Error Resume Next
    
    If Toolbar1.Buttons.Item(1).Enabled Then
        If SSTab1.Tab >= 1 Then
           ViewRecord
        Else
            If Not (rstUserList.EOF Or rstUserList.BOF) Then
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
        If SSTab1.Tab = 1 Then
            Mh3dFrame2.Enabled = True
            Mh3dFrame4.Enabled = False
            Text2.SetFocus
        Else
            Mh3dFrame2.Enabled = False
            Mh3dFrame4.Enabled = True
            Check1.SetFocus
        End If
    End If
End Sub
Public Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim UpdateFlag As Integer
    Dim HiLiteRecord As Boolean
    
    If Button.Index = 1 Then
        If rstUserMaster.State = adStateOpen Then
           rstUserMaster.Close
        End If
        rstUserMaster.Open "Select * From UserMaster Where Code = ''", CxnUserMaster, adOpenKeyset, adLockOptimistic
        ClearFields ("P")
        ClearFields ("C")
        Call LoadPrivileges("")
        If rstUserChild.State = adStateClosed Then
            SSTab1.Tab = 0
            Exit Sub
        End If
        If AddRecord(rstUserMaster) Then
            Call SetButtons(False)
            Mh3dFrame3.Enabled = True
            SSTab1.Tab = 1
            Text2.SetFocus
            blnRecordExist = False
        End If
    ElseIf Button.Index = 2 Then
        If rstUserList.RecordCount = 0 Then Exit Sub
        SSTab1.Tab = 1
        EditRecord
    ElseIf Button.Index = 3 Then
        If rstUserList.RecordCount = 0 Then Exit Sub
        SSTab1.Tab = 1
        If rstUserMaster.Fields("Level").Value = "1" Then
           Call DisplayError("Administrator Account cann't be Deleted")
        ElseIf MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            On Error Resume Next
            MdiMainMenu.MousePointer = vbHourglass
            CxnUserMaster.Execute "DELETE FROM UserMaster WHERE Code = '" & rstUserList.Fields("Code").Value & "'"
            MdiMainMenu.MousePointer = vbNormal
            If Err.Number = 0 Then
                rstUserList.Delete
                rstUserList.MoveNext
                If rstUserList.RecordCount > 0 And rstUserList.EOF Then
                    rstUserList.MoveLast
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
        SaveFields
        UpdateFlag = 0
        If UpdateRecord(rstUserMaster) Then
            If UpdatePrivileges("D") Then
                UpdateFlag = 1
                If Not UpdatePrivileges("U") Then
                     UpdateFlag = 0
                End If
            End If
        End If
        If UpdateFlag Then
            AddToList
            If rstUserMaster.State = adStateOpen Then
                rstUserMaster.Close
            End If
            rstUserMaster.CursorLocation = adUseClient
            Call SetButtons(True)
            SSTab1.Tab = 0
            ShowProgressInStatusBar True
            Timer1.Enabled = True
        Else
            DisplayError ("Failed to save the record")
            Toolbar1_ButtonClick Toolbar1.Buttons.Item(5)
        End If
    ElseIf Button.Index = 5 Then
        If CancelRecordUpdate(rstUserMaster) Then
           If rstUserMaster.State = adStateOpen Then
              rstUserMaster.Close
           End If
           rstUserMaster.CursorLocation = adUseClient
           Call SetButtons(True)
           SetButtonsForNoRecord
           SSTab1.Tab = 0
        End If
    ElseIf Button.Index = 6 Then
        SSTab1.Tab = 0
        Set DataGrid1.DataSource = Nothing
        rstUserList.ActiveConnection = CxnUserMaster
        Do While Not RefreshRecord(rstUserList)
        Loop
        Set DataGrid1.DataSource = rstUserList
        rstUserList.ActiveConnection = Nothing
        HiLiteRecord = True
    ElseIf Button.Index = 7 Then
        SSTab1.Tab = 0
        With FrmFilter
            .Combo1.AddItem "Name", 0
            .Combo1.ListIndex = 0
            Set .srcForm = Me
            .Show vbModal
        End With
        HiLiteRecord = True
    ElseIf Button.Index = 13 Then
        If rstUserList.RecordCount > 0 Then rstUserList.MoveFirst
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 14 Then
        If rstUserList.RecordCount > 0 Then
           rstUserList.MovePrevious
           If rstUserList.BOF Then
              rstUserList.MoveNext
           End If
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 15 Then
        If rstUserList.RecordCount > 0 Then
           rstUserList.MoveNext
           If rstUserList.EOF Then
              rstUserList.MovePrevious
           End If
        End If
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 16 Then
        If rstUserList.RecordCount > 0 Then rstUserList.MoveLast
        HiLiteRecord = True
        ViewRecord
    ElseIf Button.Index = 18 Then
        Call CloseForm(FrmUserMaster)
        HiLiteRecord = False
    End If
    If HiLiteRecord Then
        If Not (rstUserList.EOF Or rstUserList.BOF) Then
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
Private Sub SetButtons(bVal As Boolean)
    Toolbar1.Buttons.Item(1).Enabled = bVal
    Toolbar1.Buttons.Item(2).Enabled = bVal
    Toolbar1.Buttons.Item(3).Enabled = bVal
    Toolbar1.Buttons.Item(4).Enabled = Not bVal
    Toolbar1.Buttons.Item(5).Enabled = Not bVal
    Toolbar1.Buttons.Item(6).Enabled = bVal
    Toolbar1.Buttons.Item(7).Enabled = bVal
    Toolbar1.Buttons.Item(13).Enabled = bVal
    Toolbar1.Buttons.Item(14).Enabled = bVal
    Toolbar1.Buttons.Item(15).Enabled = bVal
    Toolbar1.Buttons.Item(16).Enabled = bVal
    Toolbar1.Buttons.Item(18).Enabled = bVal
    Mh3dFrame2.Enabled = Not bVal
    Mh3dFrame4.Enabled = False
End Sub
Private Sub SetButtonsForNoRecord()
    If rstUserList.RecordCount = 0 Then
        Toolbar1.Buttons.Item(2).Enabled = False
        Toolbar1.Buttons.Item(3).Enabled = False
        Toolbar1.Buttons.Item(13).Enabled = False
        Toolbar1.Buttons.Item(14).Enabled = False
        Toolbar1.Buttons.Item(15).Enabled = False
        Toolbar1.Buttons.Item(16).Enabled = False
    End If
End Sub
Private Sub Text2_Validate(Cancel As Boolean)
    If rstUserMaster.EOF Or rstUserMaster.BOF Then Exit Sub
    If CheckEmpty(Text2, True) Then
        Cancel = True
    ElseIf CheckDuplicate(CxnUserMaster, "UserMaster", "Code", "Name", Text2.Text, rstUserMaster.Fields("Code").Value, False) Then
        Cancel = True
    ElseIf CheckEmpty(Text3, False) Then
        Text3.Text = Text2.Text
    End If
End Sub
Private Sub Text4_Validate(Cancel As Boolean)
    If CheckEmpty(Text4, True) Then
        Cancel = True
    End If
End Sub
Private Sub Text5_Validate(Cancel As Boolean)
    If CheckEmpty(Text5, True) Then
        Cancel = True
    ElseIf UCase(Trim(Text4.Text)) <> UCase(Trim(Text5.Text)) Then
        Call DisplayError("Password Mismatch")
    End If
End Sub
Private Sub Text4_GotFocus()
    Text4.SelStart = 0
    Text4.SelLength = Len(Text4.Text)
End Sub
Private Sub Text5_GotFocus()
    Text5.SelStart = 0
    Text5.SelLength = Len(Text5.Text)
End Sub
Private Sub ViewRecord()
    ClearFields ("P")
    ClearFields ("C")
    If rstUserList.EOF Then Exit Sub
    FindRecord
    LoadFields
End Sub
Private Sub FindRecord()
    If rstUserMaster.State = adStateOpen Then
       rstUserMaster.Close
    End If
    rstUserMaster.Open "Select * From UserMaster Where Code = '" & FixQuote(rstUserList.Fields("Code").Value) & "'", CxnUserMaster, adOpenKeyset, adLockOptimistic
    If rstUserMaster.RecordCount = 0 Then
       Call DisplayError("This Record has been deleted by some Other User ! Click Ok To Refresh the List")
       Toolbar1_ButtonClick Toolbar1.Buttons.Item(6)
    End If
End Sub
Private Sub ClearFields(ByVal strType As String)
    If strType = "P" Then
        Text2.Text = ""
        Text3.Text = ""
        Text4.Text = ""
        Text5.Text = ""
        Option1.Value = False
        Option2.Value = True
        Option3.Value = False
    Else
        Check1.Value = 0
        Check2.Value = 0
        Check3.Value = 0
        Check4.Value = 0
        TreeView1.Nodes.Clear
    End If
End Sub
Private Sub LoadFields()
    If rstUserMaster.EOF Or rstUserMaster.BOF Then Exit Sub
    Text2.Text = rstUserMaster.Fields("Name").Value
    Text3.Text = rstUserMaster.Fields("PrintName").Value
    Text4.Text = oEncrypt.DecryptString(rstUserMaster.Fields("Password").Value)
    Text5.Text = oEncrypt.DecryptString(rstUserMaster.Fields("Password").Value)
    If rstUserMaster.Fields("Level").Value = "1" Then
       Option1.Value = True
    ElseIf rstUserMaster.Fields("Level").Value = "2" Then
       Option2.Value = True
    ElseIf rstUserMaster.Fields("Level").Value = "3" Then
       Option3.Value = True
    End If
    Check1.Value = Val(rstUserMaster.Fields("AllowMastersModification").Value)
    Check2.Value = Val(rstUserMaster.Fields("AllowMastersDeletion").Value)
    Check3.Value = Val(rstUserMaster.Fields("AllowTransactionsModification").Value)
    Check4.Value = Val(rstUserMaster.Fields("AllowTransactionsDeletion").Value)
    Call LoadPrivileges(rstUserMaster.Fields("Code").Value)
End Sub
Private Sub EditRecord()
    On Error GoTo ErrorHandler
    
    If rstUserMaster.RecordCount = 0 Then Exit Sub
    If rstUserMaster.Fields("Level").Value <> "3" And UserLevel = "2" Then
       Call DisplayError("You cann't Modify this Account")
       Exit Sub
    End If
    If rstUserChild.State = adStateClosed Then
        SSTab1.Tab = 0
        Exit Sub
    End If
    If rstUserMaster.State = adStateOpen Then
       rstUserMaster.Close
    End If
    rstUserMaster.CursorLocation = adUseServer
    rstUserMaster.Open "Select * From UserMaster Where Code = '" & FixQuote(rstUserList.Fields("Code").Value) & "'", CxnUserMaster, adOpenKeyset, adLockPessimistic
    MdiMainMenu.MousePointer = vbHourglass
    rstUserMaster.Fields("Printstatus") = "N"
    MdiMainMenu.MousePointer = vbNormal
    AddToList
    Call SetButtons(False)
    If rstUserMaster.Fields("Level").Value = "1" Then
       Mh3dFrame3.Enabled = False
    Else
       Mh3dFrame3.Enabled = True
    End If
    SSTab1.TabEnabled(0) = False
    Text2.SetFocus
    blnRecordExist = True
    Exit Sub
ErrorHandler:
    If Err.Number = -2147467259 Then
       Call DisplayError("Failed to Edit the record")
    End If
    MdiMainMenu.MousePointer = vbNormal
    SSTab1.Tab = 0
End Sub
Private Sub SaveFields()
    If rstUserMaster.EOF Or rstUserMaster.BOF Then Exit Sub
    If Not blnRecordExist Then
        rstUserMaster.Fields("Code").Value = GenerateCode(CxnUserMaster, "Select Max(Code) From UserMaster", 6, "0")
    End If
    rstUserMaster.Fields("Name").Value = Trim(Text2.Text)
    rstUserMaster.Fields("PrintName").Value = Trim(Text3.Text)
    rstUserMaster.Fields("Password").Value = oEncrypt.EncryptString(Trim(Text4.Text))
    If Option1.Value = True Then
       rstUserMaster.Fields("Level").Value = "1"
    ElseIf Option2.Value = True Then
       rstUserMaster.Fields("Level").Value = "2"
    ElseIf Option3.Value = True Then
       rstUserMaster.Fields("Level").Value = "3"
    End If
    rstUserMaster.Fields("AllowMastersModification").Value = IIf(Option1.Value, 1, IIf(Check1.Value, 1, 0))
    rstUserMaster.Fields("AllowMastersDeletion").Value = IIf(Option1.Value, 1, IIf(Check2.Value, 1, 0))
    rstUserMaster.Fields("AllowTransactionsModification").Value = IIf(Option1.Value, 1, IIf(Check3.Value, 1, 0))
    rstUserMaster.Fields("AllowTransactionsDeletion").Value = IIf(Option1.Value, 1, IIf(Check4.Value, 1, 0))
    rstUserMaster.Fields("PrintStatus").Value = "N"
End Sub
Private Sub AddToList()
    On Error Resume Next
    
    rstUserList.MoveFirst
    rstUserList.Find "[Code] = '" & rstUserMaster.Fields("Code").Value & "'"
    If rstUserList.EOF Then
       rstUserList.AddNew
       rstUserList.Fields("Code").Value = rstUserMaster.Fields("Code").Value
    End If
    rstUserList.Fields("Name").Value = rstUserMaster.Fields("Name").Value
    rstUserList.Update
    rstUserList.Sort = "Name Asc"
    rstUserList.Find "[Code] = '" & rstUserMaster.Fields("Code").Value & "'"
End Sub
Private Function CheckMandatoryFields() As Boolean
    If CheckEmpty(Text2.Text, False) Then
        SSTab1.Tab = 1
        Text2.SetFocus
        CheckMandatoryFields = True
    ElseIf CheckEmpty(Text3.Text, False) Then
        SSTab1.Tab = 1
        Text3.SetFocus
        CheckMandatoryFields = True
    ElseIf CheckEmpty(Text4.Text, False) Then
        SSTab1.Tab = 1
        Text4.SetFocus
        CheckMandatoryFields = True
    ElseIf UCase(Trim(Text4.Text)) <> UCase(Trim(Text5.Text)) Then
        SSTab1.Tab = 1
        Call DisplayError("Password Mismatch")
        Text4.SetFocus
        CheckMandatoryFields = True
    ElseIf CheckDuplicate(CxnUserMaster, "UserMaster", "Code", "Name", Text2.Text, rstUserMaster.Fields("Code").Value, False) Then
        SSTab1.Tab = 1
        Text2.SetFocus
        CheckMandatoryFields = True
    End If
End Function
Private Sub LoadPrivileges(ByVal strUserCode As String)
    Dim Cnt As Integer
    On Error GoTo ErrorHandler
    
    Call CreateTreeView
    If rstUserChild.State = adStateOpen Then
       rstUserChild.Close
    End If
    rstUserChild.Open "Select [Module] From UserChild Where Code = '" & FixQuote(strUserCode) & "'", CxnUserMaster, adOpenKeyset, adLockOptimistic
    rstUserChild.ActiveConnection = Nothing
    For Cnt = 1 To TreeView1.Nodes.Count
        If rstUserChild.RecordCount = 0 Then
            TreeView1.Nodes(Cnt).Checked = True
        Else
            rstUserChild.MoveFirst
            rstUserChild.Find "[Module] = '" & Mid(TreeView1.Nodes(Cnt).Key, 4, 4) & "'"
            If Not rstUserChild.EOF Then
                TreeView1.Nodes(Cnt).Checked = True
            End If
        End If
    Next
    Exit Sub
ErrorHandler:
    DisplayError ("Failed to Load Privileges")
End Sub
Private Function UpdatePrivileges(ByVal strOption As String) As Boolean
    Dim Cnt As Integer
    Dim ChildNodeSelected As Boolean
    On Error GoTo ErrorHandler
    
    UpdatePrivileges = True
    If strOption = "D" Then
        CxnUserMaster.Execute "Delete From UserChild Where Code = '" & rstUserMaster.Fields("Code").Value & "'"
    Else
        If rstUserMaster.Fields("Level").Value <> "1" Then
            For Cnt = TreeView1.Nodes.Count To 1 Step -1
                If InStr(1, "Key0100_Key0200_Key0300", TreeView1.Nodes(Cnt).Key) Then
                    If Not ChildNodeSelected Then
                        TreeView1.Nodes(Cnt).Checked = False
                    End If
                    ChildNodeSelected = False
                ElseIf TreeView1.Nodes(Cnt).Checked Then    'For Child Nodes
                    ChildNodeSelected = True
                End If
                If TreeView1.Nodes(Cnt).Checked Then
                    CxnUserMaster.Execute "Insert Into UserChild Values ('" & rstUserMaster.Fields("Code").Value & "','" & Mid(TreeView1.Nodes(Cnt).Key, 4, 4) & "')"
                End If
            Next
        End If
    End If
    Exit Function
ErrorHandler:
    UpdatePrivileges = False
End Function
Private Sub CreateTreeView()
    Dim ParentNodeKey As String
    Dim nodX As Node
    Dim Object As Object
    On Error GoTo ErrorHandler
    
    Screen.MousePointer = vbHourglass
    For Each Object In MdiMainMenu
        If TypeName(Object) = "Menu" Then
            If Object.Tag <> "" Then
                If InStr(1, "0100_0200_0300", Object.Tag) Then
                    ParentNodeKey = "Key" & Trim(Object.Tag)
                    Set nodX = TreeView1.Nodes.Add(, , ParentNodeKey, Mid(Object.Caption, 2))
                ElseIf Object.Tag <> "0112" Then
                    Set nodX = TreeView1.Nodes.Add(ParentNodeKey, tvwChild, "Key" & Trim(Object.Tag), Object.Caption)
                End If
            End If
        End If
    Next
    nodX.Expanded = False
    Set nodX = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub
ErrorHandler:
    Set nodX = Nothing
    TreeView1.Nodes.Clear
    Screen.MousePointer = vbDefault
End Sub
Private Sub TreeView1_Expand(ByVal Node As MSComctlLib.Node)
    If Not Node.Checked Then
        Node.Expanded = False
    End If
End Sub
Private Sub TreeView1_NodeCheck(ByVal Node As MSComctlLib.Node)
    Dim Cnt As Integer
    
    If InStr(1, "Key0100_Key0200_Key0300", Node.Key) Then
        If Not Node.Checked Then
            For Cnt = Node.Index + 1 To Node.Index + Node.Children
                TreeView1.Nodes(Cnt).Checked = False
            Next
            Node.Expanded = False
        End If
    End If
End Sub
Public Sub FilterRecord(ByVal SrchFor As String, ByVal SrchText As String)
    If SrchFor = "Name" Then
        rstUserList.Filter = "[Name] Like '%" & SrchText & "%'"
    End If
End Sub
Private Sub Timer1_Timer()
    On Error Resume Next
    
    MdiMainMenu.ProgressBar1.Value = MdiMainMenu.ProgressBar1.Value + 10
    If MdiMainMenu.ProgressBar1.Value = 100 Then
       Timer1.Enabled = False
       ShowProgressInStatusBar False
    End If
End Sub
