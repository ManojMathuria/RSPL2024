VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmCorrectionRegister 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Correction Register"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9735
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CorrectionRegister.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "FrmLogin"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   9735
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   9240
      Picture         =   "CorrectionRegister.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Exit"
      Top             =   105
      Width           =   375
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
      Height          =   5050
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   105
      Width           =   9015
      _Version        =   65536
      _ExtentX        =   15901
      _ExtentY        =   8908
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
      Picture         =   "CorrectionRegister.frx":0544
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
         Left            =   1200
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   105
         Width           =   7690
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
         Height          =   330
         Left            =   120
         TabIndex        =   4
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
         Caption         =   " Book Name"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "CorrectionRegister.frx":0560
         Picture         =   "CorrectionRegister.frx":057C
      End
      Begin FPSpreadADO.fpSpread fpSpread3 
         Height          =   4305
         Left            =   120
         TabIndex        =   0
         Top             =   630
         Width           =   8775
         _Version        =   524288
         _ExtentX        =   15478
         _ExtentY        =   7594
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
         SpreadDesigner  =   "CorrectionRegister.frx":0598
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   9000
         Y1              =   530
         Y2              =   530
      End
   End
End
Attribute VB_Name = "FrmCorrectionRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
    CenterForm Me
    DisableCloseButton Me
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
        KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyEscape Then
        cmdExit_Click
        KeyCode = 0
    End If
End Sub
Private Sub cmdExit_Click()
    Me.Hide
End Sub