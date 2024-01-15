VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Begin VB.Form FrmNumberRange 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Range"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4440
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "NumberRange.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "FrmLogin"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   4440
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3960
      Picture         =   "NumberRange.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Cancel"
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton cmdProceed 
      Height          =   375
      Left            =   3960
      Picture         =   "NumberRange.frx":0544
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Proceed"
      Top             =   120
      Width           =   375
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   1260
      Left            =   120
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   3735
      _Version        =   65536
      _ExtentX        =   6588
      _ExtentY        =   2222
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
      Picture         =   "NumberRange.frx":0B02
      Begin VB.TextBox Text2 
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
         IMEMode         =   3  'DISABLE
         Left            =   1320
         MaxLength       =   60
         TabIndex        =   1
         Top             =   435
         Width           =   2310
      End
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
         IMEMode         =   3  'DISABLE
         Left            =   1320
         MaxLength       =   60
         TabIndex        =   0
         Top             =   120
         Width           =   2310
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
         Height          =   330
         Left            =   120
         TabIndex        =   5
         Top             =   120
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
         Caption         =   " Starting No."
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   -2147483634
         Picture         =   "NumberRange.frx":0B1E
         Picture         =   "NumberRange.frx":0B3A
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
         Height          =   330
         Left            =   120
         TabIndex        =   6
         Top             =   435
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
         Caption         =   " Ending No."
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   -2147483634
         Picture         =   "NumberRange.frx":0B56
         Picture         =   "NumberRange.frx":0B72
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
         Height          =   380
         Left            =   120
         TabIndex        =   8
         Top             =   750
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   670
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
         Caption         =   " Print"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   -2147483634
         Picture         =   "NumberRange.frx":0B8E
         Picture         =   "NumberRange.frx":0BAA
      End
      Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame3 
         Height          =   380
         Left            =   1320
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   750
         Width           =   2310
         _Version        =   65536
         _ExtentX        =   4075
         _ExtentY        =   670
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
         Picture         =   "NumberRange.frx":0BC6
         Begin VB.ComboBox Combo1 
            Height          =   330
            ItemData        =   "NumberRange.frx":0BE2
            Left            =   0
            List            =   "NumberRange.frx":0BE4
            TabIndex        =   10
            Top             =   30
            Width           =   1935
         End
         Begin VB.CheckBox Check2 
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
            Height          =   345
            Left            =   1980
            TabIndex        =   2
            Top             =   20
            Width           =   225
         End
      End
   End
End
Attribute VB_Name = "FrmNumberRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Report As Integer
Private Sub Form_Load()

If Report = 1 Then
    Combo1.AddItem "2-Title Printing Order"
    Combo1.AddItem "3-Title Lamination Order"
    Combo1.AddItem "4-Book Binding Order"
ElseIf Report = 2 Then
    Combo1.AddItem "1-Book Printing Order"
    Combo1.AddItem "3-Title Lamination Order"
    Combo1.AddItem "4-Book Binding Order"
    
ElseIf Report = 3 Then
    Combo1.AddItem "2-Title Printing Order"
ElseIf Report = 4 Then
    Combo1.AddItem "1-Book Printing Order"
Else
    Combo1.AddItem ""
End If
CenterForm Me
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}", True
        KeyCode = 0
    ElseIf KeyCode = vbKeyEscape Then
        KeyCode = 0
        
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then cmdCancel_Click
End Sub
Private Sub cmdProceed_Click()
    Me.Hide
End Sub
Private Sub cmdCancel_Click()
    Text1.Text = "": Text2.Text = ""
    Combo1.Text = ""
    Me.Hide
End Sub
