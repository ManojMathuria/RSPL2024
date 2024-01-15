VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Begin VB.Form FrmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login Please..."
   ClientHeight    =   1590
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
   Icon            =   "Login.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "FrmLogin"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   4440
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3960
      Picture         =   "Login.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Exit"
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton cmdLogin 
      Height          =   375
      Left            =   3960
      Picture         =   "Login.frx":0544
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Login"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtUserName 
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
      Left            =   1440
      MaxLength       =   40
      TabIndex        =   1
      Top             =   240
      Width           =   2295
   End
   Begin VB.TextBox txtPassword 
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
      Left            =   1440
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   560
      Width           =   2295
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   1200
      Left            =   120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
      Width           =   3735
      _Version        =   65536
      _ExtentX        =   6588
      _ExtentY        =   2117
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
      Picture         =   "Login.frx":09C0
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
         Height          =   330
         Left            =   120
         TabIndex        =   0
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
         Caption         =   " &User Name"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   -2147483634
         Picture         =   "Login.frx":09DC
         Picture         =   "Login.frx":09F8
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
         Height          =   330
         Left            =   120
         TabIndex        =   2
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
         Caption         =   " &Password"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   -2147483634
         Picture         =   "Login.frx":0A14
         Picture         =   "Login.frx":0A30
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
         Height          =   330
         Left            =   120
         TabIndex        =   7
         Top             =   750
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
         Caption         =   " System Date"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   -2147483634
         Picture         =   "Login.frx":0A4C
         Picture         =   "Login.frx":0A68
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
         Height          =   330
         Left            =   2280
         TabIndex        =   8
         Top             =   750
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
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
         Caption         =   " Time"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   -2147483634
         Picture         =   "Login.frx":0A84
         Picture         =   "Login.frx":0AA0
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
         Height          =   330
         Left            =   1320
         TabIndex        =   9
         Top             =   750
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   582
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TintColor       =   16711935
         Caption         =   ""
         Alignment       =   0
         FillColor       =   16777215
         Picture         =   "Login.frx":0ABC
         Picture         =   "Login.frx":0AD8
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
         Height          =   330
         Left            =   2880
         TabIndex        =   10
         Top             =   750
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   582
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TintColor       =   16711935
         Caption         =   ""
         Alignment       =   0
         FillColor       =   16777215
         Picture         =   "Login.frx":0AF4
         Picture         =   "Login.frx":0B10
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   400
      Left            =   1440
      Top             =   120
   End
   Begin VB.Label Label3 
      Caption         =   "Change Password"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      MouseIcon       =   "Login.frx":0B2C
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   1320
      Width           =   1455
   End
End
Attribute VB_Name = "FrmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstUserMaster As New ADODB.Recordset
Dim oEncrypt As New clsBlowFish
Dim oRegistry As New clsRegistry
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    
    Mh3dLabel5.Caption = Format(Date, "dd-mmm-yy")
    Mh3dLabel6.Caption = Format(Time, "hh:mm:ss")
    CenterForm Me
    txtUserName.Text = oRegistry.GetRegistryValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Saral", "Last User", "")
    LoginSuccess = False
    
    Exit Sub
    
ErrorHandler:
    Call CloseForm(FrmLogin)
End Sub
Private Sub Form_Activate()
    If txtUserName.Text <> "" Then
       txtPassword.SetFocus
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       Sendkeys "{TAB}", True
       KeyCode = 0
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    BusySystemIndicator False
    Set oRegistry = Nothing
    Set oEncrypt = Nothing
    Call CloseRecordset(rstUserMaster)
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        Call CloseForm(FrmLogin)
    End If
End Sub
Private Sub cmdLogin_Click()
    If LoginSuccess Then Exit Sub
    Dim IsUserNameOk As Boolean, IsPasswordOk As Boolean
    On Error GoTo ErrorHandler
    cmdLogin.Picture = LoadPicture(App.Path & "\Icon\Yellow.Bmp")
    rstUserMaster.Open "Select Code, Name, Password, [Level], AllowMastersModification, AllowMastersDeletion, AllowTransactionsModification, AllowTransactionsDeletion From UserMaster Where Name = '" & Trim(txtUserName.Text) & "'", CxnDatabase, adOpenKeyset, adLockReadOnly
    If rstUserMaster.RecordCount <> 0 Then
        IsUserNameOk = True
        If UCase(oEncrypt.DecryptString(Trim(rstUserMaster.Fields("Password")))) = UCase(Trim(txtPassword.Text)) Then IsPasswordOk = True
    End If
    If Not IsUserNameOk Then
       cmdLogin.Picture = LoadPicture(App.Path & "\Icon\Red.Bmp")
       txtUserName.SetFocus
    ElseIf Not IsPasswordOk Then
       cmdLogin.Picture = LoadPicture(App.Path & "\Icon\Red.Bmp")
       txtPassword.SetFocus
    Else
       LoginSuccess = True
       Me.Caption = "Login Successful !"
       cmdCancel.ToolTipText = "Proceed"
       UserCode = RTrim(rstUserMaster.Fields("Code").Value)
       UserName = RTrim(rstUserMaster.Fields("Name").Value)
       UserLevel = RTrim(rstUserMaster.Fields("Level").Value)
       AllowMastersModification = Val(rstUserMaster.Fields("AllowMastersModification").Value)
       AllowMastersDeletion = Val(rstUserMaster.Fields("AllowMastersDeletion").Value)
       AllowTransactionsModification = Val(rstUserMaster.Fields("AllowTransactionsModification").Value)
       AllowTransactionsDeletion = Val(rstUserMaster.Fields("AllowTransactionsDeletion").Value)
       Call oRegistry.SetRegistryValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Saral", "Last User", Trim(txtUserName.Text))
       cmdLogin.Picture = LoadPicture(App.Path & "\Icon\Green.Bmp")
       cmdCancel.Picture = LoadPicture(App.Path & "\Icon\Run.Bmp")
       cmdCancel.SetFocus
    End If
    rstUserMaster.Close
    Exit Sub
ErrorHandler:
    If rstUserMaster.State = adStateOpen Then
         rstUserMaster.Close
    End If
End Sub
Private Sub cmdCancel_Click()
     Call CloseForm(FrmLogin)
End Sub
Private Sub txtUserName_GotFocus()
    txtUserName.SelStart = 0
    txtUserName.SelLength = Len(txtUserName.Text)
End Sub
Private Sub txtUserName_Validate(Cancel As Boolean)
    If CheckEmpty(txtUserName, False) Then
        Cancel = True
    End If
End Sub
Private Sub txtPassword_GotFocus()
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword.Text)
End Sub
Private Sub Timer1_Timer()
    Static Ticks As Integer
    
    Mh3dLabel5.Caption = Format(Date, "dd-mmm-yy")
    Mh3dLabel6.Caption = Format(Time, "hh:mm:ss")
    If FrmLogin.Caption = " " Then
       If LoginSuccess Then
          FrmLogin.Caption = "Login Successful !"
       Else
          FrmLogin.Caption = "Login Please..."
          Ticks = 0
       End If
    Else
        FrmLogin.Caption = " "
    End If
    If LoginSuccess Then
       Ticks = Ticks + 1
       If Ticks >= 5 Then
          Call CloseForm(FrmLogin)
       End If
    Else
        Ticks = 0
    End If
End Sub
Private Sub Label3_Click()
    If LoginSuccess Then Exit Sub
    On Error GoTo ErrorHandler
    
    rstUserMaster.Open "Select Name From UserMaster Where Name = '" & Trim(txtUserName.Text) & "'", CxnDatabase, adOpenKeyset, adLockReadOnly
    txtPassword.Text = ""
    If rstUserMaster.RecordCount = 0 Then
        txtUserName.SelStart = 0
        txtUserName.SelLength = Len(txtUserName.Text)
    Else
        Load FrmChangePassword
        FrmChangePassword.User = txtUserName
        FrmChangePassword.Show vbModal
    End If
    rstUserMaster.Close
    Exit Sub
ErrorHandler:
    If rstUserMaster.State = adStateOpen Then
         rstUserMaster.Close
    End If
End Sub
