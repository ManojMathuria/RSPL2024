VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Begin VB.Form FrmChangePassword 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Password..."
   ClientHeight    =   1410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4200
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ChangePassword.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "FrmLogin"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   4200
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtOldPassword 
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
      IMEMode         =   3  'DISABLE
      Left            =   1800
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox txtNewPassword 
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
      IMEMode         =   3  'DISABLE
      Left            =   1800
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   560
      Width           =   1695
   End
   Begin VB.CommandButton cmdChange 
      Height          =   375
      Left            =   3720
      Picture         =   "ChangePassword.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Change"
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3720
      Picture         =   "ChangePassword.frx":08BE
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Cancel"
      Top             =   480
      Width           =   375
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   1170
      Left            =   120
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   120
      Width           =   3495
      _Version        =   65536
      _ExtentX        =   6165
      _ExtentY        =   2064
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
      Picture         =   "ChangePassword.frx":09C0
      Begin VB.TextBox txtConfirmPassword 
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
         IMEMode         =   3  'DISABLE
         Left            =   1680
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   755
         Width           =   1695
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
         Height          =   330
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TintColor       =   16711935
         Caption         =   " &Old Password"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   -2147483634
         Picture         =   "ChangePassword.frx":09DC
         Picture         =   "ChangePassword.frx":09F8
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
         Height          =   330
         Left            =   120
         TabIndex        =   2
         Top             =   440
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TintColor       =   16711935
         Caption         =   " &New Password"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   -2147483634
         Picture         =   "ChangePassword.frx":0A14
         Picture         =   "ChangePassword.frx":0A30
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
         Height          =   330
         Left            =   120
         TabIndex        =   4
         Top             =   755
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   582
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TintColor       =   16711935
         Caption         =   " &Confirm Password"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   -2147483634
         Picture         =   "ChangePassword.frx":0A4C
         Picture         =   "ChangePassword.frx":0A68
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   400
      Left            =   1440
      Top             =   120
   End
End
Attribute VB_Name = "FrmChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public User As String
Dim rstUserMaster As New ADODB.Recordset
Dim oEncrypt As New clsBlowFish
Dim ChangeSuccess As Boolean
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    
    CenterForm Me
    ChangeSuccess = False
    rstUserMaster.CursorLocation = adUseServer
    Exit Sub
ErrorHandler:
    Call CloseForm(FrmChangePassword)
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       SendKeys "{TAB}", True
       KeyCode = 0
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    BusySystemIndicator False
    Call CloseRecordset(rstUserMaster)
    Set oEncrypt = Nothing
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        Call CloseForm(FrmChangePassword)
    End If
End Sub
Private Sub cmdChange_Click()
    If ChangeSuccess Then Exit Sub
    Dim UpdateFlag As Integer
    Dim IsOldPasswordOk As Boolean, IsNewPasswordOk As Boolean
    On Error GoTo ErrorHandler
    
    If rstUserMaster.State = adStateClosed Then
         rstUserMaster.Open "Select * From UserMaster Where Name = '" & FixQuote(User) & "'", CxnDatabase, adOpenKeyset, adLockPessimistic
    End If
    rstUserMaster.Fields("Printstatus").Value = "N"
    cmdChange.Picture = LoadPicture(App.Path & "\Icon\Yellow.Bmp")
    If UCase(oEncrypt.DecryptString(Trim(rstUserMaster.Fields("Password")))) = UCase(Trim(txtOldPassword.Text)) Then
       IsOldPasswordOk = True
    End If
    If UCase(Trim(txtNewPassword.Text)) = UCase(Trim(txtConfirmPassword.Text)) Then
       IsNewPasswordOk = True
    End If
    If Not IsOldPasswordOk Then
       Call CancelRecordUpdate(rstUserMaster)
       cmdChange.Picture = LoadPicture(App.Path & "\Icon\Red.Bmp")
       txtOldPassword.SetFocus
    ElseIf Not IsNewPasswordOk Then
       Call DisplayError("Password Mismatch")
       Call CancelRecordUpdate(rstUserMaster)
       cmdChange.Picture = LoadPicture(App.Path & "\Icon\Red.Bmp")
       txtNewPassword.SetFocus
    Else
       MdiMainMenu.MousePointer = vbHourglass
       MdiMainMenu.MousePointer = vbNormal
       rstUserMaster.Fields("Password").Value = oEncrypt.EncryptString(Trim(txtNewPassword.Text))
       If UpdateRecord(rstUserMaster) Then
            ChangeSuccess = True
            Me.Caption = "Password Change Successful !"
            cmdCancel.ToolTipText = "Proceed"
            cmdChange.Picture = LoadPicture(App.Path & "\Icon\Green.Bmp")
            cmdCancel.Picture = LoadPicture(App.Path & "\Icon\Run.Bmp")
            cmdCancel.SetFocus
       Else
            Call DisplayError("Failed to change the Password")
            Call CancelRecordUpdate(rstUserMaster)
       End If
    End If
    Exit Sub
ErrorHandler:
    If Err.Number = -2147467259 Then
       Call DisplayError("Failed to change the Password")
    End If
    MdiMainMenu.MousePointer = vbNormal
    If rstUserMaster.State = adStateOpen Then
        rstUserMaster.Close
    End If
End Sub
Private Sub cmdCancel_Click()
     Call CloseForm(FrmChangePassword)
End Sub
Private Sub txtOldPassword_GotFocus()
    txtOldPassword.SelStart = 0
    txtOldPassword.SelLength = Len(txtOldPassword.Text)
End Sub
Private Sub txtNewPassword_GotFocus()
    txtNewPassword.SelStart = 0
    txtNewPassword.SelLength = Len(txtNewPassword.Text)
End Sub
Private Sub txtConfirmPassword_GotFocus()
    txtConfirmPassword.SelStart = 0
    txtConfirmPassword.SelLength = Len(txtConfirmPassword.Text)
End Sub
Private Sub Timer1_Timer()
    Static Ticks As Integer
    
    If FrmChangePassword.Caption = " " Then
       If ChangeSuccess Then
          FrmChangePassword.Caption = "Password Change Successful !"
       Else
          FrmChangePassword.Caption = "Change Password..."
          Ticks = 0
       End If
    Else
        FrmChangePassword.Caption = " "
    End If
    If ChangeSuccess Then
       Ticks = Ticks + 1
       If Ticks >= 5 Then
          Call CloseForm(FrmChangePassword)
       End If
    Else
        Ticks = 0
    End If
End Sub
