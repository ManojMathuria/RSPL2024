VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmSplash 
   BackColor       =   &H00008000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6735
   ControlBox      =   0   'False
   Icon            =   "FrmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Left            =   960
      Top             =   3600
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   1890
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.Timer tmrMain 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   480
      Top             =   3600
   End
   Begin VB.Timer tmrCaption 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   0
      Top             =   3600
   End
   Begin MSForms.Label Label22 
      Height          =   855
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6855
      ForeColor       =   12632319
      BackColor       =   32768
      VariousPropertyBits=   8388627
      Caption         =   "Rachna Sagar Pvt Ltd."
      PicturePosition =   327683
      Size            =   "12091;1508"
      SpecialEffect   =   6
      Picture         =   "FrmSplash.frx":0442
      FontName        =   "Lucida Fax"
      FontEffects     =   1073741827
      FontHeight      =   480
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
      FontWeight      =   600
   End
   Begin MSForms.Label Label33 
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   2270
      Width           =   6855
      ForeColor       =   12648447
      BackColor       =   32768
      VariousPropertyBits=   8388627
      Caption         =   "Copyright © 2016 Rachna Sagar Pvt Ltd. All Rights Reserved."
      PicturePosition =   327683
      Size            =   "12091;661"
      FontName        =   "Lucida Fax"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   600
   End
   Begin VB.Label Label41 
      Caption         =   " Close"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5400
      TabIndex        =   3
      Top             =   1560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   90
      Index           =   0
      Left            =   5010
      TabIndex        =   2
      Top             =   1320
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label Label3 
      Caption         =   " Press Esc -  For Login"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   1560
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   " Double Click for Download New Version"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   450
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   6735
   End
End
Attribute VB_Name = "FrmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim aa As String
Dim Execonn As New ADODB.Connection
Dim Rs As New ADODB.Recordset
Dim oRegistry As New clsRegistry
Dim T As Integer

''Here is where we do all the work
'Public Sub ScrollText()
' Static i As Integer
' Dim k As Integer
' k = i Xor 1 'other label
' 'move the label left by one pixel
' Label1(i).Left = Label1(i).Left - 30
' 'other label follows like a train
' Label1(k).Left = Label1(i).Left + Label1(i).Width
' 'if engine is off screen, then make it caboose
' If Label1(k).Left = 0 Then i = k
'End Sub

'Sub SetText(StrText As String)
'    Label1(0) = StrText & Space(10)
'    Label1(1) = Label1(0)
'    Label1(0).Left = 0
'    Label1(1).Left = Label1(0).Width
'End Sub
'
'Private Sub Command1_Click()
'  SetText "Hello World - CG - VB Programmer's Group"
'  Timer1.Interval = 20
'End Sub

'Private Sub Timer1_Timer()
'    ScrollText
'End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
'Unload Me
Call CloseForm(FrmSplash)
MdiMainMenu.Show
End If
End Sub

Private Sub Form_Load()
Dim SoftwareID As Integer
Dim SoftwareVer As String

Me.Top = 100
Me.Left = 520


Label2.Caption = " Easy Publish Production Management System Version - 20 | 11.17 "
Timer1.Enabled = True
Timer1.Interval = 50


'On Error GoTo ErrorHandler
'
'Call oRegistry.SetRegistryValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Saral", "", "")
'Label2.Caption = " Saral-Production Management System"
'
'SoftwareID = IIf(oRegistry.GetRegistryValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Saral", "Saral ID", "") = "", 0, oRegistry.GetRegistryValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Saral", "Saral ID", ""))
'SoftwareVer = oRegistry.GetRegistryValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Saral", "Saral Ver", "")
'
'    MousePointer = vbHourglass
'    If Execonn.State = 1 Then Execonn.Close
'    Execonn.ConnectionString = "Provider=SQLOLEDB;Server=" & Trim(ReadFromFile("Server Name")) & ";Database=ExeVersionSystem;USER ID=sa; Password=" & Trim(ReadFromFile("Server Password")) & ""
'
'    Execonn.Open
'
'    If Rs.State = 1 Then Rs.Close
'    Rs.Open "Select Top 1 * from UploadExe where ModuleName='Saral' order by id Desc", Execonn, adOpenKeyset, adLockOptimistic
'    If Rs.RecordCount > 0 Then
'        If Rs.Fields("ID") > SoftwareID Then
'           Label2.Caption = " Double Click for Download New Version"
'           Timer1.Enabled = False
'           Timer1.Interval = 0
'           tmrCaption.Enabled = True
'           tmrMain.Enabled = True
'
'        Else
'           Label2.Caption = " Saral-Production Management System"
'           tmrCaption.Enabled = False
'           tmrMain.Enabled = False
'           Timer1.Enabled = True
'           Timer1.Interval = 50
'        End If
'    Else
'           Label2.Caption = " Saral-Production Management System"
'           tmrCaption.Enabled = False
'           tmrMain.Enabled = False
'           Timer1.Enabled = True
'           Timer1.Interval = 50
'    End If
'    MousePointer = vbNormal
'Exit Sub
'ErrorHandler:
End Sub
Private Sub Label2_DblClick()

On Error GoTo ErrorHandler
    Dim strFile As New FileSystemObject
    Dim FileFullPath As String
    MousePointer = vbHourglass
    Dim mStream As New ADODB.Stream
    Dim Rs As New ADODB.Recordset
    Dim dblRetVal As Double
    Dim iFileNumber As Integer
    Dim Execonn As New ADODB.Connection
    If strFile.FileExists("C:\Setting.inf") = False Then
        strFile.CreateTextFile ("C:\Setting.inf")
    Else
         strFile.DeleteFile ("C:\Setting.inf")
    End If
    
    FileFullPath = "C:\Setting.inf"
    iFileNumber = FreeFile
    If Overwrite Then
        Open FileFullPath For Output As #iFileNumber
    Else
        Open FileFullPath For Append As #iFileNumber
    End If
    Print #iFileNumber, App.Path
    Print #iFileNumber, App.EXEName
    SaveTextToFile = True
    
    Close #iFileNumber
    
    If Execonn.State = 1 Then Execonn.Close
    Execonn.ConnectionString = "Provider=SQLOLEDB;Server=" & Trim(ReadFromFile("Server Name")) & ";Database=ExeVersionSystem;USER ID=sa; Password=" & Trim(ReadFromFile("Server Password")) & ""
    Execonn.Open
    If Rs.State = 1 Then Rs.Close
    Rs.Open "Select Top 1 * from UploadExe where ModuleName='Saral' order by id Desc", Execonn, adOpenKeyset, adLockOptimistic
      
      With mStream
        .Type = adTypeBinary
        .Open
        .Write Rs("Image1")
        DoEvents
        .SaveToFile App.Path & "\SaralTemp.exe", adSaveCreateOverWrite
      End With
      
      Call oRegistry.SetRegistryValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Saral", "Saral ID", Rs.Fields("ID"))
      Call oRegistry.SetRegistryValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Saral", "Saral Ver", Rs.Fields("ExeVer"))
      Call oRegistry.SetRegistryValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Saral", "Saral UpdateOn", Rs.Fields("ExeDate"))
      
      dblRetVal = Shell(App.Path & "\SaralTemp.exe", vbNormalFocus)
     
      Set mStream = Nothing
      MousePointer = vbNormal
      End
Exit Sub
ErrorHandler:
Close #iFileNumber
End Sub
Private Sub Label41_Click()
'Unload Me
Call CloseForm(FrmSplash)
End Sub
Private Sub Timer1_Timer()
    On Error Resume Next
    T = T + 2
    If T > 100 Then
        'Unload Me
        Timer1.Enabled = False
        Load MdiMainMenu
        If Err.Number <> 364 Then MdiMainMenu.Show
        Call CloseForm(FrmSplash)
        Exit Sub
    End If
    ProgressBar1.Value = T
End Sub

Private Sub tmrCaption_Timer()
 Label2.ForeColor = &H800080
End Sub

Private Sub tmrMain_Timer()
 tmrMain.Interval = 690
 Label2.ForeColor = &HFFFF&
 '&HC000C0
End Sub

