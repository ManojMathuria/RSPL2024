VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmLicenceAgreement 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Licence Agreement"
   ClientHeight    =   9465
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   13440
   ClipControls    =   0   'False
   Icon            =   "LicenceAgreement.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9465
   ScaleWidth      =   13440
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2880
      TabIndex        =   16
      Top             =   8520
      Width           =   4695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Update Masters"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   13
      Top             =   9000
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Update Version"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   11400
      TabIndex        =   12
      Top             =   9000
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Agree to Terms && Conditions"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7800
      TabIndex        =   11
      Top             =   8160
      Width           =   3375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Activate &Later"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   11400
      TabIndex        =   10
      Top             =   8520
      Width           =   1725
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Activate Key"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7680
      TabIndex        =   9
      Top             =   8520
      Visible         =   0   'False
      Width           =   3645
   End
   Begin FPSpreadADO.fpSpread fpSpread1 
      Height          =   5895
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   13215
      _Version        =   524288
      _ExtentX        =   23310
      _ExtentY        =   10398
      _StockProps     =   64
      AllowMultiBlocks=   -1  'True
      EditEnterAction =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GridShowHoriz   =   0   'False
      GridShowVert    =   0   'False
      MaxRows         =   506
      OperationMode   =   1
      SelectBlockOptions=   1
      SpreadDesigner  =   "LicenceAgreement.frx":000C
      TabEnhancedShape=   1
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   1020
      Left            =   106
      Picture         =   "LicenceAgreement.frx":2093
      ScaleHeight     =   674.24
      ScaleMode       =   0  'User
      ScaleWidth      =   674.24
      TabIndex        =   1
      Top             =   135
      Width           =   1020
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   11400
      TabIndex        =   0
      Top             =   7545
      Width           =   1740
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&System Info..."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   11400
      TabIndex        =   2
      Top             =   7995
      Width           =   1725
   End
   Begin MSForms.ComboBox Combo2 
      Height          =   405
      Left            =   2880
      TabIndex        =   17
      Top             =   8520
      Visible         =   0   'False
      Width           =   4695
      VariousPropertyBits=   545282075
      BackColor       =   16777215
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "8281;714"
      MatchEntry      =   0
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Calibri"
      FontEffects     =   1073741825
      FontHeight      =   285
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   405
      Left            =   120
      TabIndex        =   15
      Top             =   8520
      Width           =   2685
      VariousPropertyBits=   545282075
      BackColor       =   16777215
      BorderStyle     =   1
      DisplayStyle    =   7
      Size            =   "4736;714"
      MatchEntry      =   0
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Calibri"
      FontEffects     =   1073741825
      FontHeight      =   285
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3000
      TabIndex        =   14
      Top             =   9000
      Width           =   8205
   End
   Begin VB.Label Label1 
      Caption         =   " Renewal Key :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   8
      Top             =   8520
      Width           =   1575
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   -30
      X2              =   19200
      Y1              =   7365
      Y2              =   7365
   End
   Begin VB.Label lblDescription 
      Caption         =   "Website: http://www.easyinfosolution.com/   email: sales@easyinfosolution.com"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   1170
      TabIndex        =   3
      Top             =   960
      Width           =   12045
   End
   Begin VB.Label lblTitle 
      Caption         =   "Easy Info Solutions International"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   1170
      TabIndex        =   5
      Top             =   240
      Width           =   12045
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   -15
      X2              =   19200
      Y1              =   7380
      Y2              =   7380
   End
   Begin VB.Label lblVersion 
      Caption         =   "Easy Publish  21|Rel 05 | 06.29 Version |Production & Inventory Management System"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1170
      TabIndex        =   6
      Top             =   660
      Width           =   12045
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   $"LicenceAgreement.frx":2BF0
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   945
      Left            =   135
      TabIndex        =   4
      Top             =   7545
      Width           =   10350
   End
End
Attribute VB_Name = "frmLicenceAgreement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Decrypt As Variant
Dim EncryptFlag As Boolean, LaterFlag As Boolean, ServerID As String, RenewFlag As Boolean
' Website hyperlink...
Private WithEvents oHuffman As clsHuffman
Attribute oHuffman.VB_VarHelpID = -1
Private oRegistry As New clsRegistry
'Private Developer As String, DueDate As String, DaysLeft As Variant
Private Declare Function ShellExecute _
                            Lib "shell32.dll" _
                            Alias "ShellExecuteA" ( _
                            ByVal hwnd As Long, _
                            ByVal lpOperation As String, _
                            ByVal lpFile As String, _
                            ByVal lpParameters As String, _
                            ByVal lpDirectory As String, _
                            ByVal nShowCmd As Long) _
                            As Long
' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"
Dim VchType As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Dim rstCompanyMaster As New ADODB.Recordset
Private Sub Check1_Click()
    If Check1.Value Then Command1.Visible = True
    If Check1.Value = 0 Then Command1.Visible = False
End Sub
Private Sub cmdSysInfo_Click()
    Call StartSysInfo
End Sub
Private Sub cmdOK_Click()
    Unload Me
End Sub
Private Sub Combo1_Change()
If Combo1.ListIndex = 0 Then 'Super User
    Text1.Visible = False: Combo2.Visible = True
            Combo2.Clear
            Combo2.AddItem "EasyPublish", 0
            Combo2.AddItem "Admin", 1
        If Trim(ReadFromFile(Combo1.Text)) = "EasyPublish" Then
            Combo2.ListIndex = 0
        ElseIf Trim(ReadFromFile(Combo1.Text)) = "Admin" Then
            Combo2.ListIndex = 1
        End If
ElseIf Combo1.ListIndex = 2 Then 'Data Base Type
    Text1.Visible = False: Combo2.Visible = True
            Combo2.Clear
            Combo2.AddItem "MS Access", 0
            Combo2.AddItem "MS SQL", 1
        If Trim(ReadFromFile(Combo1.Text)) = "MS Access" Then
            Combo2.ListIndex = 0
        ElseIf Trim(ReadFromFile(Combo1.Text)) = "MS SQL" Then
            Combo2.ListIndex = 1
        End If
ElseIf Combo1.ListIndex = 8 Then 'Client ID
    Text1.Visible = False: Combo2.Visible = True
            Combo2.Clear
            Combo2.AddItem "Printer", 0
            Combo2.AddItem "Publisher", 1
        If Trim(ReadFromFile(Combo1.Text)) = "Printer" Then
            Combo2.ListIndex = 0
        ElseIf Trim(ReadFromFile(Combo1.Text)) = "Publisher" Then
            Combo2.ListIndex = 1
        End If
End If
If Combo1.ListIndex = 1 Or Combo1.ListIndex = 3 Or Combo1.ListIndex = 4 Or Combo1.ListIndex = 5 Or Combo1.ListIndex = 6 Or Combo1.ListIndex = 7 Then
    Text1.Visible = True: Combo2.Visible = False
End If
Text1.Text = Trim(ReadFromFile(Combo1.Text))
EncryptFlag = False
Command1.Caption = IIf(Combo1.ListIndex = 0, "Update Super User", IIf(Combo1.ListIndex = 1, "Activate Key", IIf(Combo1.ListIndex = 2, "Update Database Type", IIf(Combo1.ListIndex = 3, "Update Server Name", IIf(Combo1.ListIndex = 4, "Update Server User", IIf(Combo1.ListIndex = 5, "Update Server Passward", IIf(Combo1.ListIndex = 6, "Update Tally Port", IIf(Combo1.ListIndex = 7, "Update Server ID", IIf(Combo1.ListIndex = 8, "Update Client ID", " ")))))))))
End Sub
Private Sub Command2_Click()
    Unload Me
    LaterFlag = True
End Sub
Public Function Encrypted(Decrypted As Variant, Encrypt As Variant)
Dim K As Long, N As Long
Dim Flag As Boolean
Dim Key As String
Dim Key1 As String
Dim Key2 As String
Dim Key3 As String
Dim Key4 As String
Dim dueDate As Date, sNow As Date
Key1 = "": Key2 = "": Key3 = "": Key4 = "": Key = ""
K = 0:
'Company Fundation Date 28-SEP-2-16
N = Len(Text1.Text)

'Check For Stoping Existing Encryption
Do While N <> 0 And EncryptFlag = False And Key <> "§"
If Flag = False Then Key1 = "E"
If N <> 0 Then K = K + 1: Key = (Mid(Text1.Text, K, 1)): N = N - 1:
Loop
If Key = "§" Then EncryptFlag = True

'Start Encryption
N = Len(Text1.Text): K = 0
Do While N <> 0 And EncryptFlag = False
If Flag = False Then Key1 = "E"
If N <> 0 Then K = K + 1: Key1 = Key1 + (Mid(Text1.Text, K, 1)): N = N - 1:

If Flag = False Then Key2 = Key2 + "§I"
If N <> 0 Then K = K + 1: Key2 = Key2 + Mid(Text1.Text, K, 1): N = N - 1:

If Flag = False Then Key3 = Key3 + "§S"
If N <> 0 Then K = K + 1: Key3 = Key3 + Mid(Text1.Text, K, 1): N = N - 1:

If Flag = False Then Key4 = Key4 + "§I"
If N <> 0 Then K = K + 1: Key4 = Key4 + Mid(Text1.Text, K, 1): N = N - 1:
Flag = True
Loop
If EncryptFlag = False Then Encrypt = Key1 + Key2 + Key3 + Key4 + "§" + " " + "§"
If EncryptFlag = False Then sNow = Format(Now(), "DD-MMM-YYYY")
If EncryptFlag = False Then Encrypted = True
End Function
Private Sub Command1_Click()
Decrypt = "":
If Combo1.ListIndex <> 0 And Combo1.ListIndex <> 1 And Combo1.ListIndex <> 2 And Combo1.ListIndex <> 6 And Combo1.ListIndex <> 7 And Combo1.ListIndex <> 8 Then
    If Encrypted(Trim(Text1.Text), Decrypt) Then
        If EncryptFlag = False And Text1.Text <> "" Then Text1.Text = Decrypt: EncryptFlag = True
    End If
End If
If Combo1.ListIndex = 0 Then 'Supper User
    WriteToFile Combo1.Text, Combo2.Text
ElseIf Combo1.ListIndex = 1 Then 'Renewal Key
    If Text1.Text <> "" Then
        WriteToFile "Server ID", Text1.Text + "@" + ServerID
        RenewFlag = True
        cmdOK_Click
    Else
        Text1.SetFocus
    End If
ElseIf Combo1.ListIndex = 2 Then 'Database Type
    WriteToFile Combo1.Text, Combo2.Text
ElseIf EncryptFlag = True And Combo1.ListIndex = 3 Then 'Server Name
    WriteToFile Combo1.Text, Text1.Text
ElseIf EncryptFlag = True And Combo1.ListIndex = 4 Then 'Server User
    WriteToFile Combo1.Text, Text1.Text
ElseIf EncryptFlag = True And Combo1.ListIndex = 5 Then 'Server Password
    WriteToFile Combo1.Text, Text1.Text
ElseIf Combo1.ListIndex = 6 Then 'Tally Port
    WriteToFile Combo1.Text, Text1.Text
 ElseIf Combo1.ListIndex = 7 Then 'Server ID
    'WriteToFile Combo1.Text, Text1.Text
 ElseIf Combo1.ListIndex = 8 Then 'Cleint ID
    WriteToFile Combo1.Text, Combo2.Text
 End If
    Command1.Visible = False
    Check1.Value = 0
End Sub
Private Sub Form_Load()
    Dim R As Long, C As Long
    CenterForm Me
    Me.Caption = "License Agreement"   '"About " & App.Title
    If Trim(ReadFromFile("Super User")) = "EasyPublish" Then Command3.Visible = True Else Me.Height = 9540: Command3.Visible = False
    If Trim(ReadFromFile("Super User")) = "EasyPublish" Then Command4.Visible = True Else Me.Height = 9540: Command4.Visible = False
    If Dir(App.Path & "\Icon\ICON.ICO", vbDirectory) <> "" Then Me.Icon = LoadPicture(App.Path & "\Icon\ICON.ICO")
    lblVersion.Caption = "Easy Publish |Rel  21.05 |Version " & App.Major & "." & App.Minor & "." & App.Revision & " |Production && Inventory Management System"
    'Easy Publish  21|Rel 05 | 06.29 Version |Production & Inventory Management System
    lblTitle.Caption = "Easy Info Solutions International" 'App.Title
    
    With fpSpread1
    .MaxRows = 57: .MaxCols = 2
    fpSpread1.RowHeadersShow = False
    fpSpread1.ColHeadersShow = False
    
    For R = 1 To 22
    C = 1
        fpSpread1.Col = C: fpSpread1.Row = R: fpSpread1.CellType = CellTypeEdit: fpSpread1.TypeHAlign = TypeHAlignCenter
    Next
        fpSpread1.Col = 2: fpSpread1.Row = 1: fpSpread1.CellType = CellTypeEdit: fpSpread1.TypeHAlign = TypeHAlignCenter: fpSpread1.RowsFrozen = 7
    For R = 3 To 22
    C = 2
        fpSpread1.Col = C: fpSpread1.Row = R: fpSpread1.CellType = CellTypeEdit: fpSpread1.TypeHAlign = TypeHAlignLeft: fpSpread1.TypeTextWordWrap = True: fpSpread1.RowMerge = MergeAlways
    Next
    End With
    Combo1.AddItem "Super User", 0
    Combo1.AddItem "Renewal Key", 1
    Combo1.AddItem "Database Type", 2
    Combo1.AddItem "Server Name", 3
    Combo1.AddItem "Server User", 4
    Combo1.AddItem "Server Password", 5
    Combo1.AddItem "Tally Port", 6
    Combo1.AddItem "Server ID", 7
    Combo1.AddItem "Client ID", 8
    Combo1.ListIndex = 1
If Combo1.ListIndex = 0 Then
    Combo2.AddItem "EasyPublish", 0
    Combo2.AddItem "Admin", 1
    Combo2.ListIndex = 1
ElseIf Combo1.ListIndex = 2 Then
    Combo2.AddItem "MS Access", 0
    Combo2.AddItem "MS SQL", 1
    Combo2.ListIndex = 1
ElseIf Combo1.ListIndex = 8 Then
    Combo2.AddItem "Printer", 0
    Combo2.AddItem "Publisher", 1
    Combo2.ListIndex = 1
End If
    Text1.Text = Trim(ReadFromFile(Combo1.Text))
    Command1.Caption = IIf(Combo1.ListIndex = 0, "Update Super User", IIf(Combo1.ListIndex = 1, "Activate Key", IIf(Combo1.ListIndex = 2, "Update Database Type", IIf(Combo1.ListIndex = 3, "Update Server Name", IIf(Combo1.ListIndex = 4, "Update Server User", IIf(Combo1.ListIndex = 5, "Update Server Passward", IIf(Combo1.ListIndex = 6, "Update Tally Port", IIf(Combo1.ListIndex = 7, "Update Server ID", IIf(Combo1.ListIndex = 8, "Update Client ID", "")))))))))
End Sub
Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' Try To Get System Info Program Path\Name From Registry...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' Try To Get System Info Program Path Only From Registry...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Validate Existance Of Known 32 Bit File Version
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' Error - File Can Not Be Found...
        Else
            GoTo SysInfoErr
        End If
    ' Error - Registry Entry Can Not Be Found...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "System Information Is Unavailable At This Time", vbOKOnly
End Sub
Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ                                             ' String Registry Key Data Type
        KeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
    End Select
    
    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:      ' Cleanup After An Error Has Occured...
    KeyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function
Private Sub lblDescription_Click()
Dim R As Long
      R = ShellExecute(0, "open", "http://www.easyinfosolution.com", 0, 0, 1)
End Sub
Private Sub Command4_Click()
'    If UpdateComp(CompCode, False, False, True) Then
'    'If UpdateMaster(CompCode, 1) Then
'        Call MsgBox("Successfully Updated Masters !", vbInformation, App.Title)
'    Else
'        DisplayError ("Failed to Updated Master")
'    End If
End Sub
'Private Function UpdateMaster(ByVal CompanyCode As String, ByVal WithMasters As Boolean) As Boolean
'    On Error GoTo ErrorHandler
'    UpdateMaster = True
'    cnDatabase.CursorLocation = adUseClient
'    If cnDatabase.State = adStateOpen Then cnDatabase.Close
'    If DatabaseType = "MS SQL" Then
'    ConnectionString = "Provider=SQLOLEDB;Password=" & ServerPassword & ";Persist Security Info=True;User ID=" & ServerUser & ";Initial Catalog=EP" & CompCode & ";Data Source=" & ServerName
'    cnDatabase.Open ConnectionString
'    ElseIf DatabaseType = "MS Access" Then
'    cnDatabase.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DatabasePath & "\EasyPublish." & CompanyCode & ";Persist Security Info=False;Jet OLEDB:Database Password=pubprint123!@#"
'    End If
'    cnDatabase.BeginTrans
'    If DatabaseType = "MS SQL" Then
'    'BackUpDatabse
'        cnDatabase.Execute "BACKUP DATABASE [EP" & CompCode & "] TO  DISK = N'C:\Program Files\Microsoft SQL Server\MSSQL13.MSSQLSERVER\MSSQL\Backup\EP" & CompCode & "_LogBackup_temp.bak' WITH NOFORMAT, NOINIT,  NAME = N'EP" & CompCode & " -Full Database Backup', SKIP, NOREWIND, NOUNLOAD,  STATS = 10"
'    'RestoreDatabse
'        cnDatabase.Execute "RESTORE DATABASE [EP" & CompanyCode & "] FROM  DISK = N'C:\Program Files\Microsoft SQL Server\MSSQL13.MSSQLSERVER\MSSQL\Backup\EP" & CompCode & "_LogBackup_temp.bak' WITH  FILE = 1,  MOVE N'EPM' TO N'C:\Program Files\Microsoft SQL Server\MSSQL12.MSSQLSERVER\MSSQL\DATA\EP" & CompanyCode & "M.mdf',  MOVE N'EPL' TO N'C:\Program Files\Microsoft SQL Server\MSSQL12.MSSQLSERVER\MSSQL\DATA\EP" & CompanyCode & "L.ldf',  NOUNLOAD,  STATS = 5"
'    End If
'    cnDatabase.CommitTrans
'    'CloseMainConnection
'    CompCode = CompanyCode
'    cnDatabase.CursorLocation = adUseClient
'    If cnDatabase.State = adStateOpen Then cnDatabase.Close
'    If DatabaseType = "MS SQL" Then
'        ConnectionString = "Provider=SQLOLEDB;Password=" & ServerPassword & ";Persist Security Info=True;User ID=" & ServerUser & ";Initial Catalog=EP" & CompCode & ";Data Source=" & ServerName
'    cnDatabase.Open ConnectionString
'    End If
'    cnDatabase.BeginTrans
''    cnDatabase.Execute "DELETE FROM CompanyMaster"
''        cnDatabase.Execute "INSERT INTO CompanyMaster (Code,Name,PrintName,Address1,Address2,Address3,Address4,Phone,Mobile,Fax,eMail,Website,GSTIN,CreatedFrom,MCGroup,MCPrimary,MCRepair,FinancialYearFrom,FinancialYearTo,Printstatus,TitleCombo,BankName,AccountNo,IFSC,TallyIntegration,BusyIntegration,FYCode,Alias) VALUES ('000001','" & Trim(FrmCompanyMaster.Text1.Text) & "','" & Trim(FrmCompanyMaster.Text2.Text) & "','" & Trim(FrmCompanyMaster.Text3.Text) & "','" & Trim(FrmCompanyMaster.Text4.Text) & "','" & Trim(FrmCompanyMaster.Text5.Text) & "','" & Trim(FrmCompanyMaster.Text6.Text) & "','" & Trim(FrmCompanyMaster.Text7.Text) & "','" & Trim(FrmCompanyMaster.Text11.Text) & "','" & Trim(FrmCompanyMaster.Text12.Text) & "'" & _
''                                          ",'" & Trim(FrmCompanyMaster.Text8.Text) & "','" & Trim(FrmCompanyMaster.Text9.Text) & "','" & Trim(FrmCompanyMaster.Text10.Text) & "','" & CompCode & "','0','0','0','" & Format(GetDate(FrmCompanyMaster.MhDateInput1.Text), "mm-dd-yyyy") & "','" & Format(GetDate(FrmCompanyMaster.MhDateInput2.Text), "mm-dd-yyyy") & "','N','1','" & Trim(FrmCompanyMaster.Text18.Text) & "','" & Trim(FrmCompanyMaster.Text19.Text) & "','" & Trim(FrmCompanyMaster.Text20.Text) & "','" & Trim(FrmCompanyMaster.Option1.Value) & "','" & Trim(FrmCompanyMaster.Option2.Value) & "','" & Trim(FrmCompanyMaster.Text16.Text) & "','" & Trim(FrmCompanyMaster.Text15.Text) & "')"
'
'    'Transactions 44_Tables
'        cnDatabase.Execute "DELETE FROM BookDNChild"
'        cnDatabase.Execute "DELETE FROM BookDNParent"
'        cnDatabase.Execute "DELETE FROM BookOOChild"
'        cnDatabase.Execute "DELETE FROM BookOOParent"
'        cnDatabase.Execute "DELETE FROM BookPOChild05"
'        cnDatabase.Execute "DELETE FROM BookPOChild0501"
'        cnDatabase.Execute "DELETE FROM BookPOChild06"
'        cnDatabase.Execute "DELETE FROM BookPOChild07"
'        cnDatabase.Execute "DELETE FROM BookPOChild08"
'        cnDatabase.Execute "DELETE FROM BookPOChild0801"
'        cnDatabase.Execute "DELETE FROM BookPOChild09"
'        cnDatabase.Execute "DELETE FROM BookPOChild0901"
'        cnDatabase.Execute "DELETE FROM BookPOParent"
'        cnDatabase.Execute "DELETE FROM BookRVChild"
'        cnDatabase.Execute "DELETE FROM BookRVParent"
'        cnDatabase.Execute "DELETE FROM DebitCreditParent"
'        cnDatabase.Execute "DELETE FROM DebitCreditChild"
'        cnDatabase.Execute "DELETE FROM DebitCreditOthInf"
'        cnDatabase.Execute "DELETE FROM DebitCreditRef"
'        cnDatabase.Execute "DELETE FROM JobworkBVChild"
'        cnDatabase.Execute "DELETE FROM JobworkBVOthInf"
'        cnDatabase.Execute "DELETE FROM JobworkBVRef"
'        cnDatabase.Execute "DELETE FROM JobworkBVParent"
'        cnDatabase.Execute "DELETE FROM MaterialIOChild"
'        cnDatabase.Execute "DELETE FROM MaterialIOParent"
'        cnDatabase.Execute "DELETE FROM MaterialMVChild"
'        cnDatabase.Execute "DELETE FROM MaterialMVParent"
'        cnDatabase.Execute "DELETE FROM MaterialSVChild"
'        cnDatabase.Execute "DELETE FROM MaterialSVParent"
'        cnDatabase.Execute "DELETE FROM OutsourceItemPOChild"
'        cnDatabase.Execute "DELETE FROM OutsourceItemPOParent"
'        cnDatabase.Execute "DELETE FROM PackingSlipChild"
'        cnDatabase.Execute "DELETE FROM PackingSlipParent"
'        cnDatabase.Execute "DELETE FROM PaperDNChild"
'        cnDatabase.Execute "DELETE FROM PaperDNParent"
'        cnDatabase.Execute "DELETE FROM PaperIOChild"
'        cnDatabase.Execute "DELETE FROM PaperMVChild"
'        cnDatabase.Execute "DELETE FROM PaperMVParent"
'        cnDatabase.Execute "DELETE FROM PaperPOChild"
'        cnDatabase.Execute "DELETE FROM PaperPOParent"
'        cnDatabase.Execute "DELETE FROM PrintPVChild"
'        cnDatabase.Execute "DELETE FROM PrintPVParent"
'        cnDatabase.Execute "DELETE FROM TatRVChild"
'        cnDatabase.Execute "DELETE FROM TatRVParent"
''Without MAsters
'    If Not WithMasters Then    'Delete Master
'    'Accounts Master
'        cnDatabase.Execute "DELETE FROM AccountChild04 Where CODE IN (Select Code From AccountMaster Where Right([Group],5)<'10001' AND Left(Code,1)<>'*' AND Code<> '000000')"
'        cnDatabase.Execute "DELETE FROM AccountChild05 Where CODE IN (Select Code From AccountMaster Where Right([Group],5)<'10001' AND Left(Code,1)<>'*' AND Code<> '000000')"
'        cnDatabase.Execute "DELETE FROM AccountChild06 Where CODE IN (Select Code From AccountMaster Where Right([Group],5)<'10001' AND Left(Code,1)<>'*' AND Code<> '000000')"
'        cnDatabase.Execute "DELETE FROM AccountChild07 Where CODE IN (Select Code From AccountMaster Where Right([Group],5)<'10001' AND Left(Code,1)<>'*' AND Code<> '000000')"
'        cnDatabase.Execute "DELETE FROM AccountChild08 Where CODE IN (Select Code From AccountMaster Where Right([Group],5)<'10001' AND Left(Code,1)<>'*' AND Code<> '000000')"
'        cnDatabase.Execute "DELETE FROM AccountChild0801 Where CODE IN (Select Code From AccountMaster Where Right([Group],5)<'10001' AND Left(Code,1)<>'*' AND Code<> '000000')"
'        cnDatabase.Execute "DELETE FROM AccountMaster "
'
'    'Book Master
'        cnDatabase.Execute "DELETE FROM BookChild01 Where Left(Code,1)<>'*'"
'        cnDatabase.Execute "DELETE FROM BookChild02 Where Left(Code,1)<>'*'"
'        cnDatabase.Execute "DELETE FROM BookChild03 Where Left(Code,1)<>'*'"
'        cnDatabase.Execute "DELETE FROM BookChild05 Where Left(Code,1)<>'*'"
'        cnDatabase.Execute "DELETE FROM BookChild06 Where Left(Code,1)<>'*'"
'        cnDatabase.Execute "DELETE FROM BookChild07 Where Left(Code,1)<>'*'"
'        'cnDatabase.Execute "DELETE FROM BookChild08 Where Left(Code,1)<>'*'"
'        cnDatabase.Execute "DELETE FROM BookMaster Where Left(Code,1)<>'*'"
'    'Other Masters
'        cnDatabase.Execute "DELETE FROM BookingRouteMaster Where Left(Code,1)<>'*'"
'        cnDatabase.Execute "DELETE FROM DiscountMaster "
'        cnDatabase.Execute "DELETE FROM ElementMaster Where Left(Code,1)<>'*'"
'        If MsgBox("Do You Wants to Delete 'Finish Size Masters' Also !!!" & vbCrLf & "Please Make Sure Before Process !!!", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Proceed !") = vbYes Then
'        cnDatabase.Execute "DELETE FROM FinishSizeChild Where Left(Code,1)<>'*'"
'        End If
'        cnDatabase.Execute "DELETE GeneralMaster Where Left(Code,1)<>'*' And Type='5' And Name <> 'General'"
'        cnDatabase.Execute "DELETE FROM OutsourceItemMaster Where Left(Code,1)<>'*'"
'        If MsgBox("Do You Wants to Delete 'Paper Master' Also !!!" & vbCrLf & "Please Make Sure Before Process !!!", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Proceed !") = vbYes Then
'        cnDatabase.Execute "DELETE FROM PaperMaster Where Left(Code,1)<>'*'"
'        End If
'        If MsgBox("Do You Wants to Delete 'Size Group Masters' Also !!!" & vbCrLf & "Please Make Sure Before Process !!!", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Proceed !") = vbYes Then
'        cnDatabase.Execute "DELETE FROM SizeGroupChild Where Left(Code,1)<>'*'"
'        End If
'        cnDatabase.Execute "DELETE FROM TaxMaster Where Left(Code,1)<>'*'"
'        cnDatabase.Execute "DELETE FROM TeamMemberMaster Where Left(Code,1)<>'*'"
'
'    Else
'        cnDatabase.Execute "UPDATE AccountMaster SET CreatedOn=GETDate(), ModifiedBy=Null, ModifiedOn=Null, Recordstatus='N', Printstatus='N',Opening='0'"
'        cnDatabase.Execute "UPDATE BookMaster SET CreatedOn=GETDate(), ModifiedBy=Null, ModifiedOn=Null, Recordstatus='N', Printstatus='N'"
'        cnDatabase.Execute "UPDATE PaperMaster SET CreatedOn=GETDate(), ModifiedBy=Null, ModifiedOn=Null, Recordstatus='N', Printstatus='N'"
'        cnDatabase.Execute "UPDATE OutsourceItemMaster SET CreatedOn=GETDate(), ModifiedBy=Null, ModifiedOn=Null, Recordstatus='N', Printstatus='N'"
'        cnDatabase.Execute "UPDATE TaxMaster SET CreatedOn=GETDate(), ModifiedBy=Null, ModifiedOn=Null, Recordstatus='N', Printstatus='N'"
'        cnDatabase.Execute "UPDATE TeamMemberMaster SET CreatedOn=GETDate(), ModifiedBy=Null, ModifiedOn=Null, Recordstatus='N', Printstatus='N'"
'        cnDatabase.Execute "UPDATE GeneralMaster SET CreatedOn=GETDate(), ModifiedBy=Null, ModifiedOn=Null, Recordstatus='N', Printstatus='N'"
'    End If
'
''Default Masters
'    cnDatabase.Execute "DELETE FROM BookChild"
'    cnDatabase.Execute "DELETE FROM PaperChild"
'    cnDatabase.Execute "DELETE FROM UserChild Where Code NOT IN (Select Code from UserMaster Where Level<>1)"
'    cnDatabase.Execute "DELETE FROM UserMaster Where Level<>1"
'    cnDatabase.Execute "DELETE FROM UserAction"
'    cnDatabase.Execute "DELETE FROM VchSeriesMaster Where Left(Code,1)='*'"
'    cnDatabase.Execute "UPDATE AccountMaster SET Opening='0'"
'
''General Accounts
'    cnDatabase.Execute "DELETE FROM AccountMaster Where Left(Code,1)='*'"
'''Account Masters
'    cnDatabase.Execute "DELETE FROM AccountMaster Where Code ='000000' Or Left(Code,1)='*'"
'    cnDatabase.Execute "Insert Into AccountMaster VALUES ('000000','" & Trim(FrmCompanyMaster.Text1.Text) & "','" & Trim(FrmCompanyMaster.Text2.Text) & "','000000','*12002','" & Trim(FrmCompanyMaster.Text3.Text) & "','" & Trim(FrmCompanyMaster.Text4.Text) & "','" & Trim(FrmCompanyMaster.Text5.Text) & "','" & Trim(FrmCompanyMaster.Text6.Text) & "','" & Trim(FrmCompanyMaster.Text7.Text) & "','" & Trim(FrmCompanyMaster.Text11.Text) & "','" & Trim(FrmCompanyMaster.Text10.Text) & "','" & Trim(FrmCompanyMaster.Text8.Text) & "', 1,'000001',GetDate(),Null,Null,'N','N','',0);"
'    cnDatabase.Execute "Insert Into AccountMaster VALUES ('*00001','Rate Master','Rate Master','1002','*12002','" & Trim(FrmCompanyMaster.Text3.Text) & "','" & Trim(FrmCompanyMaster.Text4.Text) & "','" & Trim(FrmCompanyMaster.Text5.Text) & "','" & Trim(FrmCompanyMaster.Text6.Text) & "','" & Trim(FrmCompanyMaster.Text7.Text) & "','" & Trim(FrmCompanyMaster.Text11.Text) & "','" & Trim(FrmCompanyMaster.Text10.Text) & "','" & Trim(FrmCompanyMaster.Text8.Text) & "', 1,'000001',GetDate(),Null,Null,'N','N','',0);"
'    cnDatabase.Execute "Insert Into AccountMaster VALUES ('*00002','Main Godown','Main Godown','1003','*99999','" & Trim(FrmCompanyMaster.Text3.Text) & "','" & Trim(FrmCompanyMaster.Text4.Text) & "','" & Trim(FrmCompanyMaster.Text5.Text) & "','" & Trim(FrmCompanyMaster.Text6.Text) & "','" & Trim(FrmCompanyMaster.Text7.Text) & "','" & Trim(FrmCompanyMaster.Text11.Text) & "','" & Trim(FrmCompanyMaster.Text10.Text) & "','" & Trim(FrmCompanyMaster.Text8.Text) & "', 1,'000001',GetDate(),Null,Null,'N','N','',0);"
'    cnDatabase.Execute "Insert Into AccountMaster VALUES ('*00003','Self Transport','Self Transport','1004','*99996','" & Trim(FrmCompanyMaster.Text3.Text) & "','" & Trim(FrmCompanyMaster.Text4.Text) & "','" & Trim(FrmCompanyMaster.Text5.Text) & "','" & Trim(FrmCompanyMaster.Text6.Text) & "','" & Trim(FrmCompanyMaster.Text7.Text) & "','" & Trim(FrmCompanyMaster.Text11.Text) & "','" & Trim(FrmCompanyMaster.Text10.Text) & "','" & Trim(FrmCompanyMaster.Text8.Text) & "', 1,'000001',GetDate(),Null,Null,'N','N','',0);"
'    cnDatabase.Execute "Insert Into AccountMaster VALUES ('*00004','Self Packer','Self Packer','1005','*99997','" & Trim(FrmCompanyMaster.Text3.Text) & "','" & Trim(FrmCompanyMaster.Text4.Text) & "','" & Trim(FrmCompanyMaster.Text5.Text) & "','" & Trim(FrmCompanyMaster.Text6.Text) & "','" & Trim(FrmCompanyMaster.Text7.Text) & "','" & Trim(FrmCompanyMaster.Text11.Text) & "','" & Trim(FrmCompanyMaster.Text10.Text) & "','" & Trim(FrmCompanyMaster.Text8.Text) & "', 1,'000001',GetDate(),Null,Null,'N','N','',0);"
'    cnDatabase.Execute "Insert Into AccountMaster VALUES ('*00005','Direct','Direct','1006','*99998','" & Trim(FrmCompanyMaster.Text3.Text) & "','" & Trim(FrmCompanyMaster.Text4.Text) & "','" & Trim(FrmCompanyMaster.Text5.Text) & "','" & Trim(FrmCompanyMaster.Text6.Text) & "','" & Trim(FrmCompanyMaster.Text7.Text) & "','" & Trim(FrmCompanyMaster.Text11.Text) & "','" & Trim(FrmCompanyMaster.Text10.Text) & "','" & Trim(FrmCompanyMaster.Text8.Text) & "', 1,'000001',GetDate(),Null,Null,'N','N','',0);"
'
''Finance Account
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01001','Cash','Cash','1001','*26007','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01002','Development Tax','Development Tax','1002','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01003','Edu. Cess on TDS','Edu. Cess on TDS','1003','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01004','Excise Duty','Excise Duty','1004','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01005','KKC on Service Tax','KKC on Service Tax','1005','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01006','SBC on Service Tax','SBC on Service Tax','1006','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01007','Service Tax','Service Tax','1007','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01008','SHE Cess on TDS','SHE Cess on TDS','1008','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01009','TDS (Commission or Brokerage)','TDS (Commission or Brokerage)','1009','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01010','TDS (Contracts to Individuals/HUF)','TDS (Contracts to Individuals/HUF)','1010','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01011','TDS (Contracts to Others)','TDS (Contracts to Others)','1011','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01012','TDS (Contracts to Transporter)','TDS (Contracts to Transporter)','1012','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01013','TDS (Interest from a Banking Co)','TDS (Interest from a Banking Co)','1013','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01014','TDS (Interest from a NonBanking Co)','TDS (Interest from a NonBanking Co)','1014','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01015','TDS (Professionals Services)','TDS (Professionals Services)','1015','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01016','TDS (Rent of Land)','TDS (Rent of Land)','1016','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01017','TDS (Rent of Plant & Machinery)','TDS (Rent of Plant & Machinery)','1017','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01018','TDS (Salary)','TDS (Salary)','1018','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01019','Advertisement & Publicity','Advertisement & Publicity','1019','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01020','Bad Debts Written Off','Bad Debts Written Off','1020','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01021','Bank Charges','Bank Charges','1021','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01022','Books & Periodicals','Books & Periodicals','1022','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01023','Charity & Donations','Charity & Donations','1023','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01024','Commission on Sales','Commission on Sales','1024','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01025','Conveyance Expenses','Conveyance Expenses','1025','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01026','Customer Entertainment Expenses','Customer Entertainment Expenses','1026','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01027','Depreciation A/c','Depreciation A/c','1027','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01028','Freight & Forwarding Charges','Freight & Forwarding Charges','1028','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01029','Legal Expenses','Legal Expenses','1029','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01030','Miscellaneous Expenses','Miscellaneous Expenses','1030','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01031','Office Maintenance Expenses','Office Maintenance Expenses','1031','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01032','Office Rent','Office Rent','1032','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01033','Postal Expenses','Postal Expenses','1033','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01034','Printing & Stationery','Printing & Stationery','1034','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01035','Rounded Off','Rounded Off','1035','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01036','Salary','Salary','1036','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01037','Sales Promotion Expenses','Sales Promotion Expenses','1037','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01038','Service Charges Paid','Service Charges Paid','1038','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01039','Staff Welfare Expenses','Staff Welfare Expenses','1039','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01040','Telephone Expenses','Telephone Expenses','1040','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01041','Travelling Expenses','Travelling Expenses','1041','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01042','Water & Electricity Expenses','Water & Electricity Expenses','1042','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01043','Capital Equipments','Capital Equipments','1043','*26016','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01044','Computers','Computers','1044','*26016','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01045','Furniture & Fixture','Furniture & Fixture','1045','*26016','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01046','Office Equipments','Office Equipments','1046','*26016','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01047','Plant & Machinery','Plant & Machinery','1047','*26016','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01048','Service Charges Receipts','Service Charges Receipts','1048','*26018','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01049','Profit & Loss','Profit & Loss','1049','*26001','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01050','Salary & Bonus Payable','Salary & Bonus Payable','1050','*26024','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01051','Purchase','Purchase','1051','*26025','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01052','Sales','Sales','1052','*26027','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01053','Earnest Money','Earnest Money','1053','*26029','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01054','Stock','Stock','1054','*26003','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01055','Easy Info Solutions International','Easy Info Solutions International','1055','*26030','E-461, Vijay Marg,Jagjeet Nagar','Delhi-110053','','','','+91-987-342-2907','','sales@easyinfosolution.com ','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'        cnDatabase.Execute "Insert Into AccountMaster VALUES ('*01056','XXX Bank','XXX Bank','1056','*26004','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0)"
'
''Booking Route Master
'        cnDatabase.Execute "DELETE FROM BookingRouteMaster Where Left(Code,1)='*'"
'        cnDatabase.Execute "Insert Into BookingRouteMaster VALUES ('*00001','NOIDA-NOIDA','NOIDA-NOIDA','24.5','N')"
'        cnDatabase.Execute "Insert Into BookingRouteMaster VALUES ('*00002','NOIDA-DELHI','NOIDA-DELHI','40','N')"
'        cnDatabase.Execute "Insert Into BookingRouteMaster VALUES ('*00003','DELHI-DELHI','DELHI-DELHI','30','N')"
'
''Element Master
'    cnDatabase.Execute "DELETE FROM ElementMaster Where Left(Code,1)='*'"
'    cnDatabase.Execute "Insert Into ElementMaster VALUES ('*00011','Text-1','Text-1','Single Sheet','8','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'    cnDatabase.Execute "Insert Into ElementMaster VALUES ('*00012','Text-2','Text-2','Multi Forms','8','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'    cnDatabase.Execute "Insert Into ElementMaster VALUES ('*00013','Text-3','Text-3','Multi Forms','8','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'    cnDatabase.Execute "Insert Into ElementMaster VALUES ('*00014','Single Form','Single Form','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'    cnDatabase.Execute "Insert Into ElementMaster VALUES ('*00015','Combo Form','Combo Form','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'    cnDatabase.Execute "Insert Into ElementMaster VALUES ('*00016','FG','FG','FG','8','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'    cnDatabase.Execute "Insert Into ElementMaster VALUES ('*00017','UFG','UFG','UFG','8','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'    cnDatabase.Execute "Insert Into ElementMaster VALUES ('*00018','Separator','Separator','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'    cnDatabase.Execute "Insert Into ElementMaster VALUES ('*00019','End Paper','End Paper','Single Sheet','4','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'    cnDatabase.Execute "Insert Into ElementMaster VALUES ('*00020','Cover','Cover','Single Sheet','4','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'    cnDatabase.Execute "Insert Into ElementMaster VALUES ('*00027','Title','Title','Single Sheet','4','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'    cnDatabase.Execute "Insert Into ElementMaster VALUES ('*00028','Title(GateFold)','Title(GateFold)','Single Sheet','6','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'    cnDatabase.Execute "Insert Into ElementMaster VALUES ('*00029','PLC','PLC','Single Sheet','4','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'    cnDatabase.Execute "Insert Into ElementMaster VALUES ('*00030','Calendar Fly Leaf','Calendar Fly Leaf','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'    cnDatabase.Execute "Insert Into ElementMaster VALUES ('*00031','Calendar Leaf','Calendar Leaf','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'    cnDatabase.Execute "Insert Into ElementMaster VALUES ('*00032','Annual Report','Annual Report','Multi Forms','8','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'    cnDatabase.Execute "Insert Into ElementMaster VALUES ('*00033','Label','Label','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'    cnDatabase.Execute "Insert Into ElementMaster VALUES ('*00034','Letter Head','Letter Head','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'    cnDatabase.Execute "Insert Into ElementMaster VALUES ('*00035','Leaflet','Leaflet','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'    cnDatabase.Execute "Insert Into ElementMaster VALUES ('*00036','Poster','Poster','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'    cnDatabase.Execute "Insert Into ElementMaster VALUES ('*00037','Sticker','Sticker','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'    cnDatabase.Execute "Insert Into ElementMaster VALUES ('*00038','Folders','Folders','Single Sheet','4','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'    cnDatabase.Execute "Insert Into ElementMaster VALUES ('*00039','Dust Cover','Dust Cover','Single Sheet','6','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'    cnDatabase.Execute "Insert Into ElementMaster VALUES ('*00040','Danglar','Danglar','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'    cnDatabase.Execute "Insert Into ElementMaster VALUES ('*00041','Carton','Carton','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'    cnDatabase.Execute "Insert Into ElementMaster VALUES ('*00042','Carton [Inner]','Carton [Inner]','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'    cnDatabase.Execute "Insert Into ElementMaster VALUES ('*00043','Carton [Outer]','Carton [Outer]','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'    cnDatabase.Execute "Insert Into ElementMaster VALUES ('*00044','Card','Card','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'    cnDatabase.Execute "Insert Into ElementMaster VALUES ('*00045','Envelope','Envelope','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'
''Finish Size Master
'    cnDatabase.Execute "DELETE FROM FinishSizeChild Where Left(Code,1)='*'"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11011','*01039','16','16','*01017')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11012','*01030','16','16','*01031')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11012','*01064','32','16','*01031')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11013','*01039','16','16','*01017')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11014','*01039','16','16','*01017')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11015','*01055','16','16','*01028')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11016','*01048','16','16','*01017')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11017','*01051','16','16','*01028')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11018','*01058','16','16','*01028')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11019','*01056','16','16','*01028')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11020','*01028','8','16','*01028')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11020','*01060','16','16','*01028')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11021','*01067','16','16','*01031')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11033','*01039','16','16','*01017')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11023','*01031','8','16','*01031')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11023','*01067','16','16','*01031')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11024','*01033','8','16','*01033')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11024','*01068','16','16','*01033')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11025','*01068','16','16','*01017')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11026','*01037','8','16','*01017')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11026','*01070','16','16','*01017')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11027','*01054','8','16','*01017')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11028','*01072','16','16','*01017')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11029','*01038','8','16','*01017')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11029','*01072','16','16','*01017')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11030','*01055','12','24','*01028')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11031','*01039','8','16','*01017')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11032','*01046','8','16','*01017')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11034','*01048','8','16','*01017')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11035','*01063','12','24','*01031')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11036','*01048','8','16','*01017')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11037','*01048','8','16','*01017')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11038','*01055','8','16','*01028')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11039','*01067','12','24','*01031')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11040','*01050','8','16','*01028')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11041','*01058','8','16','*01028')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11042','*01027','4','8','*01031')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11043','*01070','12','24','*01017')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11044','*01060','8','16','*01028')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11046','*01060','8','16','*01028')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11047','*01039','6','12','*01017')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11049','*01073','16','16','*01017')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11050','*01068','8','16','*01031')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11051','*01055','6','12','*01028')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11052','*01072','6','12','*01017')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11053','*01060','4','8','*01028')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11054','*01068','4','8','*01031')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11055','*01070','6','12','*01017')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11004','*01073','4','8','*01017')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11004','*01039','2','2','*01012')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11004','*01012','1','1','*01012')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11057','*01055','8','16','*01028')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11048','*01028','4','8','*01029')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11048','*01059','8','16','*01029')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11058','*01058','8','16','*01029')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11045','*01063','8','16','*01031')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11085','*01060','12','24','*01029')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11005','*01011','2','2','*01011')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11092','*01065','8','16','*01029')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11091','*01028','4','8','*01029')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11091','*01067','8','16','*01029')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11094','*01028','8','16','*01028')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11094','*01068','16','16','*01028')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11022','*01072','16','16','*01031')"
'    cnDatabase.Execute "Insert Into FinishSizeChild VALUES ('*11095','*01045','8','8','*01047')"
'
''Genral Master
''Size Master
'        cnDatabase.Execute "DELETE FROM GeneralMaster Where Type ='1' AND Left(Code,1)='*'"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01001','05.25X10.00','05.25X10.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01002','10.00X29.00','10.00X29.00','1','0','000001',GetDate(),'NULL',NULL,'M','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01003','11.00X14.00','11.00X14.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01004','11.50X18.00','11.50X18.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01005','12.00X18.00','12.00X18.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01006','12.00X23.00','12.00X23.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01007','12.50X18.00','12.50X18.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01008','13.00X19.00','13.00X19.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01009','14.00X19.00','14.00X19.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01010','14.00X22.00','14.00X22.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01011','15.00X10.00 (CARD)','15.00X10.00 (CARD)','1','0','000001',GetDate(),'NULL',NULL,'M','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01012','15.00X20.00 (CARD)','15.00X20.00 (CARD)','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01013','15.00X21.00','15.00X21.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01014','15.00X27.50','15.00X27.50','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01015','15.50X20.00','15.50X20.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01016','15.50X20.50','15.50X20.50','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01017','15.50X21.00','15.50X21.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01018','15.50X21.50','15.50X21.50','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01019','16.00X20.00','16.00X20.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01020','16.00X20.50','16.00X20.50','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01021','16.00X22.00','16.00X22.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01022','16.00X24.00','16.00X24.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01023','16.00X25.00','16.00X25.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01024','16.00X30.00','16.00X30.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01025','16.50X10.50','16.50X10.50','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01026','17.00X22.00','17.00X22.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01027','17.00X24.00','17.00X24.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01028','18.00X23.00','18.00X23.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01029','18.00X23.00 (Card)','18.00X23.00 (Card)','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01030','18.00X24.00','18.00X24.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01031','18.00X25.00','18.00X25.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01032','19.00X20.00','19.00X20.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01033','19.00X25.00','19.00X25.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01034','19.00X38.00','19.00X38.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01035','20.00X24.00','20.00X24.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01036','20.00X25.00','20.00X25.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01037','20.00X26.00','20.00X26.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01038','20.00X28.00','20.00X28.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01039','20.00X30.00','20.00X30.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01040','20.00X30.00(A/P)','20.00X30.00(A/P)','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01041','20.00X30.00(Card)','20.00X30.00(Card)','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01042','20.00X31.00','20.00X31.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01043','20.50X24.00','20.50X24.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01044','20.50X31.00','20.50X31.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01045','21.00X29.70 (A4)','21.00X29.70 (A4)','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01046','21.00X30.00','21.00X30.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01047','21.00X30.00(CARD)','21.00X30.00(CARD)','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01048','21.00X31.00','21.00X31.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01049','21.00X32.00','21.00X32.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01050','21.00X33.00','21.00X33.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01051','21.00X34.00','21.00X34.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01052','21.00X35.00','21.00X35.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01053','21.50X28.50','21.50X28.50','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01054','22.00X28.00','22.00X28.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01055','22.00X32.00','22.00X32.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01056','22.00X34.00','22.00X34.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01057','23.00X30.00','23.00X30.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01058','23.00X33.00','23.00X33.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01059','23.00X35.00','23.00X35.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01060','23.00X36.00','23.00X36.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01061','23.00X36.00(A/P)','23.00X36.00(A/P)','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01062','23.00X36.00(Card)','23.00X36.00(Card)','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01063','24.00X34.00','24.00X34.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01064','24.00X36.00','24.00X36.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01065','24.13X24.13','24.13X24.13','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01066','25.00X30.00','25.00X30.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01067','25.00X36.00','25.00X36.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01068','25.00X38.00','25.00X38.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01069','26.00X38.00','26.00X38.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01070','26.00X40.00','26.00X40.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01071','28.00X35.00','28.00X35.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01072','28.00X40.00','28.00X40.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01073','30.00X40.00','30.00X40.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*01074','31.50X41.50','31.50X41.50','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'
''Item Group Master
'        cnDatabase.Execute "DELETE FROM GeneralMaster Where Type ='5' AND Left(Code,1)='*'"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*05001','Activity Book','Activity Book','5','0','000001',GetDate(),'NULL',NULL,'M','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*05002','Box','Box','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*05003','CARD','CARD','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*05004','CATALOGUE','CATALOGUE','5','0','000001',GetDate(),'NULL',NULL,'M','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*05005','General','General','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*05006','GRADE 1','GRADE 1','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*05007','GRADE 2','GRADE 2','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*05008','GRADE 3','GRADE 3','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*05009','GRADE 4','GRADE 4','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*05010','GRADE 5','GRADE 5','5','0','000001',GetDate(),'NULL',NULL,'M','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*05011','JUNIOR','JUNIOR','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*05012','LEVEL 1','LEVEL 1','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*05013','LEVEL 2','LEVEL 2','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*05014','LEVEL 3','LEVEL 3','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*05015','LEVEL 4','LEVEL 4','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*05016','LEVEL 5','LEVEL 5','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*05017','LEVEL A','LEVEL A','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*05018','LEVEL B','LEVEL B','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*05019','LEVEL C','LEVEL C','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*05020','NURSERY','NURSERY','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*05021','SECONDARY STD VI','SECONDARY STD','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*05022','SENIOR','SENIOR','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*05023','SET 1','SET 1','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*05024','Item Group','Item Group','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'
''Binding Type
'        cnDatabase.Execute "DELETE FROM GeneralMaster Where Type ='6' AND Left(Code,1)='*'"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*06001','Die_Cutting','Die_Cutting','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*06002','Die_Perforation','Die_Perforation','6','0','000001',GetDate(),'NULL',NULL,'M','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*06003','Hard Bound','Hard Bound','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*06004','Perfect Binding With Sewing','Perfect Binding With Sewing','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*06005','Perfect Binding With Sewing(CD-Insert)','Perfect Binding With Sewing(CD-Insert)','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*06006','Spiral Binding','Spiral Binding','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*06007','Wirro Binding','Wirro Binding','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*06008','Cutting & Packing','Cutting & Packing','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*06009','Cutting Only','Cutting Only','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*06010','Half Die Cut','Half Die Cut','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*06011','Loose','Loose','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*06012','Pad Gumming','Pad Gumming','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*06013','Pakki Binding','Pakki Binding','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*06014','Kachchi Binding','Kachchi Binding','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*06015','Center Pinning (BOX)','Center Pinning (BOX)','6','0','000001',GetDate(),'NULL',NULL,'M','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*06016','Center Pin Binding','Center Pin Binding','6','0','000001',GetDate(),'NULL',NULL,'M','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*06017','None','None','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*06018','Perfect Binding','Perfect Binding','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
''Finishing Type
'        cnDatabase.Execute "DELETE FROM GeneralMaster Where Type ='7' AND Left(Code,1)='*'"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07001','BOPP Gloss','BOPP Gloss','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07002','BOPP Matt','BOPP Matt','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07003','Box Packing','Box Packing','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07004','Center Pin Binding','Center Pin Binding','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07005','Counting & Fabrication','Counting & Fabrication','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07006','Creasing+Folding+Packing','Creasing+Folding+Packing','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07007','Cutting and Packing','Cutting and Packing','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07008','Cutting Leaflet Only','Cutting Leaflet Only','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07009','Die Cutting Charges','Die Cutting Charges','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07010','Die Making Charges','Die Making Charges','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07011','Digital Print','Digital Print','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07012','Embossing','Embossing','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07013','Foiling Charges','Foiling Charges','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07014','Folding & Packing','Folding & Packing','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07015','Graning','Graning','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07016','Half Die Cutting Charges','Half Die Cutting Charges','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07017','Hardbound Binding','Hardbound Binding','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07018','Hologram','Hologram','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07019','Matt + Spot UV','Matt + Spot UV','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07020','Matt + Spot UV + Foiling + Embossing','Matt + Spot UV + Foiling + Embossing','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07021','Matt + Spot UV+Glitter UV','Matt + Spot UV+Glitter UV','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07022','Matt Both Side','Matt Both Side','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07023','MINI Offset JOB','MINI Offset JOB','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07024','None','None','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07025','Packing Shrink','Packing Shrink','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07026','Paper Cost','Paper Cost','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07027','Pasting Charges','Pasting Charges','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07028','Perfect Binding','Perfect Binding','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07029','Plate','Plate','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07030','Printing 4 Col','Printing 4 Col','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07031','PVC','PVC','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07032','Spot UV','Spot UV','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07033','Thermal Matt','Thermal Matt','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07034','UV Hybraid','UV Hybraid','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*07035','Varnising','Varnising','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
''Project Member/Editorial Team Master
'        cnDatabase.Execute "DELETE FROM GeneralMaster Where Type ='8' AND Left(Code,1)='*'"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*08002','Author_ABC','Author_ABC','8','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*08003','DTP_ABC','DTP_ABC','8','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*08005','Editor_ABC','Editor_ABC','8','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*08007','Graphic_ABC','Graphic_ABC','8','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*08008','PPQ_ABC','PPQ_ABC','8','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*08009','Processing_S.R.K','Processing_S.R.K','8','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*08010','Proof Reader_ABC','Proof Reader_ABC','8','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*08011','Type Setting_ABC','Type Setting_Sanjay Khanna','8','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
''Plate Master
'        cnDatabase.Execute "DELETE FROM GeneralMaster Where Type ='9' AND Left(Code,1)='*'"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*09001','CTP_Plates','CTP_Plates','9','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*09002','Nagative-Cut Pieces','Nagative-Cut Pieces','9','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*09003','Nagative-One Pieces','Nagative-One Pieces','9','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
''Size Group Master
'        cnDatabase.Execute "DELETE FROM GeneralMaster Where Type ='10' AND Left(Code,1)='*'"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*10016','Extra Large-28''''X40''''','Extra Large-28''''X40''''','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*10018','Extra Large-28''''X40''''-(Card)','Extra Large-28''''X40''''-(Card)','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*10001','Extra Large-28''''X40''''-A/P','Extra Large-28''''X40''''-A/P','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*10002','Extra Large-28''''X40''''-A/P_SPL','Extra Large-28''''X40''''-A/P_SPL','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*10003','Extra Large-30''''X40''''','Extra Large-30''''X40''''','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*10004','Extra Large-30''''X40''''-(A/P)','Extra Large-30''''X40''''-(A/P)','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*10005','Extra Large-30''''X40''''-(Card)','Extra Large-30''''X40''''-(Card)','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*10006','LARGE-23''''X36''''','LARGE-23''''X36''''','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*10007','LARGE-23''''X36''''-(A/P)','LARGE-23''''X36''''-(A/P)','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*10008','LARGE-23''''X36''''-(Card)','LARGE-23''''X36''''-(Card)','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*10019','Little-11.50''''X18.00''''','Little-11.50''''X18.00''''','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*10021','Little-11.50''''X18.00''''-(A/P)','Little-11.50''''X18.00''''-(A/P)','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*10020','Little-11.50''''X18.00''''-(Card)','Little-11.50''''X18.00''''-(Card)','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*10009','Medium-20''''X30''''','Medium-20''''X30''''','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*10010','Medium-20''''X30''''(A/P)','Medium-20''''X30''''(A/P)','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*10011','Medium-20''''X30''''(Card)','Medium-20''''X30''''(Card)','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*10012','Small-19''''X26''''','Small-19''''X26''''','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*10017','Small-19''''X26''''-(A/P)','Small-19''''X26''''-(A/P)','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*10013','Small-19''''X26''''(Card)','Small-19''''X26''''(Card)','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*10014','Web-508mm','Web-508mm','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*10015','Web-578mm','Web-578mm','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
''Finish Size Master
'        cnDatabase.Execute "DELETE FROM GeneralMaster Where Type ='11' AND Left(Code,1)='*'"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11001','05.25x10.00','05.25x10.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11002','12.00X18.00','12.00X18.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11003','12.00X23.00','12.00X23.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11004','14.00X19.00','14.00X19.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11005','15.00X10.00 (CARD)','15.00X10.00 (CARD)','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11006','15.50X20.50','15.50X20.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11007','16.00x20.00','16.00x20.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11008','16.00X24.00','16.00X24.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11009','16.50X10.50','16.50X10.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11010','17.00X22.00','17.00X22.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11011','04.00X06.87','04.00X06.87','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11012','04.25X05.50','04.25X05.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11013','04.25X07.00','04.25X07.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11014','04.37X07.00','04.37X07.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11015','04.72X07.48','04.72X07.48','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11016','05.00X07.00','05.00X07.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11017','05.00X08.00','05.00X08.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11018','05.06X07.81','05.06X07.81','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11019','05.25X08.00','05.25X08.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11020','05.50X08.50','05.50X08.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11021','05.83X08.27','05.83X08.27','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11022','06.00X08.25','06.00X08.25','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11023','06.00X08.50','06.00X08.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11024','06.00X09.00','06.00X09.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11025','06.14X09.21','06.14X09.21','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11026','06.25X09.50','06.25X09.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11027','06.63X10.25','06.63X10.25','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11028','06.69X09.61','06.69X09.61','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11029','06.75X09.50','06.75X09.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11030','07.00X07.00','07.00X07.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11031','07.00X09.00','07.00X09.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11032','07.00X10.00','07.00X10.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11033','07.25X09.50','07.25X09.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11034','07.44X09.69','07.44X09.69','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11035','07.50X07.50','07.50X07.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11036','07.50X09.25','07.50X09.25','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11037','07.50X09.50','07.50X09.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11038','07.75X10.50','07.75X10.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11039','08.00X08.00','08.00X08.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11040','08.00X10.00','08.00X10.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11041','08.00X10.88','08.00X10.88','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11042','08.00X11.25','08.00X11.25','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11043','08.25X08.25','08.25X08.25','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11044','08.25X11.00','08.25X11.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11045','08.27X11.69','08.27X11.69','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11046','08.50X08.50','08.50X08.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11047','08.50X09.00','08.50X09.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11048','08.50X11.00','08.50X11.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11049','09.00X07.00','09.00X07.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11050','09.00X12.00','09.00X12.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11051','10.00X10.00','10.00X10.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11052','11.00X13.00','11.00X13.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11053','11.00X17.00','11.00X17.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11054','11.00X18.00','11.00X18.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11055','12.00X12.00','12.00X12.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11056','18.00X23.00','18.00X23.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11057','07.75X11.25','07.75X11.25','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11058','08.00X11.00','08.00X11.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11059','04.50X01.75','04.50X01.75','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11060','11.00x15.75','11.00x15.75','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11061','11.00X16.00','11.00X16.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11062','08.25X11.75','08.25X11.75','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11063','04.00X06.00','04.00X06.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11064','20.00X30.00','20.00X30.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11065','17.50X22.50','17.50X22.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11066','11.50X08.00','11.50X08.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11067','21.00X31.00','21.00X31.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11068','05.30X08.30','05.30X08.30','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11069','11.50X10.75','11.50X10.75','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11070','08.50X10.75','08.50X10.75','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11071','02.00X03.00','02.00X03.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11072','11.50X07.00','11.50X07.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11073','05.50X19.00','05.50X19.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11074','10.25X07.50','10.25X07.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11075','07.50X13.75','07.50X13.75','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11076','07.00X02.50','07.00X02.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11077','06.50X09.50','06.50X09.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11078','04.00x07.50','04.00x07.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11079','23.00X36.00','23.00X36.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11080','15.00X20.00','15.00X20.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11081','25.00X36.00','25.00X36.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11082','09.00X14.00','09.00X14.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11083','05.25X07.00','05.25X07.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11084','08.00X10.50','08.00X10.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11085','07.50X08.50','07.50X08.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11086','03.25X04.75','03.25X04.75','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11087','09.75X11.00','09.75X11.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11088','13.50X18.00','13.50X18.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11089','07.62X11.00','07.62X11.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11090','07.36X11.00','07.36X11.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11091','08.26X11.69','08.26X11.69','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11092','09.50X09.50','09.50X09.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11093','11.69X05.20','11.69X05.20','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11094','05.75X08.25','05.75X08.25','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*11095','21.00X29.70','21.00X29.70','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
''Genral Accounts Groups
'        cnDatabase.Execute "DELETE FROM GeneralMaster Where Type ='12' AND Left(Code,1)='*'"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*12002','Account Group','Account Group','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26031')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*99996','Transporter','Transporter','12','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*99997','Packer','Transporter','12','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*99998','Deliverer','Deliverer','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26030')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*99999','Material Centre','Material Centre','12','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*12001','Binders','Binders','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26030')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*12003','Box Supplier','Box Supplier','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26030')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*12004','CD Suppliers','CD Suppliers','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26030')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*12005','FG Godown','FG Godown','12','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*12006','Laminator','Laminator','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26030')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*12007','Packaging Supplier','Packaging Supplier','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26030')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*12008','Paper Suppliers','Paper Suppliers','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26030')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*12009','Printer','Printer','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26030')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*12010','Printer & Binder','Printer & Binder','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26030')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*12011','Printer, Binder & Laminator','Printer, Binder & Laminator','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26030')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*12012','Processor & Printer','Processor & Printer','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26030')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*12013','Processor, Printer & Laminator','Processor, Printer & Laminator','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26030')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*12014','UFG Godown','UFG Godown','12','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*12015','Publisher','Publisher','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26031')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*12016','Clients','Clients','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26031')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*12017','Cons. Supplier','Cons. Supplier','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26030')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*12018','Plate Maker','Plate Maker','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26030')"
''Departments
'        cnDatabase.Execute "DELETE FROM GeneralMaster Where Type ='13' AND Left(Code,1)='*'"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*13001','Editorial Department','Editorial Department','13','0','000001',GetDate(),'NULL',NULL,'M','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*13002','Production Department','Production Department','13','0','000001',GetDate(),'NULL',NULL,'M','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*13003','Sales Department','Sales Department','13','0','000001',GetDate(),'NULL',NULL,'M','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*13004','Contracts Department and Legal Department','Contracts Department and Legal Department','13','0','000001',GetDate(),'NULL',NULL,'M','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*13005','Managing Editorial and Production','Managing Editorial and Production','13','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*13006','Creative Departments','Creative Departments','13','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*13007','Subsidiary Rights Departments','Subsidiary Rights Departments','13','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*13008','Marketing, Promotion, and Advertising Departments','Marketing, Promotion, and Advertising Departments','13','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*13009','Publicity Department','Publicity Department','13','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*13010','Publisher Website Maintenance','Publisher Website Maintenance','13','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*13011','Finance and Accounting','Finance and Accounting','13','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*13012','Information Technology (IT)','Information Technology (IT)','13','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*13013','Human Resources (HR)','Human Resources (HR)','13','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
''Designation
'        cnDatabase.Execute "DELETE FROM GeneralMaster Where Type ='14' AND Left(Code,1)='*'"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*14001','Editor-in-Chief','Editor-in-Chief','14','0','000001',GetDate(),'NULL',NULL,'M','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*14002','Managing editor','Managing editor','14','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*14003','Editors','Editors','14','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*14004','Author/Writers','Author/Writers','14','0','000001',GetDate(),'NULL',NULL,'M','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*14005','Fact-checkers','Fact-checkers','14','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*14006','Graphic Designer','Graphic Designer','14','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*14007','Production manager','Production manager','14','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*14008','DTP-Operator','DTP-Operator','14','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*14009','Proof Reader','Proof Reader','14','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
''Paper Unit Master
'        cnDatabase.Execute "DELETE FROM GeneralMaster Where Type ='15' AND Left(Code,1)='*'"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*15001','Gross','Gross','15','144','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*15002','Packet(100)','Packet(100)','15','100','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*15003','Packet(150)','Packet(150)','15','150','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*15004','Ream','Ream','15','500','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*15005','Reel','Reel','15','500','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*15006','Bundle (700)','Bundle (700)','15','700','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*15007','Packet(200)','Packet(200)','15','200','000001',GetDate(),'NULL',NULL,'M','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*15008','PACKET','PACKET','15','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*15009','Sheet','Sheet','15','1','000001',GetDate(),'NULL',NULL,'M','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*15010','Packet (250)','Packet (250)','15','250','000001',GetDate(),'NULL',NULL,'M','N','NULL')"
''Paper Unit Master
'        cnDatabase.Execute "DELETE FROM GeneralMaster Where Type ='16' AND Left(Code,1)='*'"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*16001','Coated Matt','Coated Matt','16','0.95','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*16002','Coated Gloss','Coated Gloss','16','0.9','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*16003',' Uncoated','Uncoated','16','1.35','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*16004','High Bulk','High Bulk','16','1.4','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
''Narration
'        cnDatabase.Execute "DELETE FROM GeneralMaster Where Type ='17' AND Left(Code,1)='*'"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*17001','1. Printing & Finishing Charges of','Printing & Finishing Charges of','17','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*17002','1. Text Printing Charges of','Text Printing Charges of','17','2','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*17003','2. Title Printing Charges of','Title Printing Charges of','17','3','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*17004','3. Combo Title Printing Charges of','Combo Title Printing Charges of','17','4','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*17005','4. Finishing Charges of','Finishing Charges of','17','5','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*17006','5. Binding Charges of','Binding Charges of','17','6','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*17007','7. Title Printing & Finishing Charges of','Title Printing & Finishing Charges of','17','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*17008','6. Text Printing & Finishing Charges of','Text Printing & Finishing Charges of','17','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*17009','8. Unit Cost Charges of','Unit Cost Charges of','17','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*17010','9. Unit Cost','.','17','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*17011','10 Lamination Charges','Lamination Charges','17','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*17012','11 Printed Book','Printed Book','17','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*17013','RGN-Cool Luke-Energing Exfoliator (450 g','RGN-Cool Luke-Energing Exfoliator (450 g','17','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
''HSN MASTER
'        cnDatabase.Execute "DELETE FROM GeneralMaster Where Type ='18' AND Left(Code,1)='*'"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*18001','998812','998812','18','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*18002','998912','998912','18','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*18003','4901','4901','18','0','000001',GetDate(),'NULL',NULL,'M','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*18004','49011010','49011010','18','0','000001',GetDate(),'NULL',NULL,'M','N','NULL')"
''Elements MASTER
'        cnDatabase.Execute "DELETE FROM GeneralMaster Where Type ='19'"
'        'eLEMENT mASTER mOVED TO eLEMENT mASTER
''Calculation Units MASTER
'        cnDatabase.Execute "DELETE FROM GeneralMaster Where Type ='20' AND Left(Code,1)='*'"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*20001','Per Unit','Per Unit','20','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*20002','Per Inch²','Per Inch²','20','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*20003','100 Inch²','100 Inch²','20','100','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*20004','1000 Inch²','1000 Inch²','20','1000','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*20005','Per 1000','Per 1000','20','1000','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
''Machine Master
'        cnDatabase.Execute "DELETE FROM GeneralMaster Where Type ='21' AND Left(Code,1)='*'"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*21046','Machine To Be Decide','Machine To Be Decide','21','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*21047','RYOBI - 4 Col','RYOBI - 4 Col','21','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*21048','SM 102 28x40','SM 102 28x40','21','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*21049','SM 74 20x29','SM 74 20x29','21','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*21050','Heidel 2 Col','Heidel 2 Col','21','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
''General  Unit Master
'        cnDatabase.Execute "DELETE FROM GeneralMaster Where Type ='25' AND Left(Code,1)='*'"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*25001','Kilogram','kg.','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*25002','Gram','gm.','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*25003','Milligram','mg.','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*25004','Liter','ltr.','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*25005','Milliliter','ml.','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*25006','Feet','ft.','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*25007','Inch','in.','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*25008','Meter','mtr.','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*25009','Centimeter','cm.','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*25010','Millimeter','mm.','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*25011','Piece','pec.','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*25012','Bags','bags','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*25013','Roll','roll','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*25014','Sets','sets','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*25015','Packets','packets','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*25016','Gross','gross','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*25017','Dozen','dozen','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*25018','Tonn','tonn','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
''Account Group
'        cnDatabase.Execute "DELETE FROM GeneralMaster Where Type ='26' AND Left(Code,1)='*'"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26001','Profit & Loss','Profit & Loss','26','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26002','Revenue Accounts','Revenue Accounts','26','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26003','Stock-in-hand','Stock-in-hand','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26008')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26004','Bank Accounts','Bank Accounts','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26008')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26005','Bank O/D Account','Bank O/D Account','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26022')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26006','Capital Account','Capital Account','26','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26007','Cash-in-hand','Cash-in-hand','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26008')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26008','Current Assets','Current Assets','26','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26009','Current Liabilities','Current Liabilities','26','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26010','Depreciation Res On Machine','Depreciation Res On Machine','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26016')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26011','Duties & Taxes','Duties & Taxes','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26009')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26012','Expenses (Direct/Mfg.)','Expenses (Direct/Mfg.)','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26002')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26013','Expenses (Indirect/Admn.)','Expenses (Indirect/Admn.)','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26002')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26014','File-Sundry Creditors','File-Sundry Creditors','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26030')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26015','File-Sundry Debtors','File-Sundry Debtors','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26031')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26016','Fixed Assets','Fixed Assets','26','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26017','Income (Direct/Opr.)','Income (Direct/Opr.)','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26002')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26018','Income (Indirect)','Income (Indirect)','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26002')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26019','Income Tax Advance','Income Tax Advance','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26021')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26020','Investments','Investments','26','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26021','Loans & Advances (Asset)','Loans & Advances (Asset)','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26008')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26022','Loans (Liability)','Loans (Liability)','26','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26023','Pre-Operative Expenses','Pre-Operative Expenses','26','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26024','Provisions/Expenses Payable','Provisions/Expenses Payable','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26009')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26025','Purchase','Purchase','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26002')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26026','Reserves & Surplus','Reserves & Surplus','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26006')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26027','Sale','Sale','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26002')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26028','Secured Loans','Secured Loans','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26022')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26029','Securities & Deposits (Asset)','Securities & Deposits (Asset)','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26008')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26030','Sundry Creditors','Sundry Creditors','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26009')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26031','Sundry Debtors','Sundry Debtors','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26008')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26032','Suspense Account','Suspense Account','26','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'        cnDatabase.Execute "Insert Into GeneralMaster VALUES ('*26033','Unsecured Loans','Unsecured Loans','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26022')"
''Paper Master
'        cnDatabase.Execute "DELETE FROM PaperMaster Where Left(Code,1)='*'"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00001','Art Card-200gsm-20.00X30.00in²-(50.80X76.20cm²)-7.742kg-Gloss','Art Card-200gsm-20.00X30.00in²-(50.80X76.20cm²)-7.742kg-Gloss','S','B','50.8','76.2','20','30','*15002','200','Art Card','Gloss','7.742','6','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00002','Art Card-210gsm-20.00X30.00in²-(50.80X76.20cm²)-8.129kg-Gloss','Art Card-210gsm-20.00X30.00in²-(50.80X76.20cm²)-8.129kg-Gloss','S','B','50.8','76.2','20','30','*15002','210','Art Card','Gloss','8.129','6','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00003','Art Card-220gsm-20.00X30.00in²-(50.80X76.20cm²)-8.516kg-Gloss','Art Card-220gsm-20.00X30.00in²-(50.80X76.20cm²)-8.516kg-Gloss','S','B','50.8','76.2','20','30','*15002','220','Art Card','Gloss','8.516','6','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00004','Art Card-250gsm-20.00X30.00in²-(50.80X76.20cm²)-9.677kg-Gloss','Art Card-250gsm-20.00X30.00in²-(50.80X76.20cm²)-9.677kg-Gloss','S','B','50.8','76.2','20','30','*15002','250','Art Card','Gloss','9.677','5','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00005','Art Card-200gsm-23.00X36.00in²-(58.42X91.44cm²)-10.684kg-Gloss','Art Card-200gsm-23.00X36.00in²-(58.42X91.44cm²)-10.684kg-Gloss','S','B','58.42','91.44','23','36','*15002','200','Art Card','Gloss','10.684','5','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00006','Art Card-210gsm-23.00X36.00in²-(58.42X91.44cm²)-11.218kg-Gloss','Art Card-210gsm-23.00X36.00in²-(58.42X91.44cm²)-11.218kg-Gloss','S','B','58.42','91.44','23','36','*15002','210','Art Card','Gloss','11.218','4','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00007','Art Card-220gsm-23.00X36.00in²-(58.42X91.44cm²)-11.752kg-Gloss','Art Card-220gsm-23.00X36.00in²-(58.42X91.44cm²)-11.752kg-Gloss','S','B','58.42','91.44','23','36','*15002','220','Art Card','Gloss','11.752','4','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00008','Art Card-250gsm-23.00X36.00in²-(58.42X91.44cm²)-13.355kg-Gloss','Art Card-250gsm-23.00X36.00in²-(58.42X91.44cm²)-13.355kg-Gloss','S','B','58.42','91.44','23','36','*15002','250','Art Card','Gloss','13.355','4','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00009','Art Paper-70gsm-20.00X30.00in²-(50.80X76.20cm²)-13.548kg-Gloss','Art Paper-70gsm-20.00X30.00in²-(50.80X76.20cm²)-13.548kg-Gloss','S','P','50.8','76.2','20','30','*15004','70','Art Paper','Gloss','13.548','4','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00010','Art Paper-80gsm-20.00X30.00in²-(50.80X76.20cm²)-15.484kg-Gloss','Art Paper-80gsm-20.00X30.00in²-(50.80X76.20cm²)-15.484kg-Gloss','S','P','50.8','76.2','20','30','*15004','80','Art Paper','Gloss','15.484','3','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00011','Art Paper-90gsm-20.00X30.00in²-(50.80X76.20cm²)-17.419kg-Gloss','Art Paper-90gsm-20.00X30.00in²-(50.80X76.20cm²)-17.419kg-Gloss','S','P','50.8','76.2','20','30','*15004','90','Art Paper','Gloss','17.419','3','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00012','Art Paper-100gsm-20.00X30.00in²-(50.80X76.20cm²)-19.355kg-Gloss','Art Paper-100gsm-20.00X30.00in²-(50.80X76.20cm²)-19.355kg-Gloss','S','P','50.8','76.2','20','30','*15004','100','Art Paper','Gloss','19.355','3','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00013','Art Paper-130gsm-20.00X30.00in²-(50.80X76.20cm²)-25.161kg-Gloss','Art Paper-130gsm-20.00X30.00in²-(50.80X76.20cm²)-25.161kg-Gloss','S','P','50.8','76.2','20','30','*15004','130','Art Paper','Gloss','25.161','2','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00014','Art Paper-170gsm-20.00X30.00in²-(50.80X76.20cm²)-32.903kg-Gloss','Art Paper-170gsm-20.00X30.00in²-(50.80X76.20cm²)-32.903kg-Gloss','S','P','50.8','76.2','20','30','*15004','170','Art Paper','Gloss','32.903','2','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00015','Art Paper-70gsm-23.00X36.00in²-(58.42X91.44cm²)-18.697kg-Gloss','Art Paper-70gsm-23.00X36.00in²-(58.42X91.44cm²)-18.697kg-Gloss','S','P','58.42','91.44','23','36','*15004','70','Art Paper','Gloss','18.697','3','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00016','Art Paper-80gsm-23.00X36.00in²-(58.42X91.44cm²)-21.368kg-Gloss','Art Paper-80gsm-23.00X36.00in²-(58.42X91.44cm²)-21.368kg-Gloss','S','P','58.42','91.44','23','36','*15004','80','Art Paper','Gloss','21.368','2','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00017','Art Paper-90gsm-23.00X36.00in²-(58.42X91.44cm²)-24.039kg-Gloss','Art Paper-90gsm-23.00X36.00in²-(58.42X91.44cm²)-24.039kg-Gloss','S','P','58.42','91.44','23','36','*15004','90','Art Paper','Gloss','24.039','2','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00018','Art Paper-100gsm-23.00X36.00in²-(58.42X91.44cm²)-26.71kg-Gloss','Art Paper-100gsm-23.00X36.00in²-(58.42X91.44cm²)-26.71kg-Gloss','S','P','58.42','91.44','23','36','*15004','100','Art Paper','Gloss','26.71','2','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00019','Art Paper-130gsm-23.00X36.00in²-(58.42X91.44cm²)-34.723kg-Gloss','Art Paper-130gsm-23.00X36.00in²-(58.42X91.44cm²)-34.723kg-Gloss','S','P','58.42','91.44','23','36','*15004','130','Art Paper','Gloss','34.723','1','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00020','Art Paper-170gsm-23.00X36.00in²-(58.42X91.44cm²)-45.406kg-Gloss','Art Paper-170gsm-23.00X36.00in²-(58.42X91.44cm²)-45.406kg-Gloss','S','P','58.42','91.44','23','36','*15004','170','Art Paper','Gloss','45.406','1','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00021','Paper-60gsm-20.00X30.00in²-(50.80X76.20cm²)-11.613kg-Maplitho','Paper-60gsm-20.00X30.00in²-(50.80X76.20cm²)-11.613kg-Maplitho','S','P','50.8','76.2','20','30','*15004','60','Paper','Maplitho','11.613','4','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00022','Paper-64gsm-20.00X30.00in²-(50.80X76.20cm²)-12.387kg-Maplitho','Paper-64gsm-20.00X30.00in²-(50.80X76.20cm²)-12.387kg-Maplitho','S','P','50.8','76.2','20','30','*15004','64','Paper','Maplitho','12.387','4','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00023','Paper-70gsm-20.00X30.00in²-(50.80X76.20cm²)-13.548kg-Maplitho','Paper-70gsm-20.00X30.00in²-(50.80X76.20cm²)-13.548kg-Maplitho','S','P','50.8','76.2','20','30','*15004','70','Paper','Maplitho','13.548','4','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00024','Paper-80gsm-20.00X30.00in²-(50.80X76.20cm²)-15.484kg-Maplitho','Paper-80gsm-20.00X30.00in²-(50.80X76.20cm²)-15.484kg-Maplitho','S','P','50.8','76.2','20','30','*15004','80','Paper','Maplitho','15.484','3','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00025','Paper-90gsm-20.00X30.00in²-(50.80X76.20cm²)-17.419kg-Maplitho','Paper-90gsm-20.00X30.00in²-(50.80X76.20cm²)-17.419kg-Maplitho','S','P','50.8','76.2','20','30','*15004','90','Paper','Maplitho','17.419','3','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00026','Paper-100gsm-20.00X30.00in²-(50.80X76.20cm²)-19.355kg-Maplitho','Paper-100gsm-20.00X30.00in²-(50.80X76.20cm²)-19.355kg-Maplitho','S','P','50.8','76.2','20','30','*15004','100','Paper','Maplitho','19.355','3','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00027','Paper-120gsm-20.00X30.00in²-(50.80X76.20cm²)-23.226kg-Maplitho','Paper-120gsm-20.00X30.00in²-(50.80X76.20cm²)-23.226kg-Maplitho','S','P','50.8','76.2','20','30','*15004','120','Paper','Maplitho','23.226','2','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00028','Paper-60gsm-23.00X36.00in²-(58.42X91.44cm²)-16.026kg-Maplitho','Paper-60gsm-23.00X36.00in²-(58.42X91.44cm²)-16.026kg-Maplitho','S','P','58.42','91.44','23','36','*15004','60','Paper','Maplitho','16.026','3','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00029','Paper-64gsm-23.00X36.00in²-(58.42X91.44cm²)-17.094kg-Maplitho','Paper-64gsm-23.00X36.00in²-(58.42X91.44cm²)-17.094kg-Maplitho','S','P','58.42','91.44','23','36','*15004','64','Paper','Maplitho','17.094','3','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00030','Paper-70gsm-23.00X36.00in²-(58.42X91.44cm²)-18.697kg-Maplitho','Paper-70gsm-23.00X36.00in²-(58.42X91.44cm²)-18.697kg-Maplitho','S','P','58.42','91.44','23','36','*15004','70','Paper','Maplitho','18.697','3','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00031','Paper-80gsm-23.00X36.00in²-(58.42X91.44cm²)-21.368kg-Maplitho','Paper-80gsm-23.00X36.00in²-(58.42X91.44cm²)-21.368kg-Maplitho','S','P','58.42','91.44','23','36','*15004','80','Paper','Maplitho','21.368','2','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00032','Paper-90gsm-23.00X36.00in²-(58.42X91.44cm²)-24.039kg-Maplitho','Paper-90gsm-23.00X36.00in²-(58.42X91.44cm²)-24.039kg-Maplitho','S','P','58.42','91.44','23','36','*15004','90','Paper','Maplitho','24.039','2','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00033','Paper-100gsm-23.00X36.00in²-(58.42X91.44cm²)-26.71kg-Maplitho','Paper-100gsm-23.00X36.00in²-(58.42X91.44cm²)-26.71kg-Maplitho','S','P','58.42','91.44','23','36','*15004','100','Paper','Maplitho','26.71','2','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00034','Paper-120gsm-23.00X36.00in²-(58.42X91.44cm²)-32.052kg-Maplitho','Paper-120gsm-23.00X36.00in²-(58.42X91.44cm²)-32.052kg-Maplitho','S','P','58.42','91.44','23','36','*15004','120','Paper','Maplitho','32.052','2','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00035','SBS-200gsm-20.00X30.00in²-(50.80X76.20cm²)-7.742kg-C1S','SBS-200gsm-20.00X30.00in²-(50.80X76.20cm²)-7.742kg-C1S','S','B','50.8','76.2','20','30','*15002','200','SBS','C1S','7.742','6','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00036','SBS-210gsm-20.00X30.00in²-(50.80X76.20cm²)-8.129kg-C1S','SBS-210gsm-20.00X30.00in²-(50.80X76.20cm²)-8.129kg-C1S','S','B','50.8','76.2','20','30','*15002','210','SBS','C1S','8.129','6','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00037','SBS-220gsm-20.00X30.00in²-(50.80X76.20cm²)-8.516kg-C1S','SBS-220gsm-20.00X30.00in²-(50.80X76.20cm²)-8.516kg-C1S','S','B','50.8','76.2','20','30','*15002','220','SBS','C1S','8.516','6','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00038','SBS-250gsm-20.00X30.00in²-(50.80X76.20cm²)-9.677kg-C1S','SBS-250gsm-20.00X30.00in²-(50.80X76.20cm²)-9.677kg-C1S','S','B','50.8','76.2','20','30','*15002','250','SBS','C1S','9.677','5','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00039','SBS-200gsm-23.00X36.00in²-(58.42X91.44cm²)-10.684kg-C1S','SBS-200gsm-23.00X36.00in²-(58.42X91.44cm²)-10.684kg-C1S','S','B','58.42','91.44','23','36','*15002','200','SBS','C1S','10.684','5','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00040','SBS-210gsm-23.00X36.00in²-(58.42X91.44cm²)-11.218kg-C1S','SBS-210gsm-23.00X36.00in²-(58.42X91.44cm²)-11.218kg-C1S','S','B','58.42','91.44','23','36','*15002','210','SBS','C1S','11.218','4','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00041','SBS-220gsm-23.00X36.00in²-(58.42X91.44cm²)-11.752kg-C1S','SBS-220gsm-23.00X36.00in²-(58.42X91.44cm²)-11.752kg-C1S','S','B','58.42','91.44','23','36','*15002','220','SBS','C1S','11.752','4','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into PaperMaster VALUES ('*00042','SBS-250gsm-23.00X36.00in²-(58.42X91.44cm²)-13.355kg-C1S','SBS-250gsm-23.00X36.00in²-(58.42X91.44cm²)-13.355kg-C1S','S','B','58.42','91.44','23','36','*15002','250','SBS','C1S','13.355','4','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N')"
'
''Size Group Master
'        cnDatabase.Execute "DELETE FROM SizeGroupChild Where Left(Code,1)='*'"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10003','*01067')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10003','*01068')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10003','*01070')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10003','*01072')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10003','*01073')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10007','*01061')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10011','*01047')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10006','*01050')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10006','*01051')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10006','*01056')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10006','*01058')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10006','*01060')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10006','*01063')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10006','*01064')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10006','*01059')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10012','*01017')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10012','*01020')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10012','*01021')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10012','*01027')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10012','*01028')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10012','*01030')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10012','*01031')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10012','*01033')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10012','*01013')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10013','*01012')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10013','*01015')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10013','*01016')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10013','*01019')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10013','*01029')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10013','*01018')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10018','*01069')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10009','*01036')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10009','*01037')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10009','*01038')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10009','*01039')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10009','*01046')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10009','*01048')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10009','*01054')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10009','*01057')"
'        cnDatabase.Execute "Insert Into SizeGroupChild VALUES ('*10020','*01011')"
''Tax Master
'        cnDatabase.Execute "DELETE FROM TaxMaster Where Left(Code,1)='*'"
'        cnDatabase.Execute "Insert Into TaxMaster VALUES ('*00001','Local GST 12%','Local GST 12%','L','6','6',0,'000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into TaxMaster VALUES ('*00002','IGST 12%','IGST 12%','I','0','0',12,'000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into TaxMaster VALUES ('*00003','IGST 5%','IGST 5%','I','0','0',5,'000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into TaxMaster VALUES ('*00004','Local GST 5%','Local GST 5%','L','2.5','2.5',0,'000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into TaxMaster VALUES ('*00005','Local GST 18%','Local GST 18%','L','9','9',0,'000006',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into TaxMaster VALUES ('*00006','IGST 18%','IGST 18%','I','0','0',18,'000006',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into TaxMaster VALUES ('*00007','Local GST NIL','Local GST NIL','L','0','0',0,'000001',GetDate(),'NULL',NULL,'N','N')"
'        cnDatabase.Execute "Insert Into TaxMaster VALUES ('*00008','IGST NIL','IGST NIL','I','0','0',0,'000001',GetDate(),'NULL',NULL,'N','N')"
''Vch Series Master
'        cnDatabase.Execute "DELETE FROM VchSeriesMaster Where Left(Code,1)='*'"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*00101','Main','01PF','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/Purc','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*00102','Main','01PU','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/PrJU','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*00103','Main','01PC','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/PrJC','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*00104','Main','01PJ','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/PrJW','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*00201','Main','02OF','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/PrRt','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*00202','Main','02OU','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/PrRtJU','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*00203','Main','02OC','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/PrRtJC','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*00204','Main','02OJ','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/PrRtJW','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*00301','Main','03TF','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/SlRt','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*00302','Main','03TU','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/SlRtJU','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*00303','Main','03TC','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/SlRtJC','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*00304','Main','03TJ','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/SlRtJW','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*00401','Main','04SF','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/Sale','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*00402','Main','04SU','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/SlJU','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*00403','Main','04SC','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/SlJC','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*00404','Main','04SJ','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/SlJW','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*00501','Main','05RF','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/MtRc','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*00502','Main','05FR','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/MtRcJW','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*00601','Main','06IF','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/PrRtC','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*00602','Main','06FI','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/PrRtCJW','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*00701','Main','07RF','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/SlRtC','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*00702','Main','07FR','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/SlRtCJW','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*00801','Main','08IF','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/MtIs','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*00802','Main','08FI','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/MtIsJW','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*01701','Main','17PO','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/PO','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*01801','Main','18SO','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/SO','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*01901','Main','19ST','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/STrn','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*02001','Main','20JR','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/SJrnl','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*02101','Main','21JR','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/SJrnl','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*02201','Main','22JR','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/SJrnl','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*02301','Main','23PQ','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/PQ','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*02302','Main','23UZ','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/PQU','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*02303','Main','23CZ','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/PQC','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*02304','Main','23JZ','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/PQJ','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*02305','Main','24SQ','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/SQ','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*02306','Main','24UQ','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/SQU','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*02307','Main','24CQ','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/SQC','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*02308','Main','24JQ','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/SQJ','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*05101','Main','51PI','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/Pymt','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*05201','Main','52PR','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/Rcpt','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*05301','Main','53JE','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/Jrnl','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*05401','Main','54CE','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/Cntr','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*05501','Main','55CN','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/CrNt','A')"
'        cnDatabase.Execute "Insert Into VchSeriesMaster VALUES ('*05601','Main','56DN','" & Trim(FrmCompanyMaster.Text15.Text) & "/" & "','/DrNt','A')"
''CompChild
'        cnDatabase.Execute "DELETE FROM CompChild "
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','01','1. Please send two copies of invoice.','2. Please notify us immediately if ','you are unable to ship as specified.','3. Enter this order in accordance, with the price,terms, ','delivery method and specification Listed above.','4. All disputes are subject to Our Jurisdiction Only','','SEPL/Pur/','/20-21','Purchase')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','02','1. Please send two copies of invoice.','2. Please notify us immediately if ','you are unable to ship as specified.','3. Enter this order in accordance, with the price,terms, ','delivery method and specification Listed above.','4. All disputes are subject to Our Jurisdiction Only','','SEPL/PR/','/20-21','Purchase Return')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','03','1. All disputes are subject to Our Jurisdiction Only','2. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection.','3. Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','','','SEPL/SR/','/20-21','Sale Return')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','04','1. Interest @24% p.a. will be charged if','the payment is not made in time.','2. All disputes are subject to Our Jurisdiction Only','3. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection','4. . Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','SEPL/Sale/','/20-21','Sale')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','05','1. Please send two copies of invoice.','2. Please notify us immediately if ','you are unable to ship as specified.','3. Enter this order in accordance, with the price,terms, ','delivery method and specification Listed above.','4. All disputes are subject to Our Jurisdiction Only','','SEPL/PC/','/20-21','Purchase Challan IN')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','06','','','','','','','','SEPL/PRC/','/20-21','Purchase Challan Out')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','07','1. Interest @24% p.a. will be charged if','the payment is not made in time.','2. All disputes are subject to Our Jurisdiction Only','3. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection','4. . Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','SEPL/SRC/','/20-21','Sale Challan IN')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','08','1. Interest @24% p.a. will be charged if','the payment is not made in time.','2. All disputes are subject to Our Jurisdiction Only','3. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection','4. . Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','SEPL/SC/','/20-21','Sale Challan Out')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','09','1. Interest @24% p.a. will be charged if','the payment is not made in time.','2. All disputes are subject to Our Jurisdiction Only','3. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection','4. . Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','SEPL/SJ/','/20-21','Sale Jobwork')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','10','1. Interest @24% p.a. will be charged if','the payment is not made in time.','2. All disputes are subject to Our Jurisdiction Only','3. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection','4. . Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','SEPL/SC/','/20-21','Sale Jobwork Unit Cost')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','11','1. Interest @24% p.a. will be charged if','the payment is not made in time.','2. All disputes are subject to Our Jurisdiction Only','3. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection','4. . Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','SEPL/DN/','/20-21','Challan Revesal IN')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','12','1. Interest @24% p.a. will be charged if','the payment is not made in time.','2. All disputes are subject to Delhi Jurisdiction Only','3. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection','4. . Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','SFAPL/PU/','/20-21','Challan Revesal Out')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','13','1. Interest @24% p.a. will be charged if','the payment is not made in time.','2. All disputes are subject to Our Jurisdiction Only','3. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection','4. . Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','SEPL/SC/','/20-21','Challan TO Be Billed IN')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','14','1. Interest @24% p.a. will be charged if','the payment is not made in time.','2. All disputes are subject to Our Jurisdiction Only','3. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection','4. . Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','SEPL/SC/','/20-21','Challan TO Be Billed OUT')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','15','1. Interest @24% p.a. will be charged if','the payment is not made in time.','2. All disputes are subject to Our Jurisdiction Only','3. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection','4. . Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','SEPL/SC/','/20-21','Challan Not TO Be Billed IN')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','16','1. Interest @24% p.a. will be charged if','the payment is not made in time.','2. All disputes are subject to Our Jurisdiction Only','3. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection','4. . Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','SEPL/SC/','/20-21','Challan Not TO Be Billed IOUT')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','17','1. The Deliverables shall be delivered or performed on the ','date and at the place specified in the Purchase Order.','2. Prices shall be as specified in the  Purchase  Order.','3. No increase in price shall be made or accepted unless ',' agreed in writing by Accenture.','4. The  Deliverables must conform in all respects with the','   Specifications and must be of sound.','SEPL/PO/','/20-21','Purchase Order')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','18','1. Interest @24% p.a. will be charged if','the payment is not made in time.','2. All disputes are subject to Our Jurisdiction Only','3. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection','4. . Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','SEPL/SO/','/20-21','Sale Order')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','19','1. Interest @24% p.a. will be charged if','the payment is not made in time.','2. All disputes are subject to Our Jurisdiction Only','3. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection','4. . Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','SEPL/ST/','/20-21','Stock Tranfer')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','20','','','','','','','','SEPL/RN/','/20-21','Stock Genral')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','21','1. Interest @24% p.a. will be charged if','the payment is not made in time.','2. All disputes are subject to Delhi Jurisdiction Only','3. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection','4. . Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','SFAPL/SU/','/20-21','Promotional Sale Challan Out')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','22','1. Interest @24% p.a. will be charged if','the payment is not made in time.','2. All disputes are subject to Our Jurisdiction Only','3. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection','4. . Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','SEPL/SQ/','/20-21','--')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','23','1. The price set for in Supplier’s Quotation (“Price”) are',' in  INDIA INR.','2. All Taxes shall be paid by Customer in addition to the ',' Price.','3.  Quotation (“Prices”) are valid for 30 days only.','','','SEPL/QP/','/20-21','Purchase Quotation')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','24','1. The price set for in Supplier’s Quotation (“Price”) are',' in  INDIA INR.','2. All Taxes shall be paid by Customer in addition to the ',' Price.','3.  Quotation (“Prices”) are valid for 30 days only.','','','SEPL/QS/','/20-21','Sales Quotation')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','25','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','26','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','27','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','28','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','29','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','30','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','31','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','32','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','33','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','34','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','35','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','36','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','37','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','38','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','39','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','40','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','41','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','42','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','43','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','44','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','45','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','46','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','47','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','48','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','49','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','50','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','51','','','','','','','','SEPL/PI/','/20-21','Payment')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','52','','','','','','','','SEPL/PR/','/20-21','Receipt')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','53','','','','','','','','SEPL/JE/','/20-21','Journal')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','54','','','','','','','','SEPL/CE/','/20-21','Contra')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','55','','','','','','','','SEPL/DN/','/20-21','Debit Note')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','56','','','','','','','','SEPL/CN/','/20-21','Credit Note')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','57','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','58','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','59','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','60','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','61','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','62','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','63','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','64','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','65','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','66','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','67','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','68','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','69','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','70','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','71','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','72','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','73','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','74','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','75','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','76','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','77','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','78','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','79','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','80','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','81','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','82','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','83','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','84','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','85','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','86','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','87','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','88','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','89','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','90','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','91','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','92','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','93','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','94','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','95','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','96','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','97','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','98','','','','','','','','','','')"
'        cnDatabase.Execute "Insert Into CompChild VALUES ('000001','99','','','','','','','','','','')"
'
'    cnDatabase.CommitTrans
'    'CloseMainConnection
'    Exit Function
'ErrorHandler:
'    UpdateMaster = False
'    cnDatabase.RollbackTrans
'    'CloseMainConnection
'End Function
Private Sub Command3_Click()
'    If UpdateComp(CompCode, False, False, True) Then
'    If CompCode = "" Then Call MsgBox("Please Login Company !!!", vbInformation, App.Title): Exit Sub
'        Call MsgBox("Successfully Updated Version !!!", vbInformation, App.Title)
'    Else
'        DisplayError ("Failed to Update Version")
'    End If
End Sub
'Private Function Update(ByVal CompanyCode As String, ByVal WithMasters As Boolean) As Boolean
'    'On Error GoTo ErrorHandler
'    On Error Resume Next
'    If CompCode = "" Then Call MsgBox("Please Login Company !!!", vbInformation, App.Title): Exit Function
'    Update = True
'    cnDatabase.CursorLocation = adUseClient
'    If cnDatabase.State = adStateOpen Then cnDatabase.Close
'    If DatabaseType = "MS SQL" Then
'    ConnectionString = "Provider=SQLOLEDB;Password=" & ServerPassword & ";Persist Security Info=True;User ID=" & ServerUser & ";Initial Catalog=EP" & CompCode & ";Data Source=" & ServerName
'    cnDatabase.Open ConnectionString
'    ElseIf DatabaseType = "MS Access" Then
'    cnDatabase.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DatabasePath & "\EasyPublish." & CompanyCode & ";Persist Security Info=False;Jet OLEDB:Database Password=pubprint123!@#"
'    End If
'
'    If rstCompanyMaster.State = adStateOpen Then rstCompanyMaster.Close
'    rstCompanyMaster.CursorLocation = adUseServer
'    BusySystemIndicator False
'    rstCompanyMaster.Open "SELECT * FROM CompanyMaster", cnDatabase, adOpenKeyset, adLockPessimistic
'    cnDatabase.BeginTrans
'    Dim Alias As String, compName As String
'    Dim aFlag As Boolean
'    Dim j As Long, K As Long
'        j = 0: K = 0: aFlag = True: Alias = ""
'        compName = rstCompanyMaster.Fields("Name")
'        K = Len(compName)
'    For j = 1 To K
'        If Mid(compName, j, 1) <> " " And aFlag = True Then
'            Alias = Alias + Mid(compName, j, 1)
'            aFlag = False
'        ElseIf Mid(compName, j, 1) = " " Then
'            aFlag = True
'        End If
'    Next j
''Account Master Update Table
'    cnDatabase.Execute "IF COL_LENGTH('AccountMaster', 'Notes') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE AccountMaster ADD Notes text NULL ALTER TABLE AccountMaster SET (LOCK_ESCALATION = TABLE) "
'    cnDatabase.Execute "IF COL_LENGTH('AccountMaster', 'Notes') IS NULL PRINT 'NOT Exists' ELSE Update AccountMaster Set Notes='' Where Notes IS NULL"
'    cnDatabase.Execute "IF COL_LENGTH('AccountMaster', 'Opening') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE AccountMaster ADD Opening decimal(12, 2) NULL ALTER TABLE AccountMaster SET (LOCK_ESCALATION = TABLE) "
'    cnDatabase.Execute "IF COL_LENGTH('AccountMaster', 'Opening') IS NULL PRINT 'NOT Exists' ELSE Update AccountMaster Set Opening=0 Where Opening IS NULL"
''BindingTypeChild Create Table
'    cnDatabase.Execute "IF COL_LENGTH('BindingTypeChild', 'Code') IS NOT NULL PRINT 'Exists' ELSE  CREATE TABLE BindingTypeChild(Code nvarchar(6) NOT NULL,BinderyProcess nvarchar(6) NOT NULL  ) ON [PRIMARY] ALTER TABLE dbo.BindingTypeChild SET (LOCK_ESCALATION = TABLE)"
''BookChild Create Table
'    cnDatabase.Execute "IF COL_LENGTH('BookChild', 'MaterialCentre') IS NOT NULL PRINT 'Exists' ELSE  CREATE TABLE BookChild(MaterialCentre nvarchar(6) NOT NULL,Item nvarchar(6) NOT NULL,OpBal int NOT NULL,FYCode nvarchar(6) NOT NULL ) ON [PRIMARY] ALTER TABLE BookChild SET (LOCK_ESCALATION = TABLE)"
''Book Master Update Table
'    cnDatabase.Execute "IF COL_LENGTH('BookMaster', 'Notes') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE BookMaster ADD Notes text NULL ALTER TABLE BookMaster SET (LOCK_ESCALATION = TABLE) "
'    cnDatabase.Execute "IF COL_LENGTH('BookMaster', 'Notes') IS NULL PRINT 'NOT Exists' ELSE Update BookMaster Set Notes='' Where Notes IS NULL"
''BookPOChild05
'    cnDatabase.Execute "IF COL_LENGTH('BookPOChild05', 'DeliveredQuantityC') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE BookPOChild05 ADD DeliveredQuantityC DECIMAL(12,0) NOT NULL DEFAULT (0) WITH VALUES "
'    cnDatabase.Execute "IF COL_LENGTH('BookPOChild05', 'DeliveredQuantityC') IS NULL PRINT 'NOT Exists' ELSE UPDATE BookPOChild05 SET DeliveredQuantityC=P.QuantityIssuedC+P.QuantityReceivedC FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE P.QuantityIssuedC+P.QuantityReceivedC>0"
'    cnDatabase.Execute "IF COL_LENGTH('BookPOChild05', 'DeliveredQuantityB') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE BookPOChild05 ADD DeliveredQuantityB DECIMAL(12,0) NOT NULL DEFAULT (0) WITH VALUES "
'    cnDatabase.Execute "IF COL_LENGTH('BookPOChild05', 'DeliveredQuantityB') IS NULL PRINT 'NOT Exists' ELSE UPDATE BookPOChild05 SET DeliveredQuantityB=P.QuantityIssuedB+P.QuantityReceivedB FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE P.QuantityIssuedB+P.QuantityReceivedB>0"
'    cnDatabase.Execute "IF COL_LENGTH('BookPOChild05', 'BilledMFC') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE BookPOChild05 ADD BilledMFC DECIMAL(12,0) NOT NULL DEFAULT (0) WITH VALUES "
'    cnDatabase.Execute "IF COL_LENGTH('BookPOChild05', 'BilledMFC') IS NULL PRINT 'NOT Exists' ELSE UPDATE BookPOChild05 SET BilledMFC=P.BilledAllC FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE P.BilledAllC>0"
'    cnDatabase.Execute "IF COL_LENGTH('BookPOChild05', 'BilledMFB') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE BookPOChild05 ADD BilledMFB DECIMAL(12,0) NOT NULL DEFAULT (0) WITH VALUES "
'    cnDatabase.Execute "IF COL_LENGTH('BookPOChild05', 'BilledMFB') IS NULL PRINT 'NOT Exists' ELSE UPDATE BookPOChild05 SET BilledMFB=P.BilledAllB FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE P.BilledAllB>0"
'    cnDatabase.Execute "DECLARE @sql NVARCHAR(255), @table NVARCHAR(50) " & _
'                                       "SET @table='BookPOChild05' " & _
'                                       "WHILE EXISTS (SELECT Name FROM SYS.DEFAULT_CONSTRAINTS P WHERE PARENT_OBJECT_ID=OBJECT_ID(@table) AND PARENT_COLUMN_ID IN ((SELECT column_id FROM sys.columns WHERE NAME IN ( 'QuantityIssuedC','QuantityReceivedC','QuantityIssuedB','QuantityReceivedB') AND object_id = P.PARENT_OBJECT_ID))) " & _
'                                       "BEGIN " & _
'                                       "SELECT @sql = 'ALTER TABLE '+@table+' DROP CONSTRAINT ' + (SELECT TOP 1 Name FROM SYS.DEFAULT_CONSTRAINTS P WHERE PARENT_OBJECT_ID=OBJECT_ID(@table) AND PARENT_COLUMN_ID IN ((SELECT column_id FROM sys.columns WHERE NAME IN ( 'QuantityIssuedC','QuantityReceivedC','QuantityIssuedB','QuantityReceivedB') AND object_id = P.PARENT_OBJECT_ID))) " & _
'                                       "EXEC sp_executesql @sql " & _
'                                       "End"
'    cnDatabase.Execute "IF COL_LENGTH('BookPOChild05', 'QuantityIssuedC') IS NOT NULL ALTER TABLE BookPOChild05 DROP COLUMN QuantityIssuedC,QuantityReceivedC,QuantityIssuedB,QuantityReceivedB ELSE PRINT 'Exists' "
''BookPOChild06 Table Update
'    cnDatabase.Execute "IF COL_LENGTH('BookPOChild06', 'DeliveredQuantityC') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE BookPOChild06 ADD DeliveredQuantityC DECIMAL(12,0) NOT NULL DEFAULT (0) WITH VALUES "
'    cnDatabase.Execute "IF COL_LENGTH('BookPOChild06', 'DeliveredQuantityC') IS NULL PRINT 'NOT Exists' ELSE UPDATE BookPOChild06 SET DeliveredQuantityC=P.QuantityIssuedC+P.QuantityReceivedC FROM BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code WHERE P.QuantityIssuedC+P.QuantityReceivedC>0"
'    cnDatabase.Execute "IF COL_LENGTH('BookPOChild06', 'DeliveredQuantityB') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE BookPOChild06 ADD DeliveredQuantityB DECIMAL(12,0) NOT NULL DEFAULT (0) WITH VALUES "
'    cnDatabase.Execute "IF COL_LENGTH('BookPOChild06', 'DeliveredQuantityB') IS NULL PRINT 'NOT Exists' ELSE UPDATE BookPOChild06 SET DeliveredQuantityB=P.QuantityIssuedB+P.QuantityReceivedB FROM BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code WHERE P.QuantityIssuedB+P.QuantityReceivedB>0"
'    cnDatabase.Execute "IF COL_LENGTH('BookPOChild06', 'BilledMEC') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE BookPOChild06 ADD BilledMEC DECIMAL(12,0) NOT NULL DEFAULT (0) WITH VALUES "
'    cnDatabase.Execute "IF COL_LENGTH('BookPOChild06', 'BilledMEC') IS NULL PRINT 'NOT Exists' ELSE UPDATE BookPOChild06 SET BilledMEC=P.BilledAllC FROM BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code WHERE P.BilledAllC>0"
'    cnDatabase.Execute "IF COL_LENGTH('BookPOChild06', 'BilledMEB') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE BookPOChild06 ADD BilledMEB DECIMAL(12,0) NOT NULL DEFAULT (0) WITH VALUES "
'    cnDatabase.Execute "IF COL_LENGTH('BookPOChild06', 'BilledMEB') IS NULL PRINT 'NOT Exists' ELSE UPDATE BookPOChild06 SET BilledMEB=P.BilledAllB FROM BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code WHERE P.BilledAllB>0"
'    cnDatabase.Execute "DECLARE @sql NVARCHAR(255), @table NVARCHAR(50) " & _
'                                       "SET @table='BookPOChild06' " & _
'                                       "WHILE EXISTS (SELECT Name FROM SYS.DEFAULT_CONSTRAINTS P WHERE PARENT_OBJECT_ID=OBJECT_ID(@table) AND PARENT_COLUMN_ID IN ((SELECT column_id FROM sys.columns WHERE NAME IN ( 'QuantityIssuedC','QuantityReceivedC','QuantityIssuedB','QuantityReceivedB') AND object_id = P.PARENT_OBJECT_ID))) " & _
'                                       "BEGIN " & _
'                                        "SELECT @sql = 'ALTER TABLE '+@table+' DROP CONSTRAINT ' + (SELECT TOP 1 Name FROM SYS.DEFAULT_CONSTRAINTS P WHERE PARENT_OBJECT_ID=OBJECT_ID(@table) AND PARENT_COLUMN_ID IN ((SELECT column_id FROM sys.columns WHERE NAME IN ( 'QuantityIssuedC','QuantityReceivedC','QuantityIssuedB','QuantityReceivedB') AND object_id = P.PARENT_OBJECT_ID))) " & _
'                                        "EXEC sp_executesql @sql " & _
'                                        "End"
'    cnDatabase.Execute "IF COL_LENGTH('BookPOChild06', 'QuantityIssuedC') IS NOT NULL ALTER TABLE BookPOChild06 DROP COLUMN QuantityIssuedC,QuantityReceivedC,QuantityIssuedB,QuantityReceivedB ELSE  PRINT 'Exists' "
''BookPOChild07 Table Update
'    cnDatabase.Execute "Alter Table BookPOChild07 Alter Column Number decimal(7, 3);"
'    cnDatabase.Execute "Alter Table BookPOChild07 Alter Column Rate decimal(12, 3);"
'    cnDatabase.Execute "IF COL_LENGTH('BookPOChild07', 'DeliveredQuantityC') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE BookPOChild07 ADD DeliveredQuantityC DECIMAL(12,0) NOT NULL DEFAULT (0) WITH VALUES "
'    cnDatabase.Execute "IF COL_LENGTH('BookPOChild07', 'DeliveredQuantityC') IS NULL PRINT 'NOT Exists' ELSE UPDATE BookPOChild07 SET DeliveredQuantityC=P.QuantityIssuedC+P.QuantityReceivedC FROM BookPOParent P INNER JOIN BookPOChild07 C ON P.Code=C.Code WHERE P.QuantityIssuedC+P.QuantityReceivedC>0"
'    cnDatabase.Execute "IF COL_LENGTH('BookPOChild07', 'DeliveredQuantityB') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE BookPOChild07 ADD DeliveredQuantityB DECIMAL(12,0) NOT NULL DEFAULT (0) WITH VALUES "
'    cnDatabase.Execute "IF COL_LENGTH('BookPOChild07', 'DeliveredQuantityB') IS NULL PRINT 'NOT Exists' ELSE UPDATE BookPOChild07 SET DeliveredQuantityB=P.QuantityIssuedB+P.QuantityReceivedB FROM BookPOParent P INNER JOIN BookPOChild07 C ON P.Code=C.Code WHERE P.QuantityIssuedB+P.QuantityReceivedB>0"
'    cnDatabase.Execute "IF COL_LENGTH('BookPOChild07', 'BilledMOC') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE BookPOChild07 ADD BilledMOC DECIMAL(12,0) NOT NULL DEFAULT (0) WITH VALUES "
'    cnDatabase.Execute "IF COL_LENGTH('BookPOChild07', 'BilledMOC') IS NULL PRINT 'NOT Exists' ELSE UPDATE BookPOChild07 SET BilledMOC=P.BilledAllC FROM BookPOParent P INNER JOIN BookPOChild07 C ON P.Code=C.Code WHERE P.BilledAllC>0"
'    cnDatabase.Execute "IF COL_LENGTH('BookPOChild07', 'BilledMOB') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE BookPOChild07 ADD BilledMOB DECIMAL(12,0) NOT NULL DEFAULT (0) WITH VALUES "
'    cnDatabase.Execute "IF COL_LENGTH('BookPOChild07', 'BilledMOB') IS NULL PRINT 'NOT Exists' ELSE UPDATE BookPOChild07 SET BilledMOB=P.BilledAllB FROM BookPOParent P INNER JOIN BookPOChild07 C ON P.Code=C.Code WHERE P.BilledAllB>0"
'    cnDatabase.Execute "DECLARE @sql NVARCHAR(255), @table NVARCHAR(50) " & _
'                                      "SET @table='BookPOChild07' " & _
'                                      "WHILE EXISTS (SELECT Name FROM SYS.DEFAULT_CONSTRAINTS P WHERE PARENT_OBJECT_ID=OBJECT_ID(@table) AND PARENT_COLUMN_ID IN ((SELECT column_id FROM sys.columns WHERE NAME IN ( 'QuantityIssuedC','QuantityReceivedC','QuantityIssuedB','QuantityReceivedB') AND object_id = P.PARENT_OBJECT_ID))) " & _
'                                      "BEGIN " & _
'                                          "SELECT @sql = 'ALTER TABLE '+@table+' DROP CONSTRAINT ' + (SELECT TOP 1 Name FROM SYS.DEFAULT_CONSTRAINTS P WHERE PARENT_OBJECT_ID=OBJECT_ID(@table) AND PARENT_COLUMN_ID IN ((SELECT column_id FROM sys.columns WHERE NAME IN ( 'QuantityIssuedC','QuantityReceivedC','QuantityIssuedB','QuantityReceivedB') AND object_id = P.PARENT_OBJECT_ID))) " & _
'                                          "EXEC sp_executesql @sql " & _
'                                      "End"
'    cnDatabase.Execute "IF COL_LENGTH('BookPOChild07', 'QuantityIssuedC') IS NOT NULL ALTER TABLE BookPOChild07 DROP COLUMN QuantityIssuedC,QuantityReceivedC,QuantityIssuedB,QuantityReceivedB ELSE PRINT 'Not Exists'"
''BookPOChild08 Table Update
'    cnDatabase.Execute "IF COL_LENGTH('BookPOChild08', 'DeliveredQuantityC') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE BookPOChild08 ADD DeliveredQuantityC DECIMAL(12,0) NOT NULL DEFAULT (0) WITH VALUES "
'    cnDatabase.Execute "IF COL_LENGTH('BookPOChild08', 'DeliveredQuantityC') IS NULL PRINT 'NOT Exists' ELSE UPDATE BookPOChild08 SET DeliveredQuantityC=P.QuantityIssuedC+P.QuantityReceivedC FROM BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code WHERE P.QuantityIssuedC+P.QuantityReceivedC>0"
'    cnDatabase.Execute "IF COL_LENGTH('BookPOChild08', 'DeliveredQuantityB') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE BookPOChild08 ADD DeliveredQuantityB DECIMAL(12,0) NOT NULL DEFAULT (0) WITH VALUES "
'    cnDatabase.Execute "IF COL_LENGTH('BookPOChild08', 'DeliveredQuantityB') IS NULL PRINT 'NOT Exists' ELSE UPDATE BookPOChild08 SET DeliveredQuantityB=P.QuantityIssuedB+P.QuantityReceivedB FROM BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code WHERE P.QuantityIssuedB+P.QuantityReceivedB>0"
'    cnDatabase.Execute "IF COL_LENGTH('BookPOChild08', 'BilledBNC') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE BookPOChild08 ADD BilledBNC DECIMAL(12,0) NOT NULL DEFAULT (0) WITH VALUES "
'    cnDatabase.Execute "IF COL_LENGTH('BookPOChild08', 'BilledBNC') IS NULL PRINT 'NOT Exists' ELSE UPDATE BookPOChild08 SET BilledBNC=P.BilledAllC FROM BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code WHERE P.BilledAllC>0"
'    cnDatabase.Execute "IF COL_LENGTH('BookPOChild08', 'BilledBNB') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE BookPOChild08 ADD BilledBNB DECIMAL(12,0) NOT NULL DEFAULT (0) WITH VALUES "
'    cnDatabase.Execute "IF COL_LENGTH('BookPOChild08', 'BilledBNB') IS NULL PRINT 'NOT Exists' UPDATE BookPOChild08 SET BilledBNB=P.BilledAllB FROM BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code WHERE P.BilledAllB>0"
'    cnDatabase.Execute "DECLARE @sql NVARCHAR(255), @table NVARCHAR(50) " & _
'                                      "SET @table='BookPOChild08' " & _
'                                      "WHILE EXISTS (SELECT Name FROM SYS.DEFAULT_CONSTRAINTS P WHERE PARENT_OBJECT_ID=OBJECT_ID(@table) AND PARENT_COLUMN_ID IN ((SELECT column_id FROM sys.columns WHERE NAME IN ( 'QuantityIssuedC','QuantityReceivedC','QuantityIssuedB','QuantityReceivedB') AND object_id = P.PARENT_OBJECT_ID))) " & _
'                                      "BEGIN " & _
'                                         "SELECT @sql = 'ALTER TABLE '+@table+' DROP CONSTRAINT ' + (SELECT TOP 1 Name FROM SYS.DEFAULT_CONSTRAINTS P WHERE PARENT_OBJECT_ID=OBJECT_ID(@table) AND PARENT_COLUMN_ID IN ((SELECT column_id FROM sys.columns WHERE NAME IN ( 'QuantityIssuedC','QuantityReceivedC','QuantityIssuedB','QuantityReceivedB') AND object_id = P.PARENT_OBJECT_ID))) " & _
'                                         "EXEC sp_executesql @sql " & _
'                                     "End "
'    cnDatabase.Execute "IF COL_LENGTH('BookPOChild08', 'QuantityIssuedC') IS NOT NULL ALTER TABLE BookPOChild08 DROP COLUMN QuantityIssuedC,QuantityReceivedC,QuantityIssuedB,QuantityReceivedB ELSE PRINT 'NotExists' "
''BookPOChild0801 Table Update
'    cnDatabase.Execute "IF COL_LENGTH('BookPOChild0801', 'DeliveredQuantityC') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE BookPOChild0801 ADD DeliveredQuantityC DECIMAL(12,0) NOT NULL DEFAULT (0) WITH VALUES "
'    cnDatabase.Execute "IF COL_LENGTH('BookPOChild0801', 'DeliveredQuantityC') IS NULL PRINT 'NOT Exists' ELSE UPDATE BookPOChild0801 SET DeliveredQuantityC=P.QuantityIssuedC+P.QuantityReceivedC FROM BookPOParent P INNER JOIN BookPOChild0801 C ON P.Code=C.Code WHERE P.QuantityIssuedC+P.QuantityReceivedC>0"
'    cnDatabase.Execute "IF COL_LENGTH('BookPOChild0801', 'DeliveredQuantityB') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE BookPOChild0801 ADD DeliveredQuantityB DECIMAL(12,0) NOT NULL DEFAULT (0) WITH VALUES "
'    cnDatabase.Execute "IF COL_LENGTH('BookPOChild0801', 'DeliveredQuantityB') IS NULL PRINT 'NOT Exists' ELSE UPDATE BookPOChild0801 SET DeliveredQuantityB=P.QuantityIssuedB+P.QuantityReceivedB FROM BookPOParent P INNER JOIN BookPOChild0801 C ON P.Code=C.Code WHERE P.QuantityIssuedB+P.QuantityReceivedB>0"
'    cnDatabase.Execute "IF COL_LENGTH('BookPOChild0801', 'BilledBMC') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE BookPOChild0801 ADD BilledBMC DECIMAL(12,0) NOT NULL DEFAULT (0) WITH VALUES "
'    cnDatabase.Execute "IF COL_LENGTH('BookPOChild0801', 'BilledBMC') IS NULL PRINT 'NOT Exists' ELSE UPDATE BookPOChild0801 SET BilledBMC=P.BilledAllC FROM BookPOParent P INNER JOIN BookPOChild0801 C ON P.Code=C.Code WHERE P.BilledAllC>0"
'    cnDatabase.Execute "IF COL_LENGTH('BookPOChild0801', 'BilledBMB') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE BookPOChild0801 ADD BilledBMB DECIMAL(12,0) NOT NULL DEFAULT (0) WITH VALUES "
'    cnDatabase.Execute "IF COL_LENGTH('BookPOChild0801', 'BilledBMB') IS NULL PRINT 'NOT Exists' ELSE UPDATE BookPOChild0801 SET BilledBMB=P.BilledAllB FROM BookPOParent P INNER JOIN BookPOChild0801 C ON P.Code=C.Code WHERE P.BilledAllB>0"
'    cnDatabase.Execute "DECLARE @sql NVARCHAR(255), @table NVARCHAR(50) " & _
'                                      "SET @table='BookPOChild0801' " & _
'                                      "WHILE EXISTS (SELECT Name FROM SYS.DEFAULT_CONSTRAINTS P WHERE PARENT_OBJECT_ID=OBJECT_ID(@table) AND PARENT_COLUMN_ID IN ((SELECT column_id FROM sys.columns WHERE NAME IN ( 'QuantityIssuedC','QuantityReceivedC','QuantityIssuedB','QuantityReceivedB') AND object_id = P.PARENT_OBJECT_ID))) " & _
'                                      "BEGIN " & _
'                                          "SELECT @sql = 'ALTER TABLE '+@table+' DROP CONSTRAINT ' + (SELECT TOP 1 Name FROM SYS.DEFAULT_CONSTRAINTS P WHERE PARENT_OBJECT_ID=OBJECT_ID(@table) AND PARENT_COLUMN_ID IN ((SELECT column_id FROM sys.columns WHERE NAME IN ( 'QuantityIssuedC','QuantityReceivedC','QuantityIssuedB','QuantityReceivedB') AND object_id = P.PARENT_OBJECT_ID))) " & _
'                                          "EXEC sp_executesql @sql " & _
'                                      "End "
'    cnDatabase.Execute "IF COL_LENGTH('BookPOChild0801', 'QuantityIssuedC') IS NOT NULL  ALTER TABLE BookPOChild0801 DROP COLUMN QuantityIssuedC,QuantityReceivedC,QuantityIssuedB,QuantityReceivedB ELSE PRINT 'Exists' "
''BookPOChild09
'    cnDatabase.Execute "DECLARE @sql NVARCHAR(255), @table NVARCHAR(50) " & _
'                                      "SET @table='BookPOChild09' " & _
'                                      "WHILE EXISTS (SELECT Name FROM SYS.DEFAULT_CONSTRAINTS P WHERE PARENT_OBJECT_ID=OBJECT_ID(@table) AND PARENT_COLUMN_ID IN ((SELECT column_id FROM sys.columns WHERE NAME IN ( 'QuantityIssuedC','QuantityReceivedC','QuantityIssuedB','QuantityReceivedB') AND object_id = P.PARENT_OBJECT_ID))) " & _
'                                      "BEGIN " & _
'                                          "SELECT @sql = 'ALTER TABLE '+@table+' DROP CONSTRAINT ' + (SELECT TOP 1 Name FROM SYS.DEFAULT_CONSTRAINTS P WHERE PARENT_OBJECT_ID=OBJECT_ID(@table) AND PARENT_COLUMN_ID IN ((SELECT column_id FROM sys.columns WHERE NAME IN ( 'QuantityIssuedC','QuantityReceivedC','QuantityIssuedB','QuantityReceivedB') AND object_id = P.PARENT_OBJECT_ID))) " & _
'                                          "EXEC sp_executesql @sql " & _
'                                      "End "
'    cnDatabase.Execute "IF COL_LENGTH('BookPOChild09', 'QuantityIssuedC') IS NOT NULL ALTER TABLE BookPOChild09 DROP COLUMN QuantityIssuedC,QuantityReceivedC,QuantityIssuedB,QuantityReceivedB ELSE PRINT 'Exists' "
''BookPOChild0901 Table Update
'    cnDatabase.Execute "IF COL_LENGTH('BookPOChild0901', 'DeliveredQuantityC') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE BookPOChild0901 ADD DeliveredQuantityC DECIMAL(12,0) NOT NULL DEFAULT (0) WITH VALUES "
'    cnDatabase.Execute "IF COL_LENGTH('BookPOChild0901', 'DeliveredQuantityC') IS NULL PRINT 'NOT Exists' ELSE UPDATE BookPOChild0901 SET DeliveredQuantityC=P.QuantityIssuedC+P.QuantityReceivedC FROM BookPOParent P INNER JOIN BookPOChild0901 C ON P.Code=C.Code WHERE P.QuantityIssuedC+P.QuantityReceivedC>0"
'    cnDatabase.Execute "IF COL_LENGTH('BookPOChild0901', 'DeliveredQuantityB') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE BookPOChild0901 ADD DeliveredQuantityB DECIMAL(12,0) NOT NULL DEFAULT (0) WITH VALUES "
'    cnDatabase.Execute "IF COL_LENGTH('BookPOChild0901', 'DeliveredQuantityB') IS NULL PRINT 'NOT Exists' ELSE UPDATE BookPOChild0901 SET DeliveredQuantityB=P.QuantityIssuedB+P.QuantityReceivedB FROM BookPOParent P INNER JOIN BookPOChild0901 C ON P.Code=C.Code WHERE P.QuantityIssuedB+P.QuantityReceivedB>0"
'    cnDatabase.Execute "IF COL_LENGTH('BookPOChild0901', 'BilledCFC') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE BookPOChild0901 ADD BilledCFC DECIMAL(12,0) NOT NULL DEFAULT (0) WITH VALUES "
'    cnDatabase.Execute "IF COL_LENGTH('BookPOChild0901', 'BilledCFC') IS NULL PRINT 'NOT Exists' ELSE UPDATE BookPOChild0901 SET BilledCFC=P.BilledAllC FROM BookPOParent P INNER JOIN BookPOChild0901 C ON P.Code=C.Code WHERE P.BilledAllC>0"
'    cnDatabase.Execute "IF COL_LENGTH('BookPOChild0901', 'BilledCFB') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE BookPOChild0901 ADD BilledCFB DECIMAL(12,0) NOT NULL DEFAULT (0) WITH VALUES "
'    cnDatabase.Execute "IF COL_LENGTH('BookPOChild0901', 'BilledCFB') IS NULL PRINT 'NOT Exists' ELSE UPDATE BookPOChild0901 SET BilledCFB=P.BilledAllB FROM BookPOParent P INNER JOIN BookPOChild0901 C ON P.Code=C.Code WHERE P.BilledAllB>0"
'    cnDatabase.Execute "DECLARE @sql NVARCHAR(255), @table NVARCHAR(50) " & _
'                                      "SET @table='BookPOChild0901' " & _
'                                      "WHILE EXISTS (SELECT Name FROM SYS.DEFAULT_CONSTRAINTS P WHERE PARENT_OBJECT_ID=OBJECT_ID(@table) AND PARENT_COLUMN_ID IN ((SELECT column_id FROM sys.columns WHERE NAME IN ( 'QuantityIssuedC','QuantityReceivedC','QuantityIssuedB','QuantityReceivedB') AND object_id = P.PARENT_OBJECT_ID))) " & _
'                                      "BEGIN " & _
'                                          "SELECT @sql = 'ALTER TABLE '+@table+' DROP CONSTRAINT ' + (SELECT TOP 1 Name FROM SYS.DEFAULT_CONSTRAINTS P WHERE PARENT_OBJECT_ID=OBJECT_ID(@table) AND PARENT_COLUMN_ID IN ((SELECT column_id FROM sys.columns WHERE NAME IN ( 'QuantityIssuedC','QuantityReceivedC','QuantityIssuedB','QuantityReceivedB') AND object_id = P.PARENT_OBJECT_ID))) " & _
'                                          "EXEC sp_executesql @sql " & _
'                                     "End"
'    cnDatabase.Execute "IF COL_LENGTH('BookPOChild0901', 'QuantityIssuedC') IS NOT NULL  ALTER TABLE BookPOChild0901 DROP COLUMN QuantityIssuedC,QuantityReceivedC,QuantityIssuedB,QuantityReceivedB ELSE PRINT 'Exists' "
''BookPOParent Table Update
'    cnDatabase.Execute "IF COL_LENGTH('BookPOParent', 'DeliveredQuantityC') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE BookPOParent ADD DeliveredQuantityC DECIMAL(12,0) NOT NULL DEFAULT (0) WITH VALUES "
'    cnDatabase.Execute "IF COL_LENGTH('BookPOParent', 'DeliveredQuantityC') IS NULL PRINT 'NOT Exists' ELSE UPDATE BookPOParent SET DeliveredQuantityC=QuantityIssuedC+QuantityReceivedC WHERE QuantityIssuedC+QuantityReceivedC>0"
'    cnDatabase.Execute "IF COL_LENGTH('BookPOParent', 'DeliveredQuantityB') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE BookPOParent ADD DeliveredQuantityB DECIMAL(12,0) NOT NULL DEFAULT (0) WITH VALUES "
'    cnDatabase.Execute "IF COL_LENGTH('BookPOParent', 'DeliveredQuantityB') IS NULL PRINT 'NOT Exists' ELSE UPDATE BookPOParent SET DeliveredQuantityB=QuantityIssuedB+QuantityReceivedB WHERE QuantityIssuedB+QuantityReceivedB>0"
'    cnDatabase.Execute "IF COL_LENGTH('BookPOParent', 'QuantityIssuedC') IS NOT NULL ALTER TABLE BookPOParent DROP CONSTRAINT df_QuantityIssuedC,df_QuantityReceivedC,df_QuantityIssuedB,df_QuantityReceivedB,df_QuantityIssued07C,df_QuantityReceived07C,df_QuantityIssued07B,df_QuantityReceived07B,df_QuantityIssued0801C,df_QuantityReceived0801C,df_QuantityIssued0801B,df_QuantityReceived0801B,df_BilledTextC,df_BilledTextB,df_BilledTitleC,df_BilledTitleB,df_BilledComboTitleC,df_BilledComboTitleB,df_BilledLaminationC,df_BilledLaminationB,df_BilledBOMC,df_BilledBOMB ELSE PRINT 'Exists'  "
'    cnDatabase.Execute "IF COL_LENGTH('BookPOParent', 'QuantityIssuedC') IS NOT NULL  ALTER TABLE BookPOParent DROP COLUMN QuantityIssuedC,QuantityReceivedC,QuantityIssuedB,QuantityReceivedB,QuantityIssued07C,QuantityReceived07C,QuantityIssued07B,QuantityReceived07B,QuantityIssued0801C,QuantityReceived0801C,QuantityIssued0801B,QuantityReceived0801B,BilledTextC,BilledTextB,BilledTitleC,BilledTitleB,BilledComboTitleC,BilledComboTitleB,BilledLaminationC,BilledLaminationB,BilledBOMC,BilledBOMB ELSE PRINT 'Exists' "
''Company Master Table Update Table
'    cnDatabase.Execute "IF COL_LENGTH('CompanyMaster', 'TallyIntegration') IS NOT NULL PRINT 'Exists' ELSE  Alter Table CompanyMaster Add TallyIntegration bit NOT NULL CONSTRAINT df_TallyIntegration DEFAULT '' "
'    cnDatabase.Execute "IF COL_LENGTH('CompanyMaster', 'TallyIntegration') IS NULL PRINT 'NOT Exists' ELSE Update CompanyMaster Set TallyIntegration=0"
'    cnDatabase.Execute "IF COL_LENGTH('CompanyMaster', 'BusyIntegration') IS NOT NULL PRINT 'Exists' ELSE  Alter Table CompanyMaster Add BusyIntegration bit NOT NULL CONSTRAINT df_BusyIntegration DEFAULT '' "
'    cnDatabase.Execute "IF COL_LENGTH('CompanyMaster', 'TallyIntegration') IS NULL PRINT 'NOT Exists' ELSE Update CompanyMaster Set BusyIntegration=0"
'    If FYCode = "'" Then FYCode = "'" + "00" + Right(Date, 4) + "'"
'    cnDatabase.Execute "IF COL_LENGTH('CompanyMaster', 'FYCode') IS NOT NULL PRINT 'Exists' ELSE Alter Table CompanyMaster Add FYCode nvarchar(6) NOT NULL CONSTRAINT df_FYCode DEFAULT ''  "
'    cnDatabase.Execute "IF COL_LENGTH('CompanyMaster', 'FYCode') IS NULL PRINT 'NOT Exists' ELSE Update CompanyMaster Set FYCode= " & FYCode & " Where FYCode ='' OR FYCode IS NULL"
'    'cnDatabase.Execute "Update CompanyMaster Set FYCode= " & FYCode & ""
'    cnDatabase.Execute "IF COL_LENGTH('CompanyMaster', 'Alias') IS NOT NULL PRINT 'Exists' ELSE Alter Table CompanyMaster Add Alias nvarchar(6) NOT NULL CONSTRAINT df_Alias DEFAULT '' "
'    cnDatabase.Execute "IF COL_LENGTH('CompanyMaster', 'Alias') IS NULL PRINT 'NOT Exists' ELSE Update CompanyMaster Set Alias= '" & Alias & "' Where Alias='' Or Alias IS Null"
''Comp Child Update Table
'    cnDatabase.Execute "IF COL_LENGTH('CompChild', 'VchName') IS NOT NULL PRINT 'Exists' ELSE  ALTER TABLE CompChild ADD VchName nvarchar(6) NOT NULL CONSTRAINT DF_CompChild_VchName DEFAULT ('')  "
'    cnDatabase.Execute "IF COL_LENGTH('CompChild', 'VchName') IS NULL PRINT 'NOT Exists' ELSE  Update CompChild Set VchName=''  Where VchName IS NULL"
''DebitCreditChild Create Table
'    cnDatabase.Execute "IF COL_LENGTH('DebitCreditChild', 'Code') IS NOT NULL PRINT 'Exists' ELSE CREATE TABLE DebitCreditChild (Code nvarchar(6) NOT NULL,TOA nchar(1) NOT NULL,Ref nvarchar(6) NULL,BOM nvarchar(6) NULL,Account nvarchar(6) NOT NULL,Debit decimal(12, 2) NOT NULL,Credit decimal(12, 2) NOT NULL,ShortNarration nvarchar(100) NOT NULL,SrNo tinyint NOT NULL,RefCode nvarchar(6) NULL)  ON [PRIMARY]ALTER TABLE DebitCreditChild SET (LOCK_ESCALATION = TABLE)"
''DebitCreditParent Create Table
'    cnDatabase.Execute "IF COL_LENGTH('DebitCreditParent', 'Code') IS NOT NULL PRINT 'Exists' ELSE CREATE TABLE DebitCreditParent(Code nvarchar(6) NOT NULL,Name nvarchar(25) NOT NULL,Date datetime NULL,Debit decimal(12, 2) NOT NULL,Credit decimal(12, 2) NOT NULL,LongNarration nvarchar(100) NULL,Type nvarchar(6) NOT NULL,CreatedBy nvarchar(6) NOT NULL,CreatedOn datetime NOT NULL,ModifiedBy nvarchar(6) NULL,ModifiedOn datetime NULL,RecordStatus nvarchar(1) NULL,VchSeries nvarchar(6) NULL,AutoVchNo nvarchar(10) NULL,FYCode nvarchar(6) NOT NULL,Notes text NULL)  ON [PRIMARY] TEXTIMAGE_ON [PRIMARY] ALTER TABLE DebitCreditParent ADD CONSTRAINT DF_DebitCreditParent_Debit DEFAULT ((0)) FOR Debit ALTER TABLE DebitCreditParent ADD CONSTRAINT DF_DebitCreditParent_Credit DEFAULT ((0)) FOR Credit ALTER TABLE DebitCreditParent ADD CONSTRAINT DF_DebitCreditParent_FYCode DEFAULT ('') FOR FYCode " & _
'                                      "ALTER TABLE DebitCreditParent ADD CONSTRAINT PK_DebitCreditParent PRIMARY KEY CLUSTERED (Code) WITH( STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY] ALTER TABLE DebitCreditParent SET (LOCK_ESCALATION = TABLE)"
''DebitCreditParent Update Table
'    cnDatabase.Execute "IF COL_LENGTH('DebitCreditParent', 'Notes') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE DebitCreditParent ADD Notes text NULL  ALTER TABLE DebitCreditParent SET (LOCK_ESCALATION = TABLE) "
'    cnDatabase.Execute "IF COL_LENGTH('DebitCreditParent', 'Notes') IS NULL PRINT 'NOT Exists' ELSE Update DebitCreditParent Set Notes='' Where Notes IS NULL"
''DebitCreditRef Create Table
'    cnDatabase.Execute "IF COL_LENGTH('DebitCreditRef', 'RefCode') IS NOT NULL PRINT 'Exists' ELSE CREATE TABLE DebitCreditRef(RefCode nvarchar(6) NOT NULL,Method tinyint NOT NULL,VchType nvarchar(6) NOT NULL,VchCode nvarchar(6) NOT NULL,VchNo nvarchar(25) NULL,VchDate datetime NOT NULL,Account nvarchar(6) NOT NULL,Debit decimal(12, 2) NOT NULL,Credit decimal(12, 2) NOT NULL, TOA nchar(1) NOT NULL)  ON [PRIMARY] ALTER TABLE DebitCreditRef SET (LOCK_ESCALATION = TABLE)"
''DebitCreditOthInf Create Table
'    cnDatabase.Execute "IF COL_LENGTH('DebitCreditOthInf', 'Code') IS NOT NULL PRINT 'Exists' ELSE CREATE TABLE DebitCreditOthInf(Code nvarchar(6) NOT NULL,BiltyNo nvarchar(30) NULL,BiltyDate datetime NULL,BiltyType nvarchar(30) NULL,Pkt smallint NOT NULL,Station nvarchar(30) NULL,Transport nvarchar(30) NULL,PktPicked bit NOT NULL)ON [PRIMARY]ALTER TABLE DebitCreditOthInf ADD CONSTRAINT  DF_DebitCreditOthInf_Pkt DEFAULT ((0)) FOR Pkt ALTER TABLE DebitCreditOthInf ADD CONSTRAINT DF_DebitCreditOthInf_PktPicked DEFAULT ((0)) FOR PktPicked ALTER TABLE DebitCreditOthInf SET (LOCK_ESCALATION = TABLE)"
''DiscountMaster Create Table
'    cnDatabase.Execute "IF COL_LENGTH('DiscountMaster', 'Party') IS NOT NULL PRINT 'Exists' ELSE CREATE TABLE DiscountMaster (Party nvarchar(6) NOT NULL,ItemGroup nvarchar(6) NOT NULL,[Disc%] decimal(6, 2) NOT NULL,FYCode nvarchar(6) Not NULL )  ON [PRIMARY] ALTER TABLE DiscountMaster ADD CONSTRAINT [DF_DiscountMaster_Disc%] DEFAULT ((0)) FOR [Disc%] ALTER TABLE DiscountMaster ADD CONSTRAINT PK_DiscountMaster PRIMARY KEY CLUSTERED (Party,ItemGroup) WITH( STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY] ALTER TABLE DiscountMaster SET (LOCK_ESCALATION = TABLE)"
''General Master Table Update Table
'    cnDatabase.Execute "IF COL_LENGTH('GeneralMaster', 'UnderGroup') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE GeneralMaster ADD  UnderGroup nvarchar(6) NULL   ALTER TABLE GeneralMaster SET (LOCK_ESCALATION = TABLE) "
''JobworkBVParent Update Table
'    cnDatabase.Execute "IF COL_LENGTH('JobworkBVParent', 'Notes') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE JobworkBVParent ADD Notes text NULL  ALTER TABLE JobworkBVParent SET (LOCK_ESCALATION = TABLE) "
'    cnDatabase.Execute "IF COL_LENGTH('JobworkBVParent', 'Notes') IS NULL PRINT 'NOT Exists' ELSE Update JobworkBVParent Set Notes='' Where Notes IS NULL"
'    cnDatabase.Execute "IF COL_LENGTH('JobworkBVParent', 'SalesType') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE JobworkBVParent ADD SalesType nvarchar(6) NULL  ALTER TABLE JobworkBVParent SET (LOCK_ESCALATION = TABLE) "
'    cnDatabase.Execute "IF COL_LENGTH('JobworkBVParent', 'SalesType') IS NULL PRINT 'NOT Exists' ELSE Update JobworkBVParent Set SalesType='*01052' Where SalesType IS NULL"
'    cnDatabase.Execute "IF COL_LENGTH('JobworkBVParent', 'GRDate') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE JobworkBVParent ADD GRDate nvarchar(40) NULL  ALTER TABLE JobworkBVParent SET (LOCK_ESCALATION = TABLE) "
'    cnDatabase.Execute "IF COL_LENGTH('JobworkBVParent', 'GRNo') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE JobworkBVParent ADD GRNo nvarchar(40) NULL  ALTER TABLE JobworkBVParent SET (LOCK_ESCALATION = TABLE) "
'    cnDatabase.Execute "IF COL_LENGTH('JobworkBVParent', 'Transport') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE JobworkBVParent ADD Transport nvarchar(40) NULL  ALTER TABLE JobworkBVParent SET (LOCK_ESCALATION = TABLE) "
'    cnDatabase.Execute "IF COL_LENGTH('JobworkBVParent', 'VehicleNo') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE JobworkBVParent ADD VehicleNo datetime NULL  ALTER TABLE JobworkBVParent SET (LOCK_ESCALATION = TABLE) "
'    cnDatabase.Execute "IF COL_LENGTH('JobworkBVParent', 'Station') IS NOT NULL PRINT 'Exists' ELSE ALTER TABLE JobworkBVParent ADD Station nvarchar(40) NULL  ALTER TABLE JobworkBVParent SET (LOCK_ESCALATION = TABLE) "
'    cnDatabase.Execute "IF COL_LENGTH('JobworkBVChild', 'BOM') IS NULL PRINT 'NOT Exists' ELSE ALTER TABLE JobworkBVChild ALTER COLUMN BOM NVARCHAR(18) NOT NULL"
'    cnDatabase.Execute "IF COL_LENGTH('JobworkBVChild', 'BOM') IS NULL PRINT 'NOT Exists' ELSE UPDATE JobworkBVChild SET BOM=LEFT(BOM,4)+'XXXXXXXXXXXX'+RIGHT(BOM,2) WHERE LEFT(BOM,2)='08' AND LEN(BOM)<18 "
'    cnDatabase.Execute "IF COL_LENGTH('JobworkBVChild', 'BOM') IS NULL PRINT 'NOT Exists' ELSE UPDATE JobworkBVChild SET BOM=LEFT(BOM,4)+'XXXXXXXXXXXX'+RIGHT(BOM,2) WHERE LEFT(BOM,2)='05' AND LEN(BOM)<18 "
'
'
'        'Default Master
'    'Genral Master
''Size Master_Type-1
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01001' OR Name='05.25X10.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01001','05.25X10.00','05.25X10.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01002' OR Name='10.00X29.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01002','10.00X29.00','10.00X29.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01003' OR Name='11.00X14.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01003','11.00X14.00','11.00X14.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01004' OR Name='11.50X18.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01004','11.50X18.00','11.50X18.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01005' OR Name='12.00X18.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01005','12.00X18.00','12.00X18.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01006' OR Name='12.00X23.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01006','12.00X23.00','12.00X23.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01007' OR Name='12.50X18.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01007','12.50X18.00','12.50X18.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01008' OR Name='13.00X19.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01008','13.00X19.00','13.00X19.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01009' OR Name='14.00X19.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01009','14.00X19.00','14.00X19.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01010' OR Name='14.00X22.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01010','14.00X22.00','14.00X22.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01011' OR Name='15.00X10.00 (CARD)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01011','15.00X10.00 (CARD)','15.00X10.00 (CARD)','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01012' OR Name='15.00X20.00 (CARD)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01012','15.00X20.00 (CARD)','15.00X20.00 (CARD)','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01013' OR Name='15.00X21.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01013','15.00X21.00','15.00X21.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01014' OR Name='15.00X27.50') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01014','15.00X27.50','15.00X27.50','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01015' OR Name='15.50X20.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01015','15.50X20.00','15.50X20.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01016' OR Name='15.50X20.50') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01016','15.50X20.50','15.50X20.50','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01017' OR Name='15.50X21.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01017','15.50X21.00','15.50X21.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01018' OR Name='15.50X21.50') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01018','15.50X21.50','15.50X21.50','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01019' OR Name='16.00X20.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01019','16.00X20.00','16.00X20.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01020' OR Name='16.00X20.50') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01020','16.00X20.50','16.00X20.50','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01021' OR Name='16.00X22.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01021','16.00X22.00','16.00X22.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01022' OR Name='16.00X24.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01022','16.00X24.00','16.00X24.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01023' OR Name='16.00X25.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01023','16.00X25.00','16.00X25.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01024' OR Name='16.00X30.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01024','16.00X30.00','16.00X30.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01025' OR Name='16.50X10.50') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01025','16.50X10.50','16.50X10.50','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01026' OR Name='17.00X22.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01026','17.00X22.00','17.00X22.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01027' OR Name='17.00X24.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01027','17.00X24.00','17.00X24.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01028' OR Name='18.00X23.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01028','18.00X23.00','18.00X23.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01029' OR Name='18.00X23.00 (Card)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01029','18.00X23.00 (Card)','18.00X23.00 (Card)','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01030' OR Name='18.00X24.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01030','18.00X24.00','18.00X24.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01031' OR Name='18.00X25.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01031','18.00X25.00','18.00X25.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01032' OR Name='19.00X20.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01032','19.00X20.00','19.00X20.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01033' OR Name='19.00X25.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01033','19.00X25.00','19.00X25.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01034' OR Name='19.00X38.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01034','19.00X38.00','19.00X38.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01035' OR Name='20.00X24.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01035','20.00X24.00','20.00X24.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01036' OR Name='20.00X25.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01036','20.00X25.00','20.00X25.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01037' OR Name='20.00X26.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01037','20.00X26.00','20.00X26.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01038' OR Name='20.00X28.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01038','20.00X28.00','20.00X28.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01039' OR Name='20.00X30.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01039','20.00X30.00','20.00X30.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01040' OR Name='20.00X30.00(A/P)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01040','20.00X30.00(A/P)','20.00X30.00(A/P)','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01041' OR Name='20.00X30.00(Card)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01041','20.00X30.00(Card)','20.00X30.00(Card)','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01042' OR Name='20.00X31.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01042','20.00X31.00','20.00X31.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01043' OR Name='20.50X24.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01043','20.50X24.00','20.50X24.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01044' OR Name='20.50X31.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01044','20.50X31.00','20.50X31.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01045' OR Name='21.00X29.70 (A4)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01045','21.00X29.70 (A4)','21.00X29.70 (A4)','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01046' OR Name='21.00X30.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01046','21.00X30.00','21.00X30.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01047' OR Name='21.00X30.00(CARD)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01047','21.00X30.00(CARD)','21.00X30.00(CARD)','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01048' OR Name='21.00X31.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01048','21.00X31.00','21.00X31.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01049' OR Name='21.00X32.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01049','21.00X32.00','21.00X32.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01050' OR Name='21.00X33.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01050','21.00X33.00','21.00X33.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01051' OR Name='21.00X34.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01051','21.00X34.00','21.00X34.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01052' OR Name='21.00X35.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01052','21.00X35.00','21.00X35.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01053' OR Name='21.50X28.50') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01053','21.50X28.50','21.50X28.50','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01054' OR Name='22.00X28.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01054','22.00X28.00','22.00X28.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01055' OR Name='22.00X32.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01055','22.00X32.00','22.00X32.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01056' OR Name='22.00X34.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01056','22.00X34.00','22.00X34.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01057' OR Name='23.00X30.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01057','23.00X30.00','23.00X30.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01058' OR Name='23.00X33.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01058','23.00X33.00','23.00X33.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01059' OR Name='23.00X35.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01059','23.00X35.00','23.00X35.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01060' OR Name='23.00X36.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01060','23.00X36.00','23.00X36.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01061' OR Name='23.00X36.00(A/P)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01061','23.00X36.00(A/P)','23.00X36.00(A/P)','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01062' OR Name='23.00X36.00(Card)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01062','23.00X36.00(Card)','23.00X36.00(Card)','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01063' OR Name='24.00X34.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01063','24.00X34.00','24.00X34.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01064' OR Name='24.00X36.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01064','24.00X36.00','24.00X36.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01065' OR Name='24.13X24.13') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01065','24.13X24.13','24.13X24.13','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01066' OR Name='25.00X30.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01066','25.00X30.00','25.00X30.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01067' OR Name='25.00X36.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01067','25.00X36.00','25.00X36.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01068' OR Name='25.00X38.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01068','25.00X38.00','25.00X38.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01069' OR Name='26.00X38.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01069','26.00X38.00','26.00X38.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01070' OR Name='26.00X40.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01070','26.00X40.00','26.00X40.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01071' OR Name='28.00X35.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01071','28.00X35.00','28.00X35.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01072' OR Name='28.00X40.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01072','28.00X40.00','28.00X40.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01073' OR Name='30.00X40.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01073','30.00X40.00','30.00X40.00','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01074' OR Name='31.50X41.50') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*01074','31.50X41.50','31.50X41.50','1','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
''Item Group Master_TYPE-5
'    If Trim(ReadFromFile("Client ID")) = "Publisher" Then
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*05001' OR Name='Activity Book') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*05001','Activity Book','Activity Book','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*05002' OR Name='Box') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*05002','Box','Box','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*05003' OR Name='CARD') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*05003','CARD','CARD','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*05004' OR Name='CATALOGUE') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*05004','CATALOGUE','CATALOGUE','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*05005' OR Name='General') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*05005','General','General','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*05006' OR Name='GRADE 1') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*05006','GRADE 1','GRADE 1','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*05007' OR Name='GRADE 2') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*05007','GRADE 2','GRADE 2','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*05008' OR Name='GRADE 3') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*05008','GRADE 3','GRADE 3','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*05009' OR Name='GRADE 4') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*05009','GRADE 4','GRADE 4','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*05010' OR Name='GRADE 5') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*05010','GRADE 5','GRADE 5','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*05011' OR Name='JUNIOR') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*05011','JUNIOR','JUNIOR','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*05012' OR Name='LEVEL 1') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*05012','LEVEL 1','LEVEL 1','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*05013' OR Name='LEVEL 2') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*05013','LEVEL 2','LEVEL 2','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*05014' OR Name='LEVEL 3') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*05014','LEVEL 3','LEVEL 3','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*05015' OR Name='LEVEL 4') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*05015','LEVEL 4','LEVEL 4','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*05016' OR Name='LEVEL 5') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*05016','LEVEL 5','LEVEL 5','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*05017' OR Name='LEVEL A') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*05017','LEVEL A','LEVEL A','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*05018' OR Name='LEVEL B') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*05018','LEVEL B','LEVEL B','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*05019' OR Name='LEVEL C') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*05019','LEVEL C','LEVEL C','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*05020' OR Name='NURSERY') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*05020','NURSERY','NURSERY','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*05021' OR Name='SECONDARY STD VI') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*05021','SECONDARY STD VI','SECONDARY STD','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*05022' OR Name='SENIOR') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*05022','SENIOR','SENIOR','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*05023' OR Name='SET 1') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*05023','SET 1','SET 1','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'End If
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*05024' OR Name='Item Group') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*05024','Item Group','Item Group','5','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
''Binding Type_Type-6
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*06001' OR Name='Die_Cutting') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*06001','Die_Cutting','Die_Cutting','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*06002' OR Name='Die_Perforation') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*06002','Die_Perforation','Die_Perforation','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*06003' OR Name='Hard Bound') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*06003','Hard Bound','Hard Bound','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*06004' OR Name='Perfect Binding With Sewing') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*06004','Perfect Binding With Sewing','Perfect Binding With Sewing','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*06005' OR Name='Perfect Binding With Sewing(CD-Insert)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*06005','Perfect Binding With Sewing(CD-Insert)','Perfect Binding With Sewing(CD-Insert)','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*06006' OR Name='Spiral Binding') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*06006','Spiral Binding','Spiral Binding','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*06007' OR Name='Wirro Binding') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*06007','Wirro Binding','Wirro Binding','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*06008' OR Name='Cutting & Packing') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*06008','Cutting & Packing','Cutting & Packing','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*06009' OR Name='Cutting Only') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*06009','Cutting Only','Cutting Only','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*06010' OR Name='Half Die Cut') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*06010','Half Die Cut','Half Die Cut','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*06011' OR Name='Loose') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*06011','Loose','Loose','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*06012' OR Name='Pad Gumming') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*06012','Pad Gumming','Pad Gumming','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*06013' OR Name='Pakki Binding') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*06013','Pakki Binding','Pakki Binding','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*06014' OR Name='Kachchi Binding') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*06014','Kachchi Binding','Kachchi Binding','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*06015' OR Name='Center Pinning (BOX)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*06015','Center Pinning (BOX)','Center Pinning (BOX)','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*06016' OR Name='Center Pin Binding') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*06016','Center Pin Binding','Center Pin Binding','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*06017' OR Name='None') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*06017','None','None','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*06018' OR Name='Perfect Binding') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*06018','Perfect Binding','Perfect Binding','6','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
''Finishing Type_Type-7
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07001' OR Name='BOPP Gloss') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07001','BOPP Gloss','BOPP Gloss','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07002' OR Name='BOPP Matt') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07002','BOPP Matt','BOPP Matt','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07003' OR Name='Box Packing') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07003','Box Packing','Box Packing','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07004' OR Name='Center Pin Binding') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07004','Center Pin Binding','Center Pin Binding','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07005' OR Name='Counting & Fabrication') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07005','Counting & Fabrication','Counting & Fabrication','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07006' OR Name='Creasing+Folding+Packing') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07006','Creasing+Folding+Packing','Creasing+Folding+Packing','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07007' OR Name='Cutting and Packing') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07007','Cutting and Packing','Cutting and Packing','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07008' OR Name='Cutting Leaflet Only') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07008','Cutting Leaflet Only','Cutting Leaflet Only','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07009' OR Name='Die Cutting Charges') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07009','Die Cutting Charges','Die Cutting Charges','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07010' OR Name='Die Making Charges') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07010','Die Making Charges','Die Making Charges','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07011' OR Name='Digital Print') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07011','Digital Print','Digital Print','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07012' OR Name='Embossing') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07012','Embossing','Embossing','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07013' OR Name='Foiling Charges') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07013','Foiling Charges','Foiling Charges','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07014' OR Name='Folding & Packing') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07014','Folding & Packing','Folding & Packing','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07015' OR Name='Graning') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07015','Graning','Graning','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07016' OR Name='Half Die Cutting Charges') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07016','Half Die Cutting Charges','Half Die Cutting Charges','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07017' OR Name='Hardbound Binding') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07017','Hardbound Binding','Hardbound Binding','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07018' OR Name='Hologram') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07018','Hologram','Hologram','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07019' OR Name='Matt + Spot UV') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07019','Matt + Spot UV','Matt + Spot UV','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07020' OR Name='Matt + Spot UV + Foiling + Embossing') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07020','Matt + Spot UV + Foiling + Embossing','Matt + Spot UV + Foiling + Embossing','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07021' OR Name='Matt + Spot UV+Glitter UV') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07021','Matt + Spot UV+Glitter UV','Matt + Spot UV+Glitter UV','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07022' OR Name='Matt Both Side') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07022','Matt Both Side','Matt Both Side','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07023' OR Name='MINI Offset JOB') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07023','MINI Offset JOB','MINI Offset JOB','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07024' OR Name='None') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07024','None','None','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07025' OR Name='Packing Shrink') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07025','Packing Shrink','Packing Shrink','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07026' OR Name='Paper Cost') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07026','Paper Cost','Paper Cost','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07027' OR Name='Pasting Charges') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07027','Pasting Charges','Pasting Charges','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07028' OR Name='Perfect Binding') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07028','Perfect Binding','Perfect Binding','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07029' OR Name='Plate') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07029','Plate','Plate','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07030' OR Name='Printing 4 Col') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07030','Printing 4 Col','Printing 4 Col','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07031' OR Name='PVC') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07031','PVC','PVC','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07032' OR Name='Spot UV') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07032','Spot UV','Spot UV','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07033' OR Name='Thermal Matt') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07033','Thermal Matt','Thermal Matt','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07034' OR Name='UV Hybraid') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07034','UV Hybraid','UV Hybraid','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*07035' OR Name='Varnising') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*07035','Varnising','Varnising','7','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
''Project Member/Editorial Team Master_Type-8
'If Trim(ReadFromFile("Client ID")) = "Publisher" Then
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*08002' OR Name='Author_ABC') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*08002','Author_ABC','Author_ABC','8','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*08003' OR Name='DTP_ABC') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*08003','DTP_ABC','DTP_ABC','8','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*08005' OR Name='Editor_ABC') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*08005','Editor_ABC','Editor_ABC','8','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*08007' OR Name='Graphic_ABC') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*08007','Graphic_ABC','Graphic_ABC','8','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*08008' OR Name='PPQ_ABC') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*08008','PPQ_ABC','PPQ_ABC','8','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*08009' OR Name='Processing_S.R.K') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*08009','Processing_S.R.K','Processing_S.R.K','8','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*08010' OR Name='Proof Reader_ABC') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*08010','Proof Reader_ABC','Proof Reader_ABC','8','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*08011' OR Name='Type Setting_ABC') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*08011','Type Setting_ABC','Type Setting_Sanjay Khanna','8','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'End If
''Plate Master_Type-9
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*09001' OR Name='CTP_Plates') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*09001','CTP_Plates','CTP_Plates','9','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*09002' OR Name='Nagative-Cut Pieces') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*09002','Nagative-Cut Pieces','Nagative-Cut Pieces','9','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*09003' OR Name='Nagative-One Pieces') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*09003','Nagative-One Pieces','Nagative-One Pieces','9','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
''Size Group Master-10
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10016' OR Name='Extra Large-28''''X40''''') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*10016','Extra Large-28''''X40''''','Extra Large-28''''X40''''','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10018' OR Name='Extra Large-28''''X40''''-(Card)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*10018','Extra Large-28''''X40''''-(Card)','Extra Large-28''''X40''''-(Card)','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10001' OR Name='Extra Large-28''''X40''''-A/P') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*10001','Extra Large-28''''X40''''-A/P','Extra Large-28''''X40''''-A/P','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10002' OR Name='Extra Large-28''''X40''''-A/P_SPL') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*10002','Extra Large-28''''X40''''-A/P_SPL','Extra Large-28''''X40''''-A/P_SPL','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10003' OR Name='Extra Large-30''''X40''''') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*10003','Extra Large-30''''X40''''','Extra Large-30''''X40''''','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10004' OR Name='Extra Large-30''''X40''''-(A/P)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*10004','Extra Large-30''''X40''''-(A/P)','Extra Large-30''''X40''''-(A/P)','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10005' OR Name='Extra Large-30''''X40''''-(Card)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*10005','Extra Large-30''''X40''''-(Card)','Extra Large-30''''X40''''-(Card)','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10006' OR Name='LARGE-23''''X36''''') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*10006','LARGE-23''''X36''''','LARGE-23''''X36''''','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10007' OR Name='LARGE-23''''X36''''-(A/P)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*10007','LARGE-23''''X36''''-(A/P)','LARGE-23''''X36''''-(A/P)','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10008' OR Name='LARGE-23''''X36''''-(Card)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*10008','LARGE-23''''X36''''-(Card)','LARGE-23''''X36''''-(Card)','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10019' OR Name='Little-11.50''''X18.00''''') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*10019','Little-11.50''''X18.00''''','Little-11.50''''X18.00''''','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10021' OR Name='Little-11.50''''X18.00''''-(A/P)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*10021','Little-11.50''''X18.00''''-(A/P)','Little-11.50''''X18.00''''-(A/P)','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10020' OR Name='Little-11.50''''X18.00''''-(Card)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*10020','Little-11.50''''X18.00''''-(Card)','Little-11.50''''X18.00''''-(Card)','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10009' OR Name='Medium-20''''X30''''') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*10009','Medium-20''''X30''''','Medium-20''''X30''''','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10010' OR Name='Medium-20''''X30''''(A/P)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*10010','Medium-20''''X30''''(A/P)','Medium-20''''X30''''(A/P)','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10011' OR Name='Medium-20''''X30''''(Card)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*10011','Medium-20''''X30''''(Card)','Medium-20''''X30''''(Card)','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10012' OR Name='Small-19''''X26''''') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*10012','Small-19''''X26''''','Small-19''''X26''''','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10017' OR Name='Small-19''''X26''''-(A/P)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*10017','Small-19''''X26''''-(A/P)','Small-19''''X26''''-(A/P)','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10013' OR Name='Small-19''''X26''''(Card)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*10013','Small-19''''X26''''(Card)','Small-19''''X26''''(Card)','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10014' OR Name='Web-508mm') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*10014','Web-508mm','Web-508mm','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10015' OR Name='Web-578mm') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*10015','Web-578mm','Web-578mm','10','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
''Finish Size Master_TYPE-11
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11001' OR Name='05.25x10.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11001','05.25x10.00','05.25x10.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11002' OR Name='12.00X18.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11002','12.00X18.00','12.00X18.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11003' OR Name='12.00X23.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11003','12.00X23.00','12.00X23.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11004' OR Name='14.00X19.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11004','14.00X19.00','14.00X19.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11005' OR Name='15.00X10.00 (CARD)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11005','15.00X10.00 (CARD)','15.00X10.00 (CARD)','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11006' OR Name='15.50X20.50') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11006','15.50X20.50','15.50X20.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11007' OR Name='16.00x20.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11007','16.00x20.00','16.00x20.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11008' OR Name='16.00X24.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11008','16.00X24.00','16.00X24.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11009' OR Name='16.50X10.50') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11009','16.50X10.50','16.50X10.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11010' OR Name='17.00X22.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11010','17.00X22.00','17.00X22.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11011' OR Name='04.00X06.87') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11011','04.00X06.87','04.00X06.87','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11012' OR Name='04.25X05.50') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11012','04.25X05.50','04.25X05.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11013' OR Name='04.25X07.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11013','04.25X07.00','04.25X07.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11014' OR Name='04.37X07.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11014','04.37X07.00','04.37X07.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11015' OR Name='04.72X07.48') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11015','04.72X07.48','04.72X07.48','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11016' OR Name='05.00X07.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11016','05.00X07.00','05.00X07.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11017' OR Name='05.00X08.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11017','05.00X08.00','05.00X08.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11018' OR Name='05.06X07.81') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11018','05.06X07.81','05.06X07.81','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11019' OR Name='05.25X08.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11019','05.25X08.00','05.25X08.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11020' OR Name='05.50X08.50') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11020','05.50X08.50','05.50X08.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11021' OR Name='05.83X08.27') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11021','05.83X08.27','05.83X08.27','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11022' OR Name='06.00X08.25') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11022','06.00X08.25','06.00X08.25','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11023' OR Name='06.00X08.50') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11023','06.00X08.50','06.00X08.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11024' OR Name='06.00X09.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11024','06.00X09.00','06.00X09.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11025' OR Name='06.14X09.21') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11025','06.14X09.21','06.14X09.21','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11026' OR Name='06.25X09.50') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11026','06.25X09.50','06.25X09.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11027' OR Name='06.63X10.25') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11027','06.63X10.25','06.63X10.25','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11028' OR Name='06.69X09.61') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11028','06.69X09.61','06.69X09.61','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11029' OR Name='06.75X09.50') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11029','06.75X09.50','06.75X09.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11030' OR Name='07.00X07.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11030','07.00X07.00','07.00X07.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11031' OR Name='07.00X09.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11031','07.00X09.00','07.00X09.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11032' OR Name='07.00X10.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11032','07.00X10.00','07.00X10.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11033' OR Name='07.25X09.50') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11033','07.25X09.50','07.25X09.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11034' OR Name='07.44X09.69') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11034','07.44X09.69','07.44X09.69','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11035' OR Name='07.50X07.50') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11035','07.50X07.50','07.50X07.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11036' OR Name='07.50X09.25') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11036','07.50X09.25','07.50X09.25','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11037' OR Name='07.50X09.50') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11037','07.50X09.50','07.50X09.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11038' OR Name='07.75X10.50') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11038','07.75X10.50','07.75X10.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11039' OR Name='08.00X08.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11039','08.00X08.00','08.00X08.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11040' OR Name='08.00X10.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11040','08.00X10.00','08.00X10.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11041' OR Name='08.00X10.88') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11041','08.00X10.88','08.00X10.88','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11042' OR Name='08.00X11.25') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11042','08.00X11.25','08.00X11.25','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11043' OR Name='08.25X08.25') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11043','08.25X08.25','08.25X08.25','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11044' OR Name='08.25X11.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11044','08.25X11.00','08.25X11.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11045' OR Name='08.27X11.69') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11045','08.27X11.69','08.27X11.69','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11046' OR Name='08.50X08.50') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11046','08.50X08.50','08.50X08.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11047' OR Name='08.50X09.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11047','08.50X09.00','08.50X09.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11048' OR Name='08.50X11.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11048','08.50X11.00','08.50X11.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11049' OR Name='09.00X07.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11049','09.00X07.00','09.00X07.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11050' OR Name='09.00X12.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11050','09.00X12.00','09.00X12.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11051' OR Name='10.00X10.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11051','10.00X10.00','10.00X10.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11052' OR Name='11.00X13.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11052','11.00X13.00','11.00X13.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11053' OR Name='11.00X17.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11053','11.00X17.00','11.00X17.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11054' OR Name='11.00X18.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11054','11.00X18.00','11.00X18.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11055' OR Name='12.00X12.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11055','12.00X12.00','12.00X12.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11056' OR Name='18.00X23.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11056','18.00X23.00','18.00X23.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11057' OR Name='07.75X11.25') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11057','07.75X11.25','07.75X11.25','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11058' OR Name='08.00X11.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11058','08.00X11.00','08.00X11.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11059' OR Name='04.50X01.75') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11059','04.50X01.75','04.50X01.75','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11060' OR Name='11.00x15.75') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11060','11.00x15.75','11.00x15.75','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11061' OR Name='11.00X16.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11061','11.00X16.00','11.00X16.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11062' OR Name='08.25X11.75') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11062','08.25X11.75','08.25X11.75','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11063' OR Name='04.00X06.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11063','04.00X06.00','04.00X06.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11064' OR Name='20.00X30.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11064','20.00X30.00','20.00X30.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11065' OR Name='17.50X22.50') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11065','17.50X22.50','17.50X22.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11066' OR Name='11.50X08.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11066','11.50X08.00','11.50X08.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11067' OR Name='21.00X31.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11067','21.00X31.00','21.00X31.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11068' OR Name='05.30X08.30') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11068','05.30X08.30','05.30X08.30','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11069' OR Name='11.50X10.75') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11069','11.50X10.75','11.50X10.75','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11070' OR Name='08.50X10.75') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11070','08.50X10.75','08.50X10.75','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11071' OR Name='02.00X03.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11071','02.00X03.00','02.00X03.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11072' OR Name='11.50X07.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11072','11.50X07.00','11.50X07.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11073' OR Name='05.50X19.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11073','05.50X19.00','05.50X19.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11074' OR Name='10.25X07.50') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11074','10.25X07.50','10.25X07.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11075' OR Name='07.50X13.75') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11075','07.50X13.75','07.50X13.75','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11076' OR Name='07.00X02.50') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11076','07.00X02.50','07.00X02.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11077' OR Name='06.50X09.50') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11077','06.50X09.50','06.50X09.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11078' OR Name='04.00x07.50') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11078','04.00x07.50','04.00x07.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11079' OR Name='23.00X36.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11079','23.00X36.00','23.00X36.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11080' OR Name='15.00X20.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11080','15.00X20.00','15.00X20.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11081' OR Name='25.00X36.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11081','25.00X36.00','25.00X36.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11082' OR Name='09.00X14.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11082','09.00X14.00','09.00X14.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11083' OR Name='05.25X07.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11083','05.25X07.00','05.25X07.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11084' OR Name='08.00X10.50') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11084','08.00X10.50','08.00X10.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11085' OR Name='07.50X08.50') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11085','07.50X08.50','07.50X08.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11086' OR Name='03.25X04.75') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11086','03.25X04.75','03.25X04.75','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11087' OR Name='09.75X11.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11087','09.75X11.00','09.75X11.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11088' OR Name='13.50X18.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11088','13.50X18.00','13.50X18.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11089' OR Name='07.62X11.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11089','07.62X11.00','07.62X11.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11090' OR Name='07.36X11.00') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11090','07.36X11.00','07.36X11.00','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11091' OR Name='08.26X11.69') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11091','08.26X11.69','08.26X11.69','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11092' OR Name='09.50X09.50') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11092','09.50X09.50','09.50X09.50','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11093' OR Name='11.69X05.20') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11093','11.69X05.20','11.69X05.20','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11094' OR Name='05.75X08.25') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11094','05.75X08.25','05.75X08.25','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*11095' OR Name='21.00X29.70') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*11095','21.00X29.70','21.00X29.70','11','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
''Genral Accounts Groups_TYPE-12
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*12002' OR Name='Account Group') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*12002','Account Group','Account Group','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26031')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*99996' OR Name='Transporter') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*99996','Transporter','Transporter','12','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*99997' OR Name='Packer') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*99997','Packer','Transporter','12','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*99998' OR Name='Deliverer') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*99998','Deliverer','Deliverer','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26030')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*99999' OR Name='Material Centre') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*99999','Material Centre','Material Centre','12','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*12001' OR Name='Binders') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*12001','Binders','Binders','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26030')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*12003' OR Name='Box Supplier') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*12003','Box Supplier','Box Supplier','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26030')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*12004' OR Name='CD Suppliers') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*12004','CD Suppliers','CD Suppliers','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26030')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*12005' OR Name='FG Godown') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*12005','FG Godown','FG Godown','12','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*12006' OR Name='Laminator') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*12006','Laminator','Laminator','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26030')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*12007' OR Name='Packaging Supplier') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*12007','Packaging Supplier','Packaging Supplier','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26030')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*12008' OR Name='Paper Suppliers') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*12008','Paper Suppliers','Paper Suppliers','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26030')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*12009' OR Name='Printer') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*12009','Printer','Printer','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26030')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*12010' OR Name='Printer & Binder') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*12010','Printer & Binder','Printer & Binder','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26030')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*12011' OR Name='Printer, Binder & Laminator') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*12011','Printer, Binder & Laminator','Printer, Binder & Laminator','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26030')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*12012' OR Name='Processor & Printer') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*12012','Processor & Printer','Processor & Printer','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26030')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*12013' OR Name='Processor, Printer & Laminator') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*12013','Processor, Printer & Laminator','Processor, Printer & Laminator','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26030')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*12014' OR Name='UFG Godown') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*12014','UFG Godown','UFG Godown','12','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*12015' OR Name='Publisher') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*12015','Publisher','Publisher','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26031')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*12016' OR Name='Clients') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*12016','Clients','Clients','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26031')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*12017' OR Name='Cons. Supplier') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*12017','Cons. Supplier','Cons. Supplier','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26030')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*12018' OR Name='Plate Maker') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*12018','Plate Maker','Plate Maker','12','0','000001',GetDate(),'NULL',NULL,'N','N','*26030')"
''Departments_TYPE-13
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*13001' OR Name='Editorial Department') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*13001','Editorial Department','Editorial Department','13','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*13002' OR Name='Production Department') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*13002','Production Department','Production Department','13','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*13003' OR Name='Sales Department') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*13003','Sales Department','Sales Department','13','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*13004' OR Name='Contracts Department and Legal Department') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*13004','Contracts Department and Legal Department','Contracts Department and Legal Department','13','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*13005' OR Name='Managing Editorial and Production') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*13005','Managing Editorial and Production','Managing Editorial and Production','13','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*13006' OR Name='Creative Departments') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*13006','Creative Departments','Creative Departments','13','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*13007' OR Name='Subsidiary Rights Departments') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*13007','Subsidiary Rights Departments','Subsidiary Rights Departments','13','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*13008' OR Name='Marketing, Promotion, and Advertising Departments') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*13008','Marketing, Promotion, and Advertising Departments','Marketing, Promotion, and Advertising Departments','13','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*13009' OR Name='Publicity Department') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*13009','Publicity Department','Publicity Department','13','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*13010' OR Name='Publisher Website Maintenance') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*13010','Publisher Website Maintenance','Publisher Website Maintenance','13','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*13011' OR Name='Finance and Accounting') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*13011','Finance and Accounting','Finance and Accounting','13','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*13012' OR Name='Information Technology (IT)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*13012','Information Technology (IT)','Information Technology (IT)','13','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*13013' OR Name='Human Resources (HR)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*13013','Human Resources (HR)','Human Resources (HR)','13','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
''Designation_TYPE-14
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*14001' OR Name='Editor-in-Chief') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*14001','Editor-in-Chief','Editor-in-Chief','14','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*14002' OR Name='Managing editor') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*14002','Managing editor','Managing editor','14','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*14003' OR Name='Editors') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*14003','Editors','Editors','14','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*14004' OR Name='Author/Writers') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*14004','Author/Writers','Author/Writers','14','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*14005' OR Name='Fact-checkers') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*14005','Fact-checkers','Fact-checkers','14','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*14006' OR Name='Graphic Designer') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*14006','Graphic Designer','Graphic Designer','14','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*14007' OR Name='Production manager') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*14007','Production manager','Production manager','14','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*14008' OR Name='DTP-Operator') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*14008','DTP-Operator','DTP-Operator','14','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*14009' OR Name='Proof Reader') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*14009','Proof Reader','Proof Reader','14','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
''Paper Unit Master_TYPE-15
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15001' OR Name='Gross') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*15001','Gross','Gross','15','144','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15002' OR Name='Packet(100)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*15002','Packet(100)','Packet(100)','15','100','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15003' OR Name='Packet(150)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*15003','Packet(150)','Packet(150)','15','150','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15004' OR Name='Ream') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*15004','Ream','Ream','15','500','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15005' OR Name='Reel') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*15005','Reel','Reel','15','500','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15006' OR Name='Bundle (700)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*15006','Bundle (700)','Bundle (700)','15','700','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15007' OR Name='Packet(200)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*15007','Packet(200)','Packet(200)','15','200','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15008' OR Name='PACKET') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*15008','PACKET','PACKET','15','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15009' OR Name='Sheet') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*15009','Sheet','Sheet','15','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15010' OR Name='Packet (250)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*15010','Packet (250)','Packet (250)','15','250','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
''Paper Quality Master_TYPE-16
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*16001' OR Name='Coated Matt') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*16001','Coated Matt','Coated Matt','16','0.95','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*16002' OR Name='Coated Gloss') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*16002','Coated Gloss','Coated Gloss','16','0.9','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*16003' OR Name=' Uncoated') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*16003',' Uncoated','Uncoated','16','1.35','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*16004' OR Name='High Bulk') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*16004','High Bulk','High Bulk','16','1.4','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
''Narration
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*17001' OR Name='1. Printing & Finishing Charges of') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*17001','1. Printing & Finishing Charges of','Printing & Finishing Charges of','17','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*17002' OR Name='1. Text Printing Charges of') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*17002','1. Text Printing Charges of','Text Printing Charges of','17','2','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*17003' OR Name='2. Title Printing Charges of') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*17003','2. Title Printing Charges of','Title Printing Charges of','17','3','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*17004' OR Name='3. Combo Title Printing Charges of') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*17004','3. Combo Title Printing Charges of','Combo Title Printing Charges of','17','4','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*17005' OR Name='4. Finishing Charges of') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*17005','4. Finishing Charges of','Finishing Charges of','17','5','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*17006' OR Name='5. Binding Charges of') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*17006','5. Binding Charges of','Binding Charges of','17','6','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*17007' OR Name='7. Title Printing & Finishing Charges of') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*17007','7. Title Printing & Finishing Charges of','Title Printing & Finishing Charges of','17','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*17008' OR Name='6. Text Printing & Finishing Charges of') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*17008','6. Text Printing & Finishing Charges of','Text Printing & Finishing Charges of','17','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*17009' OR Name='8. Unit Cost Charges of') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*17009','8. Unit Cost Charges of','Unit Cost Charges of','17','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*17010' OR Name='9. Unit Cost') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*17010','9. Unit Cost','.','17','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*17011' OR Name='10 Lamination Charges') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*17011','10 Lamination Charges','Lamination Charges','17','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*17012' OR Name='11 Printed Book') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*17012','11 Printed Book','Printed Book','17','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
''HSN MASTER_TYPE-18
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*18001' OR Name='998812') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*18001','998812','998812','18','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*18002' OR Name='998912') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*18002','998912','998912','18','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*18003' OR Name='4901') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*18003','4901','4901','18','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*18004' OR Name='49011010') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*18004','49011010','49011010','18','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
''Elements MASTER_TYPE-19
'        'eLEMENT mASTER mOVED TO eLEMENT mASTER
''Calculation Units MASTER_Type-20
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*20001' OR Name='Per Unit') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*20001','Per Unit','Per Unit','20','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*20002' OR Name='Per Inch²') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*20002','Per Inch²','Per Inch²','20','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*20003' OR Name='100 Inch²') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*20003','100 Inch²','100 Inch²','20','100','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*20004' OR Name='1000 Inch²') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*20004','1000 Inch²','1000 Inch²','20','1000','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*20005' OR Name='Per 1000') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*20005','Per 1000','Per 1000','20','1000','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
''Machine Master_TYPE-21
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*21046' OR Name='Machine To Be Decide') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*21046','Machine To Be Decide','Machine To Be Decide','21','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*21047' OR Name='RYOBI - 4 Col') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*21047','RYOBI - 4 Col','RYOBI - 4 Col','21','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*21048' OR Name='SM 102 28x40') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*21048','SM 102 28x40','SM 102 28x40','21','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*21049' OR Name='SM 74 20x29') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*21049','SM 74 20x29','SM 74 20x29','21','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*21050' OR Name='Heidel 2 Col') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*21050','Heidel 2 Col','Heidel 2 Col','21','0','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
''General  Unit MasterTYPE-25
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*25001' OR Name='Kilogram') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*25001','Kilogram','kg.','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*25002' OR Name='Gram') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*25002','Gram','gm.','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*25003' OR Name='Milligram') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*25003','Milligram','mg.','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*25004' OR Name='Liter') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*25004','Liter','ltr.','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*25005' OR Name='Milliliter') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*25005','Milliliter','ml.','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*25006' OR Name='Feet') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*25006','Feet','ft.','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*25007' OR Name='Inch') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*25007','Inch','in.','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*25008' OR Name='Meter') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*25008','Meter','mtr.','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*25009' OR Name='Centimeter') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*25009','Centimeter','cm.','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*25010' OR Name='Millimeter') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*25010','Millimeter','mm.','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*25011' OR Name='Piece') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*25011','Piece','pec.','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*25012' OR Name='Bags') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*25012','Bags','bags','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*25013' OR Name='Roll') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*25013','Roll','roll','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*25014' OR Name='Sets') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*25014','Sets','sets','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*25015' OR Name='Packets') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*25015','Packets','packets','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*25016' OR Name='Gross') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*25016','Gross','gross','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*25017' OR Name='Dozen') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*25017','Dozen','dozen','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*25018' OR Name='Tonn') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*25018','Tonn','tonn','25','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
''Account Group_TYPE-26
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26001' OR Name='Profit & Loss') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26001','Profit & Loss','Profit & Loss','26','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26002' OR Name='Revenue Accounts') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26002','Revenue Accounts','Revenue Accounts','26','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26003' OR Name='Stock-in-hand') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26003','Stock-in-hand','Stock-in-hand','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26008')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26004' OR Name='Bank Accounts') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26004','Bank Accounts','Bank Accounts','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26008')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26005' OR Name='Bank O/D Account') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26005','Bank O/D Account','Bank O/D Account','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26022')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26006' OR Name='Capital Account') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26006','Capital Account','Capital Account','26','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26007' OR Name='Cash-in-hand') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26007','Cash-in-hand','Cash-in-hand','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26008')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26008' OR Name='Current Assets') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26008','Current Assets','Current Assets','26','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26009' OR Name='Current Liabilities') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26009','Current Liabilities','Current Liabilities','26','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26010' OR Name='Depreciation Res On Machine') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26010','Depreciation Res On Machine','Depreciation Res On Machine','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26016')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26011' OR Name='Duties & Taxes') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26011','Duties & Taxes','Duties & Taxes','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26009')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26012' OR Name='Expenses (Direct/Mfg.)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26012','Expenses (Direct/Mfg.)','Expenses (Direct/Mfg.)','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26002')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26013' OR Name='Expenses (Indirect/Admn.)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26013','Expenses (Indirect/Admn.)','Expenses (Indirect/Admn.)','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26002')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26014' OR Name='File-Sundry Creditors') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26014','File-Sundry Creditors','File-Sundry Creditors','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26030')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26015' OR Name='File-Sundry Debtors') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26015','File-Sundry Debtors','File-Sundry Debtors','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26031')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26016' OR Name='Fixed Assets') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26016','Fixed Assets','Fixed Assets','26','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26017' OR Name='Income (Direct/Opr.)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26017','Income (Direct/Opr.)','Income (Direct/Opr.)','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26002')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26018' OR Name='Income (Indirect)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26018','Income (Indirect)','Income (Indirect)','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26002')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26019' OR Name='Income Tax Advance') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26019','Income Tax Advance','Income Tax Advance','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26021')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26020' OR Name='Investments') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26020','Investments','Investments','26','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26021' OR Name='Loans & Advances (Asset)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26021','Loans & Advances (Asset)','Loans & Advances (Asset)','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26008')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26022' OR Name='Loans (Liability)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26022','Loans (Liability)','Loans (Liability)','26','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26023' OR Name='Pre-Operative Expenses') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26023','Pre-Operative Expenses','Pre-Operative Expenses','26','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26024' OR Name='Provisions/Expenses Payable') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26024','Provisions/Expenses Payable','Provisions/Expenses Payable','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26009')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26025' OR Name='Purchase') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26025','Purchase','Purchase','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26002')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26026' OR Name='Reserves & Surplus') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26026','Reserves & Surplus','Reserves & Surplus','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26006')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26027' OR Name='Sale') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26027','Sale','Sale','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26002')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26028' OR Name='Secured Loans') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26028','Secured Loans','Secured Loans','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26022')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26029' OR Name='Securities & Deposits (Asset)') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26029','Securities & Deposits (Asset)','Securities & Deposits (Asset)','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26008')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26030' OR Name='Sundry Creditors') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26030','Sundry Creditors','Sundry Creditors','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26009')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26031' OR Name='Sundry Debtors') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26031','Sundry Debtors','Sundry Debtors','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26008')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26032' OR Name='Suspense Account') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26032','Suspense Account','Suspense Account','26','1','000001',GetDate(),'NULL',NULL,'N','N','NULL')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*26033' OR Name='Unsecured Loans') Print 'Exist' ELSE Insert Into GeneralMaster VALUES ('*26033','Unsecured Loans','Unsecured Loans','26','0','000001',GetDate(),'NULL',NULL,'N','N','*26022')"
''Finance Master
'    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01001' OR Name='Cash') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26007') Insert Into AccountMaster VALUES ('*01001','Cash','Cash','1001','*26007','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
'    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01002' OR Name='Development Tax') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26011') Insert Into AccountMaster VALUES ('*01002','Development Tax','Development Tax','1002','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
'    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01003' OR Name='Edu. Cess on TDS') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26011') Insert Into AccountMaster VALUES ('*01003','Edu. Cess on TDS','Edu. Cess on TDS','1003','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
'    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01004' OR Name='Excise Duty') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26011') Insert Into AccountMaster VALUES ('*01004','Excise Duty','Excise Duty','1004','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
'    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01005' OR Name='KKC on Service Tax') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26011') Insert Into AccountMaster VALUES ('*01005','KKC on Service Tax','KKC on Service Tax','1005','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
'    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01006' OR Name='SBC on Service Tax') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26011') Insert Into AccountMaster VALUES ('*01006','SBC on Service Tax','SBC on Service Tax','1006','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
'    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01007' OR Name='Service Tax') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26011') Insert Into AccountMaster VALUES ('*01007','Service Tax','Service Tax','1007','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
'    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01008' OR Name='SHE Cess on TDS') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26011') Insert Into AccountMaster VALUES ('*01008','SHE Cess on TDS','SHE Cess on TDS','1008','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
'    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01009' OR Name='TDS (Commission or Brokerage)') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26011') Insert Into AccountMaster VALUES ('*01009','TDS (Commission or Brokerage)','TDS (Commission or Brokerage)','1009','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
'    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01010' OR Name='TDS (Contracts to Individuals/HUF)') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26011') Insert Into AccountMaster VALUES ('*01010','TDS (Contracts to Individuals/HUF)','TDS (Contracts to Individuals/HUF)','1010','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
'    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01011' OR Name='TDS (Contracts to Others)') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26011') Insert Into AccountMaster VALUES ('*01011','TDS (Contracts to Others)','TDS (Contracts to Others)','1011','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
'    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01012' OR Name='TDS (Contracts to Transporter)') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26011') Insert Into AccountMaster VALUES ('*01012','TDS (Contracts to Transporter)','TDS (Contracts to Transporter)','1012','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
'    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01013' OR Name='TDS (Interest from a Banking Co)') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26011') Insert Into AccountMaster VALUES ('*01013','TDS (Interest from a Banking Co)','TDS (Interest from a Banking Co)','1013','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
'    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01014' OR Name='TDS (Interest from a NonBanking Co)') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26011') Insert Into AccountMaster VALUES ('*01014','TDS (Interest from a NonBanking Co)','TDS (Interest from a NonBanking Co)','1014','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
'    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01015' OR Name='TDS (Professionals Services)') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26011') Insert Into AccountMaster VALUES ('*01015','TDS (Professionals Services)','TDS (Professionals Services)','1015','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
'    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01016' OR Name='TDS (Rent of Land)') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26011') Insert Into AccountMaster VALUES ('*01016','TDS (Rent of Land)','TDS (Rent of Land)','1016','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
'    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01017' OR Name='TDS (Rent of Plant & Machinery)') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26011') Insert Into AccountMaster VALUES ('*01017','TDS (Rent of Plant & Machinery)','TDS (Rent of Plant & Machinery)','1017','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
'    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01018' OR Name='TDS (Salary)') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26011') Insert Into AccountMaster VALUES ('*01018','TDS (Salary)','TDS (Salary)','1018','*26011','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
'    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01019' OR Name='Advertisement & Publicity') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26013') Insert Into AccountMaster VALUES ('*01019','Advertisement & Publicity','Advertisement & Publicity','1019','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
'    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01020' OR Name='Bad Debts Written Off') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26013') Insert Into AccountMaster VALUES ('*01020','Bad Debts Written Off','Bad Debts Written Off','1020','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
'    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01021' OR Name='Bank Charges') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26013') Insert Into AccountMaster VALUES ('*01021','Bank Charges','Bank Charges','1021','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
'    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01022' OR Name='Books & Periodicals') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26013') Insert Into AccountMaster VALUES ('*01022','Books & Periodicals','Books & Periodicals','1022','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
'    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01023' OR Name='Charity & Donations') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26013') Insert Into AccountMaster VALUES ('*01023','Charity & Donations','Charity & Donations','1023','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
'    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01024' OR Name='Commission on Sales') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26013') Insert Into AccountMaster VALUES ('*01024','Commission on Sales','Commission on Sales','1024','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
'    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01025' OR Name='Conveyance Expenses') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26013') Insert Into AccountMaster VALUES ('*01025','Conveyance Expenses','Conveyance Expenses','1025','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
'    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01026' OR Name='Customer Entertainment Expenses') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26013') Insert Into AccountMaster VALUES ('*01026','Customer Entertainment Expenses','Customer Entertainment Expenses','1026','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
'    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01027' OR Name='Depreciation A/c') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26013') Insert Into AccountMaster VALUES ('*01027','Depreciation A/c','Depreciation A/c','1027','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
'    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01028' OR Name='Freight & Forwarding Charges') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26013') Insert Into AccountMaster VALUES ('*01028','Freight & Forwarding Charges','Freight & Forwarding Charges','1028','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
'    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01029' OR Name='Legal Expenses') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26013') Insert Into AccountMaster VALUES ('*01029','Legal Expenses','Legal Expenses','1029','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
'    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01030' OR Name='Miscellaneous Expenses') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26013') Insert Into AccountMaster VALUES ('*01030','Miscellaneous Expenses','Miscellaneous Expenses','1030','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
'    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01031' OR Name='Office Maintenance Expenses') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26013') Insert Into AccountMaster VALUES ('*01031','Office Maintenance Expenses','Office Maintenance Expenses','1031','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
'    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01032' OR Name='Office Rent') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26013') Insert Into AccountMaster VALUES ('*01032','Office Rent','Office Rent','1032','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
'    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01033' OR Name='Postal Expenses') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26013') Insert Into AccountMaster VALUES ('*01033','Postal Expenses','Postal Expenses','1033','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
'    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01034' OR Name='Printing & Stationery') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26013') Insert Into AccountMaster VALUES ('*01034','Printing & Stationery','Printing & Stationery','1034','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
'    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01035' OR Name='Rounded Off') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26013') Insert Into AccountMaster VALUES ('*01035','Rounded Off','Rounded Off','1035','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
'    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01036' OR Name='Salary') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26013') Insert Into AccountMaster VALUES ('*01036','Salary','Salary','1036','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
'    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01037' OR Name='Sales Promotion Expenses') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26013') Insert Into AccountMaster VALUES ('*01037','Sales Promotion Expenses','Sales Promotion Expenses','1037','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
'    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01038' OR Name='Service Charges Paid') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26013') Insert Into AccountMaster VALUES ('*01038','Service Charges Paid','Service Charges Paid','1038','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
'    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01039' OR Name='Staff Welfare Expenses') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26013') Insert Into AccountMaster VALUES ('*01039','Staff Welfare Expenses','Staff Welfare Expenses','1039','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
'    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01040' OR Name='Telephone Expenses') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26013') Insert Into AccountMaster VALUES ('*01040','Telephone Expenses','Telephone Expenses','1040','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
'    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01041' OR Name='Travelling Expenses') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26013') Insert Into AccountMaster VALUES ('*01041','Travelling Expenses','Travelling Expenses','1041','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
'    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01042' OR Name='Water & Electricity Expenses') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26013') Insert Into AccountMaster VALUES ('*01042','Water & Electricity Expenses','Water & Electricity Expenses','1042','*26013','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
'    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01043' OR Name='Capital Equipments') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26016') Insert Into AccountMaster VALUES ('*01043','Capital Equipments','Capital Equipments','1043','*26016','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
'    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01044' OR Name='Computers') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26016') Insert Into AccountMaster VALUES ('*01044','Computers','Computers','1044','*26016','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
'    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01045' OR Name='Furniture & Fixture') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26016') Insert Into AccountMaster VALUES ('*01045','Furniture & Fixture','Furniture & Fixture','1045','*26016','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
'    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01046' OR Name='Office Equipments') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26016') Insert Into AccountMaster VALUES ('*01046','Office Equipments','Office Equipments','1046','*26016','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
'    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01047' OR Name='Plant & Machinery') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26016') Insert Into AccountMaster VALUES ('*01047','Plant & Machinery','Plant & Machinery','1047','*26016','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
'    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01048' OR Name='Service Charges Receipts') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26018') Insert Into AccountMaster VALUES ('*01048','Service Charges Receipts','Service Charges Receipts','1048','*26018','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
'    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01049' OR Name='Profit & Loss') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26001') Insert Into AccountMaster VALUES ('*01049','Profit & Loss','Profit & Loss','1049','*26001','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
'    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01050' OR Name='Salary & Bonus Payable') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26024') Insert Into AccountMaster VALUES ('*01050','Salary & Bonus Payable','Salary & Bonus Payable','1050','*26024','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
'    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01051' OR Name='Purchase') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26025') Insert Into AccountMaster VALUES ('*01051','Purchase','Purchase','1051','*26025','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
'    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01052' OR Name='Sales') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26027') Insert Into AccountMaster VALUES ('*01052','Sales','Sales','1052','*26027','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
'    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01053' OR Name='Earnest Money') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26029') Insert Into AccountMaster VALUES ('*01053','Earnest Money','Earnest Money','1053','*26029','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
'    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01054' OR Name='Stock') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26003') Insert Into AccountMaster VALUES ('*01054','Stock','Stock','1054','*26003','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
'    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01055' OR Name='Easy Info Solutions International') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26030') Insert Into AccountMaster VALUES ('*01055','Easy Info Solutions International','Easy Info Solutions International','1055','*26030','E-461, Vijay Marg,Jagjeet Nagar','Delhi-110053','','','','+91-987-342-2907','','sales@easyinfosolution.com ','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
'    cnDatabase.Execute "IF EXISTS (SELECT * FROM AccountMaster WHERE Code='*01056' OR Name='XXX Bank') Print 'Exist' ELSE  IF EXISTS (SELECT CODE FROM GeneralMaster Where Code='*26004') Insert Into AccountMaster VALUES ('*01056','XXX Bank','XXX Bank','1056','*26004','','','','','','','','','1','000001',GetDate(),NULL,NULL,'N','N','',0) ELSE  Print 'NOT Exist'"
'
''Booking Route Master
'   cnDatabase.Execute "IF EXISTS (SELECT Code FROM BookingRouteMaster WHERE Code='*00001' OR Name='NOIDA-NOIDA') Print 'Exist' ELSE Insert Into BookingRouteMaster VALUES ('*00001','NOIDA-NOIDA','NOIDA-NOIDA','24.5','N')"
'   cnDatabase.Execute "IF EXISTS (SELECT Code FROM BookingRouteMaster WHERE Code='*00002' OR Name='NOIDA-DELHI') Print 'Exist' ELSE Insert Into BookingRouteMaster VALUES ('*00002','NOIDA-DELHI','NOIDA-DELHI','40','N')"
'   cnDatabase.Execute "IF EXISTS (SELECT Code FROM BookingRouteMaster WHERE Code='*00003' OR Name='DELHI-DELHI') Print 'Exist' ELSE Insert Into BookingRouteMaster VALUES ('*00003','DELHI-DELHI','DELHI-DELHI','30','N')"
'
''Element Master
'   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00011' OR NAME='Text-1') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00011','Text-1','Text-1','Single Sheet','8','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00012' OR NAME='Text-2') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00012','Text-2','Text-2','Multi Forms','8','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00013' OR NAME='Text-3') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00013','Text-3','Text-3','Multi Forms','8','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00014' OR NAME='Single Form') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00014','Single Form','Single Form','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00015' OR NAME='Combo Form') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00015','Combo Form','Combo Form','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00016' OR NAME='FG') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00016','FG','FG','FG','8','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00017' OR NAME='UFG') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00017','UFG','UFG','UFG','8','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00018' OR NAME='Separator') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00018','Separator','Separator','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00019' OR NAME='End Paper') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00019','End Paper','End Paper','Single Sheet','4','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00020' OR NAME='Cover') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00020','Cover','Cover','Single Sheet','4','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00027' OR NAME='Title') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00027','Title','Title','Single Sheet','4','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00028' OR NAME='Title(GateFold)') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00028','Title(GateFold)','Title(GateFold)','Single Sheet','6','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00029' OR NAME='PLC') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00029','PLC','PLC','Single Sheet','4','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00030' OR NAME='Calendar Fly Leaf') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00030','Calendar Fly Leaf','Calendar Fly Leaf','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00031' OR NAME='Calendar Leaf') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00031','Calendar Leaf','Calendar Leaf','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00032' OR NAME='Annual Report') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00032','Annual Report','Annual Report','Multi Forms','8','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00033' OR NAME='Label') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00033','Label','Label','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00034' OR NAME='Letter Head') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00034','Letter Head','Letter Head','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00035' OR NAME='Leaflet') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00035','Leaflet','Leaflet','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00036' OR NAME='Poster') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00036','Poster','Poster','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00037' OR NAME='Sticker') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00037','Sticker','Sticker','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00038' OR NAME='Folders') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00038','Folders','Folders','Single Sheet','4','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00039' OR NAME='Dust Cover') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00039','Dust Cover','Dust Cover','Single Sheet','6','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00040' OR NAME='Danglar') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00040','Danglar','Danglar','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00041' OR NAME='Carton') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00041','Carton','Carton','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00042' OR NAME='Carton [Inner]') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00042','Carton [Inner]','Carton [Inner]','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00043' OR NAME='Carton [Outer]') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00043','Carton [Outer]','Carton [Outer]','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00044' OR NAME='Card') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00044','Card','Card','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'   cnDatabase.Execute "IF EXISTS (SELECT Code FROM ElementMaster WHERE Code='*00045' OR NAME='Envelope') Print 'Exist' ELSE Insert Into ElementMaster VALUES ('*00045','Envelope','Envelope','Single Sheet','2','0','0','0','000001',GetDate(),'NULL',NULL,'N','N')"
'
''Finish Size Master
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01039' AND CODE='*01017')  Insert Into FinishSizeChild VALUES ('*11011','*01039','16','16','*01017') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01030' AND CODE='*01031')  Insert Into FinishSizeChild VALUES ('*11012','*01030','16','16','*01031') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01064' AND CODE='*01031')  Insert Into FinishSizeChild VALUES ('*11012','*01064','32','16','*01031') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01039' AND CODE='*01017')  Insert Into FinishSizeChild VALUES ('*11013','*01039','16','16','*01017') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01039' AND CODE='*01017')  Insert Into FinishSizeChild VALUES ('*11014','*01039','16','16','*01017') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01055' AND CODE='*01028')  Insert Into FinishSizeChild VALUES ('*11015','*01055','16','16','*01028') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01048' AND CODE='*01017')  Insert Into FinishSizeChild VALUES ('*11016','*01048','16','16','*01017') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01051' AND CODE='*01028')  Insert Into FinishSizeChild VALUES ('*11017','*01051','16','16','*01028') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01058' AND CODE='*01028')  Insert Into FinishSizeChild VALUES ('*11018','*01058','16','16','*01028') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01056' AND CODE='*01028')  Insert Into FinishSizeChild VALUES ('*11019','*01056','16','16','*01028') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01028' AND CODE='*01028')  Insert Into FinishSizeChild VALUES ('*11020','*01028','8','16','*01028') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01060' AND CODE='*01028')  Insert Into FinishSizeChild VALUES ('*11020','*01060','16','16','*01028') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01067' AND CODE='*01031')  Insert Into FinishSizeChild VALUES ('*11021','*01067','16','16','*01031') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01039' AND CODE='*01017')  Insert Into FinishSizeChild VALUES ('*11033','*01039','16','16','*01017') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01031' AND CODE='*01031')  Insert Into FinishSizeChild VALUES ('*11023','*01031','8','16','*01031') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01067' AND CODE='*01031')  Insert Into FinishSizeChild VALUES ('*11023','*01067','16','16','*01031') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01033' AND CODE='*01033')  Insert Into FinishSizeChild VALUES ('*11024','*01033','8','16','*01033') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01068' AND CODE='*01033')  Insert Into FinishSizeChild VALUES ('*11024','*01068','16','16','*01033') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01068' AND CODE='*01017')  Insert Into FinishSizeChild VALUES ('*11025','*01068','16','16','*01017') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01037' AND CODE='*01017')  Insert Into FinishSizeChild VALUES ('*11026','*01037','8','16','*01017') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01070' AND CODE='*01017')  Insert Into FinishSizeChild VALUES ('*11026','*01070','16','16','*01017') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01054' AND CODE='*01017')  Insert Into FinishSizeChild VALUES ('*11027','*01054','8','16','*01017') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01072' AND CODE='*01017')  Insert Into FinishSizeChild VALUES ('*11028','*01072','16','16','*01017') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01038' AND CODE='*01017')  Insert Into FinishSizeChild VALUES ('*11029','*01038','8','16','*01017') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01072' AND CODE='*01017')  Insert Into FinishSizeChild VALUES ('*11029','*01072','16','16','*01017') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01055' AND CODE='*01028')  Insert Into FinishSizeChild VALUES ('*11030','*01055','12','24','*01028') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01039' AND CODE='*01017')  Insert Into FinishSizeChild VALUES ('*11031','*01039','8','16','*01017') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01046' AND CODE='*01017')  Insert Into FinishSizeChild VALUES ('*11032','*01046','8','16','*01017') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01048' AND CODE='*01017')  Insert Into FinishSizeChild VALUES ('*11034','*01048','8','16','*01017') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01063' AND CODE='*01031')  Insert Into FinishSizeChild VALUES ('*11035','*01063','12','24','*01031') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01048' AND CODE='*01017')  Insert Into FinishSizeChild VALUES ('*11036','*01048','8','16','*01017') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01048' AND CODE='*01017')  Insert Into FinishSizeChild VALUES ('*11037','*01048','8','16','*01017') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01055' AND CODE='*01028')  Insert Into FinishSizeChild VALUES ('*11038','*01055','8','16','*01028') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01067' AND CODE='*01031')  Insert Into FinishSizeChild VALUES ('*11039','*01067','12','24','*01031') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01050' AND CODE='*01028')  Insert Into FinishSizeChild VALUES ('*11040','*01050','8','16','*01028') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01058' AND CODE='*01028')  Insert Into FinishSizeChild VALUES ('*11041','*01058','8','16','*01028') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01027' AND CODE='*01031')  Insert Into FinishSizeChild VALUES ('*11042','*01027','4','8','*01031') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01070' AND CODE='*01017')  Insert Into FinishSizeChild VALUES ('*11043','*01070','12','24','*01017') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01060' AND CODE='*01028')  Insert Into FinishSizeChild VALUES ('*11044','*01060','8','16','*01028') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01060' AND CODE='*01028')  Insert Into FinishSizeChild VALUES ('*11046','*01060','8','16','*01028') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01039' AND CODE='*01017')  Insert Into FinishSizeChild VALUES ('*11047','*01039','6','12','*01017') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01073' AND CODE='*01017')  Insert Into FinishSizeChild VALUES ('*11049','*01073','16','16','*01017') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01068' AND CODE='*01031')  Insert Into FinishSizeChild VALUES ('*11050','*01068','8','16','*01031') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01055' AND CODE='*01028')  Insert Into FinishSizeChild VALUES ('*11051','*01055','6','12','*01028') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01072' AND CODE='*01017')  Insert Into FinishSizeChild VALUES ('*11052','*01072','6','12','*01017') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01060' AND CODE='*01028')  Insert Into FinishSizeChild VALUES ('*11053','*01060','4','8','*01028') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01068' AND CODE='*01031')  Insert Into FinishSizeChild VALUES ('*11054','*01068','4','8','*01031') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01070' AND CODE='*01017')  Insert Into FinishSizeChild VALUES ('*11055','*01070','6','12','*01017') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01073' AND CODE='*01017')  Insert Into FinishSizeChild VALUES ('*11004','*01073','4','8','*01017') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01039' AND CODE='*01012')  Insert Into FinishSizeChild VALUES ('*11004','*01039','2','2','*01012') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01012' AND CODE='*01012')  Insert Into FinishSizeChild VALUES ('*11004','*01012','1','1','*01012') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01055' AND CODE='*01028')  Insert Into FinishSizeChild VALUES ('*11057','*01055','8','16','*01028') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01028' AND CODE='*01029')  Insert Into FinishSizeChild VALUES ('*11048','*01028','4','8','*01029') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01059' AND CODE='*01029')  Insert Into FinishSizeChild VALUES ('*11048','*01059','8','16','*01029') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01058' AND CODE='*01029')  Insert Into FinishSizeChild VALUES ('*11058','*01058','8','16','*01029') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01063' AND CODE='*01031')  Insert Into FinishSizeChild VALUES ('*11045','*01063','8','16','*01031') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01060' AND CODE='*01029')  Insert Into FinishSizeChild VALUES ('*11085','*01060','12','24','*01029') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01011' AND CODE='*01011')  Insert Into FinishSizeChild VALUES ('*11005','*01011','2','2','*01011') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01065' AND CODE='*01029')  Insert Into FinishSizeChild VALUES ('*11092','*01065','8','16','*01029') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01028' AND CODE='*01029')  Insert Into FinishSizeChild VALUES ('*11091','*01028','4','8','*01029') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01067' AND CODE='*01029')  Insert Into FinishSizeChild VALUES ('*11091','*01067','8','16','*01029') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01028' AND CODE='*01028')  Insert Into FinishSizeChild VALUES ('*11094','*01028','8','16','*01028') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01068' AND CODE='*01028')  Insert Into FinishSizeChild VALUES ('*11094','*01068','16','16','*01028') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01072' AND CODE='*01031')  Insert Into FinishSizeChild VALUES ('*11022','*01072','16','16','*01031') ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*01045' AND CODE='*01047')  Insert Into FinishSizeChild VALUES ('*11095','*01045','8','8','*01047') ELSE Print 'NOT Exist' "
'
''SizeGroup
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10003' AND CODE='*01067') Insert Into SizeGroupChild VALUES ('*10003','*01067')  ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10003' AND CODE='*01068') Insert Into SizeGroupChild VALUES ('*10003','*01068')  ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10003' AND CODE='*01070') Insert Into SizeGroupChild VALUES ('*10003','*01070')  ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10003' AND CODE='*01072') Insert Into SizeGroupChild VALUES ('*10003','*01072')  ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10003' AND CODE='*01073') Insert Into SizeGroupChild VALUES ('*10003','*01073')  ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10007' AND CODE='*01061') Insert Into SizeGroupChild VALUES ('*10007','*01061')  ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10011' AND CODE='*01047') Insert Into SizeGroupChild VALUES ('*10011','*01047')  ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10006' AND CODE='*01050') Insert Into SizeGroupChild VALUES ('*10006','*01050')  ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10006' AND CODE='*01051') Insert Into SizeGroupChild VALUES ('*10006','*01051')  ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10006' AND CODE='*01056') Insert Into SizeGroupChild VALUES ('*10006','*01056')  ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10006' AND CODE='*01058') Insert Into SizeGroupChild VALUES ('*10006','*01058')  ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10006' AND CODE='*01060') Insert Into SizeGroupChild VALUES ('*10006','*01060')  ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10006' AND CODE='*01063') Insert Into SizeGroupChild VALUES ('*10006','*01063')  ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10006' AND CODE='*01064') Insert Into SizeGroupChild VALUES ('*10006','*01064')  ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10006' AND CODE='*01059') Insert Into SizeGroupChild VALUES ('*10006','*01059')  ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10012' AND CODE='*01017') Insert Into SizeGroupChild VALUES ('*10012','*01017')  ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10012' AND CODE='*01020') Insert Into SizeGroupChild VALUES ('*10012','*01020')  ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10012' AND CODE='*01021') Insert Into SizeGroupChild VALUES ('*10012','*01021')  ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10012' AND CODE='*01027') Insert Into SizeGroupChild VALUES ('*10012','*01027')  ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10012' AND CODE='*01028') Insert Into SizeGroupChild VALUES ('*10012','*01028')  ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10012' AND CODE='*01030') Insert Into SizeGroupChild VALUES ('*10012','*01030')  ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10012' AND CODE='*01031') Insert Into SizeGroupChild VALUES ('*10012','*01031')  ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10012' AND CODE='*01033') Insert Into SizeGroupChild VALUES ('*10012','*01033')  ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10012' AND CODE='*01013') Insert Into SizeGroupChild VALUES ('*10012','*01013')  ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10013' AND CODE='*01012') Insert Into SizeGroupChild VALUES ('*10013','*01012')  ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10013' AND CODE='*01015') Insert Into SizeGroupChild VALUES ('*10013','*01015')  ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10013' AND CODE='*01016') Insert Into SizeGroupChild VALUES ('*10013','*01016')  ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10013' AND CODE='*01019') Insert Into SizeGroupChild VALUES ('*10013','*01019')  ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10013' AND CODE='*01029') Insert Into SizeGroupChild VALUES ('*10013','*01029')  ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10013' AND CODE='*01018') Insert Into SizeGroupChild VALUES ('*10013','*01018')  ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10018' AND CODE='*01069') Insert Into SizeGroupChild VALUES ('*10018','*01069')  ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10009' AND CODE='*01036') Insert Into SizeGroupChild VALUES ('*10009','*01036')  ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10009' AND CODE='*01037') Insert Into SizeGroupChild VALUES ('*10009','*01037')  ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10009' AND CODE='*01038') Insert Into SizeGroupChild VALUES ('*10009','*01038')  ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10009' AND CODE='*01039') Insert Into SizeGroupChild VALUES ('*10009','*01039')  ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10009' AND CODE='*01046') Insert Into SizeGroupChild VALUES ('*10009','*01046')  ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10009' AND CODE='*01048') Insert Into SizeGroupChild VALUES ('*10009','*01048')  ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10009' AND CODE='*01054') Insert Into SizeGroupChild VALUES ('*10009','*01054')  ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10009' AND CODE='*01057') Insert Into SizeGroupChild VALUES ('*10009','*01057')  ELSE Print 'NOT Exist' "
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*10020' AND CODE='*01011') Insert Into SizeGroupChild VALUES ('*10020','*01011')  ELSE Print 'NOT Exist' "
'
''Tax Master
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*00001' OR Name='Local GST 12%') Print 'Exist' ELSE Insert Into TaxMaster VALUES ('*00001','Local GST 12%','Local GST 12%','L','6','6',0,'000001',GetDate(),'NULL',NULL,'N','N')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*00002' OR Name='IGST 12%') Print 'Exist' ELSE Insert Into TaxMaster VALUES ('*00002','IGST 12%','IGST 12%','I','0','0',12,'000001',GetDate(),'NULL',NULL,'N','N')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*00003' OR Name='IGST 5%') Print 'Exist' ELSE Insert Into TaxMaster VALUES ('*00003','IGST 5%','IGST 5%','I','0','0',5,'000001',GetDate(),'NULL',NULL,'N','N')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*00004' OR Name='Local GST 5%') Print 'Exist' ELSE Insert Into TaxMaster VALUES ('*00004','Local GST 5%','Local GST 5%','L','2.5','2.5',0,'000001',GetDate(),'NULL',NULL,'N','N')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*00005' OR Name='Local GST 18%') Print 'Exist' ELSE Insert Into TaxMaster VALUES ('*00005','Local GST 18%','Local GST 18%','L','9','9',0,'000006',GetDate(),'NULL',NULL,'N','N')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*00006' OR Name='IGST 18%') Print 'Exist' ELSE Insert Into TaxMaster VALUES ('*00006','IGST 18%','IGST 18%','I','0','0',18,'000006',GetDate(),'NULL',NULL,'N','N')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*00007' OR Name='Local GST NIL') Print 'Exist' ELSE Insert Into TaxMaster VALUES ('*00007','Local GST NIL','Local GST NIL','L','0','0',0,'000001',GetDate(),'NULL',NULL,'N','N')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*00008' OR Name='IGST NIL') Print 'Exist' ELSE Insert Into TaxMaster VALUES ('*00008','IGST NIL','IGST NIL','I','0','0',0,'000001',GetDate(),'NULL',NULL,'N','N')"
'
''CompChild
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='01') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','01','1. Please send two copies of invoice.','2. Please notify us immediately if ','you are unable to ship as specified.','3. Enter this order in accordance, with the price,terms, ','delivery method and specification Listed above.','4. All disputes are subject to Our Jurisdiction Only','','SEPL/Pur/','/20-21','Purchase')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='02') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','02','1. Please send two copies of invoice.','2. Please notify us immediately if ','you are unable to ship as specified.','3. Enter this order in accordance, with the price,terms, ','delivery method and specification Listed above.','4. All disputes are subject to Our Jurisdiction Only','','SEPL/PR/','/20-21','Purchase Return')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='03') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','03','1. All disputes are subject to Our Jurisdiction Only','2. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection.','3. Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','','','SEPL/SR/','/20-21','Sale Return')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='04') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','04','1. Interest @24% p.a. will be charged if','the payment is not made in time.','2. All disputes are subject to Our Jurisdiction Only','3. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection','4. . Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','SEPL/Sale/','/20-21','Sale')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='05') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','05','1. Please send two copies of invoice.','2. Please notify us immediately if ','you are unable to ship as specified.','3. Enter this order in accordance, with the price,terms, ','delivery method and specification Listed above.','4. All disputes are subject to Our Jurisdiction Only','','SEPL/PC/','/20-21','Purchase Challan IN')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='06') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','06','','','','','','','','SEPL/PRC/','/20-21','Purchase Challan Out')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='07') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','07','1. Interest @24% p.a. will be charged if','the payment is not made in time.','2. All disputes are subject to Our Jurisdiction Only','3. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection','4. . Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','SEPL/SRC/','/20-21','Sale Challan IN')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='08') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','08','1. Interest @24% p.a. will be charged if','the payment is not made in time.','2. All disputes are subject to Our Jurisdiction Only','3. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection','4. . Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','SEPL/SC/','/20-21','Sale Challan Out')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='09') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','09','1. Interest @24% p.a. will be charged if','the payment is not made in time.','2. All disputes are subject to Our Jurisdiction Only','3. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection','4. . Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','SEPL/SJ/','/20-21','Sale Jobwork')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='10') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','10','1. Interest @24% p.a. will be charged if','the payment is not made in time.','2. All disputes are subject to Our Jurisdiction Only','3. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection','4. . Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','SEPL/SC/','/20-21','Sale Jobwork Unit Cost')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='11') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','11','1. Interest @24% p.a. will be charged if','the payment is not made in time.','2. All disputes are subject to Our Jurisdiction Only','3. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection','4. . Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','SEPL/DN/','/20-21','Challan Revesal IN')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='12') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','12','1. Interest @24% p.a. will be charged if','the payment is not made in time.','2. All disputes are subject to Delhi Jurisdiction Only','3. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection','4. . Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','SFAPL/PU/','/20-21','Challan Revesal Out')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='13') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','13','1. Interest @24% p.a. will be charged if','the payment is not made in time.','2. All disputes are subject to Our Jurisdiction Only','3. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection','4. . Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','SEPL/SC/','/20-21','Challan TO Be Billed IN')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='14') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','14','1. Interest @24% p.a. will be charged if','the payment is not made in time.','2. All disputes are subject to Our Jurisdiction Only','3. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection','4. . Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','SEPL/SC/','/20-21','Challan TO Be Billed OUT')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='15') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','15','1. Interest @24% p.a. will be charged if','the payment is not made in time.','2. All disputes are subject to Our Jurisdiction Only','3. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection','4. . Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','SEPL/SC/','/20-21','Challan Not TO Be Billed IN')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='16') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','16','1. Interest @24% p.a. will be charged if','the payment is not made in time.','2. All disputes are subject to Our Jurisdiction Only','3. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection','4. . Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','SEPL/SC/','/20-21','Challan Not TO Be Billed IOUT')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='17') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','17','1. The Deliverables shall be delivered or performed on the ','date and at the place specified in the Purchase Order.','2. Prices shall be as specified in the  Purchase  Order.','3. No increase in price shall be made or accepted unless ',' agreed in writing by Accenture.','4. The  Deliverables must conform in all respects with the','   Specifications and must be of sound.','SEPL/PO/','/20-21','Purchase Order')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='18') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','18','1. Interest @24% p.a. will be charged if','the payment is not made in time.','2. All disputes are subject to Our Jurisdiction Only','3. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection','4. . Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','SEPL/SO/','/20-21','Sale Order')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='19') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','19','1. Interest @24% p.a. will be charged if','the payment is not made in time.','2. All disputes are subject to Our Jurisdiction Only','3. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection','4. . Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','SEPL/ST/','/20-21','Stock Tranfer')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='20') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','20','','','','','','','','SEPL/RN/','/20-21','Stock Genral')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='21') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','21','1. Interest @24% p.a. will be charged if','the payment is not made in time.','2. All disputes are subject to Delhi Jurisdiction Only','3. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection','4. . Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','SFAPL/SU/','/20-21','Promotional Sale Challan Out')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='22') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','22','1. Interest @24% p.a. will be charged if','the payment is not made in time.','2. All disputes are subject to Our Jurisdiction Only','3. Rejection, if any shall be informed within one week from','the date of receipt in writing giving reason of rejection','4. . Please, Receive Following Goods in Good Condition.','after 7 days of the date of this Bill','SEPL/SQ/','/20-21','--')"
'   'Check err. String or binary data would be truncated.
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='23') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','23','1. The price set for in Supplier’s Quotation (Price) are',' in  INDIA INR.','2. All Taxes shall be paid by Customer in addition to the ',' Price.','3.  Quotation (Prices) are valid for 30 days only.','','','SEPL/QP/','/20-21','Purchase Quotation') "
'   'Check err. String or binary data would be truncated.
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='24') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','24','1. The price set for in Supplier’s Quotation (“Price”) are',' in  INDIA INR.','2. All Taxes shall be paid by Customer in addition to the ',' Price.','3.  Quotation (“Prices”) are valid for 30 days only.','','','SEPL/QS/','/20-21','Sales Quotation')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='25') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','25','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='26') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','26','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='27') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','27','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='28') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','28','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='29') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','29','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='30') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','30','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='31') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','31','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='32') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','32','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='33') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','33','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='34') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','34','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='35') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','35','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='36') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','36','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='37') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','37','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='38') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','38','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='39') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','39','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='40') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','40','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='41') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','41','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='42') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','42','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='43') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','43','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='44') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','44','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='45') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','45','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='46') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','46','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='47') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','47','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='48') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','48','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='49') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','49','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='50') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','50','','','','','','','','','','')"
'   'Check err. String or binary data would be truncated.
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='51') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','51','','','','','','','','SEPL/PI/','/20-21','Payment')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='52') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','52','','','','','','','','SEPL/PR/','/20-21','Receipt')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='53') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','53','','','','','','','','SEPL/JE/','/20-21','Journal')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='54') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','54','','','','','','','','SEPL/CE/','/20-21','Contra')"
'   'Check err. String or binary data would be truncated.
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='55') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','55','','','','','','','','SEPL/DN/','/20-21','Debit Note')"
'   'Check err. String or binary data would be truncated.
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='56') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','56','','','','','','','','SEPL/CN/','/20-21','Credit Note')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='57') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','57','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='58') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','58','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='59') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','59','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='60') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','60','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='61') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','61','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='62') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','62','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='63') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','63','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='64') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','64','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='65') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','65','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='66') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','66','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='67') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','67','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='68') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','68','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='69') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','69','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='70') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','70','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='71') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','71','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='72') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','72','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='73') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','73','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='74') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','74','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='75') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','75','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='76') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','76','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='77') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','77','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='78') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','78','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='79') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','79','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='80') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','80','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='81') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','81','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='82') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','82','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='83') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','83','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='84') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','84','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='85') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','85','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='86') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','86','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='87') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','87','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='88') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','88','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='89') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','89','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='90') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','90','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='91') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','91','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='92') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','92','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='93') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','93','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='94') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','94','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='95') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','95','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='96') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','96','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='97') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','97','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='98') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','98','','','','','','','','','','')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM CompChild WHERE Code='000001' AND VchType='99') Print 'Exist' ELSE Insert Into CompChild VALUES ('000001','99','','','','','','','','','','')"
'
''Vch Series Master
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*00101' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*00101','Main','01PF','SEPL/','/Purc','A')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*00102' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*00102','Main','01PU','SEPL/','/PrJU','A')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*00103' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*00103','Main','01PC','SEPL/','/PrJC','A')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*00104' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*00104','Main','01PJ','SEPL/','/PrJW','A')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*00201' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*00201','Main','02OF','SEPL/','/PrRt','A')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*00202' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*00202','Main','02OU','SEPL/','/PrRtJU','A')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*00203' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*00203','Main','02OC','SEPL/','/PrRtJC','A')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*00204' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*00204','Main','02OJ','SEPL/','/PrRtJW','A')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*00301' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*00301','Main','03TF','SEPL/','/SlRt','A')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*00302' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*00302','Main','03TU','SEPL/','/SlRtJU','A')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*00303' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*00303','Main','03TC','SEPL/','/SlRtJC','A')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*00304' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*00304','Main','03TJ','SEPL/','/SlRtJW','A')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*00401' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*00401','Main','04SF','SEPL/','/Sale','A')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*00402' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*00402','Main','04SU','SEPL/','/SlJU','A')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*00403' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*00403','Main','04SC','SEPL/','/SlJC','A')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*00404' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*00404','Main','04SJ','SEPL/','/SlJW','A')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*00501' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*00501','Main','05RF','SEPL/','/MtRc','A')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*00502' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*00502','Main','05FR','SEPL/','/MtRcJW','A')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*00601' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*00601','Main','06IF','SEPL/','/PrRtC','A')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*00602' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*00602','Main','06FI','SEPL/','/PrRtCJW','A')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*00701' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*00701','Main','07RF','SEPL/','/SlRtC','A')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*00702' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*00702','Main','07FR','SEPL/','/SlRtCJW','A')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*00801' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*00801','Main','08IF','SEPL/','/MtIs','A')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*00802' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*00802','Main','08FI','SEPL/','/MtIsJW','A')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*01701' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*01701','Main','17PO','SEPL/','/PO','A')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*01801' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*01801','Main','18SO','SEPL/','/SO','A')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*01901' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*01901','Main','19ST','SEPL/','/STrn','A')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*02001' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*02001','Main','20JR','SEPL/','/SJrnl','A')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*02101' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*02101','Main','21JR','SEPL/','/SJrnl','A')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*02201' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*02201','Main','22JR','SEPL/','/SJrnl','A')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*02301' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*02301','Main','23PQ','SEPL/','/PQ','A')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*02302' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*02302','Main','23UZ','SEPL/','/PQU','A')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*02303' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*02303','Main','23CZ','SEPL/','/PQC','A')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*02304' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*02304','Main','23JZ','SEPL/','/PQJ','A')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*02305' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*02305','Main','24SQ','SEPL/','/SQ','A')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*02306' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*02306','Main','24UQ','SEPL/','/SQU','A')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*02307' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*02307','Main','24CQ','SEPL/','/SQC','A')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*02308' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*02308','Main','24JQ','SEPL/','/SQJ','A')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*05101' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*05101','Main','51PI','SEPL/','/Pymt','A')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*05201' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*05201','Main','52PR','SEPL/','/Rcpt','A')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*05301' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*05301','Main','53JE','SEPL/','/Jrnl','A')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*05401' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*05401','Main','54CE','SEPL/','/Cntr','A')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*05501' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*05501','Main','55CN','SEPL/','/CrNt','A')"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM VchSeriesMaster WHERE Code='*05601' AND NAME='Main') Print 'Exist' ELSE Insert Into VchSeriesMaster VALUES ('*05601','Main','56DN','SEPL/','/DrNt','A')"
''Paper Master
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*00001' OR NAME='Art Card-200gsm-20.00X30.00in²-(50.80X76.20cm²)-7.742kg-Gloss') Print 'Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15002' AND Code='*16002') Insert Into PaperMaster VALUES ('*00001','Art Card-200gsm-20.00X30.00in²-(50.80X76.20cm²)-7.742kg-Gloss','Art Card-200gsm-20.00X30.00in²-(50.80X76.20cm²)-7.742kg-Gloss','S','B','50.8','76.2','20','30','*15002','200','Art Card','Gloss','7.742','6','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'NOT Exist'"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*00002' OR NAME='Art Card-210gsm-20.00X30.00in²-(50.80X76.20cm²)-8.129kg-Gloss') Print 'Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15002' AND Code='*16002') Insert Into PaperMaster VALUES ('*00002','Art Card-210gsm-20.00X30.00in²-(50.80X76.20cm²)-8.129kg-Gloss','Art Card-210gsm-20.00X30.00in²-(50.80X76.20cm²)-8.129kg-Gloss','S','B','50.8','76.2','20','30','*15002','210','Art Card','Gloss','8.129','6','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'NOT Exist'"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*00003' OR NAME='Art Card-220gsm-20.00X30.00in²-(50.80X76.20cm²)-8.516kg-Gloss') Print 'Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15002' AND Code='*16002') Insert Into PaperMaster VALUES ('*00003','Art Card-220gsm-20.00X30.00in²-(50.80X76.20cm²)-8.516kg-Gloss','Art Card-220gsm-20.00X30.00in²-(50.80X76.20cm²)-8.516kg-Gloss','S','B','50.8','76.2','20','30','*15002','220','Art Card','Gloss','8.516','6','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'NOT Exist'"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*00004' OR NAME='Art Card-250gsm-20.00X30.00in²-(50.80X76.20cm²)-9.677kg-Gloss') Print 'Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15002' AND Code='*16002') Insert Into PaperMaster VALUES ('*00004','Art Card-250gsm-20.00X30.00in²-(50.80X76.20cm²)-9.677kg-Gloss','Art Card-250gsm-20.00X30.00in²-(50.80X76.20cm²)-9.677kg-Gloss','S','B','50.8','76.2','20','30','*15002','250','Art Card','Gloss','9.677','5','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'NOT Exist'"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*00005' OR NAME='Art Card-200gsm-23.00X36.00in²-(58.42X91.44cm²)-10.684kg-Gloss') Print 'Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15002' AND Code='*16002') Insert Into PaperMaster VALUES ('*00005','Art Card-200gsm-23.00X36.00in²-(58.42X91.44cm²)-10.684kg-Gloss','Art Card-200gsm-23.00X36.00in²-(58.42X91.44cm²)-10.684kg-Gloss','S','B','58.42','91.44','23','36','*15002','200','Art Card','Gloss','10.684','5','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'NOT Exist'"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*00006' OR NAME='Art Card-210gsm-23.00X36.00in²-(58.42X91.44cm²)-11.218kg-Gloss') Print 'Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15002' AND Code='*16002') Insert Into PaperMaster VALUES ('*00006','Art Card-210gsm-23.00X36.00in²-(58.42X91.44cm²)-11.218kg-Gloss','Art Card-210gsm-23.00X36.00in²-(58.42X91.44cm²)-11.218kg-Gloss','S','B','58.42','91.44','23','36','*15002','210','Art Card','Gloss','11.218','4','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'NOT Exist'"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*00007' OR NAME='Art Card-220gsm-23.00X36.00in²-(58.42X91.44cm²)-11.752kg-Gloss') Print 'Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15002' AND Code='*16002') Insert Into PaperMaster VALUES ('*00007','Art Card-220gsm-23.00X36.00in²-(58.42X91.44cm²)-11.752kg-Gloss','Art Card-220gsm-23.00X36.00in²-(58.42X91.44cm²)-11.752kg-Gloss','S','B','58.42','91.44','23','36','*15002','220','Art Card','Gloss','11.752','4','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'NOT Exist'"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*00008' OR NAME='Art Card-250gsm-23.00X36.00in²-(58.42X91.44cm²)-13.355kg-Gloss') Print 'Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15002' AND Code='*16002') Insert Into PaperMaster VALUES ('*00008','Art Card-250gsm-23.00X36.00in²-(58.42X91.44cm²)-13.355kg-Gloss','Art Card-250gsm-23.00X36.00in²-(58.42X91.44cm²)-13.355kg-Gloss','S','B','58.42','91.44','23','36','*15002','250','Art Card','Gloss','13.355','4','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'NOT Exist'"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*00009' OR NAME='Art Paper-70gsm-20.00X30.00in²-(50.80X76.20cm²)-13.548kg-Gloss') Print 'Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15004' AND Code='*16002') Insert Into PaperMaster VALUES ('*00009','Art Paper-70gsm-20.00X30.00in²-(50.80X76.20cm²)-13.548kg-Gloss','Art Paper-70gsm-20.00X30.00in²-(50.80X76.20cm²)-13.548kg-Gloss','S','P','50.8','76.2','20','30','*15004','70','Art Paper','Gloss','13.548','4','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'NOT Exist'"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*00010' OR NAME='Art Paper-80gsm-20.00X30.00in²-(50.80X76.20cm²)-15.484kg-Gloss') Print 'Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15004' AND Code='*16002') Insert Into PaperMaster VALUES ('*00010','Art Paper-80gsm-20.00X30.00in²-(50.80X76.20cm²)-15.484kg-Gloss','Art Paper-80gsm-20.00X30.00in²-(50.80X76.20cm²)-15.484kg-Gloss','S','P','50.8','76.2','20','30','*15004','80','Art Paper','Gloss','15.484','3','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'NOT Exist'"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*00011' OR NAME='Art Paper-90gsm-20.00X30.00in²-(50.80X76.20cm²)-17.419kg-Gloss') Print 'Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15004' AND Code='*16002') Insert Into PaperMaster VALUES ('*00011','Art Paper-90gsm-20.00X30.00in²-(50.80X76.20cm²)-17.419kg-Gloss','Art Paper-90gsm-20.00X30.00in²-(50.80X76.20cm²)-17.419kg-Gloss','S','P','50.8','76.2','20','30','*15004','90','Art Paper','Gloss','17.419','3','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'NOT Exist'"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*00012' OR NAME='Art Paper-100gsm-20.00X30.00in²-(50.80X76.20cm²)-19.355kg-Gloss') Print 'Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15004' AND Code='*16002') Insert Into PaperMaster VALUES ('*00012','Art Paper-100gsm-20.00X30.00in²-(50.80X76.20cm²)-19.355kg-Gloss','Art Paper-100gsm-20.00X30.00in²-(50.80X76.20cm²)-19.355kg-Gloss','S','P','50.8','76.2','20','30','*15004','100','Art Paper','Gloss','19.355','3','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'NOT Exist'"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*00013' OR NAME='Art Paper-130gsm-20.00X30.00in²-(50.80X76.20cm²)-25.161kg-Gloss') Print 'Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15004' AND Code='*16002') Insert Into PaperMaster VALUES ('*00013','Art Paper-130gsm-20.00X30.00in²-(50.80X76.20cm²)-25.161kg-Gloss','Art Paper-130gsm-20.00X30.00in²-(50.80X76.20cm²)-25.161kg-Gloss','S','P','50.8','76.2','20','30','*15004','130','Art Paper','Gloss','25.161','2','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'NOT Exist'"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*00014' OR NAME='Art Paper-170gsm-20.00X30.00in²-(50.80X76.20cm²)-32.903kg-Gloss') Print 'Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15004' AND Code='*16002') Insert Into PaperMaster VALUES ('*00014','Art Paper-170gsm-20.00X30.00in²-(50.80X76.20cm²)-32.903kg-Gloss','Art Paper-170gsm-20.00X30.00in²-(50.80X76.20cm²)-32.903kg-Gloss','S','P','50.8','76.2','20','30','*15004','170','Art Paper','Gloss','32.903','2','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'NOT Exist'"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*00015' OR NAME='Art Paper-70gsm-23.00X36.00in²-(58.42X91.44cm²)-18.697kg-Gloss') Print 'Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15004' AND Code='*16002') Insert Into PaperMaster VALUES ('*00015','Art Paper-70gsm-23.00X36.00in²-(58.42X91.44cm²)-18.697kg-Gloss','Art Paper-70gsm-23.00X36.00in²-(58.42X91.44cm²)-18.697kg-Gloss','S','P','58.42','91.44','23','36','*15004','70','Art Paper','Gloss','18.697','3','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'NOT Exist'"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*00016' OR NAME='Art Paper-80gsm-23.00X36.00in²-(58.42X91.44cm²)-21.368kg-Gloss') Print 'Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15004' AND Code='*16002') Insert Into PaperMaster VALUES ('*00016','Art Paper-80gsm-23.00X36.00in²-(58.42X91.44cm²)-21.368kg-Gloss','Art Paper-80gsm-23.00X36.00in²-(58.42X91.44cm²)-21.368kg-Gloss','S','P','58.42','91.44','23','36','*15004','80','Art Paper','Gloss','21.368','2','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'NOT Exist'"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*00017' OR NAME='Art Paper-90gsm-23.00X36.00in²-(58.42X91.44cm²)-24.039kg-Gloss') Print 'Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15004' AND Code='*16002') Insert Into PaperMaster VALUES ('*00017','Art Paper-90gsm-23.00X36.00in²-(58.42X91.44cm²)-24.039kg-Gloss','Art Paper-90gsm-23.00X36.00in²-(58.42X91.44cm²)-24.039kg-Gloss','S','P','58.42','91.44','23','36','*15004','90','Art Paper','Gloss','24.039','2','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'NOT Exist'"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*00018' OR NAME='Art Paper-100gsm-23.00X36.00in²-(58.42X91.44cm²)-26.71kg-Gloss') Print 'Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15004' AND Code='*16002') Insert Into PaperMaster VALUES ('*00018','Art Paper-100gsm-23.00X36.00in²-(58.42X91.44cm²)-26.71kg-Gloss','Art Paper-100gsm-23.00X36.00in²-(58.42X91.44cm²)-26.71kg-Gloss','S','P','58.42','91.44','23','36','*15004','100','Art Paper','Gloss','26.71','2','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'NOT Exist'"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*00019' OR NAME='Art Paper-130gsm-23.00X36.00in²-(58.42X91.44cm²)-34.723kg-Gloss') Print 'Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15004' AND Code='*16002') Insert Into PaperMaster VALUES ('*00019','Art Paper-130gsm-23.00X36.00in²-(58.42X91.44cm²)-34.723kg-Gloss','Art Paper-130gsm-23.00X36.00in²-(58.42X91.44cm²)-34.723kg-Gloss','S','P','58.42','91.44','23','36','*15004','130','Art Paper','Gloss','34.723','1','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'NOT Exist'"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*00020' OR NAME='Art Paper-170gsm-23.00X36.00in²-(58.42X91.44cm²)-45.406kg-Gloss') Print 'Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15004' AND Code='*16002') Insert Into PaperMaster VALUES ('*00020','Art Paper-170gsm-23.00X36.00in²-(58.42X91.44cm²)-45.406kg-Gloss','Art Paper-170gsm-23.00X36.00in²-(58.42X91.44cm²)-45.406kg-Gloss','S','P','58.42','91.44','23','36','*15004','170','Art Paper','Gloss','45.406','1','64','*16002','0.9','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'NOT Exist'"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*00021' OR NAME='Paper-60gsm-20.00X30.00in²-(50.80X76.20cm²)-11.613kg-Maplitho') Print 'Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15004' AND Code='*16003') Insert Into PaperMaster VALUES ('*00021','Paper-60gsm-20.00X30.00in²-(50.80X76.20cm²)-11.613kg-Maplitho','Paper-60gsm-20.00X30.00in²-(50.80X76.20cm²)-11.613kg-Maplitho','S','P','50.8','76.2','20','30','*15004','60','Paper','Maplitho','11.613','4','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'NOT Exist'"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*00022' OR NAME='Paper-64gsm-20.00X30.00in²-(50.80X76.20cm²)-12.387kg-Maplitho') Print 'Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15004' AND Code='*16003') Insert Into PaperMaster VALUES ('*00022','Paper-64gsm-20.00X30.00in²-(50.80X76.20cm²)-12.387kg-Maplitho','Paper-64gsm-20.00X30.00in²-(50.80X76.20cm²)-12.387kg-Maplitho','S','P','50.8','76.2','20','30','*15004','64','Paper','Maplitho','12.387','4','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'NOT Exist'"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*00023' OR NAME='Paper-70gsm-20.00X30.00in²-(50.80X76.20cm²)-13.548kg-Maplitho') Print 'Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15004' AND Code='*16003') Insert Into PaperMaster VALUES ('*00023','Paper-70gsm-20.00X30.00in²-(50.80X76.20cm²)-13.548kg-Maplitho','Paper-70gsm-20.00X30.00in²-(50.80X76.20cm²)-13.548kg-Maplitho','S','P','50.8','76.2','20','30','*15004','70','Paper','Maplitho','13.548','4','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'NOT Exist'"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*00024' OR NAME='Paper-80gsm-20.00X30.00in²-(50.80X76.20cm²)-15.484kg-Maplitho') Print 'Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15004' AND Code='*16003') Insert Into PaperMaster VALUES ('*00024','Paper-80gsm-20.00X30.00in²-(50.80X76.20cm²)-15.484kg-Maplitho','Paper-80gsm-20.00X30.00in²-(50.80X76.20cm²)-15.484kg-Maplitho','S','P','50.8','76.2','20','30','*15004','80','Paper','Maplitho','15.484','3','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'NOT Exist'"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*00025' OR NAME='Paper-90gsm-20.00X30.00in²-(50.80X76.20cm²)-17.419kg-Maplitho') Print 'Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15004' AND Code='*16003') Insert Into PaperMaster VALUES ('*00025','Paper-90gsm-20.00X30.00in²-(50.80X76.20cm²)-17.419kg-Maplitho','Paper-90gsm-20.00X30.00in²-(50.80X76.20cm²)-17.419kg-Maplitho','S','P','50.8','76.2','20','30','*15004','90','Paper','Maplitho','17.419','3','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'NOT Exist'"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*00026' OR NAME='Paper-100gsm-20.00X30.00in²-(50.80X76.20cm²)-19.355kg-Maplitho') Print 'Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15004' AND Code='*16003') Insert Into PaperMaster VALUES ('*00026','Paper-100gsm-20.00X30.00in²-(50.80X76.20cm²)-19.355kg-Maplitho','Paper-100gsm-20.00X30.00in²-(50.80X76.20cm²)-19.355kg-Maplitho','S','P','50.8','76.2','20','30','*15004','100','Paper','Maplitho','19.355','3','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'NOT Exist'"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*00027' OR NAME='Paper-120gsm-20.00X30.00in²-(50.80X76.20cm²)-23.226kg-Maplitho') Print 'Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15004' AND Code='*16003') Insert Into PaperMaster VALUES ('*00027','Paper-120gsm-20.00X30.00in²-(50.80X76.20cm²)-23.226kg-Maplitho','Paper-120gsm-20.00X30.00in²-(50.80X76.20cm²)-23.226kg-Maplitho','S','P','50.8','76.2','20','30','*15004','120','Paper','Maplitho','23.226','2','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'NOT Exist'"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*00028' OR NAME='Paper-60gsm-23.00X36.00in²-(58.42X91.44cm²)-16.026kg-Maplitho') Print 'Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15004' AND Code='*16003') Insert Into PaperMaster VALUES ('*00028','Paper-60gsm-23.00X36.00in²-(58.42X91.44cm²)-16.026kg-Maplitho','Paper-60gsm-23.00X36.00in²-(58.42X91.44cm²)-16.026kg-Maplitho','S','P','58.42','91.44','23','36','*15004','60','Paper','Maplitho','16.026','3','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'NOT Exist'"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*00029' OR NAME='Paper-64gsm-23.00X36.00in²-(58.42X91.44cm²)-17.094kg-Maplitho') Print 'Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15004' AND Code='*16003') Insert Into PaperMaster VALUES ('*00029','Paper-64gsm-23.00X36.00in²-(58.42X91.44cm²)-17.094kg-Maplitho','Paper-64gsm-23.00X36.00in²-(58.42X91.44cm²)-17.094kg-Maplitho','S','P','58.42','91.44','23','36','*15004','64','Paper','Maplitho','17.094','3','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'NOT Exist'"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*00030' OR NAME='Paper-70gsm-23.00X36.00in²-(58.42X91.44cm²)-18.697kg-Maplitho') Print 'Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15004' AND Code='*16003') Insert Into PaperMaster VALUES ('*00030','Paper-70gsm-23.00X36.00in²-(58.42X91.44cm²)-18.697kg-Maplitho','Paper-70gsm-23.00X36.00in²-(58.42X91.44cm²)-18.697kg-Maplitho','S','P','58.42','91.44','23','36','*15004','70','Paper','Maplitho','18.697','3','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'NOT Exist'"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*00031' OR NAME='Paper-80gsm-23.00X36.00in²-(58.42X91.44cm²)-21.368kg-Maplitho') Print 'Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15004' AND Code='*16003') Insert Into PaperMaster VALUES ('*00031','Paper-80gsm-23.00X36.00in²-(58.42X91.44cm²)-21.368kg-Maplitho','Paper-80gsm-23.00X36.00in²-(58.42X91.44cm²)-21.368kg-Maplitho','S','P','58.42','91.44','23','36','*15004','80','Paper','Maplitho','21.368','2','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'NOT Exist'"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*00032' OR NAME='Paper-90gsm-23.00X36.00in²-(58.42X91.44cm²)-24.039kg-Maplitho') Print 'Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15004' AND Code='*16003') Insert Into PaperMaster VALUES ('*00032','Paper-90gsm-23.00X36.00in²-(58.42X91.44cm²)-24.039kg-Maplitho','Paper-90gsm-23.00X36.00in²-(58.42X91.44cm²)-24.039kg-Maplitho','S','P','58.42','91.44','23','36','*15004','90','Paper','Maplitho','24.039','2','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'NOT Exist'"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*00033' OR NAME='Paper-100gsm-23.00X36.00in²-(58.42X91.44cm²)-26.71kg-Maplitho') Print 'Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15004' AND Code='*16003') Insert Into PaperMaster VALUES ('*00033','Paper-100gsm-23.00X36.00in²-(58.42X91.44cm²)-26.71kg-Maplitho','Paper-100gsm-23.00X36.00in²-(58.42X91.44cm²)-26.71kg-Maplitho','S','P','58.42','91.44','23','36','*15004','100','Paper','Maplitho','26.71','2','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'NOT Exist'"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*00034' OR NAME='Paper-120gsm-23.00X36.00in²-(58.42X91.44cm²)-32.052kg-Maplitho') Print 'Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15004' AND Code='*16003') Insert Into PaperMaster VALUES ('*00034','Paper-120gsm-23.00X36.00in²-(58.42X91.44cm²)-32.052kg-Maplitho','Paper-120gsm-23.00X36.00in²-(58.42X91.44cm²)-32.052kg-Maplitho','S','P','58.42','91.44','23','36','*15004','120','Paper','Maplitho','32.052','2','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'NOT Exist'"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*00035' OR NAME='SBS-200gsm-20.00X30.00in²-(50.80X76.20cm²)-7.742kg-C1S') Print 'Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15002' AND Code='*16003') Insert Into PaperMaster VALUES ('*00035','SBS-200gsm-20.00X30.00in²-(50.80X76.20cm²)-7.742kg-C1S','SBS-200gsm-20.00X30.00in²-(50.80X76.20cm²)-7.742kg-C1S','S','B','50.8','76.2','20','30','*15002','200','SBS','C1S','7.742','6','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'NOT Exist'"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*00036' OR NAME='SBS-210gsm-20.00X30.00in²-(50.80X76.20cm²)-8.129kg-C1S') Print 'Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15002' AND Code='*16003') Insert Into PaperMaster VALUES ('*00036','SBS-210gsm-20.00X30.00in²-(50.80X76.20cm²)-8.129kg-C1S','SBS-210gsm-20.00X30.00in²-(50.80X76.20cm²)-8.129kg-C1S','S','B','50.8','76.2','20','30','*15002','210','SBS','C1S','8.129','6','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'NOT Exist'"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*00037' OR NAME='SBS-220gsm-20.00X30.00in²-(50.80X76.20cm²)-8.516kg-C1S') Print 'Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15002' AND Code='*16003') Insert Into PaperMaster VALUES ('*00037','SBS-220gsm-20.00X30.00in²-(50.80X76.20cm²)-8.516kg-C1S','SBS-220gsm-20.00X30.00in²-(50.80X76.20cm²)-8.516kg-C1S','S','B','50.8','76.2','20','30','*15002','220','SBS','C1S','8.516','6','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'NOT Exist'"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*00038' OR NAME='SBS-250gsm-20.00X30.00in²-(50.80X76.20cm²)-9.677kg-C1S') Print 'Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15002' AND Code='*16003') Insert Into PaperMaster VALUES ('*00038','SBS-250gsm-20.00X30.00in²-(50.80X76.20cm²)-9.677kg-C1S','SBS-250gsm-20.00X30.00in²-(50.80X76.20cm²)-9.677kg-C1S','S','B','50.8','76.2','20','30','*15002','250','SBS','C1S','9.677','5','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'NOT Exist'"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*00039' OR NAME='SBS-200gsm-23.00X36.00in²-(58.42X91.44cm²)-10.684kg-C1S') Print 'Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15002' AND Code='*16003') Insert Into PaperMaster VALUES ('*00039','SBS-200gsm-23.00X36.00in²-(58.42X91.44cm²)-10.684kg-C1S','SBS-200gsm-23.00X36.00in²-(58.42X91.44cm²)-10.684kg-C1S','S','B','58.42','91.44','23','36','*15002','200','SBS','C1S','10.684','5','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'NOT Exist'"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*00040' OR NAME='SBS-210gsm-23.00X36.00in²-(58.42X91.44cm²)-11.218kg-C1S') Print 'Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15002' AND Code='*16003') Insert Into PaperMaster VALUES ('*00040','SBS-210gsm-23.00X36.00in²-(58.42X91.44cm²)-11.218kg-C1S','SBS-210gsm-23.00X36.00in²-(58.42X91.44cm²)-11.218kg-C1S','S','B','58.42','91.44','23','36','*15002','210','SBS','C1S','11.218','4','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'NOT Exist'"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*00041' OR NAME='SBS-220gsm-23.00X36.00in²-(58.42X91.44cm²)-11.752kg-C1S') Print 'Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15002' AND Code='*16003') Insert Into PaperMaster VALUES ('*00041','SBS-220gsm-23.00X36.00in²-(58.42X91.44cm²)-11.752kg-C1S','SBS-220gsm-23.00X36.00in²-(58.42X91.44cm²)-11.752kg-C1S','S','B','58.42','91.44','23','36','*15002','220','SBS','C1S','11.752','4','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'NOT Exist'"
'   cnDatabase.Execute "IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*00042' OR NAME='SBS-250gsm-23.00X36.00in²-(58.42X91.44cm²)-13.355kg-C1S') Print 'Exist' ELSE IF EXISTS (SELECT *FROM GeneralMaster WHERE Code='*15002' AND Code='*16003') Insert Into PaperMaster VALUES ('*00042','SBS-250gsm-23.00X36.00in²-(58.42X91.44cm²)-13.355kg-C1S','SBS-250gsm-23.00X36.00in²-(58.42X91.44cm²)-13.355kg-C1S','S','B','58.42','91.44','23','36','*15002','250','SBS','C1S','13.355','4','64','*16003','1.35','A','000001',GetDate(),'NULL',NULL,'N','N') ELSE Print 'NOT Exist'"
'
'    cnDatabase.CommitTrans
'    Call CloseRecordset(rstCompanyMaster)
'    Screen.MousePointer = vbNormal
'    Exit Function
'ErrorHandler:
'    Update = False
'    cnDatabase.RollbackTrans
'    Call CloseRecordset(rstCompanyMaster)
'    Screen.MousePointer = vbNormal
'End Function
