VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Begin VB.Form FrmSMS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Send SMS"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8250
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SMS.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "FrmLogin"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   8250
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5730
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Cancel"
      Top             =   7785
      Width           =   855
   End
   Begin VB.CommandButton cmdProceed 
      Caption         =   "Send SMS"
      Height          =   375
      Left            =   6690
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Save"
      Top             =   7785
      Width           =   1455
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
      Height          =   7650
      Left            =   120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   105
      Width           =   8010
      _Version        =   65536
      _ExtentX        =   14129
      _ExtentY        =   13494
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
      Picture         =   "SMS.frx":0442
      Begin VB.OptionButton Option3 
         Caption         =   "For Binder"
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
         Left            =   5880
         TabIndex        =   12
         Top             =   120
         Width           =   1335
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
         Left            =   2040
         TabIndex        =   11
         Top             =   5145
         Width           =   5970
      End
      Begin VB.TextBox Text1 
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
         Left            =   2040
         TabIndex        =   10
         Top             =   4830
         Width           =   5970
      End
      Begin VB.OptionButton Option2 
         Caption         =   "For Printer"
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
         Left            =   3120
         TabIndex        =   9
         Top             =   120
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "For General"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.TextBox Text3 
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
         Height          =   1810
         Left            =   0
         TabIndex        =   0
         Top             =   5820
         Width           =   8000
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
         Height          =   340
         Left            =   0
         TabIndex        =   4
         Top             =   5460
         Width           =   8055
         _Version        =   65536
         _ExtentX        =   14208
         _ExtentY        =   600
         _StockProps     =   77
         ForeColor       =   33023
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
         Caption         =   " Type message in below given text box......"
         FillColor       =   4210816
         TextColor       =   16777215
         Picture         =   "SMS.frx":045E
         Picture         =   "SMS.frx":047A
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
         Height          =   330
         Index           =   0
         Left            =   0
         TabIndex        =   5
         Top             =   5145
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
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
         Caption         =   " Subject"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "SMS.frx":0496
         Picture         =   "SMS.frx":04B2
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4365
         Left            =   0
         TabIndex        =   6
         Top             =   480
         Width           =   8010
         _ExtentX        =   14129
         _ExtentY        =   7699
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
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
         NumItems        =   0
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
         Height          =   330
         Index           =   1
         Left            =   0
         TabIndex        =   7
         Top             =   4830
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
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
         Caption         =   " Mobile No."
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "SMS.frx":04CE
         Picture         =   "SMS.frx":04EA
      End
   End
End
Attribute VB_Name = "FrmSMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CxnConnectBusy As New ADODB.Connection
Public rstAccountList As New ADODB.Recordset
Public rstCompanyMaster As New ADODB.Recordset
Public AccountName As String
Dim MobileNo As String
Private Const E_POINTER As Long = &H80004003
Private Const S_OK As Long = 0
Private Const URL_ESCAPE_PERCENT As Long = &H1000&
Private Declare Function UrlEscape Lib "shlwapi" Alias "UrlEscapeA" (ByVal pszURL As String, ByVal pszEscaped As String, ByRef pcchEscaped As Long, ByVal dwFlags As Long) As Long
Private Sub Form_Load()
On Error GoTo ErrorHandler
    Me.Caption = "Send SMS"
    CenterForm Me
    BusySystemIndicator True
    rstCompanyMaster.Open "Select PrintName From CompanyMaster", CxnDatabase, adOpenKeyset, adLockReadOnly
    Option1.Value = True
    ListView1.Enabled = False
    BusySystemIndicator False
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    CloseForm Me
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyReturn Then
        Sendkeys "{TAB}"
        KeyCode = 0
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyS Then
        cmdProceed_Click
        KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyEscape Then
        cmdCancel_Click
        KeyCode = 0
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        Call CloseForm(Me)
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CloseRecordset(rstCompanyMaster)
    Call CloseRecordset(rstAccountList)
    Call CloseConnection(CxnConnectBusy)
End Sub
Private Sub ClearFields()
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
End Sub
Private Sub cmdProceed_Click()
Screen.MousePointer = vbHourglass
Call BusyConnect
On Error GoTo ErrorHandler
Dim rstURL As New ADODB.Recordset
If rstURL.State = adStateOpen Then rstURL.Close
rstURL.Open "SELECT M1 As URL,C1 As UserIDField,C21 As UserIDValue,C2 As PasswordField,C22 As PasswordValue,C5 As SenderIDField,C25 As SenderIDValue,C3 As MobileNoField,C4 As MessageField FROM Config WHERE RecType = 57 AND DocName=(SELECT DocName FROM Despatch..SMSConfig WHERE Company='000002')", CxnConnectBusy, adOpenKeyset, adLockReadOnly
rstURL.ActiveConnection = Nothing
If rstURL.RecordCount = 0 Then Call MsgBox("Failed to send SMS !!!", vbInformation, App.Title): Exit Sub
If Option1.Value = True Then
    If Text3.Text = "" Then Call MsgBox("Failed to send SMS !!!", vbInformation, App.Title): Exit Sub
    If Text1.Text <> "" Then
       Call SendSMS(Trim(Text1.Text), rstURL)
    End If
Else
    Dim z As Integer
    If Text3.Text = "" Then Call MsgBox("Failed to send SMS !!!", vbInformation, App.Title): Exit Sub
    For z = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(z).Checked Then
        MdiMainMenu.StatusBar1.Panels(2).Text = "Sending SMS for Printer/Binder #" & Trim(Text1.Text) & " !!!"
         Text1.Text = ListView1.ListItems.Item(z).SubItems(1)
            If Text1.Text <> "" Then
               Call SendSMS(Trim(Text1.Text), rstURL)
            End If
        End If
    Next z
    MdiMainMenu.StatusBar1.Panels(2).Text = ""
End If
Call CloseRecordset(rstURL)
Call CloseConnection(CxnConnectBusy)
Screen.MousePointer = vbNormal
Exit Sub
ErrorHandler:
    Screen.MousePointer = vbNormal
    DisplayError ("Failed to connect the busy")
    Call CloseConnection(CxnConnectBusy)
    Call CloseRecordset(rstURL)
End Sub
Private Sub cmdCancel_Click()
    Call CloseForm(Me)
End Sub
Private Sub SendSMS(ByVal MobileNo As String, ByVal rstURL As Recordset)
    Dim WinHttpReq As Object
    Dim Response As String, URL As String, VchAmt As Double
    Set WinHttpReq = CreateObject("Msxml2.XMLHTTP")
    Response = MobileNo
    
    URL = rstURL.Fields("URL").Value
    URL = Replace(URL, rstURL.Fields("UserIDField").Value, rstURL.Fields("UserIDValue").Value)
    URL = Replace(URL, rstURL.Fields("PasswordField").Value, rstURL.Fields("PasswordValue").Value)
    URL = Replace(URL, rstURL.Fields("SenderIDField").Value, rstURL.Fields("SenderIDValue").Value)
    URL = Replace(URL, rstURL.Fields("MobileNoField").Value, Response)
    Response = Text3.Text
    URL = Replace(URL, rstURL.Fields("MessageField").Value, Response)
    
    With WinHttpReq
        .Open "GET", URL, False
        .Send
        Response = .responseText
    End With
    
    If Option1.Value = True Then
        If InStr(LCase(Response), "true") > 0 Then
           Call MsgBox("Successfully Send the SMS !!!", vbInformation, App.Title)
        Else
          Call MsgBox("Successfully Send the SMS !!!", vbInformation, App.Title)
        End If
    End If
    
End Sub
Private Function URLEncode(ByVal URL As String) As String
    Dim cchEscaped As Long
    Dim HRESULT As Long
    cchEscaped = Len(URL) * 1.5
    URLEncode = String(cchEscaped, 0)
    HRESULT = UrlEscape(URL, URLEncode, cchEscaped, URL_ESCAPE_PERCENT)
    If HRESULT = E_POINTER Then URLEncode = String$(cchEscaped, 0): HRESULT = UrlEscape(URL, URLEncode, cchEscaped, URL_ESCAPE_PERCENT)
    If HRESULT <> S_OK Then DisplayError ("System error")
    URLEncode = Left$(URLEncode, cchEscaped): URLEncode = Replace$(URLEncode, "+", "%2B"): URLEncode = Replace$(URLEncode, " ", "+")
End Function
Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        For i = 1 To ListView1.ListItems.Count
            ListView1.ListItems(i).Checked = True
        Next i
    ElseIf KeyCode = vbKeyD And Shift = vbCtrlMask Then
        For i = 1 To ListView1.ListItems.Count
            ListView1.ListItems(i).Checked = False
        Next i
    End If
End Sub
Private Sub BusyConnect()
    Dim DatabaseName As String
    Dim i As Integer
    DatabaseName = Trim(ReadFromFile("Busy Database Name")): i = 0
    If ServerName = "" Or DatabaseName = "" Then Exit Sub
    Screen.MousePointer = vbHourglass
    CxnConnectBusy.CursorLocation = adUseClient
      Dim str As String
      i = InStr(1, DatabaseName, ",")
      If CxnConnectBusy.State = adStateOpen Then CxnConnectBusy.Close
      CxnConnectBusy.Open "Provider=SQLOLEDB.1;Password=" & ServerPassword & ";Persist Security Info=True;User ID=sa;Initial Catalog=" & Mid(DatabaseName, i + 1, 20) & ";Data Source=" & ServerName
End Sub
Private Sub Option1_Click()
  For i = 1 To ListView1.ListItems.Count
      ListView1.ListItems(i).Checked = False
  Next i
  ListView1.ListItems.Clear
  ListView1.Enabled = False
  ClearFields
End Sub
Private Sub Option2_Click()
    If rstAccountList.State = adStateOpen Then rstAccountList.Close
    rstAccountList.Open "Select Name,Mobile,Code From AccountMaster Where Type = '05' And Mobile<>'' Order By Name", CxnDatabase, adOpenKeyset, adLockReadOnly
    rstAccountList.ActiveConnection = Nothing
    ListView1.ListItems.Clear
    Call FillList(ListView1, "List of Printer's...", rstAccountList)
    rstAccountList.ActiveConnection = Nothing
    ListView1.Enabled = True
    ClearFields
    Exit Sub
End Sub
Private Sub Option3_Click()
    If rstAccountList.State = adStateOpen Then rstAccountList.Close
    rstAccountList.Open "Select Name,Mobile,Code From AccountMaster Where Type = '08' And Mobile<>'' Order By Name", CxnDatabase, adOpenKeyset, adLockReadOnly
    rstAccountList.ActiveConnection = Nothing
    ListView1.ListItems.Clear
    Call FillList(ListView1, "List of Binder's...", rstAccountList)
    rstAccountList.ActiveConnection = Nothing
    ListView1.Enabled = True
    ClearFields
    Exit Sub
End Sub
