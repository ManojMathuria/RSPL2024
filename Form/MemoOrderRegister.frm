VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmMemoOrderRegister 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Memo Order Register"
   ClientHeight    =   6435
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   8400
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MemoOrderRegister.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6435
   ScaleWidth      =   8400
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   8400
      _ExtentX        =   14817
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Print Preview"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Print"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exit"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3000
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MemoOrderRegister.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MemoOrderRegister.frx":0986
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MemoOrderRegister.frx":0A9A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   6060
      Left            =   45
      TabIndex        =   11
      Top             =   345
      Width           =   8310
      _Version        =   65536
      _ExtentX        =   14658
      _ExtentY        =   10689
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
      Picture         =   "MemoOrderRegister.frx":0BAE
      Begin VB.CheckBox Check3 
         Caption         =   "Repair"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5700
         TabIndex        =   4
         Top             =   53
         Value           =   1  'Checked
         Width           =   840
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Fresh"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4920
         TabIndex        =   3
         Top             =   53
         Value           =   1  'Checked
         Width           =   750
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Show All"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3840
         TabIndex        =   2
         Top             =   53
         Width           =   1095
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2875
         Left            =   0
         TabIndex        =   6
         Top             =   320
         Width           =   4150
         _ExtentX        =   7329
         _ExtentY        =   5080
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
         Height          =   330
         Left            =   0
         TabIndex        =   12
         Top             =   0
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
         Caption         =   " &From"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   -2147483634
         Picture         =   "MemoOrderRegister.frx":0BCA
         Picture         =   "MemoOrderRegister.frx":0BE6
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
         Height          =   330
         Left            =   1920
         TabIndex        =   13
         Top             =   0
         Width           =   765
         _Version        =   65536
         _ExtentX        =   1349
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
         Caption         =   " &To"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   -2147483634
         Picture         =   "MemoOrderRegister.frx":0C02
         Picture         =   "MemoOrderRegister.frx":0C1E
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   2880
         Left            =   4140
         TabIndex        =   7
         Top             =   315
         Width           =   4170
         _ExtentX        =   7355
         _ExtentY        =   5080
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin TDBDate6Ctl.TDBDate MhDateInput2 
         Height          =   330
         Left            =   2670
         TabIndex        =   1
         Top             =   0
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calendar        =   "MemoOrderRegister.frx":0C3A
         Caption         =   "MemoOrderRegister.frx":0D52
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "MemoOrderRegister.frx":0DBE
         Keys            =   "MemoOrderRegister.frx":0DDC
         Spin            =   "MemoOrderRegister.frx":0E3A
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
      Begin TDBDate6Ctl.TDBDate MhDateInput1 
         Height          =   330
         Left            =   840
         TabIndex        =   0
         Top             =   0
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calendar        =   "MemoOrderRegister.frx":0E62
         Caption         =   "MemoOrderRegister.frx":0F7A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "MemoOrderRegister.frx":0FE6
         Keys            =   "MemoOrderRegister.frx":1004
         Spin            =   "MemoOrderRegister.frx":1062
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
      Begin MSComctlLib.ListView ListView3 
         Height          =   2880
         Left            =   0
         TabIndex        =   8
         Top             =   3180
         Width           =   4150
         _ExtentX        =   7329
         _ExtentY        =   5080
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView4 
         Height          =   2880
         Left            =   4140
         TabIndex        =   9
         Top             =   3180
         Width           =   4170
         _ExtentX        =   7355
         _ExtentY        =   5080
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSForms.ComboBox Combo1 
         Height          =   330
         Left            =   6580
         TabIndex        =   5
         Top             =   0
         Width           =   1725
         VariousPropertyBits=   545282075
         BackColor       =   16777215
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "3043;582"
         MatchEntry      =   0
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Calibri"
         FontEffects     =   1073741825
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
   End
End
Attribute VB_Name = "FrmMemoOrderRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CxnBusy As New ADODB.Connection
Dim rstPrintOrderStatusRegister As New ADODB.Recordset, rstCompanyMaster As New ADODB.Recordset, rstBusy As New ADODB.Recordset, rstBoardList As New ADODB.Recordset, rstBookList As New ADODB.Recordset, rstAccountList As New ADODB.Recordset
Dim rstExist As New ADODB.Recordset
Dim rstBreakup As New ADODB.Recordset
Dim OutputTo As String
Dim RefCode As String
Public OrderType As String
Private Sub Form_Load()
    
    '01-Bookwise 02-Print Orderwise 05
    On Error GoTo ErrorHandler
    CenterForm Me
    BusySystemIndicator True
    CxnBusy.CursorLocation = adUseClient
    
    If InStr(1, "0102", OrderType) = 0 Then ListView3.Width = 8310: ListView4.Visible = False
    If OrderType = "XX" Then
       Me.Height = 3940: Mh3dFrame1.Height = 3200: ListView3.Visible = False
    Else
        rstAccountList.Open "SELECT Name,Code FROM AccountMaster WHERE Type IN ('" & IIf(InStr(1, "0102", OrderType) > 0, "05", IIf(OrderType = "YY", "08", OrderType)) & "') ORDER BY Name", CxnDatabase, adOpenKeyset, adLockReadOnly
        rstAccountList.ActiveConnection = Nothing
        Call FillList(ListView3, "List of " & IIf(InStr(1, "01020506", OrderType) > 0, "Printers", "Binders") & "...", rstAccountList)
        If InStr(1, "0102", OrderType) > 0 Then
            If rstAccountList.State = adStateOpen Then rstAccountList.Close
            rstAccountList.Open "SELECT Name,Code FROM AccountMaster WHERE Type='08' ORDER BY Name", CxnDatabase, adOpenKeyset, adLockReadOnly
            rstAccountList.ActiveConnection = Nothing
            Call FillList(ListView4, "List of Book Binders...", rstAccountList)
        End If
    End If
                  
    rstCompanyMaster.Open "SELECT PrintName FROM CompanyMaster", CxnDatabase, adOpenKeyset, adLockReadOnly
    rstBoardList.Open "SELECT Name,Code FROM GeneralMaster WHERE Type='2' ORDER BY Name", CxnDatabase, adOpenKeyset, adLockReadOnly
    rstBoardList.ActiveConnection = Nothing
    Call FillList(ListView1, "List of Boards...", rstBoardList)
    Call BookSelection(True)
    'MhDateInput1.Text = Format(DateAdd("D", -365, FinancialYearFrom), "dd-mm-yyyy")
    MhDateInput1.Text = Format(FinancialYearFrom, "dd-mm-yyyy")
    MhDateInput2.Text = Format(Date, "dd-mm-yyyy")
    Combo1.AddItem "Without Stock", 0: Combo1.AddItem "With Stock", 1: Combo1.AddItem "With Pending SO", 2: Combo1.ListIndex = 0
    BusySystemIndicator False
    Exit Sub
    
ErrorHandler:
    BusySystemIndicator False
    CloseForm Me
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       SendKeys "{TAB}", True: KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyEscape Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(3): KeyCode = 0
    ElseIf Shift = vbAltMask And KeyCode = vbKeyP Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(2): KeyCode = 0
    ElseIf Shift = vbAltMask And KeyCode = vbKeyV Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(1): KeyCode = 0
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then CloseForm Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CloseRecordset(rstCompanyMaster)
    Call CloseRecordset(rstBookList)
    Call CloseRecordset(rstBoardList)
    Call CloseRecordset(rstAccountList)
    Call CloseRecordset(rstBusy)
    Call CloseRecordset(rstPrintOrderStatusRegister)
    Call CloseRecordset(rstExist)
    Call CloseRecordset(rstBreakup)
End Sub
Private Sub MhDateInput1_Validate(Cancel As Boolean)
    If Not IsDate(GetDate(MhDateInput1.Text)) Then Cancel = True
End Sub
Private Sub MhDateInput2_Validate(Cancel As Boolean)
    If Not IsDate(GetDate(MhDateInput2.Text)) Or Format(GetDate(MhDateInput2.Text), "yyyymmdd") < Format(GetDate(MhDateInput1.Text), "yyyymmdd") Then Cancel = True
End Sub
Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
     Call BookSelection(False)
End Sub
Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    If Shift = vbCtrlMask And (KeyCode = vbKeyA Or KeyCode = vbKeyD) Then
        For i = 1 To ListView1.ListItems.Count
            ListView1.ListItems(i).Checked = IIf(KeyCode = vbKeyA, True, False)
        Next i
        Call BookSelection(IIf(KeyCode = vbKeyA, True, False))
    End If
End Sub
Private Sub ListView2_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    If Shift = vbCtrlMask And (KeyCode = vbKeyA Or KeyCode = vbKeyD) Then
        For i = 1 To ListView2.ListItems.Count
            ListView2.ListItems(i).Checked = IIf(KeyCode = vbKeyA, True, False)
        Next i
    End If
End Sub
Private Sub ListView3_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    If Shift = vbCtrlMask And (KeyCode = vbKeyA Or KeyCode = vbKeyD) Then
        For i = 1 To ListView3.ListItems.Count
            ListView3.ListItems(i).Checked = IIf(KeyCode = vbKeyA, True, False)
        Next i
    End If
End Sub
Private Sub ListView4_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    If Shift = vbCtrlMask And (KeyCode = vbKeyA Or KeyCode = vbKeyD) Then
        For i = 1 To ListView4.ListItems.Count
            ListView4.ListItems(i).Checked = IIf(KeyCode = vbKeyA, True, False)
        Next i
    End If
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    If Button.Index = 3 Then CloseForm Me: Exit Sub
    If Button.Index = 1 Then OutputTo = "S" Else OutputTo = "P"
    PrintOrderStatusRegister
End Sub
Private Sub BookSelection(ByVal SelectAll As Boolean)
    If rstBookList.State = adStateOpen Then rstBookList.Close
    rstBookList.Open "SELECT Name,Code FROM BookMaster " & IIf(SelectAll, "", "WHERE Board IN (" & SelectedItems(ListView1) & ")") & " ORDER BY Name", CxnDatabase, adOpenKeyset, adLockReadOnly
    rstBookList.ActiveConnection = Nothing
    ListView2.ListItems.Clear
    Call FillList(ListView2, "List of Books...", rstBookList)
End Sub
Private Function GetBusyOrder(ByVal xOrderNo As String)
    If rstBusy.State = adStateOpen Then rstBusy.Close
    rstBusy.Open "SELECT T1.VchNo,M.Alias,M.PrintName As BinderName,T1.Date,ABS(T3.Value1) As OrderedQuantity FROM (Tran1 T1 INNER JOIN Tran3 T3 ON T1.VchNo=T3.No) INNER JOIN Master1 M ON M.Code=T1.MasterCode1 WHERE T1.VchType=T3.VchType AND T1.VchType=13 AND T1.Date=T3.Date AND LTRIM(T3.No)='" & Trim(xOrderNo) & "'", CxnBusy, adOpenKeyset, adLockReadOnly
    rstBusy.ActiveConnection = Nothing
End Function
Private Function ConnectToBusy() As Boolean
    On Error GoTo ErrHandler
    Dim DatabaseName
    DatabaseName = Trim(ReadFromFile("Busy Database Name"))
    DatabaseName = StrReverse(Left(StrReverse(DatabaseName), InStr(1, StrReverse(DatabaseName), ",") - 1))
    If CxnBusy.State = adStateOpen Then CxnBusy.Close
    CxnBusy.Open "Provider=SQLOLEDB.1;Password=" & ServerPassword & ";Persist Security Info=True;User ID=sa;Initial Catalog=" & DatabaseName & ";Data Source=" & ServerName
    ConnectToBusy = True
ErrHandler:
End Function
Private Sub GetAllItemStock()
    Dim SQL As String
    MdiMainMenu.StatusBar1.Panels(2).Text = "Processing !!! Please Wait....."
    On Error GoTo ErrorHandler
    Screen.MousePointer = vbHourglass
    If rstBusy.State = adStateOpen Then rstBusy.Close
    SQL = "SELECT LEFT(Alias,4) As Alias,Alias As FullAlias," & _
              "(SELECT ISNULL(SUM(D1),0) FROM Tran4 WHERE MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND ParentGrp IN (SELECT Code FROM Master1 WHERE MasterType=10 AND UPPER(Name) LIKE '%" & UCase(MCGroup) & "%'))) As OpBal," & _
              "(SELECT ISNULL(SUM(0-Value1),0) FROM Tran2 WHERE VchType IN (3,9) AND RecType=2 AND Date>='" & FinancialYearFrom & "' AND Date <='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND ParentGrp IN (SELECT Code FROM Master1 WHERE MasterType=10 AND UPPER(Name) LIKE '%" & UCase(MCGroup) & "%'))) As NetSale," & _
              "(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE VchType=5 AND Date>='" & FinancialYearFrom & "' AND Date <='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND ParentGrp IN (SELECT Code FROM Master1 WHERE MasterType=10 AND UPPER(Name) LIKE '%" & UCase(MCGroup) & "%'))) As NetStockTransfer," & _
              "(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE VchType IN (4,11) AND Date>='" & FinancialYearFrom & "' AND Date <='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND ParentGrp IN (SELECT Code FROM Master1 WHERE MasterType=10 AND UPPER(Name) LIKE '%" & UCase(MCGroup) & "%'))) As NetPurchase," & _
              "(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE VchType=8 AND Date>='" & FinancialYearFrom & "' AND Date <='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND ParentGrp IN (SELECT Code FROM Master1 WHERE MasterType=10 AND UPPER(Name) LIKE '%" & UCase(MCGroup) & "%'))) As NetStockAdjustment "
    If Combo1.ListIndex = 1 Then
        SQL = SQL + "FROM Master1 M WHERE MasterType=6 AND Alias<>'' ORDER BY Left(Alias,4)"
    Else
        SQL = SQL + "," & _
                   "(SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran3 WHERE VchType=12 AND Date<='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code) As SaleOrder," & _
                   "ABS(ISNULL((SELECT ISNULL(SUM(Value1),0) FROM Tran3 WHERE RecType=4 AND Method=2 AND RefCode IN (SELECT RefCode FROM Tran3 WHERE VchType=12 AND Date<='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code)),0)) As SaleOrderSupplied " & _
                   "FROM Master1 M WHERE MasterType=6 AND Alias<>'' ORDER BY Left(Alias,4)"
    End If
    rstBusy.Open SQL, CxnBusy, adOpenKeyset, adLockReadOnly
    rstBusy.ActiveConnection = Nothing
ErrorHandler:
    On Error GoTo 0
    Screen.MousePointer = vbNormal
End Sub
Private Function GetStock(ByVal xItem As String) As String
    Dim EffStock As Long, PendingSO As Long
    On Error GoTo ErrorHandler
    Screen.MousePointer = vbHourglass
    If rstBusy.RecordCount > 0 Then rstBusy.MoveFirst
    rstBusy.Find "[Alias]='" & Left(xItem, 4) & "'"
    Do While Not rstBusy.EOF
        If rstBusy.Fields("Alias") = Left(xItem, 4) Then
            If InStr(1, "Z-Z_", Mid(rstBusy.Fields("FullAlias").Value, 6, 2), vbTextCompare) = 0 Then
                EffStock = EffStock + Val(rstBusy.Fields("OpBal").Value) - Val(rstBusy.Fields("NetSale").Value) + Val(rstBusy.Fields("NetStockTransfer").Value) + Val(rstBusy.Fields("NetPurchase").Value) + Val(rstBusy.Fields("NetStockAdjustment").Value)
                If Combo1.ListIndex = 2 Then PendingSO = PendingSO + Val(rstBusy.Fields("SaleOrder").Value) - Val(rstBusy.Fields("SaleOrderSupplied").Value)
            End If
        Else
            Exit Do
        End If
        rstBusy.MoveNext
    Loop
    GetStock = Trim(str(EffStock)) + "|" + Trim(str(PendingSO))
ErrorHandler:
    On Error GoTo 0
    Screen.MousePointer = vbNormal
End Function
Private Sub PrintOrderStatusRegister()
    Dim oExcel As Object
    Dim i As Long, K As Long, Cnt As Long, T As String
    Dim balQty As Double
    Dim SelectedBoards, SelectedBooks, SelectedPrinters, SelectedBinders, SQL, Path
    If OrderType = "XX" Or (Combo1.ListIndex > 0 And InStr(1, "0102050608", OrderType) > 0) Then If Not ConnectToBusy Then Screen.MousePointer = vbNormal: DisplayError ("Failed to connect to busy"): Exit Sub
    On Error Resume Next
    ConnectToBusy
    If Not FileExist(App.Path & "\Template\Memo Order Register.xlsx") Then Exit Sub
    Screen.MousePointer = vbHourglass
    If rstPrintOrderStatusRegister.State = adStateOpen Then rstPrintOrderStatusRegister.Close
    SelectedBoards = SelectedItems(ListView1): SelectedBooks = SelectedItems(ListView2)
    If OrderType <> "XX" Then
        If Combo1.ListIndex > 0 And InStr(1, "0102050608", OrderType) > 0 Then GetAllItemStock
        
        SQL = "SELECT P.Code,P.Name As OrderNo,P.Date As OrderDate,M2.PrintName As BookName,M2.BusyCode As Alias,M3.PrintName As BoardName,(SELECT PrintName FROM GeneralMaster WHERE Code=M2.[Size]) As BookSize,M2.FormType,(SELECT STR(Forms1) FROM BookPOChild05 WHERE Code=P.Code) As Forms1,(SELECT STR(Forms2) FROM BookPOChild05 WHERE Code=P.Code) As Forms2,(SELECT STR(Forms4) FROM BookPOChild05 WHERE Code=P.Code) As Forms4,(SELECT STR(FrontPrintingType) FROM BookPOChild06 WHERE Code=P.Code) As FrontPrintingType,(SELECT STR(BackPrintingType) FROM BookPOChild06 WHERE Code=P.Code) As BackPrintingType,FORMAT((SELECT ActualQuantity FROM BookPOChild05 WHERE Code=P.Code),'0') As TextQuantity,FORMAT((SELECT ActualQuantity FROM BookPOChild06 WHERE Code=P.Code),'0') As TitleQuantity,FORMAT((SELECT ActualQuantity FROM BookPOChild08 WHERE Code=P.Code),'0') As BookQuantity,ReceivedQuantity," & _
                  "(SELECT PrintName FROM AccountMaster WHERE Code=P.BookPrinter) As TextPrinterName,(SELECT Status FROM BookPOChild05 WHERE Code=P.Code) As TextStatus,(SELECT PrintName FROM AccountMaster WHERE Code=P.TitlePrinter) As TitlePrinterName,(SELECT Status FROM BookPOChild06 WHERE Code=P.Code) As TitleStatus,(SELECT PrintName FROM AccountMaster WHERE Code=P.Binder) As BinderName,(SELECT Status FROM BookPOChild08 WHERE Code=P.Code) As BookStatus,(SELECT PrintName FROM BookPOChild05 T INNER JOIN PaperMaster M ON T.Paper1=M.Code WHERE T.Code=P.Code) As Paper1,(SELECT PrintName FROM BookPOChild05 T INNER JOIN PaperMaster M ON T.Paper2=M.Code WHERE T.Code=P.Code) As Paper2,(SELECT PrintName FROM BookPOChild05 T INNER JOIN PaperMaster M ON T.Paper4=M.Code WHERE T.Code=P.Code) As Paper4,C.Narration,C.BillNo,C.BillDate,C.TargetDate,C.ExtendDate "
        
    End If
    If InStr(1, "0102", OrderType) > 0 Then 'Book/Print Orderwise
        SelectedPrinters = SelectedItems(ListView3): SelectedBinders = SelectedItems(ListView4)
                        
        SQL = SQL + ",(SELECT Name  FROM PrintPVParent Where Code in( Select Ref From BookPOChild05 WHERE Code=P.Code) ) As Ref1,(SELECT Name  FROM PrintPVParent Where Code in( Select Ref From BookPOChild06 WHERE Code=P.Code) ) As Ref2,(SELECT Narration  FROM PrintPVChild Where Code in( Select Ref From BookPOChild05 WHERE Code=P.Code And Book=M2.Code) ) As RefRemarks,(SELECT Warehouse1  FROM PrintPVChild Where Code in( Select Ref From BookPOChild05 WHERE Code=P.Code And Book=M2.Code) ) As RefNoida,(SELECT Warehouse2  FROM PrintPVChild Where Code in( Select Ref From BookPOChild05 WHERE Code=P.Code  And Book=M2.Code) ) As RefDaryaganj,(SELECT Warehouse3  FROM PrintPVChild Where Code in( Select Ref From BookPOChild05 WHERE Code=P.Code And Book=M2.Code) ) As Ref8No,(SELECT  BookStatus  FROM BookPOChild05 WHERE Code=P.Code) As BookStatus2,(Select AdvanceRecvdDate From BookPOChild08 WHERE Code=P.Code ) As AdvanceRecvdDate" & _
                             " FROM (((BookPOParent P LEFT JOIN BookPOChild08 C ON P.Code=C.Code) LEFT JOIN AccountMaster M1 ON P.Binder=M1.Code) LEFT JOIN BookMaster M2 ON P.Book=M2.Code) LEFT JOIN GeneralMaster M3 ON M2.Board=M3.Code " & _
                             " WHERE P.Type IN ('" & IIf(Check2.Value And Check3.Value, "F','R", IIf(Check2.Value, "F", "R")) & "') AND P.Date>=#" & GetDate(MhDateInput1.Text) & "# AND P.Date<=#" & GetDate(MhDateInput2.Text) & "# AND " & IIf(Check1.Value, "1=1", "C.Status NOT IN ('D','E','W')") & " AND " & IIf(SelectedBoards = "''", "1=1", "M3.Code IN (" & SelectedBoards & ")") & " AND " & IIf(SelectedBooks = "''", "1=1", "M2.Code IN (" & SelectedBooks & ")") & " AND " & IIf(SelectedPrinters = "''", "1=1", "P.BookPrinter IN (" & SelectedPrinters & ")") & " AND " & IIf(SelectedBinders = "''", "1=1", "M1.Code IN (" & SelectedBinders & ")") & Space(1) & _
                             "ORDER BY " & IIf(OrderType = "01", "M2.PrintName,Val(P.Name)", "Val(P.Name),M2.PrintName")
        rstPrintOrderStatusRegister.Open SQL, CxnDatabase, adOpenKeyset, adLockReadOnly
    
    End If
    If rstPrintOrderStatusRegister.RecordCount = 0 Then Screen.MousePointer = vbNormal: On Error GoTo 0: Exit Sub
    DoEvents
    Set oExcel = CreateObject("Excel.Application")
    oExcel.Workbooks.Open (App.Path & "\Template\Memo Order Register")
    oExcel.DisplayAlerts = False
    Path = "Memo Order Register"
    oExcel.Workbooks.Item(1).SaveAs (App.Path & "\Report\" & Path & " (" & CompCode & ")")
    oExcel.DisplayAlerts = True
    If OrderType <> "XX" Then
        For i = 1 To oExcel.Sheets.Count
            If InStr(1, "Sheet1", oExcel.Sheets(i).Name) = 0 Then oExcel.Sheets(i).Visible = False
        Next
        oExcel.Visible = False
        oExcel.Cells(1, "A").Value = Trim(rstCompanyMaster.Fields("PrintName").Value)
        oExcel.Cells(2, "A").Value = "Memo Order Register" & " From " & Format(MhDateInput1, "dd-MMM-yyyy") & " To " & Format(MhDateInput2, "dd-MMM-yyyy")
        i = 4: Cnt = 1
        Do While Not rstPrintOrderStatusRegister.EOF
            If OrderType = "YY" Then If CheckEmpty(rstPrintOrderStatusRegister.Fields("BillNo").Value, False) Or Left(Trim(rstPrintOrderStatusRegister.Fields("OrderNo").Value), 1) = "*" Then GoTo Continue
            oExcel.Cells(i, "A").Value = Cnt
            
            oExcel.Application.Cells(i, "B").Value = Trim(rstPrintOrderStatusRegister.Fields("Ref1").Value)
            
            oExcel.Application.Cells(i, "C").Value = Trim(rstPrintOrderStatusRegister.Fields("BookName").Value)
           
            If Trim(rstPrintOrderStatusRegister.Fields("TextPrinterName").Value) <> "" Then
               oExcel.Application.Cells(i, "D").Value = "Reprint"
             Else
               oExcel.Application.Cells(i, "D").Value = "New"
            End If
            
            oExcel.Application.Cells(i, "E").Value = Trim(rstPrintOrderStatusRegister.Fields("BookSize").Value) & "/" & IIf(rstPrintOrderStatusRegister.Fields("FormType").Value = "1", "08", IIf(rstPrintOrderStatusRegister.Fields("FormType").Value = "2", "16", IIf(rstPrintOrderStatusRegister.Fields("FormType").Value = "3", "04", IIf(rstPrintOrderStatusRegister.Fields("FormType").Value = "4", "12", IIf(rstPrintOrderStatusRegister.Fields("FormType").Value = "5", "24", IIf(rstPrintOrderStatusRegister.Fields("FormType").Value = "6", "32", "64"))))))
            
            oExcel.Application.Cells(i, "F").Value = IIf(Left(Trim(rstPrintOrderStatusRegister.Fields("OrderNo").Value), 1) = "*", Mid(Trim(rstPrintOrderStatusRegister.Fields("OrderNo").Value), 2), Trim(rstPrintOrderStatusRegister.Fields("OrderNo").Value))
            oExcel.Application.Cells(i, "G").Value = Format(rstPrintOrderStatusRegister.Fields("OrderDate").Value, "dd-MM-yy")
            oExcel.Application.Cells(i, "H").Value = rstPrintOrderStatusRegister.Fields(IIf(OrderType = "06", "TitleQuantity", IIf(OrderType = "05", "TextQuantity", "BookQuantity"))).Value
            oExcel.Application.Cells(i, "I").Value = Trim(rstPrintOrderStatusRegister.Fields("TextPrinterName").Value)
            oExcel.Application.Cells(i, "J").Value = Trim(rstPrintOrderStatusRegister.Fields("TitlePrinterName").Value)
            oExcel.Application.Cells(i, "K").Value = Trim(rstPrintOrderStatusRegister.Fields("BinderName").Value)
                        
            oExcel.Application.Cells(i, "L").Value = Trim(rstPrintOrderStatusRegister.Fields("BookStatus2").Value)
            
            If Trim(rstPrintOrderStatusRegister.Fields("BookStatus2").Value) = "N/C" Or Trim(rstPrintOrderStatusRegister.Fields("BookStatus2").Value) = "Pendind" Then
               oExcel.Application.Cells(i, "M").Value = "No"
            ElseIf Trim(rstPrintOrderStatusRegister.Fields("BookStatus2").Value) = "New" Or Trim(rstPrintOrderStatusRegister.Fields("BookStatus2").Value) = "L/C" Then
               oExcel.Application.Cells(i, "M").Value = "Yes"
            Else
               oExcel.Application.Cells(i, "M").Value = ""
            End If
            
            oExcel.Application.Cells(i, "N").Value = Trim(rstPrintOrderStatusRegister.Fields("RefNoida").Value)
            oExcel.Application.Cells(i, "O").Value = Trim(rstPrintOrderStatusRegister.Fields("RefDaryaganj").Value)
            oExcel.Application.Cells(i, "P").Value = Trim(rstPrintOrderStatusRegister.Fields("Ref8No").Value)
            oExcel.Application.Cells(i, "Q").Value = Format(rstPrintOrderStatusRegister.Fields("AdvanceRecvdDate").Value, "dd-MM-yyyy")
            
            oExcel.Application.Cells(i, "R").Value = Trim(rstPrintOrderStatusRegister.Fields("RefRemarks").Value)


            
            MdiMainMenu.StatusBar1.Panels(2).Text = "Processed record #" & Trim(str(Cnt)) & " of " & Trim(str(rstPrintOrderStatusRegister.RecordCount)) & " !!!"
            
            Cnt = Cnt + 1: i = i + 1
Continue:
            rstPrintOrderStatusRegister.MoveNext
        Loop
        MdiMainMenu.StatusBar1.Panels(2).Text = ""
        oExcel.Columns("A:AD").EntireColumn.AutoFit
       
'        oExcel.Columns("T:T").Hidden = True
'        oExcel.Columns("AC:AD").Hidden = True: oExcel.Columns("H").Hidden = True: oExcel.Columns("C").Hidden = True: oExcel.Columns("Y:AA").Hidden = True
'        If OrderType = "06" Then oExcel.Columns("H").Hidden = False
'        If OrderType = "YY" Then oExcel.Columns("O").Hidden = True: oExcel.Columns("Q").Hidden = True: oExcel.Columns("U:AB").Hidden = True: oExcel.Columns("AC:AD").Hidden = False
    End If
    oExcel.Workbooks.Item(1).Save
    Screen.MousePointer = vbNormal
    If OutputTo = "S" Then oExcel.Range("A1").Activate: oExcel.Visible = True Else oExcel.Workbooks.Item(1).PrintOut
    Set oExcel = Nothing
    Call CloseConnection(CxnBusy)
    On Error GoTo 0
End Sub
Private Function GetBreakupRecievedQty(ByVal xOrderNo As String, ByVal xRefCode As String) As String
 On Error GoTo ErrHandler
    Dim sBreakup As String
    Dim rQty As Double
    If rstBreakup.State = adStateOpen Then rstBreakup.Close
    rstBreakup.Open "SELECT Date,Value1 FROM Tran3 WHERE Tran3.RefCode =" & xRefCode & " And Tran3.RecType=4 and Tran3.VchType =4 ORDER BY Tran3.Type", CxnBusy, adOpenKeyset, adLockReadOnly
    rstBreakup.ActiveConnection = Nothing
    If rstBreakup.RecordCount > 0 Then
        Do While Not rstBreakup.EOF
           sBreakup = sBreakup & vbLf & Format(rstBreakup.Fields("Date").Value, "dd-MMM-yy") & "(" & rstBreakup.Fields("Value1").Value & ")"
           rQty = rQty + Val(rstBreakup.Fields("Value1").Value)
           rstBreakup.MoveNext
        Loop
    Else
      sBreakup = "0"
    End If
    
    If rQty > 0 Then
       GetBreakupRecievedQty = "Total = " & rQty & vbLf & sBreakup
    Else
       GetBreakupRecievedQty = sBreakup
    End If
    Set rstBreakup = Nothing
    
ErrHandler:

End Function
Private Function BreakupExist(ByVal xOrderNo As String, ByVal oQty As Double) As Boolean
  On Error GoTo ErrHandler
    RefCode = ""
    If rstExist.State = adStateOpen Then rstExist.Close
    rstExist.Open "SELECT T3.RefCode,ABS(T3.Value1) AS OrderedQuantity ,Ltrim(T1.VchNo) as BillNo,T1.VchSeriesCode ,(Select Top 1 NameAlias from Help1 as H1 where H1.NameOrAlias = 1 and H1.Code = T3.MasterCode2 ) as AccName,(Select Top 1 NameAlias from Help1 as H1 where H1.NameOrAlias = 1 and H1.Code = T3.MasterCode1) as ItemName, T3.ItemSrNo as ItemSrNo, (Select Name from Master1 where Code = T3.MasterCode1) as tmpItemName  FROM TRAN3 As T3 INNER JOIN TRAN1 As T1 ON T3.VCHCODE = T1.VCHCODE WHERE T3.RecType=4 And T3.Method=1 And T3.VchType = 13 And T3.Date>='" & GetDate(MhDateInput1.Text) & "' AND T3.Date <='" & GetDate(MhDateInput2.Text) & "' And T3.ApprovalStatus <> 2 AND T1.VchSeriesCode = 262 AND LTRIM(T1.Vchno)= '" & xOrderNo & "' Order By T3.Date, T3.Vchtype, t1.VchNo,T3.MasterCode2,T3.MasterCode1,T3.RefCode", CxnBusy, adOpenKeyset, adLockReadOnly
    rstExist.ActiveConnection = Nothing
    If rstExist.RecordCount > 0 Then
       If Val(rstExist.Fields("OrderedQuantity").Value) = Val(oQty) Then
          RefCode = rstExist.Fields("RefCode").Value
          BreakupExist = True
       Else
          RefCode = ""
          BreakupExist = False
       End If
    Else
       RefCode = ""
       BreakupExist = False
    End If
    Set rstExist = Nothing
ErrHandler:
End Function



