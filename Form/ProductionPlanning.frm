VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmProductionPlanning 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Production Planning"
   ClientHeight    =   6435
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7620
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ProductionPlanning.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6435
   ScaleWidth      =   7620
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   7620
      _ExtentX        =   13441
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
            Picture         =   "ProductionPlanning.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ProductionPlanning.frx":0986
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ProductionPlanning.frx":0A9A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   6065
      Left            =   45
      TabIndex        =   7
      Top             =   345
      Width           =   7530
      _Version        =   65536
      _ExtentX        =   13282
      _ExtentY        =   10698
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
      Picture         =   "ProductionPlanning.frx":0BAE
      Begin MSComctlLib.ListView ListView4 
         Height          =   2875
         Left            =   0
         TabIndex        =   2
         Top             =   320
         Width           =   3765
         _ExtentX        =   6641
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
         Height          =   330
         Index           =   0
         Left            =   0
         TabIndex        =   8
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
         Picture         =   "ProductionPlanning.frx":0BCA
         Picture         =   "ProductionPlanning.frx":0BE6
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
         Height          =   330
         Left            =   1920
         TabIndex        =   9
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
         Picture         =   "ProductionPlanning.frx":0C02
         Picture         =   "ProductionPlanning.frx":0C1E
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2875
         Left            =   3750
         TabIndex        =   3
         Top             =   320
         Width           =   3780
         _ExtentX        =   6668
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
            Weight          =   400
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
         Calendar        =   "ProductionPlanning.frx":0C3A
         Caption         =   "ProductionPlanning.frx":0D52
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "ProductionPlanning.frx":0DBE
         Keys            =   "ProductionPlanning.frx":0DDC
         Spin            =   "ProductionPlanning.frx":0E3A
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
         Calendar        =   "ProductionPlanning.frx":0E62
         Caption         =   "ProductionPlanning.frx":0F7A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "ProductionPlanning.frx":0FE6
         Keys            =   "ProductionPlanning.frx":1004
         Spin            =   "ProductionPlanning.frx":1062
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
      Begin MSComctlLib.ListView ListView2 
         Height          =   2880
         Left            =   0
         TabIndex        =   4
         Top             =   3180
         Width           =   3765
         _ExtentX        =   6641
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   2880
         Left            =   3750
         TabIndex        =   5
         Top             =   3180
         Width           =   3765
         _ExtentX        =   6641
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
         Left            =   3720
         TabIndex        =   10
         Top             =   0
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
         Caption         =   "  Alias Prefix"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   -2147483634
         Picture         =   "ProductionPlanning.frx":108A
         Picture         =   "ProductionPlanning.frx":10A6
      End
      Begin MSForms.ComboBox Combo1 
         Height          =   330
         Left            =   4920
         TabIndex        =   11
         Top             =   0
         Width           =   2610
         VariousPropertyBits=   545282075
         BackColor       =   16777215
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "4604;582"
         ListRows        =   3
         MatchEntry      =   0
         ShowDropButtonWhen=   1
         SpecialEffect   =   0
         FontName        =   "Calibri"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
End
Attribute VB_Name = "FrmProductionPlanning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public OrderType As String
Dim rstCompanyMaster As New ADODB.Recordset
Dim rstProductionPlanning As New ADODB.Recordset
Dim rstBookList As New ADODB.Recordset
Dim rstBoardList As New ADODB.Recordset
Dim rstClassList As New ADODB.Recordset
Dim rstGroupList As New ADODB.Recordset
Dim CxnProductionPlanning As New ADODB.Connection
Dim CxnDistributor As New ADODB.Connection
Dim K As Integer
Dim OutputTo As String
Dim BrPendingPO As Long
Dim BrPendingSO As Long
Dim rstAccountList As New ADODB.Recordset
Dim PartyParentGroups As String
'UPDATE BookMaster INNER JOIN Sheet2 ON BookMAster.Code = Sheet2.Code Set BookMaster.BusyCode = Sheet2.Alias


Private Sub Form_Load()
    On Error GoTo ErrorHandler
    Me.Caption = IIf(OrderType = "M", "Production Planning (Main", "Production Planning (Supplement") + " Orders)"
    CenterForm Me
    BusySystemIndicator True
    rstCompanyMaster.Open "Select PrintName From CompanyMaster", CxnDatabase, adOpenKeyset, adLockReadOnly
    rstGroupList.Open "Select Name, Code From GeneralMaster Where Type = '5' ORDER BY Name", CxnDatabase, adOpenKeyset, adLockReadOnly
    rstGroupList.ActiveConnection = Nothing
    Call FillList(ListView4, "List of Groups...", rstGroupList)
    rstClassList.Open "Select Name, Code From GeneralMaster Where Type = '4' ORDER BY Name", CxnDatabase, adOpenKeyset, adLockReadOnly
    rstClassList.ActiveConnection = Nothing
    Call FillList(ListView1, "List of Classes...", rstClassList)
    rstBoardList.Open "SELECT Name,Code FROM GeneralMaster WHERE Type='2' ORDER BY Name", CxnDatabase, adOpenKeyset, adLockReadOnly
    rstBoardList.ActiveConnection = Nothing
    Call FillList(ListView2, "List of Boards...", rstBoardList)
    If OrderType = "M" Then MhDateInput1.Text = "01-10-" + Trim(Year(FinancialYearFrom) - 2) Else MhDateInput1.Text = Format(FinancialYearFrom, "dd-MM-yyyy")
    If Format(FinancialYearTo, "yyyymmdd") < Format(Date, "yyyymmdd") Then MhDateInput2.Text = Format(FinancialYearTo, "dd-mm-yyyy") Else MhDateInput2.Text = Format(Date, "dd-mm-yyyy")

    Dim BusyDatebaseName As String
   
    BusyDatebaseName = Trim(ReadFromFile("Busy Database Name"))
    Combo1.AddItem Val(Right(BusyDatebaseName, 4)) - 0, 0
    Combo1.AddItem Val(Right(BusyDatebaseName, 4)) + 1, 1
    Combo1.AddItem Val(Right(BusyDatebaseName, 4)) + 2, 2
    Combo1.ListIndex = 1
    BusySystemIndicator False
    
    Exit Sub

ErrorHandler:
    BusySystemIndicator False
    CloseForm Me
    
End Sub

 Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyReturn Then
       SendKeys "{TAB}", True
       KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyEscape Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(3)
        KeyCode = 0
    ElseIf Shift = vbAltMask And KeyCode = vbKeyP Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(2)
        KeyCode = 0
    ElseIf Shift = vbAltMask And KeyCode = vbKeyV Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(1)
        KeyCode = 0
    End If
    
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then CloseForm Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    
    Call CloseRecordset(rstCompanyMaster)
    Call CloseRecordset(rstBookList)
    Call CloseRecordset(rstBoardList)
    Call CloseRecordset(rstClassList)
    Call CloseRecordset(rstGroupList)
    Call CloseRecordset(rstProductionPlanning)
    
End Sub

Private Sub MhDateInput1_Validate(Cancel As Boolean)
    
    If Not IsDate(GetDate(MhDateInput1.Text)) Then
        Cancel = True
    ElseIf OrderType = "M" And (Month(GetDate(MhDateInput1.Text)) <> 10 And Month(GetDate(MhDateInput1.Text)) <> 4) Or Day(GetDate(MhDateInput1.Text)) <> 1 Then
        Cancel = True
    ElseIf OrderType = "S" And Format(GetDate(MhDateInput1.Text), "yyyymmdd") < Format(FinancialYearFrom, "yyyymmdd") Then
        Cancel = True
    End If
    
End Sub
Private Sub MhDateInput2_Validate(Cancel As Boolean)
    
    If Not IsDate(GetDate(MhDateInput2.Text)) Then
        Cancel = True
    ElseIf Format(GetDate(MhDateInput2.Text), "yyyymmdd") < Format(GetDate(MhDateInput1.Text), "yyyymmdd") Then
        FocusSelect Me.ActiveControl
        Cancel = True
    ElseIf OrderType = "M" And Year(GetDate(MhDateInput2.Text)) - Year(GetDate(MhDateInput1.Text)) < 2 Then
        Cancel = True
    ElseIf Format(GetDate(MhDateInput2.Text), "yyyymmdd") > Format(FinancialYearTo, "yyyymmdd") Then
        Cancel = True
    End If
    
End Sub
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
Private Sub ListView2_ItemCheck(ByVal Item As MSComctlLib.ListItem)
     Call BookSelection
End Sub
Private Sub ListView2_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Dim i As Integer
    
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        For i = 1 To ListView2.ListItems.Count
            ListView2.ListItems(i).Checked = True
        Next i
        Call BookSelection
    ElseIf KeyCode = vbKeyD And Shift = vbCtrlMask Then
        For i = 1 To ListView2.ListItems.Count
            ListView2.ListItems(i).Checked = False
        Next i
        Call BookSelection
    End If
End Sub
Private Sub ListView3_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Dim i As Integer
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        For i = 1 To ListView3.ListItems.Count
            ListView3.ListItems(i).Checked = True
        Next i
    ElseIf KeyCode = vbKeyD And Shift = vbCtrlMask Then
        For i = 1 To ListView3.ListItems.Count
            ListView3.ListItems(i).Checked = False
        Next i
    End If
    
End Sub
Private Sub ListView4_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        For i = 1 To ListView4.ListItems.Count
            ListView4.ListItems(i).Checked = True
        Next i
    ElseIf KeyCode = vbKeyD And Shift = vbCtrlMask Then
        For i = 1 To ListView4.ListItems.Count
            ListView4.ListItems(i).Checked = False
        Next i
    End If
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    If Button.Index = 1 Then
        OutputTo = "S"
        If OrderType = "M" Then
          PrintProductionPlanning
        Else
          PrintProductionPlanningS
        End If
        
    ElseIf Button.Index = 2 Then
        OutputTo = "P"
        If OrderType = "M" Then
          PrintProductionPlanning
        Else
          PrintProductionPlanningS
        End If
    ElseIf Button.Index = 3 Then
        CloseForm Me
    End If
End Sub
Private Sub BookSelection()
    If rstBookList.State = adStateOpen Then rstBookList.Close
    'Dim aa As String
    'aa = "SELECT Name,BusyCode As Code FROM BookMaster WHERE [Group] IN (" & SelectedItems(ListView4) & ")  AND [Class] IN (" & SelectedItems(ListView1) & ") AND Board IN (" & SelectedItems(ListView2) & ") AND BusyCode<>'' AND LEN(BusyCode)>=4 AND Type='F' AND [Class]<>'' AND [Group]<>'' ORDER BY Name"
    rstBookList.Open "SELECT Name,BusyCode As Code FROM BookMaster WHERE [Group] IN (" & SelectedItems(ListView4) & ")  AND [Class] IN (" & SelectedItems(ListView1) & ") AND Board IN (" & SelectedItems(ListView2) & ") AND BusyCode<>'' AND LEN(BusyCode)>=4 AND Type='F' AND [Class]<>'' AND [Group]<>'' ORDER BY Name", CxnDatabase, adOpenKeyset, adLockReadOnly
    rstBookList.ActiveConnection = Nothing
    ListView3.ListItems.Clear
    Call FillList(ListView3, "List of Books...", rstBookList)
        
End Sub

Private Sub PrintProductionPlanning()
    
Dim DatabaseName As String
Dim DatabaseNameDistributor As String
    
    Dim FromDate As String, ToDate As String, FromDate02 As String, ToDate02 As String
    
    Dim oExcel As Object
    Dim i As Long, Cnt As Long
    Dim j As Long
    
    Dim Period01 As String, Period02 As String, Period03 As String
    
    On Error GoTo ErrorHandler
    
    DoEvents
    ShowProgressInStatusBar True
    MdiMainMenu.ProgressBar1.Value = 1
    
    DatabaseName = Trim(ReadFromFile("Busy Database Name"))
    If ServerName = "" Or DatabaseName = "" Then Exit Sub
    Screen.MousePointer = vbHourglass
    CxnDatabase.Execute "UPDATE BookMaster SET SaleLY1003=0,SaleTY0409=0,StockTransferLY1003=0,StockTransferTY0409=0,SpecimenLY1003=0,SpecimenTY0409=0,PendingSO=0,SaleableStock=0,POLTLY1003=0,POLY0409=0,POLY1003=0,POTY0409=0,PendingPO=0,ESO30=0,ESO60=0,ESO90=0,ESO150=0,PSO15=0,PSO30=0"
    i = 0: K = 0: j = 0
    CxnProductionPlanning.CursorLocation = adUseClient: CxnProductionPlanning.CommandTimeout = 0
    
    Do While True
        i = InStr(1, DatabaseName, ",")
        If CxnProductionPlanning.State = adStateOpen Then CxnProductionPlanning.Close
        If i = 0 Then CxnProductionPlanning.Open "Provider=SQLOLEDB.1;Password=" & ServerPassword & ";Persist Security Info=True;User ID=sa;Initial Catalog=" & Mid(DatabaseName, 1) & ";Data Source=" & ServerName Else CxnProductionPlanning.Open "Provider=SQLOLEDB.1;Password=" & ServerPassword & ";Persist Security Info=True;User ID=sa;Initial Catalog=" & Mid(DatabaseName, 1, i - 1) & ";Data Source=" & ServerName

        K = K + 1
        If rstProductionPlanning.State = adStateOpen Then rstProductionPlanning.Close
        If K = 1 Then   'Last Year Data Processing
            If OrderType = "M" Then
                If Month(GetDate(MhDateInput1.Text)) = 10 Then  'Oct-Mar
                    FromDate = "01-Oct-" + Trim(Year(GetDate(MhDateInput1.Text)) + 1): ToDate = "31-Mar-" + Trim(Year(GetDate(MhDateInput1.Text)) + 2)
                    
                    rstProductionPlanning.Open "SELECT Alias," & _
                                                                  "(SELECT ISNULL(SUM(0-Value1),0) FROM Tran2 WHERE VchType IN (3,9) AND RecType=2 AND Date>='" & FromDate & "' AND Date<='" & ToDate & "' AND MasterCode1=M.Code) As Sale01," & _
                                                                  "0 As Sale02," & _
                                                                  "(SELECT ISNULL(SUM(0-Value1),0) FROM Tran2 C INNER JOIN Tran1 P ON P.VchCode=C.VchCode WHERE C.VchType IN (4,11) AND C.Date>='" & FromDate & "' AND C.Date<='" & ToDate & "' AND C.MasterCode1=M.Code AND P.MasterCode1 IN ((SELECT Code FROM Master1 WHERE MasterType=2 AND ParentGrp IN (SELECT Code FROM Master1 WHERE MasterType=1 AND UPPER(Name) LIKE '%BRANCH%'))))+(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE VchType NOT IN (12,13,8) AND Date>='" & FromDate & "' AND Date<='" & ToDate & "' AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND ParentGrp IN (SELECT Code FROM Master1 WHERE MasterType=10 AND UPPER(Name) LIKE '%BRANCH%'))) As BranchStockTransfer01," & _
                                                                  "0 As BranchStockTransfer02," & _
                                                                  "(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE VchType NOT IN (12,13,8) AND Date>='" & FromDate & "' AND Date<='" & ToDate & "' AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND UPPER(Name) LIKE '%SPECIMEN%')) As Specimen01," & _
                                                                  "0 As Specimen02 " & _
                                                                  "FROM Master1 M WHERE MasterType=6 AND Alias<>'' AND LEFT(Alias,4) IN (" & SelectedItems(ListView3, True) & ") ORDER BY Alias", CxnProductionPlanning, adOpenKeyset, adLockReadOnly
                Else    'Apr-Sep
                    FromDate = "01-Apr-" + Trim(Year(GetDate(MhDateInput1.Text)) + 1): ToDate = "30-Sep-" + Trim(Year(GetDate(MhDateInput1.Text)) + 1)
                    FromDate02 = "01-Oct-" + Trim(Year(GetDate(MhDateInput1.Text)) + 1): ToDate02 = "31-Mar-" + Trim(Year(GetDate(MhDateInput1.Text)) + 2)
                    rstProductionPlanning.Open "SELECT Alias," & _
                                                                  "(SELECT ISNULL(SUM(0-Value1),0) FROM Tran2 WHERE VchType IN (3,9) AND RecType=2 AND Date>='" & FromDate & "' AND Date<='" & ToDate & "' AND MasterCode1=M.Code) As Sale01," & _
                                                                  "(SELECT ISNULL(SUM(0-Value1),0) FROM Tran2 WHERE VchType IN (3,9) AND RecType=2 AND Date>='" & FromDate02 & "' AND Date<='" & ToDate02 & "' AND MasterCode1=M.Code) As Sale02," & _
                                                                  "(SELECT ISNULL(SUM(0-Value1),0) FROM Tran2 C INNER JOIN Tran1 P ON P.VchCode=C.VchCode WHERE C.VchType IN (4,11) AND C.Date>='" & FromDate & "' AND C.Date<='" & ToDate & "' AND C.MasterCode1=M.Code AND P.MasterCode1 IN ((SELECT Code FROM Master1 WHERE MasterType=2 AND ParentGrp IN (SELECT Code FROM Master1 WHERE MasterType=1 AND UPPER(Name) LIKE '%BRANCH%'))))+(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE VchType NOT IN (12,13,8) AND Date>='" & FromDate & "' AND Date<='" & ToDate & "' AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND ParentGrp IN (SELECT Code FROM Master1 WHERE MasterType=10 AND UPPER(Name) LIKE '%BRANCH%'))) As BranchStockTransfer01," & _
                                                                  "(SELECT ISNULL(SUM(0-Value1),0) FROM Tran2 C INNER JOIN Tran1 P ON P.VchCode=C.VchCode WHERE C.VchType IN (4,11) AND C.Date>='" & FromDate02 & "' AND C.Date<='" & ToDate02 & "' AND C.MasterCode1=M.Code AND P.MasterCode1 IN ((SELECT Code FROM Master1 WHERE MasterType=2 AND ParentGrp IN (SELECT Code FROM Master1 WHERE MasterType=1 AND UPPER(Name) LIKE '%BRANCH%'))))+(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE VchType NOT IN (12,13,8) AND Date>='" & FromDate & "' AND Date<='" & ToDate & "' AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND ParentGrp IN (SELECT Code FROM Master1 WHERE MasterType=10 AND UPPER(Name) LIKE '%BRANCH%'))) As BranchStockTransfer02," & _
                                                                  "(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE VchType NOT IN (12,13,8) AND Date>='" & FromDate & "' AND Date<='" & ToDate & "' AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND UPPER(Name) LIKE '%SPECIMEN%')) As Specimen01," & _
                                                                  "(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE VchType NOT IN (12,13,8) AND Date>='" & FromDate02 & "' AND Date<='" & ToDate02 & "' AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND UPPER(Name) LIKE '%SPECIMEN%')) As Specimen02 " & _
                                                                  "FROM Master1 M WHERE MasterType=6 AND Alias<>'' AND LEFT(Alias,4) IN (" & SelectedItems(ListView3, True) & ") ORDER BY Alias", CxnProductionPlanning, adOpenKeyset, adLockReadOnly
                End If
            Else
                FromDate = Trim(Day(GetDate(MhDateInput2.Text))) + "-" + MonthName(Trim(Month(GetDate(MhDateInput2.Text))), True) + "-" + Trim(Year(GetDate(MhDateInput2.Text)) - 1)
                If Not IsDate(FromDate) Then FromDate = Trim(Day(GetDate(MhDateInput2.Text)) - 1) + "-" + MonthName(Trim(Month(GetDate(MhDateInput2.Text))), True) + "-" + Trim(Year(GetDate(MhDateInput2.Text)) - 1)
                
                rstProductionPlanning.Open "SELECT Alias," & _
                                                              "(SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran2 WHERE VchType=9 AND RecType=2 AND Date>='" & FromDate & "' AND Date<='" & Format(DateAdd("d", 30, FromDate), "dd-MMM-yyyy") & "' AND MasterCode1=M.Code) As Sale30," & _
                                                              "(SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran2 WHERE VchType=9 AND RecType=2 AND Date>='" & FromDate & "' AND Date<='" & Format(DateAdd("d", 60, FromDate), "dd-MMM-yyyy") & "' AND MasterCode1=M.Code) As Sale60," & _
                                                              "(SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran2 WHERE VchType=9 AND RecType=2 AND Date>='" & FromDate & "' AND Date<='" & Format(DateAdd("d", 90, FromDate), "dd-MMM-yyyy") & "' AND MasterCode1=M.Code) As Sale90," & _
                                                              "(SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran2 WHERE VchType=9 AND RecType=2 AND Date>='" & FromDate & "' AND Date<='" & Format(DateAdd("d", 150, FromDate), "dd-MMM-yyyy") & "' AND MasterCode1=M.Code) As Sale150," & _
                                                              "(SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran2 C INNER JOIN Tran1 P  ON P.VchCode=C.VchCode WHERE C.VchType=11 AND C.Date>='" & FromDate & "' AND C.Date<='" & Format(DateAdd("d", 30, FromDate), "dd-MMM-yyyy") & "' AND C.MasterCode1=M.Code AND P.MasterCode1 IN ((SELECT Code FROM Master1 WHERE MasterType=2 AND ParentGrp IN (SELECT Code FROM Master1 WHERE MasterType=1 AND UPPER(Name) LIKE '%BRANCH%'))))+(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE Value1>0 AND VchType NOT IN (12,13,8) AND Date>='" & FromDate & "' AND Date<='" & Format(DateAdd("d", 30, FromDate), "dd-MMM-yyyy") & "' AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND ParentGrp IN (SELECT Code FROM Master1 WHERE MasterType=10 AND UPPER(Name) LIKE '%BRANCH%'))) As BranchStockTransfer30," & _
                                                              "(SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran2 C INNER JOIN Tran1 P  ON P.VchCode=C.VchCode WHERE C.VchType=11 AND C.Date>='" & FromDate & "' AND C.Date<='" & Format(DateAdd("d", 60, FromDate), "dd-MMM-yyyy") & "' AND C.MasterCode1=M.Code AND P.MasterCode1 IN ((SELECT Code FROM Master1 WHERE MasterType=2 AND ParentGrp IN (SELECT Code FROM Master1 WHERE MasterType=1 AND UPPER(Name) LIKE '%BRANCH%'))))+(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE Value1>0 AND VchType NOT IN (12,13,8) AND Date>='" & FromDate & "' AND Date<='" & Format(DateAdd("d", 60, FromDate), "dd-MMM-yyyy") & "' AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND ParentGrp IN (SELECT Code FROM Master1 WHERE MasterType=10 AND UPPER(Name) LIKE '%BRANCH%'))) As BranchStockTransfer60," & _
                                                              "(SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran2 C INNER JOIN Tran1 P  ON P.VchCode=C.VchCode WHERE C.VchType=11 AND C.Date>='" & FromDate & "' AND C.Date<='" & Format(DateAdd("d", 90, FromDate), "dd-MMM-yyyy") & "' AND C.MasterCode1=M.Code AND P.MasterCode1 IN ((SELECT Code FROM Master1 WHERE MasterType=2 AND ParentGrp IN (SELECT Code FROM Master1 WHERE MasterType=1 AND UPPER(Name) LIKE '%BRANCH%'))))+(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE Value1>0 AND VchType NOT IN (12,13,8) AND Date>='" & FromDate & "' AND Date<='" & Format(DateAdd("d", 90, FromDate), "dd-MMM-yyyy") & "' AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND ParentGrp IN (SELECT Code FROM Master1 WHERE MasterType=10 AND UPPER(Name) LIKE '%BRANCH%'))) As BranchStockTransfer90," & _
                                                              "(SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran2 C INNER JOIN Tran1 P  ON P.VchCode=C.VchCode WHERE C.VchType=11 AND C.Date>='" & FromDate & "' AND C.Date<='" & Format(DateAdd("d", 150, FromDate), "dd-MMM-yyyy") & "' AND C.MasterCode1=M.Code AND P.MasterCode1 IN ((SELECT Code FROM Master1 WHERE MasterType=2 AND ParentGrp IN (SELECT Code FROM Master1 WHERE MasterType=1 AND UPPER(Name) LIKE '%BRANCH%'))))+(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE Value1>0 AND VchType NOT IN (12,13,8) AND Date>='" & FromDate & "' AND Date<='" & Format(DateAdd("d", 150, FromDate), "dd-MMM-yyyy") & "' AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND ParentGrp IN (SELECT Code FROM Master1 WHERE MasterType=10 AND UPPER(Name) LIKE '%BRANCH%'))) As BranchStockTransfer150," & _
                                                              "(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE Value1>0 AND VchType NOT IN (12,13,8) AND Date>='" & FromDate & "' AND Date<='" & Format(DateAdd("d", 30, FromDate), "dd-MMM-yyyy") & "' AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND UPPER(Name) LIKE '%SPECIMEN%')) As Specimen30," & _
                                                              "(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE Value1>0 AND VchType NOT IN (12,13,8) AND Date>='" & FromDate & "' AND Date<='" & Format(DateAdd("d", 60, FromDate), "dd-MMM-yyyy") & "' AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND UPPER(Name) LIKE '%SPECIMEN%')) As Specimen60," & _
                                                              "(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE Value1>0 AND VchType NOT IN (12,13,8) AND Date>='" & FromDate & "' AND Date<='" & Format(DateAdd("d", 90, FromDate), "dd-MMM-yyyy") & "' AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND UPPER(Name) LIKE '%SPECIMEN%')) As Specimen90," & _
                                                              "(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE Value1>0 AND VchType NOT IN (12,13,8) AND Date>='" & FromDate & "' AND Date<='" & Format(DateAdd("d", 150, FromDate), "dd-MMM-yyyy") & "' AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND UPPER(Name) LIKE '%SPECIMEN%')) As Specimen150," & _
                                                              "(SELECT ISNULL(SUM(0-Value1),0) FROM Tran2 WHERE VchType IN (3,9) AND RecType=2 AND Date>='" & "01-Apr-" + Trim(Year(GetDate(MhDateInput1.Text)) - 1) & "' AND Date<='" & "31-Mar-" + Trim(Year(GetDate(MhDateInput1.Text))) & "' AND MasterCode1=M.Code) As LYSale," & _
                                                              "(SELECT ISNULL(SUM(0-Value1),0) FROM Tran2 C INNER JOIN Tran1 P ON P.VchCode=C.VchCode WHERE C.VchType IN (4,11) AND C.Date>='" & "01-Apr-" + Trim(Year(GetDate(MhDateInput1.Text)) - 1) & "' AND C.Date<='" & "31-Mar-" + Trim(Year(GetDate(MhDateInput1.Text))) & "' AND C.MasterCode1=M.Code AND P.MasterCode1 IN ((SELECT Code FROM Master1 WHERE MasterType=2 AND ParentGrp IN (SELECT Code FROM Master1 WHERE MasterType=1 AND UPPER(Name) LIKE '%BRANCH%'))))+(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE VchType NOT IN (12,13,8) AND Date>='" & "01-Apr-" + Trim(Year(GetDate(MhDateInput1.Text)) - 1) & "' AND Date<='" & "31-Mar-" + Trim(Year(GetDate(MhDateInput1.Text))) & "' AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND ParentGrp IN (SELECT Code FROM Master1 WHERE MasterType=10 AND UPPER(Name) LIKE '%BRANCH%'))) As LYBranchStockTransfer," & _
                                                              "(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE VchType NOT IN (12,13,8) AND Date>='" & "01-Apr-" + Trim(Year(GetDate(MhDateInput1.Text)) - 1) & "' AND Date<='" & "31-Mar-" + Trim(Year(GetDate(MhDateInput1.Text))) & "' AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND UPPER(Name) LIKE '%SPECIMEN%')) As LYSpecimen," & _
                                                              "(SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran2 WHERE VchType=9 AND RecType=2 AND Date>='" & Format(DateAdd("d", -15, GetDate(MhDateInput2.Text)), "dd-MMM-yyyy") & "' AND Date<='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code) As CSale15," & _
                                                              "(SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran2 WHERE VchType=9 AND RecType=2 AND Date>='" & Format(DateAdd("d", -30, GetDate(MhDateInput2.Text)), "dd-MMM-yyyy") & "' AND Date<='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code) As CSale30," & _
                                                              "(SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran2 C INNER JOIN Tran1 P  ON P.VchCode=C.VchCode WHERE C.VchType=11 AND C.Date>='" & Format(DateAdd("d", -15, GetDate(MhDateInput2.Text)), "dd-MMM-yyyy") & "' AND C.Date<='" & GetDate(MhDateInput2.Text) & "' AND C.MasterCode1=M.Code AND P.MasterCode1 IN ((SELECT Code FROM Master1 WHERE MasterType=2 AND ParentGrp IN (SELECT Code FROM Master1 WHERE MasterType=1 AND UPPER(Name) LIKE '%BRANCH%'))))+(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE Value1>0 AND VchType NOT IN (12,13,8) AND Date>='" & Format(DateAdd("d", -15, GetDate(MhDateInput2.Text)), "dd-MMM-yyyy") & "' AND Date<='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND ParentGrp IN (SELECT Code FROM Master1 WHERE MasterType=10 AND UPPER(Name) LIKE '%BRANCH%'))) As CBranchStockTransfer15," & _
                                                              "(SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran2 C INNER JOIN Tran1 P  ON P.VchCode=C.VchCode WHERE C.VchType=11 AND C.Date>='" & Format(DateAdd("d", -30, GetDate(MhDateInput2.Text)), "dd-MMM-yyyy") & "' AND C.Date<='" & GetDate(MhDateInput2.Text) & "' AND C.MasterCode1=M.Code AND P.MasterCode1 IN ((SELECT Code FROM Master1 WHERE MasterType=2 AND ParentGrp IN (SELECT Code FROM Master1 WHERE MasterType=1 AND UPPER(Name) LIKE '%BRANCH%'))))+(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE Value1>0 AND VchType NOT IN (12,13,8) AND Date>='" & Format(DateAdd("d", -30, GetDate(MhDateInput2.Text)), "dd-MMM-yyyy") & "' AND Date<='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND ParentGrp IN (SELECT Code FROM Master1 WHERE MasterType=10 AND UPPER(Name) LIKE '%BRANCH%'))) As CBranchStockTransfer30," & _
                                                              "(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE Value1>0 AND VchType NOT IN (12,13,8) AND Date>='" & Format(DateAdd("d", -15, GetDate(MhDateInput2.Text)), "dd-MMM-yyyy") & "' AND Date<='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND UPPER(Name) LIKE '%SPECIMEN%')) As CSpecimen15," & _
                                                              "(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE Value1>0 AND VchType NOT IN (12,13,8) AND Date>='" & Format(DateAdd("d", -30, GetDate(MhDateInput2.Text)), "dd-MMM-yyyy") & "' AND Date<='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND UPPER(Name) LIKE '%SPECIMEN%')) As CSpecimen30 " & _
                                                              "FROM Master1 M WHERE MasterType=6 AND Alias<>'' AND LEFT(Alias,4) IN (" & SelectedItems(ListView3, True) & ") ORDER BY Alias", CxnProductionPlanning, adOpenKeyset, adLockReadOnly
            
            
            End If
            
            rstProductionPlanning.ActiveConnection = Nothing
            Call UpdatePPFigures("1") 'Update Sale, Stock Transfer And Specimen Figures
            MdiMainMenu.ProgressBar1.Value = MdiMainMenu.ProgressBar1.Value + 16.5
        
        Else    'Current Year Data Processing
        
            If OrderType = "M" Then
                FromDate = "01-Apr-" + Trim(Year(GetDate(MhDateInput1.Text)) + 2): ToDate = "30-Sep-" + Trim(Year(GetDate(MhDateInput1.Text)) + 2)
                rstProductionPlanning.Open "SELECT Name,Alias," & _
                                                              "(SELECT ISNULL(SUM(0-Value1),0) FROM Tran2 WHERE VchType IN (3,9) AND RecType=2 AND Date>='" & FromDate & "' AND Date<='" & ToDate & "' AND MasterCode1=M.Code) As Sale," & _
                                                              "(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE VchType=3 AND RecType=2 AND Date>='" & FromDate & "' AND Date<='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code) As SaleReturn," & _
                                                              "(SELECT ISNULL(SUM(0-Value1),0) FROM Tran2 C INNER JOIN Tran1 P ON P.VchCode=C.VchCode WHERE C.VchType IN (4,11) AND C.Date>='" & FromDate & "' AND C.Date<='" & ToDate & "' AND C.MasterCode1=M.Code AND P.MasterCode1 IN ((SELECT Code FROM Master1 WHERE MasterType=2 AND ParentGrp IN (SELECT Code FROM Master1 WHERE MasterType=1 AND UPPER(Name) LIKE '%BRANCH%'))))+(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE VchType NOT IN (12,13,8) AND Date>='" & FromDate & "' AND Date<='" & ToDate & "' AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND ParentGrp IN (SELECT Code FROM Master1 WHERE MasterType=10 AND UPPER(Name) LIKE '%BRANCH%'))) As BranchStockTransfer," & _
                                                              "(SELECT ISNULL(SUM(Value1),0) FROM Tran2 C INNER JOIN Tran1 P ON P.VchCode=C.VchCode WHERE C.VchType=4 AND C.Date>='" & FromDate & "' AND C.Date<='" & GetDate(MhDateInput2.Text) & "' AND C.MasterCode1=M.Code AND P.MasterCode1 IN ((SELECT Code FROM Master1 WHERE MasterType=2 AND ParentGrp IN (SELECT Code FROM Master1 WHERE MasterType=1 AND UPPER(Name) LIKE '%BRANCH%'))))+(SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran2 WHERE VchType NOT IN (12,13,8) AND Value1<0 AND Date>='" & FromDate & "' AND Date<='" & ToDate & "' AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND UPPER(Name) LIKE '%BRANCH%')) As BranchStockTransferReturn," & _
                                                              "(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE VchType NOT IN (12,13,8) AND Date>='" & FromDate & "' AND Date<='" & ToDate & "' AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND UPPER(Name) LIKE '%SPECIMEN%')) As Specimen," & _
                                                              "(SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran3 WHERE VchType=12 AND Date<='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code AND CM1 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND ParentGrp IN (SELECT Code FROM Master1 WHERE MasterType=10 AND UPPER(Name) LIKE '%" & UCase(MCGroup) & "%'))) As SaleOrder," & _
                                                              "(SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran3 WHERE RecType=4 AND Method=2 AND RefCode IN (SELECT RefCode FROM Tran3 WHERE VchType=12 AND Date<='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code AND CM1 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND ParentGrp IN (SELECT Code FROM Master1 WHERE MasterType=10 AND UPPER(Name) LIKE '%" & UCase(MCGroup) & "%')))) As SaleOrderSupplied," & _
                                                              "(SELECT ISNULL(SUM(D1),0) FROM Tran4 WHERE MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND ParentGrp IN (SELECT Code FROM Master1 WHERE MasterType=10 AND UPPER(Name) LIKE '%" & UCase(MCGroup) & "%'))) As OpBal," & _
                                                              "(SELECT ISNULL(SUM(0-Value1),0) FROM Tran2 WHERE VchType IN (3,9) AND RecType=2 AND Date>='" & FromDate & "' AND Date <='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND ParentGrp IN (SELECT Code FROM Master1 WHERE MasterType=10 AND UPPER(Name) LIKE '%" & UCase(MCGroup) & "%'))) As NetSale," & _
                                                              "(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE VchType=5 AND Date>='" & FromDate & "' AND Date <='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND ParentGrp IN (SELECT Code FROM Master1 WHERE MasterType=10 AND UPPER(Name) LIKE '%" & UCase(MCGroup) & "%'))) As NetStockTransfer," & _
                                                              "(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE VchType IN (4,11) AND Date>='" & FromDate & "' AND Date <='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND ParentGrp IN (SELECT Code FROM Master1 WHERE MasterType=10 AND UPPER(Name) LIKE '%" & UCase(MCGroup) & "%'))) As NetPurchase," & _
                                                              "(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE VchType=8 AND Date>='" & FromDate & "' AND Date <='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND ParentGrp IN (SELECT Code FROM Master1 WHERE MasterType=10 AND UPPER(Name) LIKE '%" & UCase(MCGroup) & "%'))) As NetStockAdjustment " & _
                                                              "FROM Master1 M WHERE MasterType=6 AND Alias<>'' AND LEFT(Alias,4) IN (" & SelectedItems(ListView3, True) & ") ORDER BY Alias", CxnProductionPlanning, adOpenKeyset, adLockReadOnly
            Else
                FromDate = Trim(Day(GetDate(MhDateInput2.Text))) + "-" + MonthName(Trim(Month(GetDate(MhDateInput2.Text))), True) + "-" + Trim(Year(GetDate(MhDateInput2.Text)) - 1)
                If Not IsDate(FromDate) Then FromDate = Trim(Day(GetDate(MhDateInput2.Text)) - 1) + "-" + MonthName(Trim(Month(GetDate(MhDateInput2.Text))), True) + "-" + Trim(Year(GetDate(MhDateInput2.Text)) - 1)
                 rstProductionPlanning.Open "SELECT Name,Alias," & _
                                                              "(SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran2 WHERE VchType=9 AND RecType=2 AND Date>='" & FromDate & "' AND Date<='" & Format(DateAdd("d", 30, FromDate), "dd-MMM-yyyy") & "' AND MasterCode1=M.Code) As Sale30,(SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran2 WHERE VchType=9 AND RecType=2 AND Date>='" & FromDate & "' AND Date<='" & Format(DateAdd("d", 60, FromDate), "dd-MMM-yyyy") & "' AND MasterCode1=M.Code) As Sale60,(SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran2 WHERE VchType=9 AND RecType=2 AND Date>='" & FromDate & "' AND Date<='" & Format(DateAdd("d", 90, FromDate), "dd-MMM-yyyy") & "' AND MasterCode1=M.Code) As Sale90,(SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran2 WHERE VchType=9 AND RecType=2 AND Date>='" & FromDate & "' AND Date<='" & Format(DateAdd("d", 150, FromDate), "dd-MMM-yyyy") & "' AND MasterCode1=M.Code) As Sale150," & _
                                                              "(SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran2 C INNER JOIN Tran1 P  ON P.VchCode=C.VchCode WHERE C.VchType=11 AND C.Date>='" & FromDate & "' AND C.Date<='" & Format(DateAdd("d", 30, FromDate), "dd-MMM-yyyy") & "' AND C.MasterCode1=M.Code AND P.MasterCode1 IN ((SELECT Code FROM Master1 WHERE MasterType=2 AND ParentGrp IN (SELECT Code FROM Master1 WHERE MasterType=1 AND UPPER(Name) LIKE '%BRANCH%'))))+(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE Value1>0 AND VchType NOT IN (12,13,8) AND Date>='" & FromDate & "' AND Date<='" & Format(DateAdd("d", 30, FromDate), "dd-MMM-yyyy") & "' AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND ParentGrp IN (SELECT Code FROM Master1 WHERE MasterType=10 AND UPPER(Name) LIKE '%BRANCH%'))) As BranchStockTransfer30," & _
                                                              "(SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran2 C INNER JOIN Tran1 P  ON P.VchCode=C.VchCode WHERE C.VchType=11 AND C.Date>='" & FromDate & "' AND C.Date<='" & Format(DateAdd("d", 60, FromDate), "dd-MMM-yyyy") & "' AND C.MasterCode1=M.Code AND P.MasterCode1 IN ((SELECT Code FROM Master1 WHERE MasterType=2 AND ParentGrp IN (SELECT Code FROM Master1 WHERE MasterType=1 AND UPPER(Name) LIKE '%BRANCH%'))))+(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE Value1>0 AND VchType NOT IN (12,13,8) AND Date>='" & FromDate & "' AND Date<='" & Format(DateAdd("d", 60, FromDate), "dd-MMM-yyyy") & "' AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND ParentGrp IN (SELECT Code FROM Master1 WHERE MasterType=10 AND UPPER(Name) LIKE '%BRANCH%'))) As BranchStockTransfer60," & _
                                                              "(SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran2 C INNER JOIN Tran1 P  ON P.VchCode=C.VchCode WHERE C.VchType=11 AND C.Date>='" & FromDate & "' AND C.Date<='" & Format(DateAdd("d", 90, FromDate), "dd-MMM-yyyy") & "' AND C.MasterCode1=M.Code AND P.MasterCode1 IN ((SELECT Code FROM Master1 WHERE MasterType=2 AND ParentGrp IN (SELECT Code FROM Master1 WHERE MasterType=1 AND UPPER(Name) LIKE '%BRANCH%'))))+(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE Value1>0 AND VchType NOT IN (12,13,8) AND Date>='" & FromDate & "' AND Date<='" & Format(DateAdd("d", 90, FromDate), "dd-MMM-yyyy") & "' AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND ParentGrp IN (SELECT Code FROM Master1 WHERE MasterType=10 AND UPPER(Name) LIKE '%BRANCH%'))) As BranchStockTransfer90," & _
                                                              "(SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran2 C INNER JOIN Tran1 P  ON P.VchCode=C.VchCode WHERE C.VchType=11 AND C.Date>='" & FromDate & "' AND C.Date<='" & Format(DateAdd("d", 150, FromDate), "dd-MMM-yyyy") & "' AND C.MasterCode1=M.Code AND P.MasterCode1 IN ((SELECT Code FROM Master1 WHERE MasterType=2 AND ParentGrp IN (SELECT Code FROM Master1 WHERE MasterType=1 AND UPPER(Name) LIKE '%BRANCH%'))))+(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE Value1>0 AND VchType NOT IN (12,13,8) AND Date>='" & FromDate & "' AND Date<='" & Format(DateAdd("d", 150, FromDate), "dd-MMM-yyyy") & "' AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND ParentGrp IN (SELECT Code FROM Master1 WHERE MasterType=10 AND UPPER(Name) LIKE '%BRANCH%'))) As BranchStockTransfer150," & _
                                                              "(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE Value1>0 AND VchType NOT IN (12,13,8) AND Date>='" & FromDate & "' AND Date<='" & Format(DateAdd("d", 30, FromDate), "dd-MMM-yyyy") & "' AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND UPPER(Name) LIKE '%SPECIMEN%')) As Specimen30,(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE Value1>0 AND VchType NOT IN (12,13,8) AND Date>='" & FromDate & "' AND Date<='" & Format(DateAdd("d", 60, FromDate), "dd-MMM-yyyy") & "' AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND UPPER(Name) LIKE '%SPECIMEN%')) As Specimen60," & _
                                                              "(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE Value1>0 AND VchType NOT IN (12,13,8) AND Date>='" & FromDate & "' AND Date<='" & Format(DateAdd("d", 90, FromDate), "dd-MMM-yyyy") & "' AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND UPPER(Name) LIKE '%SPECIMEN%')) As Specimen90,(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE Value1>0 AND VchType NOT IN (12,13,8) AND Date>='" & FromDate & "' AND Date<='" & Format(DateAdd("d", 150, FromDate), "dd-MMM-yyyy") & "' AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND UPPER(Name) LIKE '%SPECIMEN%')) As Specimen150," & _
                                                              "(SELECT ISNULL(SUM(0-Value1),0) FROM Tran2 WHERE VchType IN (3,9) AND RecType=2 AND Date>='" & "01-Apr-" + Trim(Year(GetDate(MhDateInput1.Text))) & "' AND Date<='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code) As CYSale," & _
                                                              "(SELECT ISNULL(SUM(0-Value1),0) FROM Tran2 C INNER JOIN Tran1 P ON P.VchCode=C.VchCode WHERE C.VchType IN (4,11) AND C.Date>='" & "01-Apr-" + Trim(Year(GetDate(MhDateInput1.Text))) & "' AND C.Date<='" & GetDate(MhDateInput2.Text) & "' AND C.MasterCode1=M.Code AND P.MasterCode1 IN ((SELECT Code FROM Master1 WHERE MasterType=2 AND ParentGrp IN (SELECT Code FROM Master1 WHERE MasterType=1 AND UPPER(Name) LIKE '%BRANCH%'))))+(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE VchType NOT IN (12,13,8) AND Date>='" & "01-Apr-" + Trim(Year(GetDate(MhDateInput1.Text))) & "' AND Date<='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND ParentGrp IN (SELECT Code FROM Master1 WHERE MasterType=10 AND UPPER(Name) LIKE '%BRANCH%'))) As CYBranchStockTransfer," & _
                                                              "(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE VchType NOT IN (12,13,8) AND Date>='" & "01-Apr-" + Trim(Year(GetDate(MhDateInput1.Text))) & "' AND Date<='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND UPPER(Name) LIKE '%SPECIMEN%')) As CYSpecimen," & _
                                                              "(SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran2 WHERE VchType=9 AND RecType=2 AND Date>='" & Format(DateAdd("d", -15, GetDate(MhDateInput2.Text)), "dd-MMM-yyyy") & "' AND Date<='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code) As CSale15," & _
                                                              "(SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran2 WHERE VchType=9 AND RecType=2 AND Date>='" & Format(DateAdd("d", -30, GetDate(MhDateInput2.Text)), "dd-MMM-yyyy") & "' AND Date<='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code) As CSale30," & _
                                                              "(SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran2 C INNER JOIN Tran1 P  ON P.VchCode=C.VchCode WHERE C.VchType=11 AND C.Date>='" & Format(DateAdd("d", -15, GetDate(MhDateInput2.Text)), "dd-MMM-yyyy") & "' AND C.Date<='" & GetDate(MhDateInput2.Text) & "' AND C.MasterCode1=M.Code AND P.MasterCode1 IN ((SELECT Code FROM Master1 WHERE MasterType=2 AND ParentGrp IN (SELECT Code FROM Master1 WHERE MasterType=1 AND UPPER(Name) LIKE '%BRANCH%'))))+(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE Value1>0 AND VchType NOT IN (12,13,8) AND Date>='" & Format(DateAdd("d", -15, GetDate(MhDateInput2.Text)), "dd-MMM-yyyy") & "' AND Date<='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND ParentGrp IN (SELECT Code FROM Master1 WHERE MasterType=10 AND UPPER(Name) LIKE '%BRANCH%'))) As CBranchStockTransfer15," & _
                                                              "(SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran2 C INNER JOIN Tran1 P  ON P.VchCode=C.VchCode WHERE C.VchType=11 AND C.Date>='" & Format(DateAdd("d", -30, GetDate(MhDateInput2.Text)), "dd-MMM-yyyy") & "' AND C.Date<='" & GetDate(MhDateInput2.Text) & "' AND C.MasterCode1=M.Code AND P.MasterCode1 IN ((SELECT Code FROM Master1 WHERE MasterType=2 AND ParentGrp IN (SELECT Code FROM Master1 WHERE MasterType=1 AND UPPER(Name) LIKE '%BRANCH%'))))+(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE Value1>0 AND VchType NOT IN (12,13,8) AND Date>='" & Format(DateAdd("d", -30, GetDate(MhDateInput2.Text)), "dd-MMM-yyyy") & "' AND Date<='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND ParentGrp IN (SELECT Code FROM Master1 WHERE MasterType=10 AND UPPER(Name) LIKE '%BRANCH%'))) As CBranchStockTransfer30," & _
                                                              "(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE Value1>0 AND VchType NOT IN (12,13,8) AND Date>='" & Format(DateAdd("d", -15, GetDate(MhDateInput2.Text)), "dd-MMM-yyyy") & "' AND Date<='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND UPPER(Name) LIKE '%SPECIMEN%')) As CSpecimen15," & _
                                                              "(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE Value1>0 AND VchType NOT IN (12,13,8) AND Date>='" & Format(DateAdd("d", -30, GetDate(MhDateInput2.Text)), "dd-MMM-yyyy") & "' AND Date<='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND UPPER(Name) LIKE '%SPECIMEN%')) As CSpecimen30," & _
                                                              "(SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran3 WHERE VchType=12 AND Date<='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code AND CM1 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND ParentGrp IN (SELECT Code FROM Master1 WHERE MasterType=10 AND UPPER(Name) LIKE '%" & UCase(MCGroup) & "%'))) As SaleOrder," & _
                                                              "(SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran3 WHERE RecType=4 AND Method=2 AND RefCode IN (SELECT RefCode FROM Tran3 WHERE VchType=12 AND Date<='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code AND CM1 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND ParentGrp IN (SELECT Code FROM Master1 WHERE MasterType=10 AND UPPER(Name) LIKE '%" & UCase(MCGroup) & "%')))) As SaleOrderSupplied," & _
                                                              "(SELECT ISNULL(SUM(D1),0) FROM Tran4 WHERE MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND ParentGrp IN (SELECT Code FROM Master1 WHERE MasterType=10 AND UPPER(Name) LIKE '%" & UCase(MCGroup) & "%'))) As OpBal," & _
                                                              "(SELECT ISNULL(SUM(0-Value1),0) FROM Tran2 WHERE VchType IN (3,9) AND RecType=2 AND Date>='" & FromDate & "' AND Date <='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND ParentGrp IN (SELECT Code FROM Master1 WHERE MasterType=10 AND UPPER(Name) LIKE '%" & UCase(MCGroup) & "%'))) As NetSale," & _
                                                              "(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE VchType=5 AND Date>='" & FromDate & "' AND Date <='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND ParentGrp IN (SELECT Code FROM Master1 WHERE MasterType=10 AND UPPER(Name) LIKE '%" & UCase(MCGroup) & "%'))) As NetStockTransfer," & _
                                                              "(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE VchType IN (4,11) AND Date>='" & FromDate & "' AND Date <='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND ParentGrp IN (SELECT Code FROM Master1 WHERE MasterType=10 AND UPPER(Name) LIKE '%" & UCase(MCGroup) & "%'))) As NetPurchase," & _
                                                              "(SELECT ISNULL(SUM(Value1),0) FROM Tran2 WHERE VchType=8 AND Date>='" & FromDate & "' AND Date <='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=M.Code AND MasterCode2 IN (SELECT Code FROM Master1 WHERE MasterType=11 AND ParentGrp IN (SELECT Code FROM Master1 WHERE MasterType=10 AND UPPER(Name) LIKE '%" & UCase(MCGroup) & "%'))) As NetStockAdjustment " & _
                                                              "FROM Master1 M WHERE MasterType=6 AND Alias<>'' AND LEFT(Alias,4) IN (" & SelectedItems(ListView3, True) & ") ORDER BY Alias", CxnProductionPlanning, adOpenKeyset, adLockReadOnly
            
        End If
            rstProductionPlanning.ActiveConnection = Nothing
            Call UpdatePPFigures("2") 'Update Sale,Stock Transfer,Specimen, Stock & Pending Sales Order Figures
            MdiMainMenu.ProgressBar1.Value = MdiMainMenu.ProgressBar1.Value + 16.5
        End If
        If i = 0 Then Exit Do Else DatabaseName = Mid(DatabaseName, i + 1): i = 0
    Loop
    
    DatabaseName = Trim(ReadFromFile("Saral Database Name")): If DatabaseName = "" Then Exit Sub
    i = 0: K = 0
    Do While True
    
        i = InStr(1, DatabaseName, ",")
        If CxnProductionPlanning.State = adStateOpen Then CxnProductionPlanning.Close
        If i = 0 Then CxnProductionPlanning.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DatabasePath & "\" & Mid(DatabaseName, 1) & ";Persist Security Info=False;Jet OLEDB:Database Password=RSPLILoveMyINDIA" Else CxnProductionPlanning.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DatabasePath & "\" & Mid(DatabaseName, 1, i - 1) & ";Persist Security Info=False;Jet OLEDB:Database Password=RSPLILoveMyINDIA"
        K = K + 1
        If rstProductionPlanning.State = adStateOpen Then rstProductionPlanning.Close
        If K = 1 Then
            If OrderType = "M" Then
                FromDate = "01-Oct-" & Trim(Year(GetDate(MhDateInput1.Text))): ToDate = "31-Mar-" & Trim(Year(GetDate(MhDateInput1.Text)) + 1)
                rstProductionPlanning.Open "SELECT M.BusyCode,CLng(Sum(C.ActualQuantity)) As PrintOrder FROM (BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code) INNER JOIN BookMaster M ON M.Code=P.Book WHERE P.Type='F' AND LEFT(P.Code,1)<>'*' AND P.Date>=#" & FromDate & "# AND P.Date<=#" & ToDate & "# AND RIGHT(M.BusyCode,1)<>'S' AND LEFT(M.BusyCode,4) IN (" & SelectedItems(ListView3, True) & ") GROUP BY M.BusyCode ORDER BY M.BusyCode", CxnProductionPlanning, adOpenKeyset, adLockReadOnly
                rstProductionPlanning.ActiveConnection = Nothing
                Call UpdatePPFigures("3") 'Update Print Order Figures
            End If
            MdiMainMenu.ProgressBar1.Value = MdiMainMenu.ProgressBar1.Value + 16.5
        ElseIf K = 2 Then
            If OrderType = "M" Then
                
                FromDate = "01-Apr-" & Trim(Year(GetDate(MhDateInput1.Text)) + 1): ToDate = "30-Sep-" & Trim(Year(GetDate(MhDateInput1.Text)) + 1)
                rstProductionPlanning.Open "SELECT M.BusyCode,CLng(Sum(C.ActualQuantity)) As PrintOrder FROM (BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code) INNER JOIN BookMaster M ON M.Code=P.Book WHERE P.Type='F' AND LEFT(P.Code,1)<>'*' AND P.Date>=#" & FromDate & "# AND P.Date<=#" & ToDate & "# AND RIGHT(M.BusyCode,1)<>'S' AND LEFT(M.BusyCode,4) IN (" & SelectedItems(ListView3, True) & ") GROUP BY M.BusyCode ORDER BY M.BusyCode", CxnProductionPlanning, adOpenKeyset, adLockReadOnly
                rstProductionPlanning.ActiveConnection = Nothing
                Call UpdatePPFigures("4") 'Update Print Order Figures
                MdiMainMenu.ProgressBar1.Value = MdiMainMenu.ProgressBar1.Value + 16.5
                If rstProductionPlanning.State = adStateOpen Then rstProductionPlanning.Close
                FromDate = "01-Oct-" & Trim(Year(GetDate(MhDateInput1.Text)) + 1): ToDate = "31-Mar-" & Trim(Year(GetDate(MhDateInput1.Text)) + 2)
                rstProductionPlanning.Open "SELECT M.BusyCode,CLng(Sum(C.ActualQuantity)) As PrintOrder FROM (BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code) INNER JOIN BookMaster M ON M.Code=P.Book WHERE P.Type='F' AND LEFT(P.Code,1)<>'*' AND P.Date>=#" & FromDate & "# AND P.Date<=#" & ToDate & "# AND RIGHT(M.BusyCode,1)<>'S' AND LEFT(M.BusyCode,4) IN (" & SelectedItems(ListView3, True) & ") GROUP BY M.BusyCode ORDER BY M.BusyCode", CxnProductionPlanning, adOpenKeyset, adLockReadOnly
                rstProductionPlanning.ActiveConnection = Nothing
                Call UpdatePPFigures("5") 'Update Print Order Figures
            Else
                FromDate = "01-Apr-" & Trim(Year(GetDate(MhDateInput1.Text)) - 1)
                rstProductionPlanning.Open "SELECT M.BusyCode,CLng(IIF(ISNULL(SUM(C.ActualQuantity-IIF(C.Status IN ('D','E','W'),C.ActualQuantity,P.ReceivedQuantity))),0,SUM(C.ActualQuantity-IIF(C.Status IN ('D','E','W'),C.ActualQuantity,P.ReceivedQuantity))))+(SELECT CLng(IIF(ISNULL(SUM(C.ActualQuantity-IIF(C.Status IN ('D','E','W') OR BillNo<>'',C.ActualQuantity,P.ReceivedQuantity))),0,SUM(C.ActualQuantity-IIF(C.Status IN ('D','E','W') OR BillNo<>'',C.ActualQuantity,P.ReceivedQuantity)))) FROM BookPOChild08 C INNER JOIN BookPOParent P ON P.Code=C.Code WHERE P.Type='R' AND P.Book=M.Code AND LEFT(P.Code,1)<>'*' AND P.Date>=#" & FromDate & "# AND P.Date<=#" & GetDate(MhDateInput2.Text) & "#) As PendingPrintOrder FROM (BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code) INNER JOIN BookMaster M ON M.Code=P.Book " & _
                                                              "WHERE P.Type='F' AND LEFT(P.Code,1)<>'*' AND P.Date>=#" & FromDate & "# AND P.Date<=#" & GetDate(MhDateInput2.Text) & "# AND RIGHT(M.BusyCode,1)<>'S' AND LEFT(M.BusyCode,4) IN (" & SelectedItems(ListView3, True) & ") GROUP BY M.BusyCode,M.Code ORDER BY M.BusyCode", CxnProductionPlanning, adOpenKeyset, adLockReadOnly
                rstProductionPlanning.ActiveConnection = Nothing
                Call UpdatePPFigures("4") 'Update Pending Print Order Figures
                
            End If
            MdiMainMenu.ProgressBar1.Value = MdiMainMenu.ProgressBar1.Value + 16.5
            
        Else
            FromDate = "01-Apr-" & Trim(Year(GetDate(MhDateInput1.Text)) + IIf(OrderType = "M", 2, 0))
            rstProductionPlanning.Open "SELECT M.BusyCode,CLng(Sum(C.ActualQuantity)) As PrintOrder,CLng(IIF(ISNULL(SUM(C.ActualQuantity-IIF(C.Status IN ('D','E','W'),C.ActualQuantity,P.ReceivedQuantity))),0,SUM(C.ActualQuantity-IIF(C.Status IN ('D','E','W'),C.ActualQuantity,P.ReceivedQuantity))))+(SELECT CLng(IIF(ISNULL(SUM(C.ActualQuantity-IIF(C.Status IN ('D','E','W') OR BillNo<>'',C.ActualQuantity,P.ReceivedQuantity))),0,SUM(C.ActualQuantity-IIF(C.Status IN ('D','E','W') OR BillNo<>'',C.ActualQuantity,P.ReceivedQuantity)))) FROM BookPOChild08 C INNER JOIN BookPOParent P ON P.Code=C.Code WHERE P.Type='R' AND P.Book=M.Code AND LEFT(P.Code,1)<>'*' AND P.Date>=#" & FromDate & "# AND P.Date<=#" & GetDate(MhDateInput2.Text) & "#) As PendingPrintOrder FROM (BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code) INNER JOIN BookMaster M ON M.Code=P.Book " & _
                                                          "WHERE P.Type='F' AND LEFT(P.Code,1)<>'*' AND P.Date>=#" & FromDate & "# AND P.Date<=#" & GetDate(MhDateInput2.Text) & "# AND RIGHT(M.BusyCode,1)<>'S' AND LEFT(M.BusyCode,4) IN (" & SelectedItems(ListView3, True) & ") GROUP BY M.Code,M.BusyCode ORDER BY M.BusyCode", CxnProductionPlanning, adOpenKeyset, adLockReadOnly
            rstProductionPlanning.ActiveConnection = Nothing
            Call UpdatePPFigures("6") 'Update Print Order & Pending Print Order Figures
            MdiMainMenu.ProgressBar1.Value = MdiMainMenu.ProgressBar1.Value + 16.5
        End If
        If i = 0 Then Exit Do Else DatabaseName = Mid(DatabaseName, i + 1): i = 0
    Loop
    Call CloseRecordset(rstProductionPlanning)
    Call CloseConnection(CxnProductionPlanning)
    Screen.MousePointer = vbNormal
    On Error Resume Next
    If Not FileExist(App.Path & "\Template\Production Planning.xlsx") Then Exit Sub
    Screen.MousePointer = vbHourglass
    If rstProductionPlanning.State = adStateOpen Then rstProductionPlanning.Close
    If OrderType = "M" Then
        rstProductionPlanning.Open "SELECT Code,PrintName,BusyCode As Alias,POLTLY1003,POLY0409,POLY1003,POTY0409,SaleLY1003,SaleTY0409,StockTransferLY1003,StockTransferTY0409,SpecimenLY1003,SpecimenTY0409,ESO30 As CYReturn,PendingPO,SaleableStock,PendingSO,Remarks FROM BookMaster WHERE LEFT(BusyCode,4) IN (" & SelectedItems(ListView3, True) & ") AND Type='F' AND BusyCode<>'' AND RIGHT(BusyCode,1)<>'S' ORDER BY PrintName", CxnDatabase, adOpenKeyset, adLockReadOnly
    Else
        rstProductionPlanning.Open "SELECT Code,PrintName,BusyCode As Alias,POTY0409,PendingPO,SaleableStock,PendingSO,ESO30,ESO60,ESO90,ESO150,PSO15,PSO30,SaleLY1003 As LYSale,SaleTY0409 As CYSale FROM BookMaster WHERE Type='F' AND LEFT(BusyCode,4) IN (" & SelectedItems(ListView3, True) & ") AND BusyCode<>'' AND RIGHT(BusyCode,1)<>'S' ORDER BY PrintName", CxnDatabase, adOpenKeyset, adLockReadOnly
    End If
    
    
    If rstProductionPlanning.RecordCount = 0 Then
        DisplayError ("No Record Found")
        ShowProgressInStatusBar False
        Screen.MousePointer = vbNormal
        On Error GoTo 0
        Exit Sub
    End If
    
    DoEvents
        
    'Writing To Excel

    Set oExcel = CreateObject("Excel.Application")
    oExcel.Workbooks.Open (App.Path & "\Template\Production Planning"): oExcel.DisplayAlerts = False
    oExcel.Workbooks.Item(1).SaveAs (App.Path & "\Report\Production Planning (" & CompCode & ")"): oExcel.DisplayAlerts = True
    oExcel.Sheets("Reorder Level Register").Visible = False: oExcel.Sheets("Production Planning (" & IIf(OrderType = "M", "SO", "MO") & ")").Visible = False: oExcel.Sheets("Production Planning (" & IIf(OrderType = "M", "MO", "SO") & ")").Select: oExcel.Visible = False
    oExcel.Cells(1, "A").Value = Trim(rstCompanyMaster.Fields("PrintName").Value)
    oExcel.Cells(2, "A").Value = "Production Planning (" & IIf(OrderType = "M", "Main", "Supplement") & " Orders) As On [" & Format(GetDate(MhDateInput2.Text), "dd-MMM-yyyy") & "]"
    
    If OrderType = "M" Then
        Period01 = "(" + Right(Year(GetDate(MhDateInput1.Text)), 2) + "-" + Right(Year(GetDate(MhDateInput1.Text)) + 1, 2) + ")"
        Period02 = "(" + Right(Year(GetDate(MhDateInput1.Text)) + 1, 2) + "-" + Right(Year(GetDate(MhDateInput1.Text)) + 2, 2) + ")"
        Period03 = "(" + Right(Year(GetDate(MhDateInput1.Text)) + 2, 2) + "-" + Right(Year(GetDate(MhDateInput1.Text)) + 3, 2) + ")"
        oExcel.Cells(5, "D").Value = Period01
        oExcel.Cells(5, "E").Value = Period02
        oExcel.Cells(5, "G").Value = Period03
        
        If Month(GetDate(MhDateInput1.Text)) = 10 Then
            oExcel.Cells(4, "H").Value = "Oct-Mar"
            oExcel.Cells(4, "I").Value = "Apr-Sep"
            oExcel.Cells(5, "H").Value = Period02
            oExcel.Cells(5, "I").Value = Period03
        Else
            oExcel.Cells(4, "H").Value = "Apr-Sep"
            oExcel.Cells(4, "I").Value = "Oct-Mar"
            oExcel.Cells(5, "H").Value = Period02
            oExcel.Cells(5, "I").Value = Period02
        End If
    End If
    
    i = IIf(OrderType = "M", 7, 5): Cnt = 1
    Do While Not rstProductionPlanning.EOF
        oExcel.Cells(i, "A").Value = Cnt
        oExcel.Application.Cells(i, "B").Value = Trim(rstProductionPlanning.Fields("PrintName").Value)
        oExcel.Application.Cells(i, "C").Value = Trim(rstProductionPlanning.Fields("Alias").Value)
        If OrderType = "M" Then
            
            'Print Order
            oExcel.Application.Cells(i, "D").Value = Val(rstProductionPlanning.Fields("POLTLY1003").Value)
            oExcel.Application.Cells(i, "E").Value = Val(rstProductionPlanning.Fields("POLY0409").Value)
            oExcel.Application.Cells(i, "F").Value = Val(rstProductionPlanning.Fields("POLY1003").Value)
            oExcel.Application.Cells(i, "G").Value = Val(rstProductionPlanning.Fields("POTY0409").Value)
            'Sale
            oExcel.Application.Cells(i, "H").Value = Val(rstProductionPlanning.Fields("SaleLY1003").Value)
            
            oExcel.Application.Cells(i, "I").Value = Val(rstProductionPlanning.Fields("SaleTY0409").Value)
            'Branch Transfer
            oExcel.Application.Cells(i, "J").Value = Val(rstProductionPlanning.Fields("StockTransferLY1003").Value)
            oExcel.Application.Cells(i, "K").Value = Val(rstProductionPlanning.Fields("StockTransferTY0409").Value)
            'Specimen
            oExcel.Application.Cells(i, "L").Value = Val(rstProductionPlanning.Fields("SpecimenLY1003").Value)
            oExcel.Application.Cells(i, "M").Value = Val(rstProductionPlanning.Fields("SpecimenTY0409").Value)
            'Current Return
            oExcel.Application.Cells(i, "N").Value = Val(rstProductionPlanning.Fields("CYReturn").Value)
            'Pending Print Order
            oExcel.Application.Cells(i, "O").Value = Val(rstProductionPlanning.Fields("PendingPO").Value)
            oExcel.Application.Cells(i, "P").Value = Val(rstProductionPlanning.Fields("SaleableStock").Value)
            oExcel.Application.Cells(i, "Q").Value = Val(rstProductionPlanning.Fields("PendingSO").Value)
            If i > 7 Then oExcel.Range("R" & Trim(i)).FormulaR1C1 = oExcel.Range("R7").FormulaR1C1
            If Val(oExcel.Application.Cells(i, "R")) < 0 Then oExcel.Application.Cells(i, "R").Value = 0
            oExcel.Application.Cells(i, "S").Value = rstProductionPlanning.Fields("Remarks").Value
            oExcel.Application.Cells(i, "XFD").Value = rstProductionPlanning.Fields("Code").Value
        Else
            oExcel.Application.Cells(i, "D").Value = Val(rstProductionPlanning.Fields("POTY0409").Value)
            oExcel.Application.Cells(i, "E").Value = Val(rstProductionPlanning.Fields("SaleableStock").Value)
            oExcel.Application.Cells(i, "F").Value = Val(rstProductionPlanning.Fields("PendingPO").Value)
            oExcel.Application.Cells(i, "G").Value = Val(rstProductionPlanning.Fields("PendingSO").Value)
            oExcel.Application.Cells(i, "H").Value = Val(oExcel.Application.Cells(i, "E")) + Val(oExcel.Application.Cells(i, "F")) - Val(oExcel.Application.Cells(i, "G"))
            oExcel.Application.Cells(i, "I").Value = Val(rstProductionPlanning.Fields("LYSale").Value)
            oExcel.Application.Cells(i, "J").Value = Val(rstProductionPlanning.Fields("CYSale").Value)
            oExcel.Application.Cells(i, "K").Value = Val(rstProductionPlanning.Fields("ESO30").Value)
            oExcel.Application.Cells(i, "L").Value = Val(rstProductionPlanning.Fields("ESO60").Value)
            oExcel.Application.Cells(i, "M").Value = Val(rstProductionPlanning.Fields("ESO90").Value)
            oExcel.Application.Cells(i, "N").Value = Val(rstProductionPlanning.Fields("ESO150").Value)
            oExcel.Application.Cells(i, "O").Value = Val(rstProductionPlanning.Fields("PSO15").Value)
            oExcel.Application.Cells(i, "P").Value = Val(rstProductionPlanning.Fields("PSO30").Value)
        End If
        Cnt = Cnt + 1: i = i + 1
        rstProductionPlanning.MoveNext
    Loop
    oExcel.Columns("A:B").EntireColumn.AutoFit
    oExcel.Workbooks.Item(1).Save
    Screen.MousePointer = vbNormal
    MdiMainMenu.ProgressBar1.Value = 100
    If OutputTo = "S" Then oExcel.Range("A1").Activate: oExcel.Visible = True Else oExcel.Workbooks.Item(1).PrintOut
    ShowProgressInStatusBar False
    Set oExcel = Nothing
    On Error GoTo 0
    Exit Sub
ErrorHandler:
    Screen.MousePointer = vbNormal
    DisplayError ("Failed to update Production Planning figures")
    ShowProgressInStatusBar False
    Call CloseRecordset(rstProductionPlanning)
    Call CloseConnection(CxnProductionPlanning)
    Call CloseConnection(CxnDistributor)
    
End Sub


Private Sub UpdatePPFigures(ByVal UpdationType As String)
    If rstProductionPlanning.RecordCount > 0 Then rstProductionPlanning.MoveFirst
  
    Do While Not rstProductionPlanning.EOF
        
        If UpdationType = "1" Then
            If OrderType = "M" Then
                CxnDatabase.Execute "UPDATE BookMaster SET SaleLY1003=SaleLY1003+" & Val(rstProductionPlanning.Fields("Sale01").Value) & ",StockTransferLY1003=StockTransferLY1003+" & Val(rstProductionPlanning.Fields("BranchStockTransfer01").Value) & ",SpecimenLY1003=SpecimenLY1003+" & Val(rstProductionPlanning.Fields("Specimen01").Value) & ",SaleTY0409=SaleTY0409+" & Val(rstProductionPlanning.Fields("Sale02").Value) & ",StockTransferTY0409=StockTransferTY0409+" & Val(rstProductionPlanning.Fields("BranchStockTransfer02").Value) & ",SpecimenTY0409=SpecimenTY0409+" & Val(rstProductionPlanning.Fields("Specimen02").Value) & " WHERE LEFT(BusyCode,4)='" & Left(rstProductionPlanning.Fields("Alias").Value, 4) & "'"
            Else
               CxnDatabase.Execute "UPDATE BookMaster SET ESO30=ESO30+" & Val(rstProductionPlanning.Fields("Sale30").Value) + Val(rstProductionPlanning.Fields("BranchStockTransfer30").Value) + Val(rstProductionPlanning.Fields("Specimen30").Value) & ",ESO60=ESO60+" & Val(rstProductionPlanning.Fields("Sale60").Value) + Val(rstProductionPlanning.Fields("BranchStockTransfer60").Value) + Val(rstProductionPlanning.Fields("Specimen60").Value) & ",ESO90=ESO90+" & Val(rstProductionPlanning.Fields("Sale90").Value) + Val(rstProductionPlanning.Fields("BranchStockTransfer90").Value) + Val(rstProductionPlanning.Fields("Specimen90").Value) & ",ESO150=ESO150+" & Val(rstProductionPlanning.Fields("Sale150").Value) + Val(rstProductionPlanning.Fields("BranchStockTransfer150").Value) + Val(rstProductionPlanning.Fields("Specimen150").Value) & "," & _
                                   "PSO15=PSO15+" & Val(rstProductionPlanning.Fields("CSale15").Value) + Val(rstProductionPlanning.Fields("CBranchStockTransfer15").Value) + Val(rstProductionPlanning.Fields("CSpecimen15").Value) & ",PSO30=PSO30+" & Val(rstProductionPlanning.Fields("CSale30").Value) + Val(rstProductionPlanning.Fields("CBranchStockTransfer30").Value) + Val(rstProductionPlanning.Fields("CSpecimen30").Value) & ",SaleLY1003=SaleLY1003+" & Val(rstProductionPlanning.Fields("LYSale").Value) + Val(rstProductionPlanning.Fields("LYBranchStockTransfer").Value) + Val(rstProductionPlanning.Fields("LYSpecimen").Value) & " WHERE LEFT(BusyCode,4)='" & Left(rstProductionPlanning.Fields("Alias").Value, 4) & "'"
           End If
        ElseIf UpdationType = "2" Then
            If OrderType = "M" Then
                If Month(GetDate(MhDateInput1.Text)) = 10 Then CxnDatabase.Execute "UPDATE BookMaster SET SaleTY0409=SaleTY0409+" & Val(rstProductionPlanning.Fields("Sale").Value) & ",StockTransferTY0409=StockTransferTY0409+" & Val(rstProductionPlanning.Fields("BranchStockTransfer").Value) & ",SpecimenTY0409=SpecimenTY0409+" & Val(rstProductionPlanning.Fields("Specimen").Value) & " WHERE LEFT(BusyCode,4)='" & Left(rstProductionPlanning.Fields("Alias").Value, 4) & "'"
                CxnDatabase.Execute "UPDATE BookMaster SET ESO30=ESO30+" & Val(rstProductionPlanning.Fields("BranchStockTransferReturn").Value) + Val(rstProductionPlanning.Fields("SaleReturn").Value) & " WHERE LEFT(BusyCode,4)='" & Left(rstProductionPlanning.Fields("Alias").Value, 4) & "'"  'Current Return
            Else
                CxnDatabase.Execute "UPDATE BookMaster SET ESO30=ESO30+" & Val(rstProductionPlanning.Fields("Sale30").Value) + Val(rstProductionPlanning.Fields("BranchStockTransfer30").Value) + Val(rstProductionPlanning.Fields("Specimen30").Value) & ",ESO60=ESO60+" & Val(rstProductionPlanning.Fields("Sale60").Value) + Val(rstProductionPlanning.Fields("BranchStockTransfer60").Value) + Val(rstProductionPlanning.Fields("Specimen60").Value) & ",ESO90=ESO90+" & Val(rstProductionPlanning.Fields("Sale90").Value) + Val(rstProductionPlanning.Fields("BranchStockTransfer90").Value) + Val(rstProductionPlanning.Fields("Specimen90").Value) & ",ESO150=ESO150+" & Val(rstProductionPlanning.Fields("Sale150").Value) + Val(rstProductionPlanning.Fields("BranchStockTransfer150").Value) + Val(rstProductionPlanning.Fields("Specimen150").Value) & "," & _
                                                    "PSO15=PSO15+" & Val(rstProductionPlanning.Fields("CSale15").Value) + Val(rstProductionPlanning.Fields("CBranchStockTransfer15").Value) + Val(rstProductionPlanning.Fields("CSpecimen15").Value) & ",PSO30=PSO30+" & Val(rstProductionPlanning.Fields("CSale30").Value) + Val(rstProductionPlanning.Fields("CBranchStockTransfer30").Value) + Val(rstProductionPlanning.Fields("CSpecimen30").Value) & ",SaleTY0409=SaleTY0409+" & Val(rstProductionPlanning.Fields("CYSale").Value) + Val(rstProductionPlanning.Fields("CYBranchStockTransfer").Value) + Val(rstProductionPlanning.Fields("CYSpecimen").Value) & " WHERE LEFT(BusyCode,4)='" & Left(rstProductionPlanning.Fields("Alias").Value, 4) & "'"
            End If
            If StrConv(Mid(rstProductionPlanning.Fields("Alias").Value, 6, 1), vbUpperCase) <> "Z" Then CxnDatabase.Execute "UPDATE BookMaster SET PendingSO=PendingSO+" & Val(rstProductionPlanning.Fields("SaleOrder").Value) - Val(rstProductionPlanning.Fields("SaleOrderSupplied").Value) & ",SaleableStock=SaleableStock+" & Val(rstProductionPlanning.Fields("OpBal").Value) - Val(rstProductionPlanning.Fields("NetSale").Value) + Val(rstProductionPlanning.Fields("NetStockTransfer").Value) + Val(rstProductionPlanning.Fields("NetPurchase").Value) + Val(rstProductionPlanning.Fields("NetStockAdjustment").Value) & " WHERE LEFT(BusyCode,4)='" & Left(rstProductionPlanning.Fields("Alias").Value, 4) & "'"
        ElseIf UpdationType = "3" Then
            If OrderType = "M" Then CxnDatabase.Execute "UPDATE BookMaster SET POLTLY1003=POLTLY1003+" & Val(CheckNull(rstProductionPlanning.Fields("PrintOrder").Value)) & " WHERE LEFT(BusyCode,4)='" & Left(rstProductionPlanning.Fields("BusyCode").Value, 4) & "'"
        ElseIf UpdationType = "4" Then
            If OrderType = "M" Then CxnDatabase.Execute "UPDATE BookMaster SET POLY0409=POLY0409+" & Val(CheckNull(rstProductionPlanning.Fields("PrintOrder").Value)) & " WHERE LEFT(BusyCode,4)='" & Left(rstProductionPlanning.Fields("BusyCode").Value, 4) & "'" Else CxnDatabase.Execute "UPDATE BookMaster SET PendingPO=PendingPO+" & Val(CheckNull(rstProductionPlanning.Fields("PendingPrintOrder").Value)) & " WHERE LEFT(BusyCode,4)='" & Left(rstProductionPlanning.Fields("BusyCode").Value, 4) & "'"
         ElseIf UpdationType = "5" Then
            If OrderType = "M" Then CxnDatabase.Execute "UPDATE BookMaster SET POLY1003=POLY1003+" & Val(CheckNull(rstProductionPlanning.Fields("PrintOrder").Value)) & " WHERE LEFT(BusyCode,4)='" & Left(rstProductionPlanning.Fields("BusyCode").Value, 4) & "'"
        ElseIf UpdationType = "6" Then
            CxnDatabase.Execute "UPDATE BookMaster SET POTY0409=POTY0409+" & Val(CheckNull(rstProductionPlanning.Fields("PrintOrder").Value)) & ",PendingPO=PendingPO+" & Val(CheckNull(rstProductionPlanning.Fields("PendingPrintOrder").Value)) & " WHERE LEFT(BusyCode,4)='" & Left(rstProductionPlanning.Fields("BusyCode").Value, 4) & "'"
        End If
        rstProductionPlanning.MoveNext
    Loop
End Sub

Private Sub PrintProductionPlanningS()
    Dim DatabaseName As String
    Dim DatabaseNameDistributor As String
    Dim FromDate As String, ToDate As String, FromDate02 As String, ToDate02 As String
    Dim oExcel As Object
    Dim i As Long, Cnt As Long
    Dim j As Long
    Dim Period01 As String, Period02 As String, Period03 As String
    On Error GoTo ErrorHandler
    DoEvents
    ShowProgressInStatusBar True
    MdiMainMenu.ProgressBar1.Value = 1
    DatabaseName = Trim(ReadFromFile("Busy Database Name"))
    DatabaseNameDistributor = Trim(ReadFromFile("Busy Distributor Database Name"))
    If ServerName = "" Or DatabaseName = "" Then Exit Sub
    If ServerName = "" Or DatabaseNameDistributor = "" Then Exit Sub
    Screen.MousePointer = vbHourglass
    CxnDatabase.Execute "UPDATE BookMaster SET SaleLY1003=0,SaleTY0409=0,StockTransferLY1003=0,StockTransferTY0409=0,SpecimenLY1003=0,SpecimenTY0409=0,PendingSO=0,SaleableStock=0,POLTLY1003=0,POLY0409=0,POLY1003=0,POTY0409=0,PendingPO=0,ESO30=0,ESO60=0,ESO90=0,ESO150=0,PSO15=0,PSO30=0"
    i = 0: K = 0: j = 0
    CxnProductionPlanning.CursorLocation = adUseClient: CxnProductionPlanning.CommandTimeout = 0
    CxnDistributor.CursorLocation = adUseClient: CxnDistributor.CommandTimeout = 0
    
    Do While True
        i = InStr(1, DatabaseName, ",")
        If CxnProductionPlanning.State = adStateOpen Then CxnProductionPlanning.Close
        If i = 0 Then CxnProductionPlanning.Open "Provider=SQLOLEDB.1;Password=" & ServerPassword & ";Persist Security Info=True;User ID=sa;Initial Catalog=" & Mid(DatabaseName, 1) & ";Data Source=" & ServerName Else CxnProductionPlanning.Open "Provider=SQLOLEDB.1;Password=" & ServerPassword & ";Persist Security Info=True;User ID=sa;Initial Catalog=" & Mid(DatabaseName, 1, i - 1) & ";Data Source=" & ServerName

        j = InStr(1, DatabaseNameDistributor, ",")
        If CxnDistributor.State = adStateOpen Then CxnDistributor.Close
        If j = 0 Then CxnDistributor.Open "Provider=SQLOLEDB.1;Password=" & ServerPassword & ";Persist Security Info=True;User ID=sa;Initial Catalog=" & Mid(DatabaseNameDistributor, 1) & ";Data Source=" & ServerName Else CxnDistributor.Open "Provider=SQLOLEDB.1;Password=" & ServerPassword & ";Persist Security Info=True;User ID=sa;Initial Catalog=" & Mid(DatabaseNameDistributor, 1, i - 1) & ";Data Source=" & ServerName
          Dim StrQry1 As String 'Main Qry
          Dim StrQry2 As String 'Second Qry
        K = K + 1
        If rstProductionPlanning.State = adStateOpen Then rstProductionPlanning.Close
        If K = 1 Then
        
        Else
            
            'Current Year Data Processing
            Dim LYDatabaseName As String
            Dim DLYDatabaseName As String
            LYDatabaseName = Mid(DatabaseName, 1, 16) & Val(Right(DatabaseName, 4)) - 1
            DLYDatabaseName = Mid(DatabaseNameDistributor, 1, 16) & Val(Right(DatabaseNameDistributor, 4)) - 1
            FromDate = Trim(Day(GetDate(MhDateInput2.Text))) + "-" + MonthName(Trim(Month(GetDate(MhDateInput2.Text))), True) + "-" + Trim(Year(GetDate(MhDateInput2.Text)) - 1)
            If Not IsDate(FromDate) Then FromDate = Trim(Day(GetDate(MhDateInput2.Text)) - 1) + "-" + MonthName(Trim(Month(GetDate(MhDateInput2.Text))), True) + "-" + Trim(Year(GetDate(MhDateInput2.Text)) - 1)
               
               Dim z As Integer
               Dim StrBusyCode As String
               Dim strBusyCode2 As Integer
               DoEvents
               
               For z = 1 To ListView3.ListItems.Count
                   If ListView3.ListItems(z).Checked Then
                      
                            StrBusyCode = ListView3.ListItems.Item(z).SubItems(1)
                            strBusyCode2 = ListView3.ListItems.Item(z).SubItems(1)
                            CxnDatabase.Execute "UPDATE BookMaster SET PrintedLYear=0,PrintedCYear=0,CurrentStock=0,Pending_PO=0,Pending_SO=0,EffectiveStock=0,BlrStock=0,DistbrStock=0,SpecimenLyear=0,SpecimenCYear=0,SaleLYear=0,DistbrLYear=0,SaleCYear=0,DistbrCYear=0,ESale30=0,ESale60=0,SalePDays15=0,SalePDays30=0,FinalOrder=0 WHERE BusyCode='" & StrBusyCode & "'"
                           '*Rachna Stock***
                           Dim SQLQuery As String
'                           SQLQuery = "SELECT M.Mc AS McName,M.CM as Code,M.BSStock as B1,M.PG as ParentGrp,ISNULL(S1.MTB,0) AS MainTransBal," & _
'                                " (select ISNULL(Sum(d1),0) from tran4 where tran4.mastercode1=(Select MAX(code) From Master1 Where Alias Like'" & StrBusyCode & "-N" & "%') and  tran4.mastercode2= M.Cm) AS MainOpBal," & _
'                                " (select ISNULL(Sum(d3),0) from tran4 where tran4.mastercode1=(Select MAX(code) From Master1 Where Alias Like'" & StrBusyCode & "-N" & "%') and  tran4.mastercode2= M.Cm) AS OPAmt " & _
'                                " From (Select Name as Mc,Code as CM,B1 as BSStock,ParentGrp as PG from Master1 where  CODE IN (26548)) AS M Left Join (SELECT Mastercode2, sum(value1) AS MTB,sum(value2) AS ATB " & _
'                                " From tran2 Where rectype = 2 And Mastercode1=(Select MAX(code) From Master1 Where Alias Like'" & StrBusyCode & "-N" & "%') and  Date <= '" & GetDate(MhDateInput2.Text) & "' group by Mastercode2 ) AS " & _
'                                " S1 ON (S1.Mastercode2 = M.CM)  Where M.CM IN (SELECT CODE FROM HELP1 WHERE MASTERSERIES = 26841 And MasterType= 11) ORDER BY M.Mc"
                           
                           SQLQuery = "SELECT M.Mc AS McName,M.CM as Code,M.BSStock as B1,M.PG as ParentGrp,ISNULL(S1.MTB,0) AS MainTransBal," & _
                                " (select ISNULL(Sum(d1),0) from tran4 where tran4.mastercode1=(Select MAX(code) From Master1 Where Alias Like'" & StrBusyCode & "-N" & "%'  )  and tran4.mastercode2= M.Cm) AS MainOpBal," & _
                                " (select ISNULL(Sum(d3),0) from tran4 where tran4.mastercode1=(Select MAX(code) From Master1 Where Alias Like'" & StrBusyCode & "-N" & "%'  )  and  tran4.mastercode2= M.Cm) AS OPAmt " & _
                                " From (Select Name as Mc,Code as CM,B1 as BSStock,ParentGrp as PG from Master1 where Mastertype=11 And CODE IN ((Select Code From Master1 Where " & _
                                " ParentGrp in(Select Code From Master1 Where Name='Head Office (Group)'))))  AS M Left Join (SELECT Mastercode2, sum(value1) AS MTB,sum(value2) AS ATB " & _
                                " From tran2 Where rectype = 2 And Mastercode1=(Select MAX(code) From Master1 Where Alias Like'" & StrBusyCode & "-N" & "%'  )  and  Date <= '" & GetDate(MhDateInput2.Text) & "' group by Mastercode2 ) AS " & _
                                " S1 ON (S1.Mastercode2 = M.CM)  Where M.CM IN (SELECT CODE FROM HELP1 WHERE MASTERSERIES = 26841 And MasterType= 11) ORDER BY M.Mc"
                              If rstProductionPlanning.State = adStateOpen Then rstProductionPlanning.Close
                            rstProductionPlanning.Open SQLQuery, CxnProductionPlanning, adOpenKeyset, adLockReadOnly
                            rstProductionPlanning.ActiveConnection = Nothing
                            Call UpdateStockFigures("C", StrBusyCode) 'Update Stock
                            
                            '*Specimen***
                            If rstProductionPlanning.State = adStateOpen Then rstProductionPlanning.Close
                            Dim str_qry2 As String
                            str_qry2 = "SELECT M.Mc AS McName,M.CM as Code,M.BSStock as B1,M.PG as ParentGrp,ISNULL(S1.MTB,0) AS MainTransBal, ISNULL(S1.ATB,0) AS AltTransBal,"
                            str_qry2 = str_qry2 & " (select ISNULL(Sum(d1),0) from tran4 where tran4.mastercode1=(Select MAX(code) from Master1 Where  MasterType=6 and Alias<>'' and Alias Like'" & StrBusyCode & "-N" & "%'  ) and  tran4.mastercode2= M.Cm) AS MainOpBal,"
                            str_qry2 = str_qry2 & " (select ISNULL(Sum(d2),0) from tran4 where tran4.mastercode1=(Select MAX(code) from Master1 Where  MasterType=6 and Alias<>'' and Alias Like'" & StrBusyCode & "-N" & "%'  ) and  tran4.mastercode2= M.Cm) AS AltOPBal,"
                            str_qry2 = str_qry2 & " (select ISNULL(Sum(d3),0) from tran4 where tran4.mastercode1=(Select MAX(code) from Master1 Where  MasterType=6 and Alias<>'' and Alias Like'" & StrBusyCode & "-N" & "%'  ) and  tran4.mastercode2= M.Cm) AS OPAmt"
                            str_qry2 = str_qry2 & " From (Select Name as Mc,Code as CM,B1 as BSStock,ParentGrp as PG from Master1 where Mastertype=11 And CODE IN (Select Code From Master1 Where ParentGrp in(Select Code From Master1"
                            str_qry2 = str_qry2 & " Where Name='Specimen (Group)')))  AS M Left Join (SELECT Mastercode2, sum(value1) AS MTB,sum(value2) AS ATB From tran2 Where rectype = 2 and Mastercode1=(Select MAX(code) from Master1 Where"
                            str_qry2 = str_qry2 & " MasterType=6 and Alias<>'' and Alias Like'" & StrBusyCode & "-N" & "%'  )  and  Date <= '" & GetDate(MhDateInput2.Text) & "' group by Mastercode2 ) AS S1 ON (S1.Mastercode2 = M.CM)  ORDER BY M.Mc"
                            rstProductionPlanning.Open str_qry2, CxnProductionPlanning, adOpenKeyset, adLockReadOnly
                            rstProductionPlanning.ActiveConnection = Nothing
                            Call UpdateStockFigures("SPCY", StrBusyCode) 'Specimen Current Year



                            If rstProductionPlanning.State = adStateOpen Then rstProductionPlanning.Close
                            Dim str_qry3 As String
                            str_qry3 = "SELECT M.Mc AS McName,M.CM as Code,M.BSStock as B1,M.PG as ParentGrp,ISNULL(S1.MTB,0) AS MainTransBal, ISNULL(S1.ATB,0) AS AltTransBal,"
                            str_qry3 = str_qry3 & " (select ISNULL(Sum(d1),0) from " & LYDatabaseName & "..tran4 where tran4.mastercode1 In(Select Code from Master1 Where  MasterType=6 and Alias<>'' and Alias Like'" & StrBusyCode & "-N" & "%' )  and  tran4.mastercode2= M.Cm) AS MainOpBal,"
                            str_qry3 = str_qry3 & " (select ISNULL(Sum(d2),0) from " & LYDatabaseName & "..tran4 where tran4.mastercode1 In(Select Code from Master1 Where  MasterType=6 and Alias<>'' and Alias Like'" & StrBusyCode & "-N" & "%' )  and  tran4.mastercode2= M.Cm) AS AltOPBal,"
                            str_qry3 = str_qry3 & " (select ISNULL(Sum(d3),0) from " & LYDatabaseName & "..tran4 where tran4.mastercode1 In(Select Code from Master1 Where  MasterType=6 and Alias<>'' and Alias Like'" & StrBusyCode & "-N" & "%') and  tran4.mastercode2= M.Cm) AS OPAmt"
                            str_qry3 = str_qry3 & " From (Select Name as Mc,Code as CM,B1 as BSStock,ParentGrp as PG from Master1 where Mastertype=11 And CODE IN (Select Code From Master1 Where ParentGrp in(Select Code From Master1"
                            str_qry3 = str_qry3 & " Where Name='Specimen (Group)')))  AS M Left Join (SELECT Mastercode2, sum(value1) AS MTB,sum(value2) AS ATB From " & LYDatabaseName & "..tran2 Where rectype = 2 and Mastercode1 IN(Select Code from Master1 Where"
                            str_qry3 = str_qry3 & " MasterType=6 and Alias<>'' and Alias Like'" & StrBusyCode & "-N" & "%'  ) and  DATE >= '" & Format(CDate("01-Apr-" & Right(DatabaseName, 4) - 1), "dd-MMM-yyyy") & "' AND DATE <= '" & Format(CDate("31-Mar-" & Right(DatabaseName, 4)), "dd-MMM-yyyy") & "' group by Mastercode2 ) AS S1 ON (S1.Mastercode2 = M.CM)  ORDER BY M.Mc"
                            rstProductionPlanning.Open str_qry3, CxnProductionPlanning, adOpenKeyset, adLockReadOnly
                            rstProductionPlanning.ActiveConnection = Nothing
                            Call UpdateStockFigures("SPLY", StrBusyCode) 'Specimen Last Year
                          'Call LoadPartyGroup
                           Dim strPartyGroup As String
                           
                          
                            strPartyGroup = "36650,21491,4074,4154,4147,18158,29123,29124,1035,13492,17990,6732,43651,43796,32119,15668,"
                            strPartyGroup = strPartyGroup & "18118,3016,16169,21634,3332,18119,18120,26988,10432,35119,4286,4398,36344,36345,18117,30252,16190,16191,40067,1007,1009,1010,1012,1013,1014,1016,"
                            strPartyGroup = strPartyGroup & "1017,1018,1020,1021,1023,1025,1027,1029,1030,1038,22280,1041,11630,11951,6891,15667,11631,11632,11616,11617,22227,18134,18135,18138,18142,18139,"
                            strPartyGroup = strPartyGroup & "18140,18137,18141,18148,18149,18147,4190,4284,4313,18061,21620,21617,21618,21621,21622,21623,21612,21631,21626,21627,21628,21629,21630,21633,21613,"
                            strPartyGroup = strPartyGroup & "21614,21615,21608,21610,21611,21619,21624,21609,21616,21625,21632,21916,21917,21918,21919,21903,21607,18074,18073,18076,18075,18082,18083,18084,36995,"
                            strPartyGroup = strPartyGroup & "18088,18090,40977,4191,4053,18091,18093,18092,18096,18107,18108,18111,35462,18115,32750,18113,18125,18128,4210,18126,18127,18129,18130,18454,18106,18122,"
                            strPartyGroup = strPartyGroup & "18123,4160,4287,5186,12857,18124,18069,4124,18072,18070,18132,4212,4093,4102,4273,18050,18131,18133,4783,4050,4157,4171,4302,37483,18059,37"
                            strPartyGroup = strPartyGroup & "776,13484,4051,38646,4227,12852,37547,12862,18116,42432,42433,42434,31936,18121,4165,4121,4285,12855,12858,12864,4156,32542,4137,42511,3291,18097,18098,"
                            strPartyGroup = strPartyGroup & "18099,1022,12847,4233,4279,38686,12865,24078,24079,4228,4237,12860,4107,4096,4068,42383,4188,4283,4364,18062,35120,4221,4179,4060,4086,4260,18104,18102,"
                            strPartyGroup = strPartyGroup & "18103,38000,29273,29274,29275,29276,18105,18156,18155,18068,18067,4194,4197,4058,18063,18064,4168,18065,18066,4353,18058,4218,4258,4346,4318,4103,4184,18071,"
                            strPartyGroup = strPartyGroup & "4370,4189,4097,4288,4158,18077,4172,4281,18078,4174,18080,18079,18081,18085,39068,4235,18086,4270,4239,18087,37144,37157,4098,18089,18052,18055,4319,40976,4177,"
                            strPartyGroup = strPartyGroup & "18094,4214,1011,4303,18100,18101,11975,4114,4152,4057,4072,42430,4056,4079,18110,18109,4196,18112,42428,4220,18114,6435,43192,42426,42427,4342,18162,18163,18164,"
                            strPartyGroup = strPartyGroup & "18165,4307,18160,18161,18166,18159,4170,18136,4354,4373,4092,38665,18143,18047,4146,18144,18145,4366,4261,18060,18146,18048,18152,4052,18150,18153,18151,5758,4182,4274,"
                            strPartyGroup = strPartyGroup & "18154,4324,43697,21646,21635,21636,21637,21638,21645,21639,21643,21641,21"
                            strPartyGroup = strPartyGroup & "642,21648,21640,21644,21647,4335,18095,40978"

                                                       
                           
                           '*Sale Details
                            If rstProductionPlanning.State = adStateOpen Then rstProductionPlanning.Close
                            Dim str_qry As String
                            
'***********Old Code***************************************************
                            'str_qry = " SELECT CODE,Name,"
'                            str_qry = str_qry & " ((Select ISNULL(Sum(abs(Value1)),0) From Tran2 Where Tran2.CM1=Master1.Code And  (RECTYPE = 2 OR RECTYPE = 7) AND VCHTYPE = 9 AND MasterCode1 = (Select Code from Master1 Where  MasterType=6 and Alias<>'' AND Alias Like'" & StrBusyCode & "-N" & "%') AND DATE >= '" & GetDate(MhDateInput1.Text) & "' AND DATE <= '" & GetDate(MhDateInput2.Text) & "' AND IsReturnQty=0)-(Select ISNULL(Sum(abs(Value1)),0) From Tran2 Where Tran2.CM1=Master1.Code And  (RECTYPE = 2 OR RECTYPE = 7) AND VCHTYPE = 3 AND MasterCode1 = (Select Code from Master1 Where  MasterType=6 and Alias<>'' AND Alias Like'" & StrBusyCode & "-N" & "%') AND DATE >= '" & GetDate(MhDateInput1.Text) & "' AND DATE <= '" & GetDate(MhDateInput2.Text) & "' AND IsReturnQty=0)) as SaleQty,"
'                            str_qry = str_qry & " ((Select ISNULL(Sum(abs(Value1)),0) From " & LYDatabaseName & "..Tran2 T1 Where T1.CM1=Master1.Code And  (RECTYPE = 2 OR RECTYPE = 7) AND VCHTYPE = 9 AND MasterCode1 IN (Select Code from Master1 Where  MasterType=6 and Alias<>'' and (Alias Like'" & StrBusyCode & "%' or Alias Like'" & strBusyCode2 & "%')) AND DATE >= '" & Format(CDate("01-Apr-" & Right(DatabaseName, 4) - 1), "dd-MMM-yyyy") & "' AND DATE <= '" & Format(CDate("31-Mar-" & Right(DatabaseName, 4)), "dd-MMM-yyyy") & "' And IsReturnQty=0)-"
'                            str_qry = str_qry & " (Select ISNULL(Sum(abs(Value1)),0) From " & LYDatabaseName & "..Tran2 T1 Where T1.CM1=Master1.Code And  (RECTYPE = 2 OR RECTYPE = 7) AND VCHTYPE = 3 AND MasterCode1 IN (Select Code from Master1 Where  MasterType=6 and Alias<>'' and (Alias Like'" & StrBusyCode & "%' or Alias Like'" & strBusyCode2 & "%')) AND DATE >= '" & Format(CDate("01-Apr-" & Right(DatabaseName, 4) - 1), "dd-MMM-yyyy") & "' AND DATE <= '" & Format(CDate("31-Mar-" & Right(DatabaseName, 4)), "dd-MMM-yyyy") & "' And IsReturnQty=0)) as LYSaleQty,"
'                            str_qry = str_qry & " (SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran2 WHERE VchType=9 AND RecType=2 AND Date>='" & Format(DateAdd("d", -15, GetDate(MhDateInput2.Text)), "dd-MMM-yyyy") & "' AND Date<='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=(Select Code from Master1 Where  MasterType=6 and Alias<>'' and Alias Like'" & StrBusyCode & "-N" & "%')) As PreviousSale15,"
'                            str_qry = str_qry & " (SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran2 WHERE VchType=9 AND RecType=2 AND Date>='" & Format(DateAdd("d", -30, GetDate(MhDateInput2.Text)), "dd-MMM-yyyy") & "' AND Date<='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=(Select Code from Master1 Where  MasterType=6 and Alias<>'' and Alias Like'" & StrBusyCode & "-N" & "%')) As PreviousSale30,"
'                            str_qry = str_qry & " (SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran2 WHERE VchType=9 AND RecType=2 AND Date>='" & FromDate & "' AND Date<='" & Format(DateAdd("d", 30, FromDate), "dd-MMM-yyyy") & "' AND MasterCode1=(Select Code from Master1 Where  MasterType=6 and Alias<>'' and Alias Like'" & StrBusyCode & "-N" & "%')) As NextSale30,"
'                            str_qry = str_qry & " (SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran2 WHERE VchType=9 AND RecType=2 AND Date>='" & FromDate & "' AND Date<='" & Format(DateAdd("d", 60, FromDate), "dd-MMM-yyyy") & "' AND MasterCode1=(Select Code from Master1 Where  MasterType=6 and Alias<>'' and Alias Like'" & StrBusyCode & "-N" & "%')) As NextSale60"
'                            str_qry = str_qry & " FROM MASTER1 WHERE Code in (SELECT DISTINCT CM1 FROM TRAN2 WHERE (RECTYPE = 2 OR RECTYPE = 7)  AND (VCHTYPE = 9 OR VCHTYPE = 3) AND DATE >= '" & GetDate(MhDateInput1.Text) & "' AND DATE <= '" & GetDate(MhDateInput2.Text) & "' AND MASTERCODE1 = (Select Code from Master1 Where  MasterType=6 and Alias<>'' and Alias Like'" & StrBusyCode & "-N" & "%')) AND Code IN (SELECT CODE FROM HELP1 WHERE MASTERSERIES = 26841 And MasterType= 2)  Order By Name"
'***********End Old Code***********************************************
                            
                           
                            
                            str_qry = "SELECT CODE,Name,ParentGrp,"
                            str_qry = str_qry & " (Select ISNULL(SUM(ABS(Value1)),0) From Tran2 Where Tran2.CM1=Master1.Code And  (RECTYPE = 2 OR RECTYPE = 7) AND VCHTYPE = 9 AND MasterCode1 = (Select MAX(code) from Master1 Where  MasterType=6 and Alias<>'' AND Alias Like'" & StrBusyCode & "-N" & "%'  ) AND DATE >= '" & GetDate(MhDateInput1.Text) & "' AND DATE <= '" & GetDate(MhDateInput2.Text) & "' And IsReturnQty=0) as SaleQty,"
                            str_qry = str_qry & " (Select ISNULL(SUM(ABS(Value1)),0) From Tran2 Where Tran2.CM1=Master1.Code And  (RECTYPE = 2 OR RECTYPE = 7) AND VCHTYPE = 3 AND MasterCode1 = (Select MAX(code) from Master1 Where  MasterType=6 and Alias<>'' AND Alias Like'" & StrBusyCode & "-N" & "%'  ) AND DATE >= '" & GetDate(MhDateInput1.Text) & "' AND DATE <= '" & GetDate(MhDateInput2.Text) & "' And IsReturnQty=0) as SaleRetQty,"
                            str_qry = str_qry & " (Select ISNULL(SUM(ABS(Value1)),0) From " & LYDatabaseName & "..Tran2 T1 Where T1.CM1=Master1.Code And  (RECTYPE = 2 OR RECTYPE = 7) AND VCHTYPE = 9 AND MasterCode1 = (Select top 1 code from Master1 Where  MasterType=6 and Alias<>'' and (Alias Like'" & StrBusyCode & "%' or Alias Like'" & strBusyCode2 & "-N" & "%') and Year(CreationTime)=2018) AND DATE >= '" & Format(CDate("01-Apr-" & Right(DatabaseName, 4) - 1), "dd-MMM-yyyy") & "' AND DATE <= '" & Format(CDate("31-Mar-" & Right(DatabaseName, 4)), "dd-MMM-yyyy") & "' And IsReturnQty=0) as LYSaleQty,"
                                
                            
'                            str_qry = str_qry & " ((Select ISNULL(SUM(ABS(Value1)),0) From " & LYDatabaseName & "..Tran2 T1 Where T1.CM1=Master1.Code And  (RECTYPE = 2 OR RECTYPE = 7) AND VCHTYPE = 9 AND MasterCode1 IN (Select Code from Master1 Where  MasterType=6 and Alias<>'' and (Alias Like'" & StrBusyCode & "%' or Alias Like'" & strBusyCode2 & "%')) AND DATE >= '" & Format(CDate("01-Apr-" & Right(DatabaseName, 4) - 1), "dd-MMM-yyyy") & "' AND DATE <= '" & Format(CDate("31-Mar-" & Right(DatabaseName, 4)), "dd-MMM-yyyy") & "' And IsReturnQty=0)-"
'                            str_qry = str_qry & " (Select ISNULL(SUM(ABS(Value1)),0) From " & LYDatabaseName & "..Tran2 T1 Where T1.CM1=Master1.Code And  (RECTYPE = 2 OR RECTYPE = 7) AND VCHTYPE = 3 AND MasterCode1 IN (Select Code from Master1 Where  MasterType=6 and Alias<>'' and (Alias Like'" & StrBusyCode & "%' or Alias Like'" & strBusyCode2 & "%')) AND DATE >= '" & Format(CDate("01-Apr-" & Right(DatabaseName, 4) - 1), "dd-MMM-yyyy") & "' AND DATE <= '" & Format(CDate("31-Mar-" & Right(DatabaseName, 4)), "dd-MMM-yyyy") & "' And IsReturnQty=0)) as LYSaleQty,"
'
                            
                            
                            str_qry = str_qry & " (SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran2 WHERE Tran2.CM1=Master1.Code And VchType=9 AND RecType=2 AND Date>='" & Format(DateAdd("d", -14, GetDate(MhDateInput2.Text)), "dd-MMM-yyyy") & "' AND Date<='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=(Select MAX(code) from Master1 Where  MasterType=6 and Alias<>'' AND Alias Like'" & StrBusyCode & "-N" & "%'  )) As PreviousSale15,"
                            str_qry = str_qry & " (SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran2 WHERE Tran2.CM1=Master1.Code And VchType=9 AND RecType=2 AND Date>='" & Format(DateAdd("d", -29, GetDate(MhDateInput2.Text)), "dd-MMM-yyyy") & "' AND Date<='" & GetDate(MhDateInput2.Text) & "' AND MasterCode1=(Select MAX(code) from Master1 Where  MasterType=6 and Alias<>'' AND Alias Like'" & StrBusyCode & "-N" & "%'  )) As PreviousSale30,"
                            str_qry = str_qry & " (SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran2 WHERE Tran2.CM1=Master1.Code And VchType=9 AND RecType=2 AND Date>='" & FromDate & "' AND Date<='" & Format(DateAdd("d", 29, FromDate), "dd-MMM-yyyy") & "' AND MasterCode1=(Select MAX(code) from Master1 Where  MasterType=6 and Alias<>'' AND Alias Like'" & StrBusyCode & "-N" & "%'  )) As NextSale30,"
                            str_qry = str_qry & " (SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran2 WHERE Tran2.CM1=Master1.Code And VchType=9 AND RecType=2 AND Date>='" & FromDate & "' AND Date<='" & Format(DateAdd("d", 59, FromDate), "dd-MMM-yyyy") & "' AND MasterCode1=(Select MAX(code) from Master1 Where  MasterType=6 and Alias<>'' AND Alias Like'" & StrBusyCode & "-N" & "%'  )) As NextSale60"
                            str_qry = str_qry & " FROM MASTER1 WHERE Code in (SELECT DISTINCT CM1 FROM TRAN2 WHERE (RECTYPE = 2 OR RECTYPE = 7)  AND (VCHTYPE = 9 OR VCHTYPE = 3) AND DATE >= '" & GetDate(MhDateInput1.Text) & "' AND DATE <= '" & GetDate(MhDateInput2.Text) & "' AND MASTERCODE1 = (Select MAX(code) from Master1 Where  MasterType=6 and Alias<>'' AND Alias Like'" & StrBusyCode & "%'  )"
                            str_qry = str_qry & " And CM1 in (SELECT CODE FROM MASTER1 as M1 WHERE M1.MasterType IN (2))) Order By Name"
                            rstProductionPlanning.Open str_qry, CxnProductionPlanning, adOpenKeyset, adLockReadOnly
                            rstProductionPlanning.ActiveConnection = Nothing
                            Call UpdateStockFigures("S", StrBusyCode) 'Update Sale,Expected Sales Figures
                                                    
                           
                            '*Banglore Stock***
                            SQLQuery = ""
                            SQLQuery = "SELECT M.Mc AS McName,M.CM as Code,M.BSStock as B1,M.PG as ParentGrp,ISNULL(S1.MTB,0) AS MainTransBal," & _
                                " (select ISNULL(Sum(d1),0) from tran4 where tran4.mastercode1=(Select MAX(code) From Master1 Where Alias Like'" & StrBusyCode & "-N" & "%'  ) and  tran4.mastercode2= M.Cm) AS MainOpBal," & _
                                " (select ISNULL(Sum(d3),0) from tran4 where tran4.mastercode1=(Select MAX(code) From Master1 Where Alias Like'" & StrBusyCode & "-N" & "%'  ) and  tran4.mastercode2= M.Cm) AS OPAmt " & _
                                " From (Select Name as Mc,Code as CM,B1 as BSStock,ParentGrp as PG from Master1 where Mastertype=11 And CODE IN ((Select Code From Master1 Where " & _
                                " ParentGrp in(Select Code From Master1 Where Name='Branch Group'))))  AS M Left Join (SELECT Mastercode2, sum(value1) AS MTB,sum(value2) AS ATB " & _
                                " From tran2 Where rectype = 2 And Mastercode1=(Select MAX(code) From Master1 Where Alias Like'" & StrBusyCode & "-N" & "%'  ) and  Date <= '" & GetDate(MhDateInput2.Text) & "' group by Mastercode2 ) AS " & _
                                " S1 ON (S1.Mastercode2 = M.CM)  ORDER BY M.Mc"
                            If rstProductionPlanning.State = adStateOpen Then rstProductionPlanning.Close
                            rstProductionPlanning.Open SQLQuery, CxnProductionPlanning, adOpenKeyset, adLockReadOnly
                            rstProductionPlanning.ActiveConnection = Nothing
                            Call UpdateStockFigures("B", StrBusyCode) 'Update Sale,Stock Transfer,Specimen, Stock & Pending Sales Order Figures

                           
                           
                            
                            ''*Distrubutor Stock***
                            If rstProductionPlanning.State = adStateOpen Then rstProductionPlanning.Close
                            Dim StrQry_D As String
                            StrQry_D = "SELECT M.Mc AS McName,M.CM as Code,M.BSStock as B1,M.PG as ParentGrp,ISNULL(S1.MTB,0) AS MainTransBal, ISNULL(S1.ATB,0) AS AltTransBal," & _
                            "(select Sum(d1) from tran4 where tran4.mastercode1=  (Select MAX(code) from Master1 Where  MasterType=6 and Alias<>'' and (Alias Like'" & StrBusyCode & "-N" & "%' or Alias Like'" & strBusyCode2 & "-N" & "%'  )) and  tran4.mastercode2= M.Cm) AS MainOpBal," & _
                            "(select Sum(d2) from tran4 where tran4.mastercode1= (Select MAX(code) from Master1 Where  MasterType=6 and Alias<>'' and (Alias Like'" & StrBusyCode & "-N" & "%' or Alias Like'" & strBusyCode2 & "-N" & "%'  ))" & _
                            "and  tran4.mastercode2= M.Cm) AS AltOPBal,(select Sum(d3) from tran4 where tran4.mastercode1= (Select MAX(code) from Master1 Where  MasterType=6 and Alias<>'' and (Alias Like'" & StrBusyCode & "-N" & "%' or Alias Like'" & strBusyCode2 & "-N" & "%'  )) and  tran4.mastercode2= M.Cm) AS OPAmt " & _
                            " From (Select Name as Mc,Code as CM,B1 as BSStock,ParentGrp as PG from Master1 where Mastertype=11)  AS M Left Join (SELECT Mastercode2, sum(value1) AS MTB,sum(value2) AS ATB " & _
                            "From tran2 Where rectype = 2 and Mastercode1=(Select MAX(code) from Master1 Where  MasterType=6 and Alias<>'' and (Alias Like'" & StrBusyCode & "-N" & "%' or Alias Like'" & strBusyCode2 & "-N" & "%'  )) and  Date <= '" & GetDate(MhDateInput2.Text) & "' group by Mastercode2 ) AS S1 ON (S1.Mastercode2 = M.CM) And  M.BSStock>0  ORDER BY M.Mc"
                            
                            rstProductionPlanning.Open StrQry_D, CxnDistributor, adOpenKeyset, adLockReadOnly
                            rstProductionPlanning.ActiveConnection = Nothing
                            Call UpdateStockFigures("D", StrBusyCode) 'Update Sale,Stock Transfer,Specimen, Stock & Pending Sales Order Figures



                            If rstProductionPlanning.State = adStateOpen Then rstProductionPlanning.Close
                            Dim str_qry_d As String
                            str_qry_d = "SELECT CODE,Name,"
                            str_qry_d = str_qry_d & " ((Select ISNULL(Sum(abs(Value1)),0) From Tran2 Where Tran2.CM1=Master1.Code And  (RECTYPE = 2 OR RECTYPE = 7) AND VCHTYPE = 9 AND MasterCode1 = (Select max(Code) from Master1 Where  MasterType=6 and Alias<>'' and (Alias Like'" & StrBusyCode & "-N" & "%' or Alias Like'" & strBusyCode2 & "-N" & "%'  )) AND DATE >= '" & GetDate(MhDateInput1.Text) & "' AND DATE <= '" & GetDate(MhDateInput2.Text) & "' AND IsReturnQty=0)-(Select ISNULL(Sum(abs(Value1)),0) From Tran2 Where Tran2.CM1=Master1.Code And  (RECTYPE = 2 OR RECTYPE = 7) AND VCHTYPE = 3 AND MasterCode1 = (Select Max(Code) from Master1 Where  MasterType=6 and Alias<>'' and (Alias Like'" & StrBusyCode & "-N" & "%' or Alias Like'" & strBusyCode2 & "-N" & "%'  )) AND DATE >= '" & GetDate(MhDateInput1.Text) & "' AND DATE <= '" & GetDate(MhDateInput2.Text) & "' AND IsReturnQty=0)) as DSaleQty,"
                            str_qry_d = str_qry_d & " ((Select ISNULL(Sum(abs(Value1)),0) From " & DLYDatabaseName & "..Tran2 Where Tran2.CM1=Master1.Code And  (RECTYPE = 2 OR RECTYPE = 7) AND VCHTYPE = 9 AND MasterCode1 = (Select max(Code) from Master1 Where  MasterType=6 and Alias<>'' and (Alias Like'" & StrBusyCode & "-N" & "%' or Alias Like'" & strBusyCode2 & "-N" & "%' and Year(CreationTime)=2018)) AND DATE >= '" & Format(CDate("01-Apr-" & Right(DatabaseName, 4) - 1), "dd-MMM-yyyy") & "' AND DATE <= '" & Format(CDate("31-Mar-" & Right(DatabaseName, 4)), "dd-MMM-yyyy") & "' And IsReturnQty=0)-"
                            str_qry_d = str_qry_d & " (Select ISNULL(Sum(abs(Value1)),0) From " & DLYDatabaseName & "..Tran2 Where Tran2.CM1=Master1.Code And  (RECTYPE = 2 OR RECTYPE = 7) AND VCHTYPE = 3 AND MasterCode1 = (Select max(Code) from Master1 Where  MasterType=6 and Alias<>'' and (Alias Like'" & StrBusyCode & "-N" & "%' or Alias Like'" & StrBusyCode & "-N" & "%' and Year(CreationTime)=2018)) AND DATE >= '" & Format(CDate("01-Apr-" & Right(DatabaseName, 4) - 1), "dd-MMM-yyyy") & "' AND DATE <= '" & Format(CDate("31-Mar-" & Right(DatabaseName, 4)), "dd-MMM-yyyy") & "' And IsReturnQty=0)) as LYDSaleQty"
                            str_qry_d = str_qry_d & " FROM MASTER1 WHERE Code in (SELECT DISTINCT CM1 FROM TRAN2 WHERE (RECTYPE = 2 OR RECTYPE = 7)  AND (VCHTYPE = 9 OR VCHTYPE = 3) AND DATE >= '" & GetDate(MhDateInput1.Text) & "' AND DATE <= '" & GetDate(MhDateInput2.Text) & "' AND MASTERCODE1 = (Select Max(Code) from Master1 Where  MasterType=6 and Alias<>'' and (Alias Like'" & StrBusyCode & "-N" & "%' or Alias Like'" & strBusyCode2 & "-N" & "%'  ))) Order By Name"
                            rstProductionPlanning.Open str_qry_d, CxnDistributor, adOpenKeyset, adLockReadOnly
                            rstProductionPlanning.ActiveConnection = Nothing
                            Call UpdateStockFigures("DS", StrBusyCode) 'Update Distributor Sale Figures

                           
                            '09/09/2019
                            MdiMainMenu.ProgressBar1.Value = MdiMainMenu.ProgressBar1.Value + 2

                       
                       
                       End If
                       DoEvents
               Next z
        End If
        
        If i = 0 Then Exit Do Else DatabaseName = Mid(DatabaseName, i + 1): i = 0
        If j = 0 Then Exit Do Else DatabaseNameDistributor = Mid(DatabaseNameDistributor, j + 1): j = 0
  
    Loop
    
    
    DatabaseName = Trim(ReadFromFile("Saral Database Name")): If DatabaseName = "" Then Exit Sub
    i = 0: K = 0
    Do While True
         i = InStr(1, DatabaseName, ",")
        If CxnProductionPlanning.State = adStateOpen Then CxnProductionPlanning.Close
        If i = 0 Then CxnProductionPlanning.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DatabasePath & "\" & Mid(DatabaseName, 1) & ";Persist Security Info=False;Jet OLEDB:Database Password=RSPLILoveMyINDIA" Else CxnProductionPlanning.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DatabasePath & "\" & Mid(DatabaseName, 1, i - 1) & ";Persist Security Info=False;Jet OLEDB:Database Password=RSPLILoveMyINDIA"
        K = K + 1
        If rstProductionPlanning.State = adStateOpen Then rstProductionPlanning.Close
        If K = 2 Then
           FromDate = "01-Apr-" & Trim(Year(GetDate(MhDateInput1.Text)) - 1)
           ToDate = "31-Mar-" & Trim(Year(GetDate(MhDateInput1.Text)))
           rstProductionPlanning.Open "SELECT M.BusyCode,CLng(Sum(C.ActualQuantity)) As PrintOrder,CLng(IIF(ISNULL(SUM(C.ActualQuantity-IIF(C.Status IN ('D','E','W'),C.ActualQuantity,P.ReceivedQuantity))),0,SUM(C.ActualQuantity-IIF(C.Status IN ('D','E','W'),C.ActualQuantity,P.ReceivedQuantity))))+(SELECT CLng(IIF(ISNULL(SUM(C.ActualQuantity-IIF(C.Status IN ('D','E','W') OR BillNo<>'',C.ActualQuantity,P.ReceivedQuantity))),0,SUM(C.ActualQuantity-IIF(C.Status IN ('D','E','W') OR BillNo<>'',C.ActualQuantity,P.ReceivedQuantity)))) FROM BookPOChild08 C INNER JOIN BookPOParent P ON P.Code=C.Code WHERE P.Type='R' AND P.Book=M.Code AND LEFT(P.Code,1)<>'*' AND P.Date>=#" & FromDate & "# AND P.Date<=#" & ToDate & "#) As PendingPrintOrder FROM (BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code) INNER JOIN BookMaster M ON M.Code=P.Book " & _
                                       "WHERE P.Type='F' AND LEFT(P.Code,1)<>'*' AND P.Date>=#" & FromDate & "# AND P.Date<=#" & ToDate & "# AND RIGHT(M.BusyCode,1)<>'S' AND LEFT(M.BusyCode,4) IN (" & SelectedItems(ListView3, True) & ") GROUP BY M.Code,M.BusyCode ORDER BY M.BusyCode", CxnProductionPlanning, adOpenKeyset, adLockReadOnly
           rstProductionPlanning.ActiveConnection = Nothing
           Call UpdateStockFigures("LYPO", "") 'Update Print Order & Pending Print Order Figures
           MdiMainMenu.ProgressBar1.Value = MdiMainMenu.ProgressBar1.Value + 16.5
        ElseIf K = 3 Then
           FromDate = "01-Apr-" & Trim(Year(GetDate(MhDateInput1.Text)) + IIf(OrderType = "M", 2, 0))
            
           rstProductionPlanning.Open "SELECT M.BusyCode,CLng(Sum(C.ActualQuantity)) As PrintOrder,CLng(IIF(ISNULL(SUM(C.ActualQuantity-IIF(C.Status IN ('D','E','W'),C.ActualQuantity,P.ReceivedQuantity))),0,SUM(C.ActualQuantity-IIF(C.Status IN ('D','E','W'),C.ActualQuantity,P.ReceivedQuantity))))+(SELECT CLng(IIF(ISNULL(SUM(C.ActualQuantity-IIF(C.Status IN ('D','E','W') OR BillNo<>'',C.ActualQuantity,P.ReceivedQuantity))),0,SUM(C.ActualQuantity-IIF(C.Status IN ('D','E','W') OR BillNo<>'',C.ActualQuantity,P.ReceivedQuantity)))) FROM BookPOChild08 C INNER JOIN BookPOParent P ON P.Code=C.Code WHERE P.Type='R' AND P.Book=M.Code AND LEFT(P.Code,1)<>'*' AND P.Date>=#" & FromDate & "# AND P.Date<=#" & GetDate(MhDateInput2.Text) & "#) As PendingPrintOrder FROM (BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code) INNER JOIN BookMaster M ON M.Code=P.Book " & _
                                       "WHERE P.Type='F' AND LEFT(P.Code,1)<>'*' AND P.Date>=#" & FromDate & "# AND P.Date<=#" & GetDate(MhDateInput2.Text) & "# AND RIGHT(M.BusyCode,1)<>'S' AND LEFT(M.BusyCode,4) IN (" & SelectedItems(ListView3, True) & ") GROUP BY M.Code,M.BusyCode ORDER BY M.BusyCode", CxnProductionPlanning, adOpenKeyset, adLockReadOnly
           rstProductionPlanning.ActiveConnection = Nothing
           Call UpdateStockFigures("CYPO", "") 'Update Print Order & Pending Print Order Figures
        End If
        MdiMainMenu.ProgressBar1.Value = MdiMainMenu.ProgressBar1.Value
        If i = 0 Then Exit Do Else DatabaseName = Mid(DatabaseName, i + 1): i = 0
    Loop
    Call CloseRecordset(rstProductionPlanning)
    Call CloseConnection(CxnProductionPlanning)
    Screen.MousePointer = vbNormal
    On Error Resume Next
    If Not FileExist(App.Path & "\Template\Production Planning.xlsx") Then Exit Sub
    Screen.MousePointer = vbHourglass
    If rstProductionPlanning.State = adStateOpen Then rstProductionPlanning.Close
    rstProductionPlanning.Open "SELECT * FROM BookMaster WHERE Type='F' AND LEFT(BusyCode,4) IN (" & SelectedItems(ListView3, True) & ") AND BusyCode<>'' AND RIGHT(BusyCode,1)<>'S' ORDER BY PrintName", CxnDatabase, adOpenKeyset, adLockReadOnly
    If rstProductionPlanning.RecordCount = 0 Then
        DisplayError ("No Record Found")
        ShowProgressInStatusBar False
        Screen.MousePointer = vbNormal
        On Error GoTo 0
        Exit Sub
    End If
    
    DoEvents
    'Writing To Excel
    
    Set oExcel = CreateObject("Excel.Application")
    oExcel.Workbooks.Open (App.Path & "\Template\Production Planning New"): oExcel.DisplayAlerts = False
    
    oExcel.Workbooks.Item(1).SaveAs (App.Path & "\Report\Production Planning New (" & CompCode & ")"): oExcel.DisplayAlerts = True
    oExcel.Sheets("Reorder Level Register").Visible = False: oExcel.Sheets("Production Planning (" & IIf(OrderType = "M", "SO", "MO") & ")").Visible = False: oExcel.Sheets("Production Planning (" & IIf(OrderType = "M", "MO", "SO") & ")").Select: oExcel.Visible = False
    oExcel.Cells(1, "A").Value = Trim(rstCompanyMaster.Fields("PrintName").Value)
    oExcel.Cells(2, "A").Value = "Production Planning (Supplement Orders) As On [" & Format(GetDate(MhDateInput2.Text), "dd-MMM-yyyy") & "]"
    oExcel.Cells(4, "C").Value = "01-Apr-" & Trim(Year(GetDate(MhDateInput1.Text)) - 1) & " To " & " 31-Mar-" & Trim(Year(GetDate(MhDateInput1.Text)))
    i = 5: Cnt = 1
    Do While Not rstProductionPlanning.EOF
        oExcel.Cells(i, "A").Value = Cnt
        oExcel.Application.Cells(i, "B").Value = Trim(rstProductionPlanning.Fields("PrintName").Value)
        oExcel.Application.Cells(i, "C").Value = Trim(rstProductionPlanning.Fields("PrintedLYear").Value)
        oExcel.Application.Cells(i, "D").Value = Val(rstProductionPlanning.Fields("PrintedCYear").Value)
        oExcel.Application.Cells(i, "E").Value = rstProductionPlanning.Fields("CurrentStock").Value
        oExcel.Application.Cells(i, "F").Value = Val(rstProductionPlanning.Fields("Pending_PO").Value)
        oExcel.Application.Cells(i, "G").Value = Val(rstProductionPlanning.Fields("Pending_SO").Value)
        oExcel.Application.Cells(i, "H").Value = Val(rstProductionPlanning.Fields("EffectiveStock").Value)
        oExcel.Application.Cells(i, "I").Value = Val(rstProductionPlanning.Fields("DistbrStock").Value)
        oExcel.Application.Cells(i, "J").Value = Val(rstProductionPlanning.Fields("BlrStock").Value)
        oExcel.Application.Cells(i, "K").Value = Val(rstProductionPlanning.Fields("SpecimenLyear").Value)
        oExcel.Application.Cells(i, "L").Value = Val(rstProductionPlanning.Fields("SpecimenCYear").Value)
        oExcel.Application.Cells(i, "M").Value = Val(rstProductionPlanning.Fields("SaleLYear").Value) + Val(rstProductionPlanning.Fields("DistbrLYear").Value)
        
        Dim cy As Long
        Dim dcy As Long
        
        cy = Val(rstProductionPlanning.Fields("SaleCYear").Value)
        dcy = Val(rstProductionPlanning.Fields("DistbrCYear").Value)

        oExcel.Application.Cells(i, "N").Value = Val(rstProductionPlanning.Fields("SaleCYear").Value) + Val(rstProductionPlanning.Fields("DistbrCYear").Value)
        
        oExcel.Application.Cells(i, "O").Value = Val(rstProductionPlanning.Fields("ESale30").Value)
        
        oExcel.Application.Cells(i, "P").Value = Val(rstProductionPlanning.Fields("ESale60").Value)
        oExcel.Application.Cells(i, "Q").Value = Val(rstProductionPlanning.Fields("SalePDays15").Value)
        oExcel.Application.Cells(i, "R").Value = Val(rstProductionPlanning.Fields("SalePDays30").Value)
        Cnt = Cnt + 1: i = i + 1
        rstProductionPlanning.MoveNext
    Loop
    
    oExcel.Columns("A:B").EntireColumn.AutoFit
    oExcel.Columns(6).Hide = True
    oExcel.Workbooks.Item(1).Save
    Screen.MousePointer = vbNormal
    MdiMainMenu.ProgressBar1.Value = 100
    If OutputTo = "S" Then oExcel.Range("A1").Activate: oExcel.Visible = True Else oExcel.Workbooks.Item(1).PrintOut
    ShowProgressInStatusBar False
    Set oExcel = Nothing
    On Error GoTo 0
    Exit Sub

ErrorHandler:
    Screen.MousePointer = vbNormal
    DisplayError ("Failed to update Production Planning figures")
    ShowProgressInStatusBar False
    Call CloseRecordset(rstProductionPlanning)
    Call CloseConnection(CxnProductionPlanning)
    Call CloseConnection(CxnDistributor)

End Sub

Private Sub UpdateStockFigures(ByVal UpdationType As String, ByVal BusyCode As String)
    
    If rstProductionPlanning.RecordCount > 0 Then rstProductionPlanning.MoveFirst
    Dim CurrentStock As Long, PendingPO As Long, PendingSO As Long, EffectiveStock As Long, BlrStock As Long, DisbtrStock As Long, LYSale As Long, CYSale As Long
    Dim NextSale30 As Long, NextSale60 As Long, PSale15 As Long, PSale30 As Long, SpecimenQty As Long, LYSpecimenQty As Long, DSaleQty As Long, DLYSaleQty As Long
    BrPendingPO = 0
    BrPendingSO = 0
    Do While Not rstProductionPlanning.EOF
        If UpdationType = "D" Then
           DisbtrStock = DisbtrStock + (Val(IIf(IsNull(rstProductionPlanning.Fields("MainTransBal").Value), 0, rstProductionPlanning.Fields("MainTransBal").Value)) + Val(IIf(IsNull(rstProductionPlanning.Fields("MainOpBal").Value), 0, rstProductionPlanning.Fields("MainOpBal").Value)))
           CxnDatabase.Execute "UPDATE BookMaster SET DistbrStock=" & DisbtrStock & " WHERE BusyCode='" & BusyCode & "'"
        ElseIf UpdationType = "B" Then
           
           If rstProductionPlanning.Fields("McName").Value = "Blr" Then
                If Val(rstProductionPlanning.Fields("MainTransBal").Value) > 0 Or Val(rstProductionPlanning.Fields("MainTransBal").Value) < 0 Then
                   BlrStock = BlrStock + (Val(IIf(IsNull(rstProductionPlanning.Fields("MainTransBal").Value), 0, rstProductionPlanning.Fields("MainTransBal").Value)) + Val(IIf(IsNull(rstProductionPlanning.Fields("MainOpBal").Value), 0, rstProductionPlanning.Fields("MainOpBal").Value)))
                
                      Call UpdatePendingStock(BusyCode, rstProductionPlanning.Fields("Code").Value, "Branch Group")
                
                CxnDatabase.Execute "UPDATE BookMaster SET BlrStock=" & BlrStock & " WHERE BusyCode='" & BusyCode & "'"
                CxnDatabase.Execute "UPDATE BookMaster SET BlrStock=(BlrStock - " & BrPendingSO & ")+ " & Abs(BrPendingPO) & " WHERE BusyCode='" & BusyCode & "'"
           
                End If
          
           End If
           
        ElseIf UpdationType = "DS" Then
           DSaleQty = DSaleQty + Abs(Val(rstProductionPlanning.Fields("DSaleQty").Value))
           DLYSaleQty = DLYSaleQty + Abs(Val(rstProductionPlanning.Fields("LYDSaleQty").Value))
           CxnDatabase.Execute "UPDATE BookMaster SET DistbrLYear=" & Abs(DLYSaleQty) & ",DistbrCYear=" & Abs(DSaleQty) & " WHERE BusyCode='" & BusyCode & "'"
        ElseIf UpdationType = "CYPO" Then
              CxnDatabase.Execute "UPDATE BookMaster SET PrintedCYear=" & Val(CheckNull(rstProductionPlanning.Fields("PrintOrder").Value)) & " WHERE LEFT(BusyCode,4)='" & Left(rstProductionPlanning.Fields("BusyCode").Value, 4) & "'"
        ElseIf UpdationType = "LYPO" Then
              CxnDatabase.Execute "UPDATE BookMaster SET PrintedLYear=" & Val(CheckNull(rstProductionPlanning.Fields("PrintOrder").Value)) & " WHERE LEFT(BusyCode,4)='" & Left(rstProductionPlanning.Fields("BusyCode").Value, 4) & "'"
        ElseIf UpdationType = "SPCY" Then
             SpecimenQty = SpecimenQty + (Val(rstProductionPlanning.Fields("MainTransBal").Value) + Val(rstProductionPlanning.Fields("MainOPBal").Value))
             CxnDatabase.Execute "UPDATE BookMaster SET SpecimenCYear=" & SpecimenQty & " WHERE BusyCode='" & BusyCode & "'"
        ElseIf UpdationType = "SPLY" Then
            LYSpecimenQty = LYSpecimenQty + (Val(rstProductionPlanning.Fields("MainTransBal").Value) + Val(rstProductionPlanning.Fields("MainOPBal").Value))
            CxnDatabase.Execute "UPDATE BookMaster SET SpecimenLYear=" & LYSpecimenQty & " WHERE BusyCode='" & BusyCode & "'"
        
        ElseIf UpdationType = "S" Then
            
            'CYSale = CYSale + (Val(rstProductionPlanning.Fields("SaleQty").Value) - Val(rstProductionPlanning.Fields("SaleRetQty").Value))
             
            CYSale = CYSale + Val(rstProductionPlanning.Fields("SaleQty").Value)
            LYSale = LYSale + Val(rstProductionPlanning.Fields("LYSaleQty").Value)
            NextSale30 = NextSale30 + Val(rstProductionPlanning.Fields("NextSale30").Value)
            NextSale60 = NextSale60 + Val(rstProductionPlanning.Fields("NextSale60").Value)
            PSale15 = PSale15 + Val(rstProductionPlanning.Fields("PreviousSale15").Value)
            PSale30 = PSale30 + Val(rstProductionPlanning.Fields("PreviousSale30").Value)
            CxnDatabase.Execute "UPDATE BookMaster SET SaleCYear=" & CYSale & ",SaleLYear=" & LYSale & ",ESale30=" & NextSale30 & ",Esale60=" & NextSale60 & ",SalePDays15=" & PSale15 & ",SalePDays30=" & PSale30 & ",FinalOrder=1000 WHERE BusyCode='" & BusyCode & "'"
       
        ElseIf UpdationType = "C" Then
             If Val(rstProductionPlanning.Fields("MainTransBal").Value) > 0 Or Val(rstProductionPlanning.Fields("MainTransBal").Value) < 0 Then
               CurrentStock = CurrentStock + (Val(IIf(IsNull(rstProductionPlanning.Fields("MainTransBal").Value), 0, rstProductionPlanning.Fields("MainTransBal").Value)) + Val(IIf(IsNull(rstProductionPlanning.Fields("MainOpBal").Value), 0, rstProductionPlanning.Fields("MainOpBal").Value)))
               Call UpdatePendingStock(BusyCode, rstProductionPlanning.Fields("Code").Value, "Head Office (Group)")
               CxnDatabase.Execute "UPDATE BookMaster SET CurrentStock=" & CurrentStock & ",EffectiveStock=(" & CurrentStock & " - Pending_SO)+Pending_PO WHERE BusyCode='" & BusyCode & "'"
            End If
        End If
       rstProductionPlanning.MoveNext
    Loop
    
'    If UpdationType = "B" Then
'        CxnDatabase.Execute "UPDATE BookMaster SET BlrStock=" & BlrStock & " WHERE BusyCode='" & BusyCode & "'"
'        CxnDatabase.Execute "UPDATE BookMaster SET BlrStock=(BlrStock - " & BrPendingSO & ")+ " & BrPendingPO & " WHERE BusyCode='" & BusyCode & "'"

'    End If
    
    
'    If UpdationType = "D" Then
'       'CxnDatabase.Execute "UPDATE BookMaster SET DistbrStock=" & DisbtrStock & " WHERE BusyCode='" & BusyCode & "'"
'    ElseIf UpdationType = "B" Then
'        'CxnDatabase.Execute "UPDATE BookMaster SET BlrStock=" & BlrStock & " WHERE BusyCode='" & BusyCode & "'"
'        'CxnDatabase.Execute "UPDATE BookMaster SET BlrStock=(BlrStock - " & BrPendingSO & ")+ " & BrPendingPO & " WHERE BusyCode='" & BusyCode & "'"
'
'    ElseIf UpdationType = "DS" Then
'       'CxnDatabase.Execute "UPDATE BookMaster SET DistbrLYear=" & Abs(DLYSaleQty) & ",DistbrCYear=" & Abs(DSaleQty) & " WHERE BusyCode='" & BusyCode & "'"
'    ElseIf UpdationType = "SPCY" Then
'       'CxnDatabase.Execute "UPDATE BookMaster SET SpecimenCYear=" & SpecimenQty & " WHERE BusyCode='" & BusyCode & "'"
'    ElseIf UpdationType = "SPLY" Then
'       'CxnDatabase.Execute "UPDATE BookMaster SET SpecimenLYear=" & LYSpecimenQty & " WHERE BusyCode='" & BusyCode & "'"
'    ElseIf UpdationType = "S" Then
'       'CxnDatabase.Execute "UPDATE BookMaster SET SaleCYear=" & CYSale & ",SaleLYear=" & LYSale & ",ESale30=" & NextSale30 & ",Esale60=" & NextSale60 & ",SalePDays15=" & PSale15 & ",SalePDays30=" & PSale30 & ",FinalOrder=1000 WHERE BusyCode='" & BusyCode & "'"
'    ElseIf UpdationType = "C" Then
'       'CxnDatabase.Execute "UPDATE BookMaster SET CurrentStock=" & CurrentStock & ",EffectiveStock=(" & CurrentStock & " - Pending_SO)+Pending_PO WHERE BusyCode='" & BusyCode & "'"
'    End If

End Sub
Private Sub UpdatePendingStock(ByVal ItemCode As String, ByVal RefCode As Long, ByVal StockGroup As String)
    Dim Rs As New ADODB.Recordset
    Dim PendingSO As Long, PendingPO As Long
    
    Dim fff As String
    
    
    If Rs.State = adStateOpen Then Rs.Close
    
    
    If StockGroup = "Branch Group" Then
       
       fff = "Select T3.vchtype,Tr.* From Tran3 As T3 Inner Join (Select RefCode as RCode,Sum(Value1) as Bal1,Sum(Value2) as Bal2 From Tran3 Where MASTERCODE1= (Select MAX(code) From Master1 Where Alias Like'" & ItemCode & "-N" & "%')  AND Date <= '" & GetDate(MhDateInput2.Text) & "' Group By RefCode  Having abs(Sum(Value1)) > 0.000000001  ) As Tr on Tr.RCode=T3.Refcode   Where T3.RECTYPE = 4  AND  T3.METHOD = 1   AND T3.MASTERCODE1 =  (Select MAX(code) From Master1 Where Alias Like'" & ItemCode & "-N" & "%') AND T3.cm1 =" & RefCode & "  AND Date<='" & GetDate(MhDateInput2.Text) & "'   ORDER BY T3.DATE,T3.NO"
       
       Rs.Open "Select T3.vchtype,Tr.* From Tran3 As T3 Inner Join (Select RefCode as RCode,Sum(Value1) as Bal1,Sum(Value2) as Bal2 From Tran3 Where MASTERCODE1= (Select MAX(code) From Master1 Where Alias Like'" & ItemCode & "-N" & "%')  AND Date <= '" & GetDate(MhDateInput2.Text) & "' Group By RefCode  Having abs(Sum(Value1)) > 0.000000001  ) As Tr on Tr.RCode=T3.Refcode   Where T3.RECTYPE = 4  AND  T3.METHOD = 1   AND T3.MASTERCODE1 =  (Select MAX(code) From Master1 Where Alias Like'" & ItemCode & "-N" & "%') AND T3.cm1 =" & RefCode & "  AND Date<='" & GetDate(MhDateInput2.Text) & "'   ORDER BY T3.DATE,T3.NO", CxnProductionPlanning, adOpenKeyset, adLockReadOnly
    
    
    Else
      Rs.Open "Select T3.vchtype,Tr.* From Tran3 As T3 Inner Join (Select RefCode as RCode,Sum(Value1) as Bal1,Sum(Value2) as Bal2 From Tran3 Where MASTERCODE1= (Select MAX(code) From Master1 Where Alias Like'" & ItemCode & "-N" & "%')  AND Date <= '" & GetDate(MhDateInput2.Text) & "' Group By RefCode  Having abs(Sum(Value1)) > 0.000000001  ) As Tr on Tr.RCode=T3.Refcode   Where T3.RECTYPE = 4  AND  T3.METHOD = 1   AND T3.MASTERCODE1 =  (Select MAX(code) From Master1 Where Alias Like'" & ItemCode & "-N" & "%') AND T3.cm1 in((Select Code From Master1 Where ParentGrp in(Select Code From Master1 Where Name='" & StockGroup & "')))  AND Date<='" & GetDate(MhDateInput2.Text) & "' AND T3.MasterCode2 IN (SELECT CODE FROM HELP1 WHERE MASTERSERIES = 26841 AND MasterType= 2)  ORDER BY T3.DATE,T3.NO", CxnProductionPlanning, adOpenKeyset, adLockReadOnly
    End If
    Rs.ActiveConnection = Nothing
    Dim Bal1 As Long
    Dim Bal2 As Long
    If Rs.RecordCount > 0 Then
       Rs.MoveFirst
        Do While Not Rs.EOF
           If StockGroup = "Branch Group" Then
                If Val(Rs.Fields("vchtype").Value) = 13 Then
                  BrPendingPO = BrPendingPO + Val(Rs.Fields("Bal1").Value)
                Else
                  BrPendingSO = BrPendingSO + Val(Rs.Fields("Bal1").Value)
                End If
           Else
                If Val(Rs.Fields("vchtype").Value) = 13 Then
                  PendingPO = PendingPO + Val(Rs.Fields("Bal1").Value)
                Else
                  PendingSO = PendingSO + Val(Rs.Fields("Bal1").Value)
                End If
           End If
           Rs.MoveNext
        Loop
        If StockGroup <> "Branch Group" Then
          CxnDatabase.Execute "UPDATE BookMaster SET Pending_SO=Pending_SO + " & Abs(PendingSO) & ",Pending_PO=Pending_PO+" & Abs(PendingPO) & " WHERE BusyCode='" & ItemCode & "'"
        End If
    End If
    Call CloseRecordset(Rs)
End Sub

Private Sub LoadPartyGroup()
    On Error GoTo ErrorHandler
    Screen.MousePointer = vbHourglass
        Dim ParentGroups As String, CurrentGroups As String
        ParentGroups = "36650": CurrentGroups = "36650"
        Do While True
            If rstAccountList.State = adStateOpen Then rstAccountList.Close
            rstAccountList.Open "SELECT Code FROM Master1 WHERE MasterType=1 AND ParentGrp IN (" & CurrentGroups & ")", CxnProductionPlanning, adOpenKeyset, adLockReadOnly
            If rstAccountList.RecordCount = 0 Then Exit Do
            CurrentGroups = ""
            With rstAccountList
                Do While Not .EOF
                    CurrentGroups = CurrentGroups & IIf(CurrentGroups = "", "'", ", '") & rstAccountList.Fields(0).Value & "'"
                    .MoveNext
                Loop
                CurrentGroups = IIf(CurrentGroups = "", "''", CurrentGroups)
            End With
            ParentGroups = ParentGroups & "," & CurrentGroups
            
        Loop
        PartyParentGroups = ParentGroups
'        If rstAccountList.State = adStateOpen Then rstAccountList.Close
'        Dim aaa As String
'        aaa = "SELECT replace(Name,'''',',')As Name,Code FROM Master1 WHERE ParentGrp IN (" & ParentGroups & ") ORDER BY Name"
'         rstAccountList.Open "SELECT replace(Name,'''',',')As Name,Code FROM Master1 WHERE ParentGrp IN (" & ParentGroups & ") ORDER BY Name", CxnProductionPlanning, adOpenKeyset, adLockReadOnly
'         rstAccountList.ActiveConnection = Nothing
    Screen.MousePointer = vbDefault
    On Error GoTo 0
    Exit Sub
ErrorHandler:
    MsgBox Err.Description & " !!!", vbInformation, "Error"
    Screen.MousePointer = vbDefault: rstAccountList.ActiveConnection = Nothing
End Sub






