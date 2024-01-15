VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Begin VB.Form FrmPaperStockRegister 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Paper Stock Register"
   ClientHeight    =   6915
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
   Icon            =   "PaperStockRegister.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6915
   ScaleWidth      =   7620
   Begin VB.CheckBox Check2 
      Caption         =   "Export To Excel"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   6480
      Width           =   1695
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   9
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
            Picture         =   "PaperStockRegister.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PaperStockRegister.frx":0986
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PaperStockRegister.frx":0A9A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   6065
      Left            =   45
      TabIndex        =   10
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
      Picture         =   "PaperStockRegister.frx":0BAE
      Begin VB.CheckBox Check1 
         Caption         =   "Negative Bal"
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
         Left            =   3840
         TabIndex        =   2
         Top             =   53
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Summarised"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   6180
         TabIndex        =   4
         Top             =   10
         Width           =   1350
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Detailed"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   5160
         TabIndex        =   3
         Top             =   10
         Value           =   -1  'True
         Width           =   1065
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2875
         Left            =   0
         TabIndex        =   5
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
         Left            =   0
         TabIndex        =   11
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
         Picture         =   "PaperStockRegister.frx":0BCA
         Picture         =   "PaperStockRegister.frx":0BE6
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
         Height          =   330
         Left            =   1920
         TabIndex        =   12
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
         Picture         =   "PaperStockRegister.frx":0C02
         Picture         =   "PaperStockRegister.frx":0C1E
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   2880
         Left            =   3755
         TabIndex        =   6
         Top             =   320
         Width           =   3775
         _ExtentX        =   6668
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
         Height          =   2875
         Left            =   0
         TabIndex        =   7
         Top             =   3185
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
      Begin MSComctlLib.ListView ListView4 
         Height          =   2880
         Left            =   3755
         TabIndex        =   8
         Top             =   3185
         Width           =   3775
         _ExtentX        =   6668
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
      Begin TDBDate6Ctl.TDBDate MhDateInput1 
         Height          =   330
         Left            =   840
         TabIndex        =   0
         Top             =   0
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calendar        =   "PaperStockRegister.frx":0C3A
         Caption         =   "PaperStockRegister.frx":0D52
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "PaperStockRegister.frx":0DBE
         Keys            =   "PaperStockRegister.frx":0DDC
         Spin            =   "PaperStockRegister.frx":0E3A
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
      Begin TDBDate6Ctl.TDBDate MhDateInput2 
         Height          =   330
         Left            =   2670
         TabIndex        =   1
         Top             =   0
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calendar        =   "PaperStockRegister.frx":0E62
         Caption         =   "PaperStockRegister.frx":0F7A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "PaperStockRegister.frx":0FE6
         Keys            =   "PaperStockRegister.frx":1004
         Spin            =   "PaperStockRegister.frx":1062
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
   End
End
Attribute VB_Name = "FrmPaperStockRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstCompanyMaster As New ADODB.Recordset
Dim rstPaperStockRegister As New ADODB.Recordset
Dim rstPaperSizeList As New ADODB.Recordset
Dim rstPaperGSMList As New ADODB.Recordset
Dim rstPaperList As New ADODB.Recordset
Dim rstAccountList As New ADODB.Recordset
Dim EMailID As String
Dim Attachment As String
Dim Message As String
Dim OutputTo As String
Dim PaperTbl

Private Sub Form_Load()
    On Error GoTo ErrorHandler
    CenterForm Me
    BusySystemIndicator True
    PaperTbl = "SELECT Code As Paper FROM PaperChild UNION " & _
                      "SELECT Paper FROM PaperIOChild UNION " & _
                      "SELECT Item As Paper FROM MaterialSVChild WHERE Category='2' UNION " & _
                      "SELECT Paper FROM PaperMVChild UNION " & _
                      "SELECT Paper FROM PaperDNChild UNION " & _
                      "SELECT Item As Paper FROM BookPOParent P INNER JOIN BookPOChild0801 C ON P.Code=C.Code WHERE C.Category='2' AND P.Type<>'O' AND LEFT(P.Code,1)<>'*' UNION " & _
                      "SELECT Paper FROM BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code WHERE P.Type<>'O' AND LEFT(P.Code,1)<>'*' UNION " & _
                      "SELECT Paper1 As Paper FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE P.Type<>'O' AND LEFT(P.Code,1)<>'*' UNION " & _
                      "SELECT Paper2 As Paper FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE P.Type<>'O' AND LEFT(P.Code,1)<>'*' UNION " & _
                      "SELECT Paper4 As Paper FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE P.Type<>'O' AND LEFT(P.Code,1)<>'*'"
    rstCompanyMaster.Open "SELECT PrintName,Phone,eMail FROM CompanyMaster", CxnDatabase, adOpenKeyset, adLockReadOnly
    rstPaperSizeList.Open "SELECT Name,Code FROM GeneralMaster WHERE Code IN (SELECT [Size] FROM PaperMaster P INNER JOIN (" & PaperTbl & ") As C ON P.Code=C.Paper) ORDER BY Name", CxnDatabase, adOpenKeyset, adLockReadOnly
    rstPaperSizeList.ActiveConnection = Nothing
    Call FillList(ListView1, "List of Paper Sizes...", rstPaperSizeList)
    Call LoadPaperGSMList
    Call LoadPaperList
    Call LoadAccountList
    Option1.Value = True
    MhDateInput1.Text = Format(FinancialYearFrom, "dd-mm-yyyy")
    MhDateInput2.Text = Format(IIf(Format(FinancialYearTo, "yyyymmdd") < Format(Date, "yyyymmdd"), FinancialYearTo, Date), "dd-mm-yyyy")
    BusySystemIndicator False
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    Call CloseForm(FrmPaperStockRegister)
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
    If UnloadMode = 0 Then Call CloseForm(FrmPaperStockRegister)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CloseRecordset(rstCompanyMaster)
    Call CloseRecordset(rstPaperSizeList)
    Call CloseRecordset(rstPaperGSMList)
    Call CloseRecordset(rstPaperList)
    Call CloseRecordset(rstAccountList)
    Call CloseRecordset(rstPaperStockRegister)
End Sub
Private Sub MhDateInput1_Validate(Cancel As Boolean)
    If Not IsDate(GetDate(MhDateInput1.Text)) Then Cancel = True
End Sub
Private Sub MhDateInput2_Validate(Cancel As Boolean)
    If Not IsDate(GetDate(MhDateInput2.Text)) Then
        Cancel = True
    ElseIf Format(GetDate(MhDateInput2.Text), "yyyymmdd") < Format(GetDate(MhDateInput1.Text), "yyyymmdd") Then
        FocusSelect Me.ActiveControl
        Cancel = True
    ElseIf Format(GetDate(MhDateInput2.Text), "yyyymmdd") > Format(FinancialYearTo, "yyyymmdd") Then
        Cancel = True
    End If
End Sub
Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Call LoadPaperGSMList
End Sub
Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    If (KeyCode = vbKeyA Or KeyCode = vbKeyD) And Shift = vbCtrlMask Then
        For i = 1 To ListView1.ListItems.Count
            ListView1.ListItems(i).Checked = IIf(KeyCode = vbKeyA, True, False)
        Next i
        Call LoadPaperGSMList
    End If
End Sub
Private Sub ListView2_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Call LoadPaperList
End Sub
Private Sub ListView2_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    If (KeyCode = vbKeyA Or KeyCode = vbKeyD) And Shift = vbCtrlMask Then
        For i = 1 To ListView2.ListItems.Count
            ListView2.ListItems(i).Checked = IIf(KeyCode = vbKeyA, True, False)
        Next i
        Call LoadPaperList
    End If
End Sub
Private Sub ListView3_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Call LoadAccountList
End Sub
Private Sub ListView3_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    If (KeyCode = vbKeyA Or KeyCode = vbKeyD) And Shift = vbCtrlMask Then
        For i = 1 To ListView3.ListItems.Count
            ListView3.ListItems(i).Checked = IIf(KeyCode = vbKeyA, True, False)
        Next i
        Call LoadAccountList
    End If
End Sub
Private Sub ListView4_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    If (KeyCode = vbKeyA Or KeyCode = vbKeyD) And Shift = vbCtrlMask Then
        For i = 1 To ListView4.ListItems.Count
            ListView4.ListItems(i).Checked = IIf(KeyCode = vbKeyA, True, False)
        Next i
    End If
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    If Button.Index = 1 Then
        OutputTo = "S"
        PrintPaperStockRegister
    ElseIf Button.Index = 2 Then
        OutputTo = "P"
        PrintPaperStockRegister
    ElseIf Button.Index = 3 Then
        Call CloseForm(FrmPaperStockRegister)
    End If
End Sub
Private Sub LoadPaperGSMList()
    Dim SelectedPaperSizes
    If rstPaperGSMList.State = adStateOpen Then rstPaperGSMList.Close
    SelectedPaperSizes = SelectedItems(ListView1)
    rstPaperGSMList.Open "SELECT DISTINCT GSM As Name,STR(GSM) As Code FROM PaperMaster P INNER JOIN (" & PaperTbl & ") As C ON P.Code=C.Paper WHERE " & IIf(SelectedPaperSizes = "''", "1=1", "[Size] IN (" & SelectedPaperSizes & ")") & " ORDER BY GSM", CxnDatabase, adOpenKeyset, adLockReadOnly
    rstPaperGSMList.ActiveConnection = Nothing
    ListView2.ListItems.Clear
    Call FillList(ListView2, "List of Paper GSMs...", rstPaperGSMList)
End Sub
Private Sub LoadAccountList()
    Dim SelectedPapers, AccountTbl
    If rstAccountList.State = adStateOpen Then rstAccountList.Close
    SelectedPapers = SelectedItems(ListView3)
    AccountTbl = "SELECT Account FROM PaperChild WHERE " & IIf(SelectedPapers = "''", "1=1", "Code IN (" & SelectedPapers & ")") & " UNION " & _
                          "SELECT Account FROM PaperIOChild WHERE " & IIf(SelectedPapers = "''", "1=1", "Paper IN (" & SelectedPapers & ")") & " UNION " & _
                          "SELECT Account FROM MaterialSVParent P INNER JOIN MaterialSVChild C ON P.Code=C.Code WHERE C.Category='2' AND " & IIf(SelectedPapers = "''", "1=1", "C.Item IN (" & SelectedPapers & ")") & " UNION " & _
                          "SELECT AccountFrom As Account FROM PaperMVParent P INNER JOIN PaperMVChild C ON P.Code=C.Code WHERE " & IIf(SelectedPapers = "''", "1=1", "C.Paper IN (" & SelectedPapers & ")") & " UNION " & _
                          "SELECT AccountTo As Account FROM PaperMVParent P INNER JOIN PaperMVChild C ON P.Code=C.Code WHERE " & IIf(SelectedPapers = "''", "1=1", "C.Paper IN (" & SelectedPapers & ")") & " UNION " & _
                          "SELECT Account FROM PaperDNParent P INNER JOIN PaperDNChild C ON P.Code=C.Code WHERE " & IIf(SelectedPapers = "''", "1=1", "C.Paper IN (" & SelectedPapers & ")") & " UNION " & _
                          "SELECT Binder As Account FROM BookPOParent P INNER JOIN BookPOChild0801 C ON P.Code=C.Code WHERE C.Category='2' AND P.Type<>'O' AND LEFT(P.Code,1)<>'*' AND " & IIf(SelectedPapers = "''", "1=1", "C.Item IN (" & SelectedPapers & ")") & " UNION " & _
                          "SELECT TitlePrinter As Account FROM BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code WHERE P.Type<>'O' AND LEFT(P.Code,1)<>'*' AND " & IIf(SelectedPapers = "''", "1=1", "C.Paper IN (" & SelectedPapers & ")") & " UNION " & _
                          "SELECT BookPrinter As Account FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE P.Type<>'O' AND LEFT(P.Code,1)<>'*' AND " & IIf(SelectedPapers = "''", "1=1", "C.Paper1 IN (" & SelectedPapers & ")") & " UNION " & _
                          "SELECT BookPrinter As Account FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE P.Type<>'O' AND LEFT(P.Code,1)<>'*' AND " & IIf(SelectedPapers = "''", "1=1", "C.Paper2 IN (" & SelectedPapers & ")") & " UNION " & _
                          "SELECT BookPrinter As Account FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE P.Type<>'O' AND LEFT(P.Code,1)<>'*' AND " & IIf(SelectedPapers = "''", "1=1", "C.Paper4 IN (" & SelectedPapers & ")")
                          
        'Shams need to check here
        
        'Dim AA As String
        'AA = "SELECT Name,Code FROM AccountMaster P INNER JOIN (" & AccountTbl & ") As C ON P.Code=C.Account ORDER BY Name"
        
    rstAccountList.Open "SELECT Name,Code FROM AccountMaster P INNER JOIN (" & AccountTbl & ") As C ON P.Code=C.Account ORDER BY Name", CxnDatabase, adOpenKeyset, adLockReadOnly
    rstAccountList.ActiveConnection = Nothing
    ListView4.ListItems.Clear
    Call FillList(ListView4, "List of Godowns...", rstAccountList)
End Sub
Private Sub LoadPaperList()
    Dim SelectedPaperGSMs, SelectedPaperSizes
    If rstPaperList.State = adStateOpen Then rstPaperList.Close
    SelectedPaperSizes = SelectedItems(ListView1)
    SelectedPaperGSMs = SelectedItems(ListView2)
    rstPaperList.Open "SELECT Name,Code FROM PaperMaster P INNER JOIN (" & PaperTbl & ") As C ON P.Code=C.Paper WHERE " & IIf(SelectedPaperSizes = "''" Or SelectedPaperGSMs = "''", "1=1", "[Size] IN (" & SelectedPaperSizes & ") AND STR(GSM) IN (" & SelectedPaperGSMs & ")") & " ORDER BY Name", CxnDatabase, adOpenKeyset, adLockReadOnly
    rstPaperList.ActiveConnection = Nothing
    ListView3.ListItems.Clear
    Call FillList(ListView3, "List of Papers...", rstPaperList)
End Sub
Private Sub PrintPaperStockRegister()
    Dim CRXParamDefs As CRAXDRT.ParameterFieldDefinitions
    Dim CRXParamDef As CRAXDRT.ParameterFieldDefinition
    Dim OpBalQry
    Dim SelectedPapers As String
    Dim SelectedAccounts As String
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    
    If Check2.Value = 1 Then
        rptPaperStockRegisterExcel.Text11.SetText "Paper Stock Register (" & IIf(Option1.Value, "Detailed", "Summarised") & ")"
        rptPaperStockRegisterExcel.Text12.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
        rptPaperStockRegisterExcel.Text13.SetText "From [" + Format(GetDate(MhDateInput1.Text), "dd-mm-yyyy") + "] To [" + Format(GetDate(MhDateInput2.Text), "dd-mm-yyyy") + "]"
    Else
    
        rptPaperStockRegister.Text11.SetText "Paper Stock Register (" & IIf(Option1.Value, "Detailed", "Summarised") & ")"
        rptPaperStockRegister.Text12.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
        rptPaperStockRegister.Text13.SetText "From [" + Format(GetDate(MhDateInput1.Text), "dd-mm-yyyy") + "] To [" + Format(GetDate(MhDateInput2.Text), "dd-mm-yyyy") + "]"
    End If
    
    rptPaperStockRegister.Text11.SetText "Paper Stock Register (" & IIf(Option1.Value, "Detailed", "Summarised") & ")"
    rptPaperStockRegister.Text12.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptPaperStockRegister.Text13.SetText "From [" + Format(GetDate(MhDateInput1.Text), "dd-mm-yyyy") + "] To [" + Format(GetDate(MhDateInput2.Text), "dd-mm-yyyy") + "]"
    
    
    
    If rstPaperStockRegister.State = adStateOpen Then rstPaperStockRegister.Close
    SelectedPapers = SelectedItems(ListView3)
    SelectedAccounts = SelectedItems(ListView4)
    OpBalQry = "" & _
    "(SELECT IIF(ISNULL(SUM(OpBalSheets)),0,SUM(OpBalSheets)) FROM PaperChild WHERE Code=M2.Code AND Account=M1.Code)+" & _
    "(SELECT IIF(ISNULL(SUM(QuantitySheets)),0,SUM(QuantitySheets)) FROM PaperPOParent P INNER JOIN PaperIOChild C ON P.Code=C.Code WHERE Paper=M2.Code AND Account=M1.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)+" & _
    "(SELECT IIF(ISNULL(SUM(Quantity)),0,SUM(INT(Quantity)*500+(Quantity-INT(Quantity))*1000)) FROM MaterialSVParent P INNER JOIN MaterialSVChild C ON P.Code=C.Code WHERE Category='2' AND Item=M2.Code AND Quantity>=0 AND Account=M1.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)-" & _
    "(SELECT IIF(ISNULL(SUM(Quantity)),0,ABS(SUM(FIX(Quantity)*500+(Quantity-FIX(Quantity))*1000))) FROM MaterialSVParent P INNER JOIN MaterialSVChild C ON P.Code=C.Code WHERE Category='2' AND Item=M2.Code AND Quantity<0 AND Account=M1.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)-" & _
    "(SELECT IIF(ISNULL(SUM(QuantitySheets)),0,SUM(QuantitySheets)) FROM PaperMVParent P INNER JOIN PaperMVChild C ON P.Code=C.Code WHERE Paper=M2.Code AND AccountFrom=M1.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)+" & _
    "(SELECT IIF(ISNULL(SUM(QuantitySheets)),0,SUM(QuantitySheets)) FROM PaperMVParent P INNER JOIN PaperMVChild C ON P.Code=C.Code WHERE Paper=M2.Code AND AccountTo=M1.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)-" & _
    "(SELECT IIF(ISNULL(SUM(Quantity)),0,SUM(INT(Quantity)*500+(Quantity-INT(Quantity))*1000)) FROM PaperDNParent P INNER JOIN PaperDNChild C ON P.Code=C.Code WHERE Paper=M2.Code AND Account=M1.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)-" & _
    "(SELECT IIF(ISNULL(SUM(Round(ActualQuantity*C1.Quantity,0))),0,SUM(Round(ActualQuantity*C1.Quantity,0))) FROM (BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code) INNER JOIN BookPOChild0801 C1 ON C.Code=C1.Code WHERE P.Type <> 'O' AND LEFT(P.Code,1)<>'*' AND Category='2' AND Item=M2.Code AND Binder=M1.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)-" & _
    "(SELECT IIF(ISNULL(SUM(PaperConsumptionSheets)),0,SUM(PaperConsumptionSheets)) FROM BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code WHERE P.Type <> 'O' AND LEFT(P.Code,1)<>'*' AND Paper=M2.Code AND TitlePrinter=M1.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)-" & _
    "(SELECT IIF(ISNULL(SUM(PaperConsumptionSheets1)),0,SUM(PaperConsumptionSheets1)) FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE P.Type <> 'O' AND LEFT(P.Code,1)<>'*' AND Paper1=M2.Code AND BookPrinter=M1.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)-" & _
    "(SELECT IIF(ISNULL(SUM(PaperConsumptionSheets2)),0,SUM(PaperConsumptionSheets2)) FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE P.Type <> 'O' AND LEFT(P.Code,1)<>'*' AND Paper2=M2.Code AND BookPrinter=M1.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)-" & _
    "(SELECT IIF(ISNULL(SUM(PaperConsumptionSheets4)),0,SUM(PaperConsumptionSheets4)) FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE P.Type <> 'O' AND LEFT(P.Code,1)<>'*' AND Paper4=M2.Code AND BookPrinter=M1.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)"
    'VchNo,VchDate,VchType,Particulars,BookQuantity,Forms,Quantity,GSM,GodownName,SizeName,PaperName
    
    
    rstPaperStockRegister.Open "" & _
    "SELECT * FROM (SELECT '' As VchNo,#" & CDate(GetDate(MhDateInput1.Text)) - 1 & "# As VchDate,'OB' As VchType,'Opening Balance' As Particulars,0 As BookQuantity,0.00 As Forms," & OpBalQry & " As Quantity,M2.GSM,'Godown Name : '+TRIM(M1.PrintName) As GodownName,'Size Name : '+(SELECT TRIM(PrintName) FROM GeneralMaster WHERE Code=M2.[Size]) As SizeName,'Paper Name : '+TRIM(M2.PrintName) As PaperName,0 As QtyRecd,'' As Remarks FROM AccountMaster M1,PaperMaster M2 WHERE M1.Type IN ('05','06','08','09') AND M1.Code IN (" & SelectedAccounts & ") AND M2.Code IN (" & SelectedPapers & ")) As Tbl WHERE Quantity<>0 UNION ALL " & _
    "SELECT TRIM(P.Name) As VchNo,P.Date As VchDate,'PI' As VchType,'Paper IN (FROM : '+(SELECT TRIM(PrintName) FROM AccountMaster WHERE Code=P.Supplier)+')' As Particulars,0 As BookQuantity,0.00 As Forms,QuantitySheets As Quantity,M2.GSM,'Godown Name : '+TRIM(M1.PrintName) As GodownName,'Size Name : '+(SELECT TRIM(PrintName) FROM GeneralMaster WHERE Code=M2.[Size]) As SizeName,'Paper Name : '+TRIM(M2.PrintName) As PaperName,0 As QtyRecd,'' As Remarks FROM ((PaperPOParent P INNER JOIN PaperIOChild C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON C.Account=M1.Code) INNER JOIN PaperMaster M2 ON C.Paper=M2.Code WHERE M1.Code IN (" & SelectedAccounts & ") AND M2.Code IN (" & SelectedPapers & ") AND P.Date>=#" & GetDate(MhDateInput1.Text) & "# AND P.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
    "SELECT TRIM(P.Name) As VchNo,P.Date As VchDate,'SI' As VchType,'Stock Journal (Generated)' As Particulars,0 As BookQuantity,0.00 As Forms,INT(Quantity)*500+(Quantity-INT(Quantity))*1000 As Quantity,M2.GSM,'Godown Name : '+TRIM(M1.PrintName) As GodownName,'Size Name : '+(SELECT TRIM(PrintName) FROM GeneralMaster WHERE Code=M2.[Size]) As SizeName,'Paper Name : '+TRIM(M2.PrintName) As PaperName,0 As QtyRecd,'' As Remarks FROM ((MaterialSVParent P INNER JOIN MaterialSVChild C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON P.Account=M1.Code) INNER JOIN PaperMaster M2 ON C.Item=M2.Code WHERE C.Category='2' AND C.Quantity>=0 AND M1.Code IN (" & SelectedAccounts & ") AND M2.Code IN (" & SelectedPapers & ") AND P.Date>=#" & GetDate(MhDateInput1.Text) & "# AND P.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
    "SELECT TRIM(P.Name) As VchNo,P.Date As VchDate,'SR' As VchType,'Stock Journal (Consumed)' As Particulars,0 As BookQuantity,0.00 As Forms,ABS(Fix(Quantity)*500+(Quantity-Fix(Quantity))*1000) As Quantity,M2.GSM,'Godown Name : '+TRIM(M1.PrintName) As GodownName,'Size Name : '+(SELECT TRIM(PrintName) FROM GeneralMaster WHERE Code=M2.[Size]) As SizeName,'Paper Name : '+TRIM(M2.PrintName) As PaperName,0 As QtyRecd,'' As Remarks FROM ((MaterialSVParent P INNER JOIN MaterialSVChild C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON P.Account=M1.Code) INNER JOIN PaperMaster M2 ON C.Item=M2.Code WHERE C.Category='2' AND C.Quantity<0 AND M1.Code IN (" & SelectedAccounts & ") AND M2.Code IN (" & SelectedPapers & ") AND P.Date>=#" & GetDate(MhDateInput1.Text) & "# AND P.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
    "SELECT TRIM(P.Name) As VchNo,P.Date As VchDate,'MO' As VchType,'Paper Out (To : '+(SELECT TRIM(PrintName) FROM AccountMaster WHERE Code=P.AccountTo)+')' As Particulars,0 As BookQuantity,0.00 As Forms,QuantitySheets As Quantity,M2.GSM,'Godown Name : '+TRIM(M1.PrintName) As GodownName,'Size Name : '+(SELECT TRIM(PrintName) FROM GeneralMaster WHERE Code=M2.[Size]) As SizeName,'Paper Name : '+TRIM(M2.PrintName) As PaperName,0 As QtyRecd,'' As Remarks FROM ((PaperMVParent P INNER JOIN PaperMVChild C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON P.AccountFrom=M1.Code) INNER JOIN PaperMaster M2 ON C.Paper=M2.Code WHERE M1.Code IN (" & SelectedAccounts & ") AND M2.Code IN (" & SelectedPapers & ") AND P.Date>=#" & GetDate(MhDateInput1.Text) & "# AND P.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
    "SELECT TRIM(P.Name) As VchNo,P.Date As VchDate,'MI' As VchType,'Paper IN (FROM : '+(SELECT TRIM(PrintName) FROM AccountMaster WHERE Code=P.AccountFrom)+')' As Particulars,0 As BookQuantity,0.00 As Forms,QuantitySheets As Quantity,M2.GSM,'Godown Name : '+TRIM(M1.PrintName) As GodownName,'Size Name : '+(SELECT TRIM(PrintName) FROM GeneralMaster WHERE Code=M2.[Size]) As SizeName,'Paper Name : '+TRIM(M2.PrintName) As PaperName,0 As QtyRecd,'' As Remarks FROM ((PaperMVParent P INNER JOIN PaperMVChild C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON P.AccountTo=M1.Code) INNER JOIN PaperMaster M2 ON C.Paper=M2.Code WHERE M1.Code IN (" & SelectedAccounts & ") AND M2.Code IN (" & SelectedPapers & ") AND P.Date>=#" & GetDate(MhDateInput1.Text) & "# AND P.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
    "SELECT TRIM(P.Name) As VchNo,P.Date As VchDate,'DN' As VchType,'Debit Note' As Particulars,0 As BookQuantity,0.00 As Forms,INT(Quantity)*500+(Quantity-INT(Quantity))*1000 As Quantity,M2.GSM,'Godown Name : '+TRIM(M1.PrintName) As GodownName,'Size Name : '+(SELECT TRIM(PrintName) FROM GeneralMaster WHERE Code=M2.[Size]) As SizeName,'Paper Name : '+TRIM(M2.PrintName) As PaperName,0 As QtyRecd,'' As Remarks FROM ((PaperDNParent P INNER JOIN PaperDNChild C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON P.Account=M1.Code) INNER JOIN PaperMaster M2 ON C.Paper=M2.Code WHERE M1.Code IN (" & SelectedAccounts & ") AND M2.Code IN (" & SelectedPapers & ") AND P.Date>=#" & GetDate(MhDateInput1.Text) & "# AND P.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
    "SELECT TRIM(P.Name) As VchNo,C.OrderDate As VchDate,'PC' As VchType,'Paper Consumed (Binding : '+(SELECT TRIM(PrintName) FROM BookMaster WHERE Code=P.Book)+')' As Particulars,ActualQuantity As BookQuantity,0.00 As Forms,ROUND(ActualQuantity*C1.Quantity,0) As Quantity,M2.GSM,'Godown Name : '+TRIM(M1.PrintName) As GodownName,'Size Name : '+(SELECT TRIM(PrintName) FROM GeneralMaster WHERE Code=M2.[Size]) As SizeName,'Paper Name : '+TRIM(M2.PrintName) As PaperName,0 As QtyRecd,'' As Remarks FROM (((BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code) INNER JOIN BookPOChild0801 C1 ON C.Code=C1.Code) INNER JOIN AccountMaster M1 ON P.Binder=M1.Code) INNER JOIN PaperMaster M2 ON C1.Item=M2.Code WHERE P.Type <> 'O' AND LEFT(P.Code,1)<>'*' AND C1.Category='2' AND M1.Code IN (" & SelectedAccounts & ") AND M2.Code IN (" & SelectedPapers & ") AND C.OrderDate>=#" & GetDate(MhDateInput1.Text) & "# AND C.OrderDate<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
    "SELECT TRIM(P.Name) As VchNo,C.OrderDate As VchDate,'PC' As VchType,'Paper Consumed (Title : '+(SELECT TRIM(PrintName) FROM BookMaster WHERE Code=P.Book)+') Wastage-'+TRIM(STR([PaperWastage%])+'%,Qty Recd:'+str(ReceivedQuantity)) As Particulars,ActualQuantity As BookQuantity,0.00 As Forms,PaperConsumptionSheets As Quantity,M2.GSM,'Godown Name : '+TRIM(M1.PrintName) As GodownName,'Size Name : '+(SELECT TRIM(PrintName) FROM GeneralMaster WHERE Code=M2.[Size]) As SizeName,'Paper Name : '+TRIM(M2.PrintName) As PaperName,ReceivedQuantity As QtyRecd,'' As Remarks FROM ((BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON P.TitlePrinter=M1.Code) INNER JOIN PaperMaster M2 ON C.Paper=M2.Code WHERE P.Type <> 'O' AND LEFT(P.Code,1)<>'*' AND M1.Code IN (" & SelectedAccounts & ") AND M2.Code IN (" & SelectedPapers & ") AND C.OrderDate>=#" & GetDate(MhDateInput1.Text) & "# AND C.OrderDate<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
    "SELECT TRIM(P.Name) As VchNo,C.OrderDate As VchDate,'PC' As VchType,'Paper Consumed (Book : '+(SELECT TRIM(PrintName) FROM BookMaster WHERE Code=P.Book)+') Wastage-'+TRIM(STR([PaperWastage1%])+'%,Qty Recd:'+str(ReceivedQuantity)) As Particulars,ActualQuantity As BookQuantity,Forms1 As Forms,PaperConsumptionSheets1 As Quantity,M2.GSM,'Godown Name : '+TRIM(M1.PrintName) As GodownName,'Size Name : '+(SELECT TRIM(PrintName) FROM GeneralMaster WHERE Code=M2.[Size]) As SizeName,'Paper Name : '+TRIM(M2.PrintName) As PaperName,ReceivedQuantity As QtyRecd,Remarks As Remarks FROM ((BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON P.BookPrinter=M1.Code) INNER JOIN PaperMaster M2 ON C.Paper1=M2.Code WHERE P.Type <> 'O' AND LEFT(P.Code,1)<>'*' AND M1.Code IN (" & SelectedAccounts & ") AND M2.Code IN (" & SelectedPapers & ") AND C.OrderDate>=#" & GetDate(MhDateInput1.Text) & "# AND C.OrderDate<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
    "SELECT TRIM(P.Name) As VchNo,C.OrderDate As VchDate,'PC' As VchType,'Paper Consumed (Book : '+(SELECT TRIM(PrintName) FROM BookMaster WHERE Code=P.Book)+') Wastage-'+TRIM(STR([PaperWastage2%])+'%,Qty Recd:'+str(ReceivedQuantity)) As Particulars,ActualQuantity As BookQuantity,Forms2 As Forms,PaperConsumptionSheets2 As Quantity,M2.GSM,'Godown Name : '+TRIM(M1.PrintName) As GodownName,'Size Name : '+(SELECT TRIM(PrintName) FROM GeneralMaster WHERE Code=M2.[Size]) As SizeName,'Paper Name : '+TRIM(M2.PrintName) As PaperName,ReceivedQuantity As QtyRecd,Remarks As Remarks FROM ((BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON P.BookPrinter=M1.Code) INNER JOIN PaperMaster M2 ON C.Paper2=M2.Code WHERE P.Type <> 'O' AND LEFT(P.Code,1)<>'*' AND M1.Code IN (" & SelectedAccounts & ") AND M2.Code IN (" & SelectedPapers & ") AND C.OrderDate>=#" & GetDate(MhDateInput1.Text) & "# AND C.OrderDate<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
    "SELECT TRIM(P.Name) As VchNo,C.OrderDate As VchDate,'PC' As VchType,'Paper Consumed (Book : '+(SELECT TRIM(PrintName) FROM BookMaster WHERE Code=P.Book)+') Wastage-'+TRIM(STR([PaperWastage4%])+'%,Qty Recd:'+str(ReceivedQuantity)) As Particulars,ActualQuantity As BookQuantity,Forms4 As Forms,PaperConsumptionSheets4 As Quantity,M2.GSM,'Godown Name : '+TRIM(M1.PrintName) As GodownName,'Size Name : '+(SELECT TRIM(PrintName) FROM GeneralMaster WHERE Code=M2.[Size]) As SizeName,'Paper Name : '+TRIM(M2.PrintName) As PaperName,ReceivedQuantity As QtyRecd,Remarks As Reamrks FROM ((BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON P.BookPrinter=M1.Code) INNER JOIN PaperMaster M2 ON C.Paper4=M2.Code WHERE P.Type <> 'O' AND LEFT(P.Code,1)<>'*' AND M1.Code IN (" & SelectedAccounts & ") AND M2.Code IN (" & SelectedPapers & ") AND C.OrderDate>=#" & GetDate(MhDateInput1.Text) & "# AND C.OrderDate<=#" & GetDate(MhDateInput2.Text) & "# " & _
    "ORDER BY GodownName,SizeName,PaperName,VchDate,VchNo", CxnDatabase, adOpenKeyset, adLockOptimistic
    
    
    Screen.MousePointer = vbNormal
    If rstPaperStockRegister.RecordCount = 0 Then
        On Error GoTo 0
        Exit Sub
    End If
    
    If Check2.Value = 1 Then
        'Call Export2Excel(rstPaperStockRegister)
            Dim oExcel As Object
            Dim xlsWB1 As Object
             Set oExcel = CreateObject("Excel.Application")
            
            rptPaperStockRegisterExcel.Database.SetDataSource rstPaperStockRegister, 3, 1
            rptPaperStockRegisterExcel.DiscardSavedData
            Set CRXParamDefs = rptPaperStockRegisterExcel.ParameterFields
            For Each CRXParamDef In CRXParamDefs
            If CRXParamDef.ParameterFieldName = "PF1" Then
               CRXParamDef.SetCurrentValue (IIf(Check1.Value, 0, 0.1))
            ElseIf CRXParamDef.ParameterFieldName = "PF2" Then
               CRXParamDef.SetCurrentValue (IIf(Option1.Value, "D", "S"))
            End If
            Next
            rptPaperStockRegisterExcel.EnableParameterPrompting = False
            rptPaperStockRegisterExcel.ExportOptions.DestinationType = crEDTDiskFile
            'rptPaperStockRegisterExcel.ExportOptions.DiskFileName = "D:\Rackserver\Paper Stock Register.xls"
            rptPaperStockRegisterExcel.ExportOptions.DiskFileName = (App.Path & "\Report\Paper Stock Register (" & CompCode & ")")
            rptPaperStockRegisterExcel.ExportOptions.FormatType = crEFTExcel50
            rptPaperStockRegisterExcel.ExportOptions.ExcelUseConstantColumnWidth = True
            rptPaperStockRegisterExcel.Export False
            oExcel.Visible = True
            
            oExcel.Columns(8).NumberFormat = "General"
            oExcel.Columns(9).NumberFormat = "General"
            oExcel.Columns(10).NumberFormat = "General"
            oExcel.Columns(11).NumberFormat = "General"
            
            
            
            Set xlsWB1 = oExcel.Workbooks.Open((App.Path & "\Report\Paper Stock Register (" & CompCode & ")"))
            
            'Columns(1).NumberFormat = "@"'Text
            'Columns(8).NumberFormat = "General"'General
            'Columns(3).NumberFormat = "0"'Number
            
            oExcel.Columns(9).NumberFormat = "General"
            oExcel.Columns(10).NumberFormat = "General"
            oExcel.Columns(11).NumberFormat = "General"
            

'            oExcel.Application.DisplayAlerts = True
'            oExcel.Sheets("Sheet1").Select
'            oExcel.Sheets("Sheet1").Activate
'            oExcel.Columns("A:M").EntireColumn.AutoFit
'            oExcel.Workbooks.Item(1).Save
    Else
        rptPaperStockRegister.Database.SetDataSource rstPaperStockRegister, 3, 1
        rptPaperStockRegister.DiscardSavedData
        Set CRXParamDefs = rptPaperStockRegister.ParameterFields
        For Each CRXParamDef In CRXParamDefs
            If CRXParamDef.ParameterFieldName = "PF1" Then
                CRXParamDef.SetCurrentValue (IIf(Check1.Value, 0, 0.1))
            ElseIf CRXParamDef.ParameterFieldName = "PF2" Then
                CRXParamDef.SetCurrentValue (IIf(Option1.Value, "D", "S"))
            End If
        Next
        rptPaperStockRegister.EnableParameterPrompting = False
        EMailID = "xxxxxxxxxx"
        Attachment = "Paper Stock Register"
        Message = "Dear Sir,<Br>Please find attached herewith Paper Stock Register From [" & Format(GetDate(MhDateInput1.Text), "dd-mm-yyyy") & "] To [" & Format(GetDate(MhDateInput2.Text), "dd-mm-yyyy") & "] for doing the needful at your end.<Br>Kindly inform us if you find any discrepancy in the same and acknowledge the receipt of mail.<Br><Br>" & Trim(rstCompanyMaster.Fields("PrintName").Value) & "<Br>Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & "<Br>E-Mail : <a HRef='mailto:" & Trim(rstCompanyMaster.Fields("EMail").Value) & "'>" & Trim(rstCompanyMaster.Fields("EMail").Value) & "</a>"
        If OutputTo = "S" Then
            
            FrmReportViewer.EMailID = EMailID
            FrmReportViewer.Subject = "Paper Stock Register"
            FrmReportViewer.Attachment = Attachment
            FrmReportViewer.Message = Message
            Set FrmReportViewer.Report = rptPaperStockRegister
            FrmReportViewer.Show vbModal
        Else
            rptPaperStockRegister.PrintOut
        End If
        Set rptPaperStockRegister = Nothing
        Set rptPaperStockRegisterExcel = Nothing
    End If
    On Error GoTo 0
End Sub

Private Sub Export2Excel(ByVal RsReaderInfo As ADODB.Recordset)

    
    Dim OFF1 As Integer: Dim OFF2 As Double: Dim OFF3 As Double: Dim OFF4 As Double: Dim OFF5 As Double: Dim OFF6 As Double
    
    Dim TOFF3 As Double
    Dim TOFF4 As Double
    Dim TBalance As Double
    Dim OFF7 As Double
    Dim OFF8 As Double
    
    Dim TOFF23 As Double
    Dim TOFF24 As Double
    
    Dim TOFF14 As Double
    Dim TOFF15 As Double
    
    Dim TOFF17 As Double
    Dim TOFF18 As Double
    
    Dim TOFF20 As Double
    Dim TOFF21 As Double
    
    
    
    Dim oExcel As Object
    Dim R As Long
    Dim Cnt As Long
    Dim TotalBundles As Double
    Dim PaperName As String
    Dim GSM As String
    Dim SizeName As String
    Dim GodownName As String
    
    
    On Error Resume Next
    If Not FileExist(App.Path & "\Template\Paper Stock Register.xlsx") Then Exit Sub
    Screen.MousePointer = vbHourglass
    
    If RsReaderInfo.RecordCount = 0 Then Screen.MousePointer = vbNormal: On Error GoTo 0: Exit Sub
    DoEvents
    Set oExcel = CreateObject("Excel.Application")
    oExcel.Workbooks.Open (App.Path & "\Template\Paper Stock Register")
    oExcel.DisplayAlerts = False
    oExcel.Workbooks.Item(1).SaveAs (App.Path & "\Report\Paper Stock Register (" & CompCode & ")")
    
    oExcel.Application.DisplayAlerts = True
    oExcel.Sheets("Sheet1").Select
    oExcel.Application.Cells(1, 1).Value = "RACHNA SAGAR (P) LTD."
    oExcel.Application.Cells(2, 1).Value = "Paper Stock Register From [" & Format(MhDateInput1, "dd-MMM-yyyy") & "] To [" & Format(MhDateInput2, "dd-MMM-yyyy") & "]"
    
    R = 4: Cnt = 1
    oExcel.Application.Cells(R, 4).Value = Trim(RsReaderInfo.Fields("PaperName").Value)
    oExcel.Rows(R).Font.Bold = True
    R = 5: Cnt = 1
    

'    'Fix(Sum({@FF6},{ado.PaperName})/500)+Remainder(Sum({@FF6},{ado.PaperName}),500)/1000

 
    Do While Not RsReaderInfo.EOF
        
            'If CheckNull(Trim(RsReaderInfo.Fields("VchNo").Value)) = "" Then GoTo Skip
            If (Trim(RsReaderInfo.Fields("PaperName").Value) <> PaperName) And PaperName <> "" Then
                
                Dim OFFFF As Double

                OFFFF = TOFF4
                
            
                MsgBox Int(TOFF4) + Int((TOFF4 - Int(TOFF4)) * 1000 / 500) + ((TOFF4 - Int(TOFF4)) * 1000 Mod 500) / 1000
                
                oExcel.Application.Cells(R, 8).Value = Format((TOFF3), "0.000")
                oExcel.Application.Cells(R, 9).Value = Format((TOFF4), "0.000")
                
                
                oExcel.Application.Cells(R, 10).Value = Format(Val(TOFF3 - TOFF4), "0.000")
                
                
                oExcel.Rows(R).Font.Bold = True
                
                If R > 5 Then
                 oExcel.Application.Cells(R + 5, 4).Value = Trim(RsReaderInfo.Fields("PaperName").Value)
                 oExcel.Rows(R + 5).Font.Bold = True
                End If
                
                If (Trim(RsReaderInfo.Fields("GSM").Value) <> GSM) And GSM <> "" Then
                    oExcel.Application.Cells(R + 1, 8).Value = Format((TOFF23), "0.000")
                    oExcel.Application.Cells(R + 1, 9).Value = Format((TOFF24), "0.000")
                    oExcel.Application.Cells(R + 1, 10).Value = Format(Val(TOFF23 - TOFF24), "0.000")
                    oExcel.Application.Cells(R + 1, 4).Value = "GSM Total"
                    oExcel.Rows(R + 1).Font.Bold = True
                    TOFF23 = 0
                    TOFF24 = 0
                End If
                
                If (Trim(RsReaderInfo.Fields("SizeName").Value) <> SizeName) And SizeName <> "" Then
                    oExcel.Application.Cells(R + 2, 8).Value = Format((TOFF14), "0.000")
                    oExcel.Application.Cells(R + 2, 9).Value = Format((TOFF15), "0.000")
                    oExcel.Application.Cells(R + 2, 10).Value = Format(Val(TOFF14 - TOFF15), "0.000")
                    oExcel.Application.Cells(R + 2, 4).Value = "Size Total"
                    oExcel.Rows(R + 2).Font.Bold = True
                    TOFF14 = 0
                    TOFF15 = 0
                End If
                
                If (Trim(RsReaderInfo.Fields("GodownName").Value) <> GodownName) And GodownName <> "" Then
                    oExcel.Application.Cells(R + 3, 8).Value = Format((TOFF17), "0.000")
                    oExcel.Application.Cells(R + 3, 9).Value = Format((TOFF18), "0.000")
                    oExcel.Application.Cells(R + 3, 10).Value = Format(Val(TOFF17 - TOFF18), "0.000")
                    oExcel.Application.Cells(R + 3, 4).Value = "Godown Total"
                    oExcel.Rows(R + 3).Font.Bold = True
                    TOFF17 = 0
                    TOFF18 = 0
                End If
                
                
                TOFF3 = 0
                TOFF4 = 0
                OFF7 = 0
                OFF8 = 0
                
                R = R + 6
            End If
        
'
'            If (Trim(RsReaderInfo.Fields("SizeName").Value) <> PaperName) And SizeName <> "" Then
'
'            End If
            
            
'            OFF7=Fix(Sum({@FF5},{ado.PaperName})/500)+Remainder(Sum({@FF5},{ado.PaperName}),500)/1000'In
'
'            OFF8=Fix(Sum({@FF6},{ado.PaperName})/500)+Remainder(Sum({@FF6},{ado.PaperName}),500)/1000'Out
'
'            OFF2=Fix((Sum({@FF5},{ado.PaperName})-Sum({@FF6},{ado.PaperName}))/500)+Remainder(Sum({@FF5},{ado.PaperName})-Sum({@FF6},{ado.PaperName}),500)/1000'Balance

            
            If (InStr(1, "PI_SI_MI", RsReaderInfo.Fields("VchType").Value) Or RsReaderInfo.Fields("VchType").Value = "OB") And Val(RsReaderInfo.Fields("Quantity").Value) > 0 Then
                OFF5 = Val(RsReaderInfo.Fields("Quantity").Value)
            Else
                OFF5 = 0
            End If

            If InStr(1, "SR_MO_PC_DN", RsReaderInfo.Fields("VchType").Value) Or (RsReaderInfo.Fields("VchType").Value = "OB" And Val(RsReaderInfo.Fields("Quantity").Value) < 0) Then
                OFF6 = Abs(Val(RsReaderInfo.Fields("Quantity").Value))
            Else
                OFF6 = 0
            End If

            OFF3 = Int(OFF5 / 500) + (OFF5 Mod 500) / 1000
            OFF4 = Int(OFF6 / 500) + (OFF6 Mod 500) / 1000
            
             
            oExcel.Application.Cells(R, 1).Value = Cnt
            oExcel.Application.Cells(R, 2).Value = RsReaderInfo.Fields("VchNo").Value
            oExcel.Application.Cells(R, 3).Value = Format(RsReaderInfo.Fields("VchDate").Value, "dd-MMM-yyyy")
            oExcel.Application.Cells(R, 4).Value = Trim(RsReaderInfo.Fields("Particulars").Value)
            'oExcel.Application.Cells(R, 3).Value = Trim(RsReaderInfo.Fields("PaperName").Value)
            If Val(RsReaderInfo.Fields("BookQuantity").Value) > 0 Then
               oExcel.Application.Cells(R, 5).Value = Trim(RsReaderInfo.Fields("BookQuantity").Value)
            End If
            If Val(RsReaderInfo.Fields("QtyRecd").Value) > 0 Then
               oExcel.Application.Cells(R, 6).Value = Format(Val(RsReaderInfo.Fields("QtyRecd").Value), "0.000")
            End If
            'oExcel.Application.Cells(R, 5).Value = Format(Val(RsReaderInfo.Fields("QtyRecd").Value), "0.000")
            If Val(RsReaderInfo.Fields("Forms").Value) > 0 Then
               oExcel.Application.Cells(R, 7).Value = Trim(RsReaderInfo.Fields("Forms").Value)
            End If
            If OFF3 > 0 Then
               oExcel.Application.Cells(R, 8).Value = Format((OFF3), "0.000")
            End If
            If OFF4 > 0 Then
               oExcel.Application.Cells(R, 9).Value = Format((OFF4), "0.000")
            End If
            oExcel.Application.Cells(R, 10).Value = ""
            oExcel.Application.Cells(R, 11).Value = Trim(RsReaderInfo.Fields("Remarks").Value)
    
    
    '        oExcel.Application.Cells(R, 6).Value = Format(Int(Val(RsReaderInfo.Fields("Quantity").Value)) * Val(rstPaperIssueRegister.Fields("Weight/Ream").Value), "0.000")
    '        If Val(rstPaperIssueRegister.Fields("Quantity").Value) - Int(Val(rstPaperIssueRegister.Fields("Quantity").Value)) > 0 Then oExcel.Application.Cells(R, 6).Value = Val(oExcel.Application.Cells(R, 6).Value) + ((Val(rstPaperIssueRegister.Fields("Quantity").Value) - Int(Val(rstPaperIssueRegister.Fields("Quantity").Value))) * 1000) * (Val(rstPaperIssueRegister.Fields("Weight/Ream").Value) / 500)
    '        oExcel.Application.Cells(R, 6).Value = Val(oExcel.Application.Cells(R, 6).Value) / 1000
     
        
        
        PaperName = Trim(RsReaderInfo.Fields("PaperName").Value)
        
        GSM = Trim(RsReaderInfo.Fields("GSM").Value)
        
        SizeName = Trim(RsReaderInfo.Fields("SizeName").Value)
        
        GodownName = Trim(RsReaderInfo.Fields("GodownName").Value)

'        TOFF3 = TOFF3 + OFF3
'        TOFF4 = TOFF4 + OFF4
        
        TOFF3 = TOFF3 + OFF5
        TOFF4 = TOFF4 + OFF6
        
        
        
'        TOFF20 = TOFF20 + OFF3
'        TOFF21 = TOFF20 + OFF4
        
        TOFF20 = TOFF20 + OFF5
        TOFF21 = TOFF20 + OFF6
        
        If (Trim(RsReaderInfo.Fields("GSM").Value) = GSM) And GSM <> "" Then
'         TOFF23 = TOFF23 + OFF3
'         TOFF24 = TOFF24 + OFF4
         
         TOFF23 = TOFF23 + OFF5
         TOFF24 = TOFF24 + OFF6
         
        End If
        
        If (Trim(RsReaderInfo.Fields("SizeName").Value) = SizeName) And SizeName <> "" Then
'         TOFF14 = TOFF14 + OFF3
'         TOFF15 = TOFF15 + OFF4
'
         TOFF14 = TOFF14 + OFF5
         TOFF15 = TOFF15 + OFF6
        End If
        
        If (Trim(RsReaderInfo.Fields("GodownName").Value) = GodownName) And GodownName <> "" Then
         TOFF17 = TOFF17 + OFF5
         TOFF18 = TOFF18 + OFF6
'
'         TOFF17 = TOFF17 + OFF3
'         TOFF18 = TOFF18 + OFF4
         
        End If

        Cnt = Cnt + 1: R = R + 1
Skip:
        RsReaderInfo.MoveNext
    Loop
    
   
    oExcel.Application.Cells(R + 0, 8).Value = Format((TOFF23), "0.000")
    oExcel.Application.Cells(R + 0, 9).Value = Format((TOFF24), "0.000")
    oExcel.Application.Cells(R + 0, 10).Value = Format(Val(TOFF23 - TOFF24), "0.000")
    oExcel.Application.Cells(R + 0, 4).Value = "GSM Total"
    
    oExcel.Rows(R + 0).Font.Bold = True

    oExcel.Application.Cells(R + 1, 8).Value = Format((TOFF14), "0.000")
    oExcel.Application.Cells(R + 1, 9).Value = Format((TOFF15), "0.000")
    oExcel.Application.Cells(R + 1, 10).Value = Format(Val(TOFF14 - TOFF15), "0.000")
    oExcel.Application.Cells(R + 1, 4).Value = "Size Total"
    
    oExcel.Rows(R + 1).Font.Bold = True

    oExcel.Application.Cells(R + 2, 8).Value = Format((TOFF17), "0.000")
    oExcel.Application.Cells(R + 2, 9).Value = Format((TOFF18), "0.000")
    oExcel.Application.Cells(R + 2, 10).Value = Format(Val(TOFF17 - TOFF18), "0.000")
    oExcel.Application.Cells(R + 2, 4).Value = "Godown Total"
    oExcel.Rows(R + 2).Font.Bold = True

    oExcel.Application.Cells(R + 3, 8).Value = Format((TOFF20), "0.000")
    oExcel.Application.Cells(R + 3, 9).Value = Format((TOFF21), "0.000")
    oExcel.Application.Cells(R + 3, 10).Value = Format(Val(TOFF20 - TOFF21), "0.000")
    oExcel.Application.Cells(R + 3, 4).Value = "Grand Total"
    oExcel.Rows(R + 3).Font.Bold = True

    
    oExcel.Sheets("Sheet1").Activate
    oExcel.Columns("A:M").EntireColumn.AutoFit
    oExcel.Workbooks.Item(1).Save
    Screen.MousePointer = vbNormal
    If OutputTo = "S" Then oExcel.Range("A1").Activate: oExcel.Application.Visible = True Else oExcel.Workbooks.Item(1).PrintOut
    Set oExcel = Nothing
    On Error GoTo 0
End Sub




