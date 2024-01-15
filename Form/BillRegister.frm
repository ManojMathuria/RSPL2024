VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Begin VB.Form FrmBillRegister 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bill Register"
   ClientHeight    =   6660
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6540
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "BillRegister.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6660
   ScaleWidth      =   6540
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6540
      _ExtentX        =   11536
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Print Preview"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Print"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Mail"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exit"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   7
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Paper Supplier"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Outsource Item Supplier"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Book Printer"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Title Printer"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Title Laminator"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Book Binder"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "All"
               EndProperty
            EndProperty
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
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BillRegister.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BillRegister.frx":0986
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BillRegister.frx":0A9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BillRegister.frx":0BAC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   6255
      Left            =   45
      TabIndex        =   7
      Top             =   345
      Width           =   6450
      _Version        =   65536
      _ExtentX        =   11377
      _ExtentY        =   11033
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
      Picture         =   "BillRegister.frx":0CBE
      Begin VB.OptionButton Option3 
         Caption         =   "Pending"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3765
         TabIndex        =   2
         Top             =   10
         Value           =   -1  'True
         Width           =   1020
      End
      Begin VB.OptionButton Option2 
         Caption         =   "All"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   5775
         TabIndex        =   4
         Top             =   10
         Width           =   630
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Paid"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   4905
         TabIndex        =   3
         Top             =   10
         Width           =   750
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   5940
         Left            =   0
         TabIndex        =   5
         Top             =   315
         Width           =   6450
         _ExtentX        =   11377
         _ExtentY        =   10478
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
         Caption         =   " From"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   -2147483634
         Picture         =   "BillRegister.frx":0CDA
         Picture         =   "BillRegister.frx":0CF6
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
         Height          =   330
         Left            =   1920
         TabIndex        =   9
         Top             =   0
         Width           =   645
         _Version        =   65536
         _ExtentX        =   1138
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
         Caption         =   " To"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   -2147483634
         Picture         =   "BillRegister.frx":0D12
         Picture         =   "BillRegister.frx":0D2E
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
         Calendar        =   "BillRegister.frx":0D4A
         Caption         =   "BillRegister.frx":0E62
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BillRegister.frx":0ECE
         Keys            =   "BillRegister.frx":0EEC
         Spin            =   "BillRegister.frx":0F4A
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
         Left            =   2550
         TabIndex        =   1
         Top             =   0
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calendar        =   "BillRegister.frx":0F72
         Caption         =   "BillRegister.frx":108A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "BillRegister.frx":10F6
         Keys            =   "BillRegister.frx":1114
         Spin            =   "BillRegister.frx":1172
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
Attribute VB_Name = "FrmBillRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cnBillRegister As New ADODB.Connection          'for Connection to busy
Dim rstCompanyMaster As New ADODB.Recordset
Dim rstBillRegister As New ADODB.Recordset
Dim rstAccountList As New ADODB.Recordset
Dim AccountType As String, ReportType As Byte
Dim OutputTo As String
Private Sub Form_Load()
 On Error GoTo ErrorHandler
    CenterForm Me
    BusySystemIndicator True
    rstCompanyMaster.Open "SELECT PrintName FROM CompanyMaster", CxnDatabase, adOpenKeyset, adLockReadOnly
    Option3.Value = True
    MhDateInput1.Text = Format(FinancialYearFrom, "dd-mm-yyyy")
    MhDateInput2.Text = IIf(Format(FinancialYearTo, "yyyymmdd") < Format(Date, "yyyymmdd"), Format(FinancialYearTo, "dd-mm-yyyy"), Format(Date, "dd-mm-yyyy"))
    Toolbar1_ButtonMenuClick Toolbar1.Buttons(5).ButtonMenus.Item(7)
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
    Call CloseConnection(cnBillRegister)
    Call CloseRecordset(rstCompanyMaster)
    Call CloseRecordset(rstAccountList)
    Call CloseRecordset(rstBillRegister)
End Sub
Private Sub MhDateInput1_Validate(Cancel As Boolean)
    If Not IsDate(GetDate(MhDateInput1.Text)) Then Cancel = True
End Sub
Private Sub MhDateInput2_Validate(Cancel As Boolean)
    If Not IsDate(GetDate(MhDateInput2.Text)) Then
        Cancel = True
    ElseIf Format(GetDate(MhDateInput2.Text), "yyyymmdd") < Format(GetDate(MhDateInput1.Text), "yyyymmdd") Then
        Cancel = True
    End If
End Sub
Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    If Shift = vbCtrlMask And (KeyCode = vbKeyA Or KeyCode = vbKeyD) Then
        For i = 1 To ListView1.ListItems.Count
            ListView1.ListItems(i).Checked = IIf(KeyCode = vbKeyA, True, False)
        Next i
    End If
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    If Button.Index = 3 Then CloseForm Me: Exit Sub
    If Button.Index = 1 Then OutputTo = "S" Else OutputTo = "P"
    PrintBillRegister
End Sub
Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    On Error Resume Next
    ReportType = ButtonMenu.Index
    AccountType = Choose(ButtonMenu.Index, "01", "01", "05", "06", "07", "08", "00")
    Me.Caption = "Bill Register [" & Choose(ButtonMenu.Index, "Paper Supplier", "Outsource Item Supplier", "Book Printer", "Title Printer", "Title Laminator", "Book Binder", "All Accounts") & "]"
    If rstAccountList.State = adStateOpen Then rstAccountList.Close
    ListView1.ListItems.Clear
    If AccountType <> "00" Then
        rstAccountList.Open "SELECT Name,Code FROM AccountMaster WHERE Type='" & AccountType & "' ORDER BY Name", CxnDatabase, adOpenKeyset, adLockReadOnly
    Else
        Dim AccountList As String, DatabaseName As String, i As Integer
        'Comma Separated List of Accounts in Saral
        rstAccountList.Open "SELECT Alias FROM AccountMaster WHERE Alias NOT IN ('','0')", CxnDatabase, adOpenKeyset, adLockReadOnly
        With rstAccountList
            Do While Not .EOF
                AccountList = AccountList + "'" + .Fields("Alias").Value + "',"
                .MoveNext
            Loop
            If AccountList <> "" Then AccountList = Mid(AccountList, 1, Len(AccountList) - 1) Else AccountList = "''"
            If .State = adStateOpen Then .Close
        End With
        'Connection to busy
        DatabaseName = Trim(ReadFromFile("Busy Database Name"))
        Do While True
            i = InStr(1, DatabaseName, ",")
            If i = 0 Then Exit Do Else DatabaseName = Mid(DatabaseName, i + 1): i = 0
        Loop
        If ServerName = "" Or DatabaseName = "" Then On Error GoTo 0: Exit Sub
        cnBillRegister.CursorLocation = adUseClient
        If cnBillRegister.State = adStateOpen Then cnBillRegister.Close
        cnBillRegister.Open "Provider=SQLOLEDB.1;Password=" & ServerPassword & ";Persist Security Info=True;User ID=sa;Initial Catalog=" & DatabaseName & ";Data Source=" & ServerName
        'List of Accounts in busy
        rstAccountList.Open "SELECT Name,Alias As Code FROM Master1 WHERE MasterType=2 AND Alias IN (" & AccountList & ") ORDER BY Name", cnBillRegister, adOpenKeyset, adLockReadOnly
    End If
    Call FillList(ListView1, "List of " & Choose(ButtonMenu.Index, "Paper Suppliers", "Outsource Item Suppliers", "Book Printers", "Title Printers", "Title Laminators", "Book Binders", "Accounts") & "...", rstAccountList)
End Sub
Private Sub PrintBillRegister()
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    rptBillRegister.Text12.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptBillRegister.Text9.SetText IIf(Option3.Value, "Pending ", IIf(Option1.Value, "Paid ", "")) & "Bill Register From [" + Format(GetDate(MhDateInput1.Text), "dd-MM-yyyy") + "] To [" + Format(GetDate(MhDateInput2.Text), "dd-MM-yyyy") & "]"
    rptBillRegister.Text13.SetText "Account Type : " & Choose(ReportType, "Paper Supplier", "Outsource Item Supplier", "Book Printer", "Title Printer", "Title Laminator", "Book Binder", "All")
    If rstBillRegister.State = adStateOpen Then rstBillRegister.Close
    If AccountType = "00" Then
        rstBillRegister.Open "SELECT Trim(M1.PrintName) As AccountName,IIF(OrderType='1','PB/','PT/')+'" & Right(Year(FinancialYearFrom), 2) + "-" + Right(Year(FinancialYearTo), 2) & "/'+Trim(P.Name) As VchNo,P.Date As VchDate,M2.PrintName As ItemName,P.BillNo,P.BillDate,Format(C.QuantityOther,'0.000') As Quantity,P.BillAmount,P.PaidAmount,M1.Alias FROM ((PaperPOParent P INNER JOIN PaperPOChild C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON M1.Code=P.Supplier) INNER JOIN PaperMaster M2 ON M2.Code=C.Paper WHERE P.Date>=#" & GetDate(MhDateInput1.Text) & "# AND P.Date<=#" & GetDate(MhDateInput2.Text) & "# AND " & IIf(Option3.Value, "P.BillNo=''", IIf(Option1.Value, "P.BillNo<>''", "1=1")) & " AND M1.Alias IN (" & SelectedItems(ListView1) & ") " & _
                                     "UNION " & _
                                     "SELECT Trim(M1.PrintName) As AccountName,'MI/" & Right(Year(FinancialYearFrom), 2) + "-" + Right(Year(FinancialYearTo), 2) & "/'+Trim(P.Name) As VchNo,P.Date As VchDate,M2.PrintName As ItemName,P.BillNo,P.BillDate,Format(C.Quantity,'0') As Quantity,P.BillAmount,P.PaidAmount,M1.Alias FROM ((OutsourceItemPOParent P INNER JOIN OutsourceItemPOChild C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON M1.Code=P.Supplier) INNER JOIN OutsourceItemMaster M2 ON M2.Code=C.OutsourceItem WHERE P.Date>=#" & GetDate(MhDateInput1.Text) & "# AND P.Date<=#" & GetDate(MhDateInput2.Text) & "# AND " & IIf(Option3.Value, "P.BillNo=''", IIf(Option1.Value, "P.BillNo<>''", "1=1")) & " AND M1.Alias IN (" & SelectedItems(ListView1) & ") " & _
                                     "UNION " & _
                                     "SELECT Trim(M1.PrintName) As AccountName,'BP/" & Right(Year(FinancialYearFrom), 2) & "-" & Right(Year(FinancialYearTo), 2) & "/'+Trim(P.Name) As VchNo,P.Date As VchDate,M2.PrintName As ItemName,C.BillNo,C.BillDate,Format(C.ActualQuantity,'0') As Quantity,C.BillAmount,C.PaidAmount,M1.Alias FROM ((BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON M1.Code=P.BookPrinter) INNER JOIN BookMaster M2 ON M2.Code=P.Book WHERE P.Type<>'O' AND LEFT(P.Code,1)<>'*' AND P.Date>=#" & GetDate(MhDateInput1.Text) & "# AND P.Date<=#" & GetDate(MhDateInput2.Text) & "# AND " & IIf(Option3.Value, "C.BillNo=''", IIf(Option1.Value, "C.BillNo<>''", 1)) & " AND M1.Alias IN (" & SelectedItems(ListView1) & ") " & _
                                     "UNION " & _
                                     "SELECT Trim(M1.PrintName) As AccountName,'TP/" & Right(Year(FinancialYearFrom), 2) & "-" & Right(Year(FinancialYearTo), 2) & "/'+Trim(P.Name) As VchNo,P.Date As VchDate,M2.PrintName As ItemName,C.BillNo,C.BillDate,Format(C.ActualQuantity,'0') As Quantity,C.BillAmount,C.PaidAmount,M1.Alias FROM ((BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON M1.Code=P.TitlePrinter) INNER JOIN BookMaster M2 ON M2.Code=P.Book WHERE P.Type<>'O' AND LEFT(P.Code,1)<>'*' AND P.Date>=#" & GetDate(MhDateInput1.Text) & "# AND P.Date<=#" & GetDate(MhDateInput2.Text) & "# AND " & IIf(Option3.Value, "C.BillNo=''", IIf(Option1.Value, "C.BillNo<>''", 1)) & " AND M1.Alias IN (" & SelectedItems(ListView1) & ") " & _
                                     "UNION " & _
                                     "SELECT Trim(M1.PrintName) As AccountName,'TL/" & Right(Year(FinancialYearFrom), 2) & "-" & Right(Year(FinancialYearTo), 2) & "/'+Trim(P.Name) As VchNo,P.Date As VchDate,M2.PrintName As ItemName,C.BillNo,C.BillDate,Format(C.ActualQuantity,'0') As Quantity,C.BillAmount,C.PaidAmount,M1.Alias FROM ((BookPOParent P INNER JOIN BookPOChild07 C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON M1.Code=P.Laminator) INNER JOIN BookMaster M2 ON M2.Code=P.Book WHERE P.Type<>'O' AND LEFT(P.Code,1)<>'*' AND P.Date>=#" & GetDate(MhDateInput1.Text) & "# AND P.Date<=#" & GetDate(MhDateInput2.Text) & "# AND " & IIf(Option3.Value, "C.BillNo=''", IIf(Option1.Value, "C.BillNo<>''", 1)) & " AND M1.Alias IN (" & SelectedItems(ListView1) & ") " & _
                                     "UNION " & _
                                     "SELECT Trim(M1.PrintName) As AccountName,'BB/" & Right(Year(FinancialYearFrom), 2) & "-" & Right(Year(FinancialYearTo), 2) & "/'+Trim(P.Name) As VchNo,P.Date As VchDate,M2.PrintName As ItemName,C.BillNo,C.BillDate,Format(C.ActualQuantity,'0') As Quantity,C.BillAmount,C.PaidAmount,M1.Alias FROM ((BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON M1.Code=P.Binder) INNER JOIN BookMaster M2 ON M2.Code=P.Book WHERE P.Type<>'O' AND LEFT(P.Code,1)<>'*' AND P.Date>=#" & GetDate(MhDateInput1.Text) & "# AND P.Date<=#" & GetDate(MhDateInput2.Text) & "# AND " & IIf(Option3.Value, "C.BillNo=''", IIf(Option1.Value, "C.BillNo<>''", "1=1")) & " AND M1.Alias IN (" & SelectedItems(ListView1) & ") " & _
                                     " ORDER BY AccountName,VchNo", CxnDatabase, adOpenKeyset, adLockOptimistic
        rstBillRegister.ActiveConnection = Nothing
        If rstBillRegister.RecordCount > 0 Then rstBillRegister.MoveFirst
        With rstBillRegister
            .MoveFirst
            Do While Not .EOF
                rstAccountList.MoveFirst
                rstAccountList.Find "[Code]='" & Trim(.Fields("Alias").Value) & "'"
                If Not rstAccountList.EOF Then .Fields("AccountName").Value = rstAccountList.Fields("Name").Value
                .MoveNext
            Loop
        End With
    ElseIf AccountType = "01" Then
        rstBillRegister.Open "SELECT Trim(M1.PrintName) As AccountName,Trim(P.Name) As VchNo,P.Date As VchDate,M2.PrintName As ItemName,P.BillNo,P.BillDate,Format(C.QuantityOther,'0.000') As Quantity,P.BillAmount,P.PaidAmount FROM ((PaperPOParent P INNER JOIN PaperPOChild C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON M1.Code=P.Supplier) INNER JOIN PaperMaster M2 ON M2.Code=C.Paper WHERE P.Date>=#" & GetDate(MhDateInput1.Text) & "# AND P.Date<=#" & GetDate(MhDateInput2.Text) & "# AND " & IIf(Option3.Value, "P.BillNo=''", IIf(Option1.Value, "P.BillNo<>''", "1=1")) & " AND M1.Code IN (" & SelectedItems(ListView1) & ") ORDER BY M1.PrintName,P.Name", CxnDatabase, adOpenKeyset, adLockOptimistic
    ElseIf AccountType = "02" Then
        rstBillRegister.Open "SELECT Trim(M1.PrintName) As AccountName,Trim(P.Name) As VchNo,P.Date As VchDate,M2.PrintName As ItemName,P.BillNo,P.BillDate,Format(C.Quantity,'0') As Quantity,P.BillAmount,P.PaidAmount FROM ((OutsourceItemPOParent P INNER JOIN OutsourceItemPOChild C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON M1.Code=P.Supplier) INNER JOIN OutsourceItemMaster M2 ON M2.Code=C.OutsourceItem WHERE P.Date>=#" & GetDate(MhDateInput1.Text) & "# AND P.Date<=#" & GetDate(MhDateInput2.Text) & "# AND " & IIf(Option3.Value, "P.BillNo=''", IIf(Option1.Value, "P.BillNo<>''", "1=1")) & " AND M1.Code IN (" & SelectedItems(ListView1) & ") ORDER BY M1.PrintName,P.Name", CxnDatabase, adOpenKeyset, adLockOptimistic
    ElseIf AccountType = "05" Then
        rstBillRegister.Open "SELECT Trim(M1.PrintName) As AccountName,Trim(P.Name) As VchNo,P.Date As VchDate,M2.PrintName As ItemName,C.BillNo,C.BillDate,Format(C.ActualQuantity,'0') As Quantity,C.BillAmount,C.PaidAmount FROM ((BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON M1.Code=P.BookPrinter) INNER JOIN BookMaster M2 ON M2.Code=P.Book WHERE P.Type<>'O' AND LEFT(P.Code,1)<>'*' AND P.Date>=#" & GetDate(MhDateInput1.Text) & "# AND P.Date<=#" & GetDate(MhDateInput2.Text) & "# AND " & IIf(Option3.Value, "C.BillNo=''", IIf(Option1.Value, "C.BillNo<>''", 1)) & " AND M1.Code IN (" & SelectedItems(ListView1) & ") ORDER BY M1.PrintName,P.Name", CxnDatabase, adOpenKeyset, adLockOptimistic
    ElseIf AccountType = "06" Then
        rstBillRegister.Open "SELECT Trim(M1.PrintName) As AccountName,Trim(P.Name) As VchNo,P.Date As VchDate,M2.PrintName As ItemName,C.BillNo,C.BillDate,Format(C.ActualQuantity,'0') As Quantity,C.BillAmount,C.PaidAmount FROM ((BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON M1.Code=P.TitlePrinter) INNER JOIN BookMaster M2 ON M2.Code=P.Book WHERE P.Type<>'O' AND LEFT(P.Code,1)<>'*' AND P.Date>=#" & GetDate(MhDateInput1.Text) & "# AND P.Date<=#" & GetDate(MhDateInput2.Text) & "# AND " & IIf(Option3.Value, "C.BillNo=''", IIf(Option1.Value, "C.BillNo<>''", 1)) & " AND M1.Code IN (" & SelectedItems(ListView1) & ") ORDER BY M1.PrintName,P.Name", CxnDatabase, adOpenKeyset, adLockOptimistic
    ElseIf AccountType = "07" Then
        rstBillRegister.Open "SELECT Trim(M1.PrintName) As AccountName,Trim(P.Name) As VchNo,P.Date As VchDate,M2.PrintName As ItemName,C.BillNo,C.BillDate,Format(C.ActualQuantity,'0') As Quantity,C.BillAmount,C.PaidAmount FROM ((BookPOParent P INNER JOIN BookPOChild07 C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON M1.Code=P.Laminator) INNER JOIN BookMaster M2 ON M2.Code=P.Book WHERE P.Type<>'O' AND LEFT(P.Code,1)<>'*' AND P.Date>=#" & GetDate(MhDateInput1.Text) & "# AND P.Date<=#" & GetDate(MhDateInput2.Text) & "# AND " & IIf(Option3.Value, "C.BillNo=''", IIf(Option1.Value, "C.BillNo<>''", 1)) & " AND M1.Code IN (" & SelectedItems(ListView1) & ") ORDER BY M1.PrintName,P.Name", CxnDatabase, adOpenKeyset, adLockOptimistic
    ElseIf AccountType = "08" Then
        rstBillRegister.Open "SELECT Trim(M1.PrintName) As AccountName,Trim(P.Name) As VchNo,P.Date As VchDate,M2.PrintName As ItemName,C.BillNo,C.BillDate,Format(C.ActualQuantity,'0') As Quantity,C.BillAmount,C.PaidAmount FROM ((BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON M1.Code=P.Binder) INNER JOIN BookMaster M2 ON M2.Code=P.Book WHERE P.Type<>'O' AND LEFT(P.Code,1)<>'*' AND P.Date>=#" & GetDate(MhDateInput1.Text) & "# AND P.Date<=#" & GetDate(MhDateInput2.Text) & "# AND " & IIf(Option3.Value, "C.BillNo=''", IIf(Option1.Value, "C.BillNo<>''", "1=1")) & " AND M1.Code IN (" & SelectedItems(ListView1) & ") ORDER BY M1.PrintName,P.Name", CxnDatabase, adOpenKeyset, adLockOptimistic
    End If
    Screen.MousePointer = vbNormal
    If rstBillRegister.RecordCount = 0 Then On Error GoTo 0: Exit Sub
    rstBillRegister.MoveFirst
    rptBillRegister.Database.SetDataSource rstBillRegister, 3, 1
    rptBillRegister.DiscardSavedData
    If OutputTo = "S" Then Set FrmReportViewer.Report = rptBillRegister: FrmReportViewer.Show vbModal Else rptBillRegister.PrintOut
    Set rptBillRegister = Nothing
    On Error GoTo 0
End Sub
