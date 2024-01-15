VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Begin VB.Form FrmDayBook 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Day Book"
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
   Icon            =   "DayBook.frx":0000
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
      TabIndex        =   3
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
            Picture         =   "DayBook.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DayBook.frx":0986
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "DayBook.frx":0A9A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   6255
      Left            =   45
      TabIndex        =   4
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
      Picture         =   "DayBook.frx":0BAE
      Begin MSComctlLib.ListView lvwAccount 
         Height          =   5940
         Left            =   0
         TabIndex        =   2
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
         TabIndex        =   5
         Top             =   0
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
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
         Caption         =   " From Date"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   -2147483634
         Picture         =   "DayBook.frx":0BCA
         Picture         =   "DayBook.frx":0BE6
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
         Height          =   330
         Left            =   2165
         TabIndex        =   6
         Top             =   0
         Width           =   885
         _Version        =   65536
         _ExtentX        =   1561
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
         Caption         =   " To Date"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   -2147483634
         Picture         =   "DayBook.frx":0C02
         Picture         =   "DayBook.frx":0C1E
      End
      Begin TDBDate6Ctl.TDBDate MhDateInput1 
         Height          =   330
         Left            =   1080
         TabIndex        =   0
         Top             =   0
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calendar        =   "DayBook.frx":0C3A
         Caption         =   "DayBook.frx":0D52
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "DayBook.frx":0DBE
         Keys            =   "DayBook.frx":0DDC
         Spin            =   "DayBook.frx":0E3A
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
         Left            =   3025
         TabIndex        =   1
         Top             =   0
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calendar        =   "DayBook.frx":0E62
         Caption         =   "DayBook.frx":0F7A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "DayBook.frx":0FE6
         Keys            =   "DayBook.frx":1004
         Spin            =   "DayBook.frx":1062
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
Attribute VB_Name = "FrmDayBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CxnDayBook As New ADODB.Connection          'for Connection to busy
Dim rstCompanyMaster As New ADODB.Recordset
Dim rstAccountList As New ADODB.Recordset
Dim rstDayBook As New ADODB.Recordset
Dim OutputTo As String
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    CenterForm Me
    BusySystemIndicator True
    Dim AccountList As String, DatabaseName As String, i As Integer
    'Comma Separated List of Accounts in Saral
    rstAccountList.Open "SELECT Alias FROM AccountMaster WHERE Alias NOT IN ('','0')", CxnDatabase, adOpenKeyset, adLockReadOnly
    rstAccountList.ActiveConnection = Nothing
    With rstAccountList
        .MoveFirst
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
    CxnDayBook.CursorLocation = adUseClient
    If CxnDayBook.State = adStateOpen Then CxnDayBook.Close
    CxnDayBook.Open "Provider=SQLOLEDB.1;Password=" & ServerPassword & ";Persist Security Info=True;User ID=sa;Initial Catalog=" & DatabaseName & ";Data Source=" & ServerName
    'List of Accounts in busy
    rstAccountList.Open "SELECT Name,Alias As Code FROM Master1 WHERE MasterType=2 AND Alias IN (" & AccountList & ") ORDER BY Name", CxnDayBook, adOpenKeyset, adLockReadOnly
    rstAccountList.ActiveConnection = Nothing
    Call FillList(lvwAccount, "", rstAccountList)
    If Format(Date, "yyyymmdd") > Format(FinancialYearTo, "yyyymmdd") Then
        MhDateInput1.Text = Format(FinancialYearTo, "dd-mm-yyyy"): MhDateInput2.Text = Format(FinancialYearTo, "dd-mm-yyyy")
    Else
        MhDateInput1.Text = Format(Date, "dd-mm-yyyy"): MhDateInput2.Text = Format(Date, "dd-mm-yyyy")
    End If
    BusySystemIndicator False
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    Unload Me
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
    Call CloseConnection(CxnDayBook)
    Call CloseRecordset(rstCompanyMaster)
    Call CloseRecordset(rstDayBook)
    Call CloseRecordset(rstAccountList)
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
Private Sub lvwAccount_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        For i = 1 To lvwAccount.ListItems.Count
            lvwAccount.ListItems(i).Checked = True
        Next i
    ElseIf KeyCode = vbKeyD And Shift = vbCtrlMask Then
        For i = 1 To lvwAccount.ListItems.Count
            lvwAccount.ListItems(i).Checked = False
        Next i
    End If
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    If Button.Index = 1 Then
        OutputTo = "S"
        PrintDayBook
    ElseIf Button.Index = 2 Then
        OutputTo = "P"
        PrintDayBook
    ElseIf Button.Index = 3 Then
        CloseForm Me
    End If
End Sub
Private Sub PrintDayBook()
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    If rstDayBook.State = adStateOpen Then rstDayBook.Close
    rstDayBook.Open "SELECT M.Name As SaralAccount,M.Name As BusyAccount,'BP/" & Right(Year(FinancialYearFrom), 2) & "-" & Right(Year(FinancialYearTo), 2) & "/'+Trim(P.Name) As OrderNo,C.BillNo,C.BillDate,C.PaidAmount,C.Adjustment,C.AdjustmentRemarks,C.ComputerName,M.Alias FROM (BookPOChild05 C INNER JOIN BookPOParent P ON P.Code=C.Code) INNER JOIN AccountMaster M ON P.BookPrinter=M.Code WHERE BillFeedDate>=#" & GetDate(MhDateInput1.Text) & " 00:00:00# AND BillFeedDate<=#" & GetDate(MhDateInput2.Text) & " 23:59:59# AND M.Alias IN (" & SelectedItems(lvwAccount) & ") " & _
                                 "UNION " & _
                                 "SELECT M.Name As SaralAccount,M.Name As BusyAccount,'BP/" & Right(Year(FinancialYearFrom) - 1, 2) & "-" & Right(Year(FinancialYearTo) - 1, 2) & "/'+Trim(P.Name) As OrderNo,C.BillNo,C.BillDate,C.PaidAmount,C.Adjustment,C.AdjustmentRemarks,C.ComputerName,M.Alias FROM (OBookPOChild05 C INNER JOIN OBookPOParent P ON P.Code=C.Code) INNER JOIN AccountMaster M ON P.BookPrinter=M.Code WHERE BillFeedDate>=#" & GetDate(MhDateInput1.Text) & " 00:00:00# AND BillFeedDate<=#" & GetDate(MhDateInput2.Text) & " 23:59:59# AND M.Alias IN (" & SelectedItems(lvwAccount) & ") " & _
                                 "UNION " & _
                                 "SELECT M.Name As SaralAccount,M.Name As BusyAccount,'TP/" & Right(Year(FinancialYearFrom), 2) & "-" & Right(Year(FinancialYearTo), 2) & "/'+Trim(P.Name) As OrderNo,C.BillNo,C.BillDate,C.PaidAmount,C.Adjustment,C.AdjustmentRemarks,C.ComputerName,M.Alias FROM (BookPOChild06 C INNER JOIN BookPOParent P ON P.Code=C.Code) INNER JOIN AccountMaster M ON P.TitlePrinter=M.Code WHERE BillFeedDate>=#" & GetDate(MhDateInput1.Text) & " 00:00:00# AND BillFeedDate<=#" & GetDate(MhDateInput2.Text) & " 23:59:59# AND M.Alias IN (" & SelectedItems(lvwAccount) & ") " & _
                                 "UNION " & _
                                 "SELECT M.Name As SaralAccount,M.Name As BusyAccount,'TP/" & Right(Year(FinancialYearFrom) - 1, 2) & "-" & Right(Year(FinancialYearTo) - 1, 2) & "/'+Trim(P.Name) As OrderNo,C.BillNo,C.BillDate,C.PaidAmount,C.Adjustment,C.AdjustmentRemarks,C.ComputerName,M.Alias FROM (OBookPOChild06 C INNER JOIN OBookPOParent P ON P.Code=C.Code) INNER JOIN AccountMaster M ON P.TitlePrinter=M.Code WHERE BillFeedDate>=#" & GetDate(MhDateInput1.Text) & " 00:00:00# AND BillFeedDate<=#" & GetDate(MhDateInput2.Text) & " 23:59:59# AND M.Alias IN (" & SelectedItems(lvwAccount) & ") " & _
                                 "UNION " & _
                                 "SELECT M.Name As SaralAccount,M.Name As BusyAccount,'TL/" & Right(Year(FinancialYearFrom), 2) & "-" & Right(Year(FinancialYearTo), 2) & "/'+Trim(P.Name) As OrderNo,C.BillNo,C.BillDate,C.PaidAmount,C.Adjustment,C.AdjustmentRemarks,C.ComputerName,M.Alias FROM (BookPOChild07 C INNER JOIN BookPOParent P ON P.Code=C.Code) INNER JOIN AccountMaster M ON P.Laminator=M.Code WHERE BillFeedDate>=#" & GetDate(MhDateInput1.Text) & " 00:00:00# AND BillFeedDate<=#" & GetDate(MhDateInput2.Text) & " 23:59:59# AND M.Alias IN (" & SelectedItems(lvwAccount) & ") " & _
                                 "UNION " & _
                                 "SELECT M.Name As SaralAccount,M.Name As BusyAccount,'TL/" & Right(Year(FinancialYearFrom) - 1, 2) & "-" & Right(Year(FinancialYearTo) - 1, 2) & "/'+Trim(P.Name) As OrderNo,C.BillNo,C.BillDate,C.PaidAmount,C.Adjustment,C.AdjustmentRemarks,C.ComputerName,M.Alias FROM (OBookPOChild07 C INNER JOIN OBookPOParent P ON P.Code=C.Code) INNER JOIN AccountMaster M ON P.Laminator=M.Code WHERE BillFeedDate>=#" & GetDate(MhDateInput1.Text) & " 00:00:00# AND BillFeedDate<=#" & GetDate(MhDateInput2.Text) & " 23:59:59# AND M.Alias IN (" & SelectedItems(lvwAccount) & ") " & _
                                 "UNION " & _
                                 "SELECT M.Name As SaralAccount,M.Name As BusyAccount,'BB/" & Right(Year(FinancialYearFrom), 2) + "-" + Right(Year(FinancialYearTo), 2) & "/'+Trim(P.Name) As OrderNo,C.BillNo,C.BillDate,C.PaidAmount,C.Adjustment,C.AdjustmentRemarks,C.ComputerName,M.Alias FROM (BookPOChild08 C INNER JOIN BookPOParent P ON P.Code=C.Code) INNER JOIN AccountMaster M ON P.Binder=M.Code WHERE BillFeedDate>=#" & GetDate(MhDateInput1.Text) & " 00:00:00# AND BillFeedDate<=#" & GetDate(MhDateInput2.Text) & " 23:59:59# AND M.Alias IN (" & SelectedItems(lvwAccount) & ") " & _
                                 "UNION " & _
                                 "SELECT M.Name As SaralAccount,M.Name As BusyAccount,'BB/" & Right(Year(FinancialYearFrom) - 1, 2) + "-" + Right(Year(FinancialYearTo) - 1, 2) & "/'+Trim(P.Name) As OrderNo,C.BillNo,C.BillDate,C.PaidAmount,C.Adjustment,C.AdjustmentRemarks,C.ComputerName,M.Alias FROM (OBookPOChild08 C INNER JOIN OBookPOParent P ON P.Code=C.Code) INNER JOIN AccountMaster M ON P.Binder=M.Code WHERE BillFeedDate>=#" & GetDate(MhDateInput1.Text) & " 00:00:00# AND BillFeedDate<=#" & GetDate(MhDateInput2.Text) & " 23:59:59# AND M.Alias IN (" & SelectedItems(lvwAccount) & ") " & _
                                 "UNION " & _
                                 "SELECT M.Name As SaralAccount,M.Name As BusyAccount,IIF(OrderType='1','PB/','PT/')+'" & Right(Year(FinancialYearFrom), 2) + "-" + Right(Year(FinancialYearTo), 2) & "/'+Trim(T.Name) As OrderNo,T.BillNo,T.BillDate,T.PaidAmount,T.Adjustment,T.AdjustmentRemarks,T.ComputerName,M.Alias FROM PaperPOParent T INNER JOIN AccountMaster M ON T.Supplier=M.Code WHERE BillFeedDate>=#" & GetDate(MhDateInput1.Text) & " 00:00:00# AND BillFeedDate<=#" & GetDate(MhDateInput2.Text) & " 23:59:59# AND M.Alias IN (" & SelectedItems(lvwAccount) & ") " & _
                                 "UNION " & _
                                 "SELECT M.Name As SaralAccount,M.Name As BusyAccount,'MI/" & Right(Year(FinancialYearFrom), 2) + "-" + Right(Year(FinancialYearTo), 2) & "/'+Trim(T.Name) As OrderNo,T.BillNo,T.BillDate,T.PaidAmount,T.Adjustment,T.AdjustmentRemarks,T.ComputerName,M.Alias FROM OutsourceItemPOParent T INNER JOIN AccountMaster M ON T.Supplier=M.Code WHERE BillFeedDate>=#" & GetDate(MhDateInput1.Text) & " 00:00:00# AND BillFeedDate<=#" & GetDate(MhDateInput2.Text) & " 23:59:59# AND M.Alias IN (" & SelectedItems(lvwAccount) & ")", CxnDatabase, adOpenKeyset, adLockOptimistic
    Screen.MousePointer = vbNormal
    If rstDayBook.RecordCount = 0 Then On Error GoTo 0: Exit Sub
    rstDayBook.ActiveConnection = Nothing
    If rstAccountList.RecordCount > 0 Then rstAccountList.MoveFirst
    With rstDayBook
        .MoveFirst
        Do While Not .EOF
            rstAccountList.MoveFirst
            rstAccountList.Find "[Code]='" & Trim(.Fields("Alias").Value) & "'"
            If Not rstAccountList.EOF Then .Fields("BusyAccount").Value = rstAccountList.Fields("Name").Value
            .MoveNext
        Loop
    End With
    rstDayBook.MoveFirst
    rstCompanyMaster.Open "Select PrintName From CompanyMaster", CxnDatabase, adOpenKeyset, adLockReadOnly
    rptDayBook.Text12.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptDayBook.Text7.SetText "From " + Format(GetDate(MhDateInput1.Text), "dd-mm-yyyy") + " To " + Format(GetDate(MhDateInput2.Text), "dd-mm-yyyy")
    rptDayBook.Database.SetDataSource rstDayBook, 3, 1
    rptDayBook.DiscardSavedData
    If OutputTo = "S" Then Set FrmReportViewer.Report = rptDayBook: FrmReportViewer.Show vbModal Else rptDayBook.PrintOut
    Set rptDayBook = Nothing
    On Error GoTo 0
End Sub
