VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Begin VB.Form FrmPendingPaymentRegister 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pending Payments"
   ClientHeight    =   6435
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   5325
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "PendingPaymentRegister.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6435
   ScaleWidth      =   5325
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5325
      _ExtentX        =   9393
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
            Picture         =   "PendingPaymentRegister.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PendingPaymentRegister.frx":0986
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PendingPaymentRegister.frx":0A9A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   6060
      Left            =   45
      TabIndex        =   4
      Top             =   345
      Width           =   5250
      _Version        =   65536
      _ExtentX        =   9260
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
      Picture         =   "PendingPaymentRegister.frx":0BAE
      Begin VB.CheckBox Check1 
         Caption         =   "Without Nil"
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
         TabIndex        =   7
         Top             =   53
         Width           =   1335
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   5740
         Left            =   0
         TabIndex        =   2
         Top             =   315
         Width           =   5250
         _ExtentX        =   9260
         _ExtentY        =   10134
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
         TabIndex        =   5
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
         Picture         =   "PendingPaymentRegister.frx":0BCA
         Picture         =   "PendingPaymentRegister.frx":0BE6
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
         Height          =   330
         Left            =   1920
         TabIndex        =   6
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
         Picture         =   "PendingPaymentRegister.frx":0C02
         Picture         =   "PendingPaymentRegister.frx":0C1E
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
         Calendar        =   "PendingPaymentRegister.frx":0C3A
         Caption         =   "PendingPaymentRegister.frx":0D52
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "PendingPaymentRegister.frx":0DBE
         Keys            =   "PendingPaymentRegister.frx":0DDC
         Spin            =   "PendingPaymentRegister.frx":0E3A
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
         Calendar        =   "PendingPaymentRegister.frx":0E62
         Caption         =   "PendingPaymentRegister.frx":0F7A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "PendingPaymentRegister.frx":0FE6
         Keys            =   "PendingPaymentRegister.frx":1004
         Spin            =   "PendingPaymentRegister.frx":1062
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
Attribute VB_Name = "FrmPendingPaymentRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CxnPendingPayment As New ADODB.Connection
Dim rstCompanyMaster As New ADODB.Recordset
Dim rstAccountGroupList As New ADODB.Recordset
Dim rstAccountList As New ADODB.Recordset
Dim rstPendingPayment As New ADODB.Recordset
Dim rstUnbilledPayment As New ADODB.Recordset
Dim OutputTo As String
Private Sub Form_Load()
    Dim DatabaseName As String, i As Integer
    On Error GoTo ErrorHandler
    CenterForm Me
    BusySystemIndicator True
    DatabaseName = Trim(ReadFromFile("Busy Database Name"))
      
    Do While True
        i = InStr(1, DatabaseName, ",")
        If i = 0 Then Exit Do Else DatabaseName = Mid(DatabaseName, i + 1): i = 0
    Loop
    
    If ServerName = "" Or DatabaseName = "" Then Exit Sub
    CxnPendingPayment.CursorLocation = adUseClient
    If CxnPendingPayment.State = adStateOpen Then CxnPendingPayment.Close
    CxnPendingPayment.Open "Provider=SQLOLEDB.1;Password=" & ServerPassword & ";Persist Security Info=True;User ID=sa;Initial Catalog=" & DatabaseName & ";Data Source=" & ServerName
    Call LoadPartyList
    Call FillList(ListView1, "List of Printers...", rstAccountList)
    MhDateInput1.Text = Format(FinancialYearFrom, "dd-mm-yyyy")
    If Format(FinancialYearTo, "yyyymmdd") < Format(Date, "yyyymmdd") Then
        MhDateInput2.Text = Format(FinancialYearTo, "dd-mm-yyyy")
    Else
        MhDateInput2.Text = Format(Date, "dd-mm-yyyy")
    End If
    Check1.Value = 1
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
    If UnloadMode = 0 Then
        CloseForm Me
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CloseConnection(CxnPendingPayment)
    Call CloseRecordset(rstCompanyMaster)
    Call CloseRecordset(rstPendingPayment)
    Call CloseRecordset(rstUnbilledPayment)
    Call CloseRecordset(rstAccountGroupList)
    Call CloseRecordset(rstAccountList)
End Sub
Private Sub MhDateInput1_Validate(Cancel As Boolean)
    If Not IsDate(GetDate(MhDateInput1.Text)) Then
        Cancel = True
    End If
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
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    If Button.Index = 1 Then
        OutputTo = "S"
        PrintPendingPayment
    ElseIf Button.Index = 2 Then
        OutputTo = "P"
        PrintPendingPayment
    ElseIf Button.Index = 3 Then
        CloseForm Me
    End If
End Sub
Private Sub PrintPendingPayment()
    Dim oExcel As Object
    Dim i As Long, Cnt As Long
    Dim UnbilledPayment As Double
    If Not FileExist(App.Path & "\Template\Pending Payment Register.xlsx") Then Exit Sub
    Screen.MousePointer = vbHourglass
    If rstPendingPayment.State = adStateOpen Then rstPendingPayment.Close
    rstPendingPayment.Open "SELECT Name As Account,Alias,(SELECT Name FROM Master1 WHERE Code=M.ParentGrp) As AccountGroup," & _
                                                  "(SELECT ISNULL(SUM(D1),0)  FROM Folio1 WHERE MasterCode=M.Code) As OpBal," & _
                                                  "(SELECT  ISNULL(SUM(Value1),0) FROM Tran2 WHERE MasterCode1=M.Code AND (VchType=16 OR VchType=17 OR VchType=18 OR VchType=19 OR VchType=2) AND Date>='" & GetDate(MhDateInput1.Text) & "' AND Date<='" & GetDate(MhDateInput2.Text) & "') As CurBal, " & _
                                                  "(SELECT  ISNULL(SUM(ABS(Value1)),0) FROM Tran2 WHERE MasterCode1=M.Code AND VchType=19 AND Date>='" & GetDate(MhDateInput1.Text) & "' AND Date<='" & GetDate(MhDateInput2.Text) & "') As Payment " & _
                                                  "FROM Master1 M WHERE MasterType=2 AND Code IN (" & SelectedItems(ListView1) & ") ORDER BY PrintName", CxnPendingPayment, adOpenKeyset, adLockReadOnly
    rstPendingPayment.ActiveConnection = Nothing
    If rstUnbilledPayment.State = adStateOpen Then rstUnbilledPayment.Close
    rstUnbilledPayment.Open "SELECT Alias,SUM(BillAmount) As Amount FROM (PaperPOParent P INNER JOIN PaperPOChild C ON P.Code=C.Code) INNER JOIN AccountMaster M ON M.Code=P.Supplier WHERE  P.Date>=#" & GetDate(MhDateInput1.Text) & "# AND P.Date<=#" & GetDate(MhDateInput2.Text) & "# AND P.BillNo='' GROUP BY Alias " & _
                                     "UNION ALL " & _
                                     "SELECT Alias,SUM(BillAmount) As Amount FROM (OutsourceItemPOParent P INNER JOIN OutsourceItemPOChild C ON P.Code=C.Code) INNER JOIN AccountMaster M ON M.Code=P.Supplier WHERE  P.Date>=#" & GetDate(MhDateInput1.Text) & "# AND P.Date<=#" & GetDate(MhDateInput2.Text) & "# AND P.BillNo='' GROUP BY Alias " & _
                                     "UNION ALL " & _
                                     "SELECT Alias,SUM(BillAmount) As Amount FROM (BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN AccountMaster M ON M.Code=P.BookPrinter WHERE  P.Type<>'O' AND LEFT(P.Code,1)<>'*' AND P.Date>=#" & GetDate(MhDateInput1.Text) & "# AND P.Date<=#" & GetDate(MhDateInput2.Text) & "# AND C.BillNo='' GROUP BY Alias " & _
                                     "UNION ALL " & _
                                     "SELECT Alias,SUM(BillAmount) As Amount FROM (BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code) INNER JOIN AccountMaster M ON M.Code=P.TitlePrinter WHERE  P.Type<>'O' AND LEFT(P.Code,1)<>'*' AND P.Date>=#" & GetDate(MhDateInput1.Text) & "# AND P.Date<=#" & GetDate(MhDateInput2.Text) & "# AND C.BillNo='' GROUP BY Alias " & _
                                     "UNION ALL " & _
                                     "SELECT Alias,SUM(BillAmount) As Amount FROM (BookPOParent P INNER JOIN BookPOChild07 C ON P.Code=C.Code) INNER JOIN AccountMaster M ON M.Code=P.Laminator WHERE  P.Type<>'O' AND LEFT(P.Code,1)<>'*' AND P.Date>=#" & GetDate(MhDateInput1.Text) & "# AND P.Date<=#" & GetDate(MhDateInput2.Text) & "# AND C.BillNo='' GROUP BY Alias " & _
                                     "UNION ALL " & _
                                     "SELECT Alias,SUM(BillAmount) As Amount FROM (BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code) INNER JOIN AccountMaster M ON M.Code=P.Binder WHERE  P.Type<>'O' AND LEFT(P.Code,1)<>'*' AND P.Date>=#" & GetDate(MhDateInput1.Text) & "# AND P.Date<=#" & GetDate(MhDateInput2.Text) & "# AND C.BillNo='' GROUP BY Alias " & _
                                     "ORDER BY Alias", CxnDatabase, adOpenKeyset, adLockReadOnly
    'Last year Unbilled Payments
    Dim CxnImporter As New ADODB.Connection
    Dim rstImporter As New ADODB.Recordset
    On Error GoTo ErrorHandler
    If rstCompanyMaster.State = adStateOpen Then rstCompanyMaster.Close
    rstCompanyMaster.Open "Select PrintName,CreatedFrom From CompanyMaster", CxnDatabase, adOpenKeyset, adLockReadOnly
    If rstCompanyMaster.Fields("CreatedFrom").Value <> "" Then
        CxnImporter.CursorLocation = adUseClient
        CxnImporter.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DatabasePath & "\Saral." & rstCompanyMaster.Fields("CreatedFrom").Value & ";Persist Security Info=False;Jet OLEDB:Database Password=RSPLILoveMyINDIA"
        rstImporter.Open "SELECT Alias,SUM(BillAmount) As Amount FROM (PaperPOParent P INNER JOIN PaperPOChild C ON P.Code=C.Code) INNER JOIN AccountMaster M ON M.Code=P.Supplier WHERE P.BillNo='' GROUP BY Alias " & _
                                     "UNION ALL " & _
                                     "SELECT Alias,SUM(BillAmount) As Amount FROM (OutsourceItemPOParent P INNER JOIN OutsourceItemPOChild C ON P.Code=C.Code) INNER JOIN AccountMaster M ON M.Code=P.Supplier WHERE P.BillNo='' GROUP BY Alias " & _
                                     "UNION ALL " & _
                                     "SELECT Alias,SUM(BillAmount) As Amount FROM (BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN AccountMaster M ON M.Code=P.BookPrinter WHERE  P.Type<>'O' AND LEFT(P.Code,1)<>'*' AND C.BillNo='' GROUP BY Alias " & _
                                     "UNION ALL " & _
                                     "SELECT Alias,SUM(BillAmount) As Amount FROM (BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code) INNER JOIN AccountMaster M ON M.Code=P.TitlePrinter WHERE  P.Type<>'O' AND LEFT(P.Code,1)<>'*' AND C.BillNo='' GROUP BY Alias " & _
                                     "UNION ALL " & _
                                     "SELECT Alias,SUM(BillAmount) As Amount FROM (BookPOParent P INNER JOIN BookPOChild07 C ON P.Code=C.Code) INNER JOIN AccountMaster M ON M.Code=P.Laminator WHERE  P.Type<>'O' AND LEFT(P.Code,1)<>'*' AND C.BillNo='' GROUP BY Alias " & _
                                     "UNION ALL " & _
                                     "SELECT Alias,SUM(BillAmount) As Amount FROM (BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code) INNER JOIN AccountMaster M ON M.Code=P.Binder WHERE  P.Type<>'O' AND LEFT(P.Code,1)<>'*' AND C.BillNo='' GROUP BY Alias " & _
                                     "ORDER BY Alias", CxnImporter, adOpenKeyset, adLockReadOnly
        rstImporter.ActiveConnection = Nothing
        rstCompanyMaster.ActiveConnection = Nothing
    End If
    On Error Resume Next
    DoEvents
    Set oExcel = CreateObject("Excel.Application")
    oExcel.Workbooks.Open (App.Path & "\Template\Pending Payment Register")
    oExcel.DisplayAlerts = False
    oExcel.Workbooks.Item(1).SaveAs (App.Path & "\Report\Pending Payment Register (" & CompCode & ")")
    oExcel.DisplayAlerts = True
    oExcel.Visible = False
    oExcel.Cells(1, 1).Value = Trim(rstCompanyMaster.Fields("PrintName").Value)
    oExcel.Cells(2, 1).Value = "Pending Payments As On [" & Format(MhDateInput2, "dd-MMM-yyyy") & "]"
    i = 5
    Cnt = 1
    If rstPendingPayment.RecordCount > 0 Then rstPendingPayment.MoveFirst
    Do While Not rstPendingPayment.EOF
        UnbilledPayment = 0
        If rstUnbilledPayment.RecordCount > 0 Then rstUnbilledPayment.MoveFirst
        Do While Not rstUnbilledPayment.EOF
            If CheckNull(rstPendingPayment.Fields("Alias").Value) <> "" And Trim(rstUnbilledPayment.Fields("Alias").Value) = Trim(rstPendingPayment.Fields("Alias").Value) Then UnbilledPayment = UnbilledPayment + Val(rstUnbilledPayment.Fields("Amount").Value)
            rstUnbilledPayment.MoveNext
        Loop
        If rstImporter.State = adStateOpen Then
            If rstImporter.RecordCount > 0 Then rstImporter.MoveFirst
            Do While Not rstImporter.EOF
                If CheckNull(rstPendingPayment.Fields("Alias").Value) <> "" And Trim(rstImporter.Fields("Alias").Value) = Trim(rstPendingPayment.Fields("Alias").Value) Then UnbilledPayment = UnbilledPayment + Val(rstImporter.Fields("Amount").Value)
                rstImporter.MoveNext
            Loop
        End If
        If Check1.Value And (Val(rstPendingPayment.Fields("OpBal").Value) + Val(rstPendingPayment.Fields("CurBal").Value) + UnbilledPayment <= 0) Then GoTo Continue
        oExcel.Cells(i, 1).Value = Cnt
        oExcel.Application.Cells(i, 2).Value = Trim(rstPendingPayment.Fields("Account").Value)
        oExcel.Application.Cells(i, 3).Value = Trim(rstPendingPayment.Fields("AccountGroup").Value)
        oExcel.Application.Cells(i, 4).Value = Trim(rstPendingPayment.Fields("Alias").Value)
        oExcel.Application.Cells(i, 5).Value = Val(rstPendingPayment.Fields("OpBal").Value) + Val(rstPendingPayment.Fields("CurBal").Value)
        oExcel.Application.Cells(i, 6).Value = UnbilledPayment
        oExcel.Application.Cells(i, 7).Value = Val(oExcel.Application.Cells(i, 5)) + Val(oExcel.Application.Cells(i, 6))
        oExcel.Application.Cells(i, 8).Value = Val(rstPendingPayment.Fields("Payment").Value)
        If Val(oExcel.Application.Cells(i, 6)) < 0 Then oExcel.Application.Cells(i, 6).Value = 0
        Cnt = Cnt + 1
        i = i + 1
Continue:
        rstPendingPayment.MoveNext
    Loop
    oExcel.Columns("A:G").EntireColumn.AutoFit
    oExcel.Workbooks.Item(1).Save
    Screen.MousePointer = vbNormal
    If OutputTo = "S" Then
        oExcel.Range("A1").Activate
        oExcel.Visible = True
    Else
        oExcel.Workbooks.Item(1).PrintOut
    End If
    Set oExcel = Nothing
    Call CloseRecordset(rstPendingPayment)
    Call CloseRecordset(rstImporter)
    Call CloseRecordset(rstCompanyMaster)
    Call CloseConnection(CxnImporter)
    On Error GoTo 0
    Exit Sub
ErrorHandler:
    Call CloseRecordset(rstPendingPayment)
    Call CloseRecordset(rstImporter)
    Call CloseRecordset(rstCompanyMaster)
    Call CloseConnection(CxnImporter)
    Screen.MousePointer = vbNormal
End Sub
Private Sub LoadPartyList()
    Dim ParentGroups As String, CurrentGroups As String
    If rstAccountGroupList.State = adStateOpen Then rstAccountGroupList.Close
    rstAccountGroupList.Open "SELECT Code FROM Master1 WHERE MasterType=1 AND Name LIKE '%Sundry Creditor%'", CxnPendingPayment, adOpenKeyset, adLockReadOnly
    rstAccountGroupList.ActiveConnection = Nothing
    If rstAccountGroupList.RecordCount = 0 Then Exit Sub
    Do While Not rstAccountGroupList.EOF
        ParentGroups = ParentGroups & IIf(ParentGroups = "", "", ",") & rstAccountGroupList.Fields("Code").Value: CurrentGroups = rstAccountGroupList.Fields("Code").Value
        Do While True
            If rstAccountList.State = adStateOpen Then rstAccountList.Close
            rstAccountList.Open "SELECT Code FROM Master1 WHERE MasterType=1 AND ParentGrp IN (" & CurrentGroups & ")", CxnPendingPayment, adOpenKeyset, adLockReadOnly
            If rstAccountList.RecordCount = 0 Then Exit Do
            CurrentGroups = ""
            With rstAccountList
                Do While Not .EOF
                    CurrentGroups = CurrentGroups & IIf(CurrentGroups = "", "", ",") & rstAccountList.Fields(0).Value
                    .MoveNext
                Loop
                CurrentGroups = IIf(CurrentGroups = "", "''", CurrentGroups)
            End With
            ParentGroups = ParentGroups & "," & CurrentGroups
        Loop
        rstAccountGroupList.MoveNext
    Loop
    If rstAccountList.State = adStateOpen Then rstAccountList.Close
    rstAccountList.Open "SELECT Name, Code FROM Master1 WHERE MasterType=2 AND ParentGrp IN (" & ParentGroups & ") ORDER BY Name", CxnPendingPayment, adOpenKeyset, adLockReadOnly
    rstAccountList.ActiveConnection = Nothing
End Sub
