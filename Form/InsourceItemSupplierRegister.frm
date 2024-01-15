VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Begin VB.Form FrmInsourceItemSupplierRegister 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Insource Item Purchase Order Status Register"
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
   Icon            =   "InsourceItemSupplierRegister.frx":0000
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
      TabIndex        =   8
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
            Picture         =   "InsourceItemSupplierRegister.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "InsourceItemSupplierRegister.frx":0986
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "InsourceItemSupplierRegister.frx":0A9A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   6065
      Left            =   45
      TabIndex        =   9
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
      Picture         =   "InsourceItemSupplierRegister.frx":0BAE
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
         TabIndex        =   2
         Top             =   53
         Value           =   1  'Checked
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
         Left            =   6210
         TabIndex        =   4
         Top             =   10
         Width           =   1320
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
         Width           =   1125
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
         TabIndex        =   10
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
         Picture         =   "InsourceItemSupplierRegister.frx":0BCA
         Picture         =   "InsourceItemSupplierRegister.frx":0BE6
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
         Height          =   330
         Left            =   1920
         TabIndex        =   11
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
         Picture         =   "InsourceItemSupplierRegister.frx":0C02
         Picture         =   "InsourceItemSupplierRegister.frx":0C1E
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
         Calendar        =   "InsourceItemSupplierRegister.frx":0C3A
         Caption         =   "InsourceItemSupplierRegister.frx":0D52
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "InsourceItemSupplierRegister.frx":0DBE
         Keys            =   "InsourceItemSupplierRegister.frx":0DDC
         Spin            =   "InsourceItemSupplierRegister.frx":0E3A
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
         Calendar        =   "InsourceItemSupplierRegister.frx":0E62
         Caption         =   "InsourceItemSupplierRegister.frx":0F7A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "InsourceItemSupplierRegister.frx":0FE6
         Keys            =   "InsourceItemSupplierRegister.frx":1004
         Spin            =   "InsourceItemSupplierRegister.frx":1062
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
         TabIndex        =   7
         Top             =   3180
         Width           =   7530
         _ExtentX        =   13282
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
   End
End
Attribute VB_Name = "FrmInsourceItemSupplierRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstCompanyMaster As New ADODB.Recordset
Dim rstInsourceItemSupplierRegister As New ADODB.Recordset
Dim rstBookList As New ADODB.Recordset
Dim rstBoardList As New ADODB.Recordset
Dim rstPrinterList As New ADODB.Recordset
Dim OutputTo As String
Public ItemType As String
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    
    CenterForm Me
    BusySystemIndicator True
    rstPrinterList.Open "Select Name, Code From AccountMaster Where Type = " & IIf(ItemType = "5", "'07'", "'05'") & " Order By Name", CxnDatabase, adOpenKeyset, adLockReadOnly
    Call FillList(ListView3, "List of Printers...", rstPrinterList)
    Me.Caption = IIf(ItemType = "3", "Insource Item [Fresh Book] Purchase Order Status Register", IIf(ItemType = "4", "Insource Item [Repair Book] Purchase Order Status Register", "Insource Item [Title] Purchase Order Status Register"))
    rstCompanyMaster.Open "Select PrintName From CompanyMaster", CxnDatabase, adOpenKeyset, adLockReadOnly
    rstBoardList.Open "Select Name, Code From GeneralMaster Where Type = '2' And " & IIf(ItemType = "3", "Code='000000'", "Code<>'000000'") & " Order By Name", CxnDatabase, adOpenKeyset, adLockReadOnly
    rstBoardList.ActiveConnection = Nothing
    Call FillList(ListView1, "List of Boards...", rstBoardList)
    Call BookSelection(True)
    Option1.Value = True
    MhDateInput1.Text = Format(FinancialYearFrom, "dd-mm-yyyy")
    If Format(FinancialYearTo, "yyyymmdd") < Format(Date, "yyyymmdd") Then
        MhDateInput2.Text = Format(FinancialYearTo, "dd-mm-yyyy")
    Else
        MhDateInput2.Text = Format(Date, "dd-mm-yyyy")
    End If
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
    Call CloseRecordset(rstCompanyMaster)
    Call CloseRecordset(rstBookList)
    Call CloseRecordset(rstBoardList)
    Call CloseRecordset(rstPrinterList)
    Call CloseRecordset(rstInsourceItemSupplierRegister)
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
    End If
End Sub
Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
     Call BookSelection(False)
End Sub
Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer

    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        For i = 1 To ListView1.ListItems.Count
            ListView1.ListItems(i).Checked = True
        Next i
        Call BookSelection(True)
    ElseIf KeyCode = vbKeyD And Shift = vbCtrlMask Then
        For i = 1 To ListView1.ListItems.Count
            ListView1.ListItems(i).Checked = False
        Next i
        Call BookSelection(False)
    End If
End Sub
Private Sub ListView2_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer

    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        For i = 1 To ListView2.ListItems.Count
            ListView2.ListItems(i).Checked = True
            ListView2.ListItems(i).Selected = True
        Next i
    ElseIf KeyCode = vbKeyD And Shift = vbCtrlMask Then
        For i = 1 To ListView2.ListItems.Count
            ListView2.ListItems(i).Checked = False
            ListView2.ListItems(i).Selected = False
        Next i
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
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    
    If Button.Index = 1 Then
        OutputTo = "S"
        PrintInsourceItemSupplierRegister
    ElseIf Button.Index = 2 Then
        OutputTo = "P"
        PrintInsourceItemSupplierRegister
    ElseIf Button.Index = 3 Then
        CloseForm Me
    End If
End Sub
Private Sub BookSelection(ByVal SelectAll As Boolean)
    If rstBookList.State = adStateOpen Then
        rstBookList.Close
    End If
    rstBookList.Open "Select Name, Code From BookMaster Where " & IIf(SelectAll, IIf(ItemType = "3", "Board='000000'", IIf(ItemType = "4", "Type='R' And Board<>'000000'", "Board<>'000000' AND Type='F'")), "Board In (" & SelectedItems(ListView1) & ")") & " Order by Name", CxnDatabase, adOpenKeyset, adLockReadOnly
    rstBookList.ActiveConnection = Nothing
    ListView2.ListItems.Clear
    Call FillList(ListView2, "List of Books...", rstBookList)
End Sub
Private Sub PrintInsourceItemSupplierRegister()
    Dim CRXParamDefs As CRAXDRT.ParameterFieldDefinitions
    Dim CRXParamDef As CRAXDRT.ParameterFieldDefinition
    Dim SelectedInsourceItems As String
    Dim SelectedBoards As String
    Dim SelectedPrinters As String
    
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    rptInsourceItemSupplierRegister.Text11.SetText "Insource Item [" & IIf(ItemType = "3", "Fresh Book", IIf(ItemType = "4", "Repair Book", "Title")) & "] Purchase Order Status Register (" & IIf(Option1.Value, "Detailed", "Summarised") & ")"
    rptInsourceItemSupplierRegister.Text12.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptInsourceItemSupplierRegister.Text13.SetText "From [" + Format(MhDateInput1.Text, "dd-mm-yyyy") + "] To [" + Format(MhDateInput2.Text, "dd-mm-yyyy") + "]"
    If rstInsourceItemSupplierRegister.State = adStateOpen Then
        rstInsourceItemSupplierRegister.Close
    End If
    SelectedBoards = SelectedItems(ListView1)
    SelectedInsourceItems = SelectedItems(ListView2)
    SelectedPrinters = SelectedItems(ListView3)
    If ItemType <> "5" Then
        rstInsourceItemSupplierRegister.Open "Select 'Printer Name : '+Trim(M1.PrintName) As PrinterName,'Item Name : '+Trim(M2.PrintName) As InsourceItemName,'' As GodownName,Trim(P.Name) As VchNo,'' As OrderNo,P.Date As VchDate,'PO' As VchType,ActualQuantity As FinalQuantity From AccountMaster M1,BookMaster M2,BookPOParent P,BookPOChild05 C Where M2.Code=P.Book And P.Code=C.Code And P.Type<>'O' And Left(P.Code,1)<>'*' And P.BookPrinter=M1.Code And M2.Code In (" & SelectedInsourceItems & ")  And M2.Board In (" & SelectedBoards & ") And M1.Code In (" & SelectedPrinters & ") And P.Date>=#" & GetDate(MhDateInput1.Text) & "# And P.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION " & _
                                                                       "Select 'Printer Name : '+Trim(M1.PrintName) As PrinterName,'Item Name : '+Trim(M2.PrintName) As InsourceItemName,(Select Trim(PrintName) From AccountMaster Where Code=C.Godown) As GodownName,Trim(R.Name) As VchNo,Trim(P.Name) As OrderNo,P.Date As VchDate,'PS' As VchType,Quantity As FinalQuantity " & _
                                                                       "From AccountMaster M1,BookMaster M2,MaterialIOParent P,MaterialIOChild C,BookPOParent R Where M2.Code=C.Item And C.Category='" & ItemType & "' AND P.Code=C.Code And C.Ref=R.Code And P.Source=M1.Code And M2.Code In (" & SelectedInsourceItems & ")  And M2.Board In (" & SelectedBoards & ") And M1.Code In (" & SelectedPrinters & ")  And P.Date>=#" & GetDate(MhDateInput1.Text) & "# And P.Date<=#" & GetDate(MhDateInput2.Text) & "# Order By PrinterName,InsourceItemName,VchNo,OrderNo", CxnDatabase, adOpenKeyset, adLockOptimistic
    Else
        rstInsourceItemSupplierRegister.Open "Select 'Printer Name : '+Trim(M1.PrintName) As PrinterName,'Item Name : '+Trim(M2.PrintName) As InsourceItemName,'' As GodownName,Trim(P.Name) As VchNo,'' As OrderNo,P.Date As VchDate,'PO' As VchType,ActualQuantity-(SELECT IIF(ISNULL(SUM(ActualQuantity)),0,SUM(ActualQuantity)) FROM BookPOChild08 WHERE Code=P.Code) As FinalQuantity From AccountMaster M1,BookMaster M2,BookPOParent P,BookPOChild07 C Where M2.Code=P.Book And P.Code=C.Code And P.Type<>'O' And Left(P.Code,1)<>'*' And P.Laminator=M1.Code And M2.Code In (" & SelectedInsourceItems & ")  And M2.Board In (" & SelectedBoards & ") And IIF(P.Type='F',QuantityToOffice > 0,'1') And M1.Code In (" & SelectedPrinters & ") And P.Date>=#" & GetDate(MhDateInput1.Text) & "# And P.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION " & _
                                                                       "Select 'Printer Name : '+Trim(M1.PrintName) As PrinterName,'Item Name : '+Trim(M2.PrintName) As InsourceItemName,(Select Trim(PrintName) From AccountMaster Where Code=C.Godown) As GodownName,Trim(R.Name) As VchNo,Trim(P.Name) As OrderNo,P.Date As VchDate,'PS' As VchType,Quantity As FinalQuantity " & _
                                                                       "From AccountMaster M1,BookMaster M2,MaterialIOParent P,MaterialIOChild C,BookPOParent R Where M2.Code=C.Item And C.Category='" & ItemType & "' AND P.Code=C.Code And C.Ref=R.Code And P.Source=M1.Code And M2.Code In (" & SelectedInsourceItems & ")  And M2.Board In (" & SelectedBoards & ") And M1.Code In (" & SelectedPrinters & ")  And P.Date>=#" & GetDate(MhDateInput1.Text) & "# And P.Date<=#" & GetDate(MhDateInput2.Text) & "# Order By PrinterName,InsourceItemName,VchNo,OrderNo", CxnDatabase, adOpenKeyset, adLockOptimistic
    End If
    Screen.MousePointer = vbNormal
    If rstInsourceItemSupplierRegister.RecordCount = 0 Then
        On Error GoTo 0
        Exit Sub
    End If
    rptInsourceItemSupplierRegister.Database.SetDataSource rstInsourceItemSupplierRegister, 3, 1
    rptInsourceItemSupplierRegister.DiscardSavedData
    Set CRXParamDefs = rptInsourceItemSupplierRegister.ParameterFields
    For Each CRXParamDef In CRXParamDefs
        If CRXParamDef.ParameterFieldName = "PF1" Then
            CRXParamDef.SetCurrentValue (IIf(Check1.Value, 0, 0.1))
        ElseIf CRXParamDef.ParameterFieldName = "PF2" Then
            CRXParamDef.SetCurrentValue (IIf(Option1.Value, "D", "S"))
        End If
    Next
    rptInsourceItemSupplierRegister.EnableParameterPrompting = False
    If OutputTo = "S" Then
        Set FrmReportViewer.Report = rptInsourceItemSupplierRegister
        FrmReportViewer.Show vbModal
    Else
        rptInsourceItemSupplierRegister.PrintOut
    End If
    Set rptInsourceItemSupplierRegister = Nothing
    On Error GoTo 0
End Sub
