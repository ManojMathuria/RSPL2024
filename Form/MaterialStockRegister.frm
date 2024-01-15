VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Begin VB.Form FrmMaterialStockRegister 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Material Stock Register"
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
   Icon            =   "MaterialStockRegister.frx":0000
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
            Picture         =   "MaterialStockRegister.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MaterialStockRegister.frx":0986
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MaterialStockRegister.frx":0A9A
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
      Picture         =   "MaterialStockRegister.frx":0BAE
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
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
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
         Picture         =   "MaterialStockRegister.frx":0BCA
         Picture         =   "MaterialStockRegister.frx":0BE6
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
         Picture         =   "MaterialStockRegister.frx":0C02
         Picture         =   "MaterialStockRegister.frx":0C1E
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
         Left            =   0
         TabIndex        =   7
         Top             =   3180
         Width           =   7530
         _ExtentX        =   13282
         _ExtentY        =   5080
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
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
         Calendar        =   "MaterialStockRegister.frx":0C3A
         Caption         =   "MaterialStockRegister.frx":0D52
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "MaterialStockRegister.frx":0DBE
         Keys            =   "MaterialStockRegister.frx":0DDC
         Spin            =   "MaterialStockRegister.frx":0E3A
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
         Left            =   2640
         TabIndex        =   1
         Top             =   0
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calendar        =   "MaterialStockRegister.frx":0E62
         Caption         =   "MaterialStockRegister.frx":0F7A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "MaterialStockRegister.frx":0FE6
         Keys            =   "MaterialStockRegister.frx":1004
         Spin            =   "MaterialStockRegister.frx":1062
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
Attribute VB_Name = "FrmMaterialStockRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstCompanyMaster As New ADODB.Recordset
Dim rstMaterialStockRegister As New ADODB.Recordset
Dim rstBoardList As New ADODB.Recordset
Dim rstBookList As New ADODB.Recordset
Dim rstAccountList As New ADODB.Recordset
Dim OutputTo As String
Public ReportType As String
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    
    CenterForm Me
    BusySystemIndicator True
    If ReportType = "1" Then
        Me.Caption = "Material Stock Register [Binderwise/Bookwise/Itemwise]"
    Else
        Me.Caption = "Material Stock Register [Binderwise/Itemwise]"
    End If
    rstCompanyMaster.Open "Select PrintName From CompanyMaster", CxnDatabase, adOpenKeyset, adLockReadOnly
    rstAccountList.Open "Select Name As Col0, Code From AccountMaster Where Type In ('08','09') Order By Name", CxnDatabase, adOpenKeyset, adLockReadOnly
    rstAccountList.ActiveConnection = Nothing
    Call FillList(ListView3, "List of Godowns...", rstAccountList)
    If ReportType = "1" Then
        rstBoardList.Open "Select Name,Code From GeneralMaster Where Type = '2' Order by Name", CxnDatabase, adOpenKeyset, adLockReadOnly
        rstBoardList.ActiveConnection = Nothing
        Call FillList(ListView1, IIf(ReportType = "1", "List of Boards...", "List of Item Types"), rstBoardList)
        Call BookSelection(True)
        ListView2.MultiSelect = False
    Else
        rstBoardList.Open "Select Name,Code From GeneralMaster Where Type = '0' Order by Name", CxnDatabase, adOpenKeyset, adLockOptimistic
        rstBoardList.ActiveConnection = Nothing
        rstBoardList.AddNew
        rstBoardList.Fields("Name").Value = "Outsource Item"
        rstBoardList.Fields("Code").Value = "000001"
        rstBoardList.Update
        rstBoardList.AddNew
        rstBoardList.Fields("Name").Value = "Fresh Book"
        rstBoardList.Fields("Code").Value = "000003"
        rstBoardList.Update
        rstBoardList.AddNew
        rstBoardList.Fields("Name").Value = "Repair Book"
        rstBoardList.Fields("Code").Value = "000004"
        rstBoardList.Update
        rstBoardList.AddNew
        rstBoardList.Fields("Name").Value = "Title"
        rstBoardList.Fields("Code").Value = "000005"
        rstBoardList.Update
        Call FillList(ListView1, IIf(ReportType = "1", "List of Boards...", "List of Item Types"), rstBoardList)
        ListView1.ListItems(1).Selected = True
        Call BookSelection(False)
        ListView1.MultiSelect = False
    End If
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
    Call CloseForm(Me)
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
        Call CloseForm(Me)
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CloseRecordset(rstCompanyMaster)
    Call CloseRecordset(rstBoardList)
    Call CloseRecordset(rstBookList)
    Call CloseRecordset(rstAccountList)
    Call CloseRecordset(rstMaterialStockRegister)
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
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Call BookSelection(False)
End Sub
Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer

    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        For i = 1 To ListView1.ListItems.Count
            ListView1.ListItems(i).Selected = True
        Next i
        If ReportType = "1" Then
            Call BookSelection(True)
        Else
            Call BookSelection(False)
        End If
    ElseIf KeyCode = vbKeyD And Shift = vbCtrlMask Then
        For i = 1 To ListView1.ListItems.Count
            ListView1.ListItems(i).Selected = False
        Next i
        If ReportType = "2" Then
            ListView1.ListItems(4).Selected = True
        End If
        Call BookSelection(False)
    End If
End Sub
Private Sub ListView2_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer

    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        For i = 1 To ListView2.ListItems.Count
            ListView2.ListItems(i).Selected = True
        Next i
    ElseIf KeyCode = vbKeyD And Shift = vbCtrlMask Then
        For i = 1 To ListView2.ListItems.Count
            ListView2.ListItems(i).Selected = False
        Next i
    End If
End Sub
Private Sub ListView3_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer

    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        For i = 1 To ListView3.ListItems.Count
            ListView3.ListItems(i).Selected = True
        Next i
    ElseIf KeyCode = vbKeyD And Shift = vbCtrlMask Then
        For i = 1 To ListView3.ListItems.Count
            ListView3.ListItems(i).Selected = False
        Next i
    End If
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    
    If Button.Index = 1 Then
        OutputTo = "S"
        PrintMaterialStockRegister
    ElseIf Button.Index = 2 Then
        OutputTo = "P"
        PrintMaterialStockRegister
    ElseIf Button.Index = 3 Then
        Call CloseForm(Me)
    End If
End Sub
Private Sub BookSelection(ByVal SelectAll As Boolean)
    If rstBookList.State = adStateOpen Then
        rstBookList.Close
    End If
    If ReportType = "1" Then
        rstBookList.Open "Select Name, Code From BookMaster " & IIf(SelectAll, "'1'", "Where Board In (" & SelectedItems(ListView1, False) & ")") & " Order By Name", CxnDatabase, adOpenKeyset, adLockReadOnly
    Else
        rstBookList.Open "Select Name, Code FROM " & IIf(Val(ListView1.SelectedItem.SubItems(1)) = 1, "OutsourceItemMaster", IIf(Val(ListView1.SelectedItem.SubItems(1)) = 3, "BookMaster WHERE Board='000000'", IIf(Val(ListView1.SelectedItem.SubItems(1)) = 4, "BookMaster WHERE Type='R'", "BookMaster WHERE Board<>'000000' AND Type='F'"))) & " ORDER BY Name", CxnDatabase, adOpenKeyset, adLockReadOnly
    End If
    rstBookList.ActiveConnection = Nothing
    ListView2.ListItems.Clear
    Call FillList(ListView2, "List of Books...", rstBookList)
End Sub
Private Sub PrintMaterialStockRegister()
    Dim CRXParamDefs As CRAXDRT.ParameterFieldDefinitions
    Dim CRXParamDef As CRAXDRT.ParameterFieldDefinition
    Dim OutsourceItemQuantity As String
    Dim FreshBookQuantity As String
    Dim RepairBookQuantity As String
    Dim TitleQuantity As String
    Dim SelectedBoards As String
    Dim SelectedBooks As String
    Dim SelectedAccounts As String
    Dim SQL As String
    
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    rptMaterialStockRegister.Text11.SetText "Material Stock Register (" & IIf(Option1.Value, "Detailed", "Summarised") & ")"
    rptMaterialStockRegister.Text12.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptMaterialStockRegister.Text13.SetText "From [" + Format(GetDate(MhDateInput1.Text), "dd-mm-yyyy") + "] To [" + Format(GetDate(MhDateInput2.Text), "dd-mm-yyyy") + "]"
    If rstMaterialStockRegister.State = adStateOpen Then
        rstMaterialStockRegister.Close
    End If
    If ReportType = "1" Then
        rptMaterialStockRegister.Text11.Width = IIf(Option1.Value, 10800, 15780)
        rptMaterialStockRegister.Text12.Width = IIf(Option1.Value, 10800, 15780)
        rptMaterialStockRegister.Text13.Width = IIf(Option1.Value, 10800, 15780)
        rptMaterialStockRegister.Text9.Left = IIf(Option1.Value, 8880, 13860)
        rptMaterialStockRegister.Field17.Left = IIf(Option1.Value, 9840, 14820)
        rptMaterialStockRegister.Line5.Right = IIf(Option1.Value, 10800, 15780)
        SelectedBoards = SelectedItems(ListView1, False)
        SelectedBooks = SelectedItems(ListView2, False)
        SelectedAccounts = SelectedItems(ListView3, False)
        OutsourceItemQuantity = "(SELECT IIF(ISNULL(SUM(OpBal)),0,SUM(OpBal)) FROM AccountChild0801 WHERE Category=C.Category AND Item=C.Item AND Code=A.Code)+" & _
                                                  "(SELECT IIF(ISNULL(SUM(I.Quantity)),0,SUM(I.Quantity)) FROM MaterialIOParent M,MaterialIOChild I WHERE M.Code=I.Code AND Category=C.Category AND Item=C.Item AND Godown=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)+" & _
                                                  "(SELECT IIF(ISNULL(SUM(I.Quantity)),0,SUM(I.Quantity)) FROM MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND Category=C.Category AND Item=C.Item AND Account=A.Code AND I.Quantity>=0 AND Date<#" & GetDate(MhDateInput1.Text) & "#)-" & _
                                                  "(SELECT IIF(ISNULL(SUM(I.Quantity)),0,SUM(I.Quantity)) FROM MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND Category=C.Category AND Item=C.Item AND Account=A.Code AND I.Quantity<0 AND Date<#" & GetDate(MhDateInput1.Text) & "#)-" & _
                                                  "(SELECT IIF(ISNULL(SUM(I.Quantity)),0,SUM(I.Quantity)) FROM MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND Category=C.Category AND Item=C.Item AND AccountFROM=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)+" & _
                                                  "(SELECT IIF(ISNULL(SUM(I.Quantity)),0,SUM(I.Quantity)) FROM MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND Category=C.Category AND Item=C.Item AND AccountTo=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)-" & _
                                                  "(SELECT IIF(ISNULL(SUM(I.Quantity)),0,SUM(I.Quantity*Quantity*(SELECT ActualQuantity FROM BookPOChild08 WHERE Code=M.Code))) FROM BookPOParent M,BookPOChild0801 I WHERE M.Code=I.Code AND M.Type<>'O' AND LEFT(M.Code,1)<>'*' AND Category=C.Category AND Item=C.Item AND Binder=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)"
        FreshBookQuantity = "(SELECT IIF(ISNULL(SUM(OpBal)),0,SUM(OpBal)) FROM AccountChild0801 WHERE Category=C.Category AND Item=C.Item AND Code=A.Code)+" & _
                                           "(SELECT IIF(ISNULL(SUM(I.Quantity)),0,SUM(I.Quantity)) FROM MaterialIOParent M,MaterialIOChild I WHERE M.Code=I.Code AND Category=C.Category AND Item=C.Item AND Godown=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)+" & _
                                           "(SELECT IIF(ISNULL(SUM(I.Quantity)),0,SUM(I.Quantity)) FROM MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND Category=C.Category AND Item=C.Item AND Account=A.Code AND I.Quantity>=0 AND Date<#" & GetDate(MhDateInput1.Text) & "#)-" & _
                                           "(SELECT IIF(ISNULL(SUM(I.Quantity)),0,SUM(I.Quantity)) FROM MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND Category=C.Category AND Item=C.Item AND Account=A.Code AND I.Quantity<0 AND Date<#" & GetDate(MhDateInput1.Text) & "#)-" & _
                                           "(SELECT IIF(ISNULL(SUM(I.Quantity)),0,SUM(I.Quantity)) FROM MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND Category=C.Category AND Item=C.Item AND AccountFROM=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)+" & _
                                           "(SELECT IIF(ISNULL(SUM(I.Quantity)),0,SUM(I.Quantity)) FROM MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND Category=C.Category AND Item=C.Item AND AccountTo=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)-" & _
                                           "(SELECT IIF(ISNULL(SUM(I.Quantity)),0,SUM(I.Quantity*Quantity*(SELECT ActualQuantity FROM BookPOChild08 WHERE Code=M.Code))) FROM BookPOParent M,BookPOChild0801 I WHERE M.Code=I.Code AND M.Type<>'O' AND LEFT(M.Code,1)<>'*' AND Category=C.Category AND Item=C.Item AND Binder=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)"
        RepairBookQuantity = "(SELECT IIF(ISNULL(SUM(OpBal)),0,SUM(OpBal)) FROM AccountChild0801 WHERE Category='4' AND Item=O.Code AND Code=A.Code)+" & _
                                            "(SELECT IIF(ISNULL(SUM(I.Quantity)),0,SUM(I.Quantity)) FROM MaterialIOParent M,MaterialIOChild I WHERE M.Code=I.Code AND Category='4' AND Item=O.Code AND Godown=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)+" & _
                                            "(SELECT IIF(ISNULL(SUM(I.Quantity)),0,SUM(I.Quantity)) FROM MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND Category='4' AND Item=O.Code AND Account=A.Code AND I.Quantity>=0 AND Date<#" & GetDate(MhDateInput1.Text) & "#)-" & _
                                            "(SELECT IIF(ISNULL(SUM(I.Quantity)),0,SUM(I.Quantity)) FROM MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND Category='4' AND Item=O.Code AND Account=A.Code AND I.Quantity<0 AND Date<#" & GetDate(MhDateInput1.Text) & "#)-" & _
                                            "(SELECT IIF(ISNULL(SUM(I.Quantity)),0,SUM(I.Quantity)) FROM MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND Category='4' AND Item=O.Code AND AccountFROM=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)+" & _
                                            "(SELECT IIF(ISNULL(SUM(I.Quantity)),0,SUM(I.Quantity)) FROM MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND Category='4' AND Item=O.Code AND AccountTo=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)-" & _
                                            "(SELECT IIF(ISNULL(SUM(I.Quantity)),0,SUM(I.Quantity*Quantity*(SELECT ActualQuantity FROM BookPOChild08 WHERE Code=M.Code))) FROM BookPOParent M,BookPOChild0801 I WHERE M.Code=I.Code AND M.Type<>'O' AND LEFT(M.Code,1)<>'*' AND Category='4' AND Item=O.Code AND Binder=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)"
        TitleQuantity = "(SELECT IIF(ISNULL(SUM(OpBal)),0,SUM(OpBal)) FROM AccountChild0801 WHERE Category='5' AND Item=O.Code AND Code=A.Code)+" & _
                                 "(SELECT IIF(ISNULL(SUM(I.Quantity)),0,SUM(I.Quantity)) FROM MaterialIOParent M,MaterialIOChild I WHERE M.Code=I.Code AND Category='5' AND Item=O.Code AND Godown=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)+" & _
                                 "(SELECT IIF(ISNULL(SUM(I.Quantity)),0,SUM(I.Quantity)) FROM MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND Category='5' AND Item=O.Code AND Account=A.Code AND I.Quantity>=0 AND Date<#" & GetDate(MhDateInput1.Text) & "#)-" & _
                                 "(SELECT IIF(ISNULL(SUM(I.Quantity)),0,SUM(I.Quantity)) FROM MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND Category='5' AND Item=O.Code AND Account=A.Code AND I.Quantity<0 AND Date<#" & GetDate(MhDateInput1.Text) & "#)-" & _
                                 "(SELECT IIF(ISNULL(SUM(I.Quantity)),0,SUM(I.Quantity)) FROM MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND Category='5' AND Item=O.Code AND AccountFROM=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)+" & _
                                 "(SELECT IIF(ISNULL(SUM(I.Quantity)),0,SUM(I.Quantity)) FROM MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND Category='5' AND Item=O.Code AND AccountTo=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)-" & _
                                 "(SELECT IIF(ISNULL(SUM(I.Quantity)),0,SUM(I.Quantity*Quantity*(SELECT ActualQuantity FROM BookPOChild08 WHERE Code=M.Code))) FROM BookPOParent M,BookPOChild0801 I WHERE M.Code=I.Code AND M.Type<>'O' AND LEFT(M.Code,1)<>'*' AND Category='5' AND Item=O.Code AND Binder=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)"
        SQL = SQL + "SELECT '' As VchNo,#" & CDate(GetDate(MhDateInput1.Text)) - 1 & "# As VchDate,'OB' As VchType,'Opening Balance' As Particulars," & OutsourceItemQuantity & " As Quantity,'Board Name : '+Trim(G.PrintName) As BoardName," & _
                            "'Book Name : '+Trim(B.PrintName) As BookName,'Item Name : '+Trim(O.PrintName)+' (Outsource Item)' As ItemName,'1' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster B,BookChild01 C,GeneralMaster G,OutsourceItemMaster O,AccountMaster A WHERE B.Code=C.Code AND B.Board=G.Code AND C.Item=O.Code AND C.Category='1' AND B.Code In (" & SelectedBooks & ") AND G.Code In (" & SelectedBoards & ") AND A.Code In (" & SelectedAccounts & ") AND (" & OutsourceItemQuantity & ") <> 0 UNION ALL " & _
                            "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'PI' As VchType,'Material In (From : '+(SELECT Trim(PrintName) From AccountMaster Where Code=M.Source)+')' As Particulars,I.Quantity,'Board Name : '+Trim(G.PrintName) As BoardName,'Book Name : '+Trim(B.PrintName) As BookName,'Item Name : '+Trim(O.PrintName)+' (Outsource Item)' As ItemName,'1' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster B,BookChild01 C,GeneralMaster G,OutsourceItemMaster O,AccountMaster A,MaterialIOParent M,MaterialIOChild I WHERE M.Code=I.Code AND (I.Category=C.Category AND I.Item=C.Item) AND I.Godown=A.Code AND B.Code=C.Code AND B.Board=G.Code AND (C.Item=O.Code AND C.Category='1') AND B.Code In (" & SelectedBooks & ") AND G.Code In (" & SelectedBoards & ") AND A.Code In (" & SelectedAccounts & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                            "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'SI' As VchType,'Stock Journal (Generated)' As Particulars,I.Quantity,'Board Name : '+Trim(G.PrintName) As BoardName,'Book Name : '+Trim(B.PrintName) As BookName," & _
                            "'Item Name : '+Trim(O.PrintName)+' (Outsource Item)' As ItemName,'1' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster B,BookChild01 C,GeneralMaster G,OutsourceItemMaster O,AccountMaster A,MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND (I.Category=C.Category AND I.Item=C.Item) AND M.Account=A.Code AND I.Quantity>=0 AND B.Code=C.Code AND B.Board=G.Code AND (C.Item=O.Code AND C.Category='1') AND B.Code In (" & SelectedBooks & ") AND G.Code In (" & SelectedBoards & ") AND A.Code In (" & SelectedAccounts & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                            "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'SR' As VchType,'Stock Journal (Consumed)' As Particulars,I.Quantity,'Board Name : '+Trim(G.PrintName) As BoardName,'Book Name : '+Trim(B.PrintName) As BookName," & _
                            "'Item Name : '+Trim(O.PrintName)+' (Outsource Item)' As ItemName,'1' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster B,BookChild01 C,GeneralMaster G,OutsourceItemMaster O,AccountMaster A,MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND (I.Category=C.Category AND I.Item=C.Item) AND M.Account=A.Code AND I.Quantity<0 AND B.Code=C.Code AND B.Board=G.Code AND (C.Item=O.Code AND C.Category='1') AND B.Code In (" & SelectedBooks & ") AND G.Code In (" & SelectedBoards & ") AND A.Code In (" & SelectedAccounts & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                            "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'MO' As VchType,'Material Out (To : '+(SELECT Trim(PrintName) From AccountMaster Where Code=M.AccountTo)+')' As Particulars,I.Quantity,'Board Name : '+Trim(G.PrintName) As BoardName,'Book Name : '+Trim(B.PrintName) As BookName,'Item Name : '+Trim(O.PrintName)+' (Outsource Item)' As ItemName,'1' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster B,BookChild01 C,GeneralMaster G,OutsourceItemMaster O,AccountMaster A,MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND (I.Category=C.Category AND I.Item=C.Item) AND M.AccountFrom=A.Code AND B.Code=C.Code AND B.Board=G.Code AND (C.Item=O.Code AND C.Category='1') AND B.Code In (" & SelectedBooks & ") AND G.Code In (" & SelectedBoards & ") AND A.Code In (" & SelectedAccounts & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                            "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'MI' As VchType,'Material In (From : '+(SELECT Trim(PrintName) From AccountMaster Where Code=M.AccountFrom)+')' As Particulars,I.Quantity,'Board Name : '+Trim(G.PrintName) As BoardName,'Book Name : '+Trim(B.PrintName) As BookName,'Item Name : '+Trim(O.PrintName)+' (Outsource Item)' As ItemName,'1' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster B,BookChild01 C,GeneralMaster G,OutsourceItemMaster O,AccountMaster A,MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND (I.Category=C.Category AND I.Item=C.Item) AND M.AccountTo=A.Code AND B.Code=C.Code AND B.Board=G.Code AND (C.Item=O.Code AND C.Category='1') AND B.Code In (" & SelectedBooks & ") AND G.Code In (" & SelectedBoards & ") AND A.Code In (" & SelectedAccounts & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                            "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'PC' As VchType,'Material Consumed' As Particulars,I.Quantity*(SELECT ActualQuantity FROM BookPOChild08 WHERE Code=M.Code),'Board Name : '+Trim(G.PrintName) As BoardName,'Book Name : '+Trim(B.PrintName) As BookName,'Item Name : '+Trim(O.PrintName)+' (Outsource Item)' As ItemName,'1' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster B,BookChild01 C,GeneralMaster G,OutsourceItemMaster O,AccountMaster A,BookPOParent M,BookPOChild0801 I WHERE M.Code=I.Code AND M.Type<>'O' AND LEFT(M.Code,1)<>'*' AND (I.Category=C.Category AND I.Item=C.Item) AND M.Binder=A.Code AND B.Code=C.Code AND B.Board=G.Code AND (C.Item=O.Code AND C.Category='1') AND B.Code In (" & SelectedBooks & ") AND G.Code In (" & SelectedBoards & ") AND A.Code In (" & SelectedAccounts & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL "
        SQL = SQL + "SELECT '' As VchNo,#" & CDate(GetDate(MhDateInput1.Text)) - 1 & "# As VchDate,'OB' As VchType,'Opening Balance' As Particulars," & FreshBookQuantity & " As Quantity,'Board Name : '+Trim(G.PrintName) As BoardName," & _
                            "'Book Name : '+Trim(B.PrintName) As BookName,'Item Name : '+Trim(O.PrintName)+' (Fresh Book)' As ItemName,'3' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster B,BookChild01 C,GeneralMaster G,BookMaster O,AccountMaster A WHERE B.Code=C.Code AND B.Board=G.Code AND C.Item=O.Code AND C.Category='3' AND B.Code In (" & SelectedBooks & ") AND G.Code In (" & SelectedBoards & ") AND A.Code In (" & SelectedAccounts & ") AND (" & FreshBookQuantity & ") <> 0 UNION ALL " & _
                            "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'PI' As VchType,'Material In (From : '+(SELECT Trim(PrintName) From AccountMaster Where Code=M.Source)+')' As Particulars,I.Quantity,'Board Name : '+Trim(G.PrintName) As BoardName,'Book Name : '+Trim(B.PrintName) As BookName,'Item Name : '+Trim(O.PrintName)+' (Fresh Book)' As ItemName,'3' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster B,BookChild01 C,GeneralMaster G,BookMaster O,AccountMaster A,MaterialIOParent M,MaterialIOChild I WHERE M.Code=I.Code AND (I.Category=C.Category AND I.Item=C.Item) AND I.Godown=A.Code AND B.Code=C.Code AND B.Board=G.Code AND (C.Item=O.Code AND C.Category='3') AND B.Code In (" & SelectedBooks & ") AND G.Code In (" & SelectedBoards & ") AND A.Code In (" & SelectedAccounts & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                            "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'SI' As VchType,'Stock Journal (Generated)' As Particulars,I.Quantity,'Board Name : '+Trim(G.PrintName) As BoardName,'Book Name : '+Trim(B.PrintName) As BookName," & _
                            "'Item Name : '+Trim(O.PrintName)+' (Fresh Book)' As ItemName,'3' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster B,BookChild01 C,GeneralMaster G,BookMaster O,AccountMaster A,MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND (I.Category=C.Category AND I.Item=C.Item) AND M.Account=A.Code AND I.Quantity>=0 AND B.Code=C.Code AND B.Board=G.Code AND (C.Item=O.Code AND C.Category='3') AND B.Code In (" & SelectedBooks & ") AND G.Code In (" & SelectedBoards & ") AND A.Code In (" & SelectedAccounts & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                            "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'SR' As VchType,'Stock Journal (Consumed)' As Particulars,I.Quantity,'Board Name : '+Trim(G.PrintName) As BoardName,'Book Name : '+Trim(B.PrintName) As BookName," & _
                            "'Item Name : '+Trim(O.PrintName)+' (Fresh Book)' As ItemName,'3' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster B,BookChild01 C,GeneralMaster G,BookMaster O,AccountMaster A,MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND (I.Category=C.Category AND I.Item=C.Item) AND M.Account=A.Code AND I.Quantity<0 AND B.Code=C.Code AND B.Board=G.Code AND (C.Item=O.Code AND C.Category='3') AND B.Code In (" & SelectedBooks & ") AND G.Code In (" & SelectedBoards & ") AND A.Code In (" & SelectedAccounts & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                            "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'MO' As VchType,'Material Out (To : '+(SELECT Trim(PrintName) From AccountMaster Where Code=M.AccountTo)+')' As Particulars,I.Quantity,'Board Name : '+Trim(G.PrintName) As BoardName,'Book Name : '+Trim(B.PrintName) As BookName,'Item Name : '+Trim(O.PrintName)+' (Fresh Book)' As ItemName,'3' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster B,BookChild01 C,GeneralMaster G,BookMaster O,AccountMaster A,MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND (I.Category=C.Category AND I.Item=C.Item) AND M.AccountFrom=A.Code AND B.Code=C.Code AND B.Board=G.Code AND (C.Item=O.Code AND C.Category='3') AND B.Code In (" & SelectedBooks & ") AND G.Code In (" & SelectedBoards & ") AND A.Code In (" & SelectedAccounts & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                            "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'MI' As VchType,'Material In (From : '+(SELECT Trim(PrintName) From AccountMaster Where Code=M.AccountFrom)+')' As Particulars,I.Quantity,'Board Name : '+Trim(G.PrintName) As BoardName,'Book Name : '+Trim(B.PrintName) As BookName,'Item Name : '+Trim(O.PrintName)+' (Fresh Book)' As ItemName,'3' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster B,BookChild01 C,GeneralMaster G,BookMaster O,AccountMaster A,MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND (I.Category=C.Category AND I.Item=C.Item) AND M.AccountTo=A.Code AND B.Code=C.Code AND B.Board=G.Code AND (C.Item=O.Code AND C.Category='3') AND B.Code In (" & SelectedBooks & ") AND G.Code In (" & SelectedBoards & ") AND A.Code In (" & SelectedAccounts & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                            "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'PC' As VchType,'Material Consumed' As Particulars,I.Quantity*(SELECT ActualQuantity FROM BookPOChild08 WHERE Code=M.Code),'Board Name : '+Trim(G.PrintName) As BoardName,'Book Name : '+Trim(B.PrintName) As BookName,'Item Name : '+Trim(O.PrintName)+' (Fresh Book)' As ItemName,'3' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster B,BookChild01 C,GeneralMaster G,BookMaster O,AccountMaster A,BookPOParent M,BookPOChild0801 I WHERE M.Code=I.Code AND M.Type<>'O' AND LEFT(M.Code,1)<>'*' AND (I.Category=C.Category AND I.Item=C.Item) AND M.Binder=A.Code AND B.Code=C.Code AND B.Board=G.Code AND (C.Item=O.Code AND C.Category='3') AND B.Code In (" & SelectedBooks & ") AND G.Code In (" & SelectedBoards & ") AND A.Code In (" & SelectedAccounts & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL "
        SQL = SQL + "SELECT '' As VchNo,#" & CDate(GetDate(MhDateInput1.Text)) - 1 & "# As VchDate,'OB' As VchType,'Opening Balance' As Particulars," & RepairBookQuantity & " As Quantity,'Board Name : '+Trim(G.PrintName) As BoardName," & _
                            "'Book Name : '+Trim(B.PrintName) As BookName,'Item Name : '+Trim(O.PrintName)+' (Repair Book)' As ItemName,'4' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster B,BookMaster O,GeneralMaster G,AccountMaster A WHERE O.Type='R' AND Left(B.BusyCode,6)=Left(O.BusyCode,6) AND B.Board=G.Code AND B.Code In (" & SelectedBooks & ") AND G.Code In (" & SelectedBoards & ") AND A.Code In (" & SelectedAccounts & ") AND (" & RepairBookQuantity & ") <> 0 UNION ALL " & _
                            "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'PI' As VchType,'Material In (From : '+(SELECT Trim(PrintName) From AccountMaster Where Code=M.Source)+')' As Particulars,I.Quantity,'Board Name : '+Trim(G.PrintName) As BoardName,'Book Name : '+Trim(B.PrintName) As BookName,'Item Name : '+Trim(O.PrintName)+' (Repair Book)' As ItemName,'4' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster B,BookMaster O,GeneralMaster G,AccountMaster A,MaterialIOParent M,MaterialIOChild I WHERE M.Code=I.Code AND (I.Category='4' AND I.Item=O.Code) AND I.Godown=A.Code AND Left(B.BusyCode,6)=Left(O.BusyCode,6) AND O.Type='R' AND B.Board=G.Code AND B.Code In (" & SelectedBooks & ") AND G.Code In (" & SelectedBoards & ") AND A.Code In (" & SelectedAccounts & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                            "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'SI' As VchType,'Stock Journal (Generated)' As Particulars,I.Quantity,'Board Name : '+Trim(G.PrintName) As BoardName,'Book Name : '+Trim(B.PrintName) As BookName," & _
                            "'Item Name : '+Trim(O.PrintName)+' (Repair Book)' As ItemName,'4' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster B,BookMaster O,GeneralMaster G,AccountMaster A,MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND (I.Category='4' AND I.Item=O.Code) AND M.Account=A.Code AND I.Quantity>=0 AND Left(B.BusyCode,6)=Left(O.BusyCode,6) AND O.Type='R' AND B.Board=G.Code AND B.Code In (" & SelectedBooks & ") AND G.Code In (" & SelectedBoards & ") AND A.Code In (" & SelectedAccounts & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                            "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'SR' As VchType,'Stock Journal (Consumed)' As Particulars,I.Quantity,'Board Name : '+Trim(G.PrintName) As BoardName,'Book Name : '+Trim(B.PrintName) As BookName," & _
                            "'Item Name : '+Trim(O.PrintName)+' (Repair Book)' As ItemName,'4' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster B,BookMaster O,GeneralMaster G,AccountMaster A,MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND (I.Category='4' AND I.Item=O.Code) AND M.Account=A.Code AND I.Quantity<0 AND Left(B.BusyCode,6)=Left(O.BusyCode,6) AND O.Type='R' AND B.Board=G.Code AND B.Code In (" & SelectedBooks & ") AND G.Code In (" & SelectedBoards & ") AND A.Code In (" & SelectedAccounts & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                            "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'MO' As VchType,'Material Out (To : '+(SELECT Trim(PrintName) From AccountMaster Where Code=M.AccountTo)+')' As Particulars,I.Quantity,'Board Name : '+Trim(G.PrintName) As BoardName,'Book Name : '+Trim(B.PrintName) As BookName,'Item Name : '+Trim(O.PrintName)+' (Repair Book)' As ItemName,'4' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster B,BookMaster O,GeneralMaster G,AccountMaster A,MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND (I.Category='4' AND I.Item=O.Code) AND M.AccountFrom=A.Code AND Left(B.BusyCode,6)=Left(O.BusyCode,6) AND O.Type='R' AND B.Board=G.Code AND B.Code In (" & SelectedBooks & ") AND G.Code In (" & SelectedBoards & ") AND A.Code In (" & SelectedAccounts & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                            "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'MI' As VchType,'Material In (From : '+(SELECT Trim(PrintName) From AccountMaster Where Code=M.AccountFrom)+')' As Particulars,I.Quantity,'Board Name : '+Trim(G.PrintName) As BoardName,'Book Name : '+Trim(B.PrintName) As BookName,'Item Name : '+Trim(O.PrintName)+' (Repair Book)' As ItemName,'4' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster B,BookMaster O,GeneralMaster G,AccountMaster A,MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND (I.Category='4' AND I.Item=O.Code) AND M.AccountTo=A.Code AND Left(B.BusyCode,6)=Left(O.BusyCode,6) AND O.Type='R' AND B.Board=G.Code AND B.Code In (" & SelectedBooks & ") AND G.Code In (" & SelectedBoards & ") AND A.Code In (" & SelectedAccounts & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                            "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'PC' As VchType,'Material Consumed' As Particulars,I.Quantity*(SELECT ActualQuantity FROM BookPOChild08 WHERE Code=M.Code),'Board Name : '+Trim(G.PrintName) As BoardName,'Book Name : '+Trim(B.PrintName) As BookName,'Item Name : '+Trim(O.PrintName)+' (Repair Book)' As ItemName,'4' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster B,BookMaster O,GeneralMaster G,AccountMaster A,BookPOParent M,BookPOChild0801 I WHERE M.Code=I.Code AND M.Type<>'O' AND LEFT(M.Code,1)<>'*' AND (I.Category='4' AND I.Item=O.Code) AND M.Binder=A.Code AND Left(B.BusyCode,6)=Left(O.BusyCode,6) AND O.Type='R' AND B.Board=G.Code AND B.Code In (" & SelectedBooks & ") AND G.Code In (" & SelectedBoards & ") AND A.Code In (" & SelectedAccounts & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL "
        SQL = SQL + "SELECT '' As VchNo,#" & CDate(GetDate(MhDateInput1.Text)) - 1 & "# As VchDate,'OB' As VchType,'Opening Balance' As Particulars," & TitleQuantity & " As Quantity,'Board Name : '+Trim(G.PrintName) As BoardName," & _
                            "'Book Name : '+Trim(B.PrintName) As BookName,'Item Name : '+Trim(O.PrintName)+' (Title)' As ItemName,'5' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster B,BookMaster O,GeneralMaster G,AccountMaster A WHERE B.Code=O.Code AND B.Board=G.Code AND B.Code In (" & SelectedBooks & ") AND G.Code In (" & SelectedBoards & ") AND A.Code In (" & SelectedAccounts & ") AND (" & TitleQuantity & ") <> 0 UNION ALL " & _
                            "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'PI' As VchType,'Material In (From : '+(SELECT Trim(PrintName) From AccountMaster Where Code=M.Source)+')' As Particulars,I.Quantity,'Board Name : '+Trim(G.PrintName) As BoardName,'Book Name : '+Trim(B.PrintName) As BookName,'Item Name : '+Trim(O.PrintName)+' (Title)' As ItemName,'5' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster B,BookMaster O,GeneralMaster G,AccountMaster A,MaterialIOParent M,MaterialIOChild I WHERE M.Code=I.Code AND (I.Category='5' AND I.Item=O.Code) AND I.Godown=A.Code AND B.Code=O.Code AND B.Board=G.Code AND B.Code In (" & SelectedBooks & ") AND G.Code In (" & SelectedBoards & ") AND A.Code In (" & SelectedAccounts & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                            "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'SI' As VchType,'Stock Journal (Generated)' As Particulars,I.Quantity,'Board Name : '+Trim(G.PrintName) As BoardName,'Book Name : '+Trim(B.PrintName) As BookName," & _
                            "'Item Name : '+Trim(O.PrintName)+' (Title)' As ItemName,'5' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster B,BookMaster O,GeneralMaster G,AccountMaster A,MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND (I.Category='5' AND I.Item=O.Code) AND M.Account=A.Code AND I.Quantity>=0 AND B.Code=O.Code AND B.Board=G.Code AND B.Code In (" & SelectedBooks & ") AND G.Code In (" & SelectedBoards & ") AND A.Code In (" & SelectedAccounts & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                            "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'SR' As VchType,'Stock Journal (Consumed)' As Particulars,I.Quantity,'Board Name : '+Trim(G.PrintName) As BoardName,'Book Name : '+Trim(B.PrintName) As BookName," & _
                            "'Item Name : '+Trim(O.PrintName)+' (Title)' As ItemName,'5' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster B,BookMaster O,GeneralMaster G,AccountMaster A,MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND (I.Category='5' AND I.Item=O.Code) AND M.Account=A.Code AND I.Quantity<0 AND B.Code=O.Code AND B.Board=G.Code AND B.Code In (" & SelectedBooks & ") AND G.Code In (" & SelectedBoards & ") AND A.Code In (" & SelectedAccounts & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                            "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'MO' As VchType,'Material Out (To : '+(SELECT Trim(PrintName) From AccountMaster Where Code=M.AccountTo)+')' As Particulars,I.Quantity,'Board Name : '+Trim(G.PrintName) As BoardName,'Book Name : '+Trim(B.PrintName) As BookName,'Item Name : '+Trim(O.PrintName)+' (Title)' As ItemName,'5' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster B,BookMaster O,GeneralMaster G,AccountMaster A,MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND (I.Category='5' AND I.Item=O.Code) AND M.AccountFrom=A.Code AND B.Code=O.Code AND B.Board=G.Code AND B.Code In (" & SelectedBooks & ") AND G.Code In (" & SelectedBoards & ") AND A.Code In (" & SelectedAccounts & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                            "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'MI' As VchType,'Material In (From : '+(SELECT Trim(PrintName) From AccountMaster Where Code=M.AccountFrom)+')' As Particulars,I.Quantity,'Board Name : '+Trim(G.PrintName) As BoardName,'Book Name : '+Trim(B.PrintName) As BookName,'Item Name : '+Trim(O.PrintName)+' (Title)' As ItemName,'5' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster B,BookMaster O,GeneralMaster G,AccountMaster A,MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND (I.Category='5' AND I.Item=O.Code) AND M.AccountTo=A.Code AND B.Code=O.Code AND B.Board=G.Code AND B.Code In (" & SelectedBooks & ") AND G.Code In (" & SelectedBoards & ") AND A.Code In (" & SelectedAccounts & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                            "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'PC' As VchType,'Material Consumed' As Particulars,I.Quantity*(SELECT ActualQuantity FROM BookPOChild08 WHERE Code=M.Code),'Board Name : '+Trim(G.PrintName) As BoardName,'Book Name : '+Trim(B.PrintName) As BookName,'Item Name : '+Trim(O.PrintName)+' (Title)' As ItemName,'5' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster B,BookMaster O,GeneralMaster G,AccountMaster A,BookPOParent M,BookPOChild0801 I WHERE M.Code=I.Code AND M.Type<>'O' AND LEFT(M.Code,1)<>'*' AND (I.Category='5' AND I.Item=O.Code) AND M.Binder=A.Code AND B.Code=O.Code AND B.Board=G.Code AND B.Code In (" & SelectedBooks & ") AND G.Code In (" & SelectedBoards & ") AND A.Code In (" & SelectedAccounts & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# "
        rstMaterialStockRegister.Open SQL & "ORDER BY GodownName,BoardName,BookName,ItemType,ItemName,VchDate,VchNo", CxnDatabase, adOpenKeyset, adLockReadOnly
    Else
        SelectedBooks = SelectedItems(ListView2, False)
        SelectedAccounts = SelectedItems(ListView3, False)
        If Val(ListView1.SelectedItem.SubItems(1)) = 1 Then
            OutsourceItemQuantity = "(SELECT IIF(ISNULL(SUM(OpBal)),0,SUM(OpBal)) FROM AccountChild0801 WHERE Category='1' AND Item=O.Code AND Code=A.Code)+" & _
                                                      "(SELECT IIF(ISNULL(SUM(I.Quantity)),0,SUM(I.Quantity)) FROM MaterialIOParent M,MaterialIOChild I WHERE M.Code=I.Code AND Category='1' AND Item=O.Code AND Godown=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)+" & _
                                                      "(SELECT IIF(ISNULL(SUM(I.Quantity)),0,SUM(I.Quantity)) FROM MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND Category='1' AND Item=O.Code AND Account=A.Code AND I.Quantity>=0 AND Date<#" & GetDate(MhDateInput1.Text) & "#)-" & _
                                                      "(SELECT IIF(ISNULL(SUM(I.Quantity)),0,SUM(I.Quantity)) FROM MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND Category='1' AND Item=O.Code AND Account=A.Code AND I.Quantity<0 AND Date<#" & GetDate(MhDateInput1.Text) & "#)-" & _
                                                      "(SELECT IIF(ISNULL(SUM(I.Quantity)),0,SUM(I.Quantity)) FROM MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND Category='1' AND Item=O.Code AND AccountFROM=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)+" & _
                                                      "(SELECT IIF(ISNULL(SUM(I.Quantity)),0,SUM(I.Quantity)) FROM MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND Category='1' AND Item=O.Code AND AccountTo=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)-" & _
                                                      "(SELECT IIF(ISNULL(SUM(I.Quantity)),0,SUM(I.Quantity*Quantity*(SELECT ActualQuantity FROM BookPOChild08 WHERE Code=M.Code))) FROM BookPOParent M,BookPOChild0801 I WHERE M.Code=I.Code AND M.Type<>'O' AND LEFT(M.Code,1)<>'*' AND Category='1' AND Item=O.Code AND Binder=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)"
            SQL = "SELECT '' As VchNo,#" & CDate(GetDate(MhDateInput1.Text)) - 1 & "# As VchDate,'OB' As VchType,'Opening Balance' As Particulars," & OutsourceItemQuantity & " As Quantity,'' As BoardName,'' As BookName,'Item Name : '+Trim(O.PrintName)+' (Outsource Item)' As ItemName,'1' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM OutsourceItemMaster O,AccountMaster A WHERE A.Code In (" & SelectedAccounts & ") AND O.Code In (" & SelectedBooks & ") AND (" & OutsourceItemQuantity & ") <> 0 UNION ALL " & _
                      "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'PI' As VchType,'Material In (From : '+(SELECT Trim(PrintName) From AccountMaster Where Code=M.Source)+')' As Particulars,I.Quantity,'' As BoardName,'' As BookName,'Item Name : '+Trim(O.PrintName)+' (Outsource Item)' As ItemName,'1' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM OutsourceItemMaster O,AccountMaster A,MaterialIOParent M,MaterialIOChild I WHERE M.Code=I.Code AND (I.Category='1' AND I.Item=O.Code) AND I.Godown=A.Code AND A.Code In (" & SelectedAccounts & ") AND O.Code In (" & SelectedBooks & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                      "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'SI' As VchType,'Stock Journal (Generated)' As Particulars,I.Quantity,'' As BoardName,'' As BookName,'Item Name : '+Trim(O.PrintName)+' (Outsource Item)' As ItemName,'1' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM OutsourceItemMaster O,AccountMaster A,MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND (I.Category='1' AND I.Item=O.Code) AND M.Account=A.Code AND I.Quantity>=0 AND A.Code In (" & SelectedAccounts & ") AND O.Code In (" & SelectedBooks & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                      "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'SR' As VchType,'Stock Journal (Consumed)' As Particulars,I.Quantity,'' As BoardName,'' As BookName,'Item Name : '+Trim(O.PrintName)+' (Outsource Item)' As ItemName,'1' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM OutsourceItemMaster O,AccountMaster A,MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND (I.Category='1' AND I.Item=O.Code) AND M.Account=A.Code AND I.Quantity<0 AND A.Code In (" & SelectedAccounts & ") AND O.Code In (" & SelectedBooks & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                      "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'MO' As VchType,'Material Out (To : '+(SELECT Trim(PrintName) From AccountMaster Where Code=M.AccountTo)+')' As Particulars,I.Quantity,'' As BoardName,'' As BookName,'Item Name : '+Trim(O.PrintName)+' (Outsource Item)' As ItemName,'1' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM OutsourceItemMaster O,AccountMaster A,MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND (I.Category='1' AND I.Item=O.Code) AND M.AccountFrom=A.Code AND A.Code In (" & SelectedAccounts & ") AND O.Code In (" & SelectedBooks & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                      "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'MI' As VchType,'Material In (From : '+(SELECT Trim(PrintName) From AccountMaster Where Code=M.AccountFrom)+')' As Particulars,I.Quantity,'' As BoardName,'' As BookName,'Item Name : '+Trim(O.PrintName)+' (Outsource Item)' As ItemName,'1' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM OutsourceItemMaster O,AccountMaster A,MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND (I.Category='1' AND I.Item=O.Code) AND M.AccountTo=A.Code AND A.Code In (" & SelectedAccounts & ") AND O.Code In (" & SelectedBooks & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                      "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'PC' As VchType,'Material Consumed' As Particulars,I.Quantity*(SELECT ActualQuantity FROM BookPOChild08 WHERE Code=M.Code),'' As BoardName,'' As BookName,'Item Name : '+Trim(O.PrintName)+' (Outsource Item)' As ItemName,'1' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM OutsourceItemMaster O,AccountMaster A,BookPOParent M,BookPOChild0801 I WHERE M.Code=I.Code AND M.Type<>'O' AND LEFT(M.Code,1)<>'*' AND (I.Category='1' AND I.Item=O.Code) AND M.Binder=A.Code AND A.Code In (" & SelectedAccounts & ") AND O.Code In (" & SelectedBooks & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# "
        ElseIf Val(ListView1.SelectedItem.SubItems(1)) = 3 Then
            FreshBookQuantity = "(SELECT IIF(ISNULL(SUM(OpBal)),0,SUM(OpBal)) FROM AccountChild0801 WHERE Category='3' AND Item=O.Code AND Code=A.Code)+" & _
                                               "(SELECT IIF(ISNULL(SUM(I.Quantity)),0,SUM(I.Quantity)) FROM MaterialIOParent M,MaterialIOChild I WHERE M.Code=I.Code AND Category='3' AND Item=O.Code AND Godown=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)+" & _
                                               "(SELECT IIF(ISNULL(SUM(I.Quantity)),0,SUM(I.Quantity)) FROM MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND Category='3' AND Item=O.Code AND Account=A.Code AND I.Quantity>=0 AND Date<#" & GetDate(MhDateInput1.Text) & "#)-" & _
                                               "(SELECT IIF(ISNULL(SUM(I.Quantity)),0,SUM(I.Quantity)) FROM MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND Category='3' AND Item=O.Code AND Account=A.Code AND I.Quantity<0 AND Date<#" & GetDate(MhDateInput1.Text) & "#)-" & _
                                               "(SELECT IIF(ISNULL(SUM(I.Quantity)),0,SUM(I.Quantity)) FROM MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND Category='3' AND Item=O.Code AND AccountFROM=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)+" & _
                                               "(SELECT IIF(ISNULL(SUM(I.Quantity)),0,SUM(I.Quantity)) FROM MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND Category='3' AND Item=O.Code AND AccountTo=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)-" & _
                                               "(SELECT IIF(ISNULL(SUM(I.Quantity)),0,SUM(I.Quantity*Quantity*(SELECT ActualQuantity FROM BookPOChild08 WHERE Code=M.Code))) FROM BookPOParent M,BookPOChild0801 I WHERE M.Code=I.Code AND M.Type<>'O' AND LEFT(M.Code,1)<>'*' AND Category='3' AND Item=O.Code AND Binder=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)"
            SQL = "SELECT '' As VchNo,#" & CDate(GetDate(MhDateInput1.Text)) - 1 & "# As VchDate,'OB' As VchType,'Opening Balance' As Particulars," & FreshBookQuantity & " As Quantity,'' As BoardName,'' As BookName,'Item Name : '+Trim(O.PrintName)+' (Fresh Book)' As ItemName,'3' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster O,AccountMaster A WHERE A.Code In (" & SelectedAccounts & ") AND O.Code In (" & SelectedBooks & ") AND (" & FreshBookQuantity & ") <> 0 UNION ALL " & _
                      "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'PI' As VchType,'Material In (From : '+(SELECT Trim(PrintName) From AccountMaster Where Code=M.Source)+')' As Particulars,I.Quantity,'' As BoardName,'' As BookName,'Item Name : '+Trim(O.PrintName)+' (Fresh Book)' As ItemName,'3' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster O,AccountMaster A,MaterialIOParent M,MaterialIOChild I WHERE M.Code=I.Code AND (I.Category='3' AND I.Item=O.Code) AND I.Godown=A.Code AND A.Code In (" & SelectedAccounts & ") AND O.Code In (" & SelectedBooks & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                      "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'SI' As VchType,'Stock Journal (Generated)' As Particulars,I.Quantity,'' As BoardName,'' As BookName,'Item Name : '+Trim(O.PrintName)+' (Fresh Book)' As ItemName,'3' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster O,AccountMaster A,MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND (I.Category='3' AND I.Item=O.Code) AND M.Account=A.Code AND I.Quantity>=0 AND A.Code In (" & SelectedAccounts & ") AND O.Code In (" & SelectedBooks & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                      "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'SR' As VchType,'Stock Journal (Consumed)' As Particulars,I.Quantity,'' As BoardName,'' As BookName,'Item Name : '+Trim(O.PrintName)+' (Fresh Book)' As ItemName,'3' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster O,AccountMaster A,MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND (I.Category='3' AND I.Item=O.Code) AND M.Account=A.Code AND I.Quantity<0 AND A.Code In (" & SelectedAccounts & ") AND O.Code In (" & SelectedBooks & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                      "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'MO' As VchType,'Material Out (To : '+(SELECT Trim(PrintName) From AccountMaster Where Code=M.AccountTo)+')' As Particulars,I.Quantity,'' As BoardName,'' As BookName,'Item Name : '+Trim(O.PrintName)+' (Fresh Book)' As ItemName,'3' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster O,AccountMaster A,MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND (I.Category='3' AND I.Item=O.Code) AND M.AccountFrom=A.Code AND A.Code In (" & SelectedAccounts & ") AND O.Code In (" & SelectedBooks & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                      "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'MI' As VchType,'Material In (From : '+(SELECT Trim(PrintName) From AccountMaster Where Code=M.AccountFrom)+')' As Particulars,I.Quantity,'' As BoardName,'' As BookName,'Item Name : '+Trim(O.PrintName)+' (Fresh Book)' As ItemName,'3' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster O,AccountMaster A,MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND (I.Category='3' AND I.Item=O.Code) AND M.AccountTo=A.Code AND A.Code In (" & SelectedAccounts & ") AND O.Code In (" & SelectedBooks & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                      "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'PC' As VchType,'Material Consumed' As Particulars,I.Quantity*(SELECT ActualQuantity FROM BookPOChild08 WHERE Code=M.Code),'' As BoardName,'' As BookName,'Item Name : '+Trim(O.PrintName)+' (Fresh Book)' As ItemName,'3' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster O,AccountMaster A,BookPOParent M,BookPOChild0801 I WHERE M.Code=I.Code AND M.Type<>'O' AND LEFT(M.Code,1)<>'*' AND (I.Category='3' AND I.Item=O.Code) AND M.Binder=A.Code AND A.Code In (" & SelectedAccounts & ") AND O.Code In (" & SelectedBooks & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# "
        ElseIf Val(ListView1.SelectedItem.SubItems(1)) = 4 Then
            RepairBookQuantity = "(SELECT IIF(ISNULL(SUM(OpBal)),0,SUM(OpBal)) FROM AccountChild0801 WHERE Category='4' AND Item=O.Code AND Code=A.Code)+" & _
                                                "(SELECT IIF(ISNULL(SUM(I.Quantity)),0,SUM(I.Quantity)) FROM MaterialIOParent M,MaterialIOChild I WHERE M.Code=I.Code AND Category='4' AND Item=O.Code AND Godown=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)+" & _
                                                "(SELECT IIF(ISNULL(SUM(I.Quantity)),0,SUM(I.Quantity)) FROM MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND Category='4' AND Item=O.Code AND Account=A.Code AND I.Quantity>=0 AND Date<#" & GetDate(MhDateInput1.Text) & "#)-" & _
                                                "(SELECT IIF(ISNULL(SUM(I.Quantity)),0,SUM(I.Quantity)) FROM MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND Category='4' AND Item=O.Code AND Account=A.Code AND I.Quantity<0 AND Date<#" & GetDate(MhDateInput1.Text) & "#)-" & _
                                                "(SELECT IIF(ISNULL(SUM(I.Quantity)),0,SUM(I.Quantity)) FROM MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND Category='4' AND Item=O.Code AND AccountFROM=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)+" & _
                                                "(SELECT IIF(ISNULL(SUM(I.Quantity)),0,SUM(I.Quantity)) FROM MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND Category='4' AND Item=O.Code AND AccountTo=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)-" & _
                                                "(SELECT IIF(ISNULL(SUM(I.Quantity)),0,SUM(I.Quantity*Quantity*(SELECT ActualQuantity FROM BookPOChild08 WHERE Code=M.Code))) FROM BookPOParent M,BookPOChild0801 I WHERE M.Code=I.Code AND M.Type<>'O' AND LEFT(M.Code,1)<>'*' AND Category='4' AND Item=O.Code AND Binder=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)"
            SQL = "SELECT '' As VchNo,#" & CDate(GetDate(MhDateInput1.Text)) - 1 & "# As VchDate,'OB' As VchType,'Opening Balance' As Particulars," & RepairBookQuantity & " As Quantity,'' As BoardName,'' As BookName,'Item Name : '+Trim(O.PrintName)+' (Repair Book)' As ItemName,'4' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster O,AccountMaster A WHERE O.Type='R' AND A.Code In (" & SelectedAccounts & ") AND O.Code In (" & SelectedBooks & ") AND (" & RepairBookQuantity & ") <> 0 UNION ALL " & _
                      "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'PI' As VchType,'Material In (From : '+(SELECT Trim(PrintName) From AccountMaster Where Code=M.Source)+')' As Particulars,I.Quantity,'' As BoardName,'' As BookName,'Item Name : '+Trim(O.PrintName)+' (Repair Book)' As ItemName,'4' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster O,AccountMaster A,MaterialIOParent M,MaterialIOChild I WHERE M.Code=I.Code AND (I.Category='4' AND I.Item=O.Code) AND I.Godown=A.Code AND O.Type='R' AND A.Code In (" & SelectedAccounts & ") AND O.Code In (" & SelectedBooks & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                      "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'SI' As VchType,'Stock Journal (Generated)' As Particulars,I.Quantity,'' As BoardName,'' As BookName,'Item Name : '+Trim(O.PrintName)+' (Repair Book)' As ItemName,'4' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster O,AccountMaster A,MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND (I.Category='4' AND I.Item=O.Code) AND M.Account=A.Code AND I.Quantity>=0 AND O.Type='R' AND A.Code In (" & SelectedAccounts & ") AND O.Code In (" & SelectedBooks & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                      "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'SR' As VchType,'Stock Journal (Consumed)' As Particulars,I.Quantity,'' As BoardName,'' As BookName,'Item Name : '+Trim(O.PrintName)+' (Repair Book)' As ItemName,'4' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster O,AccountMaster A,MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND (I.Category='4' AND I.Item=O.Code) AND M.Account=A.Code AND I.Quantity<0 AND O.Type='R' AND A.Code In (" & SelectedAccounts & ") AND O.Code In (" & SelectedBooks & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                      "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'MO' As VchType,'Material Out (To : '+(SELECT Trim(PrintName) From AccountMaster Where Code=M.AccountTo)+')' As Particulars,I.Quantity,'' As BoardName,'' As BookName,'Item Name : '+Trim(O.PrintName)+' (Repair Book)' As ItemName,'4' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster O,AccountMaster A,MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND (I.Category='4' AND I.Item=O.Code) AND M.AccountFrom=A.Code AND O.Type='R' AND A.Code In (" & SelectedAccounts & ") AND O.Code In (" & SelectedBooks & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                      "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'MI' As VchType,'Material In (From : '+(SELECT Trim(PrintName) From AccountMaster Where Code=M.AccountFrom)+')' As Particulars,I.Quantity,'' As BoardName,'' As BookName,'Item Name : '+Trim(O.PrintName)+' (Repair Book)' As ItemName,'4' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster O,AccountMaster A,MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND (I.Category='4' AND I.Item=O.Code) AND M.AccountTo=A.Code AND O.Type='R' AND A.Code In (" & SelectedAccounts & ") AND O.Code In (" & SelectedBooks & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                      "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'PC' As VchType,'Material Consumed' As Particulars,I.Quantity*(SELECT ActualQuantity FROM BookPOChild08 WHERE Code=M.Code),'' As BoardName,'' As BookName,'Item Name : '+Trim(O.PrintName)+' (Repair Book)' As ItemName,'4' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster O,AccountMaster A,BookPOParent M,BookPOChild0801 I WHERE M.Code=I.Code AND M.Type<>'O' AND LEFT(M.Code,1)<>'*' AND (I.Category='4' AND I.Item=O.Code) AND M.Binder=A.Code AND O.Type='R' AND A.Code In (" & SelectedAccounts & ") AND O.Code In (" & SelectedBooks & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# "
        ElseIf Val(ListView1.SelectedItem.SubItems(1)) = 5 Then
            TitleQuantity = "(SELECT IIF(ISNULL(SUM(OpBal)),0,SUM(OpBal)) FROM AccountChild0801 WHERE Category='5' AND Item=O.Code AND Code=A.Code)+" & _
                                     "(SELECT IIF(ISNULL(SUM(I.Quantity)),0,SUM(I.Quantity)) FROM MaterialIOParent M,MaterialIOChild I WHERE M.Code=I.Code AND Category='5' AND Item=O.Code AND Godown=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)+" & _
                                     "(SELECT IIF(ISNULL(SUM(I.Quantity)),0,SUM(I.Quantity)) FROM MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND Category='5' AND Item=O.Code AND Account=A.Code AND I.Quantity>=0 AND Date<#" & GetDate(MhDateInput1.Text) & "#)-" & _
                                     "(SELECT IIF(ISNULL(SUM(I.Quantity)),0,SUM(I.Quantity)) FROM MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND Category='5' AND Item=O.Code AND Account=A.Code AND I.Quantity<0 AND Date<#" & GetDate(MhDateInput1.Text) & "#)-" & _
                                     "(SELECT IIF(ISNULL(SUM(I.Quantity)),0,SUM(I.Quantity)) FROM MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND Category='5' AND Item=O.Code AND AccountFROM=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)+" & _
                                     "(SELECT IIF(ISNULL(SUM(I.Quantity)),0,SUM(I.Quantity)) FROM MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND Category='5' AND Item=O.Code AND AccountTo=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)-" & _
                                     "(SELECT IIF(ISNULL(SUM(I.Quantity)),0,SUM(I.Quantity*Quantity*(SELECT ActualQuantity FROM BookPOChild08 WHERE Code=M.Code))) FROM BookPOParent M,BookPOChild0801 I WHERE M.Code=I.Code AND M.Type<>'O' AND LEFT(M.Code,1)<>'*' AND Category='5' AND Item=O.Code AND Binder=A.Code AND Date<#" & GetDate(MhDateInput1.Text) & "#)"
            SQL = "SELECT '' As VchNo,#" & CDate(GetDate(MhDateInput1.Text)) - 1 & "# As VchDate,'OB' As VchType,'Opening Balance' As Particulars," & TitleQuantity & " As Quantity,'' As BoardName,'' As BookName,'Item Name : '+Trim(O.PrintName)+' (Title)' As ItemName,'5' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster O,AccountMaster A WHERE A.Code In (" & SelectedAccounts & ") AND O.Code In (" & SelectedBooks & ") AND (" & TitleQuantity & ") <> 0 UNION ALL " & _
                      "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'PI' As VchType,'Material In (From : '+(SELECT Trim(PrintName) From AccountMaster Where Code=M.Source)+')' As Particulars,I.Quantity,'' As BoardName,'' As BookName,'Item Name : '+Trim(O.PrintName)+' (Title)' As ItemName,'5' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster O,AccountMaster A,MaterialIOParent M,MaterialIOChild I WHERE M.Code=I.Code AND (I.Category='5' AND I.Item=O.Code) AND I.Godown=A.Code AND A.Code In (" & SelectedAccounts & ") AND O.Code In (" & SelectedBooks & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                      "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'SI' As VchType,'Stock Journal (Generated)' As Particulars,I.Quantity,'' As BoardName,'' As BookName,'Item Name : '+Trim(O.PrintName)+' (Title)' As ItemName,'5' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster O,AccountMaster A,MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND (I.Category='5' AND I.Item=O.Code) AND M.Account=A.Code AND I.Quantity>=0 AND A.Code In (" & SelectedAccounts & ") AND O.Code In (" & SelectedBooks & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                      "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'SR' As VchType,'Stock Journal (Consumed)' As Particulars,I.Quantity,'' As BoardName,'' As BookName,'Item Name : '+Trim(O.PrintName)+' (Title)' As ItemName,'5' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster O,AccountMaster A,MaterialSVParent M,MaterialSVChild I WHERE M.Code=I.Code AND (I.Category='5' AND I.Item=O.Code) AND M.Account=A.Code AND I.Quantity<0 AND A.Code In (" & SelectedAccounts & ") AND O.Code In (" & SelectedBooks & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                      "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'MO' As VchType,'Material Out (To : '+(SELECT Trim(PrintName) From AccountMaster Where Code=M.AccountTo)+')' As Particulars,I.Quantity,'' As BoardName,'' As BookName,'Item Name : '+Trim(O.PrintName)+' (Title)' As ItemName,'5' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster O,AccountMaster A,MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND (I.Category='5' AND I.Item=O.Code) AND M.AccountFrom=A.Code AND A.Code In (" & SelectedAccounts & ") AND O.Code In (" & SelectedBooks & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                      "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'MI' As VchType,'Material In (From : '+(SELECT Trim(PrintName) From AccountMaster Where Code=M.AccountFrom)+')' As Particulars,I.Quantity,'' As BoardName,'' As BookName,'Item Name : '+Trim(O.PrintName)+' (Title)' As ItemName,'5' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster O,AccountMaster A,MaterialMVParent M,MaterialMVChild I WHERE M.Code=I.Code AND (I.Category='5' AND I.Item=O.Code) AND M.AccountTo=A.Code AND A.Code In (" & SelectedAccounts & ") AND O.Code In (" & SelectedBooks & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# UNION ALL " & _
                      "SELECT Trim(M.Name) As VchNo,M.Date As VchDate,'PC' As VchType,'Material Consumed' As Particulars,I.Quantity*(SELECT ActualQuantity FROM BookPOChild08 WHERE Code=M.Code),'' As BoardName,'' As BookName,'Item Name : '+Trim(O.PrintName)+' (Title)' As ItemName,'5' As ItemType,'Godown Name : '+Trim(A.PrintName) As GodownName FROM BookMaster O,AccountMaster A,BookPOParent M,BookPOChild0801 I WHERE M.Code=I.Code AND M.Type<>'O' AND LEFT(M.Code,1)<>'*' AND (I.Category='5' AND I.Item=O.Code) AND M.Binder=A.Code AND A.Code In (" & SelectedAccounts & ") AND O.Code In (" & SelectedBooks & ") And M.Date>=#" & GetDate(MhDateInput1.Text) & "# And M.Date<=#" & GetDate(MhDateInput2.Text) & "# "
        End If
        rstMaterialStockRegister.Open SQL & "ORDER BY GodownName,ItemType,ItemName,VchDate,VchNo", CxnDatabase, adOpenKeyset, adLockReadOnly
    End If
    Screen.MousePointer = vbNormal
    If rstMaterialStockRegister.RecordCount = 0 Then
        On Error GoTo 0
        Exit Sub
    End If
    rptMaterialStockRegister.Database.SetDataSource rstMaterialStockRegister, 3, 1
    rptMaterialStockRegister.DiscardSavedData
    Set CRXParamDefs = rptMaterialStockRegister.ParameterFields
    For Each CRXParamDef In CRXParamDefs
        If CRXParamDef.ParameterFieldName = "PF1" Then
            CRXParamDef.SetCurrentValue (IIf(Check1.Value, 0, 0.1))
        ElseIf CRXParamDef.ParameterFieldName = "PF2" Then
            CRXParamDef.SetCurrentValue (IIf(Option1.Value, "D", "S"))
        ElseIf CRXParamDef.ParameterFieldName = "PF3" Then
            CRXParamDef.SetCurrentValue (ReportType)
        End If
    Next
    rptMaterialStockRegister.EnableParameterPrompting = False
    If ReportType = "1" Then
        If Option2.Value Then
            rptMaterialStockRegister.PaperOrientation = crLandscape
        Else
            rptMaterialStockRegister.PaperOrientation = crPortrait
        End If
    Else
        rptMaterialStockRegister.PaperOrientation = crPortrait
    End If
    If OutputTo = "S" Then
        Set FrmReportViewer.Report = rptMaterialStockRegister
        FrmReportViewer.Show vbModal
    Else
        rptMaterialStockRegister.PrintOut
    End If
    Set rptMaterialStockRegister = Nothing
    On Error GoTo 0
End Sub

