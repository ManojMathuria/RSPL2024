VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmPrintPlanningRegister 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Planning Register"
   ClientHeight    =   3570
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
   Icon            =   "PrintPlanningRegister.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3570
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
            Picture         =   "PrintPlanningRegister.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PrintPlanningRegister.frx":0986
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PrintPlanningRegister.frx":0A9A
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
      Picture         =   "PrintPlanningRegister.frx":0BAE
      Begin VB.OptionButton Option2 
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
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   6090
         TabIndex        =   3
         Top             =   10
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "With Nil"
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
         Left            =   3840
         TabIndex        =   2
         Top             =   10
         Width           =   1095
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   2875
         Left            =   3755
         TabIndex        =   5
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
      Begin MSComctlLib.ListView ListView1 
         Height          =   2875
         Left            =   0
         TabIndex        =   4
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
         Picture         =   "PrintPlanningRegister.frx":0BCA
         Picture         =   "PrintPlanningRegister.frx":0BE6
      End
      Begin MSMask.MaskEdBox MhDateInput1 
         Height          =   330
         Left            =   840
         TabIndex        =   0
         Top             =   0
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##-##-####"
         PromptChar      =   " "
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
         Height          =   330
         Left            =   1920
         TabIndex        =   9
         Top             =   0
         Width           =   760
         _Version        =   65536
         _ExtentX        =   1341
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
         Picture         =   "PrintPlanningRegister.frx":0C02
         Picture         =   "PrintPlanningRegister.frx":0C1E
      End
      Begin MSMask.MaskEdBox MhDateInput2 
         Height          =   330
         Left            =   2670
         TabIndex        =   1
         Top             =   0
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##-##-####"
         PromptChar      =   " "
      End
   End
End
Attribute VB_Name = "FrmPrintPlanningRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstCompanyMaster As New ADODB.Recordset
Dim rstPrintPlanningRegister As New ADODB.Recordset
Dim rstBookList As New ADODB.Recordset
Dim rstBoardList As New ADODB.Recordset
Dim OutputTo As String
Public PlanningType As String
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    
    CenterForm Me
    BusySystemIndicator True
    If PlanningType = "1" Then
        Me.Caption = "Print Planning Register [Book]"
    Else
        Me.Caption = "Print Planning Register [Title]"
    End If
    rstCompanyMaster.Open "Select PrintName From CompanyMaster", CxnDatabase, adOpenKeyset, adLockReadOnly
    rstBoardList.Open "Select Name, Code From GeneralMaster Where Type = '2' Order By Name", CxnDatabase, adOpenKeyset, adLockReadOnly
    rstBoardList.ActiveConnection = Nothing
    Call FillList(ListView1, "List of Boards...", rstBoardList)
    Call BookSelection(True)
    Option2.Value = True
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
Private Sub Form_Unload(Cancel As Integer)
    Call CloseRecordset(rstCompanyMaster)
    Call CloseRecordset(rstBookList)
    Call CloseRecordset(rstBoardList)
    Call CloseRecordset(rstPrintPlanningRegister)
End Sub
Private Sub MhDateInput1_GotFocus()
    FocusSelect Me.ActiveControl
End Sub
Private Sub MhDateInput1_Validate(Cancel As Boolean)
    If Not ValidateDate(Me.ActiveControl) Then
        Cancel = True
    End If
End Sub
Private Sub MhDateInput2_GotFocus()
    FocusSelect Me.ActiveControl
End Sub
Private Sub MhDateInput2_Validate(Cancel As Boolean)
    If Not ValidateDate(Me.ActiveControl) Then
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
        Next i
    ElseIf KeyCode = vbKeyD And Shift = vbCtrlMask Then
        For i = 1 To ListView2.ListItems.Count
            ListView2.ListItems(i).Checked = False
        Next i
    End If
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    
    If Button.Index = 1 Then
        OutputTo = "S"
        PrintPrintPlanningRegister
    ElseIf Button.Index = 2 Then
        OutputTo = "P"
        PrintPrintPlanningRegister
    ElseIf Button.Index = 3 Then
        Unload Me
    End If
End Sub
Private Sub BookSelection(ByVal SelectAll As Boolean)
    If rstBookList.State = adStateOpen Then
        rstBookList.Close
    End If
    rstBookList.Open "Select Name, Code From BookMaster " & IIf(SelectAll, "", "Where Board In (" & SelectedItems(ListView1) & ")") & " Order by Name", CxnDatabase, adOpenKeyset, adLockReadOnly
    rstBookList.ActiveConnection = Nothing
    ListView2.ListItems.Clear
    Call FillList(ListView2, "List of Books...", rstBookList)
End Sub
Private Sub PrintPrintPlanningRegister()
    Dim CRXParamDefs As CRAXDRT.ParameterFieldDefinitions
    Dim CRXParamDef As CRAXDRT.ParameterFieldDefinition
    Dim SelectedBooks As String
    Dim SelectedBoards As String
    
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    rptPrintPlanningRegister.Text11.SetText "Print Planning Register (" & IIf(PlanningType = "1", "Book", "Title") & ")"
    rptPrintPlanningRegister.Text12.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptPrintPlanningRegister.Text13.SetText "From [" + Format(GetDate(MhDateInput1.Text), "dd-mm-yyyy") + "] To [" + Format(GetDate(MhDateInput2.Text), "dd-mm-yyyy") + "]"
    If rstPrintPlanningRegister.State = adStateOpen Then
        rstPrintPlanningRegister.Close
    End If
    SelectedBoards = SelectedItems(ListView1)
    SelectedBooks = SelectedItems(ListView2)
    If PlanningType = "1" Then
        rstPrintPlanningRegister.Open "Select 'Board Name : '+Trim(GeneralMaster.PrintName) As BoardName,'Book Name : '+Trim(BookMaster.PrintName) As BookName,'' As PressName,Trim(PrintPVParent.Name) As VchNo,'' As OrderNo,PrintPVParent.Date As VchDate,'PP' As VchType,Quantity From GeneralMaster,BookMaster,PrintPVParent,PrintPVChild Where GeneralMaster.Code=BookMaster.Board And BookMaster.Code In (" & SelectedBooks & ") And GeneralMaster.Code In (" & SelectedBoards & ") And BookMaster.Code=PrintPVChild.Book And PrintPVParent.Code=PrintPVChild.Code And PrintPVParent.Date>=#" & GetDate(MhDateInput1.Text) & "# And  PrintPVParent.Date <= #" & GetDate(MhDateInput2.Text) & "# And PrintPVParent.PlanningType='" & PlanningType & "' Union " & _
                                                             "Select 'Board Name : '+Trim(GeneralMaster.PrintName) As BoardName,'Book Name : '+Trim(BookMaster.PrintName) As BookName,(Select Trim(PrintName) From AccountMaster Where Code=BookPOParent.BookPrinter) As PressName,Trim(PrintPVParent.Name) As VchNo,Trim(BookPOParent.Name) As OrderNo,BookPOParent.Date As VchDate,'PR' As VchType,ActualQuantity As Quantity From GeneralMaster,BookMaster,BookPOParent,BookPOChild05,PrintPVParent Where GeneralMaster.Code=BookMaster.Board And BookMaster.Code In (" & SelectedBooks & ")  And GeneralMaster.Code In (" & SelectedBoards & ")  And BookMaster.Code=BookPOParent.Book And BookPOParent.Code=BookPOChild05.Code And BookPOChild05.Ref=PrintPVParent.Code And PrintPVParent.Date>=#" & GetDate(MhDateInput1.Text) & "# And  PrintPVParent.Date<=#" & GetDate(MhDateInput2.Text) & "# And PrintPVParent.PlanningType='" & PlanningType & "' " & _
                                                             "Order By BoardName,BookName,VchNo,OrderNo", CxnDatabase, adOpenKeyset, adLockOptimistic
    Else
        rstPrintPlanningRegister.Open "Select 'Board Name : '+Trim(GeneralMaster.PrintName) As BoardName,'Book Name : '+Trim(BookMaster.PrintName) As BookName,'' As PressName,Trim(PrintPVParent.Name) As VchNo,'' As OrderNo,PrintPVParent.Date As VchDate,'PP' As VchType,Quantity From GeneralMaster,BookMaster,PrintPVParent,PrintPVChild Where GeneralMaster.Code=BookMaster.Board And BookMaster.Code In (" & SelectedBooks & ") And GeneralMaster.Code In (" & SelectedBoards & ") And BookMaster.Code=PrintPVChild.Book And PrintPVParent.Code=PrintPVChild.Code And PrintPVParent.Date>=#" & GetDate(MhDateInput1.Text) & "# And  PrintPVParent.Date <= #" & GetDate(MhDateInput2.Text) & "# And PrintPVParent.PlanningType='" & PlanningType & "' Union " & _
                                                             "Select 'Board Name : '+Trim(GeneralMaster.PrintName) As BoardName,'Book Name : '+Trim(BookMaster.PrintName) As BookName,(Select Trim(PrintName) From AccountMaster Where Code=BookPOParent.TitlePrinter) As PressName,Trim(PrintPVParent.Name) As VchNo,Trim(BookPOParent.Name) As OrderNo,BookPOParent.Date As VchDate,'PR' As VchType,ActualQuantity As Quantity From GeneralMaster,BookMaster,BookPOParent,BookPOChild06,PrintPVParent Where GeneralMaster.Code=BookMaster.Board And BookMaster.Code In (" & SelectedBooks & ")  And GeneralMaster.Code In (" & SelectedBoards & ")  And BookMaster.Code=BookPOParent.Book And BookPOParent.Code=BookPOChild06.Code And BookPOChild06.Ref=PrintPVParent.Code And PrintPVParent.Date>=#" & GetDate(MhDateInput1.Text) & "# And  PrintPVParent.Date<=#" & GetDate(MhDateInput2.Text) & "# And PrintPVParent.PlanningType='" & PlanningType & "' " & _
                                                             "Order By BoardName,BookName,VchNo,OrderNo", CxnDatabase, adOpenKeyset, adLockOptimistic
    End If
    Screen.MousePointer = vbNormal
    If rstPrintPlanningRegister.RecordCount = 0 Then
        On Error GoTo 0
        Exit Sub
    End If
    rptPrintPlanningRegister.Database.SetDataSource rstPrintPlanningRegister, 3, 1
    rptPrintPlanningRegister.DiscardSavedData
    Set CRXParamDefs = rptPrintPlanningRegister.ParameterFields
    For Each CRXParamDef In CRXParamDefs
        If CRXParamDef.ParameterFieldName = "PF1" Then
            CRXParamDef.SetCurrentValue (IIf(Option1, 1, 0))
            Exit For
        End If
    Next
    rptPrintPlanningRegister.EnableParameterPrompting = False
    If OutputTo = "S" Then
        Set FrmReportViewer.Report = rptPrintPlanningRegister
        FrmReportViewer.Show vbModal
    Else
        rptPrintPlanningRegister.PrintOut
    End If
    Set rptPrintPlanningRegister = Nothing
    On Error GoTo 0
End Sub
