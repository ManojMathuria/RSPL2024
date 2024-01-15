VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmTatRegister 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tat Register"
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
            Picture         =   "TatRegister.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TatRegister.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TatRegister.frx":0658
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   3200
      Left            =   45
      TabIndex        =   7
      Top             =   345
      Width           =   7530
      _Version        =   65536
      _ExtentX        =   13282
      _ExtentY        =   5644
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
      Picture         =   "TatRegister.frx":076C
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
         Left            =   3720
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
         Left            =   6090
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
         Left            =   5040
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
         Picture         =   "TatRegister.frx":0788
         Picture         =   "TatRegister.frx":07A4
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
         Caption         =   " &To"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   -2147483634
         Picture         =   "TatRegister.frx":07C0
         Picture         =   "TatRegister.frx":07DC
      End
      Begin MSMask.MaskEdBox MhDateInput2 
         Height          =   330
         Left            =   2550
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
Attribute VB_Name = "FrmTatRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstCompanyMaster As New ADODB.Recordset
Dim rstTatRegister As New ADODB.Recordset
Dim rstPrinterList As New ADODB.Recordset
Dim OutputTo As String
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    
    CenterForm Me
    BusySystemIndicator True
    rstCompanyMaster.Open "Select PrintName From CompanyMaster", CxnDatabase, adOpenKeyset, adLockReadOnly
    rstPrinterList.Open "Select Name, Code From AccountMaster Where Type In ('05','06') Order by Name", CxnDatabase, adOpenKeyset, adLockReadOnly
    rstPrinterList.ActiveConnection = Nothing
    Call FillList(ListView1, "List of Printers...", rstPrinterList)
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
    Call CloseForm(FrmTatRegister)
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
        Call CloseForm(FrmTatRegister)
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CloseRecordset(rstCompanyMaster)
    Call CloseRecordset(rstPrinterList)
    Call CloseRecordset(rstTatRegister)
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
        PrintTatRegister
    ElseIf Button.Index = 2 Then
        OutputTo = "P"
        PrintTatRegister
    ElseIf Button.Index = 3 Then
        Call CloseForm(FrmTatRegister)
    End If
End Sub
Private Sub PrintTatRegister()
    Dim CRXParamDefs As CRAXDRT.ParameterFieldDefinitions
    Dim CRXParamDef As CRAXDRT.ParameterFieldDefinition
    Dim SelectedPrinters As String
    Dim TatQuantity As Long
    
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    rptTatRegister.Text11.SetText "Tat Register (" & IIf(Option1.Value, "Detailed", "Summarised") & ")"
    rptTatRegister.Text12.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
    rptTatRegister.Text13.SetText "From [" + Format(GetDate(MhDateInput1.Text), "dd-mm-yyyy") + "] To [" + Format(GetDate(MhDateInput2.Text), "dd-mm-yyyy") + "]"
    If rstTatRegister.State = adStateOpen Then
        rstTatRegister.Close
    End If
    SelectedPrinters = SelectedItems(ListView1)
    TatQuantity = "(Select iif(IsNull(Sum(OpBalTat)),0,Sum(OpBalTat)) From PaperChild Where Account=AccountMaster.Code)+" & _
                           "(Select iif(IsNull(Sum(Tat)),0,Sum(Tat)) From PaperIOParent,PaperIOChild Where PaperIOParent.Code=PaperIOChild.Code And Account=AccountMaster.Code And OrderType='1' And Date<#" & GetDate(MhDateInput1.Text) & "#)-" & _
                           "(Select iif(IsNull(Sum(Tat)),0,Sum(Tat)) From PaperMVParent,PaperMVChild Where PaperMVParent.Code=PaperMVChild.Code And AccountFrom=AccountMaster.Code And MovementType='1' And Date<#" & GetDate(MhDateInput1.Text) & "#)+" & _
                           "(Select iif(IsNull(Sum(Tat)),0,Sum(Tat)) From PaperMVParent,PaperMVChild Where PaperMVParent.Code=PaperMVChild.Code And AccountTo=AccountMaster.Code And MovementType='1' And Date<#" & GetDate(MhDateInput1.Text) & "#)-" & _
                           "(Select iif(IsNull(Sum(Quantity)),0,Sum(Quantity)) From TatRVParent,TatRVChild Where TatRVParent.Code=TatRVChild.Code And Printer=AccountMaster.Code And Date<#" & GetDate(MhDateInput1.Text) & "#)"
    rstTatRegister.Open "Select 'Press Name : '+Trim(PrintName) As PressName,'' As VchNo,#" & CDate(GetDate(MhDateInput1.Text)) - 1 & "# As VchDate,'OB' As VchType,'Opening Balance' As Particulars," & TatQuantity & " As Quantity From AccountMaster Where Code In (" & SelectedPrinters & ") And (" & TatQuantity & ") <> 0 Union " & _
                                      "Select 'Press Name : '+Trim(PrintName) As PressName,Trim(PaperIOParent.Name) As VchNo,Date As VchDate,'TI' As VchType,'Tat Issued' As Particulars,Tat As Quantity From AccountMaster,PaperIOParent,PaperIOChild Where AccountMaster.Code=Account And PaperIOParent.Code=PaperIOChild.Code And OrderType='1' And AccountMaster.Code In (" & SelectedPrinters & ") And Date>=#" & GetDate(MhDateInput1.Text) & "# And Date<=#" & GetDate(MhDateInput2.Text) & "# Union " & _
                                       "Select 'Press Name : '+Trim(PrintName) As PressName,Trim(PaperMVParent.Name) As VchNo,Date As VchDate,'MO' As VchType,'Tat Out' As Particulars,Tat As Quantity From AccountMaster,PaperMVParent,PaperMVChild Where AccountMaster.Code=AccountFrom And PaperMVParent.Code=PaperMVChild.Code And MovementType='1' And AccountMaster.Code In (" & SelectedPrinters & ") And Date>=#" & GetDate(MhDateInput1.Text) & "# And Date<=#" & GetDate(MhDateInput2.Text) & "# Union " & _
                                       "Select 'Press Name : '+Trim(PrintName) As PressName,Trim(PaperMVParent.Name) As VchNo,Date As VchDate,'MI' As VchType,'Tat In' As Particulars,Tat As Quantity From AccountMaster,PaperMVParent,PaperMVChild Where AccountMaster.Code=AccountTo And PaperMVParent.Code=PaperMVChild.Code And MovementType='1' And AccountMaster.Code In (" & SelectedPrinters & ") And Date>=#" & GetDate(MhDateInput1.Text) & "# And Date<=#" & GetDate(MhDateInput2.Text) & "# Union " & _
                                       "Select 'Press Name : '+Trim(PrintName) As PressName,Trim(TatRVParent.Name) As VchNo,Date As VchDate,'TR' As VchType,'Tat Received' As Particulars,Quantity From AccountMaster,TatRVParent,TatRVChild Where AccountMaster.Code=Printer And TatRVParent.Code=TatRVChild.Code And AccountMaster.Code In (" & SelectedPrinters & ") And Date>=#" & GetDate(MhDateInput1.Text) & "# And Date<=#" & GetDate(MhDateInput2.Text) & "# Order By PressName,VchDate,VchNo", CxnDatabase, adOpenKeyset, adLockOptimistic
    Screen.MousePointer = vbNormal
    If rstTatRegister.RecordCount = 0 Then
        On Error GoTo 0
        Exit Sub
    End If
    rptTatRegister.Database.SetDataSource rstTatRegister, 3, 1
    rptTatRegister.DiscardSavedData
    Set CRXParamDefs = rptTatRegister.ParameterFields
    For Each CRXParamDef In CRXParamDefs
        If CRXParamDef.ParameterFieldName = "PF1" Then
            CRXParamDef.SetCurrentValue (IIf(Check1.Value, 0, 0.1))
        ElseIf CRXParamDef.ParameterFieldName = "PF2" Then
            CRXParamDef.SetCurrentValue (IIf(Option1.Value, "D", "S"))
        End If
    Next
    rptTatRegister.EnableParameterPrompting = False
    If OutputTo = "S" Then
        Set FrmReportViewer.Report = rptTatRegister
        FrmReportViewer.Show vbModal
    Else
        rptTatRegister.PrintOut
    End If
    Set rptTatRegister = Nothing
    On Error GoTo 0
End Sub
