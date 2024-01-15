VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Begin VB.Form FrmPaperIssueRegister 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Paper Issue Register"
   ClientHeight    =   6660
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7605
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "PaperIssueRegister.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6660
   ScaleWidth      =   7605
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   7605
      _ExtentX        =   13414
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
            Picture         =   "PaperIssueRegister.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PaperIssueRegister.frx":0986
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PaperIssueRegister.frx":0A9A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   6255
      Left            =   45
      TabIndex        =   8
      Top             =   345
      Width           =   7515
      _Version        =   65536
      _ExtentX        =   13256
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
      Picture         =   "PaperIssueRegister.frx":0BAE
      Begin VB.OptionButton Option3 
         Caption         =   "With Bill"
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
         Left            =   5400
         TabIndex        =   3
         Top             =   10
         Width           =   1365
      End
      Begin VB.OptionButton Option2 
         Caption         =   "All"
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
         Left            =   6840
         TabIndex        =   4
         Top             =   10
         Value           =   -1  'True
         Width           =   525
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Without Bill"
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
         Width           =   1365
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   5940
         Left            =   0
         TabIndex        =   5
         Top             =   320
         Width           =   3765
         _ExtentX        =   6641
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
         TabIndex        =   9
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
         Picture         =   "PaperIssueRegister.frx":0BCA
         Picture         =   "PaperIssueRegister.frx":0BE6
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
         Height          =   330
         Left            =   1920
         TabIndex        =   10
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
         Picture         =   "PaperIssueRegister.frx":0C02
         Picture         =   "PaperIssueRegister.frx":0C1E
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
         Calendar        =   "PaperIssueRegister.frx":0C3A
         Caption         =   "PaperIssueRegister.frx":0D52
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "PaperIssueRegister.frx":0DBE
         Keys            =   "PaperIssueRegister.frx":0DDC
         Spin            =   "PaperIssueRegister.frx":0E3A
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
         Calendar        =   "PaperIssueRegister.frx":0E62
         Caption         =   "PaperIssueRegister.frx":0F7A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "PaperIssueRegister.frx":0FE6
         Keys            =   "PaperIssueRegister.frx":1004
         Spin            =   "PaperIssueRegister.frx":1062
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
         Height          =   5940
         Left            =   3755
         TabIndex        =   6
         Top             =   320
         Width           =   3765
         _ExtentX        =   6641
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
End
Attribute VB_Name = "FrmPaperIssueRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstCompanyMaster As New ADODB.Recordset
Dim rstPaperIssueRegister As New ADODB.Recordset
Dim rstSupplierList As New ADODB.Recordset
Dim rstPrinterList As New ADODB.Recordset
Dim OutputTo As String
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    
    CenterForm Me
    BusySystemIndicator True
    rstCompanyMaster.Open "Select PrintName From CompanyMaster", CxnDatabase, adOpenKeyset, adLockReadOnly
    rstSupplierList.Open "Select Name, Code From AccountMaster Where Type ='01' Order By Name", CxnDatabase, adOpenKeyset, adLockReadOnly
    rstSupplierList.ActiveConnection = Nothing
    rstPrinterList.Open "Select Name,Code From AccountMaster Where Type IN ('05','06','08') Order By Name", CxnDatabase, adOpenKeyset, adLockReadOnly
    rstPrinterList.ActiveConnection = Nothing
    Call FillList(ListView1, "List of Suppliers...", rstSupplierList)
    Call FillList(ListView2, "List of Printers...", rstPrinterList)
    MhDateInput1.Text = Format(FinancialYearFrom, "dd-mm-yyyy")
    If Format(FinancialYearTo, "yyyymmdd") < Format(Date, "yyyymmdd") Then
        MhDateInput2.Text = Format(FinancialYearTo, "dd-mm-yyyy")
    Else
        MhDateInput2.Text = Format(Date, "dd-mm-yyyy")
    End If
    Option1.Value = True
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
    Call CloseRecordset(rstSupplierList)
    Call CloseRecordset(rstPrinterList)
    Call CloseRecordset(rstPaperIssueRegister)
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
        PrintPapeIssueRegister
    ElseIf Button.Index = 2 Then
        OutputTo = "P"
        PrintPapeIssueRegister
    ElseIf Button.Index = 3 Then
        CloseForm Me
    End If
End Sub
Private Sub PrintPapeIssueRegister()
    Dim oExcel As Object
    Dim R As Long
    Dim Cnt As Long
    Dim TotalBundles As Double
    On Error Resume Next
    If Not FileExist(App.Path & "\Template\Paper Issue Register.xlsx") Then Exit Sub
    Screen.MousePointer = vbHourglass
    If rstPaperIssueRegister.State = adStateOpen Then rstPaperIssueRegister.Close
    rstPaperIssueRegister.Open "SELECT T.Date,S.PrintName As SupplierName,C2.QuantityOther As Quantity,I.PrintName As PaperName,I.[Weight/Ream],P.PrintName As PrinterName,T.BillNo,IIF(T.OrderType='1','B-','T-')+Trim(T.Name) As PONo,T.ChallanNo,IIF(ISNULL(T.DeliveryStartDate),'',FORMAT(T.DeliveryStartDate,'dd-MM-yy'))+' To '+IIF(ISNULL(T.DeliveryEndDate),'',FORMAT(T.DeliveryEndDate,'dd-MM-yy')),C2.Code,C2.Account,C2.Paper,C2.Narration " & _
                                                  "FROM ((((PaperPOParent T INNER JOIN PaperPOChild C1 ON C1.Code=T.Code) INNER JOIN PaperIOChild C2 ON C2.Code=T.Code) INNER JOIN AccountMaster S ON S.Code=T.Supplier) INNER JOIN AccountMaster P ON P.Code=C2.Account) INNER JOIN PaperMaster I ON I.Code=C2.Paper WHERE T.Date>=#" & GetDate(MhDateInput1.Text) & "# AND T.Date<=#" & GetDate(MhDateInput2.Text) & "# AND C2.Account IN (" & SelectedItems(ListView2) & ") AND T.Supplier IN (" & SelectedItems(ListView1) & ") ORDER BY P.PrintName,I.PrintName,T.Name", CxnDatabase, adOpenKeyset, adLockReadOnly
    If rstPaperIssueRegister.RecordCount = 0 Then Screen.MousePointer = vbNormal: On Error GoTo 0: Exit Sub
    DoEvents
    Set oExcel = CreateObject("Excel.Application")
    oExcel.Workbooks.Open (App.Path & "\Template\Paper Issue Register")
    oExcel.DisplayAlerts = False
    oExcel.Workbooks.Item(1).SaveAs (App.Path & "\Report\Paper Issue Register (" & CompCode & ")")
    oExcel.Application.DisplayAlerts = True
    oExcel.Sheets("Sheet1").Select
    oExcel.Application.Cells(1, 1).Value = Trim(rstCompanyMaster.Fields("PrintName").Value)
    oExcel.Application.Cells(2, 1).Value = "Paper Issue Register From [" & Format(MhDateInput1, "dd-MMM-yyyy") & "] To [" & Format(MhDateInput2, "dd-MMM-yyyy") & "]"
    R = 5: Cnt = 1
    Do While Not rstPaperIssueRegister.EOF
        If Not Option2.Value Then
            If Option1.Value Then
                If CheckNull(Trim(rstPaperIssueRegister.Fields("BillNo").Value)) <> "" Then GoTo Skip
            Else
                If CheckNull(Trim(rstPaperIssueRegister.Fields("BillNo").Value)) = "" Then GoTo Skip
            End If
        End If
        oExcel.Application.Cells(R, 1).Value = Cnt
        oExcel.Application.Cells(R, 2).Value = Format(rstPaperIssueRegister.Fields("Date").Value, "dd-MMM-yyyy")
        oExcel.Application.Cells(R, 3).Value = Trim(rstPaperIssueRegister.Fields("SupplierName").Value)
        oExcel.Application.Cells(R, 4).Value = Trim(rstPaperIssueRegister.Fields("PaperName").Value)
        oExcel.Application.Cells(R, 5).Value = Format(Val(rstPaperIssueRegister.Fields("Quantity").Value), "0.000")
        oExcel.Application.Cells(R, 6).Value = Format(Int(Val(rstPaperIssueRegister.Fields("Quantity").Value)) * Val(rstPaperIssueRegister.Fields("Weight/Ream").Value), "0.000")
        If Val(rstPaperIssueRegister.Fields("Quantity").Value) - Int(Val(rstPaperIssueRegister.Fields("Quantity").Value)) > 0 Then oExcel.Application.Cells(R, 6).Value = Val(oExcel.Application.Cells(R, 6).Value) + ((Val(rstPaperIssueRegister.Fields("Quantity").Value) - Int(Val(rstPaperIssueRegister.Fields("Quantity").Value))) * 1000) * (Val(rstPaperIssueRegister.Fields("Weight/Ream").Value) / 500)
        oExcel.Application.Cells(R, 6).Value = Val(oExcel.Application.Cells(R, 6).Value) / 1000
        oExcel.Application.Cells(R, 7).Value = Trim(rstPaperIssueRegister.Fields("PrinterName").Value)
        oExcel.Application.Cells(R, 8).Value = CheckNull(Trim(rstPaperIssueRegister.Fields("BillNo").Value))
        oExcel.Application.Cells(R, 9).Value = CheckNull(Trim(rstPaperIssueRegister.Fields("PONo").Value))
        oExcel.Application.Cells(R, 10).Value = CheckNull(Trim(rstPaperIssueRegister.Fields("Narration").Value))
        oExcel.Application.Cells(R, 11).Value = CheckNull(Trim(rstPaperIssueRegister.Fields("ChallanNo").Value))
        oExcel.Application.Cells(R, 12).Value = CheckNull(Trim(rstPaperIssueRegister.Fields("DeliveryDate").Value))
        oExcel.Application.Cells(R, 13).Value = CheckNull(Trim(rstPaperIssueRegister.Fields("BiltyNo").Value))
        oExcel.Application.Cells(R, 14).Value = CheckNull(Trim(rstPaperIssueRegister.Fields("BiltyDate").Value))
        oExcel.Application.Cells(R, 16382).Value = CheckNull(Trim(rstPaperIssueRegister.Fields("Paper").Value))
        oExcel.Application.Cells(R, 16383).Value = CheckNull(Trim(rstPaperIssueRegister.Fields("Account").Value))
        oExcel.Application.Cells(R, 16384).Value = CheckNull(Trim(rstPaperIssueRegister.Fields("Code").Value))
        Cnt = Cnt + 1: R = R + 1
Skip:
        rstPaperIssueRegister.MoveNext
    Loop
    oExcel.Sheets("Sheet1").Activate
    oExcel.Columns("A:M").EntireColumn.AutoFit
    oExcel.Workbooks.Item(1).Save
    Screen.MousePointer = vbNormal
    If OutputTo = "S" Then oExcel.Range("A1").Activate: oExcel.Application.Visible = True Else oExcel.Workbooks.Item(1).PrintOut
    Set oExcel = Nothing
    On Error GoTo 0
End Sub
