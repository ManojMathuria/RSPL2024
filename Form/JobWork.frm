VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmJobWork 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Job Work Register"
   ClientHeight    =   3600
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   5235
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "JobWork.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   5235
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5235
      _ExtentX        =   9234
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
            Picture         =   "JobWork.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JobWork.frx":0986
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JobWork.frx":0A9A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   3180
      Left            =   45
      TabIndex        =   4
      Top             =   345
      Width           =   5150
      _Version        =   65536
      _ExtentX        =   9084
      _ExtentY        =   5609
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
      Picture         =   "JobWork.frx":0BAE
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
         Height          =   330
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   615
         _Version        =   65536
         _ExtentX        =   1085
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
         Picture         =   "JobWork.frx":0BCA
         Picture         =   "JobWork.frx":0BE6
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
         Height          =   330
         Left            =   1680
         TabIndex        =   6
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
         Picture         =   "JobWork.frx":0C02
         Picture         =   "JobWork.frx":0C1E
      End
      Begin TDBDate6Ctl.TDBDate MhDateInput2 
         Height          =   330
         Left            =   2310
         TabIndex        =   1
         Top             =   0
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calendar        =   "JobWork.frx":0C3A
         Caption         =   "JobWork.frx":0D52
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "JobWork.frx":0DBE
         Keys            =   "JobWork.frx":0DDC
         Spin            =   "JobWork.frx":0E3A
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
         Left            =   600
         TabIndex        =   0
         Top             =   0
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   582
         Calendar        =   "JobWork.frx":0E62
         Caption         =   "JobWork.frx":0F7A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "JobWork.frx":0FE6
         Keys            =   "JobWork.frx":1004
         Spin            =   "JobWork.frx":1062
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
      Begin MSForms.ComboBox Combo1 
         Height          =   330
         Left            =   3390
         TabIndex        =   2
         Top             =   0
         Width           =   1725
         VariousPropertyBits=   545282075
         BackColor       =   16777215
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "3043;582"
         MatchEntry      =   0
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Calibri"
         FontEffects     =   1073741825
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
   End
End
Attribute VB_Name = "FrmJobWork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstPrintJobWorkRegister As New ADODB.Recordset
Dim rstCompanyMaster As New ADODB.Recordset
Dim rstAccountList As New ADODB.Recordset
Dim rstExist As New ADODB.Recordset
Dim OutputTo As String
Dim RefCode As String
Public OrderType As String
Public ReportType As String
Private Sub Combo1_Change()
If ReportType = 1 Then
    If Combo1.Text = "Book" Then
       OrderType = 1
    Else
       OrderType = 2
    End If
Else
    OrderType = 0
End If
End Sub
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    CenterForm Me
    BusySystemIndicator True
    MhDateInput1.Text = Format(FinancialYearFrom, "dd-mm-yyyy")
    If Format(FinancialYearTo, "yyyymmdd") < Format(Date, "yyyymmdd") Then
        MhDateInput2.Text = Format(FinancialYearTo, "dd-mm-yyyy")
    Else
        MhDateInput2.Text = Format(Date, "dd-mm-yyyy")
    End If
    If ReportType = 1 Then
       Combo1.AddItem "Book", 0: Combo1.AddItem "Title", 1: Combo1.ListIndex = 0
    Else
       Combo1.AddItem "First Report", 0: Combo1.AddItem "Second Report", 1: Combo1.ListIndex = 0
    End If
    BusySystemIndicator False
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    CloseForm Me
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       Sendkeys "{TAB}", True: KeyCode = 0
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
    Call CloseRecordset(rstCompanyMaster)
    Call CloseRecordset(rstAccountList)
    Call CloseRecordset(rstPrintJobWorkRegister)
End Sub
Private Sub MhDateInput1_Validate(Cancel As Boolean)
    If Not IsDate(GetDate(MhDateInput1.Text)) Then Cancel = True
End Sub
Private Sub MhDateInput2_Validate(Cancel As Boolean)
    If Not IsDate(GetDate(MhDateInput2.Text)) Or Format(GetDate(MhDateInput2.Text), "yyyymmdd") < Format(GetDate(MhDateInput1.Text), "yyyymmdd") Then Cancel = True
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    If Button.Index = 3 Then CloseForm Me: Exit Sub
    If Button.Index = 1 Then OutputTo = "S" Else OutputTo = "P"
    PrintJobWorkRegister
End Sub

Private Sub PrintJobWorkRegister()
    Dim oExcel As Object
    Dim i As Long, Cnt As Long
    Dim Path As String
    Dim str As String
    Screen.MousePointer = vbHourglass
    DoEvents
    If rstCompanyMaster.State = adStateOpen Then
        rstCompanyMaster.Close
    End If
    rstCompanyMaster.Open "SELECT PrintName,Address1,Address2,Address3,Address4,Phone,Fax,eMail FROM CompanyMaster", CxnDatabase, adOpenKeyset, adLockReadOnly
    If rstPrintJobWorkRegister.State = adStateOpen Then
        rstPrintJobWorkRegister.Close
    End If
    If ReportType = 1 Then
       str = "SELECT P.Code,TRIM(P.Name) As OrderNo,ChallanNo,DeliveryStartDate As ChallanDate,TRIM(M1.PrintName) As SupplierName,[VAT%],VAT,BillAmount,TRIM(M2.PrintName) As PaperName,QuantityOther,[Weight/Ream],QuantityKg,[Rate/Kg],Amount,BillNo,BillDate,(Select Distinct TIN From AccountMaster Where Code In(Select Top 1 Account From PaperIOChild Where Code=P.Code)) As GSTIN FROM ((PaperPOParent P LEFT JOIN PaperPOChild C ON P.Code=C.Code) LEFT JOIN AccountMaster M1 ON P.Supplier=M1.Code) LEFT JOIN PaperMaster M2 ON M2.Code=C.Paper WHERE P.Date>=#" & GetDate(MhDateInput1.Text) & "# And P.Date<=#" & GetDate(MhDateInput2.Text) & "# And P.OrderType='" & OrderType & "' ORDER BY M2.PrintName"
    Else
        If Combo1.Text = "Second Report" Then
            str = "SELECT T.Code,T.Name As OrderNo,C.BillAmount As BinderBillAmount,"
            str = str & " M.Name As BookName,ReceivedQuantity As RecvdQty,"
            str = str & " (SELECT Name FROM AccountMaster WHERE Code=T.Binder) As BinderName,"
            str = str & " (SELECT TIN FROM AccountMaster WHERE Code=T.Binder) As BinderGSTIN,(SELECT OrderDate FROM BookPOChild08 WHERE Code=T.Code) As BinderBillDate"
            str = str & " FROM (BookPOParent T INNER JOIN bookPOChild08 C ON T.Code=C.Code) INNER JOIN BookMaster M ON T.Book=M.Code WHERE T.Type = 'F' AND LEFT(T.Code,1)<>'*' And T.Date>=#" & GetDate(MhDateInput1.Text) & "# And T.Date<=#" & GetDate(MhDateInput2.Text) & "#  ORDER BY T.Name"
        Else
            str = "SELECT T.Code,T.Name As OrderNo,C.BillAmount As BookPrinterBillAmount,"
            str = str & " (SELECT BillAmount FROM BookPOChild08  WHERE Code=T.Code And T.Type = 'F' AND LEFT(T.Code,1)<>'*') As BAmount,"
            str = str & " (SELECT BillAmount FROM BookPOChild06 WHERE Code=T.Code) As TBillAmount,"
            str = str & " (SELECT BillAmount FROM BookPOChild07 WHERE Code=T.Code) As LBillAmount,"
            
            str = str & "(SELECT BiltyNo FROM PaperPOParent WHERE Code IN(Select Code From PaperPOChildRef Where Book+Paper+Ref=T.Book+C.Paper1+C.Ref)) As ChallanNo1,"
            str = str & "(SELECT BiltyNo FROM PaperPOParent WHERE Code IN(Select Code From PaperPOChildRef Where Book+Paper+Ref=T.Book+C.Paper2+C.Ref)) As ChallanNo2,"
            str = str & "(SELECT BiltyNo FROM PaperPOParent WHERE Code IN(Select Code From PaperPOChildRef Where Book+Paper+Ref=T.Book+C.Paper4+C.Ref)) As ChallanNo3,"
            str = str & "(SELECT BiltyDate FROM PaperPOParent WHERE Code IN(Select Code From PaperPOChildRef Where Book+Paper+Ref=T.Book+C.Paper1+C.Ref)) As ChallanDate1,"
            str = str & "(SELECT BiltyDate FROM PaperPOParent WHERE Code IN(Select Code From PaperPOChildRef Where Book+Paper+Ref=T.Book+C.Paper2+C.Ref)) As ChallanDate2,"
            str = str & "(SELECT BiltyDate FROM PaperPOParent WHERE Code IN(Select Code From PaperPOChildRef Where Book+Paper+Ref=T.Book+C.Paper4+C.Ref)) As ChallanDate3,"
            
            str = str & " B.BillingQuantity  As ActualQuantity,T.ReceivedQuantity As RecvdQty,"
            str = str & "(SELECT Name FROM AccountMaster WHERE Code=T.BookPrinter) As BookPrinterName,"
            str = str & " (SELECT TIN FROM AccountMaster WHERE Code=T.BookPrinter) As BookPrinterGSTIN,(SELECT Name FROM AccountMaster WHERE Code=T.TitlePrinter) As TitlePrinterName,(SELECT TIN FROM AccountMaster WHERE Code=T.TitlePrinter) As TitlePrinterGSTIN,(SELECT Name FROM AccountMaster WHERE Code=T.Laminator) As LaminatorName,(SELECT TIN FROM AccountMaster WHERE Code=T.Laminator) As LaminatorGSTIN,(SELECT Name FROM AccountMaster WHERE Code=T.Binder) As BinderName,"
            str = str & " (SELECT TIN FROM AccountMaster WHERE Code=T.Binder) As BinderGSTIN,(SELECT BillNo FROM BookPOChild08 WHERE Code=T.Code) As BinderBillNo,"
            str = str & "(SELECT OrderDate FROM BookPOChild08 WHERE Code=T.Code) As BinderBillDate,"
            str = str & " C.BookStatus "
            str = str & " FROM ( BookPOParent T INNER JOIN bookPOChild08 B ON T.Code=B.Code ) INNER JOIN bookPOChild05 C ON T.Code=C.Code WHERE T.Type = 'F' AND LEFT(T.Code,1)<>'*' And T.Date>=#" & GetDate(MhDateInput1.Text) & "# And T.Date<=#" & GetDate(MhDateInput2.Text) & "#  ORDER BY T.Name"
        
        End If
    End If
    rstPrintJobWorkRegister.Open str, CxnDatabase, adOpenKeyset, adLockOptimistic
    On Error Resume Next
    If Not FileExist(App.Path & "\Template\Job Work.xlsx") Then Exit Sub
  
    
    If rstPrintJobWorkRegister.RecordCount = 0 Then
        DisplayError ("No Record Found")
        ShowProgressInStatusBar False
        Screen.MousePointer = vbNormal
        On Error GoTo 0
        Exit Sub
    End If
     Screen.MousePointer = vbHourglass
    DoEvents
    Set oExcel = CreateObject("Excel.Application")
    oExcel.Workbooks.Open (App.Path & "\Template\Job Work"): oExcel.DisplayAlerts = False
    Path = IIf(ReportType = 1, "Goods Sent Job Work Register", "Goods Received Job Work Register")
    oExcel.Workbooks.Item(1).SaveAs (App.Path & "\Report\" & Path & " (" & CompCode & ")")
    oExcel.Sheets("Job Work (" & IIf(ReportType = 1, "Goods Sent", "Goods Received") & ")").Visible = False
    oExcel.Visible = False

    oExcel.Cells(1, "A").Value = "RACHNA SAGAR (P)LTD."
    oExcel.Cells(2, "A").Value = IIf(ReportType = 1, "Goods Sent Job Work Register", "Goods Received Job Work Register") & " From " & Format(MhDateInput1, "dd-MMM-yyyy") & " To " & Format(MhDateInput2, "dd-MMM-yyyy")
   
    i = 5: Cnt = 1
    If ReportType = 1 Then
        oExcel.Sheets("Sheet2").Visible = False: oExcel.Sheets("Sheet1").Select
        Do While Not rstPrintJobWorkRegister.EOF
            
            oExcel.Cells(i, "A").Value = Cnt
            oExcel.Application.Cells(i, "B").Value = Trim(rstPrintJobWorkRegister.Fields("GSTIN").Value)
            oExcel.Application.Cells(i, "C").Value = Trim(rstPrintJobWorkRegister.Fields("ChallanNo").Value)
            oExcel.Application.Cells(i, "D").Value = Trim(rstPrintJobWorkRegister.Fields("ChallanDate").Value)
            oExcel.Application.Cells(i, "E").Value = Trim(rstPrintJobWorkRegister.Fields("PaperName").Value)
            oExcel.Application.Cells(i, "F").Value = "KGS"
            oExcel.Application.Cells(i, "G").Value = Trim(rstPrintJobWorkRegister.Fields("QuantityKg").Value)
            oExcel.Application.Cells(i, "H").Value = Trim(rstPrintJobWorkRegister.Fields("Amount").Value)
            oExcel.Application.Cells(i, "I").Value = "Input"
            oExcel.Application.Cells(i, "J").Value = Val(rstPrintJobWorkRegister.Fields("VAT%").Value) / 2 & " %"
            oExcel.Application.Cells(i, "K").Value = Val(rstPrintJobWorkRegister.Fields("VAT%").Value) / 2 & " %"
            oExcel.Application.Cells(i, "L").Value = "" 'rstPrintJobWorkRegister.Fields("VAT%").Value & " %"
            oExcel.Application.Cells(i, "M").Value = "" 'rstPrintJobWorkRegister.Fields("VAT%").Value
            oExcel.Application.Cells(i, "N").Value = rstPrintJobWorkRegister.Fields("OrderNo").Value
            
            rstPrintJobWorkRegister.MoveNext
            Cnt = Cnt + 1: i = i + 1
            
        Loop
    End If
    If ReportType = 2 Then
       oExcel.Sheets("Sheet1").Visible = False: oExcel.Sheets("Sheet2").Select
        Do While Not rstPrintJobWorkRegister.EOF
            If Combo1.Text = "Second Report" Then
            
               If Trim(rstPrintJobWorkRegister.Fields("BinderGSTIN").Value) <> "" Then
   
                  oExcel.Cells(i, "A").Value = Cnt
                  oExcel.Application.Cells(i, "B").Value = Trim(rstPrintJobWorkRegister.Fields("BinderGSTIN").Value)
                  oExcel.Application.Cells(i, "C").Value = "Yes"
                  oExcel.Application.Cells(i, "D").Value = "BB" & Trim(rstPrintJobWorkRegister.Fields("OrderNo").Value)
                  oExcel.Application.Cells(i, "E").Value = Trim(rstPrintJobWorkRegister.Fields("BinderBillDate").Value)
                  oExcel.Application.Cells(i, "F").Value = ""
                  oExcel.Application.Cells(i, "G").Value = ""
                  oExcel.Application.Cells(i, "H").Value = ""
                  oExcel.Application.Cells(i, "I").Value = ""
                  oExcel.Application.Cells(i, "J").Value = ""
                  oExcel.Application.Cells(i, "K").Value = "Books"
                  oExcel.Application.Cells(i, "L").Value = "No of Books"
                  oExcel.Application.Cells(i, "M").Value = rstPrintJobWorkRegister.Fields("RecvdQty").Value
                  oExcel.Application.Cells(i, "N").Value = Val(rstPrintJobWorkRegister.Fields("BinderBillAmount").Value)
                  oExcel.Application.Cells(i, "O").Value = Trim(rstPrintJobWorkRegister.Fields("OrderNo").Value)
                  oExcel.Application.Cells(i, "P").Value = rstPrintJobWorkRegister.Fields("ActualQuantity").Value
               
               End If
             Else
                
                If Trim(rstPrintJobWorkRegister.Fields("BookPrinterGSTIN").Value) <> "" Then
                    oExcel.Cells(i, "A").Value = Cnt
                    oExcel.Application.Cells(i, "B").Value = Trim(rstPrintJobWorkRegister.Fields("BookPrinterGSTIN").Value)
                    If Trim(rstPrintJobWorkRegister.Fields("BookPrinterGSTIN").Value) = Trim(rstPrintJobWorkRegister.Fields("BinderGSTIN").Value) Then
                        oExcel.Application.Cells(i, "C").Value = "Yes"
                        oExcel.Application.Cells(i, "F").Value = ""
                        oExcel.Application.Cells(i, "G").Value = ""
                        oExcel.Application.Cells(i, "H").Value = ""
                    Else
                        oExcel.Application.Cells(i, "C").Value = "No"
                        oExcel.Application.Cells(i, "F").Value = "BB" & Trim(rstPrintJobWorkRegister.Fields("OrderNo").Value)
                        oExcel.Application.Cells(i, "G").Value = Trim(rstPrintJobWorkRegister.Fields("BinderBillDate").Value)
                        oExcel.Application.Cells(i, "H").Value = Trim(rstPrintJobWorkRegister.Fields("BinderGSTIN").Value)
                    End If
                    
                    If Not IsNull(Trim(rstPrintJobWorkRegister.Fields("ChallanNo1").Value)) Then
                        oExcel.Application.Cells(i, "D").Value = Trim(rstPrintJobWorkRegister.Fields("ChallanNo1").Value)
                        oExcel.Application.Cells(i, "E").Value = Trim(rstPrintJobWorkRegister.Fields("ChallanDate1").Value)
                    End If
                    
                    If Not IsNull(Trim(rstPrintJobWorkRegister.Fields("ChallanNo2").Value)) Then
                        oExcel.Application.Cells(i, "D").Value = Trim(rstPrintJobWorkRegister.Fields("ChallanNo2").Value)
                        oExcel.Application.Cells(i, "E").Value = Trim(rstPrintJobWorkRegister.Fields("ChallanDate2").Value)
                    End If
                    
                    If Not IsNull(Trim(rstPrintJobWorkRegister.Fields("ChallanNo3").Value)) Then
                        oExcel.Application.Cells(i, "D").Value = Trim(rstPrintJobWorkRegister.Fields("ChallanNo3").Value)
                        oExcel.Application.Cells(i, "E").Value = Trim(rstPrintJobWorkRegister.Fields("ChallanDate3").Value)
                    End If
                    
                    oExcel.Application.Cells(i, "I").Value = ""
                    oExcel.Application.Cells(i, "J").Value = ""
                    oExcel.Application.Cells(i, "K").Value = "Books"
                    oExcel.Application.Cells(i, "L").Value = "No of Books"
                    oExcel.Application.Cells(i, "M").Value = rstPrintJobWorkRegister.Fields("RecvdQty").Value
                    oExcel.Application.Cells(i, "N").Value = rstPrintJobWorkRegister.Fields("BookPrinterBillAmount").Value
                    oExcel.Application.Cells(i, "O").Value = Trim(rstPrintJobWorkRegister.Fields("OrderNo").Value)
                    oExcel.Application.Cells(i, "P").Value = rstPrintJobWorkRegister.Fields("ActualQuantity").Value
                    
                End If
                
            End If
             rstPrintJobWorkRegister.MoveNext
            Cnt = Cnt + 1: i = i + 1
         Loop
    End If
    'oExcel.Columns("A:M").EntireColumn.AutoFit
    oExcel.Workbooks.Item(1).Save
    Screen.MousePointer = vbNormal
    If OutputTo = "S" Then oExcel.Range("A1").Activate: oExcel.Visible = True Else oExcel.Workbooks.Item(1).PrintOut
    Set oExcel = Nothing
    On Error GoTo 0
    Exit Sub
'ErrorHandler:
End Sub




