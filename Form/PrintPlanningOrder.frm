VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Begin VB.Form FrmPrintPlanningOrder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Planning Order"
   ClientHeight    =   5160
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6180
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "PrintPlanningOrder.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5160
   ScaleWidth      =   6180
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6180
      _ExtentX        =   10901
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
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
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PrintPlanningOrder.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PrintPlanningOrder.frx":0986
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PrintPlanningOrder.frx":0A9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PrintPlanningOrder.frx":0BAE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   4850
      Left            =   45
      TabIndex        =   2
      Top             =   345
      Width           =   6090
      _Version        =   65536
      _ExtentX        =   10742
      _ExtentY        =   8555
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
      Picture         =   "PrintPlanningOrder.frx":0CC0
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   385
         Left            =   840
         TabIndex        =   4
         Top             =   4425
         Width           =   5245
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4410
         Left            =   0
         TabIndex        =   0
         Top             =   0
         Width           =   6090
         _ExtentX        =   10742
         _ExtentY        =   7779
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
         Height          =   385
         Left            =   0
         TabIndex        =   3
         Top             =   4425
         Width           =   855
         _Version        =   65536
         _ExtentX        =   1508
         _ExtentY        =   679
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
         Caption         =   " E-Mail :"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   -2147483634
         Picture         =   "PrintPlanningOrder.frx":0CDC
         Picture         =   "PrintPlanningOrder.frx":0CF8
      End
   End
End
Attribute VB_Name = "FrmPrintPlanningOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstCompanyMaster As New ADODB.Recordset
Dim rstPrintPlanningOrder As New ADODB.Recordset
Dim rstPrintList As New ADODB.Recordset
Dim rstPrintPlanningList As New ADODB.Recordset
Dim OutputTo As String
Public PlanningType As String
Dim oOutlook As New Outlook.Application
Dim Attachment As String
Dim Message As String

Private Sub Form_Activate()
  Text1.SetFocus
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandler
    CenterForm Me
    BusySystemIndicator True
    If PlanningType = "1" Then
        Me.Caption = "Print Planning Order [Book]"
    Else
        Me.Caption = "Print Planning Order [Title]"
    End If
    rstCompanyMaster.Open "Select PrintName From CompanyMaster", CxnDatabase, adOpenKeyset, adLockReadOnly
    rstPrintPlanningList.Open "Select '(' + trim(PrintPVParent.Name) + ')'+Particulars As Name,PrintPVParent.Code As Code From PrintPVParent Where PlanningType = '1' Order By PrintPVParent.Name", CxnDatabase, adOpenKeyset, adLockReadOnly
    rstPrintPlanningList.ActiveConnection = Nothing
    Call FillList(ListView1, "List of Memo...", rstPrintPlanningList)
    Call BookSelection(True)
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
    Call CloseRecordset(rstPrintList)
    Call CloseRecordset(rstPrintPlanningList)
    Call CloseRecordset(rstPrintPlanningOrder)
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

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    If Button.Index = 1 Then
        OutputTo = "S"
        PrintPrintPlanningRegister
    ElseIf Button.Index = 2 Then
        OutputTo = "P"
        PrintPrintPlanningRegister
    ElseIf Button.Index = 3 Then
        OutputTo = "M"
        PrintPrintPlanningRegister
    ElseIf Button.Index = 4 Then
        Unload Me
    End If
End Sub
Private Sub BookSelection(ByVal SelectAll As Boolean)
    If rstPrintList.State = adStateOpen Then
        rstPrintList.Close
    End If
    rstPrintList.Open "Select Name, Code From PrintPVParent " & IIf(SelectAll, "", "Where Code In (" & SelectedItems(ListView1) & ")") & " Order by Name", CxnDatabase, adOpenKeyset, adLockReadOnly
    rstPrintList.ActiveConnection = Nothing
  
End Sub
Private Sub PrintPrintPlanningRegister()
       
    Dim oOutlookMsg As Outlook.MailItem, RecordAffected As Integer
    Dim rstCompanyMaster As New ADODB.Recordset, Prefix As String
    Dim MemoNo As String
    
    Dim NePrint As String
        
    Dim Note As String
    Dim SelectedMemo As String
    Dim oExcel As Object
    Dim i As Long, Cnt As Long
    On Error GoTo ErrorHandler
    
    DoEvents
    
    rstCompanyMaster.Open "SELECT PrintName,Address1,Address2,Address3,Address4,Phone,Fax,eMail FROM CompanyMaster", CxnDatabase, adOpenKeyset, adLockReadOnly
    If rstPrintPlanningOrder.State = adStateOpen Then rstPrintPlanningOrder.Close
    SelectedMemo = SelectedItems(ListView1)
    rstPrintPlanningOrder.Open "Select Trim(PrintPVParent.Name) As VchNo,[Date] As VchDate,Trim(PrintName) As BookName,(Select Trim(PrintName) From GeneralMaster Where Code = BookMaster.Board) As BoardName,(Select Trim(PrintName) From GeneralMaster Where Code = BookMaster.[Size]) As SizeName,(Select Trim(PrintName) From AccountMaster Where Code=BookMaster.BookPrinter) As BookPrinter,(Select Trim(PrintName) From AccountMaster Where Code=BookMaster.TitlePrinter) As TitlePrinter,(Select Trim(PrintName) From AccountMaster Where Code=BookMaster.Laminator) As Laminator,(Select Trim(PrintName) From AccountMaster Where Code=BookMaster.BinderFresh) As Binder,Quantity,PrintPVChild.Forms,PrintPVChild.BookSize,[PaperWastage%],PaperConsumption,PrintPVChild.Narration,PrintPVChild.Warehouse1 As Noida,PrintPVChild.Warehouse2 As DistributorDaryaganj,PrintPVChild.Warehouse3 As 8No  " & _
                                   "From (PrintPVParent Inner Join PrintPVChild On (PrintPVParent.Code = PrintPVChild.Code And PlanningType = '" & PlanningType & "' And PrintPVParent.Code in(" & SelectedMemo & "))) Inner Join BookMaster On PrintPVChild.Book = BookMaster.Code Order By BookMaster.PrintName", CxnDatabase, adOpenKeyset, adLockOptimistic
    
    Screen.MousePointer = vbNormal
    
    On Error Resume Next
    If Not FileExist(App.Path & "\Template\Print Planning Order.xlsx") Then Exit Sub
    Screen.MousePointer = vbHourglass

    If rstPrintPlanningOrder.RecordCount = 0 Then
        DisplayError ("No Record Found")
        ShowProgressInStatusBar False
        Screen.MousePointer = vbNormal
        On Error GoTo 0
        Exit Sub
    End If
    DoEvents
  
    Set oExcel = CreateObject("Excel.Application")
    oExcel.Workbooks.Open (App.Path & "\Template\Print Planning Order"): oExcel.DisplayAlerts = False
    oExcel.Workbooks.Item(1).SaveAs (App.Path & "\Report\Print Planning Order (" & CompCode & ")"): oExcel.DisplayAlerts = True
       
    oExcel.Sheets("Print Planning Order (" & IIf(PlanningType = "1", "Book", "Title") & ")").Visible = False
    'oExcel.Sheets("Print Planning Order (" & IIf(PlanningType = "1", "Book", "Title") & ")").Select
    oExcel.Visible = False
    
    oExcel.Cells(1, "A").Value = rstCompanyMaster.Fields("PrintName").Value
    oExcel.Cells(2, "A").Value = "MEMO ORDER"
    MemoNo = Trim(rstPrintPlanningOrder.Fields("VchNo").Value)
    oExcel.Cells(3, "A").Value = "Memo No: " & Trim(rstPrintPlanningOrder.Fields("VchNo").Value)
    oExcel.Cells(3, "G").Value = "Date: " & Format(Trim(rstPrintPlanningOrder.Fields("VchDate").Value), "dd-MM-yyyy")
   
    i = 6: Cnt = 1
    
    Do While Not rstPrintPlanningOrder.EOF
        oExcel.Cells(i, "A").Value = Cnt
        oExcel.Application.Cells(i, "B").Value = Trim(rstPrintPlanningOrder.Fields("BookName").Value)
   
        oExcel.Application.Cells(i, "C").Value = Val(rstPrintPlanningOrder.Fields("Quantity").Value)
        
        If rstPrintPlanningOrder.Fields("BookPrinter").Value <> "" Then
           oExcel.Application.Cells(i, "D").Value = "Reprint"
        Else
           oExcel.Application.Cells(i, "D").Value = "New"
        End If
               
        oExcel.Application.Cells(i, "E").Value = Val(rstPrintPlanningOrder.Fields("Forms").Value)
        
        oExcel.Application.Cells(i, "F").Value = rstPrintPlanningOrder.Fields("BookSize").Value
        
        oExcel.Application.Cells(i, "G").Value = rstPrintPlanningOrder.Fields("BookPrinter").Value
        
        oExcel.Application.Cells(i, "H").Value = rstPrintPlanningOrder.Fields("Binder").Value
   
        oExcel.Application.Cells(i, "I").Value = Val(rstPrintPlanningOrder.Fields("Noida").Value)
        oExcel.Application.Cells(i, "J").Value = Val(rstPrintPlanningOrder.Fields("DistributorDaryaganj").Value)
        oExcel.Application.Cells(i, "K").Value = Val(rstPrintPlanningOrder.Fields("8No").Value)
        oExcel.Application.Cells(i, "L").Value = Val(rstPrintPlanningOrder.Fields("PaperConsumption").Value)
        oExcel.Application.Cells(i, "M").Value = rstPrintPlanningOrder.Fields("Narration").Value
        oExcel.Application.Cells(i, "N").Value = ""
        
        With oExcel.Worksheets("Sheet1").Rows(i)
         .RowHeight = .RowHeight * 2
        End With
        Cnt = Cnt + 1: i = i + 1
        rstPrintPlanningOrder.MoveNext
    Loop
    
    'oExcel.Cells(Cnt + 7, "B").Value = "Mr. Neeraj Gupta                             Mr. S C Gupta                          Mr. Anil Verma                             Mr. Tarun Sharma                              Mr. Saurabh Jain                           Mr. Ashwani"
    'oExcel.Range("B" & Cnt + 7 & ":" & "N" & Cnt + 7).Select
    'With oExcel.Selection
    '    .MergeCells = True
    'End With

    'With oExcel.Worksheets("Sheet1").Rows(Cnt + 7)
    ' .RowHeight = .RowHeight * 2
    'End With
    'oExcel.Columns("A:K").EntireColumn.AutoFit

    oExcel.Workbooks.Item(1).Save
    Screen.MousePointer = vbNormal
    MdiMainMenu.ProgressBar1.Value = 100
    
    If OutputTo = "S" Then
       oExcel.Range("A1").Activate: oExcel.Visible = True
    End If
    If OutputTo = "P" Then
       oExcel.Workbooks.Item(1).PrintOut
    End If
        
    If OutputTo = "M" Then
        If Text1.Text = "" Then: Exit Sub
        Attachment = Trim(rstPrintPlanningOrder.Fields("VchNo").Value)
        Attachment = Mid(Attachment, InStr(4, Attachment, "/") + 1)
        Message = "Dear Sir,<Br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Please find attached herewith Memo No : #" & MemoNo & " for doing the needful at your end. An early execution of the order will be highly appreciated.<Br>Kindly acknowledge the receipt of mail and confirm the date of execution of order.<Br><Br>" & IIf(Note = "", "", "<b><u>Note : " & Note & "</b></u><Br><Br>") & "Thanks & Regards<Br>Production Department<Br>" & Trim(rstCompanyMaster.Fields("PrintName").Value) & "<Br>Phone : " & Trim(rstCompanyMaster.Fields("Phone").Value) & "<Br>E-Mail : <a HRef='mailto:" & Trim(rstCompanyMaster.Fields("EMail").Value) & "'>" & Trim(rstCompanyMaster.Fields("EMail").Value) & "</a>"
        
        Set oOutlookMsg = oOutlook.CreateItem(olMailItem)
        With oOutlookMsg
            .To = Trim(Text1.Text)
            .Subject = IIf(PlanningType = "1", "Book", "Title") & " Memo No : #" & MemoNo
            .HTMLBody = "<Font Face='Calibri' Size='3'>" & Message & "</a>" & "</Font>"
            .Attachments.Add (App.Path & "\Report\" & "Print Planning Order (" & CompCode & ")" & ".xlsx")
            .Importance = olImportanceHigh
            .ReadReceiptRequested = True
            If CheckEmpty(.To, False) Then .Display Else .Send
        End With
        Set oOutlookMsg = Nothing
    End If
    ShowProgressInStatusBar False
    
    Set oExcel = Nothing
    On Error GoTo 0
    Exit Sub
    
ErrorHandler:
    Screen.MousePointer = vbNormal
    DisplayError ("Failed to print Planning order")
    ShowProgressInStatusBar False
    Call CloseRecordset(rstPrintPlanningOrder)
    Call CloseRecordset(rstCompanyMaster)
        
End Sub



