VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form FrmBookList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "List of Books [All]"
   ClientHeight    =   8070
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   9600
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "BookList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8070
   ScaleWidth      =   9600
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9600
      _ExtentX        =   16933
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
            Object.ToolTipText     =   "Exit"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   4
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "New"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Revised"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Old"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "All"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
      Begin VB.CheckBox Check 
         Caption         =   "All (Including Received)"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   230
         Left            =   7420
         TabIndex        =   4
         Top             =   60
         Visible         =   0   'False
         Width           =   2265
      End
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
            Picture         =   "BookList.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BookList.frx":0986
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BookList.frx":0A9A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   7710
      Left            =   15
      TabIndex        =   3
      Top             =   345
      Width           =   9570
      _Version        =   65536
      _ExtentX        =   16880
      _ExtentY        =   13600
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
      Picture         =   "BookList.frx":0BAE
      Begin MSComctlLib.ListView ListView1 
         Height          =   7710
         Left            =   0
         TabIndex        =   0
         Top             =   0
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   13600
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
      Begin MSComctlLib.ListView ListView2 
         Height          =   7710
         Left            =   3755
         TabIndex        =   1
         Top             =   0
         Width           =   5820
         _ExtentX        =   10266
         _ExtentY        =   13600
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
   End
End
Attribute VB_Name = "FrmBookList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstCompanyMaster As New ADODB.Recordset, rstBookList As New ADODB.Recordset, rstBoardList As New ADODB.Recordset
Dim BookType As String
Dim OutputTo As String
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    BookType = "A": Check.Value = 1
    CenterForm Me
    BusySystemIndicator True
    rstCompanyMaster.Open "SELECT PrintName FROM CompanyMaster", CxnDatabase, adOpenKeyset, adLockReadOnly
    rstBoardList.Open "SELECT Name,Code FROM GeneralMaster WHERE Type='2' ORDER BY Name", CxnDatabase, adOpenKeyset, adLockReadOnly
    rstBoardList.ActiveConnection = Nothing
    Call FillList(ListView1, "List of Boards...", rstBoardList)
    Call BookSelection(True)
    BusySystemIndicator False
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    Call CloseForm(FrmBookList)
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
    If UnloadMode = 0 Then Call CloseForm(FrmBookList)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CloseRecordset(rstCompanyMaster)
    Call CloseRecordset(rstBoardList)
    Call CloseRecordset(rstBookList)
End Sub
Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
     Call BookSelection(False)
End Sub
Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    If Shift = vbCtrlMask And (KeyCode = vbKeyA Or KeyCode = vbKeyD) Then
        For i = 1 To ListView1.ListItems.Count
            ListView1.ListItems(i).Checked = IIf(KeyCode = vbKeyA, True, False)
        Next i
        Call BookSelection(IIf(KeyCode = vbKeyA, True, False))
    End If
End Sub
Private Sub ListView2_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    If Shift = vbCtrlMask And (KeyCode = vbKeyA Or KeyCode = vbKeyD) Then
        For i = 1 To ListView2.ListItems.Count
            ListView2.ListItems(i).Checked = IIf(KeyCode = vbKeyA, True, False)
        Next i
    End If
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    If Button.Index = 3 Then CloseForm Me: Exit Sub
    If Button.Index = 1 Then OutputTo = "S" Else OutputTo = "P"
    PrintBookList
End Sub
Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    On Error Resume Next
    BookType = Choose(ButtonMenu.Index, "N", "R", "O", "A")
    If BookType = "A" Then Check.Visible = False: Check.Value = 1 Else Check.Visible = True: Check.Value = 0
    Me.Caption = "List of Books [" & Choose(ButtonMenu.Index, "New", "Revised", "Old", "All") & "]"
End Sub
Private Sub BookSelection(ByVal SelectAll As Boolean)
    If rstBookList.State = adStateOpen Then rstBookList.Close
    rstBookList.Open "SELECT Name,Code FROM BookMaster " & IIf(SelectAll, "", "WHERE Board IN (" & SelectedItems(ListView1) & ")") & " ORDER BY Name", CxnDatabase, adOpenKeyset, adLockReadOnly
    rstBookList.ActiveConnection = Nothing
    ListView2.ListItems.Clear
    Call FillList(ListView2, "List of Books...", rstBookList)
End Sub
Private Sub PrintBookList()
    Dim oExcel As Object
    Dim i As Long, Cnt As Long
    Dim SelectedBoards, SelectedBooks
    On Error Resume Next
    If Not FileExist(App.Path & "\Template\Print Order Status Register.xlsx") Then Exit Sub
    Screen.MousePointer = vbHourglass
    If rstBookList.State = adStateOpen Then rstBookList.Close
    MdiMainMenu.StatusBar1.Panels(2).Text = "Processing !!! Please Wait....."
    SelectedBoards = SelectedItems(ListView1): SelectedBooks = SelectedItems(ListView2)
   
    Dim strQry As String
    
    strQry = "SELECT M1.PrintName As BookName,M1.BusyCode As Alias,M1.Price,M1.ISBN,M1.FormType,M2.PrintName As SizeName,M1.OneColorForms,M1.TwoColorForms,M1.FourColorForms,M1.TitleFrontColor,M1.TitleBackColor,(SELECT PrintName FROM AccountMaster WHERE Code=M1.BookPrinter) As BookPrinter,(SELECT PrintName FROM AccountMaster WHERE Code=M1.TitlePrinter) As TitlePrinter,(SELECT PrintName FROM AccountMaster WHERE Code=M1.Laminator) As Laminator,(SELECT PrintName FROM AccountMaster WHERE Code=M1.BinderFresh) As Binder," & _
                                 "(SELECT TOP 1 PrintName FROM (BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN PaperMaster M ON C.Paper1=M.Code WHERE P.Type='F' AND P.Book=M1.Code AND " & IIf(BookType = "A", "1=1", "C.Processing='" & BookType & "'") & " AND " & IIf(Check.Value, "1=1", "ReceivedQuantity=0") & " ORDER BY FORMAT(Date,'YYYYMMDD')+P.Code DESC) As Paper1," & _
                                 "(SELECT TOP 1 PrintName FROM (BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN PaperMaster M ON C.Paper2=M.Code WHERE P.Type='F' AND P.Book=M1.Code AND " & IIf(BookType = "A", "1=1", "C.Processing='" & BookType & "'") & " AND " & IIf(Check.Value, "1=1", "ReceivedQuantity=0") & " ORDER BY FORMAT(Date,'YYYYMMDD')+P.Code DESC) As Paper2," & _
                                 "(SELECT TOP 1 PrintName FROM (BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN PaperMaster M ON C.Paper4=M.Code WHERE P.Type='F' AND P.Book=M1.Code AND " & IIf(BookType = "A", "1=1", "C.Processing='" & BookType & "'") & " AND " & IIf(Check.Value, "1=1", "ReceivedQuantity=0") & " ORDER BY FORMAT(Date,'YYYYMMDD')+P.Code DESC) As Paper4," & _
                                 "(SELECT TOP 1 PrintName FROM (OBookPOParent P INNER JOIN OBookPOChild05 C ON P.Code=C.Code) INNER JOIN PaperMaster M ON C.Paper1=M.Code WHERE P.Type='F' AND P.Book=M1.Code AND " & IIf(BookType = "A", "1=1", "C.Processing='" & BookType & "'") & " AND " & IIf(Check.Value, "1=1", "ReceivedQuantity=0") & " ORDER BY FORMAT(Date,'YYYYMMDD')+P.Code DESC) As OPaper1," & _
                                 "(SELECT TOP 1 PrintName FROM (OBookPOParent P INNER JOIN OBookPOChild05 C ON P.Code=C.Code) INNER JOIN PaperMaster M ON C.Paper2=M.Code WHERE P.Type='F' AND P.Book=M1.Code AND " & IIf(BookType = "A", "1=1", "C.Processing='" & BookType & "'") & " AND " & IIf(Check.Value, "1=1", "ReceivedQuantity=0") & " ORDER BY FORMAT(Date,'YYYYMMDD')+P.Code DESC) As OPaper2," & _
                                 "(SELECT TOP 1 PrintName FROM (OBookPOParent P INNER JOIN OBookPOChild05 C ON P.Code=C.Code) INNER JOIN PaperMaster M ON C.Paper4=M.Code WHERE P.Type='F' AND P.Book=M1.Code AND " & IIf(BookType = "A", "1=1", "C.Processing='" & BookType & "'") & " AND " & IIf(Check.Value, "1=1", "ReceivedQuantity=0") & " ORDER BY FORMAT(Date,'YYYYMMDD')+P.Code DESC) As OPaper4," & _
                                 "(SELECT TOP 1 Date FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE P.Type='F' AND P.Book=M1.Code AND " & IIf(BookType = "A", "1=1", "C.Processing='" & BookType & "'") & " AND " & IIf(Check.Value, "1=1", "ReceivedQuantity=0") & " ORDER BY FORMAT(Date,'YYYYMMDD')+P.Code DESC) As LastPODate05," & _
                                 "(SELECT TOP 1 Date FROM OBookPOParent P INNER JOIN OBookPOChild05 C ON P.Code=C.Code WHERE P.Type='F' AND P.Book=M1.Code AND " & IIf(BookType = "A", "1=1", "C.Processing='" & BookType & "'") & " AND " & IIf(Check.Value, "1=1", "ReceivedQuantity=0") & " ORDER BY FORMAT(Date,'YYYYMMDD')+P.Code DESC) As OLastPODate05," & _
                                 "(SELECT TOP 1 Date FROM BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code WHERE P.Type='F' AND P.Book=M1.Code AND " & IIf(Check.Value, "1=1", "ReceivedQuantity=0") & " ORDER BY FORMAT(Date,'YYYYMMDD')+P.Code DESC) As LastPODate08," & _
                                 "(SELECT TOP 1 Date FROM OBookPOParent P INNER JOIN OBookPOChild08 C ON P.Code=C.Code WHERE P.Type='F' AND P.Book=M1.Code AND " & IIf(Check.Value, "1=1", "ReceivedQuantity=0") & " ORDER BY FORMAT(Date,'YYYYMMDD')+P.Code DESC) As OLastPODate08 " & _
                                 "FROM BookMaster M1 INNER JOIN GeneralMaster M2 ON M1.[Size]=M2.Code WHERE M1.Type='F' AND " & IIf(SelectedBoards = "''", "1=1", "M1.Board IN (" & SelectedBoards & ")") & " AND " & IIf(SelectedBooks = "''", "1=1", "M1.Code IN (" & SelectedBooks & ")") & " AND " & IIf(BookType = "A", "1=1", "(M1.Code IN (SELECT Book FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE P.Book=M1.Code AND P.Type='F' AND " & IIf(Check.Value, "1=1", "P.ReceivedQuantity = 0") & " AND C.Processing='" & BookType & "')  OR M1.Code IN (SELECT Book FROM OBookPOParent P INNER JOIN OBookPOChild05 C ON P.Code=C.Code WHERE P.Book=M1.Code AND P.Type='F' AND " & IIf(Check.Value, "1=1", "P.ReceivedQuantity = 0") & " AND C.Processing='" & BookType & "'))") & " ORDER BY M1.PrintName"

    
    rstBookList.Open "SELECT M1.PrintName As BookName,M1.BusyCode As Alias,M1.Price,M1.ISBN,M1.FormType,M2.PrintName As SizeName,M1.OneColorForms,M1.TwoColorForms,M1.FourColorForms,M1.TitleFrontColor,M1.TitleBackColor,(SELECT PrintName FROM AccountMaster WHERE Code=M1.BookPrinter) As BookPrinter,(SELECT PrintName FROM AccountMaster WHERE Code=M1.TitlePrinter) As TitlePrinter,(SELECT PrintName FROM AccountMaster WHERE Code=M1.Laminator) As Laminator,(SELECT PrintName FROM AccountMaster WHERE Code=M1.BinderFresh) As Binder," & _
                                 "(SELECT TOP 1 PrintName FROM (BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN PaperMaster M ON C.Paper1=M.Code WHERE P.Type='F' AND P.Book=M1.Code AND " & IIf(BookType = "A", "1=1", "C.Processing='" & BookType & "'") & " AND " & IIf(Check.Value, "1=1", "ReceivedQuantity=0") & " ORDER BY FORMAT(Date,'YYYYMMDD')+P.Code DESC) As Paper1," & _
                                 "(SELECT TOP 1 PrintName FROM (BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN PaperMaster M ON C.Paper2=M.Code WHERE P.Type='F' AND P.Book=M1.Code AND " & IIf(BookType = "A", "1=1", "C.Processing='" & BookType & "'") & " AND " & IIf(Check.Value, "1=1", "ReceivedQuantity=0") & " ORDER BY FORMAT(Date,'YYYYMMDD')+P.Code DESC) As Paper2," & _
                                 "(SELECT TOP 1 PrintName FROM (BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code) INNER JOIN PaperMaster M ON C.Paper4=M.Code WHERE P.Type='F' AND P.Book=M1.Code AND " & IIf(BookType = "A", "1=1", "C.Processing='" & BookType & "'") & " AND " & IIf(Check.Value, "1=1", "ReceivedQuantity=0") & " ORDER BY FORMAT(Date,'YYYYMMDD')+P.Code DESC) As Paper4," & _
                                 "(SELECT TOP 1 PrintName FROM (OBookPOParent P INNER JOIN OBookPOChild05 C ON P.Code=C.Code) INNER JOIN PaperMaster M ON C.Paper1=M.Code WHERE P.Type='F' AND P.Book=M1.Code AND " & IIf(BookType = "A", "1=1", "C.Processing='" & BookType & "'") & " AND " & IIf(Check.Value, "1=1", "ReceivedQuantity=0") & " ORDER BY FORMAT(Date,'YYYYMMDD')+P.Code DESC) As OPaper1," & _
                                 "(SELECT TOP 1 PrintName FROM (OBookPOParent P INNER JOIN OBookPOChild05 C ON P.Code=C.Code) INNER JOIN PaperMaster M ON C.Paper2=M.Code WHERE P.Type='F' AND P.Book=M1.Code AND " & IIf(BookType = "A", "1=1", "C.Processing='" & BookType & "'") & " AND " & IIf(Check.Value, "1=1", "ReceivedQuantity=0") & " ORDER BY FORMAT(Date,'YYYYMMDD')+P.Code DESC) As OPaper2," & _
                                 "(SELECT TOP 1 PrintName FROM (OBookPOParent P INNER JOIN OBookPOChild05 C ON P.Code=C.Code) INNER JOIN PaperMaster M ON C.Paper4=M.Code WHERE P.Type='F' AND P.Book=M1.Code AND " & IIf(BookType = "A", "1=1", "C.Processing='" & BookType & "'") & " AND " & IIf(Check.Value, "1=1", "ReceivedQuantity=0") & " ORDER BY FORMAT(Date,'YYYYMMDD')+P.Code DESC) As OPaper4," & _
                                 "(SELECT TOP 1 Date FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE P.Type='F' AND P.Book=M1.Code AND " & IIf(BookType = "A", "1=1", "C.Processing='" & BookType & "'") & " AND " & IIf(Check.Value, "1=1", "ReceivedQuantity=0") & " ORDER BY FORMAT(Date,'YYYYMMDD')+P.Code DESC) As LastPODate05," & _
                                 "(SELECT TOP 1 Date FROM OBookPOParent P INNER JOIN OBookPOChild05 C ON P.Code=C.Code WHERE P.Type='F' AND P.Book=M1.Code AND " & IIf(BookType = "A", "1=1", "C.Processing='" & BookType & "'") & " AND " & IIf(Check.Value, "1=1", "ReceivedQuantity=0") & " ORDER BY FORMAT(Date,'YYYYMMDD')+P.Code DESC) As OLastPODate05," & _
                                 "(SELECT TOP 1 Date FROM BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code WHERE P.Type='F' AND P.Book=M1.Code AND " & IIf(Check.Value, "1=1", "ReceivedQuantity=0") & " ORDER BY FORMAT(Date,'YYYYMMDD')+P.Code DESC) As LastPODate08," & _
                                 "(SELECT TOP 1 Date FROM OBookPOParent P INNER JOIN OBookPOChild08 C ON P.Code=C.Code WHERE P.Type='F' AND P.Book=M1.Code AND " & IIf(Check.Value, "1=1", "ReceivedQuantity=0") & " ORDER BY FORMAT(Date,'YYYYMMDD')+P.Code DESC) As OLastPODate08 " & _
                                 "FROM BookMaster M1 INNER JOIN GeneralMaster M2 ON M1.[Size]=M2.Code WHERE M1.Type='F' AND " & IIf(SelectedBoards = "''", "1=1", "M1.Board IN (" & SelectedBoards & ")") & " AND " & IIf(SelectedBooks = "''", "1=1", "M1.Code IN (" & SelectedBooks & ")") & " AND " & IIf(BookType = "A", "1=1", "(M1.Code IN (SELECT Book FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE P.Book=M1.Code AND P.Type='F' AND " & IIf(Check.Value, "1=1", "P.ReceivedQuantity = 0") & " AND C.Processing='" & BookType & "')  OR M1.Code IN (SELECT Book FROM OBookPOParent P INNER JOIN OBookPOChild05 C ON P.Code=C.Code WHERE P.Book=M1.Code AND P.Type='F' AND " & IIf(Check.Value, "1=1", "P.ReceivedQuantity = 0") & " AND C.Processing='" & BookType & "'))") & " ORDER BY M1.PrintName", CxnDatabase, adOpenKeyset, adLockOptimistic



    If rstBookList.RecordCount = 0 Then Screen.MousePointer = vbNormal: On Error GoTo 0: Exit Sub
    DoEvents
    Set oExcel = CreateObject("Excel.Application")
    oExcel.Workbooks.Open (App.Path & "\Template\Book List")
    oExcel.DisplayAlerts = False
    oExcel.Workbooks.Item(1).SaveAs (App.Path & "\Report\Book List (" & CompCode & ")")
    oExcel.DisplayAlerts = True
    oExcel.Visible = False
    oExcel.Cells(1, "A").Value = Trim(rstCompanyMaster.Fields("PrintName").Value)
    oExcel.Cells(2, "A").Value = "List of Books" & IIf(BookType = "N", " (New)", IIf(BookType = "R", " (Revised)", IIf(BookType = "O", " (Old)", "")))
    i = 5: Cnt = 1
    Do While Not rstBookList.EOF
        
        oExcel.Cells(i, "A").Value = Cnt
        oExcel.Application.Cells(i, "B").Value = Trim(rstBookList.Fields("BookName").Value)
        oExcel.Application.Cells(i, "C").Value = Trim(rstBookList.Fields("Alias").Value)
        oExcel.Application.Cells(i, "D").Value = Val(rstBookList.Fields("Price").Value)
        oExcel.Application.Cells(i, "E").Value = Trim(rstBookList.Fields("ISBN").Value)
        oExcel.Application.Cells(i, "F").Value = Trim(rstBookList.Fields("SizeName").Value)
        oExcel.Application.Cells(i, "G").Value = IIf(rstBookList.Fields("FormType").Value = "1", "08", IIf(rstBookList.Fields("FormType").Value = "2", "16", IIf(rstBookList.Fields("FormType").Value = "3", "04", IIf(rstBookList.Fields("FormType").Value = "4", "12", IIf(rstBookList.Fields("FormType").Value = "5", "24", IIf(rstBookList.Fields("FormType").Value = "6", "32", "64"))))))
        If Val(CheckNull(rstBookList.Fields("OneColorForms").Value)) <> 0 Then oExcel.Application.Cells(i, "H").Value = Format(Val(rstBookList.Fields("OneColorForms").Value), "0.00")
        If Val(CheckNull(rstBookList.Fields("TwoColorForms").Value)) <> 0 Then oExcel.Application.Cells(i, "I").Value = Format(Val(rstBookList.Fields("TwoColorForms").Value), "0.00")
        If Val(CheckNull(rstBookList.Fields("FourColorForms").Value)) <> 0 Then oExcel.Application.Cells(i, "J").Value = Format(Val(rstBookList.Fields("FourColorForms").Value), "0.00")
        If Val(CheckNull(rstBookList.Fields("TitleFrontColor").Value)) <> 0 Then oExcel.Application.Cells(i, "K").Value = Trim(rstBookList.Fields("TitleFrontColor").Value)
        If Val(CheckNull(rstBookList.Fields("TitleBackColor").Value)) <> 0 Then oExcel.Application.Cells(i, "K").Value = oExcel.Application.Cells(i, "K").Value & "+" & Trim(rstBookList.Fields("TitleBackColor").Value)
        oExcel.Application.Cells(i, "L").Value = Trim(CheckNull(rstBookList.Fields("BookPrinter").Value))
        oExcel.Application.Cells(i, "M").Value = Trim(CheckNull(rstBookList.Fields("TitlePrinter").Value))
        oExcel.Application.Cells(i, "N").Value = Trim(CheckNull(rstBookList.Fields("Laminator").Value))
        oExcel.Application.Cells(i, "O").Value = Trim(CheckNull(rstBookList.Fields("Binder").Value))
        
        If Not CheckEmpty(rstBookList.Fields("Paper1").Value, False) Then
            oExcel.Application.Cells(i, "U").Value = rstBookList.Fields("Paper1").Value
        ElseIf Not CheckEmpty(rstBookList.Fields("Paper2").Value, False) Then
            oExcel.Application.Cells(i, "U").Value = rstBookList.Fields("Paper2").Value
        ElseIf Not CheckEmpty(rstBookList.Fields("Paper4").Value, False) Then
            oExcel.Application.Cells(i, "U").Value = rstBookList.Fields("Paper4").Value
        ElseIf Not CheckEmpty(rstBookList.Fields("OPaper1").Value, False) Then
            oExcel.Application.Cells(i, "U").Value = rstBookList.Fields("OPaper1").Value
        ElseIf Not CheckEmpty(rstBookList.Fields("OPaper2").Value, False) Then
            oExcel.Application.Cells(i, "U").Value = rstBookList.Fields("OPaper2").Value
        ElseIf Not CheckEmpty(rstBookList.Fields("OPaper4").Value, False) Then
            oExcel.Application.Cells(i, "U").Value = rstBookList.Fields("OPaper4").Value
        End If
        
                
        If Not CheckEmpty(rstBookList.Fields("LastPODate05").Value, False) Then
            oExcel.Application.Cells(i, "W").Value = Format(rstBookList.Fields("LastPODate05").Value, "dd-MM-yyyy")
        ElseIf Not CheckEmpty(rstBookList.Fields("OLastPODate05").Value, False) Then
            oExcel.Application.Cells(i, "W").Value = Format(rstBookList.Fields("OLastPODate05").Value, "dd-MM-yyyy")
        ElseIf Not CheckEmpty(rstBookList.Fields("LastPODate08").Value, False) Then
            oExcel.Application.Cells(i, "W").Value = Format(rstBookList.Fields("LastPODate08").Value, "dd-MM-yyyy")
        ElseIf Not CheckEmpty(rstBookList.Fields("OLastPODate08").Value, False) Then
            oExcel.Application.Cells(i, "W").Value = Format(rstBookList.Fields("OLastPODate08").Value, "dd-MM-yyyy")
        End If
        
        Cnt = Cnt + 1: i = i + 1
        MdiMainMenu.StatusBar1.Panels(2).Text = "Processed record #" & Trim(str(Cnt)) & " of " & Trim(str(rstBookList.RecordCount)) & " !!!"
        rstBookList.MoveNext
   
    Loop
    oExcel.Range("Q5:T" & Trim(str(i - 1))).Formula = oExcel.Range("Q5:T5").Formula
    MdiMainMenu.StatusBar1.Panels(2).Text = ""
    'oExcel.Sheets("Book List").Activate
    oExcel.Columns("A:W").EntireColumn.AutoFit
    oExcel.Columns("C").Hidden = True
    oExcel.Workbooks.Item(1).Save
    Screen.MousePointer = vbNormal
    If OutputTo = "S" Then oExcel.Range("A1").Activate: oExcel.Visible = True Else oExcel.Workbooks.Item(1).PrintOut
    Set oExcel = Nothing
    On Error GoTo 0
End Sub
