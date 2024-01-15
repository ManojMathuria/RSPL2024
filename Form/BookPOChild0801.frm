VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmBookPOChild0801 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Material Consumption Details"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8280
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "BookPOChild0801.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "FrmLogin"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   8280
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Height          =   375
      Left            =   7815
      Picture         =   "BookPOChild0801.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Cancel"
      Top             =   465
      Width           =   375
   End
   Begin VB.CommandButton cmdProceed 
      Height          =   375
      Left            =   7815
      Picture         =   "BookPOChild0801.frx":0544
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Proceed"
      Top             =   105
      Width           =   375
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
      Height          =   4250
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   105
      Width           =   7575
      _Version        =   65536
      _ExtentX        =   13361
      _ExtentY        =   7497
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
      Picture         =   "BookPOChild0801.frx":0646
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc1"
         Height          =   330
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   60
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   105
         Width           =   5895
      End
      Begin FPSpreadADO.fpSpread fpSpread1 
         Height          =   3495
         Left            =   120
         TabIndex        =   0
         Top             =   645
         Width           =   7335
         _Version        =   524288
         _ExtentX        =   12938
         _ExtentY        =   6165
         _StockProps     =   64
         EditEnterAction =   5
         EditModePermanent=   -1  'True
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   4
         MaxRows         =   100
         OperationMode   =   2
         SpreadDesigner  =   "BookPOChild0801.frx":0662
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
         Height          =   330
         Left            =   120
         TabIndex        =   4
         Top             =   105
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
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
         Caption         =   " Book Name"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "BookPOChild0801.frx":0C6C
         Picture         =   "BookPOChild0801.frx":0C88
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   8300
         Y1              =   540
         Y2              =   540
      End
   End
End
Attribute VB_Name = "FrmBookPOChild0801"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public BinderCode As String
Public BookCode As String
Public OrderCode As Variant
Public BookQuantity As Long
Public rstBookPOChild0801 As New ADODB.Recordset
Dim rstOutsourceItemList As New ADODB.Recordset
Dim rstPaperList As New ADODB.Recordset
Dim rstFreshBookList As New ADODB.Recordset
Dim rstRepairBookList As New ADODB.Recordset
Dim OutsourceItem As String
Dim Paper As String
Dim FreshBook As String
Dim RepairBook As String
Dim Title As String
Dim EditMode As Boolean
Private Sub Form_Load()
    Dim i As Integer
    On Error GoTo ErrorHandler
    
    CenterForm Me
    BusySystemIndicator True
    DisableCloseButton Me
    AbortPO = False
    Text2.Text = Trim(FrmBookPrintOrder.Text3.Text)
    rstOutsourceItemList.Open "Select Name,'1'+Code As NCode From OutsourceItemMaster Order By Name", CxnDatabase, adOpenKeyset, adLockOptimistic
    rstPaperList.Open "Select Name,'2'+Code As NCode From PaperMaster Order By Name", CxnDatabase, adOpenKeyset, adLockOptimistic
    rstFreshBookList.Open "Select Name,Board,'3'+Code As NCode From BookMaster Where Type='F' Order By Name", CxnDatabase, adOpenKeyset, adLockOptimistic
    rstRepairBookList.Open "Select Name,'4'+Code As NCode From BookMaster Where Type='R' Order By Name", CxnDatabase, adOpenKeyset, adLockOptimistic
    rstOutsourceItemList.ActiveConnection = Nothing
    rstPaperList.ActiveConnection = Nothing
    rstFreshBookList.ActiveConnection = Nothing
    rstRepairBookList.ActiveConnection = Nothing
    Call RefreshDropDownList("A")
    With fpSpread1
        .Col = 4
        .ColHidden = True
        .Col = 5
        .ColHidden = True
        .ClearRange 1, 1, .MaxCols, .MaxRows, True
        If CheckNull(OrderCode) = "" Then
            If rstBookPOChild0801.State = adStateOpen Then
                rstBookPOChild0801.Close
            End If
            rstBookPOChild0801.Open "SELECT Category,Item,Quantity FROM BookChild01 WHERE Code='" & BookCode & "'", CxnDatabase, adOpenKeyset, adLockOptimistic
            rstBookPOChild0801.ActiveConnection = Nothing
        End If
        If rstBookPOChild0801.RecordCount > 0 Then
            rstBookPOChild0801.MoveFirst
            i = 0
            Do While Not rstBookPOChild0801.EOF
                i = i + 1
                .SetText 1, i, IIf(rstBookPOChild0801.Fields("Category").Value = "1", "Outsource Item", IIf(rstBookPOChild0801.Fields("Category").Value = "2", "Paper", IIf(rstBookPOChild0801.Fields("Category").Value = "3", "Fresh Book", IIf(rstBookPOChild0801.Fields("Category").Value = "4", "Repair Book", "Title"))))
                .Col = 2
                .TypeComboBoxList = IIf(rstBookPOChild0801.Fields("Category").Value = "1", OutsourceItem, IIf(rstBookPOChild0801.Fields("Category").Value = "2", Paper, IIf(rstBookPOChild0801.Fields("Category").Value = "4", RepairBook, IIf(rstBookPOChild0801.Fields("Category").Value = "3", FreshBook, Title))))
                If rstBookPOChild0801.Fields("Category").Value = "1" Then
                   If rstOutsourceItemList.RecordCount > 0 Then rstOutsourceItemList.MoveFirst
                   rstOutsourceItemList.Find "[NCode]='" & FixQuote(rstBookPOChild0801.Fields("Category").Value + rstBookPOChild0801.Fields("Item").Value) & "'"
                   If Not rstOutsourceItemList.EOF Then
                        fpSpread1.SetText 2, i, rstOutsourceItemList.Fields("Name").Value
                   End If
                ElseIf rstBookPOChild0801.Fields("Category").Value = "2" Then
                   If rstPaperList.RecordCount > 0 Then rstPaperList.MoveFirst
                   rstPaperList.Find "[NCode]='" & FixQuote(rstBookPOChild0801.Fields("Category").Value + rstBookPOChild0801.Fields("Item").Value) & "'"
                   If Not rstPaperList.EOF Then
                        fpSpread1.SetText 2, i, rstPaperList.Fields("Name").Value
                   End If
                ElseIf rstBookPOChild0801.Fields("Category").Value = "4" Then
                   If rstRepairBookList.RecordCount > 0 Then rstRepairBookList.MoveFirst
                   rstRepairBookList.Find "[NCode]='" & FixQuote(rstBookPOChild0801.Fields("Category").Value + rstBookPOChild0801.Fields("Item").Value) & "'"
                   If Not rstRepairBookList.EOF Then
                        fpSpread1.SetText 2, i, rstRepairBookList.Fields("Name").Value
                   End If
                Else
                   If rstFreshBookList.RecordCount > 0 Then rstFreshBookList.MoveFirst
                   rstFreshBookList.Find "[NCode]='" & FixQuote("3" + rstBookPOChild0801.Fields("Item").Value) & "'"
                   If Not rstFreshBookList.EOF Then
                        fpSpread1.SetText 2, i, rstFreshBookList.Fields("Name").Value
                   End If
                End If
                .SetText 3, i, Val(rstBookPOChild0801.Fields("Quantity").Value)
                .SetText 4, i, rstBookPOChild0801.Fields("Category").Value + rstBookPOChild0801.Fields("Item").Value
                rstBookPOChild0801.MoveNext
            Loop
        End If
    End With
    BusySystemIndicator False
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    Call CloseForm(Me)
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyReturn Then
        If Me.ActiveControl.Name <> "fpSpread1" Then
            SendKeys "{TAB}"
            KeyCode = 0
        End If
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyS Then
        If Not EditMode Then
            cmdProceed_Click
            KeyCode = 0
        End If
    ElseIf Shift = 0 And KeyCode = vbKeyEscape Then
        If Not EditMode Then
            cmdCancel_Click
            KeyCode = 0
        End If
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
       Call CloseForm(Me)
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CloseRecordset(rstOutsourceItemList)
    Call CloseRecordset(rstPaperList)
    Call CloseRecordset(rstFreshBookList)
    Call CloseRecordset(rstRepairBookList)
End Sub
Private Function CheckMandatoryFields() As Boolean
    If CheckItem() Then
       fpSpread1.SetFocus
       CheckMandatoryFields = True
    End If
End Function
Private Sub cmdProceed_Click()
    If CheckMandatoryFields Then Exit Sub
    SaveFields
    Call CloseForm(Me)
End Sub
Private Sub cmdCancel_Click()
    Call CloseForm(Me)
End Sub
Private Sub SaveFields()
    Dim i As Integer, Category As Variant, Item As Variant, Qty As Variant
    
    With rstBookPOChild0801
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            .Delete adAffectCurrent
            .MoveNext
        Loop
    End With
    With fpSpread1
        For i = 1 To .DataRowCnt
            .GetText 3, i, Qty
            If Val(Qty) > 0 Then
                .GetText 1, i, Category
                .GetText 4, i, Item
                With rstBookPOChild0801
                    .AddNew
                    .Fields("Category").Value = IIf(Category = "Outsource Item", "1", IIf(Category = "Paper", "2", IIf(Category = "Fresh Book", "3", IIf(Category = "Repair Book", "4", "5"))))
                    .Fields("Item").Value = Right(Item, 6)
                    .Fields("Quantity").Value = Qty
                    .Update
                End With
            End If
        Next
    End With
End Sub
Private Sub fpSpread1_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask And KeyCode = vbKeyD Then
        If UserLevel = "3" Then Call DisplayError("You don't have the rights to delete BOM Item"): Exit Sub
        If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") = vbYes Then
            fpSpread1.DeleteRows fpSpread1.ActiveRow, 1
            fpSpread1.SetFocus
        End If
    ElseIf Shift = 0 And KeyCode = vbKeyF5 Then
        Call RefreshDropDownList("R")
    End If
End Sub
Private Sub fpSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim ActiveCellVal As Variant, Category As Variant
    
    fpSpread1.GetText Col, Row, ActiveCellVal
    If ActiveCellVal = "" Then
        Cancel = True
        Exit Sub
    End If
    fpSpread1.GetText 1, Row, Category
    If Col = 1 Then
        fpSpread1.Col = 2
        fpSpread1.TypeComboBoxList = IIf(Category = "Outsource Item", OutsourceItem, IIf(Category = "Paper", Paper, IIf(Category = "Repair Book", RepairBook, IIf(Category = "Fresh Book", FreshBook, Title))))
    ElseIf Col = 2 Then
        If Category = "Outsource Item" Then
           If rstOutsourceItemList.RecordCount > 0 Then rstOutsourceItemList.MoveFirst
           rstOutsourceItemList.Find "[Name]='" & FixQuote(ActiveCellVal) & "'"
           If Not rstOutsourceItemList.EOF Then
                fpSpread1.SetText 4, Row, rstOutsourceItemList.Fields("NCode").Value
           End If
        ElseIf Category = "Paper" Then
           If rstPaperList.RecordCount > 0 Then rstPaperList.MoveFirst
           rstPaperList.Find "[Name]='" & FixQuote(ActiveCellVal) & "'"
           If Not rstPaperList.EOF Then
                fpSpread1.SetText 4, Row, rstPaperList.Fields("NCode").Value
           End If
        ElseIf Category = "Repair Book" Then
           If rstRepairBookList.RecordCount > 0 Then rstRepairBookList.MoveFirst
           rstRepairBookList.Find "[Name]='" & FixQuote(ActiveCellVal) & "'"
           If Not rstRepairBookList.EOF Then
                fpSpread1.SetText 4, Row, rstRepairBookList.Fields("NCode").Value
           End If
        Else
           If rstFreshBookList.RecordCount > 0 Then rstFreshBookList.MoveFirst
           rstFreshBookList.Find "[Name]='" & FixQuote(ActiveCellVal) & "'"
           If Not rstFreshBookList.EOF Then
                fpSpread1.SetText 4, Row, rstFreshBookList.Fields("NCode").Value
           End If
        End If
    End If
End Sub
Private Function CheckItem() As Boolean
    Dim i As Integer, Item As Variant, Category As Variant, Qty As Variant, BalanceQuantity As Double
    
    CheckItem = False
    For i = 1 To fpSpread1.DataRowCnt
        fpSpread1.SetActiveCell 1, i
        fpSpread1.GetText 4, i, Item
        fpSpread1.GetText 1, i, Category
        fpSpread1.GetText 3, i, Qty
        If Category = "Outsource Item" Then
            If Left(Item, 1) <> "1" Then
                CheckItem = True
            End If
        ElseIf Category = "Paper" Then
            If Left(Item, 1) <> "2" Then
                CheckItem = True
            End If
        ElseIf Category = "Repair Book" Then
            If Left(Item, 1) <> "4" Then
                CheckItem = True
            End If
        Else
            If Left(Item, 1) <> "3" And Left(Item, 1) <> "5" Then
                CheckItem = True
            End If
        End If
        If CheckItem Then
            DisplayError "Data mismatch in row #" & Trim(str(i))
            Exit For
        End If
        If Category = "Paper" Then
            BalanceQuantity = CalculatePaperBalance(BinderCode, Right(Item, 6), CheckNull(OrderCode), "BPOB")
        Else
            BalanceQuantity = CalculateMaterialBalance(BinderCode, Category, Right(Item, 6), CheckNull(OrderCode), "PO")
        End If
        If FrmBookPrintOrder.BookPOType <> "O" Then
            If Val(Qty) * BookQuantity > Round(Val(BalanceQuantity), 3) Then
                If UserLevel <= 2 Then
                    If MsgBox("Stock (" & Format(Val(BalanceQuantity - Val(Qty) * BookQuantity), "0.000") & ") of the Item in row #" & Trim(str(i)) & " is going negative ! Would you like to continue ?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Proceed !") = vbNo Then Exit Function
                Else
                    Call DisplayError("Cann't Save ! Stock (" & Format(Val(BalanceQuantity - Val(Qty) * BookQuantity), "0.000") & ") of the Item in row #" & Trim(str(i)) & " is going negative"): AbortPO = True: Exit Function
                End If
            End If
        End If
    Next
End Function
Private Sub fpSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    EditMode = IIf(Mode = 1, True, False)
End Sub
Private Sub RefreshDropDownList(ByVal xType As String)
    If xType = "R" Then
        rstOutsourceItemList.ActiveConnection = CxnDatabase
        Do While Not RefreshRecord(rstOutsourceItemList)
        Loop
        rstOutsourceItemList.ActiveConnection = Nothing
        rstPaperList.ActiveConnection = CxnDatabase
        Do While Not RefreshRecord(rstPaperList)
        Loop
        rstPaperList.ActiveConnection = Nothing
        rstFreshBookList.ActiveConnection = CxnDatabase
        Do While Not RefreshRecord(rstFreshBookList)
        Loop
        rstFreshBookList.ActiveConnection = Nothing
        rstRepairBookList.ActiveConnection = CxnDatabase
        Do While Not RefreshRecord(rstRepairBookList)
        Loop
        rstRepairBookList.ActiveConnection = Nothing
        OutsourceItem = "": Paper = "": FreshBook = "": RepairBook = "": Title = ""
    End If
    Do While Not rstOutsourceItemList.EOF
        If OutsourceItem = "" Then
            OutsourceItem = rstOutsourceItemList.Fields("Name").Value
        Else
            OutsourceItem = OutsourceItem + Chr$(9) + rstOutsourceItemList.Fields("Name").Value
        End If
        rstOutsourceItemList.MoveNext
    Loop
    Do While Not rstPaperList.EOF
        If Paper = "" Then
            Paper = rstPaperList.Fields("Name").Value
        Else
            Paper = Paper + Chr$(9) + rstPaperList.Fields("Name").Value
        End If
        rstPaperList.MoveNext
    Loop
    rstFreshBookList.Filter = "[Board]='000000'"
    Do While Not rstFreshBookList.EOF
        If FreshBook = "" Then
            FreshBook = rstFreshBookList.Fields("Name").Value
        Else
            FreshBook = FreshBook + Chr$(9) + rstFreshBookList.Fields("Name").Value
        End If
        rstFreshBookList.MoveNext
    Loop
    rstFreshBookList.Filter = adFilterNone
    Do While Not rstFreshBookList.EOF
        If Title = "" Then
            Title = rstFreshBookList.Fields("Name").Value
        Else
            Title = Title + Chr$(9) + rstFreshBookList.Fields("Name").Value
        End If
        rstFreshBookList.MoveNext
    Loop
    Do While Not rstRepairBookList.EOF
        If RepairBook = "" Then
            RepairBook = rstRepairBookList.Fields("Name").Value
        Else
            RepairBook = RepairBook + Chr$(9) + rstRepairBookList.Fields("Name").Value
        End If
        rstRepairBookList.MoveNext
    Loop
End Sub
