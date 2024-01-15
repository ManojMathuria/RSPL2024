VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmAccountChild07 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Laminator Rate Detail"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6315
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "AccountChild07.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "FrmLogin"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   6315
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Height          =   375
      Left            =   5845
      Picture         =   "AccountChild07.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Cancel"
      Top             =   465
      Width           =   375
   End
   Begin VB.CommandButton cmdProceed 
      Height          =   375
      Left            =   5845
      Picture         =   "AccountChild07.frx":0544
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Save"
      Top             =   105
      Width           =   375
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
      Height          =   3785
      Left            =   120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   105
      Width           =   5610
      _Version        =   65536
      _ExtentX        =   9895
      _ExtentY        =   6676
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
      Picture         =   "AccountChild07.frx":0646
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1560
         MaxLength       =   40
         TabIndex        =   1
         Top             =   730
         Width           =   3925
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   5
         Top             =   100
         Width           =   3925
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1560
         MaxLength       =   40
         TabIndex        =   0
         Top             =   425
         Width           =   3925
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
         Height          =   330
         Left            =   120
         TabIndex        =   7
         Top             =   420
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
         Caption         =   " Book Size"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "AccountChild07.frx":0662
         Picture         =   "AccountChild07.frx":067E
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
         Height          =   330
         Left            =   120
         TabIndex        =   8
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
         Caption         =   " Laminator Name"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "AccountChild07.frx":069A
         Picture         =   "AccountChild07.frx":06B6
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel9 
         Height          =   330
         Left            =   120
         TabIndex        =   9
         Top             =   735
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
         Caption         =   " Lamination Type"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "AccountChild07.frx":06D2
         Picture         =   "AccountChild07.frx":06EE
      End
      Begin FPSpreadADO.fpSpread fpSpread1 
         Height          =   2430
         Left            =   120
         TabIndex        =   2
         Top             =   1245
         Width           =   5370
         _Version        =   524288
         _ExtentX        =   9472
         _ExtentY        =   4286
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
         MaxCols         =   1
         MaxRows         =   8
         OperationMode   =   2
         SpreadDesigner  =   "AccountChild07.frx":070A
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   5640
         Y1              =   1140
         Y2              =   1140
      End
   End
End
Attribute VB_Name = "FrmAccountChild07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rstAccountChild As New ADODB.Recordset
Public rstSizeList As New ADODB.Recordset
Public rstLaminationTypeList As New ADODB.Recordset
Public AccountName As String
Dim SizeCode As String
Dim LaminationTypeCode As String
Dim EditMode As Boolean
Private Sub Form_Load()
    CenterForm Me
    Text2.Text = Trim(AccountName)
    ClearFields
    LoadFields
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
    Set rstAccountChild = Nothing
    Set rstSizeList = Nothing
    Set rstLaminationTypeList = Nothing
End Sub
Private Sub ClearFields()
    Text3.Text = ""
    Text1.Text = ""
    fpSpread1.ClearRange 1, 1, fpSpread1.MaxCols, fpSpread1.MaxRows, True
End Sub
Private Sub LoadFields()
    Dim Cnt As Integer
    
    If rstAccountChild.RecordCount = 0 Then Exit Sub
    If Not CheckEmpty(rstAccountChild.Fields("Size").Value, False) Then
        Text3.Text = rstAccountChild.Fields("SizeName").Value
        Text1.Text = rstAccountChild.Fields("LaminationTypeName").Value
        For Cnt = 1 To 8
            fpSpread1.SetText 1, Cnt, Val(rstAccountChild.Fields("Rate" & IIf(Cnt = 1, "04", IIf(Cnt = 2, "06", IIf(Cnt = 3, "08", IIf(Cnt = 4, "12", IIf(Cnt = 5, "16", IIf(Cnt = 6, "24", IIf(Cnt = 7, "32", "64")))))))).Value)
        Next
    End If
End Sub
Private Sub SaveFields()
    Dim Rate As Variant, Cnt As Integer
    
    rstAccountChild.Fields("Size").Value = SizeCode
    rstAccountChild.Fields("LaminationType").Value = LaminationTypeCode
    rstAccountChild.Fields("SizeName").Value = Trim(Text3.Text)
    rstAccountChild.Fields("LaminationTypeName").Value = Trim(Text1.Text)
    For Cnt = 1 To 8
        fpSpread1.GetText 1, Cnt, Rate
        rstAccountChild.Fields("Rate" & IIf(Cnt = 1, "04", IIf(Cnt = 2, "06", IIf(Cnt = 3, "08", IIf(Cnt = 4, "12", IIf(Cnt = 5, "16", IIf(Cnt = 6, "24", IIf(Cnt = 7, "32", "64")))))))).Value = Val(Rate)
    Next
End Sub
Private Sub Text3_Change()
    If Text3.Text = " " Then
        Text3.Text = "?"
        SendKeys "{TAB}"
    End If
End Sub
Private Sub Text3_Validate(Cancel As Boolean)
    Dim SearchString As String
    
    SearchString = FixQuote(Text3.Text)
    If rstSizeList.RecordCount = 0 Then
       DisplayError ("No Record in Size Master")
       Cancel = True
       Exit Sub
    Else
       rstSizeList.MoveFirst
    End If
    rstSizeList.Find "[Col0] = '" & RTrim(SearchString) & "'"
    If rstSizeList.EOF Then
        SelectionType = "S"
        SizeCode = ""
        Call LoadSelectionList(rstSizeList, "List of Sizes...", "Name")
        SearchOrder = 0
        Call DisplaySelectionList(Text3, SizeCode)
        Call CloseForm(FrmSelectionList)
        If CheckEmpty(Text3.Text, False) Then
            Text3.Text = "?"
        End If
        If RTrim(SizeCode) <> "" Then
            SendKeys "{TAB}"
        End If
        Cancel = True
        Exit Sub
    ElseIf (rstAccountChild.Fields("SizeName").Value <> Trim(Text3.Text)) Or (CheckEmpty(rstAccountChild.Fields("SizeName").Value, False)) Then
        If CheckDuplicateEntry Then
            Call DisplayError("Duplicate Entry")
            Text3.SelStart = 0
            Text3.SelLength = Len(Text3.Text)
            Cancel = True
            Exit Sub
        End If
    End If
    SizeCode = rstSizeList.Fields("Code").Value
End Sub
Private Sub Text1_Change()
    If Text1.Text = " " Then
        Text1.Text = "?"
        SendKeys "{TAB}"
    End If
End Sub
Private Sub Text1_Validate(Cancel As Boolean)
    Dim SearchString As String
    
    SearchString = FixQuote(Text1.Text)
    If rstLaminationTypeList.RecordCount = 0 Then
       DisplayError ("No Record in Lamination Type Master")
       Cancel = True
       Exit Sub
    Else
       rstLaminationTypeList.MoveFirst
    End If
    rstLaminationTypeList.Find "[Col0] = '" & RTrim(SearchString) & "'"
    If rstLaminationTypeList.EOF Then
        SelectionType = "S"
        LaminationTypeCode = ""
        Call LoadSelectionList(rstLaminationTypeList, "List of Lamination Types...", "Name")
        SearchOrder = 0
        Call DisplaySelectionList(Text1, LaminationTypeCode)
        Call CloseForm(FrmSelectionList)
        If CheckEmpty(Text1.Text, False) Then
            Text1.Text = "?"
        End If
        If RTrim(LaminationTypeCode) <> "" Then
            SendKeys "{TAB}"
        End If
        Cancel = True
        Exit Sub
    ElseIf (rstAccountChild.Fields("LaminationTypeName").Value <> Trim(Text1.Text)) Or (CheckEmpty(rstAccountChild.Fields("LaminationTypeName").Value, False)) Then
        If CheckDuplicateEntry Then
            Call DisplayError("Duplicate Entry")
            Text1.SelStart = 0
            Text1.SelLength = Len(Text1.Text)
            Cancel = True
            Exit Sub
        End If
    End If
    LaminationTypeCode = rstLaminationTypeList.Fields("Code").Value
End Sub
Private Sub cmdProceed_Click()
    Dim Control As Object
    
    If CheckMandatoryFields Then Exit Sub
    SaveFields
    Me.Tag = "T"
    If Val(rstAccountChild.Fields("Rate04").Value) + Val(rstAccountChild.Fields("Rate06").Value) + Val(rstAccountChild.Fields("Rate08").Value) + Val(rstAccountChild.Fields("Rate12").Value) + Val(rstAccountChild.Fields("Rate16").Value) + Val(rstAccountChild.Fields("Rate24").Value) + Val(rstAccountChild.Fields("Rate32").Value) + Val(rstAccountChild.Fields("Rate64").Value) = 0 Then
        rstAccountChild.Fields("Size").Value = ""
    End If
    rstAccountChild.Update
    Call CloseForm(Me)
End Sub
Private Sub cmdCancel_Click()
    Call CloseForm(Me)
End Sub
Private Function CheckMandatoryFields() As Boolean
    If CheckEmpty(Text3.Text, False) Then
        Text3.SetFocus
        CheckMandatoryFields = True
    ElseIf Not CheckExists(Text3, "Col0", rstSizeList, SizeCode) Then
        Text3.SetFocus
        CheckMandatoryFields = True
    ElseIf CheckEmpty(Text1.Text, False) Then
        Text1.SetFocus
        CheckMandatoryFields = True
    ElseIf Not CheckExists(Text1, "Col0", rstLaminationTypeList, LaminationTypeCode) Then
        Text1.SetFocus
        CheckMandatoryFields = True
    End If
End Function
Private Function CheckDuplicateEntry() As Boolean
    Dim dblBookMark As Double
    
    If rstAccountChild.RecordCount = 0 Then Exit Function
    If Not (rstAccountChild.EOF Or rstAccountChild.BOF) Then
        dblBookMark = rstAccountChild.Bookmark
    End If
    rstAccountChild.MoveFirst
    Do While Not rstAccountChild.EOF
        If rstAccountChild.Fields("SizeName").Value = Trim(Text3.Text) And rstAccountChild.Fields("LaminationTypeName").Value = Trim(Text1.Text) Then
            CheckDuplicateEntry = True
            Exit Do
        End If
        rstAccountChild.MoveNext
    Loop
    If dblBookMark <> 0 Then
        rstAccountChild.Bookmark = dblBookMark
    Else
        rstAccountChild.MoveLast
    End If
End Function
Private Sub fpSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    EditMode = IIf(Mode = 1, True, False)
End Sub
