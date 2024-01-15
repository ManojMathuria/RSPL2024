VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb8.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmAccountChild08 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Binder Rate Detail"
   ClientHeight    =   4635
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
   Icon            =   "AccountChild08.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "FrmLogin"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   6315
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Height          =   375
      Left            =   5845
      Picture         =   "AccountChild08.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Cancel"
      Top             =   465
      Width           =   375
   End
   Begin VB.CommandButton cmdProceed 
      Height          =   375
      Left            =   5845
      Picture         =   "AccountChild08.frx":0544
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Save"
      Top             =   105
      Width           =   375
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
      Height          =   4440
      Left            =   120
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   105
      Width           =   5610
      _Version        =   65536
      _ExtentX        =   9895
      _ExtentY        =   7832
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
      Picture         =   "AccountChild08.frx":0646
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
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   7
         Top             =   100
         Width           =   3570
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
         Left            =   1920
         MaxLength       =   40
         TabIndex        =   0
         Top             =   425
         Width           =   3570
      End
      Begin VB.TextBox Text4 
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
         Left            =   1920
         MaxLength       =   40
         TabIndex        =   1
         Top             =   735
         Width           =   3570
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
         Height          =   330
         Left            =   120
         TabIndex        =   9
         Top             =   420
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
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
         Picture         =   "AccountChild08.frx":0662
         Picture         =   "AccountChild08.frx":067E
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
         Height          =   330
         Left            =   120
         TabIndex        =   10
         Top             =   105
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
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
         Caption         =   " Binder Name"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "AccountChild08.frx":069A
         Picture         =   "AccountChild08.frx":06B6
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel9 
         Height          =   330
         Left            =   120
         TabIndex        =   11
         Top             =   735
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
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
         Caption         =   " Bind Type"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "AccountChild08.frx":06D2
         Picture         =   "AccountChild08.frx":06EE
      End
      Begin FPSpreadADO.fpSpread fpSpread1 
         Height          =   2430
         Left            =   120
         TabIndex        =   4
         Top             =   1890
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
         MaxCols         =   5
         MaxRows         =   8
         OperationMode   =   2
         SpreadDesigner  =   "AccountChild08.frx":070A
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput1 
         Height          =   330
         Left            =   1920
         TabIndex        =   2
         Top             =   1050
         Width           =   3570
         _Version        =   65536
         _ExtentX        =   6297
         _ExtentY        =   582
         Calculator      =   "AccountChild08.frx":0E0B
         Caption         =   "AccountChild08.frx":0E2B
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild08.frx":0E97
         Keys            =   "AccountChild08.frx":0EB5
         Spin            =   "AccountChild08.frx":0EFF
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###########0.00"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "###########0.00"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999999999.99
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   1245189
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber MhRealInput2 
         Height          =   330
         Left            =   1920
         TabIndex        =   3
         Top             =   1365
         Width           =   3570
         _Version        =   65536
         _ExtentX        =   6297
         _ExtentY        =   582
         Calculator      =   "AccountChild08.frx":0F27
         Caption         =   "AccountChild08.frx":0F47
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "AccountChild08.frx":0FB3
         Keys            =   "AccountChild08.frx":0FD1
         Spin            =   "AccountChild08.frx":101B
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###########0.00"
         EditMode        =   1
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "###########0.00"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999999999999.99
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ""
         ShowContextMenu =   1
         ValueVT         =   1245189
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel2 
         Height          =   330
         Left            =   120
         TabIndex        =   12
         Top             =   1050
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
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
         Caption         =   " Pkt Packing Rate"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "AccountChild08.frx":1043
         Picture         =   "AccountChild08.frx":105F
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
         Height          =   330
         Left            =   120
         TabIndex        =   13
         Top             =   1365
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
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
         Caption         =   " Box Packing Rate"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   16777215
         Picture         =   "AccountChild08.frx":107B
         Picture         =   "AccountChild08.frx":1097
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   5640
         Y1              =   1790
         Y2              =   1790
      End
   End
End
Attribute VB_Name = "FrmAccountChild08"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rstAccountChild As New ADODB.Recordset
Public rstSizeList As New ADODB.Recordset
Public rstBindingTypeList As New ADODB.Recordset
Public AccountName As String
Dim SizeCode As String
Dim BindingTypeCode As String
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
    If UnloadMode = 0 Then Call CloseForm(Me)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set rstAccountChild = Nothing
    Set rstSizeList = Nothing
End Sub
Private Sub ClearFields()
    Text3.Text = ""
    Text4.Text = ""
    MhRealInput1.Text = "0.00"
    MhRealInput2.Text = "0.00"
    fpSpread1.ClearRange 1, 1, fpSpread1.MaxCols, fpSpread1.MaxRows, True
End Sub
Private Sub LoadFields()
    Dim Cnt As Integer
    If rstAccountChild.RecordCount = 0 Then Exit Sub
    If Not CheckEmpty(rstAccountChild.Fields("Size").Value, False) Then
        Text3.Text = rstAccountChild.Fields("SizeName").Value
        Text4.Text = rstAccountChild.Fields("BindingTypeName").Value
        MhRealInput1.Text = Format(rstAccountChild.Fields("PktPackRate").Value, "0.00")
        MhRealInput2.Text = Format(rstAccountChild.Fields("BoxPackRate").Value, "0.00")
        For Cnt = 1 To 8
            fpSpread1.SetText 1, Cnt, Val(rstAccountChild.Fields("Range" & IIf(Cnt = 1, "04", IIf(Cnt = 2, "06", IIf(Cnt = 3, "08", IIf(Cnt = 4, "12", IIf(Cnt = 5, "16", IIf(Cnt = 6, "24", IIf(Cnt = 7, "32", "64")))))))).Value)
            fpSpread1.SetText 2, Cnt, Val(rstAccountChild.Fields("FormFoldRate" & IIf(Cnt = 1, "04", IIf(Cnt = 2, "06", IIf(Cnt = 3, "08", IIf(Cnt = 4, "12", IIf(Cnt = 5, "16", IIf(Cnt = 6, "24", IIf(Cnt = 7, "32", "64")))))))).Value)
            fpSpread1.SetText 3, Cnt, Val(rstAccountChild.Fields("FormPasteRate" & IIf(Cnt = 1, "04", IIf(Cnt = 2, "06", IIf(Cnt = 3, "08", IIf(Cnt = 4, "12", IIf(Cnt = 5, "16", IIf(Cnt = 6, "24", IIf(Cnt = 7, "32", "64")))))))).Value)
            fpSpread1.SetText 4, Cnt, Val(rstAccountChild.Fields("FormStitchRate" & IIf(Cnt = 1, "04", IIf(Cnt = 2, "06", IIf(Cnt = 3, "08", IIf(Cnt = 4, "12", IIf(Cnt = 5, "16", IIf(Cnt = 6, "24", IIf(Cnt = 7, "32", "64")))))))).Value)
            fpSpread1.SetText 5, Cnt, Val(rstAccountChild.Fields("Rate/Book" & IIf(Cnt = 1, "04", IIf(Cnt = 2, "06", IIf(Cnt = 3, "08", IIf(Cnt = 4, "12", IIf(Cnt = 5, "16", IIf(Cnt = 6, "24", IIf(Cnt = 7, "32", "64")))))))).Value)
        Next
    End If
End Sub
Private Sub SaveFields()
    Dim Cnt As Integer, Value As Variant
    rstAccountChild.Fields("Size").Value = SizeCode
    rstAccountChild.Fields("SizeName").Value = Trim(Text3.Text)
    rstAccountChild.Fields("BindingType").Value = BindingTypeCode
    rstAccountChild.Fields("BindingTypeName").Value = Trim(Text4.Text)
    rstAccountChild.Fields("PktPackRate").Value = Val(MhRealInput1.Text)
    rstAccountChild.Fields("BoxPackRate").Value = Val(MhRealInput2.Text)
    For Cnt = 1 To 8
        fpSpread1.GetText 1, Cnt, Value
        rstAccountChild.Fields("Range" & IIf(Cnt = 1, "04", IIf(Cnt = 2, "06", IIf(Cnt = 3, "08", IIf(Cnt = 4, "12", IIf(Cnt = 5, "16", IIf(Cnt = 6, "24", IIf(Cnt = 7, "32", "64")))))))).Value = Val(Value)
        fpSpread1.GetText 2, Cnt, Value
        rstAccountChild.Fields("FormFoldRate" & IIf(Cnt = 1, "04", IIf(Cnt = 2, "06", IIf(Cnt = 3, "08", IIf(Cnt = 4, "12", IIf(Cnt = 5, "16", IIf(Cnt = 6, "24", IIf(Cnt = 7, "32", "64")))))))).Value = Val(Value)
        fpSpread1.GetText 3, Cnt, Value
        rstAccountChild.Fields("FormPasteRate" & IIf(Cnt = 1, "04", IIf(Cnt = 2, "06", IIf(Cnt = 3, "08", IIf(Cnt = 4, "12", IIf(Cnt = 5, "16", IIf(Cnt = 6, "24", IIf(Cnt = 7, "32", "64")))))))).Value = Val(Value)
        fpSpread1.GetText 4, Cnt, Value
        rstAccountChild.Fields("FormStitchRate" & IIf(Cnt = 1, "04", IIf(Cnt = 2, "06", IIf(Cnt = 3, "08", IIf(Cnt = 4, "12", IIf(Cnt = 5, "16", IIf(Cnt = 6, "24", IIf(Cnt = 7, "32", "64")))))))).Value = Val(Value)
        fpSpread1.GetText 5, Cnt, Value
        rstAccountChild.Fields("Rate/Book" & IIf(Cnt = 1, "04", IIf(Cnt = 2, "06", IIf(Cnt = 3, "08", IIf(Cnt = 4, "12", IIf(Cnt = 5, "16", IIf(Cnt = 6, "24", IIf(Cnt = 7, "32", "64")))))))).Value = Val(Value)
    Next
End Sub
Private Sub Text3_Change()
    If Text3.Text = " " Then Text3.Text = "?": SendKeys "{TAB}"
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
    Else
        SizeCode = rstSizeList.Fields("Code").Value
    End If
End Sub
Private Sub Text4_Change()
    If Text4.Text = " " Then
        Text4.Text = "?"
        SendKeys "{TAB}"
    End If
End Sub
Private Sub Text4_Validate(Cancel As Boolean)
    Dim SearchString As String
    SearchString = FixQuote(Text4.Text)
    If rstBindingTypeList.RecordCount = 0 Then
       DisplayError ("No Record in Bind Type Master")
       Cancel = True
       Exit Sub
    Else
       rstBindingTypeList.MoveFirst
    End If
    rstBindingTypeList.Find "[Col0] = '" & RTrim(SearchString) & "'"
    If rstBindingTypeList.EOF Then
        SelectionType = "S"
        BindingTypeCode = ""
        Call LoadSelectionList(rstBindingTypeList, "List of Bind Types...", "Name")
        SearchOrder = 0
        Call DisplaySelectionList(Text4, BindingTypeCode)
        Call CloseForm(FrmSelectionList)
        If CheckEmpty(Text4.Text, False) Then
            Text4.Text = "?"
        End If
        If RTrim(BindingTypeCode) <> "" Then
            SendKeys "{TAB}"
        End If
        Cancel = True
    Else
        BindingTypeCode = rstBindingTypeList.Fields("Code").Value
    End If
End Sub
Private Sub cmdProceed_Click()
    Dim Control As Object
    If CheckMandatoryFields Then Exit Sub
    SaveFields
    If Val(rstAccountChild.Fields("Range04").Value) + Val(rstAccountChild.Fields("Range06").Value) + Val(rstAccountChild.Fields("Range08").Value) + Val(rstAccountChild.Fields("Range12").Value) + Val(rstAccountChild.Fields("Range16").Value) + Val(rstAccountChild.Fields("Range24").Value) + Val(rstAccountChild.Fields("Range32").Value) + Val(rstAccountChild.Fields("Range64").Value) + Val(rstAccountChild.Fields("FormStitchRate04").Value) + Val(rstAccountChild.Fields("FormStitchRate06").Value) + Val(rstAccountChild.Fields("FormStitchRate08").Value) + Val(rstAccountChild.Fields("FormStitchRate12").Value) + Val(rstAccountChild.Fields("FormStitchRate16").Value) + Val(rstAccountChild.Fields("FormStitchRate24").Value) + Val(rstAccountChild.Fields("FormStitchRate32").Value) + Val(rstAccountChild.Fields("FormStitchRate64").Value) + _
       Val(rstAccountChild.Fields("FormPasteRate04").Value) + Val(rstAccountChild.Fields("FormPasteRate06").Value) + Val(rstAccountChild.Fields("FormPasteRate08").Value) + Val(rstAccountChild.Fields("FormPasteRate12").Value) + Val(rstAccountChild.Fields("FormPasteRate16").Value) + Val(rstAccountChild.Fields("FormPasteRate24").Value) + Val(rstAccountChild.Fields("FormPasteRate32").Value) + Val(rstAccountChild.Fields("FormPasteRate64").Value) + Val(rstAccountChild.Fields("FormFoldRate04").Value) + Val(rstAccountChild.Fields("FormFoldRate06").Value) + Val(rstAccountChild.Fields("FormFoldRate08").Value) + Val(rstAccountChild.Fields("FormFoldRate12").Value) + Val(rstAccountChild.Fields("FormFoldRate16").Value) + Val(rstAccountChild.Fields("FormFoldRate24").Value) + Val(rstAccountChild.Fields("FormFoldRate32").Value) + Val(rstAccountChild.Fields("FormFoldRate64").Value) + _
       Val(rstAccountChild.Fields("Rate/Book04").Value) + Val(rstAccountChild.Fields("Rate/Book06").Value) + Val(rstAccountChild.Fields("Rate/Book08").Value) + Val(rstAccountChild.Fields("Rate/Book12").Value) + Val(rstAccountChild.Fields("Rate/Book16").Value) + Val(rstAccountChild.Fields("Rate/Book24").Value) + Val(rstAccountChild.Fields("Rate/Book32").Value) + Val(rstAccountChild.Fields("Rate/Book64").Value) = 0 Then
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
    ElseIf CheckEmpty(Text4.Text, False) Then
        Text4.SetFocus
        CheckMandatoryFields = True
    ElseIf Not CheckExists(Text4, "Col0", rstBindingTypeList, BindingTypeCode) Then
        Text4.SetFocus
        CheckMandatoryFields = True
    End If
End Function
Private Sub fpSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    EditMode = IIf(Mode = 1, True, False)
End Sub
