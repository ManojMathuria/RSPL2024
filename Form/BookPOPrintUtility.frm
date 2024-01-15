VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Begin VB.Form FrmBookPOPrintUtility 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Book Print Order Print Utility"
   ClientHeight    =   4365
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   3840
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "BookPOPrintUtility.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4365
   ScaleWidth      =   3840
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   3840
      _ExtentX        =   6773
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Mail"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Print"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exit"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   3980
      Left            =   45
      TabIndex        =   9
      Top             =   345
      Width           =   3765
      _Version        =   65536
      _ExtentX        =   6641
      _ExtentY        =   7020
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
      Picture         =   "BookPOPrintUtility.frx":0442
      Begin VB.TextBox txtBookOrder 
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
         Left            =   1200
         MaxLength       =   100
         TabIndex        =   7
         Top             =   3650
         Width           =   2565
      End
      Begin VB.TextBox txtBinding 
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
         Left            =   1200
         MaxLength       =   100
         TabIndex        =   6
         Top             =   3330
         Width           =   2565
      End
      Begin VB.TextBox txtLamination 
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
         Left            =   1200
         MaxLength       =   100
         TabIndex        =   5
         Top             =   3020
         Width           =   2565
      End
      Begin VB.TextBox txtTitlePrinting 
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
         Left            =   1200
         MaxLength       =   100
         TabIndex        =   4
         Top             =   2700
         Width           =   2565
      End
      Begin VB.TextBox txtBookPrinting 
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
         Left            =   1200
         MaxLength       =   100
         TabIndex        =   3
         Top             =   2385
         Width           =   2565
      End
      Begin VB.TextBox MhRealInput16 
         Alignment       =   1  'Right Justify
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
         Left            =   2670
         MaxLength       =   13
         TabIndex        =   1
         Top             =   0
         Width           =   1095
      End
      Begin VB.TextBox MhRealInput15 
         Alignment       =   1  'Right Justify
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
         Left            =   840
         MaxLength       =   13
         TabIndex        =   0
         Top             =   0
         Width           =   1095
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1910
         Left            =   0
         TabIndex        =   2
         Top             =   320
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   3360
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
         Picture         =   "BookPOPrintUtility.frx":045E
         Picture         =   "BookPOPrintUtility.frx":047A
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
         Picture         =   "BookPOPrintUtility.frx":0496
         Picture         =   "BookPOPrintUtility.frx":04B2
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel4 
         Height          =   330
         Left            =   0
         TabIndex        =   13
         Top             =   2385
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
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
         Caption         =   " Book Printing"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   -2147483634
         Picture         =   "BookPOPrintUtility.frx":04CE
         Picture         =   "BookPOPrintUtility.frx":04EA
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel5 
         Height          =   330
         Left            =   0
         TabIndex        =   14
         Top             =   2700
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
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
         Caption         =   " Title Printing"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   -2147483634
         Picture         =   "BookPOPrintUtility.frx":0506
         Picture         =   "BookPOPrintUtility.frx":0522
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel6 
         Height          =   330
         Left            =   0
         TabIndex        =   15
         Top             =   3020
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
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
         Caption         =   " Lamination"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   -2147483634
         Picture         =   "BookPOPrintUtility.frx":053E
         Picture         =   "BookPOPrintUtility.frx":055A
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel7 
         Height          =   330
         Left            =   0
         TabIndex        =   16
         Top             =   3330
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
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
         Caption         =   " Binding"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   -2147483634
         Picture         =   "BookPOPrintUtility.frx":0576
         Picture         =   "BookPOPrintUtility.frx":0592
      End
      Begin Mh3dlblLib.Mh3dLabel lblBookOrder 
         Height          =   330
         Left            =   0
         TabIndex        =   17
         Top             =   3650
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
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
         Caption         =   " Book Order"
         Alignment       =   0
         FillColor       =   8421376
         TextColor       =   -2147483634
         Picture         =   "BookPOPrintUtility.frx":05AE
         Picture         =   "BookPOPrintUtility.frx":05CA
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   3750
         Y1              =   2300
         Y2              =   2300
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3000
      Top             =   1080
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
            Picture         =   "BookPOPrintUtility.frx":05E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BookPOPrintUtility.frx":06FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BookPOPrintUtility.frx":080C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Mh3dlblLib.Mh3dLabel Mh3dLabel3 
      Height          =   330
      Left            =   120
      TabIndex        =   12
      Top             =   2760
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
      Picture         =   "BookPOPrintUtility.frx":0920
      Picture         =   "BookPOPrintUtility.frx":093C
   End
End
Attribute VB_Name = "FrmBookPOPrintUtility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstBookPrintOrder As New ADODB.Recordset
Public PrintUtility As Boolean
Private Sub Form_Load()
    Dim i As Integer
    On Error GoTo ErrorHandler
    Unload FrmBookPrintOrder
    CenterForm Me
    BusySystemIndicator True
    ListView1.ColumnHeaders.Add 1, , "List of Order Types"
    For i = 1 To 6
        ListView1.ListItems.Add , , Choose(i, "Book Printing Order", "Title Printing Order", "Title Lamination Order", "Book Binding Order", "Book Order", "Box Label")
        ListView1.ListItems(i).Checked = True
    Next
    LockWindowUpdate ListView1.hwnd
    ListView1.ColumnHeaders(1).Width = ListView1.Width
    LockWindowUpdate 0
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
    ElseIf Shift = vbAltMask And KeyCode = vbKeyM Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(1)
        KeyCode = 0
    ElseIf Shift = vbAltMask And KeyCode = vbKeyP Then
        Toolbar1_ButtonClick Toolbar1.Buttons.Item(2)
        KeyCode = 0
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        Call CloseForm(Me)
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CloseRecordset(rstBookPrintOrder)
End Sub
Private Sub MhRealInput15_GotFocus()
    FocusSelect Me.ActiveControl
End Sub
Private Sub MhRealInput15_KeyPress(KeyAscii As Integer)
    ValidateKey MhRealInput15, KeyAscii, 0
End Sub
Private Sub MhRealInput15_Validate(Cancel As Boolean)
    If Not ValidateNumber(Me.ActiveControl, 0) Then
        Cancel = True
    ElseIf Val(MhRealInput15.Text) <= 0 Then
        MhRealInput15.SetFocus
        FocusSelect Me.ActiveControl
        Cancel = True
    End If
End Sub
Private Sub MhRealInput16_GotFocus()
    FocusSelect Me.ActiveControl
End Sub
Private Sub MhRealInput16_KeyPress(KeyAscii As Integer)
    ValidateKey MhRealInput16, KeyAscii, 0
End Sub
Private Sub MhRealInput16_Validate(Cancel As Boolean)
    If Not ValidateNumber(Me.ActiveControl, 0) Then
        Cancel = True
    ElseIf Val(MhRealInput16.Text) <= 0 Or Val(MhRealInput16.Text) < Val(MhRealInput15.Text) Then
        MhRealInput16.SetFocus
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
        Call PrintBookPrintOrder("M")
    ElseIf Button.Index = 2 Then
        Call PrintBookPrintOrder("P")
    ElseIf Button.Index = 3 Then
        Call CloseForm(Me)
    End If
End Sub
Private Sub PrintBookPrintOrder(ByVal OutputType As String)
    Screen.MousePointer = vbHourglass
    If rstBookPrintOrder.State = adStateOpen Then
        rstBookPrintOrder.Close
    End If
    rstBookPrintOrder.Open "Select Code From BookPOParent Where Type<>'O' And Left(Code,1)<>'*' And Name >= '" & Pad(Trim(MhRealInput15.Text), Space(1), 10, "L") & "' And Name <= '" & Pad(Trim(MhRealInput16.Text), Space(1), 10, "L") & "' Order By BookPOParent.Name", CxnDatabase, adOpenKeyset, adLockReadOnly
    Screen.MousePointer = vbNormal
    If rstBookPrintOrder.RecordCount = 0 Then
        On Error GoTo 0
        Exit Sub
    End If
    Do While Not rstBookPrintOrder.EOF
        PrintUtility = True
        If ListView1.ListItems(1).Checked Then Call FrmBookPrintOrder.PrintBookPrintingOrder(rstBookPrintOrder.Fields("Code").Value, txtBookPrinting.Text, OutputType)
        If ListView1.ListItems(2).Checked Then Call FrmBookPrintOrder.PrintTitlePrintingOrder(rstBookPrintOrder.Fields("Code").Value, txtTitlePrinting.Text, OutputType)
        If ListView1.ListItems(3).Checked Then Call FrmBookPrintOrder.PrintTitleLaminationOrder(rstBookPrintOrder.Fields("Code").Value, txtLamination.Text, OutputType)
        If ListView1.ListItems(4).Checked Then Call FrmBookPrintOrder.PrintBookBindingOrder(rstBookPrintOrder.Fields("Code").Value, txtBinding.Text, OutputType)
        If ListView1.ListItems(5).Checked Then Call FrmBookPrintOrder.PrintBookOrder(rstBookPrintOrder.Fields("Code").Value, txtBookOrder.Text, OutputType)
        If ListView1.ListItems(6).Checked Then If OutputType = "P" Then Call FrmBookPrintOrder.PrintBookBoxLabel(rstBookPrintOrder.Fields("Code").Value, txtBinding.Text, OutputType)
        rstBookPrintOrder.MoveNext
    Loop
    PrintUtility = False
End Sub
