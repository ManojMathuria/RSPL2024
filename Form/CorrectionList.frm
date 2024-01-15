VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form FrmCorrectionList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "List of Corrections"
   ClientHeight    =   7365
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   9420
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CorrectionList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7365
   ScaleWidth      =   9420
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9420
      _ExtentX        =   16616
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
      Begin VB.CheckBox Check1 
         Caption         =   "All Entries"
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
         Left            =   8250
         TabIndex        =   4
         Top             =   70
         Width           =   1140
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
            Picture         =   "CorrectionList.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CorrectionList.frx":0986
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CorrectionList.frx":0A9A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame1 
      Height          =   6975
      Left            =   45
      TabIndex        =   3
      Top             =   345
      Width           =   9330
      _Version        =   65536
      _ExtentX        =   16457
      _ExtentY        =   12303
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
      Picture         =   "CorrectionList.frx":0BAE
      Begin MSComctlLib.ListView ListView1 
         Height          =   6975
         Left            =   0
         TabIndex        =   0
         Top             =   0
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   12303
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
      Begin MSComctlLib.ListView ListView2 
         Height          =   6975
         Left            =   3990
         TabIndex        =   1
         Top             =   0
         Width           =   5340
         _ExtentX        =   9419
         _ExtentY        =   12303
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
Attribute VB_Name = "FrmCorrectionList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstCompanyMaster As New ADODB.Recordset
Dim rstBookList As New ADODB.Recordset
Dim rstBoardList As New ADODB.Recordset
Dim rstCorrectionList As New ADODB.Recordset
Dim OutputTo As String
Public Department As String
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    
    CenterForm Me
    BusySystemIndicator True
    rstCompanyMaster.Open "Select PrintName From CompanyMaster", CxnDatabase, adOpenKeyset, adLockReadOnly
    rstBoardList.Open "Select Name, Code From GeneralMaster Where Type = '2' Order By Name", CxnDatabase, adOpenKeyset, adLockReadOnly
    rstBoardList.ActiveConnection = Nothing
    Call FillList(ListView1, "List of Boards...", rstBoardList)
    Call BookSelection(True)
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
        Call CloseForm(Me)
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CloseRecordset(rstCompanyMaster)
    Call CloseRecordset(rstBoardList)
    Call CloseRecordset(rstBookList)
    Call CloseRecordset(rstCorrectionList)
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
        PrintCorrectionList
    ElseIf Button.Index = 2 Then
        OutputTo = "P"
        PrintCorrectionList
    ElseIf Button.Index = 3 Then
        Call CloseForm(Me)
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
Private Sub PrintCorrectionList()
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    rptCorrectionList.Text12.SetText Trim(rstCompanyMaster.Fields("PrintName").Value)
    If rstCorrectionList.State = adStateOpen Then
        rstCorrectionList.Close
    End If
    rstCorrectionList.Open "SELECT M1.PrintName As BookName,M2.PrintName As BoardName,C.Correction,C.ArrivedOn,C.RectifiedOn FROM (BookMaster M1 INNER JOIN BookChild02 C ON M1.Code=C.Code) INNER JOIN GeneralMaster M2 ON M1.Board=M2.Code " & _
                                 "WHERE C.Department='" & Department & "' AND " & IIf(Check1.Value, "1=1", "ISNULL(RectifiedOn)") & " AND M2.Code In (" & SelectedItems(ListView1) & ") AND M1.Code In (" & SelectedItems(ListView2) & ") ORDER BY M2.PrintName,M1.PrintName,C.SNo", CxnDatabase, adOpenKeyset, adLockOptimistic
    Screen.MousePointer = vbNormal
    If rstCorrectionList.RecordCount = 0 Then
        On Error GoTo 0
        Exit Sub
    End If
    rptCorrectionList.Database.SetDataSource rstCorrectionList, 3, 1
    rptCorrectionList.DiscardSavedData
    If OutputTo = "S" Then
        Set FrmReportViewer.Report = rptCorrectionList
        FrmReportViewer.Show vbModal
    Else
        rptCorrectionList.PrintOut
    End If
    Set rptCorrectionList = Nothing
    On Error GoTo 0
End Sub
