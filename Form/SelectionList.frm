VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form FrmSelectionList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "List of..."
   ClientHeight    =   5040
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
   Icon            =   "SelectionList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   7620
   StartUpPosition =   3  'Windows Default
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
      Height          =   330
      Left            =   530
      TabIndex        =   1
      ToolTipText     =   "Find"
      Top             =   4625
      Width           =   7050
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   5
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
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Filter"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Refresh"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "SelectionList.frx":0442
      Height          =   4210
      Left            =   45
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   345
      Width           =   7530
      _ExtentX        =   13282
      _ExtentY        =   7435
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      BackColor       =   16776960
      HeadLines       =   1
      RowHeight       =   18
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "Col0"
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Col1"
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Col2"
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Col3"
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         ScrollBars      =   3
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         Locked          =   -1  'True
         BeginProperty Column00 
            Locked          =   -1  'True
            ColumnWidth     =   5580.284
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   1395.213
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   4289.953
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1289.764
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
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SelectionList.frx":0457
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SelectionList.frx":099B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox TxtCode 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   720
      TabIndex        =   4
      Top             =   2640
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.TextBox TxtName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   720
      TabIndex        =   3
      Top             =   2280
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   2160
      Width           =   375
   End
   Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
      Height          =   330
      Left            =   45
      TabIndex        =   6
      Top             =   4625
      Width           =   495
      _Version        =   65536
      _ExtentX        =   873
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
      Caption         =   " Find"
      Alignment       =   0
      FillColor       =   8421376
      TextColor       =   16777215
      Picture         =   "SelectionList.frx":0AF7
      Picture         =   "SelectionList.frx":0B13
   End
End
Attribute VB_Name = "FrmSelectionList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PrevStr As String
Dim dblBookMark As Double
Dim FormCaption As String
Public rstSelectionList As New ADODB.Recordset
Public FindFieldName As String
Private Sub Form_Activate()
    
    Dim Ctr As Integer
    If rstSelectionList.RecordCount = 0 Then Exit Sub
    Set DataGrid1.DataSource = rstSelectionList
    Do While DataGrid1.SelBookmarks.Count > 0
       DataGrid1.SelBookmarks.Remove 0
       
    Loop
    
    FormCaption = Me.Caption

    Command1.SetFocus
    
    DoEvents
    
    dblBookMark = 0
    
    TxtName.Text = LTrim(TxtName.Text)
    
    If Not CheckEmpty(TxtName, False) Then
    
        rstSelectionList.Sort = IIf(SearchOrder = 0, "Col0 Asc", "Col1 Asc")
        FindFieldName = IIf(SearchOrder = 0, "Col0", "Col1")
        rstSelectionList.MoveFirst
        For Ctr = 1 To Len(TxtName.Text)
            rstSelectionList.Find "[" & RTrim(FindFieldName) & "] Like '" & FixQuote(PrevStr & Mid(TxtName.Text, Ctr, 1)) & "%'"
            If rstSelectionList.EOF Then
                Exit For
            Else
                PrevStr = PrevStr & Mid(TxtName.Text, Ctr, 1)
                dblBookMark = rstSelectionList.Bookmark
            End If
        Next
        Text1.Text = PrevStr & Mid(TxtName.Text, Ctr, 1)
        If SearchOrder <> 0 Then
            Text1.Text = ""
            PrevStr = ""
       End If
       
    End If
    
    If SearchOrder <> 0 Then
        rstSelectionList.Sort = "Col0 Asc"
        FindFieldName = "Col0"
    End If
    If TxtName.Text <> "" And dblBookMark <> 0 Then
        rstSelectionList.Bookmark = dblBookMark
    End If
    If SelectionType <> "M" Then
        If (Not rstSelectionList.EOF) And (Not rstSelectionList.BOF) Then
            With DataGrid1.SelBookmarks
                If .Count <> 0 Then .Remove 0
                .Add DataGrid1.Bookmark
            End With
        End If
    End If
    SendKeys "{TAB}": SendKeys "{End}"
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyReturn Then
        DataGrid1_DblClick
        KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyEscape Then
        TxtCode.Text = ""
        PrevStr = ""
        rstSelectionList.Filter = adFilterNone
        rstSelectionList.Sort = vbNullString
        Me.Hide
        KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyF4 Then
       Toolbar1_ButtonClick Toolbar1.Buttons.Item(1)
    ElseIf Shift = 0 And KeyCode = vbKeyF5 Then
       Toolbar1_ButtonClick Toolbar1.Buttons.Item(2)
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
       Call Form_KeyDown(vbKeyEscape, 0)
       Cancel = 1
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set rstSelectionList = Nothing
End Sub
Private Sub Text1_Change()
    If rstSelectionList.RecordCount = 0 Then Exit Sub
    rstSelectionList.MoveFirst
    If Text1.Text <> "" Then
        rstSelectionList.Find "[" & RTrim(FindFieldName) & "] Like '" & FixQuote(Text1.Text) & "%'"
        If rstSelectionList.EOF Then
            rstSelectionList.MoveFirst
            If PrevStr <> "" And Len(Text1.Text) > 1 Then
                If dblBookMark <> 0 Then
                    rstSelectionList.Bookmark = dblBookMark
                End If
            Else
                PrevStr = ""
            End If
            Beep
            Text1.Text = PrevStr
            SendKeys "{End}"
        Else
            PrevStr = Text1.Text
            dblBookMark = DataGrid1.Bookmark
        End If
    Else
        PrevStr = ""
    End If
    If SelectionType <> "M" Then
       If (Not rstSelectionList.EOF) And (Not rstSelectionList.BOF) Then
           With DataGrid1.SelBookmarks
               If .Count <> 0 Then .Remove 0
               .Add DataGrid1.Bookmark
           End With
       End If
    End If
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim KeyProcessed As Boolean
    
    If rstSelectionList.RecordCount = 0 Then Exit Sub
    If Shift = 0 And KeyCode = vbKeyUp Then
        With rstSelectionList
            .MovePrevious
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyBack Then
        With rstSelectionList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyDown Then
        With rstSelectionList
            .MoveNext
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageUp Then
        With rstSelectionList
            .Move (-1) * (DataGrid1.VisibleRows - 1)
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageUp Then
        With rstSelectionList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageDown Then
        With rstSelectionList
            .Move DataGrid1.VisibleRows - 1
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageDown Then
        With rstSelectionList
            .MoveLast
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    End If
    If KeyProcessed Then
        If SelectionType <> "M" Then
           With DataGrid1.SelBookmarks
               If .Count <> 0 Then .Remove 0
               .Add DataGrid1.Bookmark
           End With
        End If
        KeyProcessed = False
        KeyCode = 0
    End If
End Sub
Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
    If ColIndex = 0 Then
       If FindFieldName <> "Col0" Then
          FindFieldName = "Col0"
          rstSelectionList.Sort = "Col0 Asc"
       End If
    ElseIf ColIndex = 1 Then
       If FindFieldName <> "Col1" Then
          FindFieldName = "Col1"
          rstSelectionList.Sort = "Col1 Asc"
       End If
    ElseIf ColIndex = 2 Then
       If FindFieldName <> "Col2" Then
          FindFieldName = "Col2"
          rstSelectionList.Sort = "Col2 Asc"
       End If
    End If
    DataGrid1.ClearSelCols
    If SelectionType <> "M" Then
       If (Not rstSelectionList.EOF) And (Not rstSelectionList.BOF) Then
           With DataGrid1.SelBookmarks
               If .Count <> 0 Then .Remove 0
               .Add DataGrid1.Bookmark
           End With
       End If
    End If
    Text1.Text = ""
    Text1.SetFocus
End Sub
Private Sub DataGrid1_DblClick()
  If SelectionType <> "M" Then
     If (Not rstSelectionList.EOF) And (Not rstSelectionList.BOF) Then
        TxtName.Text = rstSelectionList.Fields(IIf(SearchOrder = 1, 1, 0)).Value
        
        TxtCode.Text = rstSelectionList.Fields("Code").Value
        PrevStr = ""
        
        rstSelectionList.Filter = adFilterNone
        rstSelectionList.Sort = vbNullString
        Me.Hide
     End If
  Else
     DataGrid1.SelBookmarks.Add DataGrid1.Bookmark
  End If
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    
    If Button.Index = 1 Then
        With FrmFilter
            .Combo1.AddItem DataGrid1.Columns(0).Caption, 0
            If DataGrid1.Columns(1).Visible Then
                .Combo1.AddItem DataGrid1.Columns(1).Caption, 1
            End If
            If DataGrid1.Columns(2).Visible Then
                .Combo1.AddItem DataGrid1.Columns(2).Caption, 2
            End If
            If DataGrid1.Columns(3).Visible Then
                .Combo1.AddItem DataGrid1.Columns(3).Caption, 3
            End If
            .Combo1.ListIndex = 0
            Set .srcForm = Me
            .Show vbModal
        End With
    ElseIf Button.Index = 2 Then
        Set DataGrid1.DataSource = Nothing
        rstSelectionList.ActiveConnection = CxnDatabase
        Do While Not RefreshRecord(rstSelectionList)
        Loop
        Set DataGrid1.DataSource = rstSelectionList
        rstSelectionList.ActiveConnection = Nothing
    End If
    If SelectionType <> "M" Then
       If (Not rstSelectionList.EOF) And (Not rstSelectionList.BOF) Then
           With DataGrid1.SelBookmarks
               If .Count <> 0 Then .Remove 0
               .Add DataGrid1.Bookmark
           End With
       End If
    End If
    Text1.SetFocus
End Sub
Public Sub FilterRecord(ByVal SrchFor As String, ByVal SrchText As String)

    If SrchFor = DataGrid1.Columns(0).Caption Then
        rstSelectionList.Filter = "[Col0] Like '%" & SrchText & "%'"
        
    ElseIf SrchFor = DataGrid1.Columns(1).Caption Then
        rstSelectionList.Filter = "[Col1] Like '%" & SrchText & "%'"
    ElseIf SrchFor = DataGrid1.Columns(2).Caption Then
        rstSelectionList.Filter = "[Col2] Like '%" & SrchText & "%'"
    ElseIf SrchFor = DataGrid1.Columns(3).Caption Then
        rstSelectionList.Filter = "[Col3] =" & Val(SrchText)
    End If
End Sub
