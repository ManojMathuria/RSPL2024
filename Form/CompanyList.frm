VERSION 5.00
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmCompanyList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "List of Companies..."
   ClientHeight    =   5040
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7620
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CompanyList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   7620
   StartUpPosition =   2  'CenterScreen
   Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
      Height          =   330
      Left            =   45
      TabIndex        =   2
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
      Picture         =   "CompanyList.frx":0442
      Picture         =   "CompanyList.frx":045E
   End
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
      TabIndex        =   0
      ToolTipText     =   "Find"
      Top             =   4625
      Width           =   7050
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "CompanyList.frx":047A
      Height          =   4210
      Left            =   45
      TabIndex        =   1
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "Col0"
         Caption         =   "Name"
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
         Caption         =   "Financial Year"
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         ScrollBars      =   3
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         Locked          =   -1  'True
         BeginProperty Column00 
            Locked          =   -1  'True
            ColumnWidth     =   4350.047
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   2594.835
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmCompanyList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PrevStr As String
Dim dblBookMark As Double
Dim CompanyExists As Boolean
Dim rstCompanyMaster As New ADODB.Recordset
Dim rstCompanyList As New ADODB.Recordset
Private Sub Form_Load()
    Dim Cnt As Integer
    On Error GoTo ErrorHandler
    BusySystemIndicator True
    CenterForm Me
    CxnDatabase.CursorLocation = adUseClient
    For Cnt = 1 To 999
        
        If Dir(DatabasePath & "\Saral." & Pad(Cnt, "0", 3, "L")) <> "" Then
            If CxnDatabase.State = adStateOpen Then CxnDatabase.Close
            CxnDatabase.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DatabasePath & "\Saral." & Pad(Cnt, "0", 3, "L") & ";Persist Security Info=False;Jet OLEDB:Database Password=RSPLILoveMyINDIA"
            If chkFieldExists("GSTIN", "CompanyMaster") Then
                rstCompanyMaster.Open "Select Name+' [" & Pad(Cnt, "0", 3, "L") & "]' As Col01, Mid(Format(FinancialYearFrom,'dd-mm-yyyy'),1,2)+'-'+Mid(Format(FinancialYearFrom,'dd-mmm-yyyy'),4,3)+'-'+Mid(Format(FinancialYearFrom,'dd-mm-yyyy'),7,4)+' To '+Mid(Format(FinancialYearTo,'dd-mm-yyyy'),1,2)+'-'+Mid(Format(FinancialYearTo,'dd-mmm-yyyy'),4,3)+'-'+Mid(Format(FinancialYearTo,'dd-mm-yyyy'),7,4) As Col02, '" & Pad(Cnt, "0", 3, "L") & "' As Col03, MCGroup, FinancialYearFrom, FinancialYearTo,CIN,PAN,GSTIN,CreatedFrom From CompanyMaster Order By Name", CxnDatabase, adOpenKeyset, adLockReadOnly
            Else
                rstCompanyMaster.Open "Select Name+' [" & Pad(Cnt, "0", 3, "L") & "]' As Col01, Mid(Format(FinancialYearFrom,'dd-mm-yyyy'),1,2)+'-'+Mid(Format(FinancialYearFrom,'dd-mmm-yyyy'),4,3)+'-'+Mid(Format(FinancialYearFrom,'dd-mm-yyyy'),7,4)+' To '+Mid(Format(FinancialYearTo,'dd-mm-yyyy'),1,2)+'-'+Mid(Format(FinancialYearTo,'dd-mmm-yyyy'),4,3)+'-'+Mid(Format(FinancialYearTo,'dd-mm-yyyy'),7,4) As Col02, '" & Pad(Cnt, "0", 3, "L") & "' As Col03, MCGroup, FinancialYearFrom, FinancialYearTo,'' As CIN,'' As PAN,'' As GSTIN,CreatedFrom From CompanyMaster Order By Name", CxnDatabase, adOpenKeyset, adLockReadOnly
            End If
            If rstCompanyList.State = adStateClosed Then
                If chkFieldExists("GSTIN", "CompanyMaster") Then
                    rstCompanyList.Open "Select Name As Col01, Name As Col02, Name As Col03, MCGroup, FinancialYearFrom, FinancialYearTo,CIN,PAN,GSTIN,CreatedFrom From CompanyMaster Where Code = '' Order By Name", CxnDatabase, adOpenKeyset, adLockOptimistic
                Else
                    rstCompanyList.Open "Select Name As Col01, Name As Col02, Name As Col03, MCGroup, FinancialYearFrom, FinancialYearTo,Name As CIN,Name As PAN,Name As GSTIN,CreatedFrom From CompanyMaster Where Code = '' Order By Name", CxnDatabase, adOpenKeyset, adLockOptimistic
                End If
               Set rstCompanyList.ActiveConnection = Nothing
            End If
            If rstCompanyMaster.RecordCount > 0 Then
                CompanyExists = True
                rstCompanyList.AddNew
                rstCompanyList.Fields("Col01").Value = rstCompanyMaster.Fields("Col01").Value
                rstCompanyList.Fields("Col02").Value = rstCompanyMaster.Fields("Col02").Value
                rstCompanyList.Fields("Col03").Value = rstCompanyMaster.Fields("Col03").Value
                rstCompanyList.Fields("MCGroup").Value = rstCompanyMaster.Fields("MCGroup").Value
                rstCompanyList.Fields("FinancialYearFrom").Value = rstCompanyMaster.Fields("FinancialYearFrom").Value
                rstCompanyList.Fields("FinancialYearTo").Value = rstCompanyMaster.Fields("FinancialYearTo").Value
                rstCompanyList.Fields("CIN").Value = rstCompanyMaster.Fields("CIN").Value
                rstCompanyList.Fields("PAN").Value = rstCompanyMaster.Fields("PAN").Value
                rstCompanyList.Fields("GSTIN").Value = rstCompanyMaster.Fields("GSTIN").Value
                rstCompanyList.Fields("CreatedFrom").Value = rstCompanyMaster.Fields("CreatedFrom").Value
                rstCompanyList.Update
            End If
        End If
    Next
    If Not CompanyExists Then
        DisplayError ("No Company Exists")
        Call CloseForm(Me)
        Exit Sub
    End If
    rstCompanyList.Sort = "FinancialYearFrom DESC,Col01"
    DataGrid1.Columns(0).DataField = rstCompanyList.Fields(0).Name
    DataGrid1.Columns(1).DataField = rstCompanyList.Fields(1).Name
    Set DataGrid1.DataSource = rstCompanyList
    If (Not rstCompanyList.EOF) And (Not rstCompanyList.BOF) Then
        With DataGrid1.SelBookmarks
            .Add rstCompanyList.Bookmark
            If .Count <> 0 Then .Remove 0
            .Add rstCompanyList.Bookmark
        End With
    End If
    BusySystemIndicator False
    Exit Sub
ErrorHandler:
    BusySystemIndicator False
    DisplayError ("Failed to connect to database")
    Call CloseForm(Me)
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyReturn Then
        DataGrid1_DblClick
        KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyEscape Then
        CompCode = ""
        Call CloseForm(Me)
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
    Call CloseRecordset(rstCompanyList)
End Sub
Private Sub Text1_Change()
    If rstCompanyList.RecordCount = 0 Then Exit Sub
    rstCompanyList.MoveFirst
    If Text1.Text <> "" Then
        rstCompanyList.Find "[Col01] Like '" & FixQuote(Text1.Text) & "%'"
        If rstCompanyList.EOF Then
            rstCompanyList.MoveFirst
            If PrevStr <> "" And Len(Text1.Text) > 1 Then
                If dblBookMark <> 0 Then
                    rstCompanyList.Bookmark = dblBookMark
                End If
            Else
                PrevStr = ""
            End If
            Beep
            DisplayError ("Spelling Error")
            Text1.Text = PrevStr
            Sendkeys "{End}"
        Else
            PrevStr = Text1.Text
            dblBookMark = DataGrid1.Bookmark
        End If
    Else
        PrevStr = ""
    End If
    If (Not rstCompanyList.EOF) And (Not rstCompanyList.BOF) Then
        With DataGrid1.SelBookmarks
            .Add rstCompanyList.Bookmark
            If .Count <> 0 Then .Remove 0
            .Add rstCompanyList.Bookmark
        End With
    End If
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim KeyProcessed As Boolean
    
    If rstCompanyList.RecordCount = 0 Then Exit Sub
    If Shift = 0 And KeyCode = vbKeyUp Then
        With rstCompanyList
            .MovePrevious
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyBack Then
        With rstCompanyList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyDown Then
        With rstCompanyList
            .MoveNext
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageUp Then
        With rstCompanyList
            .Move (-1) * (DataGrid1.VisibleRows - 1)
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageUp Then
        With rstCompanyList
            .MoveFirst
            If .BOF Then .MoveFirst
        End With
        KeyProcessed = True
    ElseIf Shift = 0 And KeyCode = vbKeyPageDown Then
        With rstCompanyList
            .Move DataGrid1.VisibleRows - 1
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyPageDown Then
        With rstCompanyList
            .MoveLast
            If .EOF Then .MoveLast
        End With
        KeyProcessed = True
    End If
    If KeyProcessed Then
        With DataGrid1.SelBookmarks
            If .Count <> 0 Then .Remove 0
                .Add rstCompanyList.Bookmark
        End With
        KeyProcessed = False
        KeyCode = 0
    End If
End Sub
Private Sub DataGrid1_DblClick()
    If (Not rstCompanyList.EOF) And (Not rstCompanyList.BOF) Then
        CompCode = rstCompanyList.Fields("Col03").Value
        MCGroup = rstCompanyList.Fields("MCGroup").Value
        FinancialYearFrom = rstCompanyList.Fields("FinancialYearFrom").Value
        FinancialYearTo = rstCompanyList.Fields("FinancialYearTo").Value
        COMPANY_CIN = "CIN: " + rstCompanyList.Fields("CIN").Value
        COMPANY_PAN = "PAN: " + rstCompanyList.Fields("PAN").Value
        COMPANY_GSTIN = rstCompanyList.Fields("GSTIN").Value
        CreatedFrom = rstCompanyList.Fields("CreatedFrom").Value
        If CxnDatabase.State = adStateOpen Then CxnDatabase.Close
        CxnDatabase.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DatabasePath & "\Saral." & CompCode & ";Persist Security Info=False;Jet OLEDB:Database Password=RSPLILoveMyINDIA"
         Call CloseForm(Me)
    End If
End Sub
