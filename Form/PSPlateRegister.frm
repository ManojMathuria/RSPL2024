VERSION 5.00
Object = "{3AE5AE83-A6DA-101B-9313-00AA00575482}#1.0#0"; "mhfram32.ocx"
Object = "{886939C3-7807-101C-BB03-00AA00575482}#1.0#0"; "mhlabl32.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form FrmPSPlateRegister 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PS Plate Register"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15120
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "PSPlateRegister.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "FrmLogin"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrintRegister 
      Height          =   375
      Left            =   14640
      Picture         =   "PSPlateRegister.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Exit"
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton cmdExit 
      Height          =   375
      Left            =   14640
      Picture         =   "PSPlateRegister.frx":0544
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Exit"
      Top             =   105
      Width           =   375
   End
   Begin Mh3dfrmLibCtl.Mh3dFrame Mh3dFrame2 
      Height          =   6855
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   105
      Width           =   14415
      _Version        =   65536
      _ExtentX        =   25426
      _ExtentY        =   12091
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
      Picture         =   "PSPlateRegister.frx":0646
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
         Left            =   1200
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   3
         Top             =   105
         Width           =   13095
      End
      Begin FPSpreadADO.fpSpread fpSpread1 
         Height          =   6105
         Left            =   120
         TabIndex        =   0
         Top             =   630
         Width           =   14175
         _Version        =   524288
         _ExtentX        =   25003
         _ExtentY        =   10769
         _StockProps     =   64
         EditEnterAction =   5
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
         MaxCols         =   10
         OperationMode   =   2
         SpreadDesigner  =   "PSPlateRegister.frx":0662
      End
      Begin Mh3dlblLib.Mh3dLabel Mh3dLabel1 
         Height          =   330
         Left            =   120
         TabIndex        =   4
         Top             =   105
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
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
         Picture         =   "PSPlateRegister.frx":0DB3
         Picture         =   "PSPlateRegister.frx":0DCF
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   14400
         Y1              =   525
         Y2              =   525
      End
   End
End
Attribute VB_Name = "FrmPSPlateRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public BookCode As String
Public BookName As String
Public OrderCode As String
Public OrderDate As Date
Public OrderType As String
Public PlateType As String
Dim DatabaseName As String
Dim CxnImporter As New ADODB.Connection
Dim rstImporter As New ADODB.Recordset
Dim rstPSPlateRegister As New ADODB.Recordset
Private Sub Form_Load()
    Dim i As Integer
    On Error GoTo ErrorHandler
    CenterForm Me
    Text2.Text = Trim(BookName)
    fpSpread1.ClearRange 1, 1, fpSpread1.MaxCols, fpSpread1.MaxRows, True
    DatabaseName = Trim(ReadFromFile("Saral Database Name"))
    Screen.MousePointer = vbHourglass
    CxnImporter.CursorLocation = adUseClient
    If CxnImporter.State = adStateOpen Then
        CxnImporter.Close
    End If
'    i = InStr(1, DatabaseName, ",")
'    DatabaseName = Mid(DatabaseName, i + 1, InStr(i + 1, DatabaseName, ",") - i - 1)
'    DatabaseName = IIf(Val(CompCode) < 11, "Saral.00" & Val(CompCode) - 1, "Saral.0" & Val(CompCode) - 1)
    DatabaseName = "Saral." & CreatedFrom
    CxnImporter.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DatabasePath & "\" & DatabaseName & ";Persist Security Info=False;Jet OLEDB:Database Password=RSPLILoveMyINDIA"

    
    
    rstImporter.Open "SELECT P.Name As OrderNo,C.OrderDate,M1.PrintName As PrinterName,C.Processing,C.PlateType" & PlateType & ",C.ActualQuantity As Quantity,C.PlateRate" & PlateType & " As Rate,C.BillNo,C.BillDate,M2.PrintName,C.Remarks FROM ((BookPOParent P INNER JOIN BookPOChild" & OrderType & " C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON P." & IIf(OrderType = "06", "TitlePrinter", "BookPrinter") & "=M1.Code) INNER JOIN BookMaster M2 ON P.Book=M2.Code WHERE P.Type<>'O' AND LEFT(P.Code,1)<>'*' AND M2.Code='" & BookCode & "' AND C.OrderDate>=#" & GetDate(Format(DateAdd("d", -365, CDate(OrderDate)), "dd-mm-yyyy")) & "# " & _
                                 "ORDER BY M1.PrintName,C.OrderDate", CxnImporter, adOpenKeyset, adLockReadOnly
    rstImporter.ActiveConnection = Nothing
    
    
    
    
    
    rstPSPlateRegister.Open "SELECT P.Name As OrderNo,C.OrderDate,M1.PrintName As PrinterName,C.Processing,C.PlateType" & PlateType & ",C.ActualQuantity As Quantity,C.PlateRate" & PlateType & " As Rate,C.BillNo,C.BillDate,M2.PrintName,C.Remarks FROM ((BookPOParent P INNER JOIN BookPOChild" & OrderType & " C ON P.Code=C.Code) INNER JOIN AccountMaster M1 ON P." & IIf(OrderType = "06", "TitlePrinter", "BookPrinter") & "=M1.Code) INNER JOIN BookMaster M2 ON P.Book=M2.Code WHERE P.Type<>'O' AND LEFT(P.Code,1)<>'*' AND M2.Code='" & BookCode & "' AND C.Code<'" & IIf(OrderCode = "", "999999", OrderCode) & "' AND C.OrderDate<=#" & GetDate(Format(OrderDate, "dd-mm-yyyy")) & "# " & _
                                 "ORDER BY M1.PrintName,C.OrderDate", CxnDatabase, adOpenKeyset, adLockReadOnly
    rstPSPlateRegister.ActiveConnection = Nothing
    i = 1
    If rstImporter.RecordCount > 0 Then rstImporter.MoveFirst
    
    Do While Not rstImporter.EOF
        fpSpread1.SetText 1, i, Trim(rstImporter.Fields("OrderNo").Value)
        fpSpread1.SetText 2, i, Format(rstImporter.Fields("OrderDate").Value, "dd-mm-yyyy")
        fpSpread1.SetText 3, i, Trim(rstImporter.Fields("PrinterName").Value)
        fpSpread1.SetText 4, i, IIf(rstImporter.Fields("Processing").Value = "O", "", "New")
        fpSpread1.SetText 5, i, IIf(rstImporter.Fields("PlateType" & PlateType).Value = "1", "Deepatch", IIf(rstImporter.Fields("PlateType" & PlateType).Value = "2", "PS", IIf(rstImporter.Fields("PlateType" & PlateType).Value = "3", "Wipeon", "CTP")))
        fpSpread1.SetText 6, i, Val(rstImporter.Fields("Quantity").Value)
        fpSpread1.SetText 7, i, Val(rstImporter.Fields("Rate").Value)
        fpSpread1.SetText 8, i, Trim(rstImporter.Fields("BillNo").Value)
        fpSpread1.SetText 9, i, Format(rstImporter.Fields("BillDate").Value, "dd-mm-yyyy")
        fpSpread1.SetText 10, i, Trim(rstImporter.Fields("Remarks").Value)
        i = i + 1
        rstImporter.MoveNext
    Loop
    If rstPSPlateRegister.RecordCount > 0 Then rstPSPlateRegister.MoveFirst
    Do While Not rstPSPlateRegister.EOF
        fpSpread1.SetText 1, i, Trim(rstPSPlateRegister.Fields("OrderNo").Value)
        fpSpread1.SetText 2, i, Format(rstPSPlateRegister.Fields("OrderDate").Value, "dd-mm-yyyy")
        fpSpread1.SetText 3, i, Trim(rstPSPlateRegister.Fields("PrinterName").Value)
        fpSpread1.SetText 4, i, IIf(rstPSPlateRegister.Fields("Processing").Value = "O", "", "New")
        fpSpread1.SetText 5, i, IIf(rstPSPlateRegister.Fields("PlateType" & PlateType).Value = "1", "Deepatch", IIf(rstPSPlateRegister.Fields("PlateType" & PlateType).Value = "2", "PS", IIf(rstPSPlateRegister.Fields("PlateType" & PlateType).Value = "3", "Wipeon", "CTP")))
        fpSpread1.SetText 6, i, Val(rstPSPlateRegister.Fields("Quantity").Value)
        fpSpread1.SetText 7, i, Val(rstPSPlateRegister.Fields("Rate").Value)
        fpSpread1.SetText 8, i, Trim(rstPSPlateRegister.Fields("BillNo").Value)
        fpSpread1.SetText 9, i, Format(rstPSPlateRegister.Fields("BillDate").Value, "dd-mm-yyyy")
        i = i + 1
        rstPSPlateRegister.MoveNext
    Loop
    Screen.MousePointer = vbNormal
    Exit Sub
ErrorHandler:
    Screen.MousePointer = vbNormal
    Call CloseForm(Me)
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyReturn Then
        Sendkeys "{TAB}"
        KeyCode = 0
    ElseIf Shift = 0 And KeyCode = vbKeyEscape Then
        cmdExit_Click
        KeyCode = 0
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        Call CloseForm(Me)
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call CloseRecordset(rstImporter)
    Call CloseRecordset(rstPSPlateRegister)
    Call CloseConnection(CxnImporter)
End Sub
Private Sub cmdExit_Click()
    Call CloseForm(Me)
End Sub
Private Sub cmdPrintRegister_Click()
    If Not FileExist(App.Path & "\Template\Plate Usage Register.xlsx") Then Exit Sub
    Dim oExcel As Object
    Screen.MousePointer = vbHourglass
    Set oExcel = CreateObject("Excel.Application")
    oExcel.Workbooks.Open (App.Path & "\Template\Plate Usage Register")
    oExcel.DisplayAlerts = False
    oExcel.Workbooks.Item(1).SaveAs (App.Path & "\Report\Plate Usage Register (" & CompCode & ")")
    oExcel.DisplayAlerts = True
    oExcel.Application.Cells(2, "A").Value = "Book Name : " & Trim(Text2.Text)
    Dim i As Integer, K As Integer, R As Integer, CellVal As Variant
    R = 3
    For i = 1 To fpSpread1.DataRowCnt
        K = K + 1: R = R + 1
        oExcel.Application.Cells(R, "A").Value = K
        fpSpread1.GetText 1, i, CellVal
        oExcel.Application.Cells(R, "B").Value = CellVal
        fpSpread1.GetText 2, i, CellVal
        oExcel.Application.Cells(R, "C").Value = CellVal
        fpSpread1.GetText 3, i, CellVal
        oExcel.Application.Cells(R, "D").Value = CellVal
        fpSpread1.GetText 4, i, CellVal
        oExcel.Application.Cells(R, "E").Value = CellVal
        fpSpread1.GetText 5, i, CellVal
        oExcel.Application.Cells(R, "F").Value = CellVal
        fpSpread1.GetText 6, i, CellVal
        oExcel.Application.Cells(R, "G").Value = CellVal
        fpSpread1.GetText 7, i, CellVal
        oExcel.Application.Cells(R, "H").Value = CellVal
        fpSpread1.GetText 8, i, CellVal
        oExcel.Application.Cells(R, "I").Value = CellVal
        fpSpread1.GetText 9, i, CellVal
        oExcel.Application.Cells(R, "J").Value = CellVal
    Next
    oExcel.Columns("A:J").EntireColumn.AutoFit
    oExcel.Workbooks.Item(1).Save
    Screen.MousePointer = vbNormal
    oExcel.Range("A1").Activate
    oExcel.Workbooks.Item(1).PrintOut
    oExcel.Workbooks.Close
    oExcel.Quit
    Set oExcel = Nothing
End Sub
