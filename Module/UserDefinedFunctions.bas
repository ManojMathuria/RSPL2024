Attribute VB_Name = "ModUserDefinedFunctions"
Option Explicit
Public CxnDatabase As New ADODB.Connection
Public FinancialYearFrom As Date
Public FinancialYearTo As Date
Global FSO As New FileSystemObject
Global CompCode As String
Global CreatedFrom As String
Global MCGroup As String
Global DatabasePath As String
Global SearchOrder As Integer
Global SelectionType As String
Global LoginSuccess As Boolean
Global UserCode As String
Global UserName As String
Global UserLevel As String
Global AllowMastersModification As Integer
Global AllowMastersDeletion As Integer
Global AllowTransactionsModification As Integer
Global AllowTransactionsDeletion As Integer

Global ServerName As String
Global ServerPassword As String
Global ConnectionString As String
Global LoginPassword As String
Global AbortPO As Boolean
Dim LocalHwnd As Long
Dim LocalPrevWndProc As Long
Dim MyControl As Object
Dim LeftHand_Odd() As Variant
Dim LeftHand_Even() As Variant
Dim Right_Hand() As Variant
Dim Parity() As Variant
Dim BarH As Long
Dim xObj As Object
Dim Xpos As Long, xTop As Long
Public Const WM_SETTEXT = &HC
Public Const WM_CLOSE = &H10
Public Const SM_CXSCREEN = 0
Public Const SM_CYSCREEN = 1
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWNORMAL = 1
Public Const SC_RESTORE = &HF120&
Public Const WM_SYSCOMMAND = &H112
Public Const CB_SETDROPPEDWIDTH = &H160
Public Const CB_SHOWDROPDOWN = &H14F
Public Const CB_GETITEMHEIGHT = &H154
Public Const WM_USER As Long = &H400
Public Const SB_GETRECT As Long = (WM_USER + 10)
Public Const MIIM_STATE = &H1
Public Const MIIM_ID = &H2
Public Const MIIM_SUBMENU = &H4
Public Const MIIM_TYPE = &H10
Public Const MFT_SEPARATOR = &H800
Public Const MFT_STRING = &H0
Public Const MFS_ENABLED = &H0
Public Const TPM_LEFTALIGN = &H0
Public Const TPM_TOPALIGN = &H0
Public Const TPM_NONOTIFY = &H80
Public Const TPM_RETURNCMD = &H100
Public Const TPM_LEFTBUTTON = &H0
Public Const MF_BYPOSITION = &H400&
Public Const AW_BLEND = &H80000 ' Uses a fade effect. This flag can be used only if hwnd is a top-level window.
Public Const AW_HIDE = &H10000 ' Hides the window. By default, the window is shown.
Private Const GWL_WNDPROC = -4
Private Const WM_MOUSEWHEEL = &H20A
Public Type POINTAPI
        x As Long
        Y As Long
End Type
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Public Type WINDOWPLACEMENT
        Length As Long
        flags As Long
        showCmd As Long
        ptMinPosition As POINTAPI
        ptMaxPosition As POINTAPI
        rcNormalPosition As RECT
End Type
Public Type POINT_TYPE
    x As Long
    Y As Long
End Type

Public Type MENUITEMINFO
        cbSize As Long
        fMask As Long
        fType As Long
        fState As Long
        wID As Long
        hSubMenu As Long
        hbmpChecked As Long
        hbmpUnchecked As Long
        dwItemData As Long
        dwTypeData As String
        cch As Long
End Type

Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Public Declare Function SetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Public Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function CreatePopupMenu Lib "user32.dll" () As Long
Public Declare Function DestroyMenu Lib "user32.dll" (ByVal hMenu As Long) As Long
Public Declare Function InsertMenuItem Lib "user32.dll" Alias "InsertMenuItemA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As MENUITEMINFO) As Long
Public Declare Function TrackPopupMenu Lib "user32.dll" (ByVal hMenu As Long, ByVal uFlags As Long, ByVal x As Long, ByVal Y As Long, ByVal nReserved As Long, ByVal hwnd As Long, ByVal prcRect As Long) As Long
Public Declare Function GetCursorPos Lib "user32.dll" (lpPoint As POINT_TYPE) As Long
Public Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Public Declare Function WritePrinter Lib "winspool.drv" (ByVal hPrinter As Long, pBuf As Any, ByVal cdBuf As Long, pcWritten As Long) As Long
Public Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, pDefault As Any) As Long
Public Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Public Declare Function StartDocPrinter Lib "winspool.drv" Alias "StartDocPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pDocInfo As Byte) As Long
Public Declare Function EndDocPrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Public Declare Function AnimateWindow Lib "user32" (ByVal hwnd As Long, ByVal dwTime As Long, ByVal dwFlags As Long) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Public COMPANY_CIN  As String
Public COMPANY_GSTIN As String
Public COMPANY_PAN As String
Public Function CheckEmpty(ByVal strExpression As Variant, ByVal xDspMsg As Boolean) As Boolean
    If LTrim(RTrim(strExpression)) = "" Or IsNull(strExpression) Then
       If xDspMsg Then
          DisplayError ("Mandatory Field")
       End If
       CheckEmpty = True
    End If
End Function

Public Function CheckNull(ByVal Expression As Variant) As String
    If IsNull(Expression) Then
        CheckNull = ""
    Else
        CheckNull = Expression
    End If
End Function


Public Sub LoadSelectionList(ByRef xRecordset As Recordset, ByVal xListHeader As String, ByVal xColumn1Header As String, Optional ByVal xColumn2Header As String, Optional ByVal xColumn3Header As String, Optional ByVal xColumn4Header As String)
    
    Set FrmSelectionList.rstSelectionList = xRecordset
    
    FrmSelectionList.rstSelectionList.Sort = "Col0 Asc"
    FrmSelectionList.Caption = xListHeader
    FrmSelectionList.DataGrid1.Columns(0).Caption = xColumn1Header
    If xColumn2Header = "" Then
       FrmSelectionList.Width = 6320
       FrmSelectionList.DataGrid1.Width = 6130
       FrmSelectionList.Text1.Width = 5650
       FrmSelectionList.DataGrid1.Columns(1).Visible = False
    Else
       FrmSelectionList.Width = 7710
       FrmSelectionList.DataGrid1.Width = 7530
       FrmSelectionList.Text1.Width = 7050
       FrmSelectionList.DataGrid1.Columns(1).Visible = True
       FrmSelectionList.DataGrid1.Columns(1).Caption = xColumn2Header
    End If
    If xColumn3Header = "" Then
        FrmSelectionList.DataGrid1.Columns(2).Visible = False
    Else
        FrmSelectionList.DataGrid1.Columns(2).Visible = True
        FrmSelectionList.DataGrid1.Columns(2).Caption = xColumn3Header
    End If
    If xColumn4Header = "" Then
        FrmSelectionList.DataGrid1.Columns(3).Visible = False
    Else
        FrmSelectionList.DataGrid1.Columns(3).Visible = True
        FrmSelectionList.DataGrid1.Columns(3).Caption = xColumn4Header
    End If
    FrmSelectionList.FindFieldName = "Col0"
    FrmSelectionList.Top = (Screen.Height - FrmSelectionList.Height) / 2
    FrmSelectionList.Left = Screen.Width - FrmSelectionList.Width - 130
    Load FrmSelectionList
End Sub
Public Sub DisplaySelectionList(ByRef xNameTextBox As TextBox, ByRef xCode As String)
    FrmSelectionList.Text1.Text = ""
    FrmSelectionList.TxtName = xNameTextBox.Text
    FrmSelectionList.TxtCode = xCode
    FrmSelectionList.Top = (Screen.Height - FrmSelectionList.Height) / 2
    FrmSelectionList.Left = Screen.Width - FrmSelectionList.Width - 130
    FrmSelectionList.Show vbModal
    xNameTextBox.Text = FrmSelectionList.TxtName
    xCode = FrmSelectionList.TxtCode
End Sub
Public Function CheckExists(ByVal xTextBox As TextBox, ByVal xField As String, ByVal xRecordset As Recordset, ByRef xCode As String) As Boolean
    xTextBox.Text = LTrim(RTrim(xTextBox.Text))
    If xRecordset.RecordCount > 0 Then
        xRecordset.MoveFirst
    End If
    xRecordset.Find "[" & xField & "] = '" & FixQuote(xTextBox.Text) & "'"
    If Not xRecordset.EOF Then
       CheckExists = True
       xCode = xRecordset.Fields("Code").Value
    End If
End Function
Public Function FixAPIString(ByVal strInput As String) As String
    ' strips Trailing Nulls From strings returned by API Function
    Dim intPosition As Integer
    
    intPosition = InStr(1, strInput, Chr(0))
    If intPosition Then
        FixAPIString = Left(strInput, intPosition - 1)
    Else
        FixAPIString = strInput     ' No Nulls Found
    End If
End Function
Public Function GetWindowsSystemDirectory() As String
    Dim DirectoryPath As String * 255
    
    Call GetSystemDirectory(DirectoryPath, 255)
    DirectoryPath = FixAPIString(DirectoryPath)
    GetWindowsSystemDirectory = RTrim(DirectoryPath)
End Function
Public Function GetWindowsTempDirectory() As String
    Dim DirectoryPath As String * 255
    Call GetTempPath(255, DirectoryPath)
    DirectoryPath = FixAPIString(DirectoryPath)
    GetWindowsTempDirectory = RTrim(DirectoryPath)
End Function
Public Sub WriteToFile(ByVal xKeyName As String, ByVal xString As String)
    WritePrivateProfileString "Saral", xKeyName, xString, App.Path + "\" + IIf(CheckEmpty(Command$, False), "Saral", Command$) + ".ini"
End Sub
Public Function ReadFromFile(ByVal xKeyName As String) As String
    Dim sReturn As String * 255
    GetPrivateProfileString "Saral", xKeyName, "", sReturn, 255, App.Path + "\" + IIf(CheckEmpty(Command$, False), "Saral", Command$) + ".ini"
    sReturn = FixAPIString(sReturn)
    ReadFromFile = RTrim(sReturn)
End Function
Public Sub BusySystemIndicator(ByVal bVal As Boolean)
    If bVal Then MdiMainMenu.MousePointer = vbHourglass Else MdiMainMenu.MousePointer = vbNormal
End Sub
Public Function NumberToWords(ByVal xNumber As Double, Optional ByVal blnPrintPaise As Boolean) As String
    Dim Amount As String, Paise As String
    Dim Crore As String, Lakh As String, Thousand As String, Hundred As String, Ten As String
    
    NumberToWords = "Rupees "
    Amount = Format(xNumber, "000000000.00")
    Paise = Format((xNumber - Int(xNumber)) * 100, "00")
    Crore = Mid(Amount, 1, 2)
    Lakh = Mid(Amount, 3, 2)
    Thousand = Mid(Amount, 5, 2)
    Hundred = Mid(Amount, 7, 1)
    Ten = Mid(Amount, 8, 2)
    If Val(Crore) > 0 Then
       NumberToWords = NumberToWords + Words(Crore) + " Crore "
    End If
    If Val(Lakh) > 0 Then
       NumberToWords = NumberToWords + Words(Lakh) + " Lakh "
    End If
    If Val(Thousand) > 0 Then
       NumberToWords = NumberToWords + Words(Thousand) + " Thousand "
    End If
    If Val(Hundred) > 0 Then
       NumberToWords = NumberToWords + Words(Hundred) + " Hundred "
    End If
    If Val(Ten) > 0 Then
       NumberToWords = NumberToWords + Words(Ten) + Space(1)
    End If
    If Val(Crore) = 0 And Val(Lakh) = 0 And Val(Thousand) = 0 And Val(Hundred) = 0 And Val(Ten) = 0 Then
       NumberToWords = NumberToWords + "Nil "
    End If
    If blnPrintPaise Then
       If Val(Paise) > 0 Then
          NumberToWords = NumberToWords + "And Paise "
          NumberToWords = NumberToWords + Words(Paise) + Space(1)
       Else
          'NumberToWords = NumberToWords + "And Paise Nil "
       End If
    End If
    NumberToWords = NumberToWords + "Only"
End Function
Public Function Words(ByVal xNumber As String) As String
    Const Ones = "One   Two   Three Four  Five  Six   Seven Eight Nine"
    Const Tens = "Ten     Twenty  Thirty  Forty   Fifty   Sixty   Seventy Eighty  Ninety"
    Const Teens = "Eleven    Twelve    Thirteen  Fourteen  Fifteen   Sixteen   Seventeen Eighteen  Nineteen"
    
    If Val(xNumber) >= 1 And Val(xNumber) <= 9 Then
       Words = Words + LTrim(RTrim(Mid(Ones, Val(xNumber) + (Val(xNumber) - 1) * 5, 6)))
    ElseIf Val(xNumber) >= 11 And Val(xNumber) <= 19 Then
       Words = Words + LTrim(RTrim(Mid(Teens, Val(Mid(xNumber, 2, 1)) + (Val(Mid(xNumber, 2, 1)) - 1) * 9, 10)))
    ElseIf (Val(xNumber) >= 20 And Val(xNumber) <= 99) Or Val(xNumber) = 10 Then
       Words = Words + LTrim(RTrim(Mid(Tens, Val(Mid(xNumber, 1, 1)) + (Val(Mid(xNumber, 1, 1)) - 1) * 7, 8)))
       If Mid(xNumber, 2, 1) <> "0" Then
          Words = Words + Space(1) + LTrim(RTrim(Mid(Ones, Val(Mid(xNumber, 2, 1)) + (Val(Mid(xNumber, 2, 1)) - 1) * 5, 6)))
       End If
    End If
End Function
Public Function FileExist(ByVal szFileName As String) As Boolean
    Dim nFileNumber As Integer
    On Error Resume Next
    
    nFileNumber = FreeFile
    Open szFileName For Input As nFileNumber
    If Err.Number = 0 Then
        FileExist = True
    End If
    Close nFileNumber
    Err.Clear
End Function
Public Function RestorePreviousInstance(ByVal strClass As String, ByVal strPreviousTitle As String) As Boolean
  Dim lngHandle As Long
  Dim WinLocation  As WINDOWPLACEMENT
  Dim lngRetVal   As Long
     
' VB6 uses class name "ThunderRT6FormDC"
' Including the class name for the compiled EXE class prevents the routine
' from finding and attempting to activate the project form of the same name.
  lngHandle = FindWindow(strClass, strPreviousTitle)
  DoEvents
' If application is already executing
  If lngHandle > 0 Then
      ' Get the current window state of the previous instance
     WinLocation.Length = Len(WinLocation)
     lngRetVal = GetWindowPlacement(lngHandle, WinLocation)
     ' if the WinLocation.showCmd member indicates that
     ' the window is currently minimized, it needs
     ' to be restored.
     If WinLocation.showCmd = SW_SHOWMINIMIZED Then
        With WinLocation
                .Length = Len(WinLocation)
                .flags = 0&
                .showCmd = SW_SHOWNORMAL
        End With
        lngRetVal = SetWindowPlacement(lngHandle, WinLocation)
      End If
      ' Bring the window to the front and make the active window.
      ' Without this, it may remain hidden behind other windows.
      lngRetVal = SetForegroundWindow(lngHandle)
      DoEvents
      RestorePreviousInstance = True
  End If
End Function
Public Function RestoreForm(ByVal wHandle As Long) As Boolean
  If IsIconic(wHandle) Then
    Call PostMessage(wHandle, WM_SYSCOMMAND, SC_RESTORE, 0)
    RestoreForm = True
  End If
End Function
Public Sub SetComboBoxDroppedWidth(ByVal xForm As Form, ByVal xComboBox As ComboBox, ByVal NumItemsToDisplay As Integer, ByVal DroppedWidth As Integer, ByVal ShowDropDown As Boolean)
    Dim pt As POINTAPI
    Dim rc As RECT
    Dim cWidth As Long
    Dim newHeight As Long
    Dim oldScaleMode As Long
    Dim itemHeight As Long
    
    If TypeOf xComboBox.Parent Is Frame Then Exit Sub
    oldScaleMode = xForm.ScaleMode
    xForm.ScaleMode = vbPixels
    cWidth = xComboBox.Width
    itemHeight = SendMessage(xComboBox.hwnd, CB_GETITEMHEIGHT, 0, ByVal 0)
    newHeight = itemHeight * (NumItemsToDisplay + 2)
    Call GetWindowRect(xComboBox.hwnd, rc)
    pt.x = rc.Left
    pt.Y = rc.Top
    Call SendMessage(xComboBox.hwnd, CB_SETDROPPEDWIDTH, ByVal DroppedWidth, ByVal 0)
    Call ScreenToClient(xForm.hwnd, pt)
    Call MoveWindow(xComboBox.hwnd, pt.x, pt.Y, xComboBox.Width, newHeight, True)
    Call SendMessage(xComboBox.hwnd, CB_SHOWDROPDOWN, ShowDropDown, ByVal 0)
    xForm.ScaleMode = oldScaleMode
End Sub
Public Function TimeDiff(STime As Date, ETime As Date) As String
    'Example : TimeDiff(Time, "05:45:00 PM")
    Dim TimeSecs, Hrs As Double
    Dim strSeconds As String
    Dim strMinutes As String
    Dim strHours As String
    If ETime < STime Then Exit Function
    'Get Total Number of seconds difference
    TimeSecs = DateDiff("S", STime, ETime)
    strHours = Int(TimeSecs / 3600)
    strMinutes = Int((TimeSecs Mod 3600) / 60)
    strSeconds = (TimeSecs Mod 3600) Mod 60
    TimeDiff = IIf(Len(strHours) = 1, String(2 - Len(strHours), "0") + strHours, strHours) + ":" + String(2 - Len(strMinutes), "0") + strMinutes + ":" + String(2 - Len(strSeconds), "0") + strSeconds
End Function
Public Function ProperCase(ByVal strInput As String) As String
     ProperCase = StrConv(strInput, vbProperCase)
End Function
Public Sub ShowProgressInStatusBar(ByVal bShowProgressBar As Boolean)
    Dim tRC As RECT

    If bShowProgressBar Then
        'Get the size of the Panel Rectangle from the status Bar
        SendMessage MdiMainMenu.StatusBar1.hwnd, SB_GETRECT, 0, tRC
        'And convert it to twips....
        With tRC
            .Top = (.Top * Screen.TwipsPerPixelY)
            .Left = (.Left * Screen.TwipsPerPixelX)
            .Bottom = (.Bottom * Screen.TwipsPerPixelY) - .Top
            .Right = (.Right * Screen.TwipsPerPixelX) - .Left
        End With
        'Now Reparent the ProgressBar to the statusbar
        With MdiMainMenu.ProgressBar1
            SetParent .hwnd, MdiMainMenu.StatusBar1.hwnd
            .Move tRC.Left, tRC.Top, tRC.Right, tRC.Bottom
            .Visible = True
            .Value = 0
        End With
    Else
        'Reparent the progress bar back to the form and hide it
        SetParent MdiMainMenu.ProgressBar1.hwnd, MdiMainMenu.Picture1.hwnd
        MdiMainMenu.ProgressBar1.Visible = False
    End If
End Sub
Public Function FixQuote(ByVal strInput As String) As String
    If InStr(1, strInput, "'") > 0 Then
       FixQuote = Replace(strInput, "'", "''")
    Else
       FixQuote = strInput
    End If
End Function
Public Sub CloseMsgBox(ByVal MsgTitle As String)
   Static hwnd As Long
   Static Ticks As Long
   
   If hwnd = 0 Then
      hwnd = FindWindow(vbNullString, MsgTitle)
   End If
   Ticks = Ticks + 1
   Call SendMessage(hwnd, WM_SETTEXT, 0, ByVal MsgTitle)
   If Ticks >= 5 Then
      Call SendMessage(hwnd, WM_CLOSE, 0, ByVal 0&)
      hwnd = 0
      Ticks = 0
   End If
End Sub
Public Sub DisplayError(ByVal strErrorMsg As String)
    On Error Resume Next
    Beep
    MsgBox RTrim(LTrim(strErrorMsg)) & " !!!", vbExclamation, "Error !"
    Err.Clear
End Sub
Public Sub CloseForm(ByRef xForm As Form)
    On Error GoTo ErrorHandler
    Unload xForm
    Set xForm = Nothing
    Exit Sub
ErrorHandler:
End Sub
Public Function GenerateCode(ByVal xConnection As ADODB.Connection, ByVal strSQL As String, intLen, ByVal strFillChar As String) As Variant
    On Error GoTo ErrorHandler
    Dim Rs As New ADODB.Recordset
    Dim xCode As String
    
    Rs.Open strSQL, xConnection, adOpenKeyset, adLockReadOnly
    If IsNull(Rs.Fields(0).Value) Then
       xCode = "0"
    Else
       xCode = Val(Rs.Fields(0).Value)
    End If
    GenerateCode = Pad(RTrim(Val(xCode) + 1), strFillChar, intLen, "L")
    Rs.Close
    Set Rs = Nothing
    Exit Function
ErrorHandler:
    If Rs.State = adStateOpen Then
       Rs.Close
    End If
    Set Rs = Nothing
    GenerateCode = Null
End Function
Public Function CheckDuplicate(ByVal xConnection As ADODB.Connection, ByVal TableName As String, ByVal SelectField As String, ByVal SearchField As String, ByVal SearchValue As String, Optional ByVal CheckValue As Variant, Optional ByVal AskToContinue As Boolean) As Boolean
    On Error GoTo ErrorHandler
    Dim Rs As New ADODB.Recordset
    
    Rs.Open "Select " & SelectField & " From " & TableName & " Where LTrim(RTrim(" & SearchField & ")) = '" & FixQuote(RTrim(LTrim(SearchValue))) & "'", xConnection, adOpenKeyset, adLockReadOnly
    If Rs.RecordCount <> 0 Then
       If (CheckValue <> Rs.Fields(0).Value) Or CheckEmpty(CheckValue, False) Then
          If AskToContinue Then
              Beep
              If MsgBox("      Duplicate Entry !" & vbCrLf & "Would you like to continue ?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirm Proceed !") = vbYes Then
                 CheckDuplicate = False
              Else
                 CheckDuplicate = True
              End If
          Else
              DisplayError ("Duplicate Entry")
              CheckDuplicate = True
          End If
       End If
    End If
    Rs.Close
    Set Rs = Nothing
    Exit Function
ErrorHandler:
    If Rs.State = adStateOpen Then
       Rs.Close
    End If
    Set Rs = Nothing
    CheckDuplicate = True
End Function
Public Function AddRecord(ByRef Rs As ADODB.Recordset) As Boolean
    On Error GoTo ErrorHandler
    Rs.AddNew
    AddRecord = True
    Exit Function
ErrorHandler:
End Function
Public Sub DeleteRecord(ByRef Rs As ADODB.Recordset, ByVal xCode As String)
    On Error GoTo ErrorHandler
    
    If Rs.EOF And Rs.BOF Then Exit Sub
    
    If MsgBox("Are you sure to delete the Record?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") <> vbYes Then
        Exit Sub
    End If
    
    Rs.MoveFirst
    Rs.Find "[Code] ='" & FixQuote(xCode) & "'"
    If Not Rs.EOF Then
        MdiMainMenu.MousePointer = vbHourglass
        Rs.Delete
        Rs.MoveNext
    End If
    MdiMainMenu.MousePointer = vbNormal
    Exit Sub
ErrorHandler:
    DisplayError ("Failed to delete the record")
    Rs.CancelUpdate
    MdiMainMenu.MousePointer = vbNormal
End Sub
Public Function UpdateRecord(ByRef Rs As ADODB.Recordset) As Boolean
    Dim strErrorMessage As String, oError As Error, blnAdd As Boolean
    On Error Resume Next 'This also clears the Error object
    
    UpdateRecord = True 'No Error
    blnAdd = Rs.EditMode = adEditAdd
    Rs.ActiveConnection.Errors.Clear
    Screen.MousePointer = vbHourglass 'The update might take a while
    Rs.Update
    Screen.MousePointer = vbNormal
    Select Case Err.Number
        Case 0 'Check the underlying Connection Object for Errors too. Provider-specific Errors don't show up in the above Error Trap
            If Rs.ActiveConnection.Errors.Count = 0 Then 'No provider-specific errors - save was successful
               If blnAdd Then
                  If Rs.CursorLocation = adUseClient Then
                     Rs.Resync adAffectCurrent 'show default field values that may have been entered by the database
                  End If
               End If
            ElseIf Rs.ActiveConnection.Errors.Count <> 0 Then
               For Each oError In Rs.ActiveConnection.Errors
                     strErrorMessage = strErrorMessage & oError.Description & vbCr
               Next
               If Rs.ActiveConnection.Errors.Count >= 1 Then strErrorMessage = "The following error" & IIf(Rs.ActiveConnection.Errors.Count > 1, "s were", " was") & " reported by the provider: " & vbCr & strErrorMessage
               Call DisplayError(strErrorMessage) 'display all the errors
               UpdateRecord = False 'Error
                'Leave the save and Cancel Buttons showing so that user can backtrack
            End If
        Case Else
           UpdateRecord = False 'Error
    End Select
End Function
Public Function CancelRecordUpdate(ByRef Rs As ADODB.Recordset) As Boolean
    On Error GoTo ErrorHandler
    
    If Rs.EditMode = adEditAdd Or Rs.EditMode = adEditInProgress Then
        Rs.CancelUpdate
    End If
    CancelRecordUpdate = True
    Exit Function
ErrorHandler:
End Function
Public Function RefreshRecord(ByRef Rs As ADODB.Recordset) As Boolean
    On Error GoTo ErrorHandler
    ShowProgressInStatusBar (True)
    MdiMainMenu.ProgressBar1.Value = 50
    Screen.MousePointer = vbHourglass
    Rs.Requery
    MdiMainMenu.ProgressBar1.Value = 100
    ShowProgressInStatusBar (False)
    Screen.MousePointer = vbNormal
    RefreshRecord = True
  Exit Function
ErrorHandler:
    Screen.MousePointer = vbNormal
End Function
Function Pad(ByVal strExpression As String, ByVal strFillChar As String, ByVal intLength As Integer, ByVal strAlignment As String) As String
    
    If Len(strExpression) > intLength Then
       strExpression = Left(strExpression, intLength)
    End If
    If StrConv(strAlignment, vbUpperCase) = "C" Then
       Dim lPad, rPad As Integer
       lPad = Int((intLength - Len(strExpression)) / 2)
       rPad = intLength - Len(strExpression) - lPad
       Pad = String(lPad, strFillChar) & strExpression & String(rPad, strFillChar)
    ElseIf StrConv(strAlignment, vbUpperCase) = "R" Then
        Pad = strExpression & String(intLength - Len(strExpression), strFillChar)
    ElseIf StrConv(strAlignment, vbUpperCase) = "L" Then
        Pad = String(intLength - Len(strExpression), strFillChar) & strExpression
    End If
    
End Function
Public Sub CloseRecordset(ByRef xRecordset As ADODB.Recordset)
    On Error GoTo ErrorHandler
    If xRecordset.State = adStateOpen Then
       xRecordset.Close
    End If
    Set xRecordset = Nothing
    Exit Sub
    
ErrorHandler:
End Sub
Public Sub CloseConnection(ByVal xConnection As ADODB.Connection)
    On Error GoTo ErrorHandler
    If xConnection.State = adStateOpen Then
       xConnection.Close
    End If
    Set xConnection = Nothing
    Exit Sub
ErrorHandler:
End Sub
Public Function DirExist(ByVal strDir As String) As Boolean
    On Error Resume Next
    ChDir strDir
    If Err.Number <> 76 Then DirExist = True
End Function
Public Sub SetMdiButtons(bVal As Boolean, Optional ByVal EnablePrintButtons As Boolean, Optional ByVal EnableMailButton As Boolean)
    Dim Ctr As Integer
    For Ctr = 1 To 17
        MdiMainMenu.Toolbar1.Buttons(Ctr).Enabled = bVal
        If Ctr = 9 Or Ctr = 10 Then
            MdiMainMenu.Toolbar1.Buttons(8).Visible = EnablePrintButtons
            MdiMainMenu.Toolbar1.Buttons(Ctr).Visible = EnablePrintButtons
            MdiMainMenu.Toolbar1.Buttons(Ctr).Enabled = EnablePrintButtons
        End If
    Next
    MdiMainMenu.Toolbar1.Buttons(11).Visible = EnableMailButton
    MdiMainMenu.Toolbar1.Buttons(11).Enabled = EnableMailButton
End Sub
Public Sub EnableChildMenu(Optional ByVal EnablePrintButtons As Boolean, Optional ByVal EnableMailButton As Boolean)
    MdiMainMenu.Toolbar1.Buttons(18).ToolTipText = "Close"
    Call SetMdiButtons(True, EnablePrintButtons, EnableMailButton)
End Sub
Public Sub DisableChildMenu()
    Call SetMdiButtons(False)
    MdiMainMenu.Toolbar1.Buttons(18).ToolTipText = "Exit"
End Sub
Public Sub CenterForm(frm As Form)
  Dim m_lngRetVal As Long
  Dim ClientRect As RECT     'Holds the area that the form is to be centered in
  Dim TaskBarRect As RECT     'Holds the TaskBar area if in Win95
  Dim x As Variant  'temp LeftPosition
  Dim Y As Variant  'temp TopPosition

  If frm.MDIChild Then ' Check if the form is a MDIChild.
      ' Center it in the MDIParent.
      GetClientRect GetParent(frm.hwnd), ClientRect
  Else  'Center it in the available desktop area.
      ' Get the Desktop area
      Call GetClientRect(GetDesktopWindow(), ClientRect)
      ' Check for the Task Bar.
      m_lngRetVal = FindWindow("Shell_TrayWnd", vbNullString)
      ' If there is a taskbar, ie WIN95 then adjust the ClientRect.
      If m_lngRetVal Then
          Call GetWindowRect(m_lngRetVal, TaskBarRect)
          If (TaskBarRect.Right - TaskBarRect.Left) > (TaskBarRect.Bottom - TaskBarRect.Top) Then
              ' TaskBar at the Top of Screen.
              If TaskBarRect.Top <= 0 Then
                  ClientRect.Top = ClientRect.Top + TaskBarRect.Bottom
                  ' TaskBar at the Bottom of Screen.
              Else
                  ClientRect.Bottom = ClientRect.Bottom - (TaskBarRect.Bottom - TaskBarRect.Top)
              End If
          Else
              ' TaskBar is on the Left side of the Screen.
              If TaskBarRect.Left <= 0 Then
                  ClientRect.Left = ClientRect.Left + TaskBarRect.Right
                  ' TaskBar is on the Right side of the Screen.
              Else
                  ClientRect.Right = ClientRect.Right - (TaskBarRect.Right - TaskBarRect.Left)
              End If
          End If   '[TaskBar on Top of Screen?]
      End If
  End If
' Center the Form
  With frm
       x = (((ClientRect.Right - ClientRect.Left) * Screen.TwipsPerPixelX) - .Width) / 2
       Y = (((ClientRect.Bottom - ClientRect.Top) * Screen.TwipsPerPixelY) - .Height) / 2
       .Move x, Y
  End With
End Sub
Public Function SetPrinterMode(ByVal strCode As String, ByVal blnFlag As Boolean) As String
    If strCode = "cndn" Then
       If blnFlag Then
          SetPrinterMode = Chr(15)
      Else
          SetPrinterMode = Chr(18)
      End If
    ElseIf strCode = "12cpi" Then
       If blnFlag Then
          SetPrinterMode = Chr(27) & Chr(77)
      Else
          SetPrinterMode = Chr(27) & Chr(80)
      End If
    ElseIf strCode = "15cpi" Then
       If blnFlag Then
          SetPrinterMode = Chr(27) & Chr(103)
      Else
          SetPrinterMode = Chr(27) & Chr(80)
      End If
    ElseIf strCode = "19cpi" Then '12CPI Condensed
       If blnFlag Then
          SetPrinterMode = Chr(15) & Chr(27) & Chr(77)
      Else
          SetPrinterMode = Chr(15) & Chr(27) & Chr(80)
      End If
    ElseIf strCode = "bold" Then
       If blnFlag Then
          SetPrinterMode = Chr(27) & Chr(69)
      Else
          SetPrinterMode = Chr(27) & Chr(70)
      End If
      
    ElseIf strCode = "dblstrk" Then
       If blnFlag Then
          SetPrinterMode = Chr(27) & Chr(71)
      Else
          SetPrinterMode = Chr(27) & Chr(72)
      End If
    ElseIf strCode = "ulin" Then
       If blnFlag Then
          SetPrinterMode = Chr(27) & Chr(45) & Chr(49)
      Else
          SetPrinterMode = Chr(27) & Chr(45) & Chr(48)
      End If
    ElseIf strCode = "dwth" Then
       If blnFlag Then
          SetPrinterMode = Chr(27) & Chr(87) & Chr(49)
      Else
          SetPrinterMode = Chr(27) & Chr(87) & Chr(48)
      End If
    ElseIf strCode = "8lpi" Then
       If blnFlag Then
          SetPrinterMode = Chr(27) & Chr(48)
      Else
          SetPrinterMode = Chr(27) & Chr(50)
      End If
    ElseIf strCode = "init" Then
          SetPrinterMode = Chr(27) & Chr(64)
    ElseIf strCode = "ejec" Then
          SetPrinterMode = Chr(12)
    End If
    
End Function
Public Function GetTemporaryFileName() As String
    Dim lngReturnVal As Long
    Dim strTempPath As String * 255
    Dim strTempFileName As String * 255
    On Error GoTo TempNameErr
    lngReturnVal = GetTempPath(254, strTempPath)
    lngReturnVal = GetTempFileName(strTempPath & "\", "", 0, strTempFileName)
    GetTemporaryFileName = strTempFileName
    Exit Function
TempNameErr:
    Call DisplayError("Cannot retrieve Temporary FileName")
End Function
Public Function DisplayPopupMenu(ByVal hwnd As Long, Optional ByVal intMenu As Integer) As Integer
    Dim hPopupMenu1 As Long ' Handle to the popup menu to display
    Dim mii1 As MENUITEMINFO ' Describes menu items to add
    Dim CurPos As POINT_TYPE ' Holds the current mouse coordinates
    Dim menusel As Long ' ID of what the user selected in the popup menu
    Dim retVal As Long ' Generic return value
    Dim strOption As String
    Dim Cnt As Integer
    'Create the popup menus which are initialy empty.
    hPopupMenu1 = CreatePopupMenu()
    'Create the structure which is the base for all Menus:
    With mii1
        .cbSize = Len(mii1) ' The size of this structure.
        .fMask = MIIM_STATE Or MIIM_ID Or MIIM_TYPE Or MIIM_SUBMENU ' Which elements of the structure to use.
    End With
    If intMenu = 0 Then
        For Cnt = 0 To 2
            With mii1
                .fType = MFT_STRING
                .fState = MFS_ENABLED
                .wID = 1006 - Cnt
                strOption = IIf(Cnt = 0, "Add Record", IIf(Cnt = 1, "Edit Record", "Delete Record"))
                .dwTypeData = strOption
                .cch = Len(strOption)
                .hSubMenu = 0
            End With
            retVal = InsertMenuItem(hPopupMenu1, Cnt, 1, mii1)
        Next
    ElseIf intMenu = 1 Then
        For Cnt = 0 To 5
            With mii1
                .fType = MFT_STRING
                .fState = MFS_ENABLED
                .wID = 1006 - Cnt
                strOption = Choose(Cnt + 1, "Book Printing Order", "Title Printing Order", "Title Lamination Order", "Book Binding Order", "Book Order", "Box Label")
                .dwTypeData = strOption
                .cch = Len(strOption)
                .hSubMenu = 0
            End With
            retVal = InsertMenuItem(hPopupMenu1, Cnt, 1, mii1)
        Next
    ElseIf intMenu = 2 Then
        For Cnt = 0 To 1
            With mii1
                .fType = MFT_STRING
                .fState = MFS_ENABLED
                .wID = 1006 - Cnt
                strOption = Choose(Cnt + 1, "Purchase Order", "Issue Voucher")
                .dwTypeData = strOption
                .cch = Len(strOption)
                .hSubMenu = 0
            End With
            retVal = InsertMenuItem(hPopupMenu1, Cnt, 1, mii1)
        Next
    ElseIf intMenu = 3 Then
        For Cnt = 0 To 2
            With mii1
                .fType = MFT_STRING
                .fState = MFS_ENABLED
                .wID = 1006 - Cnt
                strOption = Choose(Cnt + 1, "Original Delivery Challan", "Duplicate Devlivery Challan", "Triplicate Delivery Challan")
                .dwTypeData = strOption
                .cch = Len(strOption)
                .hSubMenu = 0
            End With
            retVal = InsertMenuItem(hPopupMenu1, Cnt, 1, mii1)
        Next
    End If
    retVal = GetCursorPos(CurPos)
    menusel = TrackPopupMenu(hPopupMenu1, TPM_TOPALIGN Or TPM_NONOTIFY Or TPM_RETURNCMD Or TPM_LEFTALIGN Or TPM_LEFTBUTTON, CurPos.x, CurPos.Y, 0, hwnd, 0)
    retVal = DestroyMenu(hPopupMenu1)
    DisplayPopupMenu = 1007 - menusel
End Function
Public Function CalculateConsumption(ByVal xPaperType As String, ByVal xQuantity As Long, ByVal xForms As Double, ByVal xWastage As Double) As Double
    If xPaperType = "1" Then    'Book
        CalculateConsumption = CLng(xQuantity * xForms * (100 + xWastage) / 100)
    Else    'Title
        CalculateConsumption = Format((xQuantity / 2) * ((100 + xWastage) / 100), "#0")
    End If
    CalculateConsumption = CLng(Val(CalculateConsumption) / 2)
    CalculateConsumption = Int(Val(CalculateConsumption) / 500) & "." & Format(Val(CalculateConsumption) Mod 500, "000")
End Function
Public Function CalculatePaperBalance(ByVal strAccountCode As String, ByVal strPaperCode As String, ByVal strVoucherCode As String, ByVal strVoucherType) As Long
    Dim rstPaperBalance As New ADODB.Recordset
    On Error GoTo ErrorHandler
    If rstPaperBalance.State = adStateOpen Then rstPaperBalance.Close
    If rstPaperBalance.State = adStateOpen Then rstPaperBalance.Close
    If InStr(1, "PMVT_PMVB", strVoucherType) > 0 Then   'Paper Movement
        rstPaperBalance.Open "SELECT FORMAT((SELECT SUM(QuantitySheets) FROM PaperIOChild WHERE Paper=M.Code AND Account='" & strAccountCode & "'),0) As Col0," & _
                                             "FORMAT((SELECT SUM(INT(Quantity)*500+(Quantity-INT(Quantity))*1000) FROM MaterialSVParent P INNER JOIN MaterialSVChild C ON P.Code=C.Code WHERE Category='2' AND Item=M.Code AND Quantity>=0 AND Account='" & strAccountCode & "'),0) As Col1," & _
                                             "FORMAT((SELECT ABS(SUM(FIX(Quantity)*500+(Quantity-FIX(Quantity))*1000)) FROM MaterialSVParent P INNER JOIN MaterialSVChild C ON P.Code=C.Code WHERE Category='2' AND Item=M.Code AND Quantity<0 AND Account='" & strAccountCode & "'),0) As Col2," & _
                                             "FORMAT((SELECT SUM(QuantitySheets) FROM PaperMVParent P INNER JOIN PaperMVChild C ON P.Code=C.Code WHERE Paper=M.Code AND AccountFrom='" & strAccountCode & "' AND P.Code<>'" & strVoucherCode & "'),0) As Col3," & _
                                             "FORMAT((SELECT SUM(QuantitySheets) FROM PaperMVParent P INNER JOIN PaperMVChild C ON P.Code=C.Code WHERE Paper=M.Code AND AccountTo='" & strAccountCode & "' AND P.Code<>'" & strVoucherCode & "'),0) As Col4," & _
                                             "FORMAT((SELECT OpBalSheets FROM PaperChild WHERE Code=M.Code AND Account='" & strAccountCode & "'),0) As Col5," & _
                                             "FORMAT((SELECT SUM(PaperConsumptionSheets1) FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE P.Type<>'O' AND LEFT(P.Code,1)<>'*' AND Paper1=M.Code AND BookPrinter='" & strAccountCode & "'),0) As Col6," & _
                                             "FORMAT((SELECT SUM(PaperConsumptionSheets2) FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE P.Type<>'O' AND LEFT(P.Code,1)<>'*' AND Paper2=M.Code AND BookPrinter='" & strAccountCode & "'),0) As Col7," & _
                                             "FORMAT((SELECT SUM(PaperConsumptionSheets4) FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE P.Type<>'O' AND LEFT(P.Code,1)<>'*' AND Paper4=M.Code AND BookPrinter='" & strAccountCode & "'),0) As Col8," & _
                                             "FORMAT((SELECT SUM(PaperConsumptionSheets) FROM BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code WHERE P.Type<>'O' AND LEFT(P.Code,1)<>'*' AND Paper=M.Code AND TitlePrinter='" & strAccountCode & "'),0) As Col9," & _
                                             "FORMAT((SELECT SUM(ROUND(ActualQuantity*C2.Quantity,0)) FROM (BookPOParent P INNER JOIN BookPOChild08 C1 ON P.Code=C1.Code) INNER JOIN BookPOChild0801 C2 ON C1.Code=C2.Code WHERE P.Type<>'O' AND LEFT(P.Code,1)<>'*' AND C2.Category='2' AND C2.Item=M.Code AND BookPrinter='" & strAccountCode & "'),0) As Col10," & _
                                             "FORMAT((SELECT INT(Quantity)*500+(Quantity-INT(Quantity))*1000 FROM PaperDNParent P INNER JOIN PaperDNChild C ON P.Code=C.Code WHERE P.Account='" & strAccountCode & "' AND C.Paper=M.Code),0) As Col11 " & _
                                             "FROM PaperMaster M WHERE Code='" & strPaperCode & "'", CxnDatabase, adOpenKeyset, adLockReadOnly
    ElseIf InStr(1, "BPOT_BPOB", strVoucherType) > 0 Then   'Print Order
        rstPaperBalance.Open "SELECT FORMAT((SELECT SUM(QuantitySheets) FROM PaperIOChild WHERE Paper=M.Code AND Account='" & strAccountCode & "'),0) As Col0," & _
                                             "FORMAT((SELECT SUM(INT(Quantity)*500+(Quantity-INT(Quantity))*1000) FROM MaterialSVParent P INNER JOIN MaterialSVChild C ON P.Code=C.Code WHERE Category='2' AND Item=M.Code AND Quantity>=0 AND Account='" & strAccountCode & "'),0) As Col1," & _
                                             "FORMAT((SELECT ABS(SUM(FIX(Quantity)*500+(Quantity-FIX(Quantity))*1000)) FROM MaterialSVParent P INNER JOIN MaterialSVChild C ON P.Code=C.Code WHERE Category='2' AND Item=M.Code AND Quantity<0 AND Account='" & strAccountCode & "'),0) As Col2," & _
                                             "FORMAT((SELECT SUM(QuantitySheets) FROM PaperMVParent P INNER JOIN PaperMVChild C ON P.Code=C.Code WHERE Paper=M.Code AND AccountFrom='" & strAccountCode & "'),0) As Col3," & _
                                             "FORMAT((SELECT SUM(QuantitySheets) FROM PaperMVParent P INNER JOIN PaperMVChild C ON P.Code=C.Code WHERE Paper=M.Code AND AccountTo='" & strAccountCode & "'),0) As Col4," & _
                                             "FORMAT((SELECT OpBalSheets FROM PaperChild WHERE Code=M.Code AND Account='" & strAccountCode & "'),0) As Col5," & _
                                             "FORMAT((SELECT SUM(PaperConsumptionSheets1) FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE P.Type<>'O' AND LEFT(P.Code,1)<>'*' AND Paper1=M.Code AND BookPrinter='" & strAccountCode & "' AND P.Code<>'" & strVoucherCode & "'),0) As Col6," & _
                                             "FORMAT((SELECT SUM(PaperConsumptionSheets2) FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE P.Type<>'O' AND LEFT(P.Code,1)<>'*' AND Paper2=M.Code AND BookPrinter='" & strAccountCode & "' AND P.Code<>'" & strVoucherCode & "'),0) As Col7," & _
                                             "FORMAT((SELECT SUM(PaperConsumptionSheets4) FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE P.Type<>'O' AND LEFT(P.Code,1)<>'*' AND Paper4=M.Code AND BookPrinter='" & strAccountCode & "' AND P.Code<>'" & strVoucherCode & "'),0) As Col8," & _
                                             "FORMAT((SELECT SUM(PaperConsumptionSheets) FROM BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code WHERE P.Type<>'O' AND LEFT(P.Code,1)<>'*' AND Paper=M.Code AND TitlePrinter='" & strAccountCode & "' AND P.Code<>'" & strVoucherCode & "'),0) As Col9," & _
                                             "FORMAT((SELECT SUM(ROUND(ActualQuantity*C2.Quantity,0)) FROM (BookPOParent P INNER JOIN BookPOChild08 C1 ON P.Code=C1.Code) INNER JOIN BookPOChild0801 C2 ON C1.Code=C2.Code WHERE P.Type<>'O' AND LEFT(P.Code,1)<>'*' AND C2.Category='2' AND C2.Item=M.Code AND BookPrinter='" & strAccountCode & "' AND P.Code<>'" & strVoucherCode & "'),0) As Col10," & _
                                             "FORMAT((SELECT INT(Quantity)*500+(Quantity-INT(Quantity))*1000 FROM PaperDNParent P INNER JOIN PaperDNChild C ON P.Code=C.Code WHERE P.Account='" & strAccountCode & "' AND C.Paper=M.Code),0) As Col11 " & _
                                             "FROM PaperMaster M WHERE Code='" & strPaperCode & "'", CxnDatabase, adOpenKeyset, adLockReadOnly
    End If
    
    If rstPaperBalance.RecordCount > 0 Then
        CalculatePaperBalance = Val(CheckNull(rstPaperBalance.Fields("Col0").Value)) + Val(CheckNull(rstPaperBalance.Fields("Col1").Value)) - Val(CheckNull(rstPaperBalance.Fields("Col2").Value)) - Val(CheckNull(rstPaperBalance.Fields("Col3").Value)) + Val(CheckNull(rstPaperBalance.Fields("Col4").Value)) + Val(CheckNull(rstPaperBalance.Fields("Col5").Value)) - Val(CheckNull(rstPaperBalance.Fields("Col6").Value)) - Val(CheckNull(rstPaperBalance.Fields("Col7").Value)) - Val(CheckNull(rstPaperBalance.Fields("Col8").Value)) - Val(CheckNull(rstPaperBalance.Fields("Col9").Value)) - Val(CheckNull(rstPaperBalance.Fields("Col10").Value)) - Val(CheckNull(rstPaperBalance.Fields("Col11").Value))
    Else
        CalculatePaperBalance = 0
    End If
    
    Call CloseRecordset(rstPaperBalance)
    Exit Function
    
ErrorHandler:
    Call CloseRecordset(rstPaperBalance)
End Function
Public Function CalculateMaterialBalance(ByVal strAccountCode As String, ByVal strCategory As String, ByVal strItemCode As String, ByVal strVoucherCode As String, ByVal strVoucherType) As Double
    Dim rstMaterialBalance As New ADODB.Recordset
    Dim Category As String
    On Error GoTo ErrorHandler
    If rstMaterialBalance.State = adStateOpen Then rstMaterialBalance.Close
    Category = IIf(strCategory = "Outsource Item", "1", IIf(strCategory = "Fresh Book", "3", IIf(strCategory = "Repair Book", "4", "5")))
    rstMaterialBalance.Open "SELECT FORMAT((SELECT SUM(Quantity) FROM MaterialIOChild WHERE Category='" & Category & "' AND Item=M.Code AND Godown='" & strAccountCode & "'),'0.000') AS Col0,FORMAT((SELECT SUM(Quantity) FROM MaterialSVParent,MaterialSVChild WHERE MaterialSVParent.Code=MaterialSVChild.Code AND Quantity>=0 AND Category='" & Category & "' AND Item=M.Code AND Account='" & strAccountCode & "'),'0.000') AS Col1,FORMAT((SELECT SUM(ABS(Quantity)) FROM MaterialSVParent,MaterialSVChild WHERE MaterialSVParent.Code=MaterialSVChild.Code AND Quantity<0 AND Category='" & Category & "' AND Item=M.Code AND Account='" & strAccountCode & "'),'0.000') AS Col2,FORMAT((SELECT SUM(Quantity) FROM MaterialMVParent,MaterialMVChild WHERE MaterialMVParent.Code=MaterialMVChild.Code AND Category='" & Category & "' AND Item=M.Code AND AccountFrom='" & strAccountCode & "' AND " & IIf(strVoucherType = "MV", "MaterialMVParent.Code<>'" & strVoucherCode & "'", "1") & "),'0.000') AS Col3," & _
                                              "FORMAT((SELECT SUM(Quantity) FROM MaterialMVParent,MaterialMVChild WHERE MaterialMVParent.Code=MaterialMVChild.Code AND Category='" & Category & "' AND Item=M.Code AND AccountTo='" & strAccountCode & "' AND " & IIf(strVoucherType = "MV", "MaterialMVParent.Code<>'" & strVoucherCode & "'", "1") & "),'0.000') AS Col4,FORMAT((SELECT OpBal FROM AccountChild0801 WHERE Category='" & Category & "' AND Item=M.Code AND Code='" & strAccountCode & "'),'0.000') AS Col5,FORMAT((SELECT SUM(Quantity*(SELECT ActualQuantity FROM BookPOChild08 WHERE Code=BookPOParent.Code)) FROM BookPOParent,BookPOChild0801 WHERE BookPOParent.Code=BookPOChild0801.Code AND Category='" & Category & "' AND Item=M.Code AND Binder='" & strAccountCode & "' AND " & IIf(strVoucherType = "PO", "BookPOParent.Code<>'" & strVoucherCode & "'", "1") & "),'0.000') AS Col6 From " & IIf(Category = "1", "OutsourceItemMaster", "BookMaster") & " M " & _
                                              "WHERE Code='" & strItemCode & "'", CxnDatabase, adOpenKeyset, adLockReadOnly
    If rstMaterialBalance.RecordCount > 0 Then
        CalculateMaterialBalance = Val(CheckNull(rstMaterialBalance.Fields("Col0").Value)) + Val(CheckNull(rstMaterialBalance.Fields("Col1").Value)) - Val(CheckNull(rstMaterialBalance.Fields("Col2").Value)) - Val(CheckNull(rstMaterialBalance.Fields("Col3").Value)) + Val(CheckNull(rstMaterialBalance.Fields("Col4").Value)) + Val(CheckNull(rstMaterialBalance.Fields("Col5").Value)) - Val(CheckNull(rstMaterialBalance.Fields("Col6").Value))
    Else
        CalculateMaterialBalance = 0
    End If
    Call CloseRecordset(rstMaterialBalance)
    Exit Function
ErrorHandler:
    Call CloseRecordset(rstMaterialBalance)
End Function
Public Sub FocusSelect(ByVal xTextBox As Object)
    On Error Resume Next
    If Len(xTextBox.Text) = 0 Then Exit Sub
    xTextBox.SelStart = 0
    xTextBox.SelLength = Len(xTextBox.Text)
End Sub
Public Sub ValidateKey(ByRef xTextBox As TextBox, ByRef KeyAscii As Integer, ByVal DecimalPlaces As Integer)
    Select Case KeyAscii
        Case vbKey0, vbKey1, vbKey2, vbKey3, vbKey4, vbKey5, vbKey6, vbKey7, vbKey8, vbKey9, vbKeyBack
        Case vbKeyDelete
            If DecimalPlaces = 0 Or InStr(xTextBox.Text, ".") <> 0 Then KeyAscii = 0
        Case vbKeyInsert
            If xTextBox.SelStart <> 0 Or InStr(xTextBox.Text, "-") <> 0 Then KeyAscii = 0
        Case Else
            KeyAscii = 0
    End Select
End Sub
Public Function ValidateNumber(ByRef xTextBox As TextBox, ByVal DecimalPlaces As Integer) As Boolean
    
    ValidateNumber = False
    If IsNumeric(xTextBox.Text) Then
        If DecimalPlaces > 0 Then
            xTextBox.Text = Format(Val(xTextBox.Text), "0." + String(DecimalPlaces - 1, "#") + "0")
        Else
            xTextBox.Text = Format(Val(xTextBox.Text), "0")
        End If
        ValidateNumber = True
    Else
        FocusSelect xTextBox
    End If
    
End Function
Public Function ValidateDate(ByVal xMaskEdBox As Object, Optional AllowBlank As Boolean) As Boolean
    ValidateDate = True
    If AllowBlank = True And xMaskEdBox.Text = "  -  -    " Then Exit Function
    If Val(Left(xMaskEdBox.Text, 2)) > 0 And Val(Mid(xMaskEdBox.Text, 4, 2)) > 0 And Val(Mid(xMaskEdBox.Text, 4, 2)) <= 12 And Val(Right(xMaskEdBox.Text, 4)) > 0 And Len(Trim(xMaskEdBox.Text)) = 10 Then
        If IsDate(Left(xMaskEdBox.Text, 2) & "-" & MonthName(Mid(xMaskEdBox.Text, 4, 2), True) & "-" & Right(xMaskEdBox.Text, 4)) Then
            Exit Function
        End If
    End If
    ValidateDate = False: FocusSelect xMaskEdBox
End Function
Public Function GetDate(ByVal strInput As String) As String
    If strInput = "  -  -    " Then
        GetDate = "Null"
    Else
        GetDate = CStr(Left(strInput, 2)) & "-" & MonthName(Mid(strInput, 4, 2), True) & "-" & CStr(Right(strInput, 4))
    End If
End Function
Public Function FillList(ByVal lvwName As ListView, ByVal ColHdr As String, ByRef xRecordset As Recordset) As String
    Dim LITem As ListItem
    
    If xRecordset.RecordCount = 0 Then Exit Function
    DoEvents
    lvwName.ColumnHeaders.Add 1, , ColHdr
    lvwName.ColumnHeaders.Add 2, , ""
    xRecordset.MoveFirst
    
    Do While Not xRecordset.EOF
        Set LITem = lvwName.ListItems.Add(, , xRecordset.Fields(0).Value)
        LITem.ListSubItems.Add , , xRecordset.Fields(1).Value
        xRecordset.MoveNext
    Loop
    
    LockWindowUpdate lvwName.hwnd
    lvwName.ColumnHeaders(1).Width = lvwName.Width
    lvwName.ColumnHeaders(2).Width = 0
    LockWindowUpdate 0
    
        
End Function
Public Function SelectedItems(ByVal lvwName As ListView, Optional ByVal lvwWithCheckBox As Boolean = True) As String
    Dim i As Integer
    
    For i = 1 To lvwName.ListItems.Count
        If lvwWithCheckBox Then
            If lvwName.ListItems(i).Checked Then
                SelectedItems = SelectedItems + IIf(SelectedItems = "", "'", ", '") + lvwName.ListItems.Item(i).SubItems(1) + "'"
            End If
        Else
            If lvwName.ListItems(i).Selected Then
                SelectedItems = SelectedItems + IIf(SelectedItems = "", "'", ", '") + lvwName.ListItems.Item(i).SubItems(1) + "'"
            End If
        End If
    Next i
    SelectedItems = IIf(SelectedItems = "", "''", SelectedItems)
End Function
Public Function DisableCloseButton(frm As Form) As Boolean
    
    Dim lHndSysMenu As Long, lAns1 As Long, lAns2 As Long
    lHndSysMenu = GetSystemMenu(frm.hwnd, 0)
    lAns1 = RemoveMenu(lHndSysMenu, 6, MF_BYPOSITION) 'Remove close button
    lAns2 = RemoveMenu(lHndSysMenu, 5, MF_BYPOSITION) 'Remove seperator bar
    DisableCloseButton = (lAns1 <> 0 And lAns2 <> 0) 'Return True if both calls were successful
    
End Function
Public Function bVerifySum10(ByVal ISBN As String) As Boolean
    If Len(ISBN) < 13 Then bVerifySum10 = False: Exit Function
    If Len(Trim(ISBN)) < 13 Or Mid(Trim(ISBN), 12, 1) <> "-" Or InStr(1, "0123456789X", Right(Trim(ISBN), 1)) = 0 Or Len(Replace(ISBN, "-", "")) <> 10 Then bVerifySum10 = False: Exit Function
    ISBN = Replace(ISBN, "-", "")
    Dim i As Integer, K As Integer
    For K = 10 To 2 Step -1
        i = i + CInt(Val(Mid(ISBN, (10 - (K - 1)), 1))) * K
    Next
    If (i Mod 11) = 0 And Mid(ISBN, 10, 1) = "0" Then
        bVerifySum10 = True: Exit Function
    ElseIf UCase(Mid(ISBN, 10, 1)) = "X" Then
        i = i + 10
    Else
        i = i + CInt(Mid(ISBN, 10, 1))
    End If
    If Not ((i Mod 11) = 0) Then bVerifySum10 = False: Exit Function
    bVerifySum10 = True
    On Error GoTo 0
End Function
Public Function bVerifySum13(ByVal ISBN As String) As Boolean
    If Len(Trim(ISBN)) < 17 Or Mid(ISBN, 1, 3) <> "978" Or Mid(Trim(ISBN), 16, 1) <> "-" Or InStr(1, "0123456789", Right(Trim(ISBN), 1)) = 0 Or Len(Replace(ISBN, "-", "")) <> 13 Then bVerifySum13 = False: Exit Function
    ISBN = Replace(ISBN, "-", "")
    Dim i As Integer, K As Integer
    i = 30
    For K = 4 To 12 Step 2
        i = i + CInt(Mid(ISBN, K - 1, 1)) + (3 * CInt(Mid(ISBN, K, 1)))
    Next
    If Not (Mid(ISBN, 13, 1) = Trim(str((10 - i Mod 10) Mod 10))) Then bVerifySum13 = False: Exit Function
    bVerifySum13 = True
    On Error GoTo 0
End Function
Public Sub UpdateUserAction(ByVal Activity As String, ByVal Action As String, ByVal Description As String, ByVal xConnection As ADODB.Connection)
    If UserName = "sa" Then Exit Sub
    On Error GoTo ErrorHandler
    Dim lpBuff As String * 1024
    GetComputerName lpBuff, Len(lpBuff)
    xConnection.Execute "INSERT INTO UserAction VALUES('" & UserName & "','" & Format(Now(), "dd-MMM-yyyy hh:mm:ss") & "','" & Activity & "','" & Action & "','" & Description & "','" & Left(lpBuff, (InStr(1, lpBuff, vbNullChar)) - 1) & "')"
    Exit Sub
ErrorHandler:
    Call DisplayError("Faied to update User Log")
End Sub
Private Function WindowProc(ByVal Lwnd As Long, ByVal Lmsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    Dim MouseKeys As Long
    Dim Rotation As Long
    Dim Xpos As Long
    Dim Ypos As Long
    If Lmsg = WM_MOUSEWHEEL Then
        MouseKeys = wParam And 65535
        Rotation = wParam / 65536
        Xpos = lParam And 65535
        Ypos = lParam / 65536
        
        'determine if mouse wheel is being moved up or down
                
        If Rotation = -120 Then
           'call scroll method of datagrid and specify the number of columns and rows to scroll through DataGrid.Scroll colNum, rowNum
           MyControl.Scroll 0, 3
            
        Else
            MyControl.Scroll 0, -3
        End If
        
    End If
    WindowProc = CallWindowProc(LocalPrevWndProc, Lwnd, Lmsg, wParam, lParam)
End Function
Public Sub WheelHook(PassedControl As Object)
    On Error Resume Next
    Set MyControl = PassedControl
    LocalHwnd = PassedControl.hwnd
    LocalPrevWndProc = SetWindowLong(LocalHwnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub
Public Sub WheelUnHook()
    Dim WorkFlag As Long
    On Error Resume Next
    WorkFlag = SetWindowLong(LocalHwnd, GWL_WNDPROC, LocalPrevWndProc)
    Set MyControl = Nothing
End Sub
Public Function FMod(a As Variant, B As Variant) As Variant 'Floating Point Modulus
   FMod = a - Int(a / B) * B + CLng(Sgn(a) <> Sgn(B)) * B
End Function
Public Function chkFieldExists(ByVal strField As String, ByVal strTable As String) As Boolean
    Dim oRecordset As New ADODB.Recordset, i As Integer
    oRecordset.Open "SELECT * FROM " & strTable, CxnDatabase, adOpenKeyset, adLockReadOnly
    For i = 0 To oRecordset.Fields.Count - 1
        If strField = oRecordset.Fields.Item(i).Name Then chkFieldExists = True: Call CloseRecordset(oRecordset): Exit Function
    Next i
    Call CloseRecordset(oRecordset)
End Function
Public Sub Sendkeys(Text As Variant, Optional Wait As Boolean = False)
   Dim WshShell As Object
   Set WshShell = CreateObject("wscript.shell")
   WshShell.Sendkeys CStr(Text), Wait
        
        Set WshShell = Nothing
End Sub

