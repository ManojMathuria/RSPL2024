VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form FrmReportViewer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report Viewer"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ReportViewer.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5800
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin VB.Menu MnuEMail 
      Caption         =   "E-Mail"
   End
End
Attribute VB_Name = "FrmReportViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Report As New Report
Public Subject As String
Public Message As String
Public EMailID As String
Public CCID As String
Public Attachment As String
Private Sub Form_Load()
    With CRViewer1
        .ReportSource = Report
        .Zoom 100
        .ViewReport
    End With
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call CloseForm(FrmReportViewer)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set Report = Nothing
End Sub
Private Sub Form_Resize()
    With CRViewer1
        .Top = 0
        .Left = 0
        .Width = ScaleWidth
        .Height = ScaleHeight
    End With
End Sub
Private Sub CRViewer1_RefreshButtonClicked(UseDefault As Boolean)
    Report.PrinterSetup (FrmReportViewer.hwnd)
    UseDefault = True
End Sub
Private Sub MnuEMail_Click()
   Dim oOutlook As New Outlook.Application
   Dim oOutlookMsg As Outlook.MailItem
    If EMailID = "" Then
        Report.PrintOut True   ' Print Report With Prompt
        Exit Sub
    End If
    On Error Resume Next
    Report.ExportOptions.FormatType = crEFTPortableDocFormat    ' Set the Export Format As .Pdf
    Report.ExportOptions.DestinationType = crEDTDiskFile
    Report.ExportOptions.DiskFileName = App.Path & "\Report\" & Trim(Attachment) & ".Pdf"
    Report.Export False
   Set oOutlookMsg = oOutlook.CreateItem(olMailItem)
    With oOutlookMsg
        .To = EMailID
        If CCID <> "" Then .CC = CCID
        .Subject = Subject
        .HTMLBody = "<Font Face='Calibri' Size='3'>" & Message & "</Font>"
        .Attachments.Add (App.Path & "\Report\" & Trim(Attachment) & ".Pdf")
        .Importance = olImportanceHigh
        .ReadReceiptRequested = True
        .Display
    End With
   Set oOutlookMsg = Nothing
   Set oOutlook = Nothing
End Sub
