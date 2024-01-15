VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{01646141-065C-11D4-8ED3-00E07D815373}#1.0#0"; "MBBrowse.ocx"
Begin VB.MDIForm MdiMainMenu 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Easy Publish Production Management System Version - 20 | 11.17 "
   ClientHeight    =   8490
   ClientLeft      =   165
   ClientTop       =   150
   ClientWidth     =   14970
   Icon            =   "MainMenu.frx":0000
   LinkTopic       =   "MdiMainMenu"
   LockControls    =   -1  'True
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picBackdrop 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   14910
      TabIndex        =   4
      Top             =   7815
      Visible         =   0   'False
      Width           =   14970
      Begin VB.PictureBox picStretched 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   7260
         Left            =   2040
         ScaleHeight     =   7260
         ScaleWidth      =   4095
         TabIndex        =   6
         Top             =   600
         Width           =   4095
      End
      Begin VB.PictureBox picOriginal 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   5055
         Left            =   240
         Picture         =   "MainMenu.frx":0442
         ScaleHeight     =   5055
         ScaleWidth      =   6105
         TabIndex        =   5
         Top             =   120
         Width           =   6105
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   60000
      Left            =   5880
      Top             =   3960
   End
   Begin MBBrowse.BrowseFF BrowseFF1 
      Left            =   720
      Top             =   1440
      _ExtentX        =   1085
      _ExtentY        =   1085
      ReturnOnlyFSDirs=   -1  'True
      ShowCurrentPath =   0   'False
      ShowEditBox     =   0   'False
      ValidatePath    =   0   'False
      StartUpPosition =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   375
      Width           =   14970
      _ExtentX        =   26405
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   18
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Add [Ctrl+A]"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Edit [Ctrl+E]"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Delete [Ctrl+D]"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Save [Ctrl+S]"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Cancel"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Refresh [F5]"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Filter"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Print [Ctrl+P]"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Print Preview [Ctrl+V]"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Mail"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "First [Ctrl+F]"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Previous [Ctrl+P]"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Next [Ctrl+N]"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Last [Ctrl+L]"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exit"
            ImageIndex      =   15
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1440
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":B3DF
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":B923
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":BE67
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":BF7B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":C08F
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":C1A3
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":C2FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":C843
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":C957
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":CE9B
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":CFAF
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":D0C3
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":D1D7
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":D2EB
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMenu.frx":D3FF
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   14940
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   14970
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   8190
      Width           =   14970
      _ExtentX        =   26405
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7232
            MinWidth        =   6306
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   7232
            MinWidth        =   6306
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5283
            MinWidth        =   4357
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3572
            MinWidth        =   2646
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   2427
            MinWidth        =   1501
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   5400
      Top             =   3960
   End
   Begin VB.Menu MnuCompany 
      Caption         =   "&Company"
      Begin VB.Menu MnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu MnuClose 
         Caption         =   "Close"
         Enabled         =   0   'False
      End
      Begin VB.Menu MnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuEdit 
         Caption         =   "Edit"
         Enabled         =   0   'False
      End
      Begin VB.Menu MnuLine6 
         Caption         =   "-"
      End
      Begin VB.Menu MnuDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu MnuLine4 
         Caption         =   "-"
      End
      Begin VB.Menu MnuBackup 
         Caption         =   "Backup"
      End
      Begin VB.Menu MnuRestore 
         Caption         =   "Restore"
      End
      Begin VB.Menu MnuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu MnuMasters 
      Caption         =   "&Masters"
      Enabled         =   0   'False
      Tag             =   "0100"
      Begin VB.Menu MnuAccount 
         Caption         =   "Account"
         Tag             =   "0101"
         Begin VB.Menu MnuSupplier 
            Caption         =   "Supplier"
         End
         Begin VB.Menu MnuPrinter 
            Caption         =   "Printer"
            Begin VB.Menu MnuBookPrinter 
               Caption         =   "Book"
            End
            Begin VB.Menu MnuTitlePrinter 
               Caption         =   "Title"
            End
         End
         Begin VB.Menu MnuLaminator 
            Caption         =   "Laminator"
         End
         Begin VB.Menu MnuBinder 
            Caption         =   "Binder"
         End
         Begin VB.Menu MnuProcessor 
            Caption         =   "Processor"
         End
         Begin VB.Menu MnuGodown 
            Caption         =   "Godown"
         End
      End
      Begin VB.Menu MnuLine7 
         Caption         =   "-"
      End
      Begin VB.Menu MnuBook 
         Caption         =   "Book"
         Tag             =   "0102"
         Begin VB.Menu MnuFreshBook 
            Caption         =   "Fresh"
         End
         Begin VB.Menu MnuRepairBook 
            Caption         =   "Repair"
         End
      End
      Begin VB.Menu MnuClass 
         Caption         =   "Class"
         Tag             =   "0103"
      End
      Begin VB.Menu MnuBoard 
         Caption         =   "Board"
         Tag             =   "0104"
      End
      Begin VB.Menu MnuSubject 
         Caption         =   "Subject"
         Tag             =   "0105"
      End
      Begin VB.Menu MnuGroup 
         Caption         =   "Group"
         Tag             =   "0106"
      End
      Begin VB.Menu MnuBindingType 
         Caption         =   "Binding Type"
         Tag             =   "0107"
      End
      Begin VB.Menu MnuLaminationType 
         Caption         =   "Lamination Type"
         Tag             =   "0108"
      End
      Begin VB.Menu MnuLine5 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSize 
         Caption         =   "Size"
         Tag             =   "0109"
      End
      Begin VB.Menu MnuLine8 
         Caption         =   "-"
      End
      Begin VB.Menu MnuPaper 
         Caption         =   "Paper"
         Tag             =   "0110"
      End
      Begin VB.Menu MnuLine9 
         Caption         =   "-"
      End
      Begin VB.Menu MnuOutsourceItem 
         Caption         =   "Outsource Item"
         Tag             =   "0111"
      End
      Begin VB.Menu MnuLine15 
         Caption         =   "-"
      End
      Begin VB.Menu MnuEditorial 
         Caption         =   "Editorial"
         Tag             =   "0113"
      End
      Begin VB.Menu MnuLine1500 
         Caption         =   "-"
      End
      Begin VB.Menu MnuUser 
         Caption         =   "User"
         Tag             =   "0112"
      End
   End
   Begin VB.Menu MnuTransactions 
      Caption         =   "&Transactions"
      Enabled         =   0   'False
      Tag             =   "0200"
      Begin VB.Menu MnuPrintPlanning 
         Caption         =   "Print Planning"
         Tag             =   "0201"
         Begin VB.Menu MnuBookPrintPlanning 
            Caption         =   "Book"
         End
         Begin VB.Menu MnuTitlePrintPlanning 
            Caption         =   "Title"
         End
      End
      Begin VB.Menu MnuBookPrintOrder 
         Caption         =   "Book Print Order"
         Tag             =   "0202"
         Begin VB.Menu MnuFreshBookPrintOrder 
            Caption         =   "Fresh"
         End
         Begin VB.Menu MnuRepairBookPrintOrder 
            Caption         =   "Repair"
         End
      End
      Begin VB.Menu MnuPOStatusUpdation 
         Caption         =   "Print Order Status Updation"
         Tag             =   "0203"
         Begin VB.Menu MnuPOStatusUpdation01 
            Caption         =   "Title Printing"
         End
         Begin VB.Menu MnuPOStatusUpdation02 
            Caption         =   "Book Printing"
         End
         Begin VB.Menu MnuPOStatusUpdation03 
            Caption         =   "Book Binding"
         End
         Begin VB.Menu MnuPOStatusUpdation04 
            Caption         =   "Production Planning"
         End
         Begin VB.Menu MnuPOStatusUpdation05 
            Caption         =   "Debit Notes"
         End
      End
      Begin VB.Menu MnuBookProcessOrder 
         Caption         =   "Book Process Order"
         Tag             =   "0204"
         Visible         =   0   'False
         Begin VB.Menu MnuBookProcessOrder01 
            Caption         =   "Fresh"
         End
         Begin VB.Menu MnuBookProcessOrder02 
            Caption         =   "Repair"
         End
         Begin VB.Menu MnuBookProcessOrder03 
            Caption         =   "Title"
         End
      End
      Begin VB.Menu MnuLine12 
         Caption         =   "-"
      End
      Begin VB.Menu MnuPaperPurchaseOrder 
         Caption         =   "Paper Purchase Order"
         Tag             =   "0207"
         Begin VB.Menu MnuBookPaperPurchaseOrder 
            Caption         =   "Book"
         End
         Begin VB.Menu MnuTitlePaperPurchaseOrder 
            Caption         =   "Title"
         End
      End
      Begin VB.Menu MnuPaperMovement 
         Caption         =   "Paper Movement"
         Tag             =   "0210"
         Begin VB.Menu MnuBookPaperMovement 
            Caption         =   "Book"
         End
         Begin VB.Menu MnuTitlePaperMovement 
            Caption         =   "Title"
         End
      End
      Begin VB.Menu MnuIOStatusUpdation 
         Caption         =   "Paper Issue Status Updation"
         Tag             =   "0209"
      End
      Begin VB.Menu MnuLine10 
         Caption         =   "-"
      End
      Begin VB.Menu MnuOutsourceItemPurchaseOrder 
         Caption         =   "Outsource Item Purchase Order"
         Tag             =   "0211"
      End
      Begin VB.Menu MnuLine99 
         Caption         =   "-"
      End
      Begin VB.Menu MnuMaterialIssueOrder 
         Caption         =   "Material Issue Order"
         Tag             =   "0213"
      End
      Begin VB.Menu MnuMaterialMovement 
         Caption         =   "Material Movement"
         Tag             =   "0214"
      End
      Begin VB.Menu MnuLine16 
         Caption         =   "-"
      End
      Begin VB.Menu MnuStockJournal 
         Caption         =   "Stock Journal"
         Tag             =   "0215"
      End
      Begin VB.Menu MnuLine100 
         Caption         =   "-"
      End
      Begin VB.Menu MnuDebitNote 
         Caption         =   "Debit Note"
         Visible         =   0   'False
         Begin VB.Menu MnuBookDebitNote 
            Caption         =   "Book"
         End
         Begin VB.Menu MnuPaperDebitNote 
            Caption         =   "Paper"
         End
         Begin VB.Menu MnuMiscDebitNote 
            Caption         =   "Miscellaneous"
         End
      End
      Begin VB.Menu OutsourceItemPurchaseOrderOld 
         Caption         =   "Outsource Item Purchase Order Old"
         Visible         =   0   'False
      End
      Begin VB.Menu MnuDeliveryChallan 
         Caption         =   "Delivery Challan"
         Begin VB.Menu MnuBookDeliveryChallan 
            Caption         =   "Book"
         End
         Begin VB.Menu MnuTitleDeliveryChallan 
            Caption         =   "Title"
         End
      End
   End
   Begin VB.Menu MnuReports 
      Caption         =   "&Reports"
      Enabled         =   0   'False
      Tag             =   "0300"
      Begin VB.Menu MnuMemoOrderRegister 
         Caption         =   "Memo Order Register"
      End
      Begin VB.Menu MnuBookPrintPlanningOrder 
         Caption         =   "Print Planning Order"
         Begin VB.Menu MnuBookPrintPlanningOrderBook 
            Caption         =   "Book"
         End
         Begin VB.Menu MnuBookPrintPlanningOrderTitle 
            Caption         =   "Title"
         End
      End
      Begin VB.Menu MnuPrintPlanningRegister 
         Caption         =   "Print Planning Register"
         Tag             =   "0301"
         Begin VB.Menu MnuBookPrintPlanningRegister 
            Caption         =   "Book"
         End
         Begin VB.Menu MnuTitlePrintPlanningRegister 
            Caption         =   "Title"
         End
      End
      Begin VB.Menu MnuPOStatusRegister 
         Caption         =   "Print Order Status Register"
         Tag             =   "0302"
         Begin VB.Menu MnuPOStatusRegister01 
            Caption         =   "Bookwise"
         End
         Begin VB.Menu MnuPOStatusRegister05 
            Caption         =   "Print Orderwise"
         End
         Begin VB.Menu MnuPOStatusRegister02 
            Caption         =   "Title Printerwise"
         End
         Begin VB.Menu MnuPOStatusRegister03 
            Caption         =   "Book Printerwise"
         End
         Begin VB.Menu MnuPOStatusRegister04 
            Caption         =   "Book Binderwise"
         End
         Begin VB.Menu MnuPOStatusRegister06 
            Caption         =   "Busy"
         End
      End
      Begin VB.Menu MnuLine50 
         Caption         =   "-"
      End
      Begin VB.Menu MnuPaperIssueRegister 
         Caption         =   "Paper Issue Register"
         Tag             =   "0303"
      End
      Begin VB.Menu MnuPaperStockRegister 
         Caption         =   "Paper Stock Register"
         Tag             =   "0304"
      End
      Begin VB.Menu MnuLine51 
         Caption         =   "-"
      End
      Begin VB.Menu MnuMaterialStockRegister 
         Caption         =   "Material Stock Register"
         Tag             =   "0305"
         Begin VB.Menu MnuMaterialStockRegister01 
            Caption         =   "Binderwise/Bookwise/Itemwise"
         End
         Begin VB.Menu MnuMaterialStockRegister02 
            Caption         =   "Binderwise/Itemwise"
         End
      End
      Begin VB.Menu MnuLine52 
         Caption         =   "-"
      End
      Begin VB.Menu MnuPOStatusReg 
         Caption         =   "Purchase Order Status Register"
         Tag             =   "0306"
         Begin VB.Menu MnuPOStatusReg02 
            Caption         =   "Outsource Item"
         End
         Begin VB.Menu MnuPOStatusReg03 
            Caption         =   "Insource Item"
            Begin VB.Menu MnuPOStatusReg0301 
               Caption         =   "Refresh Book"
            End
            Begin VB.Menu MnuPOStatusReg0302 
               Caption         =   "Repair Book"
            End
            Begin VB.Menu MnuPOStatusReg0303 
               Caption         =   "Title"
            End
         End
      End
      Begin VB.Menu MnuLine53 
         Caption         =   "-"
      End
      Begin VB.Menu MnuBillRegister 
         Caption         =   "Bill Register"
         Tag             =   "0307"
      End
      Begin VB.Menu MnuDayBook 
         Caption         =   "Day Book"
      End
      Begin VB.Menu MnuPendingPaymentRegister 
         Caption         =   "Pending Payment Register"
         Tag             =   "0308"
      End
      Begin VB.Menu MnuPendingDNRegister 
         Caption         =   "Pending Debit Notes Register"
         Tag             =   "0309"
      End
      Begin VB.Menu MnuLine54 
         Caption         =   "-"
      End
      Begin VB.Menu MnuBookList 
         Caption         =   "List of Books"
         Tag             =   "0310"
      End
      Begin VB.Menu MnuCorrectionList 
         Caption         =   "List of Corrections"
         Tag             =   "0311"
         Begin VB.Menu MnuCorrectionList01 
            Caption         =   "Production"
         End
         Begin VB.Menu MnuCorrectionList02 
            Caption         =   "Editorial"
         End
      End
      Begin VB.Menu MnuLine55 
         Caption         =   "-"
      End
      Begin VB.Menu MnuProductionPlanning 
         Caption         =   "Production Planning"
         Tag             =   "0313"
         Begin VB.Menu MnuProductionPlanning01 
            Caption         =   "Main Orders"
         End
         Begin VB.Menu MnuProductionPlanning02 
            Caption         =   "Supplement Orders"
         End
      End
      Begin VB.Menu MnuLine44 
         Caption         =   "-"
      End
      Begin VB.Menu MnuPrintUtilities 
         Caption         =   "Print Utilities"
         Tag             =   "0315"
         Begin VB.Menu MnuBookPOPrintUtility 
            Caption         =   "Book Print Order"
         End
         Begin VB.Menu MnuPaperPOPrintUtility 
            Caption         =   "Paper Purchase Order"
         End
         Begin VB.Menu MnuOpBal 
            Caption         =   "Opening Balance"
         End
      End
   End
   Begin VB.Menu MnuUtilities 
      Caption         =   "&Utilities"
      Enabled         =   0   'False
      Begin VB.Menu MnuCalculator 
         Caption         =   "Calculator"
      End
      Begin VB.Menu MnuBookReceipt 
         Caption         =   "Book Receipt"
         Tag             =   "0206"
      End
      Begin VB.Menu MnuCostSheet 
         Caption         =   "Cost Sheet"
      End
      Begin VB.Menu MnuImportBal 
         Caption         =   "Import Balances"
         Begin VB.Menu MnuImportBal01 
            Caption         =   "Print Order"
         End
         Begin VB.Menu MnuImportBal02 
            Caption         =   "Paper"
         End
         Begin VB.Menu MnuImportBal03 
            Caption         =   "Outsource Item"
         End
      End
      Begin VB.Menu MnuItemDetail 
         Caption         =   "Item Detail"
      End
      Begin VB.Menu MnuJobWork 
         Caption         =   "Job Work"
         Begin VB.Menu MnuBookJobWork 
            Caption         =   "Goods Sent"
         End
         Begin VB.Menu MnuTitleJobWork 
            Caption         =   "Goods Received"
         End
      End
      Begin VB.Menu MnuSMS 
         Caption         =   "Send SMS"
      End
   End
   Begin VB.Menu MnuWindow 
      Caption         =   "&Window"
      Enabled         =   0   'False
      WindowList      =   -1  'True
      Begin VB.Menu MnuTileHorizontally 
         Caption         =   "Tile Horizontally"
      End
      Begin VB.Menu MnuTileVertically 
         Caption         =   "Tile Vertically"
      End
      Begin VB.Menu MnuCascade 
         Caption         =   "Cascade"
      End
      Begin VB.Menu MnuArrangeIcons 
         Caption         =   "Arrange Icons"
      End
      Begin VB.Menu MnuMinimizeAll 
         Caption         =   "Minimize All"
      End
      Begin VB.Menu MnuCloseAll 
         Caption         =   "Close All"
      End
   End
End
Attribute VB_Name = "MdiMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private WithEvents oHuffman As clsHuffman
Attribute oHuffman.VB_VarHelpID = -1
    Private oRegistry As New clsRegistry
    Private Developer As String
    Dim strFile As New FileSystemObject
    'For MDI Picture
    Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
    Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
    Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
    Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Sub MDIForm_Load()
    If GetSystemMetrics(SM_CXSCREEN) < 800 Or GetSystemMetrics(SM_CYSCREEN) < 600 Then
        Call MsgBox("Saral requires atleast 800 x 600 screen resolution.", vbInformation, "Cannot Continue !")
        Call CloseForm(MdiMainMenu)
        Exit Sub
    End If
    If Dir(App.Path & "\Database", vbDirectory) = "" Then FSO.CreateFolder App.Path & "\Database"
    If Dir(App.Path & "\Saral.ini") = "" Then
        WriteToFile "Database Path", App.Path & "\Database"
        WriteToFile "Busy Database Name", "BusyComp0001_db1" & Trim(str(Year(Date)))
        WriteToFile "Saral Database Name", "Saral"
    End If
    DatabasePath = Trim(ReadFromFile("Database Path"))
'    ServerName = Trim(ReadFromFile("Server Name"))
'    ServerPassword = Trim(ReadFromFile("Server Password"))
    OutsourceItemPurchaseOrderOld.Visible = False
    'Developer = "Developed By Sanjeev Gupta Ph:+91-9968096291" & Space(68) And Mohd Shamshad Alam Ph. +91-9958926213
End Sub
Private Sub MDIForm_Resize()
    
    Dim client_rect As RECT
    Dim client_hwnd As Long
    picStretched.Move 0, 0, ScaleWidth, ScaleHeight
    
    'Copy the original picture into picStretched.
    picStretched.PaintPicture picOriginal.Picture, -20, -40, picStretched.ScaleWidth, picStretched.ScaleHeight, -8, -8, picOriginal.ScaleWidth, picOriginal.ScaleHeight
    'Set the MDI form's picture.
    Picture = picStretched.Image
    'Invalidate the picture.
    client_hwnd = FindWindowEx(Me.hwnd, 0, "MDIClient", vbNullChar)
    GetClientRect client_hwnd, client_rect
    InvalidateRect client_hwnd, client_rect, 1
    If Me.WindowState <> vbMinimized Then Me.WindowState = vbMaximized
End Sub
Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    If Not MnuOpen.Enabled Then
        MsgBox "           Cannot Quit till You have Company Open." & vbCrLf & "Kindly make sure to Close the Company before Quitting !!!", vbExclamation, "Cannot Close !"
        Cancel = 1
       'WriteToFile "Server Name", "Saral" 'Comment By Shamshad
        'WriteToFile "Server Name", "Rackserver"
        Exit Sub
    Else
        Call CloseForm(MdiMainMenu)
        If strFile.FileExists(App.Path & "\SaralTemp.exe") = True Then
           Kill App.Path & "\Saral.exe"
           Name App.Path & "\SaralTemp.exe" As App.Path & "\Saral.exe"
        End If
    End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    On Error GoTo ErrorHandler
    Set oHuffman = Nothing
    Set FSO = Nothing
    Set oRegistry = Nothing
    CloseMainConnection
    If Not CxnDatabase Is Nothing Then Set CxnDatabase = Nothing
    Call AnimateWindow(Me.hwnd, CInt(500), AW_HIDE Or AW_BLEND)
    Exit Sub
ErrorHandler:
End Sub
Private Sub MnuBookDeliveryChallan_Click()
 On Error Resume Next
    If Not IsFormLoaded("Delivery Challan [Book]") Then
        Dim FrmBookPaperDeliveryChallan As New FrmPaperDeliveryChallan
        FrmBookPaperDeliveryChallan.OrderType = "1"
        Load FrmBookPaperDeliveryChallan
        If Err.Number <> 364 Then
            FrmBookPaperDeliveryChallan.Show
        End If
    End If
End Sub

Private Sub MnuSMS_Click()
   On Error Resume Next
    Load FrmSMS
    If Err.Number <> 364 Then
        FrmSMS.Show
    End If
End Sub

Private Sub MnuTitleDeliveryChallan_Click()
    On Error Resume Next
    If Not IsFormLoaded("Delivery Challan [Title]") Then
        Dim FrmTitlePaperDeliveryChallan As New FrmPaperDeliveryChallan
        FrmTitlePaperDeliveryChallan.OrderType = "2"
        Load FrmTitlePaperDeliveryChallan
        If Err.Number <> 364 Then
            FrmTitlePaperDeliveryChallan.Show
        End If
    End If
End Sub
Private Sub MnuBookJobWork_Click()
    On Error Resume Next
    FrmJobWork.ReportType = "1"
    Load FrmJobWork
    If Err.Number <> 364 Then
        FrmJobWork.Show
    End If
End Sub
Private Sub MnuBookPrintPlanningOrderBook_Click()
  On Error Resume Next
    
    If Not IsFormLoaded("Print Planning Order") Then
        Dim FrmPrintPlanningOrder As New FrmPrintPlanningOrder
        FrmPrintPlanningOrder.PlanningType = "1"
        Load FrmPrintPlanningOrder
        If Err.Number <> 364 Then
            FrmPrintPlanningOrder.Show
        End If
    End If
    
End Sub

Private Sub MnuBookPrintPlanningOrderTitle_Click()
  On Error Resume Next
    If Not IsFormLoaded("Print Planning Order") Then
        Dim FrmPrintPlanningOrder As New FrmPrintPlanningOrder
        FrmPrintPlanningOrder.PlanningType = "2"
        Load FrmPrintPlanningOrder
        If Err.Number <> 364 Then
            FrmPrintPlanningOrder.Show
        End If
    End If
End Sub

Private Sub mnuCalculator_Click()
    On Error Resume Next
    If Not RestorePreviousInstance("SciCalc", "Calculator") Then Shell "Calc.Exe", vbNormalFocus
End Sub

Private Sub MnuMemoOrderRegister_Click()
  
  On Error Resume Next
    If Not IsFormLoaded("Print Planning Order") Then
        Dim FrmPrintMemoOrderRegister As New FrmMemoOrderRegister
        FrmMemoOrderRegister.OrderType = "02"
        Load FrmMemoOrderRegister
        If Err.Number <> 364 Then
            FrmMemoOrderRegister.Show
        End If
    End If
    
End Sub

Private Sub MnuTileHorizontally_Click()
    MdiMainMenu.Arrange vbTileHorizontal
End Sub
Private Sub MnuTileVertically_Click()
    MdiMainMenu.Arrange vbTileVertical
End Sub
Private Sub MnuArrangeIcons_Click()
    MdiMainMenu.Arrange vbArrangeIcons
End Sub
Private Sub MnuCascade_Click()
    MdiMainMenu.Arrange vbCascade
End Sub
Private Sub MnuExit_Click()
    If MnuClose.Enabled Then
       MnuClose_Click
    End If
    If Forms.Count <= 1 Then
       Call CloseForm(MdiMainMenu)
    End If
End Sub
Private Sub MnuMinimizeAll_Click()
    Dim Form As Form
    For Each Form In Forms
        If Not TypeOf Form Is MDIForm Then
            Form.WindowState = vbMinimized
        End If
    Next Form
End Sub
Private Sub MnuCloseAll_Click()
    Dim Form As Form
    For Each Form In Forms
        If Not TypeOf Form Is MDIForm Then
            Unload Form
            Set Form = Nothing
        End If
    Next Form
End Sub
Private Sub MnuOpen_Click()
    
    Dim rstCompanyMaster As New ADODB.Recordset
    On Error GoTo OpenError
    Load FrmCompanyList
    FrmCompanyList.Show vbModal
    
    If CompCode <> "" Then
        BusySystemIndicator True
        CloseMainConnection
        CxnDatabase.CursorLocation = adUseClient
        ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DatabasePath & "\Saral." & CompCode & ";Persist Security Info=False;Jet OLEDB:Database Password=RSPLILoveMyINDIA"
        CxnDatabase.Open ConnectionString
        Load FrmLogin
        FrmLogin.Show vbModal
        If LoginSuccess Then
            StatusBar1.Panels(3).Text = "User Name : " & Trim(UserName)
            SetMenuOptions (True)
            rstCompanyMaster.Open "Select Name, '-Financial Year From '+Mid(Format(FinancialYearFrom,'dd-mm-yyyy'),1,2)+'-'+Mid(Format(FinancialYearFrom,'dd-mmm-yyyy'),4,3)+'-'+Mid(Format(FinancialYearFrom,'dd-mm-yyyy'),7,4)+' To '+Mid(Format(FinancialYearTo,'dd-mm-yyyy'),1,2)+'-'+Mid(Format(FinancialYearTo,'dd-mmm-yyyy'),4,3)+'-'+Mid(Format(FinancialYearTo,'dd-mm-yyyy'),7,4),ServerName,ServerPassword From CompanyMaster", CxnDatabase, adOpenKeyset, adLockReadOnly
            MdiMainMenu.Caption = "Easy Publish Production Management System Version - 20 | 11.17  [" & Trim(rstCompanyMaster.Fields("Name").Value) & Trim(rstCompanyMaster.Fields(1).Value) & "]"
            ServerName = Trim(rstCompanyMaster.Fields("ServerName").Value)
            ServerPassword = Trim(rstCompanyMaster.Fields("ServerPassword").Value)
            Call CloseRecordset(rstCompanyMaster)
            Exit Sub
        End If
        
    End If
    
    CloseMainConnection
    BusySystemIndicator False
    Exit Sub
OpenError:

    If Not rstCompanyMaster Is Nothing Then Set rstCompanyMaster = Nothing
    CloseMainConnection
    BusySystemIndicator False
    

End Sub
Private Sub MnuClose_Click()
    Dim Form As Form
'    If Forms.Count > 1 Then
'        MsgBox "            Cannot Close the Company till You have Open Forms." & vbCrLf & "Kindly make sure to Close all the Forms before Closing the Company !!!", vbExclamation, "Cannot Close !"
'   End If
    For Each Form In Forms
        If Not TypeOf Form Is MDIForm Then
            Unload Form
            Set Form = Nothing
        End If
    Next Form
    If Forms.Count <= 1 Then
        CloseMainConnection
        SetMenuOptions (False)
        MdiMainMenu.Caption = "Easy Publish Production Management System Version - 20 | 11.17 "
        StatusBar1.Panels(3).Text = ""
    End If
End Sub
Private Sub SetMenuOptions(bVal As Boolean)
    Dim Object As Object
    Dim rstUserChild As New ADODB.Recordset
    On Error GoTo ErrorHandler
    MnuOpen.Enabled = Not bVal
'   MnuNew.Enabled = Not bVal
    MnuClose.Enabled = bVal
    MnuEdit.Enabled = bVal
    MnuDelete.Enabled = Not bVal
    MnuBackup.Enabled = Not bVal
    MnuRestore.Enabled = Not bVal
    MnuUtilities.Enabled = bVal
    MnuWindow.Enabled = bVal
    If bVal Then
        rstUserChild.Open "Select [Module] From UserChild Where Code = '" & FixQuote(UserCode) & "' Order by [Module]", CxnDatabase, adOpenKeyset, adLockReadOnly
        For Each Object In Me
            If TypeName(Object) = "Menu" Then
                If Object.Tag <> "" Then
                    If UserLevel <> "1" Then
                        rstUserChild.MoveFirst
                        rstUserChild.Find "[Module] = '" & Trim(Object.Tag) & "'"
                        Object.Enabled = IIf(rstUserChild.EOF, False, True)
                        Object.Visible = IIf(rstUserChild.EOF, False, True)
                    Else
                        Object.Enabled = True
                    End If
                End If
            End If
        Next
    Else
        MnuMasters.Enabled = bVal
        MnuTransactions.Enabled = bVal
        MnuReports.Enabled = bVal
    End If
    'If User Left from the Job
    Dim dVal As String, lVal As Variant
    dVal = "31-MAY-2024"
    lVal = DateDiff("d", Format(Date, "dd-MMM-yyyy"), dVal)
    If lVal <= 0 Then AllowMastersModification = 0: AllowMastersDeletion = 0: AllowTransactionsModification = 0: AllowTransactionsDeletion = 0: MnuTransactions.Visible = False:  MnuTransactions.Enabled = False
ErrorHandler:
    Call CloseRecordset(rstUserChild)
End Sub

Private Sub MnuTitleJobWork_Click()
    On Error Resume Next
    FrmJobWork.ReportType = "2"
    Load FrmJobWork
    If Err.Number <> 364 Then
        FrmJobWork.Show
    End If
End Sub

Private Sub OutsourceItemPurchaseOrderOld_Click()
  On Error Resume Next
    Load FrmOutsourceItemPurchaseOrderOld
    If Err.Number <> 364 Then
        FrmOutsourceItemPurchaseOrderOld.Show
    End If
End Sub

Private Sub StatusBar1_PanelDblClick(ByVal Panel As MSComctlLib.Panel)
    
    If Panel.Index = 4 Or Panel.Index = 5 Then
        On Error Resume Next
        Shell "Control.Exe Date/Time", vbNormalFocus
    End If
    
End Sub
Private Sub Timer1_Timer()
    On Error Resume Next
    Static Counter As Integer
    Counter = Counter + 1
    StatusBar1.Panels(2).Text = Left(Developer, Counter)
    If Counter = Len(Developer) Then
        Counter = 0
    End If
    
    StatusBar1.Panels(4).Text = WeekdayName(Weekday(Date), True, vbSunday) + ", " + MonthName(Month(Date), True) + str$(Day(Date)) + ", " + Right(str$(Year(Date)), 2)
    StatusBar1.Panels(5).Text = Left(Time, 8)

End Sub
Private Sub Timer2_Timer()
'    Static T As Long
'    T = T + 60000
'    If T / 60000 = 240 Then
'            MnuBookReceipt_Click
'        T = 0
'    End If
  If Time() = "2:02:00 PM" Then If UserName = "Saral" Then MnuBookReceipt_Click
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index <= 17 Then
        If ActiveForm.Toolbar1.Buttons.Item(Button.Index).Enabled Then
            ActiveForm.Toolbar1_ButtonClick ActiveForm.Toolbar1.Buttons.Item(Button.Index)
        End If
    Else
        If Toolbar1.Buttons(1).Enabled Then 'Company Open
            If ActiveForm.Toolbar1.Buttons.Item(Button.Index).Enabled Then
                ActiveForm.Toolbar1_ButtonClick ActiveForm.Toolbar1.Buttons.Item(Button.Index)
            End If
        Else
            MnuExit_Click
        End If
    End If
End Sub
Private Sub oCreate_PercentDone(ByVal Percent As Integer)
    MdiMainMenu.ProgressBar1.Value = Percent
End Sub
Private Sub MnuDelete_Click()
    On Error Resume Next
    Load FrmCompanyList
    If Err.Number <> 364 Then
        FrmCompanyList.Caption = "Select Company To Delete..."
        FrmCompanyList.Show vbModal
        On Error GoTo ErrorHandler
        If CompCode <> "" Then
            Load FrmLogin
            If Err.Number <> 364 Then
                FrmLogin.Show vbModal
                If LoginSuccess Then
                    If UserLevel <> "1" Then
                        Call MsgBox("You don't have authority to Delete a Company !", vbInformation, App.Title)
                        CompCode = ""
                    End If
                End If
            End If
        End If
    End If
    CloseMainConnection
    If CompCode = "" Or (Not LoginSuccess) Then
        Exit Sub
    End If
    If MsgBox("Are you sure to delete the Company?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Delete !") <> vbYes Then
        Exit Sub
    End If
    MdiMainMenu.MousePointer = vbHourglass
    FSO.DeleteFile DatabasePath & "\Saral." & CompCode
    Call MsgBox("Successfully deleted the company !", vbInformation, App.Title)
    MdiMainMenu.MousePointer = vbNormal
    Exit Sub
ErrorHandler:
    DisplayError ("Failed to delete the company")
End Sub
Private Sub MnuBackup_Click()
    On Error Resume Next
    Dim strDestination As String
    
    Load FrmCompanyList
    If Err.Number <> 364 Then
        FrmCompanyList.Caption = "Select Company To Backup..."
        FrmCompanyList.Show vbModal
        On Error GoTo ErrorHandler
        If CompCode <> "" Then
            BrowseFF1.Caption = "Select Destination..."
            BrowseFF1.InitialFolder = App.Path & "\Backup"
            BrowseFF1.IncludeFiles = False
            If BrowseFF1.Browse = True Then
                strDestination = BrowseFF1.SelectedItem.Name
            End If
        End If
    End If
    CloseMainConnection
    If Len(strDestination) = 0 Or CompCode = "" Then
        Exit Sub
    End If
    If Right(strDestination, 1) <> "\" Then
        strDestination = RTrim(strDestination) & "\"
    End If
    strDestination = strDestination & CStr(Format(Date, "yyyymmdd")) & "." & CompCode
    MdiMainMenu.MousePointer = vbHourglass
    ShowProgressInStatusBar True
    If Dir(strDestination) <> "" Then
        Kill strDestination
    End If
    Set oHuffman = New clsHuffman
    Call oHuffman.EncodeFile(DatabasePath & "\Saral." & CompCode, strDestination, False)
    Set oHuffman = Nothing
    ShowProgressInStatusBar False
    MdiMainMenu.MousePointer = vbNormal
    Exit Sub

ErrorHandler:
    DisplayError ("Failed to backup the data of the company")
    
End Sub
Private Sub oHuffman_Progress(Percent As Integer)
    MdiMainMenu.ProgressBar1.Value = Percent
    DoEvents
End Sub
Private Sub MnuRestore_Click()
    On Error Resume Next
    Dim strSource As String
    Load FrmCompanyList
    If Err.Number <> 364 Then
        FrmCompanyList.Caption = "Select Company To Restore..."
        FrmCompanyList.Show vbModal
        On Error GoTo ErrorHandler
        If CompCode <> "" Then
            BrowseFF1.Caption = "Select Source..."
            BrowseFF1.InitialFolder = App.Path & "\Backup"
            BrowseFF1.IncludeFiles = True
            If BrowseFF1.Browse = True Then
                strSource = BrowseFF1.SelectedItem.Name
            End If
        End If
    End If
    CloseMainConnection
    If Len(strSource) = 0 Or CompCode = "" Then
        Exit Sub
    End If
    If Right(strSource, 3) <> CompCode Then
        DisplayError ("Failed to restore the data of the company")
        Exit Sub
    End If
       
    MdiMainMenu.MousePointer = vbHourglass
    ShowProgressInStatusBar True
    Set oHuffman = New clsHuffman
    Call oHuffman.DecodeFile(strSource, DatabasePath & "\Saral." & CompCode)
    Set oHuffman = Nothing
    ShowProgressInStatusBar False
    MdiMainMenu.MousePointer = vbNormal
    Exit Sub
ErrorHandler:
    DisplayError ("Failed to restore the data of the company")
End Sub
Private Sub MnuClass_Click()
    On Error Resume Next
    If Not IsFormLoaded("Class Master") Then
        Dim FrmClassMaster As New FrmGeneralMaster
        FrmClassMaster.MasterType = "4"
        Load FrmClassMaster
        If Err.Number <> 364 Then
            FrmClassMaster.Caption = "Class Master"
            FrmClassMaster.Show
        End If
    End If
End Sub
Private Sub MnuBoard_Click()
    On Error Resume Next
    If Not IsFormLoaded("Board Master") Then
        Dim FrmBoardMaster As New FrmGeneralMaster
        FrmBoardMaster.MasterType = "2"
        Load FrmBoardMaster
        If Err.Number <> 364 Then
            FrmBoardMaster.Caption = "Board Master"
            FrmBoardMaster.Show
        End If
    End If
End Sub
Private Sub MnuSubject_Click()
    On Error Resume Next
    If Not IsFormLoaded("Subject Master") Then
        Dim FrmSubjectMaster As New FrmGeneralMaster
        FrmSubjectMaster.MasterType = "3"
        Load FrmSubjectMaster
        If Err.Number <> 364 Then
            FrmSubjectMaster.Caption = "Subject Master"
            FrmSubjectMaster.Show
        End If
    End If
End Sub
Private Sub MnuGroup_Click()
    On Error Resume Next
    If Not IsFormLoaded("Group Master") Then
        Dim FrmGroupMaster As New FrmGeneralMaster
        FrmGroupMaster.MasterType = "5"
        Load FrmGroupMaster
        If Err.Number <> 364 Then
            FrmGroupMaster.Caption = "Group Master"
            FrmGroupMaster.Show
        End If
    End If
End Sub
Private Sub MnuBindingType_Click()
    On Error Resume Next
    If Not IsFormLoaded("Binding Type Master") Then
        Dim FrmBindingTypeMaster As New FrmGeneralMaster
        FrmBindingTypeMaster.MasterType = "6"
        Load FrmBindingTypeMaster
        If Err.Number <> 364 Then
            FrmBindingTypeMaster.Caption = "Binding Type Master"
            FrmBindingTypeMaster.Show
        End If
    End If
End Sub
Private Sub MnuLaminationType_Click()
    On Error Resume Next
    If Not IsFormLoaded("Lamination Type Master") Then
        Dim FrmLaminationTypeMaster As New FrmGeneralMaster
        FrmLaminationTypeMaster.MasterType = "7"
        Load FrmLaminationTypeMaster
        If Err.Number <> 364 Then
            FrmLaminationTypeMaster.Caption = "Lamination Type Master"
            FrmLaminationTypeMaster.Show
        End If
    End If
End Sub
Private Sub MnuSize_Click()
    On Error Resume Next
    If Not IsFormLoaded("Size Master") Then
        Dim FrmSizeMaster As New FrmGeneralMaster
        FrmSizeMaster.MasterType = "1"
        Load FrmSizeMaster
        If Err.Number <> 364 Then
            FrmSizeMaster.Caption = "Size Master"
            FrmSizeMaster.Show
        End If
    End If
End Sub
Private Sub MnuSupplier_Click()
    On Error Resume Next
    If Not IsFormLoaded("Supplier Master") Then
        Dim FrmSupplierMaster As New FrmAccountMaster
        FrmSupplierMaster.AccountType = "01"
        Load FrmSupplierMaster
        If Err.Number <> 364 Then
            FrmSupplierMaster.Caption = "Supplier Master"
            FrmSupplierMaster.Show
        End If
    End If
End Sub
Private Sub MnuBookPrinter_Click()
    On Error Resume Next
    If Not IsFormLoaded("Book Printer Master") Then
        Dim FrmBookPrinterMaster As New FrmAccountMaster
        FrmBookPrinterMaster.AccountType = "05"
        Load FrmBookPrinterMaster
        If Err.Number <> 364 Then
            FrmBookPrinterMaster.Caption = "Book Printer Master"
            FrmBookPrinterMaster.Show
        End If
    End If
End Sub
Private Sub MnuTitlePrinter_Click()
    On Error Resume Next
    If Not IsFormLoaded("Title Printer Master") Then
        Dim FrmTitlePrinterMaster As New FrmAccountMaster
        FrmTitlePrinterMaster.AccountType = "06"
        Load FrmTitlePrinterMaster
        If Err.Number <> 364 Then
            FrmTitlePrinterMaster.Caption = "Title Printer Master"
            FrmTitlePrinterMaster.Show
        End If
    End If
End Sub
Private Sub MnuLaminator_Click()
    On Error Resume Next
    If Not IsFormLoaded("Laminator Master") Then
        Dim FrmLaminatorMaster As New FrmAccountMaster
        FrmLaminatorMaster.AccountType = "07"
        Load FrmLaminatorMaster
        If Err.Number <> 364 Then
            FrmLaminatorMaster.Caption = "Laminator Master"
            FrmLaminatorMaster.Show
        End If
    End If
End Sub
Private Sub MnuBinder_Click()
    On Error Resume Next
    If Not IsFormLoaded("Binder Master") Then
        Dim FrmBinderMaster As New FrmAccountMaster
        FrmBinderMaster.AccountType = "08"
        Load FrmBinderMaster
        If Err.Number <> 364 Then
            FrmBinderMaster.Caption = "Binder Master"
            FrmBinderMaster.Show
        End If
    End If
End Sub
Private Sub MnuProcessor_Click()
    On Error Resume Next
    If Not IsFormLoaded("Processor Master") Then
        Dim FrmProcessorMaster As New FrmAccountMaster
        FrmProcessorMaster.AccountType = "04"
        Load FrmProcessorMaster
        If Err.Number <> 364 Then
            FrmProcessorMaster.Caption = "Processor Master"
            FrmProcessorMaster.Show
        End If
    End If
End Sub
Private Sub MnuGodown_Click()
    On Error Resume Next
    If Not IsFormLoaded("Godown Master") Then
        Dim FrmGodownMaster As New FrmAccountMaster
        FrmGodownMaster.AccountType = "09"
        Load FrmGodownMaster
        If Err.Number <> 364 Then
            FrmGodownMaster.Caption = "Godown Master"
            FrmGodownMaster.Show
        End If
    End If
End Sub
Private Sub MnuPaper_Click()
    On Error Resume Next
    Load FrmPaperMaster
    If Err.Number <> 364 Then
        FrmPaperMaster.Show
    End If
End Sub
Private Sub MnuOutsourceItem_Click()
    On Error Resume Next
    Load FrmOutsourceItemMaster
    If Err.Number <> 364 Then
        FrmOutsourceItemMaster.Show
    End If
End Sub
Private Sub MnuEditorial_Click()
    On Error Resume Next
    Load FrmEditorialMaster
    If Err.Number <> 364 Then FrmEditorialMaster.Show
End Sub
Private Sub MnuFreshBook_Click()
    On Error Resume Next
    FrmBookMaster.BookType = "F"
    Load FrmBookMaster
    If Err.Number <> 364 Then
        FrmBookMaster.Show
    End If
End Sub
Private Sub MnuRepairBook_Click()
    On Error Resume Next
    FrmBookMaster.BookType = "R"
    Load FrmBookMaster
    If Err.Number <> 364 Then
        FrmBookMaster.Show
    End If
End Sub
Private Sub MnuUser_Click()
    On Error Resume Next
    Load FrmUserMaster
    If Err.Number <> 364 Then
        FrmUserMaster.Show
    End If
End Sub
Private Sub MnuBookPaperPurchaseOrder_Click()
    On Error Resume Next
    If Not IsFormLoaded("Paper Purchase Order [Book]") Then
        Dim FrmBookPaperPurchaseOrder As New FrmPaperPurchaseOrder
        FrmBookPaperPurchaseOrder.OrderType = "1"
        Load FrmBookPaperPurchaseOrder
        If Err.Number <> 364 Then
            FrmBookPaperPurchaseOrder.Show
        End If
    End If
End Sub
Private Sub MnuTitlePaperPurchaseOrder_Click()
    On Error Resume Next
    If Not IsFormLoaded("Paper Purchase Order [Title]") Then
        Dim FrmTitlePaperPurchaseOrder As New FrmPaperPurchaseOrder
        FrmTitlePaperPurchaseOrder.OrderType = "2"
        Load FrmTitlePaperPurchaseOrder
        If Err.Number <> 364 Then
            FrmTitlePaperPurchaseOrder.Show
        End If
    End If
End Sub
Private Sub MnuBookPaperMovement_Click()
    On Error Resume Next
    If Not IsFormLoaded("Paper Movement [Book]") Then
        Dim FrmBookPaperMovement As New FrmPaperMovement
        FrmBookPaperMovement.MovementType = "1"
        Load FrmBookPaperMovement
        If Err.Number <> 364 Then
            FrmBookPaperMovement.Show
        End If
    End If
End Sub
Private Sub MnuTitlePaperMovement_Click()
    On Error Resume Next
    If Not IsFormLoaded("Paper Movement [Title]") Then
        Dim FrmTitlePaperMovement As New FrmPaperMovement
        FrmTitlePaperMovement.MovementType = "2"
        Load FrmTitlePaperMovement
        If Err.Number <> 364 Then
            FrmTitlePaperMovement.Show
        End If
    End If
End Sub
Private Sub MnuOutsourceItemPurchaseOrder_Click()
    On Error Resume Next
    Load FrmOutsourceItemPurchaseOrder
    If Err.Number <> 364 Then
        FrmOutsourceItemPurchaseOrder.Show
    End If
End Sub

Private Sub MnuMaterialIssueOrder_Click()
    On Error Resume Next
    Load FrmMaterialIssueOrder
    If Err.Number <> 364 Then
        FrmMaterialIssueOrder.Show
    End If
End Sub
Private Sub MnuMaterialMovement_Click()
    On Error Resume Next
    Load FrmMaterialMovement
    If Err.Number <> 364 Then
        FrmMaterialMovement.Show
    End If
End Sub
Private Sub MnuStockJournal_Click()
    On Error Resume Next
    Load FrmStockJournal
    If Err.Number <> 364 Then
        FrmStockJournal.Show
    End If
End Sub
Private Sub MnuBookPrintPlanning_Click()
    On Error Resume Next
    If Not IsFormLoaded("Print Planning [Book]") Then
        Dim FrmBookPrintPlanning As New FrmPrintPlanning
        FrmBookPrintPlanning.PlanningType = "1"
        Load FrmBookPrintPlanning
        If Err.Number <> 364 Then
            FrmBookPrintPlanning.Show
        End If
    End If
End Sub
Private Sub MnuTitlePrintPlanning_Click()
    On Error Resume Next
    If Not IsFormLoaded("Print Planning [Title]") Then
        Dim FrmTitlePrintPlanning As New FrmPrintPlanning
        FrmTitlePrintPlanning.PlanningType = "2"
        Load FrmTitlePrintPlanning
        If Err.Number <> 364 Then
            FrmTitlePrintPlanning.Show
        End If
    End If
End Sub
Private Sub MnuFreshBookPrintOrder_Click()
    On Error Resume Next
    If Not IsFormLoaded("Book Print Order [Fresh]") And Not IsFormLoaded("Book Print Order [Repair]") And Not IsFormLoaded("Cost Sheet") Then
        FrmBookPrintOrder.BookPOType = "F"
        Load FrmBookPrintOrder
        If Err.Number <> 364 Then
            FrmBookPrintOrder.Show
        End If
    End If
End Sub
Private Sub MnuRepairBookPrintOrder_Click()
    On Error Resume Next
    If Not IsFormLoaded("Book Print Order [Fresh]") And Not IsFormLoaded("Book Print Order [Repair]") And Not IsFormLoaded("Cost Sheet") Then
        FrmBookPrintOrder.BookPOType = "R"
        Load FrmBookPrintOrder
        If Err.Number <> 364 Then
            FrmBookPrintOrder.Show
        End If
    End If
End Sub
Private Sub MnuBookDebitNote_Click()
    On Error Resume Next
    Load FrmBookDebitNote
    If Err.Number <> 364 Then FrmBookDebitNote.Show
End Sub
Private Sub MnuPaperDebitNote_Click()
    On Error Resume Next
    Load FrmPaperDebitNote
    If Err.Number <> 364 Then FrmPaperDebitNote.Show
End Sub
'Private Sub MnuBookProcessOrder01_Click()   'Title
'    On Error Resume Next
'    If Not IsFormLoaded("Book Process Order [Fresh]") Then
'        Dim FrmBookProcessOrder01 As New FrmBookProcessOrder
'        FrmBookProcessOrder01.OrderType = "F"
'        Load FrmBookProcessOrder01
'        If Err.Number <> 364 Then
'            FrmBookProcessOrder01.Show
'        End If
'    End If
'End Sub
'Private Sub MnuBookProcessOrder02_Click()   'Title
'    On Error Resume Next
'    If Not IsFormLoaded("Book Process Order [Repair]") Then
'        Dim FrmBookProcessOrder02 As New FrmBookProcessOrder
'        FrmBookProcessOrder02.OrderType = "T"
'        Load FrmBookProcessOrder02
'        If Err.Number <> 364 Then
'            FrmBookProcessOrder02.Show
'        End If
'    End If
'End Sub
'Private Sub MnuBookProcessOrder03_Click()   'Title
'    On Error Resume Next
'    If Not IsFormLoaded("Book Process Order [Title]") Then
'        Dim FrmBookProcessOrder03 As New FrmBookProcessOrder
'        FrmBookProcessOrder03.OrderType = "T"
'        Load FrmBookProcessOrder03
'        If Err.Number <> 364 Then
'            FrmBookProcessOrder03.Show
'        End If
'    End If
'End Sub

Private Sub MnuCostSheet_Click()
    On Error Resume Next
    If Not IsFormLoaded("Book Print Order [Fresh]") And Not IsFormLoaded("Book Print Order [Repair]") And Not IsFormLoaded("Cost Sheet") Then
        FrmBookPrintOrder.BookPOType = "O"
        Load FrmBookPrintOrder
        If Err.Number <> 364 Then
            FrmBookPrintOrder.Show
        End If
    End If
End Sub
Private Sub MnuDayBook_Click()
    On Error Resume Next
    Load FrmDayBook
    If Err.Number <> 364 Then FrmDayBook.Show
End Sub
Private Sub MnuPendingPaymentRegister_Click()
    On Error Resume Next
    Load FrmPendingPaymentRegister
    If Err.Number <> 364 Then
        FrmPendingPaymentRegister.Show
    End If
End Sub
Private Sub MnuProductionPlanning01_Click()
    On Error Resume Next
    FrmProductionPlanning.OrderType = "M"
    Load FrmProductionPlanning
    
    If Err.Number <> 364 Then
        FrmProductionPlanning.Show
    End If
    
End Sub
Private Sub MnuProductionPlanning02_Click()
    On Error Resume Next
    FrmProductionPlanning.OrderType = "S"
    Load FrmProductionPlanning
    
    If Err.Number <> 364 Then
        FrmProductionPlanning.Show
    End If
    
End Sub
Private Sub MnuIOStatusUpdation_Click()
    Dim oExcel As Object
    Dim i As Long
    On Error GoTo ErrorHandler
    If Not FileExist(App.Path & "\Report\Paper Issue Register (" & CompCode & ").xlsx") Then DisplayError ("Failed to Update the Paper Issue Order(s) Status"):          Exit Sub
    If MsgBox("Are you sure to Proceed?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Proceed !") = vbYes Then
        Screen.MousePointer = vbHourglass
        DoEvents
        Set oExcel = CreateObject("Excel.Application")
        oExcel.Workbooks.Open (App.Path & "\Report\Paper Issue Register (" & CompCode & ")")
        CxnDatabase.BeginTrans
        For i = 5 To 1048576
            If Trim(oExcel.Application.Cells(i, 16384)) = "" Then Exit For
            CxnDatabase.Execute "UPDATE PaperIOChild SET Narration='" & Trim(oExcel.Application.Cells(i, 10)) & "' WHERE Code='" & Trim(oExcel.Application.Cells(i, 16384)) & "' AND Paper='" & oExcel.Application.Cells(i, 16382) & "' AND Account='" & oExcel.Application.Cells(i, 16383) & "'"
        Next
        CxnDatabase.CommitTrans
        Call MsgBox("Successfully Updated the Paper Issue Order(s) Status !", vbInformation, App.Title)
        oExcel.Workbooks.Close
        Set oExcel = Nothing
        Screen.MousePointer = vbNormal
    End If
    Exit Sub
ErrorHandler:
    
    CxnDatabase.RollbackTrans
    oExcel.Workbooks.Close
    Set oExcel = Nothing
    Screen.MousePointer = vbNormal
    DisplayError ("Failed to Update the Paper Issue Order(s) Status")
    
End Sub
Private Sub MnuPOStatusUpdation01_Click()   'Title Printing Status Updation
    Dim oExcel As Object
    Dim i As Long, K As Long
    On Error GoTo ErrorHandler
    If Not FileExist(App.Path & "\Report\Print Order Status Register (Title Printerwise) (" & CompCode & ").xlsx") Then DisplayError ("Failed to Update the Title Printing Order(s) Status"): Exit Sub
    If MsgBox("Are you sure to Proceed?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Proceed !") = vbYes Then
        Screen.MousePointer = vbHourglass
        DoEvents
        Set oExcel = CreateObject("Excel.Application")
        oExcel.Workbooks.Open (App.Path & "\Report\Print Order Status Register (Title Printerwise) (" & CompCode & ")")
        CxnDatabase.BeginTrans
        For i = 5 To 1048576
            If Trim(oExcel.Application.Cells(i, "XFB")) = "" Then Exit For
            For K = 39 To 54
                If Trim(oExcel.Application.Cells(i, "Q")) <> "Not Reqd." Then If Trim(oExcel.Application.Cells(i, "Q")) = Trim(oExcel.Application.Cells(K, "XFD")) Then CxnDatabase.Execute "UPDATE BookPOChild06 SET Status='" & Trim(oExcel.Application.Cells(K, "XFC")) & "',Narration='" & Trim(oExcel.Application.Cells(i, "T")) & "' WHERE Code='" & Trim(oExcel.Application.Cells(i, "XFB")) & "'": Exit For
            Next
        Next
        For i = 5 To 1048576
            If Trim(oExcel.Application.Cells(i, "XFB")) = "" Then Exit For
            For K = 22 To 37
                If Trim(oExcel.Application.Cells(i, "O")) <> "Not Reqd." Then If Trim(oExcel.Application.Cells(i, "O")) = Trim(oExcel.Application.Cells(K, "XFD")) Then CxnDatabase.Execute "UPDATE BookPOChild05 SET Status='" & Trim(oExcel.Application.Cells(K, "XFC")) & "' WHERE Code='" & Trim(oExcel.Application.Cells(i, "XFB")) & "'": Exit For
            Next
        Next
        For i = 5 To 1048576
            If Trim(oExcel.Application.Cells(i, "XFB")) = "" Then Exit For
            For K = 6 To 20
                If Trim(oExcel.Application.Cells(i, "S")) <> "Not Reqd." Then If Trim(oExcel.Application.Cells(i, "S")) = Trim(oExcel.Application.Cells(K, "XFD")) Then CxnDatabase.Execute "UPDATE BookPOChild08 SET Status='" & Trim(oExcel.Application.Cells(K, "XFC")) & "' WHERE Code='" & Trim(oExcel.Application.Cells(i, "XFB")) & "'": Exit For
            Next
        Next
        CxnDatabase.CommitTrans
        Call MsgBox("Successfully Updated the Title Printing Order(s) status !", vbInformation, App.Title)
        oExcel.DisplayAlerts = False: oExcel.Workbooks.Close: oExcel.DisplayAlerts = True: Set oExcel = Nothing
        Screen.MousePointer = vbNormal
    End If
    Exit Sub
ErrorHandler:
    CxnDatabase.RollbackTrans
    oExcel.Workbooks.Close: Set oExcel = Nothing
    Screen.MousePointer = vbNormal
    DisplayError ("Failed to update the Title Printing Order(s) status")
End Sub
Private Sub MnuPOStatusUpdation02_Click()   'Book Printing Status Updation
    Dim oExcel As Object
    Dim i As Long, K As Long
    On Error GoTo ErrorHandler
    If Not FileExist(App.Path & "\Report\Print Order Status Register (Book Printerwise) (" & CompCode & ").xlsx") Then DisplayError ("Failed to Update the Book Printing Order(s) Status"): Exit Sub
    If MsgBox("Are you sure to Proceed?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Proceed !") = vbYes Then
        Screen.MousePointer = vbHourglass
        DoEvents
        Set oExcel = CreateObject("Excel.Application")
        oExcel.Workbooks.Open (App.Path & "\Report\Print Order Status Register (Book Printerwise) (" & CompCode & ")")
        CxnDatabase.BeginTrans
        For i = 5 To 1048576
            If Trim(oExcel.Application.Cells(i, "XFB")) = "" Then Exit For
            For K = 22 To 37
                If Trim(oExcel.Application.Cells(i, "O")) <> "Not Reqd." Then If Trim(oExcel.Application.Cells(i, "O")) = Trim(oExcel.Application.Cells(K, "XFD")) Then CxnDatabase.Execute "UPDATE BookPOChild05 SET Status='" & Trim(oExcel.Application.Cells(K, "XFC")) & "',Narration='" & Trim(oExcel.Application.Cells(i, "T")) & "' WHERE Code='" & Trim(oExcel.Application.Cells(i, "XFB")) & "'": Exit For
            Next
        Next
        For i = 5 To 1048576
            If Trim(oExcel.Application.Cells(i, "XFB")) = "" Then Exit For
            
            For K = 39 To 54
                If Trim(oExcel.Application.Cells(i, "Q")) <> "Not Reqd." Then If Trim(oExcel.Application.Cells(i, "Q")) = Trim(oExcel.Application.Cells(K, "XFD")) Then CxnDatabase.Execute "UPDATE BookPOChild06 SET Status='" & Trim(oExcel.Application.Cells(K, "XFC")) & "' WHERE Code='" & Trim(oExcel.Application.Cells(i, "XFB")) & "'": Exit For
            Next
            
        Next
        For i = 5 To 1048576
            If Trim(oExcel.Application.Cells(i, "XFB")) = "" Then Exit For
            For K = 6 To 20
                If Trim(oExcel.Application.Cells(i, "S")) <> "Not Reqd." Then If Trim(oExcel.Application.Cells(i, "S")) = Trim(oExcel.Application.Cells(K, "XFD")) Then CxnDatabase.Execute "UPDATE BookPOChild08 SET Status='" & Trim(oExcel.Application.Cells(K, "XFC")) & "' WHERE Code='" & Trim(oExcel.Application.Cells(i, "XFB")) & "'": Exit For
            Next
        Next
        CxnDatabase.CommitTrans
        Call MsgBox("Successfully Updated the Book Printing Order(s) status !", vbInformation, App.Title)
        oExcel.DisplayAlerts = False: oExcel.Workbooks.Close: oExcel.DisplayAlerts = True: Set oExcel = Nothing
        Screen.MousePointer = vbNormal
    End If
    Exit Sub
ErrorHandler:
    CxnDatabase.RollbackTrans
    oExcel.Workbooks.Close: Set oExcel = Nothing
    Screen.MousePointer = vbNormal
    DisplayError ("Failed to update the Book Printing Order(s) status")
End Sub
Private Sub MnuPOStatusUpdation03_Click()   'Book Binding Status Updation
    Dim oExcel As Object
    Dim i As Long, K As Long
    On Error GoTo ErrorHandler
    
    If Not FileExist(App.Path & "\Report\Print Order Status Register (Book Binderwise) (" & CompCode & ").xlsx") Then DisplayError ("Failed to Update the Book Binding Order(s) Status"): Exit Sub
    
    If MsgBox("Are you sure to Proceed?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Proceed !") = vbYes Then
        Screen.MousePointer = vbHourglass
        DoEvents
        Set oExcel = CreateObject("Excel.Application")
        oExcel.Workbooks.Open (App.Path & "\Report\Print Order Status Register (Book Binderwise) (" & CompCode & ")")
        CxnDatabase.BeginTrans
        For i = 5 To 1048576
            If Trim(oExcel.Application.Cells(i, "XFB")) = "" Then Exit For
            For K = 6 To 20
                If Trim(oExcel.Application.Cells(i, "S")) <> "Not Reqd." Then If Trim(oExcel.Application.Cells(i, "S")) = Trim(oExcel.Application.Cells(K, "XFD")) Then CxnDatabase.Execute "UPDATE BookPOChild08 SET Status='" & Trim(oExcel.Application.Cells(K, "XFC")) & "',Narration='" & Trim(oExcel.Application.Cells(i, "T")) & "' WHERE Code='" & Trim(oExcel.Application.Cells(i, "XFB")) & "'": Exit For
            Next
        Next
        For i = 5 To 1048576
            If Trim(oExcel.Application.Cells(i, "XFB")) = "" Then Exit For
            For K = 39 To 54
                If Trim(oExcel.Application.Cells(i, "Q")) <> "Not Reqd." Then If Trim(oExcel.Application.Cells(i, "Q")) = Trim(oExcel.Application.Cells(K, "XFD")) Then CxnDatabase.Execute "UPDATE BookPOChild06 SET Status='" & Trim(oExcel.Application.Cells(K, "XFC")) & "' WHERE Code='" & Trim(oExcel.Application.Cells(i, "XFB")) & "'": Exit For
            Next
        Next
        For i = 5 To 1048576
            
            If Trim(oExcel.Application.Cells(i, "XFB")) = "" Then Exit For
            For K = 22 To 37
                If Trim(oExcel.Application.Cells(i, "O")) <> "Not Reqd." Then If Trim(oExcel.Application.Cells(i, "O")) = Trim(oExcel.Application.Cells(K, "XFD")) Then CxnDatabase.Execute "UPDATE BookPOChild05 SET Status='" & Trim(oExcel.Application.Cells(K, "XFC")) & "' WHERE Code='" & Trim(oExcel.Application.Cells(i, "XFB")) & "'": Exit For
            Next
            
        Next
        
        CxnDatabase.CommitTrans
        Call MsgBox("Successfully updated the Book Binding Order(s) status !", vbInformation, App.Title)
        oExcel.DisplayAlerts = False: oExcel.Workbooks.Close: oExcel.DisplayAlerts = True: Set oExcel = Nothing
        Screen.MousePointer = vbNormal
        
    End If
    Exit Sub
ErrorHandler:
    CxnDatabase.RollbackTrans
    oExcel.Workbooks.Close: Set oExcel = Nothing
    Screen.MousePointer = vbNormal
    DisplayError ("Failed to update the Book Binding Order(s) status")
    
End Sub
Private Sub MnuPOStatusUpdation04_Click()  'Production Planning Updation
    Dim oExcel As Object
    Dim i As Long
    On Error GoTo ErrorHandler
    If Not FileExist(App.Path & "\Report\Production Planning.xlsx") Then DisplayError ("Failed to Update the Reorder Level"): Exit Sub
    If MsgBox("Are you sure to Proceed?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Proceed !") = vbYes Then
        Screen.MousePointer = vbHourglass
        DoEvents
        Set oExcel = CreateObject("Excel.Application")
        oExcel.Workbooks.Open (App.Path & "\Report\Production Planning")
        oExcel.Sheets("Production Planning (MO)").Activate
        CxnDatabase.BeginTrans
        For i = 7 To 1048576
            If Trim(oExcel.Application.Cells(i, 16384)) = "" Then Exit For
            CxnDatabase.Execute "UPDATE BookMaster SET Remarks='" & oExcel.Application.Cells(i, "S") & "' WHERE Code='" & Trim(oExcel.Application.Cells(i, "XFD")) & "'"
        Next
        CxnDatabase.CommitTrans
        Call MsgBox("Successfully updated the Production Planning !", vbInformation, App.Title)
        oExcel.DisplayAlerts = False
        oExcel.Workbooks.Close
        oExcel.DisplayAlerts = True
        Set oExcel = Nothing
        Screen.MousePointer = vbNormal
    End If
    Exit Sub
ErrorHandler:
    CxnDatabase.RollbackTrans
    oExcel.Workbooks.Close
    Set oExcel = Nothing
    Screen.MousePointer = vbNormal
    DisplayError ("Failed to Update the Production Planning")
End Sub
Private Sub MnuPOStatusUpdation05_Click()
    Dim oExcel As Object
    Dim i As Long, K As Long
    On Error GoTo ErrorHandler
    If Not FileExist(App.Path & "\Report\Pending Debit Note Register (" & CompCode & ").xlsx") Then DisplayError ("Failed to Update the Debit Note(s) Details"): Exit Sub
    If MsgBox("Are you sure to Proceed?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Proceed !") = vbYes Then
        Screen.MousePointer = vbHourglass
        DoEvents
        Set oExcel = CreateObject("Excel.Application")
        oExcel.Workbooks.Open (App.Path & "\Report\Pending Debit Note Register (" & CompCode & ")")
        CxnDatabase.BeginTrans
        For i = 5 To 1048576
            If Trim(oExcel.Application.Cells(i, 16382)) = "" Then Exit For
            For K = 5 To 30
                If Trim(oExcel.Application.Cells(i, 14)) = Trim(oExcel.Application.Cells(K, 16384)) Then
                    CxnDatabase.Execute "UPDATE BookPOChild08 SET Status='" & Trim(oExcel.Application.Cells(K, 16383)) & "',DNDetails='" & Trim(oExcel.Application.Cells(i, 15)) & "',CNDetails='" & Trim(oExcel.Application.Cells(i, 34)) & "' WHERE Code='" & Trim(oExcel.Application.Cells(i, 16382)) & "'"
                    Exit For
                End If
            Next
        Next
        CxnDatabase.CommitTrans
        Call MsgBox("Successfully Updated the Debit Note(s) Details !", vbInformation, App.Title)
        oExcel.DisplayAlerts = False
        oExcel.Workbooks.Close
        oExcel.DisplayAlerts = True
        Set oExcel = Nothing
        Screen.MousePointer = vbNormal
    End If
    Exit Sub
ErrorHandler:
    CxnDatabase.RollbackTrans
    oExcel.Workbooks.Close
    Set oExcel = Nothing
    Screen.MousePointer = vbNormal
    DisplayError ("Failed to Update the Debit Note(s) Details")
End Sub
Private Sub MnuBookReceipt_Click()
    If Trim(ReadFromFile("Book Receipt")) = "" Or Trim(ReadFromFile("Book Receipt")) = "N" Then Exit Sub
    Dim CxnImporter As New ADODB.Connection
    Dim rstImporter As New ADODB.Recordset
    Dim DatabaseName As String
    Dim i As Integer
    On Error GoTo ErrorHandler
    DatabaseName = Trim(ReadFromFile("Busy Database Name")): i = 0
    If ServerName = "" Or DatabaseName = "" Then Exit Sub
    Screen.MousePointer = vbHourglass
    CxnDatabase.BeginTrans
    CxnImporter.CursorLocation = adUseClient
    CxnDatabase.Execute "UPDATE BookPOParent SET ReceivedQuantity=0"
    Do While True
        
        Dim str As String
        i = InStr(1, DatabaseName, ",")
        If CxnImporter.State = adStateOpen Then CxnImporter.Close
        If i = 0 Then CxnImporter.Open "Provider=SQLOLEDB.1;Password=" & ServerPassword & ";Persist Security Info=True;User ID=sa;Initial Catalog=" & Mid(DatabaseName, 1) & ";Data Source=" & ServerName Else CxnImporter.Open "Provider=SQLOLEDB.1;Password=" & ServerPassword & ";Persist Security Info=True;User ID=sa;Initial Catalog=" & Mid(DatabaseName, 1, i - 1) & ";Data Source=" & ServerName
        If rstImporter.State = adStateOpen Then rstImporter.Close
        
        rstImporter.Open "SELECT * FROM (SELECT RefCode,No,Date,MasterCode1,SUM(ABS(Value1)) As OrderedQuantity,(SELECT ISNULL(SUM(ABS(Value1)),0) FROM Tran3 WHERE (VchType=2 OR VchType=4) AND RecType=4 AND RefCode=T.RefCode) As ReceivedQuantity FROM Tran3 T WHERE VchType=13 GROUP BY RefCode,No,Date,MasterCode1) As Tbl WHERE ReceivedQuantity>0 AND ISNUMERIC(No)=1 ORDER BY No", CxnImporter, adOpenKeyset, adLockReadOnly        'MasterCode1=BookCode

        rstImporter.ActiveConnection = Nothing
        If rstImporter.RecordCount > 0 Then rstImporter.MoveFirst
        Do While Not rstImporter.EOF
            DoEvents
            MdiMainMenu.StatusBar1.Panels(2).Text = "Updating PO #" & Trim(rstImporter.Fields("No").Value) & " !!!"
            CxnDatabase.Execute "UPDATE BookPOParent P,BookPOChild05 C SET C.Status='D' WHERE P.Code=C.Code AND IIF(LEFT(P.Code,1)='*',MID(TRIM(Name),2),TRIM(Name))='" & Trim(rstImporter.Fields("No").Value) & "' AND P.Type<>'O' AND Format(Date,'dd-MMM-yyyy')='" & Format(rstImporter.Fields("Date").Value, "dd-MMM-yyyy") & "'"
            CxnDatabase.Execute "UPDATE BookPOParent P,BookPOChild06 C SET C.Status='D' WHERE P.Code=C.Code AND IIF(LEFT(P.Code,1)='*',MID(TRIM(Name),2),TRIM(Name))='" & Trim(rstImporter.Fields("No").Value) & "' AND P.Type<>'O' AND Format(Date,'dd-MMM-yyyy')='" & Format(rstImporter.Fields("Date").Value, "dd-MMM-yyyy") & "'"
            CxnDatabase.Execute "UPDATE BookPOParent SET ReceivedQuantity=ReceivedQuantity+" & Val(rstImporter.Fields("ReceivedQuantity").Value) & " WHERE IIF(LEFT(Code,1)='*',MID(TRIM(Name),2),TRIM(Name))='" & Trim(rstImporter.Fields("No").Value) & "' AND Type<>'O' AND Format(Date,'dd-MMM-yyyy')='" & Format(rstImporter.Fields("Date").Value, "dd-MMM-yyyy") & "'"
            CxnDatabase.Execute "UPDATE BookPOParent SET BPODStatus=1,TPODStatus=1,TLODStatus=1,BBODStatus=1 WHERE IIF(LEFT(Code,1)='*',MID(TRIM(Name),2),TRIM(Name))='" & Trim(rstImporter.Fields("No").Value) & "' AND Type<>'O' AND Format(Date,'dd-MMM-yyyy')='" & Format(rstImporter.Fields("Date").Value, "dd-MMM-yyyy") & "'"
            rstImporter.MoveNext
        Loop
        'Price Updation
'        If rstImporter.State = adStateOpen Then rstImporter.Close
'        rstImporter.Open "SELECT Alias,D2 FROM Master1 WHERE MasterType=6 AND Alias<>'' AND (LEFT(UPPER(Name),2)<>'Z_' AND LEFT(UPPER(Name),2)<>'Z-') ORDER BY Alias", CxnImporter, adOpenKeyset, adLockReadOnly
'        rstImporter.ActiveConnection = Nothing
'        If rstImporter.RecordCount > 0 Then rstImporter.MoveFirst
'        Do While Not rstImporter.EOF
'            CxnDatabase.Execute "UPDATE BookMaster SET Price=" & Val(rstImporter.Fields("D2").Value) & " WHERE LEFT(BusyCode,6)='" & Left(rstImporter.Fields("Alias").Value, 6) & "'"
'            rstImporter.MoveNext
'        Loop
        If i = 0 Then Exit Do Else DatabaseName = Mid(DatabaseName, i + 1): i = 0
    Loop
    MdiMainMenu.StatusBar1.Panels(2).Text = ""
    CxnDatabase.Execute "UPDATE BookPOParent P,BookPOChild08 C SET C.Status='' WHERE P.Code=C.Code AND P.Type<>'O' AND C.Status='D'"
    CxnDatabase.Execute "UPDATE BookPOParent P,BookPOChild08 C SET C.Status='D' WHERE P.Code=C.Code AND P.Type<>'O' AND (P.ReceivedQuantity+C.AdjustQuantity>=C.ActualQuantity-INT(C.ActualQuantity*0.2/100) OR INT(C.ActualQuantity*0.2/100)+C.AdjustQuantity>=C.ActualQuantity)"
    CxnDatabase.Execute "UPDATE BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code SET P.ReceivedQuantity=P.ReceivedQuantity+C.AdjustQuantity"
    CxnDatabase.CommitTrans
    On Error Resume Next
    If CxnImporter.State = adStateOpen Then
        Dim RecordsAffected As Long
        If rstImporter.State = adStateOpen Then rstImporter.Close
        rstImporter.Open "SELECT TRIM(P.Name) As VchNo,Date As VchDate,M.Alias As Laminator FROM ((BookPOParent P INNER JOIN BookPOChild08 C1 ON P.Code=C1.Code) INNER JOIN BookPOChild07 C2 ON P.Code=C2.Code) INNER JOIN AccountMaster M ON P.Laminator=M.Code WHERE P.Type<>'O' AND LEFT(P.Code,1)<>'*' ORDER BY P.Name", CxnDatabase, adOpenKeyset, adLockReadOnly
        rstImporter.ActiveConnection = Nothing
        If rstImporter.RecordCount > 0 Then rstImporter.MoveFirst
        Do While Not rstImporter.EOF
            DoEvents
            MdiMainMenu.StatusBar1.Panels(2).Text = "Updating Alias for PO #" & Trim(rstImporter.Fields("VchNo").Value) & " !!!"
            CxnImporter.Execute "UPDATE VchOtherInfo SET OF1='" & rstImporter.Fields("Laminator").Value & "' WHERE VchCode IN (SELECT VchCode FROM Tran1 WHERE VchType=13 AND LTRIM(VchNo)='" & rstImporter.Fields("VchNo").Value & "' AND Date='" & Format(rstImporter.Fields("VchDate").Value, "dd-MMM-yyyy") & "')", RecordsAffected
            If RecordsAffected = 0 Then CxnImporter.Execute "INSERT INTO VchOtherInfo (VchCode,OF1) VALUES ((SELECT VchCode FROM Tran1 WHERE VchType=13 AND LTRIM(VchNo)='" & rstImporter.Fields("VchNo").Value & "' AND Date='" & Format(rstImporter.Fields("VchDate").Value, "dd-MMM-yyyy") & "'),'" & rstImporter.Fields("Laminator").Value & "')"
            rstImporter.MoveNext
        Loop
    End If
    Call CloseRecordset(rstImporter)
    Call CloseConnection(CxnImporter)
    Screen.MousePointer = vbNormal
    Exit Sub
    
ErrorHandler:
    CxnDatabase.RollbackTrans
    Screen.MousePointer = vbNormal
    DisplayError ("Failed to import the Book Receipts")
    Call CloseRecordset(rstImporter)
    Call CloseConnection(CxnImporter)
End Sub
Private Sub MnuBookPrintPlanningRegister_Click()
    On Error Resume Next
    If Not IsFormLoaded("Print Planning Register [Book]") Then
        Dim FrmBookPrintPlanningRegister As New FrmPrintPlanningRegister
        FrmBookPrintPlanningRegister.PlanningType = "1"
        Load FrmBookPrintPlanningRegister
        If Err.Number <> 364 Then
            FrmBookPrintPlanningRegister.Show
        End If
    End If
End Sub
Private Sub MnuTitlePrintPlanningRegister_Click()
    On Error Resume Next
    If Not IsFormLoaded("Print Planning Register [Title]") Then
        Dim FrmTitlePrintPlanningRegister As New FrmPrintPlanningRegister
        FrmTitlePrintPlanningRegister.PlanningType = "2"
        Load FrmTitlePrintPlanningRegister
        If Err.Number <> 364 Then
            FrmTitlePrintPlanningRegister.Show
        End If
    End If
End Sub
Private Sub MnuPOStatusRegister01_Click()
    On Error Resume Next
    
    If Not IsFormLoaded("Print Order Status Register [Bookwise]") Then
        Dim FrmPrintOrderStatusRegister01 As New FrmPrintOrderStatusRegister
        FrmPrintOrderStatusRegister01.OrderType = "01"
        Load FrmPrintOrderStatusRegister01
        If Err.Number <> 364 Then
            FrmPrintOrderStatusRegister01.Show
        End If
    End If
    
End Sub
Private Sub MnuPOStatusRegister05_Click()
    On Error Resume Next
    
    If Not IsFormLoaded("Print Order Status Register [Print Orderwise]") Then
        Dim FrmPrintOrderStatusRegister05 As New FrmPrintOrderStatusRegister
        FrmPrintOrderStatusRegister05.OrderType = "02"
        Load FrmPrintOrderStatusRegister05
        If Err.Number <> 364 Then
            FrmPrintOrderStatusRegister05.Show
        End If
    End If
    
End Sub
Private Sub MnuPOStatusRegister06_Click()
    On Error Resume Next
    If Not IsFormLoaded("Print Order Status Register [Busy]") Then
        Dim FrmPrintOrderStatusRegisterXX As New FrmPrintOrderStatusRegister
        FrmPrintOrderStatusRegisterXX.OrderType = "XX"
        Load FrmPrintOrderStatusRegisterXX
        If Err.Number <> 364 Then FrmPrintOrderStatusRegisterXX.Show
    End If
End Sub
Private Sub MnuPendingDNRegister_Click()
    On Error Resume Next
    If Not IsFormLoaded("Print Order Status Register [Debit Note]") Then
        Dim FrmPrintOrderStatusRegisterYY As New FrmPrintOrderStatusRegister
        FrmPrintOrderStatusRegisterYY.OrderType = "YY"
        Load FrmPrintOrderStatusRegisterYY
        If Err.Number <> 364 Then FrmPrintOrderStatusRegisterYY.Show
    End If
End Sub
Private Sub MnuPOStatusRegister02_Click()
    On Error Resume Next
    If Not IsFormLoaded("Print Order Status Register [Title Printerwise]") Then
        Dim FrmPrintOrderStatusRegister02 As New FrmPrintOrderStatusRegister
        FrmPrintOrderStatusRegister02.OrderType = "06"
        Load FrmPrintOrderStatusRegister02
        If Err.Number <> 364 Then FrmPrintOrderStatusRegister02.Show
    End If
End Sub
Private Sub MnuPOStatusRegister03_Click()
    On Error Resume Next
    If Not IsFormLoaded("Print Order Status Register [Book Printerwise]") Then
        Dim FrmPrintOrderStatusRegister03 As New FrmPrintOrderStatusRegister
        FrmPrintOrderStatusRegister03.OrderType = "05"
        Load FrmPrintOrderStatusRegister03
        If Err.Number <> 364 Then
            FrmPrintOrderStatusRegister03.Show
        End If
    End If
End Sub
Private Sub MnuPOStatusRegister04_Click()
    On Error Resume Next
    If Not IsFormLoaded("Print Order Status Register [Book Binderwise]") Then
        Dim FrmPrintOrderStatusRegister04 As New FrmPrintOrderStatusRegister
        FrmPrintOrderStatusRegister04.OrderType = "08"
        Load FrmPrintOrderStatusRegister04
        If Err.Number <> 364 Then
            FrmPrintOrderStatusRegister04.Show
        End If
    End If
End Sub
Private Sub MnuPOStatusReg02_Click()
    On Error Resume Next
    Load FrmOutsourceItemSupplierRegister
    If Err.Number <> 364 Then
        FrmOutsourceItemSupplierRegister.Show
    End If
End Sub
Private Sub MnuPOStatusReg0301_Click()
    On Error Resume Next
    If Not IsFormLoaded("Insource Item [Fresh Book] Purchase Order Status Register") Then
        Dim FrmInSourceItem03SupplierRegister As New FrmInsourceItemSupplierRegister
        FrmInSourceItem03SupplierRegister.ItemType = "3"
        Load FrmInSourceItem03SupplierRegister
        If Err.Number <> 364 Then
            FrmInSourceItem03SupplierRegister.Show
        End If
    End If
End Sub
Private Sub MnuPOStatusReg0302_Click()
    On Error Resume Next
    If Not IsFormLoaded("Insource Item [Repair Book] Purchase Order Status Register") Then
        Dim FrmInSourceItem04SupplierRegister As New FrmInsourceItemSupplierRegister
        FrmInSourceItem04SupplierRegister.ItemType = "4"
        Load FrmInSourceItem04SupplierRegister
        If Err.Number <> 364 Then
            FrmInSourceItem04SupplierRegister.Show
        End If
    End If
End Sub
Private Sub MnuPOStatusReg0303_Click()
    On Error Resume Next
    If Not IsFormLoaded("Insource Item [Title] Purchase Order Status Register") Then
        Dim FrmInSourceItem05SupplierRegister As New FrmInsourceItemSupplierRegister
        FrmInSourceItem05SupplierRegister.ItemType = "5"
        Load FrmInSourceItem05SupplierRegister
        If Err.Number <> 364 Then
            FrmInSourceItem05SupplierRegister.Show
        End If
    End If
End Sub
Private Sub MnuPaperIssueRegister_Click()
    On Error Resume Next
    Load FrmPaperIssueRegister
    If Err.Number <> 364 Then
        FrmPaperIssueRegister.Show
    End If
End Sub
Private Sub MnuPaperStockRegister_Click()
    On Error Resume Next
    Load FrmPaperStockRegister
    If Err.Number <> 364 Then FrmPaperStockRegister.Show
End Sub
Private Sub MnuMaterialStockRegister01_Click()
    On Error Resume Next
    If Not IsFormLoaded("Material Stock Register [Binderwise/Bookwise/Itemwise]") Then
        Dim FrmMaterialStockRegister01 As New FrmMaterialStockRegister
        FrmMaterialStockRegister01.ReportType = "1"
        Load FrmMaterialStockRegister01
        If Err.Number <> 364 Then
            FrmMaterialStockRegister01.Show
        End If
    End If
End Sub
Private Sub MnuMaterialStockRegister02_Click()
    On Error Resume Next
    If Not IsFormLoaded("Material Stock Register [Binderwise/Itemwise]") Then
        Dim FrmMaterialStockRegister02 As New FrmMaterialStockRegister
        FrmMaterialStockRegister02.ReportType = "2"
        Load FrmMaterialStockRegister02
        If Err.Number <> 364 Then
            FrmMaterialStockRegister02.Show
        End If
    End If
End Sub
Private Sub MnuBillRegister_Click()
    On Error Resume Next
    Load FrmBillRegister
    If Err.Number <> 364 Then FrmBillRegister.Show
End Sub
Private Sub MnuBookPOPrintUtility_Click()
    On Error Resume Next
    Load FrmBookPOPrintUtility
    If Err.Number <> 364 Then FrmBookPOPrintUtility.Show
End Sub
Private Sub MnuPaperPOPrintUtility_Click()
    On Error Resume Next
    Load FrmPaperPOPrintUtility
    If Err.Number <> 364 Then FrmPaperPOPrintUtility.Show
End Sub
Private Sub MnuOpBal_Click()
    Dim oExcel As Object
    Dim i As Long, Cnt As Long
    Dim rstPaperOpBal As New ADODB.Recordset
    Dim rstCompanyMaster As New ADODB.Recordset
    On Error Resume Next
    
    If Not FileExist(App.Path & "\Template\Opening Balance.xlsx") Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    If rstPaperOpBal.State = adStateOpen Then
        rstPaperOpBal.Close
    End If
    rstCompanyMaster.Open "Select PrintName From CompanyMaster", CxnDatabase, adOpenKeyset, adLockReadOnly
    rstPaperOpBal.Open "SELECT M2.PrintName As GodownName,M1.PrintName As PaperName,[Weight/Ream],C.OpBalOther,C.OpBalSheets,C.OpBalTat FROM (PaperChild C INNER JOIN PaperMaster M1 ON M1.Code=C.Code) INNER JOIN AccountMaster M2 ON M2.Code=C.Account ORDER BY M2.PrintName,M1.PrintName", CxnDatabase, adOpenKeyset, adLockReadOnly
    If rstPaperOpBal.RecordCount = 0 Then
        Screen.MousePointer = vbNormal
        On Error GoTo 0
        Exit Sub
    End If
    DoEvents
    Set oExcel = CreateObject("Excel.Application")
    oExcel.Workbooks.Open (App.Path & "\Template\Opening Balance")
    oExcel.DisplayAlerts = False
    oExcel.Workbooks.Item(1).SaveAs (App.Path & "\Report\Opening Balance (" & CompCode & ")")
    oExcel.DisplayAlerts = True
    oExcel.Sheets("Sheet1").Select
    oExcel.Visible = False
    oExcel.Cells(1, 1).Value = Trim(rstCompanyMaster.Fields("PrintName").Value)
    oExcel.Cells(2, 1).Value = "Opening Balance As On [" & Format(FinancialYearFrom, "dd-mm-yyyy") & "]"
    i = 4
    Cnt = 1
    Do While Not rstPaperOpBal.EOF
        oExcel.Application.Cells(i, 1).Value = Trim(rstPaperOpBal.Fields("GodownName").Value)
        oExcel.Application.Cells(i, 2).Value = Trim(rstPaperOpBal.Fields("PaperName").Value)
        oExcel.Application.Cells(i, 3).Value = Val(rstPaperOpBal.Fields("OpBalOther").Value)
        oExcel.Application.Cells(i, 4).Value = Val(rstPaperOpBal.Fields("Weight/Ream").Value)
        oExcel.Application.Cells(i, 5).Value = Val(rstPaperOpBal.Fields("OpBalSheets").Value)
        oExcel.Application.Cells(i, 6).Value = Round(Val(rstPaperOpBal.Fields("OpBalSheets").Value) / 500, 3)
        oExcel.Application.Cells(i, 7).Value = Val(rstPaperOpBal.Fields("TatOpBal").Value)
        oExcel.Application.Cells(i, 9).Value = Val(oExcel.Application.Cells(i, 6).Value) * Val(oExcel.Application.Cells(i, 4).Value)
        Cnt = Cnt + 1
        i = i + 1
        rstPaperOpBal.MoveNext
    Loop
    oExcel.Sheets("Sheet1").Activate
    oExcel.Columns("A:J").EntireColumn.AutoFit
    oExcel.Workbooks.Item(1).Save
    Screen.MousePointer = vbNormal
    oExcel.Range("A1").Activate
    oExcel.Visible = True
    Set oExcel = Nothing
    Call CloseRecordset(rstCompanyMaster)
    Call CloseRecordset(rstPaperOpBal)
    On Error GoTo 0
End Sub
Private Sub MnuBookList_Click()
    On Error Resume Next
    Load FrmBookList
    If Err.Number <> 364 Then
        FrmBookList.Show
    End If
End Sub
Private Sub MnuCorrectionList01_Click()
    On Error Resume Next
    Load FrmCorrectionList
    FrmCorrectionList.Department = "P"
    If Err.Number <> 364 Then FrmCorrectionList.Show
End Sub
Private Sub MnuCorrectionList02_Click()
    On Error Resume Next
    Load FrmCorrectionList
    FrmCorrectionList.Department = "E"
    If Err.Number <> 364 Then FrmCorrectionList.Show
End Sub
Private Sub MnuImportBal01_Click()  'Print Order
    Dim CxnImporter As New ADODB.Connection
    Dim rstCompanyMaster As New ADODB.Recordset
    Dim rstImporter00 As New ADODB.Recordset
    Dim rstImporter05 As New ADODB.Recordset
    Dim rstImporter06 As New ADODB.Recordset
    Dim rstImporter07 As New ADODB.Recordset
    Dim rstImporter08 As New ADODB.Recordset
    Dim i As Integer
    Dim SQL As String
    On Error GoTo ErrorHandler
    BusySystemIndicator True
    rstCompanyMaster.Open "Select CreatedFrom From CompanyMaster", CxnDatabase, adOpenKeyset, adLockReadOnly
    If rstCompanyMaster.Fields("CreatedFrom").Value <> "" Then
        If MsgBox("Are you sure to Proceed?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Proceed !") = vbYes Then
            rstCompanyMaster.ActiveConnection = Nothing
            CxnImporter.CursorLocation = adUseClient
            CxnImporter.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DatabasePath & "\Saral." & rstCompanyMaster.Fields("CreatedFrom").Value & ";Persist Security Info=False;Jet OLEDB:Database Password=RSPLILoveMyINDIA"
            rstImporter00.Open "SELECT P.* FROM BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code WHERE Type<>'O' AND LEFT(P.Code,1)<>'*' AND (C.ActualQuantity-P.ReceivedQuantity)>0 AND C.Status NOT IN ('E','D','W') ORDER BY P.Code", CxnImporter, adOpenKeyset, adLockReadOnly
            rstImporter00.ActiveConnection = Nothing
            rstImporter05.Open "SELECT C05.* FROM (BookPOParent P INNER JOIN BookPOChild08 C08 ON P.Code=C08.Code) INNER JOIN BookPOChild05 C05 ON P.Code=C05.Code WHERE Type<>'O' AND LEFT(P.Code,1)<>'*' AND (C08.ActualQuantity-P.ReceivedQuantity)>0 AND C08.Status NOT IN ('E','D','W') ORDER BY P.Code", CxnImporter, adOpenKeyset, adLockReadOnly
            rstImporter05.ActiveConnection = Nothing
            rstImporter06.Open "SELECT C06.* FROM (BookPOParent P INNER JOIN BookPOChild08 C08 ON P.Code=C08.Code) INNER JOIN BookPOChild06 C06 ON P.Code=C06.Code WHERE Type<>'O' AND LEFT(P.Code,1)<>'*' AND (C08.ActualQuantity-P.ReceivedQuantity)>0 AND C08.Status NOT IN ('E','D','W') ORDER BY P.Code", CxnImporter, adOpenKeyset, adLockReadOnly
            rstImporter06.ActiveConnection = Nothing
            rstImporter07.Open "SELECT C07.* FROM (BookPOParent P INNER JOIN BookPOChild08 C08 ON P.Code=C08.Code) INNER JOIN BookPOChild07 C07 ON P.Code=C07.Code WHERE Type<>'O' AND LEFT(P.Code,1)<>'*' AND (C08.ActualQuantity-P.ReceivedQuantity)>0 AND C08.Status NOT IN ('E','D','W') ORDER BY P.Code", CxnImporter, adOpenKeyset, adLockReadOnly
            rstImporter07.ActiveConnection = Nothing
            rstImporter08.Open "SELECT C.* FROM BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code WHERE Type<>'O' AND LEFT(P.Code,1)<>'*' AND (C.ActualQuantity-P.ReceivedQuantity)>0 AND C.Status NOT IN ('E','D','W') ORDER BY P.Code", CxnImporter, adOpenKeyset, adLockReadOnly
            rstImporter08.ActiveConnection = Nothing
            CxnDatabase.BeginTrans
            CxnDatabase.Execute "DELETE FROM BookPOParent WHERE LEFT(Code,1)='*'"
            CxnDatabase.Execute "DELETE FROM BookPOChild05 WHERE LEFT(Code,1)='*'"
            CxnDatabase.Execute "DELETE FROM BookPOChild06 WHERE LEFT(Code,1)='*'"
            CxnDatabase.Execute "DELETE FROM BookPOChild07 WHERE LEFT(Code,1)='*'"
            CxnDatabase.Execute "DELETE FROM BookPOChild08 WHERE LEFT(Code,1)='*'"
            Do While Not rstImporter00.EOF
                SQL = "INSERT INTO BookPOParent VALUES ('*" + Right(rstImporter00.Fields(0).Value, 5) + "','" + Pad("*" + Trim(rstImporter00.Fields(1).Value), Space(1), 10, "L") + "',"
                For i = 2 To rstImporter00.Fields.Count - 1
                    If IsNull(rstImporter00.Fields(i).Value) Then
                        SQL = SQL + "Null,"
                    ElseIf rstImporter00.Fields(i).Type = adVarWChar Then
                        SQL = SQL + "'" + rstImporter00.Fields(i).Value + "',"
                    ElseIf rstImporter00.Fields(i).Type = adDate Then
                        SQL = SQL + "#" + Format(rstImporter00.Fields(i).Value, "mm-dd-yyyy") + "#,"
                    ElseIf rstImporter00.Fields(i).Type = adNumeric Then
                        SQL = SQL + Trim(str(rstImporter00.Fields(i).Value)) + ","
                    ElseIf rstImporter00.Fields(i).Type = adBoolean Then
                        SQL = SQL + Trim(str(rstImporter00.Fields(i).Value)) + ","
                    End If
                Next
                
                SQL = Left(SQL, Len(SQL) - 1)
                SQL = SQL + ")"
                CxnDatabase.Execute SQL
                rstImporter00.MoveNext
            Loop
            Do While Not rstImporter05.EOF
                SQL = "INSERT INTO BookPOChild05 VALUES ('*" + Right(rstImporter05.Fields(0).Value, 5) + "',"
                For i = 1 To rstImporter05.Fields.Count - 1
                    If IsNull(rstImporter05.Fields(i).Value) Then
                        SQL = SQL + "Null,"
                    ElseIf rstImporter05.Fields(i).Type = adVarWChar Then
                        SQL = SQL + "'" + rstImporter05.Fields(i).Value + "',"
                    ElseIf rstImporter05.Fields(i).Type = adDate Then
                        SQL = SQL + "#" + Format(rstImporter05.Fields(i).Value, "mm-dd-yyyy") + "#,"
                    ElseIf rstImporter05.Fields(i).Type = adNumeric Then
                        SQL = SQL + Trim(str(rstImporter05.Fields(i).Value)) + ","
                    End If
                Next
                SQL = Left(SQL, Len(SQL) - 1)
                SQL = SQL + ")"
                CxnDatabase.Execute SQL
                rstImporter05.MoveNext
            Loop
            Do While Not rstImporter06.EOF
                SQL = "INSERT INTO BookPOChild06 VALUES ('*" + Right(rstImporter06.Fields(0).Value, 5) + "',"
                For i = 1 To rstImporter06.Fields.Count - 1
                    If IsNull(rstImporter06.Fields(i).Value) Then
                        SQL = SQL + "Null,"
                    ElseIf rstImporter06.Fields(i).Type = adVarWChar Then
                        SQL = SQL + "'" + rstImporter06.Fields(i).Value + "',"
                    ElseIf rstImporter06.Fields(i).Type = adDate Then
                        SQL = SQL + "#" + Format(rstImporter06.Fields(i).Value, "mm-dd-yyyy") + "#,"
                    ElseIf rstImporter06.Fields(i).Type = adNumeric Then
                        SQL = SQL + Trim(str(rstImporter06.Fields(i).Value)) + ","
                    End If
                Next
                SQL = Left(SQL, Len(SQL) - 1)
                SQL = SQL + ")"
                CxnDatabase.Execute SQL
                rstImporter06.MoveNext
            Loop
            Do While Not rstImporter07.EOF
                SQL = "INSERT INTO BookPOChild07 VALUES ('*" + Right(rstImporter07.Fields(0).Value, 5) + "',"
                For i = 1 To rstImporter07.Fields.Count - 1
                    If IsNull(rstImporter07.Fields(i).Value) Then
                        SQL = SQL + "Null,"
                    ElseIf rstImporter07.Fields(i).Type = adVarWChar Then
                        SQL = SQL + "'" + rstImporter07.Fields(i).Value + "',"
                    ElseIf rstImporter07.Fields(i).Type = adDate Then
                        SQL = SQL + "#" + Format(rstImporter07.Fields(i).Value, "mm-dd-yyyy") + "#,"
                    ElseIf rstImporter07.Fields(i).Type = adNumeric Then
                        SQL = SQL + Trim(str(rstImporter07.Fields(i).Value)) + ","
                    End If
                Next
                SQL = Left(SQL, Len(SQL) - 1)
                SQL = SQL + ")"
                CxnDatabase.Execute SQL
                rstImporter07.MoveNext
            Loop
            Do While Not rstImporter08.EOF
                SQL = "INSERT INTO BookPOChild08 VALUES ('*" + Right(rstImporter08.Fields(0).Value, 5) + "',"
                For i = 1 To rstImporter08.Fields.Count - 1
                    If IsNull(rstImporter08.Fields(i).Value) Then
                        SQL = SQL + "Null,"
                    ElseIf rstImporter08.Fields(i).Type = adVarWChar Then
                        SQL = SQL + "'" + rstImporter08.Fields(i).Value + "',"
                    ElseIf rstImporter08.Fields(i).Type = adDate Then
                        SQL = SQL + "#" + Format(rstImporter08.Fields(i).Value, "mm-dd-yyyy") + "#,"
                    ElseIf rstImporter08.Fields(i).Type = adNumeric Then
                        SQL = SQL + Trim(str(rstImporter08.Fields(i).Value)) + ","
                    End If
                Next
                SQL = Left(SQL, Len(SQL) - 1)
                SQL = SQL + ")"
                CxnDatabase.Execute SQL
                rstImporter08.MoveNext
            Loop
            CxnDatabase.CommitTrans
            Call MsgBox("Successfully imported the Balances !", vbInformation, App.Title)
        End If
    Else
        Call MsgBox("Nothing To Import !", vbInformation, App.Title)
    End If
    Call CloseRecordset(rstImporter00)
    Call CloseRecordset(rstImporter05)
    Call CloseRecordset(rstImporter06)
    Call CloseRecordset(rstImporter07)
    Call CloseRecordset(rstImporter08)
    Call CloseRecordset(rstCompanyMaster)
    Call CloseConnection(CxnImporter)
    BusySystemIndicator False
    Exit Sub
ErrorHandler:
    If CxnImporter.State = adStateOpen Then CxnDatabase.RollbackTrans
    BusySystemIndicator False
    DisplayError ("Failed to import the Balances")
    Call CloseRecordset(rstImporter00)
    Call CloseRecordset(rstImporter05)
    Call CloseRecordset(rstImporter06)
    Call CloseRecordset(rstImporter07)
    Call CloseRecordset(rstImporter08)
    Call CloseRecordset(rstCompanyMaster)
    Call CloseConnection(CxnImporter)
End Sub
Private Sub MnuImportBal02_Click()
    Dim CxnImporter As New ADODB.Connection
    Dim rstCompanyMaster As New ADODB.Recordset
    Dim rstImporter As New ADODB.Recordset
    Dim ClBal As Double
    On Error GoTo ErrorHandler
    BusySystemIndicator True
    rstCompanyMaster.Open "Select CreatedFrom From CompanyMaster", CxnDatabase, adOpenKeyset, adLockReadOnly
    If rstCompanyMaster.Fields("CreatedFrom").Value <> "" Then
        If MsgBox("Are you sure to Proceed?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Proceed !") = vbYes Then
            CxnImporter.CursorLocation = adUseClient
            CxnImporter.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DatabasePath & "\Saral." & rstCompanyMaster.Fields("CreatedFrom").Value & ";Persist Security Info=False;Jet OLEDB:Database Password=RSPLILoveMyINDIA"
            Dim Tbl As String
            Tbl = "SELECT Code As Paper,Account FROM PaperChild WHERE Code<>'' AND Account<>'' UNION SELECT Paper,Account FROM PaperIOChild WHERE Paper<>'' AND Account<>'' UNION SELECT Item As Paper,Account FROM MaterialSVParent P INNER JOIN MaterialSVChild C ON P.Code=C.Code WHERE Category='2' AND Item<>'' AND Account<>'' UNION SELECT Paper,AccountFrom As Account FROM PaperMVParent P INNER JOIN PaperMVChild C ON P.Code=C.Code WHERE Paper<>'' AND AccountFrom<>'' UNION SELECT Paper,AccountTo As Account FROM PaperMVParent P INNER JOIN PaperMVChild C ON P.Code=C.Code WHERE Paper<>'' AND AccountTo<>'' UNION " & _
             "SELECT Item As Paper,Binder As Account FROM (BookPOParent P INNER JOIN BookPOChild08 C1 ON P.Code=C1.Code) INNER JOIN BookPOChild0801 C2 ON C1.Code=C2.Code WHERE C2.Category='2' AND P.Type<>'O' AND LEFT(P.Code,1)<>'*' AND Item<>'' AND Binder<>'' UNION SELECT Paper,TitlePrinter As Account FROM BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code WHERE P.Type<>'O' AND LEFT(P.Code,1)<>'*' AND Paper<>'' AND TitlePrinter<>'' UNION SELECT Paper1 As Paper,BookPrinter As Account FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE P.Type<>'O' AND LEFT(P.Code,1)<>'*' AND Paper1<>'' AND BookPrinter<>'' UNION SELECT Paper2 As Paper,BookPrinter As Account FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE P.Type<>'O' AND LEFT(P.Code,1)<>'*' AND Paper2<>'' AND BookPrinter<>'' UNION SELECT Paper4 As Paper,BookPrinter As Account FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE P.Type<>'O' AND LEFT(P.Code,1)<>'*' AND Paper4<>'' AND BookPrinter<>''"
            rstImporter.Open "SELECT Account,Paper," & _
                                         "FORMAT((SELECT SUM(OpBalSheets) FROM PaperChild Where Code=T.Paper And Account=T.Account),0) As OpBal," & _
                                         "FORMAT((SELECT SUM(QuantitySheets) FROM PaperPOParent P INNER JOIN PaperIOChild C ON P.Code=C.Code WHERE C.Account=T.Account AND C.Paper=T.Paper),0) As IN1," & _
                                         "FORMAT((SELECT SUM(INT(Quantity)*500+(Quantity-INT(Quantity))*1000) FROM MaterialSVParent P INNER JOIN MaterialSVChild C ON P.Code=C.Code WHERE P.Account=T.Account AND C.Category='2' AND C.Item=T.Paper AND C.Quantity>=0),0) As IN2," & _
                                         "FORMAT((SELECT SUM(ABS(Fix(Quantity)*500+(Quantity-Fix(Quantity))*1000)) FROM MaterialSVParent P INNER JOIN MaterialSVChild C ON P.Code=C.Code WHERE P.Account=T.Account AND C.Category='2' AND C.Item=T.Paper AND C.Quantity<0),0) As OUT1," & _
                                         "FORMAT((SELECT SUM(QuantitySheets) FROM PaperMVParent P INNER JOIN PaperMVChild C ON P.Code=C.Code WHERE P.AccountFrom=T.Account AND C.Paper=T.Paper),0) As OUT2," & _
                                         "FORMAT((SELECT SUM(QuantitySheets) FROM PaperMVParent P INNER JOIN PaperMVChild C ON P.Code=C.Code WHERE P.AccountTo=T.Account AND C.Paper=T.Paper),0) As IN3," & _
                                         "FORMAT((SELECT SUM(ROUND(ActualQuantity*C1.Quantity,0)) FROM (BookPOParent P INNER JOIN BookPOChild08 C ON P.Code=C.Code) INNER JOIN BookPOChild0801 C1 ON C.Code=C1.Code WHERE P.Binder=T.Account AND C1.Category='2' AND C1.Item=T.Paper AND P.Type<>'O' AND LEFT(P.Code,1)<>'*'),0) As OUT3," & _
                                         "FORMAT((SELECT SUM(PaperConsumptionSheets) FROM BookPOParent P INNER JOIN BookPOChild06 C ON P.Code=C.Code WHERE P.TitlePrinter=T.Account AND C.Paper=T.Paper AND P.Type<>'O' AND LEFT(P.Code,1)<>'*'),0) As OUT4," & _
                                         "FORMAT((SELECT SUM(PaperConsumptionSheets1) FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE P.BookPrinter=T.Account AND C.Paper1=T.Paper AND P.Type<>'O' AND LEFT(P.Code,1)<>'*'),0) As OUT5," & _
                                         "FORMAT((SELECT SUM(PaperConsumptionSheets2) FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE P.BookPrinter=T.Account AND C.Paper2=T.Paper AND P.Type<>'O' AND LEFT(P.Code,1)<>'*'),0) As OUT6," & _
                                         "FORMAT((SELECT SUM(PaperConsumptionSheets4) FROM BookPOParent P INNER JOIN BookPOChild05 C ON P.Code=C.Code WHERE P.BookPrinter=T.Account AND C.Paper4=T.Paper AND P.Type<>'O' AND LEFT(P.Code,1)<>'*'),0) As OUT7, " & _
                                         "FORMAT((SELECT INT(Quantity)*500+(Quantity-INT(Quantity))*1000 FROM PaperDNParent P INNER JOIN PaperDNChild C ON P.Code=C.Code WHERE P.Account=T.Account AND C.Paper=T.Paper),0) As OUT8 " & _
                                         "FROM (" & Tbl & ") As T ORDER BY Account,Paper", CxnImporter, adOpenKeyset, adLockReadOnly
            
            rstImporter.ActiveConnection = Nothing
            
            rstCompanyMaster.ActiveConnection = Nothing
            CxnDatabase.BeginTrans
            CxnDatabase.Execute "Delete From PaperChild Where Imported = 'Y'"
            Do While Not rstImporter.EOF
                ClBal = Val(CheckNull(rstImporter.Fields("OpBal").Value)) + Val(CheckNull(rstImporter.Fields("IN1").Value)) + Val(CheckNull(rstImporter.Fields("IN2").Value)) + Val(CheckNull(rstImporter.Fields("IN3").Value)) - Val(CheckNull(rstImporter.Fields("OUT1").Value)) - Val(CheckNull(rstImporter.Fields("OUT2").Value)) - Val(CheckNull(rstImporter.Fields("OUT3").Value)) - Val(CheckNull(rstImporter.Fields("OUT4").Value)) - Val(CheckNull(rstImporter.Fields("OUT5").Value)) - Val(CheckNull(rstImporter.Fields("OUT6").Value)) - Val(CheckNull(rstImporter.Fields("OUT7").Value)) - Val(CheckNull(rstImporter.Fields("OUT8").Value))
                If ClBal <> 0 Then CxnDatabase.Execute "Insert Into PaperChild Values ('" & rstImporter.Fields("Paper").Value & "','" & rstImporter.Fields("Account").Value & "'," & CLng(Fix(ClBal / 500)) + ((ClBal Mod 500) / 1000) & "," & ClBal & ",0,'Y')"
                rstImporter.MoveNext
            Loop
            CxnDatabase.CommitTrans
            Call MsgBox("Successfully imported the Balances !", vbInformation, App.Title)
        End If
    Else
        Call MsgBox("Nothing To Import !", vbInformation, App.Title)
    End If
    Call CloseRecordset(rstImporter)
    Call CloseRecordset(rstCompanyMaster)
    Call CloseConnection(CxnImporter)
    BusySystemIndicator False
    Exit Sub
ErrorHandler:
    If CxnImporter.State = adStateOpen Then CxnDatabase.RollbackTrans
    BusySystemIndicator False
    DisplayError ("Failed to import the Balances")
    Call CloseRecordset(rstImporter)
    Call CloseRecordset(rstCompanyMaster)
    Call CloseConnection(CxnImporter)
End Sub
Private Sub MnuImportBal03_Click()
    
    Dim CxnImporter As New ADODB.Connection
    Dim rstCompanyMaster As New ADODB.Recordset
    Dim rstImporter As New ADODB.Recordset
    Dim ClBal As Double
    On Error GoTo ErrorHandler
    
    BusySystemIndicator True
    rstCompanyMaster.Open "Select CreatedFrom From CompanyMaster", CxnDatabase, adOpenKeyset, adLockReadOnly
    If rstCompanyMaster.Fields("CreatedFrom").Value <> "" Then
        If MsgBox("Are you sure to Proceed?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Proceed !") = vbYes Then
            CxnImporter.CursorLocation = adUseClient
            CxnImporter.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DatabasePath & "\Saral." & rstCompanyMaster.Fields("CreatedFrom").Value & ";Persist Security Info=False;Jet OLEDB:Database Password=RSPLILoveMyINDIA"
            'Outsource Items
            rstImporter.Open "SELECT DISTINCT A.Code,O.Code," & _
                                         "FORMAT((SELECT SUM(OpBal) FROM AccountChild0801 Where Category+Item='1'+O.Code AND Code=A.Code),0) As OpBal," & _
                                         "FORMAT((SELECT SUM(Quantity) FROM MaterialIOParent M INNER JOIN MaterialIOChild I ON M.Code=I.Code WHERE I.Category+I.Item='1'+O.Code AND I.Godown=A.Code),0) As IN1," & _
                                         "FORMAT((SELECT SUM(Quantity) FROM MaterialSVParent M INNER JOIN MaterialSVChild I ON M.Code=I.Code WHERE I.Category+I.Item='1'+O.Code AND M.Account=A.Code AND I.Quantity>=0),0) As IN2," & _
                                         "FORMAT((SELECT SUM(ABS(Quantity)) FROM MaterialSVParent M INNER JOIN MaterialSVChild I ON M.Code=I.Code WHERE I.Category+I.Item='1'+O.Code AND M.Account=A.Code AND I.Quantity<0),0) As OUT1," & _
                                         "FORMAT((SELECT SUM(Quantity) FROM MaterialMVParent M INNER JOIN MaterialMVChild I ON M.Code=I.Code WHERE I.Category+I.Item='1'+O.Code AND M.AccountFrom=A.Code),0) As OUT2," & _
                                         "FORMAT((SELECT SUM(Quantity) FROM MaterialMVParent M INNER JOIN MaterialMVChild I ON M.Code=I.Code WHERE I.Category+I.Item='1'+O.Code AND M.AccountTo=A.Code),0) As IN3," & _
                                         "FORMAT((SELECT SUM(Quantity*(SELECT ActualQuantity FROM BookPOChild08 WHERE Code=M.Code)) FROM BookPOParent M INNER JOIN BookPOChild0801 I ON M.Code=I.Code WHERE M.Type<>'O' AND LEFT(M.Code,1)<>'*' AND I.Category+I.Item='1'+O.Code AND M.Binder=A.Code),0) As OUT3 " & _
                                         "FROM OutsourceItemMaster O,AccountMaster A WHERE A.Type In ('08','09') ORDER BY A.Code,O.Code", CxnImporter, adOpenKeyset, adLockReadOnly
            rstImporter.ActiveConnection = Nothing
            rstCompanyMaster.ActiveConnection = Nothing
            CxnDatabase.BeginTrans
            CxnDatabase.Execute "Delete From AccountChild0801 Where Imported = 'Y'"
            Do While Not rstImporter.EOF
                ClBal = Val(CheckNull(rstImporter.Fields("OpBal").Value)) + Val(CheckNull(rstImporter.Fields("IN1").Value)) + Val(CheckNull(rstImporter.Fields("IN2").Value)) + Val(CheckNull(rstImporter.Fields("IN3").Value)) - Val(CheckNull(rstImporter.Fields("OUT1").Value)) - Val(CheckNull(rstImporter.Fields("OUT2").Value)) - Val(CheckNull(rstImporter.Fields("OUT3").Value))
                If ClBal <> 0 Then
                    CxnDatabase.Execute "Insert Into AccountChild0801 Values ('" & rstImporter.Fields("A.Code").Value & "','1','" & rstImporter.Fields("O.Code").Value & "'," & ClBal & ",'Y')"
                End If
                rstImporter.MoveNext
            Loop
            If rstImporter.State = adStateOpen Then rstImporter.Close
            'Fresh Books
            
            
            
            rstImporter.Open "SELECT A.Code,O.Code," & _
                                         "FORMAT((SELECT SUM(OpBal) FROM AccountChild0801 Where Category+Item='3'+O.Code AND Code=A.Code),0) As OpBal," & _
                                         "FORMAT((SELECT SUM(Quantity) FROM MaterialIOParent M INNER JOIN MaterialIOChild I ON M.Code=I.Code WHERE I.Category+I.Item='3'+O.Code AND I.Godown=A.Code),0) As IN1," & _
                                         "FORMAT((SELECT SUM(Quantity) FROM MaterialSVParent M INNER JOIN MaterialSVChild I ON M.Code=I.Code WHERE I.Category+I.Item='3'+O.Code AND M.Account=A.Code AND I.Quantity>=0),0) As IN2," & _
                                         "FORMAT((SELECT SUM(ABS(Quantity)) FROM MaterialSVParent M INNER JOIN MaterialSVChild I ON M.Code=I.Code WHERE I.Category+I.Item='3'+O.Code AND M.Account=A.Code AND I.Quantity<0),0) As OUT1," & _
                                         "FORMAT((SELECT SUM(Quantity) FROM MaterialMVParent M INNER JOIN MaterialMVChild I ON M.Code=I.Code WHERE I.Category+I.Item='3'+O.Code AND M.AccountFrom=A.Code),0) As OUT2," & _
                                         "FORMAT((SELECT SUM(Quantity) FROM MaterialMVParent M INNER JOIN MaterialMVChild I ON M.Code=I.Code WHERE I.Category+I.Item='3'+O.Code AND M.AccountTo=A.Code),0) As IN3," & _
                                         "FORMAT((SELECT SUM(Quantity*(SELECT ActualQuantity FROM BookPOChild08 WHERE Code=M.Code)) FROM BookPOParent M INNER JOIN BookPOChild0801 I ON M.Code=I.Code WHERE M.Type<>'O' AND LEFT(M.Code,1)<>'*' AND I.Category+I.Item='3'+O.Code AND M.Binder=A.Code),0) As OUT3 " & _
                                         "FROM BookMaster O,AccountMaster A WHERE O.Board='000000' AND A.Type In ('08','09') ORDER BY A.Code,O.Code", CxnImporter, adOpenKeyset, adLockReadOnly
            rstImporter.ActiveConnection = Nothing
            
            Do While Not rstImporter.EOF
                ClBal = Val(CheckNull(rstImporter.Fields("OpBal").Value)) + Val(CheckNull(rstImporter.Fields("IN1").Value)) + Val(CheckNull(rstImporter.Fields("IN2").Value)) + Val(CheckNull(rstImporter.Fields("IN3").Value)) - Val(CheckNull(rstImporter.Fields("OUT1").Value)) - Val(CheckNull(rstImporter.Fields("OUT2").Value)) - Val(CheckNull(rstImporter.Fields("OUT3").Value))
                If ClBal <> 0 Then
                    CxnDatabase.Execute "Insert Into AccountChild0801 Values ('" & rstImporter.Fields("A.Code").Value & "','3','" & rstImporter.Fields("O.Code").Value & "'," & ClBal & ",'Y')"
                End If
                rstImporter.MoveNext
            Loop
            If rstImporter.State = adStateOpen Then rstImporter.Close
            'Repair Books
            rstImporter.Open "SELECT A.Code,O.Code," & _
                                         "FORMAT((SELECT SUM(OpBal) FROM AccountChild0801 Where Category+Item='4'+O.Code AND Code=A.Code),0) As OpBal," & _
                                         "FORMAT((SELECT SUM(Quantity) FROM MaterialIOParent M INNER JOIN MaterialIOChild I ON M.Code=I.Code WHERE I.Category+I.Item='4'+O.Code AND I.Godown=A.Code),0) As IN1," & _
                                         "FORMAT((SELECT SUM(Quantity) FROM MaterialSVParent M INNER JOIN MaterialSVChild I ON M.Code=I.Code WHERE I.Category+I.Item='4'+O.Code AND M.Account=A.Code AND I.Quantity>=0),0) As IN2," & _
                                         "FORMAT((SELECT SUM(ABS(Quantity)) FROM MaterialSVParent M INNER JOIN MaterialSVChild I ON M.Code=I.Code WHERE I.Category+I.Item='4'+O.Code AND M.Account=A.Code AND I.Quantity<0),0) As OUT1," & _
                                         "FORMAT((SELECT SUM(Quantity) FROM MaterialMVParent M INNER JOIN MaterialMVChild I ON M.Code=I.Code WHERE I.Category+I.Item='4'+O.Code AND M.AccountFrom=A.Code),0) As OUT2," & _
                                         "FORMAT((SELECT SUM(Quantity) FROM MaterialMVParent M INNER JOIN MaterialMVChild I ON M.Code=I.Code WHERE I.Category+I.Item='4'+O.Code AND M.AccountTo=A.Code),0) As IN3," & _
                                         "FORMAT((SELECT SUM(Quantity*(SELECT ActualQuantity FROM BookPOChild08 WHERE Code=M.Code)) FROM BookPOParent M INNER JOIN BookPOChild0801 I ON M.Code=I.Code WHERE M.Type<>'O' AND LEFT(M.Code,1)<>'*' AND I.Category+I.Item='4'+O.Code AND M.Binder=A.Code),0) As OUT3 " & _
                                         "FROM BookMaster O,AccountMaster A WHERE O.Type='R' AND A.Type In ('08','09') ORDER BY A.Code,O.Code", CxnImporter, adOpenKeyset, adLockReadOnly
            rstImporter.ActiveConnection = Nothing
            Do While Not rstImporter.EOF
                ClBal = Val(CheckNull(rstImporter.Fields("OpBal").Value)) + Val(CheckNull(rstImporter.Fields("IN1").Value)) + Val(CheckNull(rstImporter.Fields("IN2").Value)) + Val(CheckNull(rstImporter.Fields("IN3").Value)) - Val(CheckNull(rstImporter.Fields("OUT1").Value)) - Val(CheckNull(rstImporter.Fields("OUT2").Value)) - Val(CheckNull(rstImporter.Fields("OUT3").Value))
                If ClBal <> 0 Then
                    CxnDatabase.Execute "Insert Into AccountChild0801 Values ('" & rstImporter.Fields("A.Code").Value & "','4','" & rstImporter.Fields("O.Code").Value & "'," & ClBal & ",'Y')"
                End If
                rstImporter.MoveNext
            Loop
            If rstImporter.State = adStateOpen Then rstImporter.Close
            'Title
            rstImporter.Open "SELECT A.Code,O.Code," & _
                                         "FORMAT((SELECT SUM(OpBal) FROM AccountChild0801 Where Category+Item='5'+O.Code AND Code=A.Code),0) As OpBal," & _
                                         "FORMAT((SELECT SUM(Quantity) FROM MaterialIOParent M INNER JOIN MaterialIOChild I ON M.Code=I.Code WHERE I.Category+I.Item='5'+O.Code AND I.Godown=A.Code),0) As IN1," & _
                                         "FORMAT((SELECT SUM(Quantity) FROM MaterialSVParent M INNER JOIN MaterialSVChild I ON M.Code=I.Code WHERE I.Category+I.Item='5'+O.Code AND M.Account=A.Code AND I.Quantity>=0),0) As IN2," & _
                                         "FORMAT((SELECT SUM(ABS(Quantity)) FROM MaterialSVParent M INNER JOIN MaterialSVChild I ON M.Code=I.Code WHERE I.Category+I.Item='5'+O.Code AND M.Account=A.Code AND I.Quantity<0),0) As OUT1," & _
                                         "FORMAT((SELECT SUM(Quantity) FROM MaterialMVParent M INNER JOIN MaterialMVChild I ON M.Code=I.Code WHERE I.Category+I.Item='5'+O.Code AND M.AccountFrom=A.Code),0) As OUT2," & _
                                         "FORMAT((SELECT SUM(Quantity) FROM MaterialMVParent M INNER JOIN MaterialMVChild I ON M.Code=I.Code WHERE I.Category+I.Item='5'+O.Code AND M.AccountTo=A.Code),0) As IN3," & _
                                         "FORMAT((SELECT SUM(Quantity*(SELECT ActualQuantity FROM BookPOChild08 WHERE Code=M.Code)) FROM BookPOParent M INNER JOIN BookPOChild0801 I ON M.Code=I.Code WHERE M.Type<>'O' AND LEFT(M.Code,1)<>'*' AND I.Category+I.Item='5'+O.Code AND M.Binder=A.Code),0) As OUT3 " & _
                                         "FROM BookMaster O,AccountMaster A WHERE O.Board<>'000000' AND O.Type='F' AND A.Type In ('08','09') ORDER BY A.Code,O.Code", CxnImporter, adOpenKeyset, adLockReadOnly
            rstImporter.ActiveConnection = Nothing
            Do While Not rstImporter.EOF
                ClBal = Val(CheckNull(rstImporter.Fields("OpBal").Value)) + Val(CheckNull(rstImporter.Fields("IN1").Value)) + Val(CheckNull(rstImporter.Fields("IN2").Value)) + Val(CheckNull(rstImporter.Fields("IN3").Value)) - Val(CheckNull(rstImporter.Fields("OUT1").Value)) - Val(CheckNull(rstImporter.Fields("OUT2").Value)) - Val(CheckNull(rstImporter.Fields("OUT3").Value))
                If ClBal <> 0 Then
                    CxnDatabase.Execute "Insert Into AccountChild0801 Values ('" & rstImporter.Fields("A.Code").Value & "','5','" & rstImporter.Fields("O.Code").Value & "'," & ClBal & ",'Y')"
                End If
                rstImporter.MoveNext
            Loop
            CxnDatabase.CommitTrans
            Call MsgBox("Successfully imported the Balances !", vbInformation, App.Title)
        End If
    Else
        Call MsgBox("Nothing To Import !", vbInformation, App.Title)
    End If
    
    Call CloseRecordset(rstImporter)
    Call CloseRecordset(rstCompanyMaster)
    Call CloseConnection(CxnImporter)
    BusySystemIndicator False
    
    Exit Sub
    
ErrorHandler:

    If CxnImporter.State = adStateOpen Then CxnDatabase.RollbackTrans
    BusySystemIndicator False
    DisplayError ("Failed to import the Balances")
    
    Call CloseRecordset(rstImporter)
    Call CloseRecordset(rstCompanyMaster)
    Call CloseConnection(CxnImporter)
   
End Sub
Private Sub CloseMainConnection()
    If CxnDatabase.State = adStateOpen Then
        CxnDatabase.Close
    End If
End Sub
Private Function IsFormLoaded(ByVal FormCaption As String) As Boolean
    Dim Form As Form
    IsFormLoaded = False
    For Each Form In Forms
        If Form.Caption = FormCaption Then
            IsFormLoaded = True
            Exit For
        End If
    Next Form
End Function
