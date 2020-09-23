VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm Main_Menu 
   BackColor       =   &H8000000C&
   Caption         =   "TBS' Business Accounting v1.0"
   ClientHeight    =   5535
   ClientLeft      =   2490
   ClientTop       =   3570
   ClientWidth     =   11055
   Icon            =   "Main_Menu1.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar tbToolBar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   1290
      Visible         =   0   'False
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlIcons1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   14
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "New"
            Object.ToolTipText     =   "Create New Working Database"
            Object.Tag             =   ""
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open Working Database"
            Object.Tag             =   ""
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save Working Database"
            Object.Tag             =   ""
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Print"
            Object.ToolTipText     =   "Print"
            Object.Tag             =   ""
            ImageKey        =   "Print"
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Calendar"
            Object.ToolTipText     =   "Show Calendar"
            Object.Tag             =   ""
            ImageIndex      =   19
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Contacts"
            Object.ToolTipText     =   "Show Contact"
            Object.Tag             =   ""
            ImageKey        =   "contacts"
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "InventoryList"
            Object.ToolTipText     =   "Show Inventory List"
            Object.Tag             =   ""
            ImageKey        =   "Books"
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Help"
            Object.ToolTipText     =   "Off-Line Help"
            Object.Tag             =   ""
            ImageKey        =   "Help"
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "stop"
            Object.ToolTipText     =   "Stop Connection"
            Object.Tag             =   ""
            ImageKey        =   "stop"
         EndProperty
      EndProperty
   End
   Begin ComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   0
      Top             =   420
      Visible         =   0   'False
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   1535
      ButtonWidth     =   1349
      ButtonHeight    =   1376
      Appearance      =   1
      ImageList       =   "imlIcons"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   17
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "New"
            Key             =   "New"
            Object.ToolTipText     =   "Create New Working Database"
            Object.Tag             =   ""
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Open"
            Key             =   "Open"
            Object.ToolTipText     =   "Open Working Database"
            Object.Tag             =   ""
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Save"
            Key             =   "Save"
            Object.ToolTipText     =   "Save Working Database As"
            Object.Tag             =   ""
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Preview"
            Key             =   "Preview"
            Object.ToolTipText     =   "Print Preview"
            Object.Tag             =   "Preview"
            ImageKey        =   "Preview"
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Print"
            Key             =   "Print"
            Object.ToolTipText     =   "Print"
            Object.Tag             =   ""
            ImageKey        =   "Print"
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Cut"
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut"
            Object.Tag             =   ""
            ImageKey        =   "Cut"
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Copy"
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            Object.Tag             =   ""
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Caption         =   "Paste"
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            Object.Tag             =   ""
            ImageKey        =   "Paste"
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Calendar"
            Key             =   "Calendar"
            Object.ToolTipText     =   "Show Calendar & Appoinment"
            Object.Tag             =   ""
            ImageKey        =   "Calendar"
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Contacts"
            Key             =   "Contacts"
            Object.ToolTipText     =   "Show Contact"
            Object.Tag             =   ""
            ImageKey        =   "Contacts"
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Inventory"
            Key             =   "InventoryList"
            Object.ToolTipText     =   "Show Inventory List"
            Object.Tag             =   ""
            ImageKey        =   "Inventory"
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Help"
            Key             =   "Help"
            Object.ToolTipText     =   "Off-Line Help"
            Object.Tag             =   ""
            ImageKey        =   "Help"
         EndProperty
         BeginProperty Button17 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Close"
            Key             =   "stop"
            Object.ToolTipText     =   "Stop Connection"
            Object.Tag             =   ""
            ImageKey        =   "Stop"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   741
      BandCount       =   1
      Picture         =   "Main_Menu1.frx":030A
      EmbossPicture   =   -1  'True
      _CBWidth        =   11055
      _CBHeight       =   420
      _Version        =   "6.7.8862"
      MinHeight1      =   360
      Width1          =   5505
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   780
         Left            =   10880
         MouseIcon       =   "Main_Menu1.frx":13FD8
         MousePointer    =   99  'Custom
         Picture         =   "Main_Menu1.frx":142E2
         ScaleHeight     =   780
         ScaleWidth      =   4320
         TabIndex        =   8
         Top             =   25
         Width           =   4320
      End
      Begin TBS.TransTBWrapper TransTBWrapper1 
         Height          =   285
         Left            =   210
         TabIndex        =   7
         Top             =   60
         Visible         =   0   'False
         Width           =   1695
         _extentx        =   2990
         _extenty        =   503
         transparent     =   -1  'True
      End
   End
   Begin VB.PictureBox picStatusBar 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   11055
      TabIndex        =   2
      Top             =   5235
      Width           =   11055
      Begin VB.PictureBox picIcon 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   14760
         Picture         =   "Main_Menu1.frx":15DD7
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   5
         Top             =   30
         Width           =   480
      End
      Begin VB.TextBox txtLogon 
         BackColor       =   &H80000000&
         Height          =   285
         Left            =   2910
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "Logon"
         Top             =   10
         Width           =   2175
      End
      Begin VB.TextBox txtConnection 
         BackColor       =   &H80000000&
         Height          =   285
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "Connection"
         Top             =   10
         Width           =   2895
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   600
      Top             =   1920
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   120
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComctlLib.ImageList imgIcon 
      Left            =   2280
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main_Menu1.frx":160E1
            Key             =   "off"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main_Menu1.frx":163FB
            Key             =   "on"
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList imlIcons1 
      Left            =   1680
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483644
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   19
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main_Menu1.frx":16715
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main_Menu1.frx":16827
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main_Menu1.frx":16939
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main_Menu1.frx":16A4B
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main_Menu1.frx":16B5D
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main_Menu1.frx":16C6F
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main_Menu1.frx":16D81
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main_Menu1.frx":16E93
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main_Menu1.frx":16FA5
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main_Menu1.frx":170B7
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main_Menu1.frx":171C9
            Key             =   "Align Left"
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main_Menu1.frx":172DB
            Key             =   "Align Right"
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main_Menu1.frx":173ED
            Key             =   "Center"
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main_Menu1.frx":174FF
            Key             =   "stop"
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main_Menu1.frx":17819
            Key             =   "Calculator"
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main_Menu1.frx":17B33
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main_Menu1.frx":17E4D
            Key             =   "contacts"
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main_Menu1.frx":18167
            Key             =   "Books"
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main_Menu1.frx":18481
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList imlIcons 
      Left            =   1080
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483644
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   16
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main_Menu1.frx":1879B
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main_Menu1.frx":18AB5
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main_Menu1.frx":18DCF
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main_Menu1.frx":190E9
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main_Menu1.frx":19403
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main_Menu1.frx":1971D
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main_Menu1.frx":19A37
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main_Menu1.frx":19D51
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main_Menu1.frx":1A06B
            Key             =   "Inventory"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main_Menu1.frx":1A385
            Key             =   "Calculator"
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main_Menu1.frx":1A69F
            Key             =   "Contacts"
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main_Menu1.frx":1A9B9
            Key             =   "Calendar"
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main_Menu1.frx":1ACD3
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main_Menu1.frx":1AFED
            Key             =   "Locked"
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main_Menu1.frx":1B307
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Main_Menu1.frx":1B621
            Key             =   "Approved"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "New Company"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open Company"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "Close Company"
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save Company As"
         Begin VB.Menu mnuNewDb 
            Caption         =   "New DataBase"
         End
         Begin VB.Menu mnuBackup 
            Caption         =   "Backup"
         End
         Begin VB.Menu mnuBackupServer 
            Caption         =   "Backup on Secured Server"
         End
      End
      Begin VB.Menu mnusenkang 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTips1 
         Caption         =   "Bulletin Board"
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "Properties"
      End
      Begin VB.Menu mnuSenkang2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHistory 
         Caption         =   "History"
      End
      Begin VB.Menu mnuErrorLog 
         Caption         =   "Error Log"
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrintPreview 
         Caption         =   "Print Preview"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "Print"
      End
      Begin VB.Menu mnuFileBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "Paste"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mnuSetup 
      Caption         =   "&Setup"
      Begin VB.Menu mnuDataSYS_Setup_Checklist 
         Caption         =   "Checklist"
      End
      Begin VB.Menu mnuSetupBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDataChartOfAccounts 
         Caption         =   "Chart Of Accounts"
      End
      Begin VB.Menu mnuSetupBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDataSYS_Setup_Company 
         Caption         =   "Company"
      End
      Begin VB.Menu mnuCustomers1 
         Caption         =   "Customers"
         Begin VB.Menu mnuDataAR_Customer 
            Caption         =   "Customers"
         End
         Begin VB.Menu mnuDataLIST_Customer_Types 
            Caption         =   "Customer Types"
         End
      End
      Begin VB.Menu mnuPayment 
         Caption         =   "Payment"
         Begin VB.Menu mnuDataLIST_Payment_Methods 
            Caption         =   "Payment Methods"
         End
         Begin VB.Menu mnuDataLIST_Payment_Terms 
            Caption         =   "Payment Terms"
         End
         Begin VB.Menu mnuDatafrmLISTCreditCards 
            Caption         =   "Credit Cards"
         End
      End
      Begin VB.Menu mnuTax 
         Caption         =   "Tax"
         Begin VB.Menu mnuTaxAuthorities 
            Caption         =   "Tax Authorities"
         End
         Begin VB.Menu mnuTaxGroup 
            Caption         =   "Tax Group"
         End
      End
      Begin VB.Menu mnuVendors 
         Caption         =   "Vendors"
         Begin VB.Menu mnuDataAP_Vendor 
            Caption         =   "Vendors"
         End
         Begin VB.Menu mnuDataLIST_Vendor_Types 
            Caption         =   "Vendor Types"
         End
      End
      Begin VB.Menu mnuInventory1 
         Caption         =   "Inventory"
         Begin VB.Menu mnuInventoryItems 
            Caption         =   "Items"
         End
         Begin VB.Menu mnuDataSYS_Setup_Inventory 
            Caption         =   "Inventory"
         End
         Begin VB.Menu mnuDataLIST_Item_Catagories 
            Caption         =   "Item Catagories"
         End
      End
      Begin VB.Menu mnuProject1 
         Caption         =   "Projects"
         Begin VB.Menu mnuProjects 
            Caption         =   "Projects"
         End
         Begin VB.Menu mnuDataLIST_Project_Types 
            Caption         =   "Project Type"
         End
      End
      Begin VB.Menu mnuAccounting1 
         Caption         =   "Accounting"
         Begin VB.Menu mnuAcctPreferences 
            Caption         =   "Accounting Preferences"
         End
         Begin VB.Menu mnuDataSYS_Setup_Banks 
            Caption         =   "Banks"
         End
         Begin VB.Menu mnuDataSYS_Setup_Periods 
            Caption         =   "Periods"
         End
      End
      Begin VB.Menu mnuRecurring 
         Caption         =   "Recurring Type"
      End
      Begin VB.Menu mnuDataLIST_Shipping_Methods 
         Caption         =   "Shipping Methods"
      End
      Begin VB.Menu mnuDataSYS_Setup_Purchases 
         Caption         =   "Purchasing Preferences"
      End
      Begin VB.Menu mnuDataSYS_Setup_Sales 
         Caption         =   "Sales Preferences"
      End
   End
   Begin VB.Menu mnuSales 
      Caption         =   "Sa&les"
      Begin VB.Menu mnuSalesBatchPosting 
         Caption         =   "Active AR Transaction"
      End
      Begin VB.Menu sengkangAR 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalesQuoteEntry 
         Caption         =   "Quote Entry"
      End
      Begin VB.Menu mnuSalesOrderEntry 
         Caption         =   "Order Entry"
      End
      Begin VB.Menu mnuSalesSalesEntry 
         Caption         =   "Invoice"
      End
      Begin VB.Menu mnuSalesReturnEntry 
         Caption         =   "Return"
      End
      Begin VB.Menu mnuSalesCreditMemo 
         Caption         =   "Credit Memo"
      End
      Begin VB.Menu mnuSalesMemo 
         Caption         =   "Sales Memo"
      End
      Begin VB.Menu sengkangAR1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalesCashReceipts 
         Caption         =   "Cash Receipts"
      End
      Begin VB.Menu mnuRecurrSales 
         Caption         =   "Recurring Sales"
      End
   End
   Begin VB.Menu mnuPurchasing 
      Caption         =   "&Purchasing"
      Begin VB.Menu mnuPurchasingBatchPosting 
         Caption         =   "Active AP Transaction"
      End
      Begin VB.Menu sengkangAP 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPurchasingPurchases 
         Caption         =   "Purchase Order"
      End
      Begin VB.Menu mnuPurchasingReceiving 
         Caption         =   "Receiving"
      End
      Begin VB.Menu mnuPurchasingVoucher 
         Caption         =   "Voucher"
      End
      Begin VB.Menu mnuPurchasingCreditMemo 
         Caption         =   "Credit Memo"
      End
      Begin VB.Menu mnuRMA 
         Caption         =   "RMA"
      End
      Begin VB.Menu sengkangAP1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPayments 
         Caption         =   "Payments"
         Begin VB.Menu mnuPurchasingCashPayments 
            Caption         =   "Cash Payments"
         End
         Begin VB.Menu mnuPurchasingPayManyVendors 
            Caption         =   "Pay Many Vendors"
         End
      End
      Begin VB.Menu mnuRecurrPurchases 
         Caption         =   "Recurring Purchases"
      End
      Begin VB.Menu mnuRecurrPayments 
         Caption         =   "Recurring Payments"
      End
   End
   Begin VB.Menu mnuInventory 
      Caption         =   "&Inventory"
      Begin VB.Menu mnuInventoryInventoryAdjustment 
         Caption         =   "Inventory Adjustment"
      End
      Begin VB.Menu mnuInventoryInventoryProduction 
         Caption         =   "Inventory Production"
      End
   End
   Begin VB.Menu mnuAccounting 
      Caption         =   "&Accounting"
      Begin VB.Menu mnuAccountingChartOfAccounts 
         Caption         =   "Chart of Accounts"
      End
      Begin VB.Menu sengkangBank 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAccountingGLEntry 
         Caption         =   "GL Entry"
      End
      Begin VB.Menu mnuRecurrGL 
         Caption         =   "Recurring GL"
      End
      Begin VB.Menu mnuSengkangGL 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAccountingBankTransactions 
         Caption         =   "Bank Transactions"
      End
      Begin VB.Menu mnuAccountingBankReconsiliation 
         Caption         =   "Bank Reconciliation"
      End
      Begin VB.Menu mnuSengkangAcct 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCheck 
         Caption         =   "Check Management"
      End
      Begin VB.Menu mnuFinaceCharges 
         Caption         =   "Asess Finance Charges"
      End
      Begin VB.Menu mnuCloseMonth 
         Caption         =   "Closing Month Procedure"
      End
      Begin VB.Menu mnuYearEnd 
         Caption         =   "Year End Procedures"
         Begin VB.Menu mnuAgeARnAP 
            Caption         =   "Age Receivables and Payables"
         End
         Begin VB.Menu mnuGLprocess 
            Caption         =   "General Ledger Processing"
         End
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuAdvertise 
         Caption         =   "Advertising"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "Status Bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "Tool Bar"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuCloseActive 
         Caption         =   "Close Active Form"
      End
      Begin VB.Menu mnuCloseAll 
         Caption         =   "Close All"
      End
      Begin VB.Menu mnuWindowBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "Cascade"
      End
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "Tile Horizontal"
      End
      Begin VB.Menu mnuWindowTileVertical 
         Caption         =   "Tile Vertical"
      End
      Begin VB.Menu mnuWindowArrangeIcons 
         Caption         =   "Arrange Icons"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuTips 
         Caption         =   "Accounting Tips"
      End
      Begin VB.Menu mnuOfflinesupport 
         Caption         =   "Off-Line Support"
      End
      Begin VB.Menu mnuBrowserDocs 
         Caption         =   "On-Line Support"
      End
      Begin VB.Menu mnuHelpContents 
         Caption         =   "Contents"
      End
      Begin VB.Menu mnuHelpSearchForHelpOn 
         Caption         =   "Search"
      End
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu mnuAccount 
      Caption         =   "Account"
      Visible         =   0   'False
      Begin VB.Menu mnuLookup 
         Caption         =   "Account Lookup"
      End
      Begin VB.Menu mnuAdd 
         Caption         =   "Add New Data"
      End
      Begin VB.Menu mnuCheck1 
         Caption         =   "Check Mangement"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "Main_Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Const EM_UNDO = &HC7
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hWnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'fMainForm
Public OnGoing As Boolean

Private Sub MDIForm_Load()
Dim WhichToolBar As Integer
Dim ToolBarPic As String
Dim FileExist As String

    ToolBarPic = ""
    Call GetStartUp(WhichToolBar, ToolBarPic)
    If ToolBarPic <> "" Then
      If ToolBarPic = "Default" Then
      ElseIf ToolBarPic = "None" Then
        Set CoolBar1.Picture = Nothing
      Else
        FileExist = Dir(ToolBarPic)
        If FileExist <> "" Then
            Set CoolBar1.Picture = Nothing
            Set CoolBar1.Picture = LoadPicture(ToolBarPic)
        End If
      End If
    End If
    'Put the Toolbar Wrapper controls in the Coolbar band
    Set TransTBWrapper1.Container = CoolBar1
    Set CoolBar1.Bands(1).Child = TransTBWrapper1
    ' put the toolbar into the toolbar wrapper
    If WhichToolBar = 1 Then
        Set TransTBWrapper1.Toolbar = tbToolBar
    Else
        Set TransTBWrapper1.Toolbar = tbToolBar1
    End If
    Picture2.Top = 30
    Picture2.Left = 11020
        
    Me.Left = GetSetting(App.title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.title, "Settings", "MainHeight", 6500)
    'LoadNewDoc
    
    'If gblApplicationConnectString = App.Path & "\properties.nos" Then
    '    fMainForm.txtConnection = "There is no Working database"
    '    MenuStatus False, False, False
    'Else
    '    fMainForm.txtConnection = gblApplicationConnectString
    '    IntroCheckList
    'End If
    
    
End Sub


Private Sub LoadNewDoc()
    'Static lDocumentCount As Long
    'Dim frmD As Main_Document
    'lDocumentCount = lDocumentCount + 1
    'Set frmD = New Main_Document
    'frmD.Caption = "Document " & lDocumentCount
    'frmD.Show
End Sub

Private Sub MDIForm_Resize()
If Me.WindowState = 0 Then
    'sbStatusBar.Panels(3).Width = 1440
    'sbStatusBar.Panels(4).Width = 1440
    'sbStatusBar.Panels(1).Width = (Me.Width - sbStatusBar.Panels(3).Width - sbStatusBar.Panels(4).Width) / 2 - 150
    'sbStatusBar.Panels(2).Width = (Me.Width - sbStatusBar.Panels(3).Width - sbStatusBar.Panels(4).Width) / 2 - 150
End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.title, "Settings", "MainLeft", Me.Left
        SaveSetting App.title, "Settings", "MainTop", Me.Top
        SaveSetting App.title, "Settings", "MainWidth", Me.Width
        SaveSetting App.title, "Settings", "MainHeight", Me.Height
    End If
    CloseAllMDIChild
    'unload the form
    ' it is VERY important to set the wrapper's Toolbar property
    ' to Nothing before the form is unloaded
    CoolBar1.Visible = False
    Set TransTBWrapper1.Toolbar = Nothing
    End
End Sub

Private Sub mnuAccountingBankReconsiliation_Click()
    'MsgBox " Accounting - Bank Reconsiliation", vbInformation, Me.Caption
    frm_Bank_Reconciliation.Show
    frm_Bank_Reconciliation.ZOrder 0
End Sub

Private Sub mnuAccountingBankTransactions_Click()
    'MsgBox " Accounting - Bank Transactions", vbInformation, Me.Caption
    frm_Bank_Transaction.Show
    frm_Bank_Transaction.ZOrder 0
End Sub

Private Sub mnuAccountingChartOfAccounts_Click()
    'Dim f As New frm_SYS_Setup_Chart_Of_Accounts
    frm_SYS_Setup_Chart_Of_Accounts.Show
    frm_SYS_Setup_Chart_Of_Accounts.ZOrder 0
End Sub

Private Sub mnuAccountingGLEntry_Click()
'    MsgBox " Accounting - GL Entry", vbInformation, Me.Caption
    frm_GL_Entry.Show
    frm_GL_Entry.ZOrder 0
End Sub

Private Sub mnuAcctPreferences_Click()
    frm_SYS_Setup_Accounting_Preferences.Show
    frm_SYS_Setup_Accounting_Preferences.ZOrder 0
End Sub

Private Sub mnuAdvertise_Click()
If mnuAdvertise.Checked = True Then
    mnuAdvertise.Checked = False
    Picture2.Visible = False
Else
    mnuAdvertise.Checked = True
    Picture2.Visible = True
End If
End Sub

Private Sub mnuAgeARnAP_Click()
Dim Response%

  Response% = MsgBox("Are you sure want to Age Receivables and Payables?", vbYesNo, "Close Month")
  If Response% = vbNo Then Exit Sub
  
  ShowStatus True
  
  Dim db As ADODB.Connection
  Set db = New ADODB.Connection
  db.CursorLocation = adUseClient
  db.Open gblADOProvider

  Dim RollDate As Variant
  Dim ID$
  
  'Age Receivables & Payables
  'Ask or warn about printing reports?????
  'Set RollDate = last day of fiscal year
  'Loop through each customer and call RollCustomer
  
  RollDate = LookRecord("[SYS COM Fiscal End Date]", "[SYS Company]", db)

  Dim rsCustomer As Recordset
  Set rsCustomer = New ADODB.Recordset
  rsCustomer.Open "SELECT [AR CUST Customer ID] FROM [AR Customer]", db, adOpenKeyset, adLockOptimistic, adCmdText

  'rsCustomer.Index = "PrimaryKey"
  If rsCustomer.RecordCount = 0 Then Exit Sub
  
  db.BeginTrans
  
  rsCustomer.MoveFirst
  Do While Not rsCustomer.EOF
    ID$ = rsCustomer("AR CUST Customer ID")
    Call CustomerRollOver(ID$, RollDate, db)
    rsCustomer.MoveNext
  Loop
  rsCustomer.Close
  Set rsCustomer = Nothing
  
  Dim rsVendor As Recordset
  Set rsVendor = New ADODB.Recordset
  rsVendor.Open "SELECT [AP VEN ID] FROM [AP Vendor]", db, adOpenKeyset, adLockOptimistic, adCmdText

  If rsVendor.RecordCount = 0 Then Exit Sub
  rsVendor.MoveFirst
  Do While Not rsVendor.EOF
    ID$ = rsVendor("AP VEN ID")
    Call VendorRollOver(ID$, RollDate, db)
    rsVendor.MoveNext
  Loop

  ShowStatus False
  'cmdAging.ForeColor = 255
  MsgBox "Age receivables and payables complete.", , "Close Year"
  
  db.CommitTrans
  
  db.Close
  Set db = Nothing

End Sub

Private Sub mnuBackup_Click()
ShowStatus True
'MenuStatus False, False
    Dim sFile As String

    With dlgCommonDialog
        .InitDir = App.Path & "\database"
        .DialogTitle = "Save Working Database As"
        .CancelError = False
        .Flags = cdlOFNPathMustExist + cdlOFNOverwritePrompt
        .Filter = "All Files (*.mdb)|*.mdb"
        .ShowSave
        
        If Len(.FileName) = 0 Then
            ShowStatus False
            Exit Sub
        End If
        sFile = .FileName
    End With
   Dim fs
   Set fs = CreateObject("Scripting.FileSystemObject")
   fs.CopyFile gblApplicationConnectString, sFile
   Set fs = Nothing
'MenuStatus True, True
ShowStatus False
End Sub

Private Sub mnuBackupServer_Click()
    MsgBox "No Connection" & vbCr & gblApplicationConnectString, vbInformation, "Error"
    Exit Sub
'Dim CompactJRO As jro.JetEngine
'Set CompactJRO = New jro.JetEngine
    'CompactJRO.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=d:\\nwind2.mdb", _
"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=d:\\abbc2.mdb;Jet OLEDB:Engine Type=4"
'CompactJRO.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & gblApplicationConnectString, _
'"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & gblApplicationConnectString & "1;Jet OLEDB:Engine Type=4"
'    MsgBox "Your request is done", vbInformation, "Error"

End Sub


Private Sub mnuBrowserDocs_Click()
    callWebPage "http://www.tbstech.com"
'    Main_Browser.Show
End Sub

Public Sub callWebPage(WebAdd As String)
Dim iret As Long
    iret = ShellExecute(Me.hWnd, vbNullString, WebAdd, vbNullString, "c:\", 1)
End Sub


Private Sub mnuCheck_Click()
    frm_Check_Management.Show
    frm_Check_Management.ZOrder 0
End Sub

Private Sub mnuCloseActive_Click()
CloseAllActive = True
    If Forms.count = 1 Then Exit Sub
    Unload ActiveForm
CloseAllActive = False
End Sub

Private Sub mnuCloseAll_Click()
    CloseAllMDIChild
End Sub

Private Sub mnuCloseMonth_Click()
  Dim db As ADODB.Connection
  Set db = New ADODB.Connection
  db.CursorLocation = adUseClient
  db.Open gblADOProvider
  
  Dim Response%

  Response% = MsgBox("About to close month, continue?", vbYesNo, "Close Month")
  If Response% = vbYes Then
    db.BeginTrans
        Call MonthEndSales(db)
        Call MonthEndPurchases(db)
    db.CommitTrans
    'Mark period as closed
    
  '  Dim PeriodToPost%
  '  Dim PeriodClosed%
  '  Dim rsCompany As Recordset
  '  Set rsCompany = db2.OpenRecordset("SYS Company")
  '  Call VerifyPeriod(Now, PeriodToPost%, PeriodClosed%)
  '
  '  rsCompany.Edit
  '    rsCompany("SYS COM P" & Trim(CStr(PeriodToPost%)) & " Closed") = True
  '  rsCompany.UPDATE
  
    MsgBox "Month has been closed.", , "Month End Processing"
    db.Close
    Set db = Nothing
  End If
End Sub

Private Sub mnuDataChartOfAccounts_Click()
    'Dim f As New frm_SYS_Setup_Chart_Of_Accounts
    'f.Show
    frm_SYS_Setup_Chart_Of_Accounts.Show
    frm_SYS_Setup_Chart_Of_Accounts.ZOrder 0
End Sub

Private Sub mnuDataSYS_Setup_Sales_Click()
    'Dim f As New frm_SYS_Setup_Sales
    'f.Show
    frm_SYS_Setup_Sales.Show
    frm_SYS_Setup_Sales.ZOrder 0
End Sub

Private Sub mnuDataSYS_Setup_Purchases_Click()
    'Dim f As New frm_SYS_Setup_Purchases
    'f.Show
    frm_SYS_Setup_Purchases.Show
    frm_SYS_Setup_Purchases.ZOrder 0
End Sub

Private Sub mnuDataSYS_Setup_Periods_Click()
    'Dim f As New frm_SYS_Setup_Periods
    'f.Show
    frm_SYS_Setup_Periods.Show
    frm_SYS_Setup_Periods.ZOrder 0
End Sub

Private Sub mnuDataSYS_Setup_Inventory_Click()
    'Dim f As New frm_SYS_Setup_Inventory
    'f.Show
    frm_SYS_Setup_Inventory.Show
    frm_SYS_Setup_Inventory.ZOrder 0
End Sub


Private Sub mnuDataSYS_Setup_Company_Click()
    'Dim f As New frm_SYS_Setup_Company
    'f.Show
    frm_SYS_Setup_Company.Show
    frm_SYS_Setup_Company.ZOrder 0
End Sub

Private Sub mnuDataSYS_Setup_Checklist_Click()
    'Dim f As New frm_SYS_Setup_Checklist
    'f.Show
    frm_SYS_Setup_Checklist.Show
    frm_SYS_Setup_Checklist.ZOrder 0
End Sub

Private Sub mnuDataSYS_Setup_Banks_Click()
    'Dim f As New frm_SYS_Setup_Banks
    'f.Show
    frm_SYS_Setup_Banks.Show
    frm_SYS_Setup_Banks.ZOrder 0
End Sub

Private Sub mnuDataLIST_Vendor_Types_Click()
    'Dim f As New frm_LIST_Vendor_Types
    'f.Show
    'frm_LIST_Vendor_Types.Show
    frm_LIST_ALL_Types.ListType "frm_LIST_Vendor_Types"
    frm_LIST_ALL_Types.ZOrder 0
End Sub

Private Sub mnuDataLIST_Shipping_Methods_Click()
    'Dim f As New frm_LIST_Shipping_Methods
    'f.Show
    'frm_LIST_Shipping_Methods.Show
    frm_LIST_ALL_Types.ListType "frm_LIST_Shipping_Methods"
    frm_LIST_ALL_Types.ZOrder 0
End Sub

Private Sub mnuDataLIST_Project_Types_Click()
    'Dim f As New frm_LIST_Project_Types
    'f.Show
    frm_LIST_Project_Types.Show
    frm_LIST_Project_Types.ZOrder 0
End Sub

Private Sub mnuDataLIST_Payment_Terms_Click()
    'Dim f As New frm_LIST_Payment_Terms
    'f.Show
    frm_LIST_Payment_Terms.Show
    frm_LIST_Payment_Terms.ZOrder 0
End Sub

Private Sub mnuDataLIST_Payment_Methods_Click()
    'Dim f As New frm_LIST_Payment_Methods
    'f.Show
    'frm_LIST_Payment_Methods.Show
    frm_LIST_ALL_Types.ListType "frm_LIST_Payment_Methods"
    frm_LIST_ALL_Types.ZOrder 0
End Sub

Private Sub mnuDataLIST_Item_Catagories_Click()
    'Dim f As New frm_LIST_Item_Catagories
    'f.Show
    'frm_LIST_Item_Catagories.Show
    frm_LIST_ALL_Types.ListType "frm_LIST_Item_Catagories"
    frm_LIST_ALL_Types.ZOrder 0

End Sub

Private Sub mnuDataLIST_Customer_Types_Click()
    'Dim f As New frm_LIST_Customer_Types
    'f.Show
    frm_LIST_ALL_Types.ListType "frm_LIST_Customer_Types"
    frm_LIST_ALL_Types.ZOrder 0
End Sub

Private Sub mnuDatafrmLISTCreditCards_Click()
    'Dim f As New frm_LIST_Credit_Cards
    'f.Show
    frm_LIST_Credit_Cards.Show
    frm_LIST_Credit_Cards.ZOrder 0
End Sub

Private Sub mnuDataAR_Ship_To_Click()
  frm_AR_Cust_Ship_To.Show
  frm_AR_Cust_Ship_To.ZOrder 0
End Sub

Private Sub mnuDataAR_Customer_Click()
    frm_AR_Customer.Show
    frm_AR_Customer.ZOrder 0
End Sub

Private Sub mnuDataAP_Vendor_Click()
    'Dim f As New frm_AP_Vendor
    'f.Show
    frm_AP_Vendor.Show
    frm_AP_Vendor.ZOrder 0
End Sub

Private Sub mnuErrorLog_Click()
    frm_History.WhichToExpose 0
    frm_History.ZOrder 0
End Sub

Private Sub mnuFinaceCharges_Click()
Dim Found As String
Dim db As ADODB.Connection
  Set db = New ADODB.Connection
  db.CursorLocation = adUseServer
  db.Open gblADOProvider
  
  Found = LookRecord("[SYS COM Finance Charges YN]", "[SYS Company]", db)
  If Trim(Found) = "" Then
    MsgBox "Your company is not set up to assess finance charges!"
    Exit Sub
  End If
  db.Close
  Set db = Nothing
  
  frm_Finance_Charges.Show
  frm_Finance_Charges.ZOrder 0
  
End Sub

Private Sub mnuGLprocess_Click()

  Dim NetProfit@
  Dim Response%
  
  Response% = MsgBox("This process must be done at the end of the year. Are you sure to start the process now?", vbYesNo)
  If Response% = vbNo Then Exit Sub
  
  Dim db As ADODB.Connection
  Set db = New ADODB.Connection
  db.CursorLocation = adUseClient
  db.Open gblADOProvider
  
  'On Error GoTo cmdGL_Click_Error

  'Get the Net Profit
  NetProfit@ = ComputeNetProfit@("CY", 13, db)

  'Get last day of year
  gLastDayOfYear = LookRecord("[SYS COM Fiscal End Date]", "[SYS Company]", db)

  'Get the retained earnings account
  gREA$ = NZ(LookRecord("[SYS COM Retained Earnings Acct]", "[SYS Company]", db))
  'If ValidAccount(gREA$) = False Then
  '  MsgBox "Retained earnings account is not correct in Accounting Preferences!", , "Error"
  '  Exit Sub
  'End If

  ShowStatus True
  
  db.BeginTrans
  
  'Put balances in previous year
  Call RollCOA(db)

  'Roll P&L Accounts
  Call WriteRetainedEarnings(db)

  'Roll Balance Sheet Accounts
  Call RollCOABalances(db)
  
  'Delete GL Transactions
  Call DeleteGLTrans(db)
  
  'Change APT type to BEGBAL
  Dim rs As Recordset
  Set rs = New ADODB.Recordset
  rs.Open "SELECT [GL Trans Type] FROM [GL Transaction] where [GL Trans Type] = 'APT'", db, adOpenKeyset, adLockOptimistic, adCmdText
  
  'On Error Resume Next
  If rs.RecordCount > 0 Then
    rs.MoveFirst
    Do While Not rs.EOF
      'rs.Edit
        rs("GL Trans Type") = "BEGBAL"
      rs.Update
      rs.MoveNext
    Loop
  End If
  rs.Close
  
  Dim SQLstatement As String
  Dim rsCompany As Recordset
  Set rsCompany = New ADODB.Recordset
  
  Dim X%
  For X% = 1 To 13
    SQLstatement = SQLstatement & "[SYS COM P" & Trim(Str(X%)) & " Date]," & "[SYS COM P" & Trim(Str(X%)) & " Closed],"
  Next X%
  
  rsCompany.Open "SELECT [SYS COM Fiscal Start Date],[SYS COM Fiscal End Date]," & _
  "[SYS COM Fiscal Year]," & SQLstatement & " FORM [SYS Company]", db, adOpenKeyset, adLockOptimistic, adCmdText

  'Compute new GL Periods and open all of them
  Dim TempDate As Variant
  For X% = 1 To 13
    rsCompany.MoveFirst
    TempDate = rsCompany("SYS COM P" & Trim(Str(X%)) & " Date")
    TempDate = DateAdd("yyyy", 1, TempDate)
    'rsCompany.Edit
      rsCompany("SYS COM P" & Trim(Str(X%)) & " Date") = Format(TempDate, "Short Date")
      rsCompany("SYS COM P" & Trim(Str(X%)) & " Closed") = False
    rsCompany.Update
  Next X%

  'Hit Fiscal Start Date
  'rsCompany.MoveFirst
  TempDate = rsCompany("SYS COM Fiscal Start Date")
  TempDate = DateAdd("yyyy", 1, TempDate)
  'rsCompany.Edit
    rsCompany("SYS COM Fiscal Start Date") = Format(TempDate, "Short Date")
  rsCompany.Update
  
  'Hit Fiscal End Date
  'rsCompany.MoveFirst
  TempDate = rsCompany("SYS COM Fiscal End Date")
  TempDate = DateAdd("yyyy", 1, TempDate)
  'rsCompany.Edit
    rsCompany("SYS COM Fiscal End Date") = Format(TempDate, "Short Date")
  rsCompany.Update
  
  'Hit Fiscal Year
  'rsCompany.MoveFirst
  TempDate = rsCompany("SYS COM Fiscal Year")
  TempDate = TempDate + 1
  'rsCompany.Edit
    rsCompany("SYS COM Fiscal Year") = TempDate
  rsCompany.Update

  'Do COA processing from Company Card to rewrite future balances
  Call ResetCYGLBalances(db)
  
  db.CommitTrans
  
  ShowStatus False '22-2 first floor mm  cmdGL.ForeColor = 255
  MsgBox "Year end processing complete.", , "Year End Assistant"
  
  db.Close
  Set db = Nothing
  
  Exit Sub
cmdGL_Click_Error:
  Call ErrorLog("fmainform", "mnuGLprocess_Click", Now, Err.Number, Err.Description, True, db)
  Resume Next

End Sub

Private Sub mnuHistory_Click()
    frm_History.WhichToExpose 1
    frm_History.ZOrder 0
End Sub

Private Sub mnuInventoryInventoryAdjustment_Click()
    frm_INV_Adjust.Show
    frm_INV_Adjust.ZOrder 0
End Sub

Private Sub mnuInventoryInventoryProduction_Click()
    frm_INV_Production.Show
    frm_INV_Production.ZOrder
End Sub

Private Sub mnuInventoryItems_Click()
    'Dim f As New frm_SYS_Setup_Items
    'f.Show
    frm_SYS_Setup_Items.Show
    frm_SYS_Setup_Items.ZOrder 0
End Sub

Private Sub mnuNewDb_Click()
ShowStatus True
'MenuStatus False, False
    Dim sFile As String

    With dlgCommonDialog
        .InitDir = App.Path & "\database"
        .DialogTitle = "Save Working Database As"
        .CancelError = False
        .Flags = cdlOFNPathMustExist + cdlOFNOverwritePrompt
        .Filter = "All Files (*.mdb)|*.mdb"
        .ShowSave
        
        If Len(.FileName) = 0 Then
            ShowStatus False
            Exit Sub
        End If
        sFile = .FileName
    End With
   Dim fs
   Set fs = CreateObject("Scripting.FileSystemObject")
   fs.CopyFile gblApplicationConnectString, sFile
   Set fs = Nothing
   'CloseAllMDIChild
   'DbConnectionString sFile
   
    'If gblApplicationConnectString = App.Path & "\properties.nos" Then
    '    fMainForm.txtConnection = "There is no Working database"
    'Else
    '    fMainForm.txtConnection = gblApplicationConnectString
    'End If
'MenuStatus True, True
ShowStatus False
End Sub

Private Sub mnuOfflinesupport_Click()
    frm_Help.Show
    frm_Help.ZOrder 0
End Sub

'Private Sub mnuPayrollPayEmployees_Click()
'    frm_Pay_Employees.Show
'    frm_Pay_Employees.ZOrder 0
'End Sub

'Private Sub mnuPayrollType_Click()
'    frm_SYS_Setup_Payroll.Show
'    frm_SYS_Setup_Payroll.ZOrder 0
'End Sub

'Private Sub mnuPayrollViodChecks_Click()
'    frm_Pay_Voids.Show
'    frm_Pay_Voids.ZOrder 0
    'MsgBox "Payroll - Void Checks", vbInformation, Me.Caption
'End Sub

Private Sub mnuProjects_Click()
    'Dim f As New frm_AR_Cust_Projects
    'f.Show
    frm_AR_Cust_Projects.Show
    frm_AR_Cust_Projects.ZOrder 0
End Sub

Private Sub mnuPurchasingBatchPosting_Click()
    frm_AP_Batch_Posting.Show
    frm_AP_Batch_Posting.ZOrder 0
End Sub

Private Sub mnuPurchasingCashPayments_Click()
    frm_AP_Cash_Payments.Show
    frm_AP_Cash_Payments.ZOrder 0
End Sub

Private Sub mnuPurchasingCreditMemo_Click()
    'Dim f As New frm_AP_Credit_Entry
    frm_AP_Credit_Entry.Show
    frm_AP_Credit_Entry.ZOrder 0
End Sub

Private Sub mnuPurchasingPayManyVendors_Click()
    'Dim f As New frm_AP_Pay_Many_Vendors
    frm_AP_Pay_Many_Vendors.Show
    frm_AP_Pay_Many_Vendors.ZOrder 0
End Sub

Private Sub mnuPurchasingPurchases_Click()
    'Dim f As New frm_AP_Purchase_Entry
    frm_AP_Purchase_Entry.Show
    frm_AP_Purchase_Entry.ZOrder 0
End Sub


Private Sub mnuPurchasingReceiving_Click()
    'Dim f As New frm_AP_Receiving_Entry
    frm_AP_Receiving_Entry.Show
    frm_AP_Receiving_Entry.ZOrder 0
End Sub

Private Sub mnuPurchasingVoucher_Click()
    'Dim f As New frm_AP_Voucher_Entry
    frm_AP_Voucher_Entry.Show
    frm_AP_Voucher_Entry.ZOrder 0
End Sub

Private Sub mnuRecurrGL_Click()
    frm_Recurring.RequestType 1
End Sub

Private Sub mnuRecurring_Click()
    'frm_LIST_ALL_Types.Show
    frm_LIST_ALL_Types.ListType "frm_LIST_ALL_Types"
    frm_LIST_ALL_Types.ZOrder 0
End Sub


Private Sub mnuRecurrPayments_Click()
    frm_Recurring.RequestType 4
End Sub

Private Sub mnuRecurrPurchases_Click()
    frm_Recurring.RequestType 3
End Sub

Private Sub mnuRecurrSales_Click()
    frm_Recurring.RequestType 2
End Sub

Private Sub mnuRMA_Click()
    frm_AP_RMA_Entry.Show
    frm_AP_RMA_Entry.ZOrder 0
End Sub

Private Sub mnuSalesBatchPosting_Click()
    'Dim f As New frm_AR_Batch_Posting
  'If frm_AR_Batch_Posting.WindowState = 0 Then Exit Sub
  If LoadBatchSalesDetail = True Then
    frm_AR_Batch_Posting.Show
    frm_AR_Batch_Posting.ZOrder 0
  End If
End Sub
Private Function LoadBatchSalesDetail() As Boolean
Dim db As ADODB.Connection
  Set db = New ADODB.Connection
  db.CursorLocation = adUseServer
  db.Open gblADOProvider

  'On Error GoTo LoadBatchSalesDetail_Error

  Dim rsBatch As ADODB.Recordset
  Set rsBatch = New ADODB.Recordset
  rsBatch.Open "[SYS Sales Batch]", db, adOpenStatic, adLockOptimistic, adCmdTable
  
  Dim rsSales As ADODB.Recordset
  Set rsSales = New ADODB.Recordset
  rsSales.Open "SELECT * FROM [AR Sales] where [AR SALE Posted YN] = false", db, adOpenStatic, adLockOptimistic
  
  db.Execute "DELETE * FROM [SYS Sales Batch]"
    
  If rsSales.RecordCount = 0 Then
    MsgBox "There are no unposted sales transactions."
    LoadBatchSalesDetail = False
    Exit Function
  End If
  
  ShowStatus True
  
  rsSales.MoveFirst
  Do While Not rsSales.EOF
    rsBatch.AddNew
      rsBatch("Post YN") = True
      rsBatch("Document ID") = rsSales("AR SALE Ext Document #")
      rsBatch("Customer ID") = rsSales("AR SALE Customer ID")
      rsBatch("Document Type") = rsSales("AR SALE Document Type")
      rsBatch("Date") = rsSales("AR SALE Date")
      rsBatch("Amount") = rsSales("AR SALE Total")
    rsBatch.Update
    rsSales.MoveNext
  Loop
  
  ShowStatus False
  
  LoadBatchSalesDetail = True
  
  Set rsBatch = Nothing
  Set rsSales = Nothing
  db.Close
  Set db = Nothing
  Exit Function
LoadBatchSalesDetail_Error:
  Call ErrorLog("AR Batch Posting", "LoadBatchSalesDetail", Now, Err.Number, Err.Description, True, db)
  Resume Next

End Function

Private Sub mnuSalesCashReceipts_Click()
    'Dim f As New frm_AR_Cash_Receipts
    frm_AR_Cash_Receipts.Show
    frm_AR_Cash_Receipts.ZOrder 0
End Sub

Private Sub mnuSalesCreditMemo_Click()
    'Dim f As New frm_AR_Credit_Entry
    frm_AR_Credit_Entry.Show
    frm_AR_Credit_Entry.ZOrder 0
End Sub

Private Sub mnuSalesMemo_Click()
    frm_AR_Sales_Memo_Entry.Show
    frm_AR_Sales_Memo_Entry.ZOrder 0
End Sub

Private Sub mnuSalesOrderEntry_Click()
    'Dim f As New frm_AR_Order_Entry
    frm_AR_Order_Entry.Show
    frm_AR_Order_Entry.ZOrder 0
End Sub

Private Sub mnuSalesQuoteEntry_Click()
    'Dim f As New frm_AR_Quote_Entry
    frm_AR_Quote_Entry.Show
    frm_AR_Quote_Entry.ZOrder 0
End Sub

Private Sub mnuSalesReturnEntry_Click()
    'Dim f As New frm_AR_Return_Entry
    frm_AR_Return_Entry.Show
    frm_AR_Return_Entry.ZOrder 0
End Sub

Private Sub mnuSalesSalesEntry_Click()
    'Dim f As New frm_AR_Sales_Entry
    frm_AR_Sales_Entry.Show
    frm_AR_Sales_Entry.ZOrder 0
End Sub

Private Sub mnuShowReport_Click()
    frm_prnPreview.Show
    frm_prnPreview.ZOrder 0
End Sub

Private Sub mnuTaxAuthorities_Click()
    frm_SYS_Setup_Tax_Authorities.Show
    frm_SYS_Setup_Tax_Authorities.ZOrder 0
End Sub

Private Sub mnuTaxGroup_Click()
    frm_SYS_Setup_Tax_Group.Show
    frm_SYS_Setup_Tax_Group.ZOrder
End Sub

Private Sub mnuTips_Click()
    frmTip.TypeOfWise 2
    frmTip.ZOrder 0
End Sub

Private Sub mnuTips1_Click()
    frmTip.TypeOfWise 1
    frmTip.ZOrder 0
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As ComctlLib.Button)
    'On Error Resume Next
    Select Case Button.Key
        Case "New"
            mnuFileNew_Click
        Case "Open"
            mnuFileOpen_Click
        Case "Save"
            mnuBackup_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Cut"
            mnuEditCut_Click
        Case "Copy"
            mnuEditCopy_Click
        Case "Paste"
            mnuEditPaste_Click
        Case "Sort Ascending"
            'ToDo: Add 'Sort Ascending' button code.
            MsgBox "Add 'Sort Ascending' button code."
        Case "Sort Descending"
            'ToDo: Add 'Sort Descending' button code.
            MsgBox "Add 'Sort Descending' button code."
        Case "Calendar"
            'Open the Calendar Form When Button is Pressed
            Menu_Calendar.Show vbModal
            Menu_Calendar.ZOrder 0
        Case "Contacts"
            'Open the Contacts Form When Button is Pressed
            'frm_Menu_Contacts.Show
            frm_AR_Customer.ShowList
        Case "InventoryList"
            'Open the Inventory Summary Form When Button is Pressed
            'frm_Menu_Inventory.Show
            frm_SYS_Setup_Items.PriceList
        Case "Help"
            frm_Help.Show
        Case "stop"
            mnuFileClose_Click
        End Select
End Sub

Private Sub mnuHelpAbout_Click()
    Main_About.Show vbModal, Me
End Sub

Private Sub mnuHelpSearchForHelpOn_Click()
'    Dim nRet As Integer
    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
'    If Len(App.HelpFile) = 0 Then
'        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
'    Else
        'On Error Resume Next
'        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 261, 0)
'        If Err Then
'            MsgBox Err.Description
'        End If
'    End If

End Sub

Private Sub mnuHelpContents_Click()
'    Dim nRet As Integer
    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
'    If Len(App.HelpFile) = 0 Then
'        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
'    Else
'        'On Error Resume Next
'        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
'        If Err Then
'            MsgBox Err.Description
'        End If
'    End If

End Sub


Private Sub mnuWindowArrangeIcons_Click()
    Me.Arrange vbArrangeIcons
End Sub

Private Sub mnuWindowTileVertical_Click()
    Me.Arrange vbTileVertical
End Sub

Private Sub mnuWindowTileHorizontal_Click()
    Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuWindowCascade_Click()
    Me.Arrange vbCascade
End Sub

Private Sub mnuToolsOptions_Click()
    Main_Options.Show vbModal, Me
End Sub

Private Sub mnuViewOptions_Click()
    Main_Options.Show vbModal, Me
End Sub

Private Sub mnuViewRefresh_Click()
    'ToDo: Add 'mnuViewRefresh_Click' code.
    MsgBox "Add 'mnuViewRefresh_Click' code."
End Sub

Private Sub mnuViewStatusBar_Click()
    mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
    picStatusBar.Visible = mnuViewStatusBar.Checked
End Sub

Private Sub mnuViewToolbar_Click()
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    CoolBar1.Visible = mnuViewToolbar.Checked
End Sub

Private Sub mnuEditPaste_Click()
On Error Resume Next
    If Forms.count > 1 Then
        If Me.ActiveForm.ActiveControl <> "False" Then
            If Me.ActiveForm.ActiveControl.Locked = True Or Me.ActiveForm.ActiveControl.Enabled = False Then
                MsgBox "Cannot Paste Locked or disable controls", vbInformation, "Information"
                Exit Sub
            End If
           Me.ActiveForm.ActiveControl = Clipboard.GetText
        End If
    Else
        MsgBox "No active Form loaded", vbInformation, "Information"
    End If
End Sub

Private Sub mnuEditCopy_Click()
On Error Resume Next
'Dim FrmActive As Form
Dim TempText As TextBox

'Set FrmActive = Me.ActiveForm.ActiveControl
    If Forms.count > 1 Then
        If Me.ActiveForm.ActiveControl <> "False" Then
            Clipboard.SetText Me.ActiveForm.ActiveControl.SelText
            If Me.ActiveForm.ActiveControl.SelText = "" Then
                mnuEditPaste.Enabled = False
                tbToolBar.Buttons("Paste").Enabled = False
            Else
                mnuEditPaste.Enabled = True
                tbToolBar.Buttons("Paste").Enabled = True
            End If
        End If
    Else
        MsgBox "No active Form Loaded", vbInformation, "Information"
    End If

End Sub

Private Sub mnuEditCut_Click()
On Error Resume Next
'Dim FrmActive As Form

    If Forms.count > 1 Then
        'Set FrmActive = Me.ActiveForm.ActiveControl
        If Me.ActiveForm.ActiveControl <> "False" Then
            If Me.ActiveForm.ActiveControl.Locked = True Or Me.ActiveForm.ActiveControl.Enabled = False Then
                MsgBox "Cannot cut Locked or disable controls", vbInformation, "Information"
                Exit Sub
            End If
            Clipboard.SetText Me.ActiveForm.ActiveControl.SelText
            If Me.ActiveForm.ActiveControl.SelText = "" Then
                mnuEditPaste.Enabled = False
                tbToolBar.Buttons("Paste").Enabled = False
            Else
                mnuEditPaste.Enabled = True
                tbToolBar.Buttons("Paste").Enabled = True
                Clipboard.SetText Me.ActiveForm.ActiveControl
                Me.ActiveForm.ActiveControl.SelText = vbNullString
            End If
        End If
    Else
        MsgBox "No active Form loaded", vbInformation, "Information"
    End If
    
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileSend_Click()
    'ToDo: Add 'mnuFileSend_Click' code.
    MsgBox "Add 'mnuFileSend_Click' code."
End Sub

Private Sub mnuFilePrint_Click()
    If Forms.count > 1 Then
    Else
        MsgBox "There is nothing to print", vbInformation, "Information"
    End If
    'On Error Resume Next
'    If ActiveForm Is Nothing Then Exit Sub
    

'    With dlgCommonDialog
'        .DialogTitle = "Print"
'        .CancelError = True
'        .Flags = cdlPDReturnDC + cdlPDNoPageNums
'        If ActiveForm.rtfText.SelLength = 0 Then
'            .Flags = .Flags + cdlPDAllPages
'        Else
'            .Flags = .Flags + cdlPDSelection
'        End If
'        .ShowPrinter
        'If Err <> MSComDlg.cdlCancel Then
        '    ActiveForm.rtfText.SelPrint .hDC
        'End If
'    End With

End Sub

Private Sub mnuFilePrintPreview_Click()
    If Forms.count > 1 Then
    Else
        MsgBox "There is nothing to print", vbInformation, "Information"
    End If
End Sub

Private Sub mnuFileProperties_Click()
    Main_Options.Show vbModal, Me
End Sub

Private Sub mnuFileSaveAll_Click()
    'ToDo: Add 'mnuFileSaveAll_Click' code.
    MsgBox "Add 'mnuFileSaveAll_Click' code."
End Sub

Private Sub mnuFileSave_Click()
    Dim sFile As String
    If Left$(ActiveForm.Caption, 8) = "Document" Then
        With dlgCommonDialog
            .DialogTitle = "Save"
            .CancelError = False
            'ToDo: set the flags and attributes of the common dialog control
            .Filter = "All Files (*.*)|*.*"
            .ShowSave
            If Len(.FileName) = 0 Then
                Exit Sub
            End If
            sFile = .FileName
        End With
        ActiveForm.rtfText.SaveFile sFile
    Else
        sFile = ActiveForm.Caption
        ActiveForm.rtfText.SaveFile sFile
    End If

End Sub

Private Sub mnuFileOpen_Click()
ShowStatus True
'MenuStatus False, False, False
    Dim sFile As String

    With dlgCommonDialog
        .DialogTitle = "Open Working Database"
        .CancelError = False
        .Flags = cdlOFNPathMustExist + cdlOFNOverwritePrompt
        .Filter = "All Files (*.mdb)|*.mdb"
        .ShowOpen
        If Len(.FileName) = 0 Then
            ShowStatus False
            Exit Sub
        End If
        sFile = .FileName
    End With
If dlgCommonDialog.Flags = 3074 Then
   If gblApplicationConnectString = sFile Then
        MsgBox "That is current working database.", vbInformation, "Information"
        ShowStatus False
        Exit Sub
   Else
        CloseAllMDIChild
        DbConnectionString sFile
        SaveCompany
         If gblApplicationConnectString = App.Path & "\properties.nos" Then
             fMainForm.txtConnection = "There is no Working database"
         Else
             fMainForm.txtConnection = gblApplicationConnectString
         End If
   End If
'MenuStatus True, True
End If
ShowStatus False
End Sub

Private Sub mnuFileNew_Click()
ShowStatus True
'MenuStatus False, False
    Dim sFile As String

    With dlgCommonDialog
        .InitDir = App.Path & "\database"
        .DialogTitle = "New Working Database"
        .CancelError = False
        .Flags = cdlOFNPathMustExist + cdlOFNOverwritePrompt
        .Filter = "All Files (*.mdb)|*.mdb"
        .FilterIndex = 2
        .ShowSave
        If Len(.FileName) = 0 Then
            ShowStatus False
            Exit Sub
        End If
        sFile = .FileName
   'MsgBox .Flags
   'Exit Sub
   End With
If dlgCommonDialog.Flags = 3074 Then
   Dim fs
   Set fs = CreateObject("Scripting.FileSystemObject")
   fs.CopyFile App.Path & "\masterDB.mdb", sFile
   Set fs = Nothing
   CloseAllMDIChild
   DbConnectionString sFile
   SaveCompany
    If gblApplicationConnectString = App.Path & "\properties.nos" Then
        fMainForm.txtConnection = "There is no Working database"
    Else
        fMainForm.txtConnection = gblApplicationConnectString
    End If
   'MenuStatus True, True
End If
ShowStatus False
End Sub

Public Sub CloseAllMDIChild()
On Error GoTo exit_sub
CloseAllActive = True
'MsgBox Me.ActiveForm
While Forms.count > 1

    Unload Me.ActiveForm
    
Wend
exit_sub:
CloseAllActive = False
End Sub

Private Sub mnuFileClose_Click()
   CloseAllMDIChild
   DbConnectionString App.Path & "\properties.nos"
   gblADOProvider = gblBasicADOProvider
   fMainForm.txtConnection = "No Connection"
   MenuStatus False, False, False
    'MsgBox "Add 'mnuFileClose_Click' code."
End Sub

Public Sub MenuStatus(EnabledStatus As Boolean, EnabledFile As Boolean, CheckLiskIntro As Boolean)

    mnuEdit.Enabled = EnabledStatus
    mnuSetup.Enabled = EnabledStatus
    
        mnuSales.Enabled = CheckLiskIntro
        mnuPurchasing.Enabled = CheckLiskIntro
        mnuInventory.Enabled = CheckLiskIntro
        mnuAccounting.Enabled = CheckLiskIntro
        'mnuPayroll.Enabled = CheckLiskIntro
        'mnuShowReport.Enabled = CheckLiskIntro
        mnuFinaceCharges.Enabled = CheckLiskIntro
        mnuCloseMonth.Enabled = CheckLiskIntro
    'mnuTools.Enabled = EnabledStatus
    'mnuWindow.Enabled = EnabledStatus
    'mnuHelp.Enabled = EnabledStatus
    tbToolBar.Buttons("InventoryList").Enabled = EnabledStatus
    tbToolBar.Buttons("Contacts").Enabled = EnabledStatus
    
    tbToolBar.Buttons("Save").Enabled = EnabledStatus
    tbToolBar.Buttons("stop").Enabled = EnabledStatus
    
    mnuFileClose.Enabled = EnabledFile
    mnuFileSaveAs.Enabled = EnabledFile
    mnuFileProperties.Enabled = EnabledFile
    mnuHistory.Enabled = EnabledFile
    mnuErrorLog.Enabled = EnabledFile
    'mnuFilePageSetup.Enabled = EnabledFile
    mnuFilePrintPreview.Enabled = EnabledFile
    mnuFilePrint.Enabled = EnabledFile
    
    'mnuPayrollPayEmployees.Enabled = False
    'mnuPayrollViodChecks.Enabled = False
    'mnuFileSend.Enabled = False
End Sub

Private Sub tbToolBar1_ButtonClick(ByVal Button As ComctlLib.Button)
On Error Resume Next
    Select Case Button.Key
        Case "New"
            mnuFileNew_Click
        Case "Open"
            mnuFileOpen_Click
        Case "Save"
            mnuBackup_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Cut"
            mnuEditCut_Click
        Case "Copy"
            mnuEditCopy_Click
        Case "Paste"
            mnuEditPaste_Click
        Case "Sort Ascending"
            'ToDo: Add 'Sort Ascending' button code.
            MsgBox "Add 'Sort Ascending' button code."
        Case "Sort Descending"
            'ToDo: Add 'Sort Descending' button code.
            MsgBox "Add 'Sort Descending' button code."
        Case "Calendar"
            'Open the Calendar Form When Button is Pressed
            Menu_Calendar.Show
            Menu_Calendar.ZOrder 0
        Case "Contacts"
            'Open the Contacts Form When Button is Pressed
            'frm_Menu_Contacts.Show
            frm_AR_Customer.ShowList
        Case "InventoryList"
            'Open the Inventory Summary Form When Button is Pressed
            frm_SYS_Setup_Items.PriceList
        Case "Help"
            frm_Help.Show
        Case "stop"
            mnuFileClose_Click
        End Select
End Sub

Private Sub Timer1_Timer()
    If Me.Visible = True Then
        If gblApplicationConnectString = App.Path & "\properties.nos" Or gblApplicationConnectString = "" Then
            fMainForm.txtConnection = "There is no Working database"
            MenuStatus False, False, False
        Else
            fMainForm.txtConnection = gblApplicationConnectString
            IntroCheckList
        End If
        frmTip.TypeOfWise 0
    End If
End Sub

Private Sub txtConnection_GotFocus()
    TxtGotFocus txtConnection
End Sub

Private Sub txtLogon_GotFocus()
    TxtGotFocus txtLogon
End Sub
