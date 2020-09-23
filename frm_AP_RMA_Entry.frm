VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_AP_RMA_Entry 
   Caption         =   "RMA Entry"
   ClientHeight    =   8850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15045
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8850
   ScaleWidth      =   15045
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox picPrimary 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8550
      Left            =   0
      ScaleHeight     =   8550
      ScaleWidth      =   15075
      TabIndex        =   0
      Top             =   480
      Width           =   15075
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   2640
         ScaleHeight     =   330
         ScaleWidth      =   1935
         TabIndex        =   113
         Top             =   360
         Width           =   1935
         Begin VB.CommandButton cmdLookupVEND 
            Height          =   285
            Left            =   0
            Picture         =   "frm_AP_RMA_Entry.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Get Customer"
            Top             =   0
            Width           =   375
         End
         Begin VB.Label lblweb 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "WebPage"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   480
            MouseIcon       =   "frm_AP_RMA_Entry.frx":030A
            MousePointer    =   99  'Custom
            TabIndex        =   115
            Top             =   60
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Label lblmail 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "E-Mail"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   1440
            MouseIcon       =   "frm_AP_RMA_Entry.frx":0614
            MousePointer    =   99  'Custom
            TabIndex        =   114
            Top             =   60
            Visible         =   0   'False
            Width           =   435
         End
      End
      Begin VB.PictureBox picMajorbutton 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   8140
         ScaleHeight     =   375
         ScaleWidth      =   900
         TabIndex        =   51
         Top             =   240
         Width           =   900
         Begin VB.CommandButton cmdSmallBig 
            Caption         =   ">>"
            Height          =   375
            Left            =   440
            TabIndex        =   57
            ToolTipText     =   "Enlarge/Shrink"
            Top             =   0
            Width           =   460
         End
         Begin VB.CommandButton Command2 
            Height          =   375
            Left            =   0
            Picture         =   "frm_AP_RMA_Entry.frx":091E
            Style           =   1  'Graphical
            TabIndex        =   119
            ToolTipText     =   "Enlarge/Shrink"
            Top             =   0
            Width           =   460
         End
      End
      Begin VB.PictureBox pcMajor 
         BorderStyle     =   0  'None
         Height          =   6375
         Left            =   50
         ScaleHeight     =   6375
         ScaleWidth      =   11535
         TabIndex        =   73
         Top             =   0
         Width           =   11535
         Begin VB.Frame frFirst 
            Height          =   3015
            Left            =   0
            TabIndex        =   78
            Top             =   0
            Width           =   11535
            Begin VB.TextBox txtFieldsVendor 
               DataField       =   "AP PO Vendor ID"
               Height          =   285
               Index           =   0
               Left            =   1200
               Locked          =   -1  'True
               TabIndex        =   1
               Top             =   360
               Width           =   1455
            End
            Begin VB.Frame Frame1 
               Height          =   2775
               Left            =   9120
               TabIndex        =   79
               Top             =   120
               Width           =   2295
               Begin VB.TextBox txtfields 
                  DataField       =   "AP PO Vendor Invoice No"
                  Height          =   285
                  Index           =   14
                  Left            =   1080
                  TabIndex        =   3
                  Top             =   1320
                  Width           =   1095
               End
               Begin VB.TextBox txtSalesPerson 
                  DataField       =   "AP PO Ordered By"
                  Height          =   285
                  Index           =   0
                  Left            =   1080
                  Locked          =   -1  'True
                  TabIndex        =   35
                  Top             =   2400
                  Width           =   1095
               End
               Begin VB.TextBox txtfields 
                  DataField       =   "AP PO Date"
                  BeginProperty DataFormat 
                     Type            =   1
                     Format          =   "MM/dd/yyyy"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   3
                  EndProperty
                  Height          =   285
                  Index           =   4
                  Left            =   1080
                  Locked          =   -1  'True
                  TabIndex        =   34
                  Top             =   2040
                  Width           =   1095
               End
               Begin VB.TextBox txtfields 
                  DataField       =   "AP PO Ext Document No"
                  Height          =   285
                  Index           =   12
                  Left            =   1080
                  Locked          =   -1  'True
                  TabIndex        =   33
                  Top             =   1680
                  Width           =   1095
               End
               Begin VB.Image imgOpen 
                  Height          =   495
                  Left            =   120
                  Picture         =   "frm_AP_RMA_Entry.frx":0C28
                  Top             =   240
                  Visible         =   0   'False
                  Width           =   1980
               End
               Begin VB.Label lblLabels 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000004&
                  Caption         =   "Vendor Inv:"
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   14
                  Left            =   120
                  TabIndex        =   83
                  Top             =   1320
                  Width           =   975
               End
               Begin VB.Image imgPosted 
                  Height          =   540
                  Left            =   120
                  Picture         =   "frm_AP_RMA_Entry.frx":14E8
                  Top             =   240
                  Width           =   2250
               End
               Begin VB.Label lblLabels 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000000&
                  Caption         =   "Returned By:"
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   28
                  Left            =   120
                  TabIndex        =   82
                  Top             =   2400
                  Width           =   975
               End
               Begin VB.Label lblLabels 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000004&
                  Caption         =   "Date:"
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   4
                  Left            =   120
                  TabIndex        =   81
                  Top             =   2040
                  Width           =   975
               End
               Begin VB.Label lblLabels 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000004&
                  Caption         =   "RMA No:"
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   12
                  Left            =   120
                  TabIndex        =   80
                  Top             =   1680
                  Width           =   975
               End
            End
            Begin VB.TextBox txtfields 
               DataField       =   "AP PO Document Type"
               Height          =   285
               Index           =   3
               Left            =   9600
               Locked          =   -1  'True
               TabIndex        =   85
               Top             =   1200
               Width           =   1335
            End
            Begin VB.TextBox txtfields 
               DataField       =   "AP PO Document No"
               Height          =   285
               Index           =   2
               Left            =   9600
               Locked          =   -1  'True
               TabIndex        =   84
               Top             =   1800
               Width           =   1335
            End
            Begin VB.TextBox txtFieldsVendor 
               DataField       =   "AP PO Remit Fax"
               Height          =   285
               Index           =   18
               Left            =   7800
               Locked          =   -1  'True
               TabIndex        =   32
               Top             =   2520
               Width           =   1215
            End
            Begin VB.TextBox txtFieldsVendor 
               DataField       =   "AP PO Remit Phone"
               Height          =   285
               Index           =   17
               Left            =   5760
               Locked          =   -1  'True
               TabIndex        =   31
               Top             =   2520
               Width           =   1215
            End
            Begin VB.TextBox txtFieldsVendor 
               DataField       =   "AP PO Fax"
               Height          =   285
               Index           =   9
               Left            =   3240
               Locked          =   -1  'True
               TabIndex        =   23
               Top             =   2520
               Width           =   1215
            End
            Begin VB.TextBox txtFieldsVendor 
               DataField       =   "AP PO Phone"
               Height          =   285
               Index           =   8
               Left            =   1200
               Locked          =   -1  'True
               TabIndex        =   22
               Top             =   2520
               Width           =   1215
            End
            Begin VB.TextBox txtFieldsVendor 
               DataField       =   "AP PO Remit Address 1"
               Height          =   285
               Index           =   11
               Left            =   5760
               Locked          =   -1  'True
               TabIndex        =   25
               Top             =   1080
               Width           =   3255
            End
            Begin VB.TextBox txtFieldsVendor 
               DataField       =   "AP PO Remit Name"
               Height          =   285
               Index           =   10
               Left            =   5760
               Locked          =   -1  'True
               TabIndex        =   24
               Top             =   720
               Width           =   3255
            End
            Begin VB.TextBox txtFieldsVendor 
               DataField       =   "AP PO Address 1"
               Height          =   285
               Index           =   2
               Left            =   1200
               Locked          =   -1  'True
               TabIndex        =   16
               Top             =   1080
               Width           =   3255
            End
            Begin VB.TextBox txtFieldsVendor 
               DataField       =   "AP PO Address 2"
               Height          =   285
               Index           =   3
               Left            =   1200
               Locked          =   -1  'True
               TabIndex        =   17
               Top             =   1440
               Width           =   3255
            End
            Begin VB.TextBox txtFieldsVendor 
               DataField       =   "AP PO City"
               Height          =   285
               Index           =   4
               Left            =   1200
               Locked          =   -1  'True
               TabIndex        =   18
               Top             =   1800
               Width           =   975
            End
            Begin VB.TextBox txtFieldsVendor 
               DataField       =   "AP PO Postal"
               Height          =   285
               Index           =   6
               Left            =   3720
               Locked          =   -1  'True
               TabIndex        =   20
               Top             =   1800
               Width           =   735
            End
            Begin VB.TextBox txtFieldsVendor 
               DataField       =   "AP PO State"
               Height          =   285
               Index           =   5
               Left            =   2760
               Locked          =   -1  'True
               TabIndex        =   19
               Top             =   1800
               Width           =   495
            End
            Begin VB.TextBox txtFieldsVendor 
               DataField       =   "AP PO Country"
               Height          =   285
               Index           =   7
               Left            =   1200
               Locked          =   -1  'True
               TabIndex        =   21
               Top             =   2160
               Width           =   3255
            End
            Begin VB.TextBox txtFieldsVendor 
               DataField       =   "AP PO Vendor Name"
               Height          =   285
               Index           =   1
               Left            =   1200
               Locked          =   -1  'True
               TabIndex        =   15
               Top             =   720
               Width           =   3255
            End
            Begin VB.TextBox txtFieldsVendor 
               DataField       =   "AP PO Remit Country"
               Height          =   285
               Index           =   16
               Left            =   5760
               Locked          =   -1  'True
               TabIndex        =   30
               Top             =   2160
               Width           =   3255
            End
            Begin VB.TextBox txtFieldsVendor 
               DataField       =   "AP PO Remit State"
               Height          =   285
               Index           =   14
               Left            =   7320
               Locked          =   -1  'True
               TabIndex        =   28
               Top             =   1800
               Width           =   495
            End
            Begin VB.TextBox txtFieldsVendor 
               DataField       =   "AP PO Remit Postal"
               Height          =   285
               Index           =   15
               Left            =   8280
               Locked          =   -1  'True
               TabIndex        =   29
               Top             =   1800
               Width           =   735
            End
            Begin VB.TextBox txtFieldsVendor 
               DataField       =   "AP PO Remit City"
               Height          =   285
               Index           =   13
               Left            =   5760
               Locked          =   -1  'True
               TabIndex        =   27
               Top             =   1800
               Width           =   975
            End
            Begin VB.TextBox txtFieldsVendor 
               DataField       =   "AP PO Remit Address 2"
               Height          =   285
               Index           =   12
               Left            =   5760
               Locked          =   -1  'True
               TabIndex        =   26
               Top             =   1440
               Width           =   3255
            End
            Begin VB.Label label 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000001&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Remitt Address"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FFFF&
               Height          =   375
               Left            =   4800
               TabIndex        =   123
               Top             =   240
               Width           =   3300
            End
            Begin VB.Label Label1 
               Caption         =   "Document Type:"
               Height          =   255
               Index           =   4
               Left            =   9600
               TabIndex        =   104
               Top             =   960
               Width           =   1335
            End
            Begin VB.Label Label1 
               Caption         =   "Document #:"
               Height          =   255
               Index           =   3
               Left            =   9600
               TabIndex        =   103
               Top             =   1560
               Width           =   1335
            End
            Begin VB.Label lblLabels 
               Caption         =   "Telephone:"
               Height          =   255
               Index           =   17
               Left            =   4800
               TabIndex        =   102
               Top             =   2520
               Width           =   975
            End
            Begin VB.Label lblLabels 
               Caption         =   "Facsimile:"
               Height          =   255
               Index           =   16
               Left            =   7080
               TabIndex        =   101
               Top             =   2520
               Width           =   735
            End
            Begin VB.Label lblLabels 
               Caption         =   "Telephone:"
               Height          =   255
               Index           =   15
               Left            =   240
               TabIndex        =   100
               Top             =   2520
               Width           =   975
            End
            Begin VB.Label lblLabels 
               Caption         =   "Facsimile:"
               Height          =   255
               Index           =   2
               Left            =   2520
               TabIndex        =   99
               Top             =   2520
               Width           =   735
            End
            Begin VB.Label lblLabels 
               Caption         =   "  Zip:"
               Height          =   255
               Index           =   19
               Left            =   7800
               TabIndex        =   98
               Top             =   1800
               Width           =   495
            End
            Begin VB.Label lblLabels 
               Caption         =   "  State:"
               Height          =   255
               Index           =   11
               Left            =   6720
               TabIndex        =   97
               Top             =   1800
               Width           =   615
            End
            Begin VB.Label lblLabels 
               Caption         =   "City:"
               Height          =   255
               Index           =   7
               Left            =   4800
               TabIndex        =   96
               Top             =   1800
               Width           =   975
            End
            Begin VB.Label lblLabels 
               Caption         =   "  Zip:"
               Height          =   255
               Index           =   9
               Left            =   3240
               TabIndex        =   95
               Top             =   1800
               Width           =   495
            End
            Begin VB.Label lblLabels 
               Caption         =   "  State:"
               Height          =   255
               Index           =   5
               Left            =   2160
               TabIndex        =   94
               Top             =   1800
               Width           =   615
            End
            Begin VB.Label lblLabels 
               Caption         =   "Address:"
               Height          =   255
               Index           =   8
               Left            =   240
               TabIndex        =   93
               Top             =   1080
               Width           =   975
            End
            Begin VB.Label lblLabels 
               Caption         =   "City:"
               Height          =   255
               Index           =   10
               Left            =   240
               TabIndex        =   92
               Top             =   1800
               Width           =   975
            End
            Begin VB.Label lblLabels 
               Caption         =   "Country:"
               Height          =   255
               Index           =   13
               Left            =   240
               TabIndex        =   91
               Top             =   2160
               Width           =   975
            End
            Begin VB.Label lblLabels 
               Caption         =   "Name:"
               Height          =   255
               Index           =   6
               Left            =   240
               TabIndex        =   90
               Top             =   720
               Width           =   975
            End
            Begin VB.Label lblLabels 
               Caption         =   "Name:"
               Height          =   255
               Index           =   0
               Left            =   4800
               TabIndex        =   89
               Top             =   720
               Width           =   975
            End
            Begin VB.Label lblLabels 
               Caption         =   "Country:"
               Height          =   255
               Index           =   29
               Left            =   4800
               TabIndex        =   88
               Top             =   2160
               Width           =   975
            End
            Begin VB.Label lblLabels 
               Caption         =   "Address:"
               Height          =   255
               Index           =   21
               Left            =   4800
               TabIndex        =   87
               Top             =   1080
               Width           =   975
            End
            Begin VB.Label lblfields 
               Caption         =   "Customer ID:"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   86
               Top             =   360
               Width           =   1035
            End
         End
         Begin VB.Frame frThird 
            Height          =   855
            Left            =   0
            TabIndex        =   74
            Top             =   4800
            Width           =   11535
            Begin VB.CheckBox chkrecurr 
               Caption         =   "Recurring Entry"
               DataField       =   "AP PO Recurring YN"
               Height          =   195
               Left            =   7320
               TabIndex        =   126
               Top             =   480
               Width           =   1455
            End
            Begin VB.CheckBox ChkCompCheck 
               Caption         =   "Manual Check"
               DataField       =   "AP PO Computer Check"
               Height          =   195
               Left            =   9600
               TabIndex        =   125
               Top             =   480
               Width           =   1575
            End
            Begin VB.ComboBox cbPurchase 
               DataField       =   "AP PO Payment Method"
               Height          =   315
               Index           =   16
               ItemData        =   "frm_AP_RMA_Entry.frx":23F3
               Left            =   2520
               List            =   "frm_AP_RMA_Entry.frx":23F5
               TabIndex        =   10
               Text            =   "cbPurchase"
               Top             =   440
               Width           =   1335
            End
            Begin VB.CommandButton cmdUpdatedua 
               Height          =   300
               Index           =   16
               Left            =   3840
               Picture         =   "frm_AP_RMA_Entry.frx":23F7
               Style           =   1  'Graphical
               TabIndex        =   11
               ToolTipText     =   "Refresh Payment Methods"
               Top             =   440
               Width           =   375
            End
            Begin VB.CommandButton cmdUpdatedua 
               Height          =   280
               Index           =   5
               Left            =   1560
               Picture         =   "frm_AP_RMA_Entry.frx":2701
               Style           =   1  'Graphical
               TabIndex        =   9
               ToolTipText     =   "Refresh Payment Terms"
               Top             =   440
               Width           =   375
            End
            Begin VB.CommandButton cmdUpdatedua 
               Height          =   300
               Index           =   2
               Left            =   6360
               Picture         =   "frm_AP_RMA_Entry.frx":2A0B
               Style           =   1  'Graphical
               TabIndex        =   13
               ToolTipText     =   "Refresh  Ship Via"
               Top             =   440
               Width           =   375
            End
            Begin VB.ComboBox cbPurchase 
               DataField       =   "AP PO Payment Terms"
               Height          =   315
               Index           =   5
               ItemData        =   "frm_AP_RMA_Entry.frx":2D15
               Left            =   120
               List            =   "frm_AP_RMA_Entry.frx":2D17
               TabIndex        =   8
               Text            =   "cbPurchase"
               Top             =   440
               Width           =   1455
            End
            Begin VB.ComboBox cbPurchase 
               DataField       =   "AP PO Ship Method"
               Height          =   315
               Index           =   2
               Left            =   4800
               TabIndex        =   12
               Text            =   "cbPurchase"
               Top             =   440
               Width           =   1575
            End
            Begin VB.Label lblfields 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Payment Methods "
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   16
               Left            =   2520
               TabIndex        =   77
               Top             =   195
               Width           =   1695
            End
            Begin VB.Label lblfields 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Ship Via"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   2
               Left            =   4800
               TabIndex        =   76
               Top             =   195
               Width           =   1935
            End
            Begin VB.Label lblfields 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Payment Terms"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   5
               Left            =   120
               TabIndex        =   75
               Top             =   200
               Width           =   1815
            End
         End
         Begin VB.Frame frSecond 
            Height          =   855
            Left            =   0
            TabIndex        =   105
            Top             =   3000
            Width           =   11535
            Begin VB.TextBox txtfields 
               DataField       =   "AP PO Check Date"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "MM/dd/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   3
               EndProperty
               Enabled         =   0   'False
               Height          =   285
               Index           =   7
               Left            =   9600
               Locked          =   -1  'True
               TabIndex        =   117
               Top             =   480
               Width           =   1455
            End
            Begin VB.CommandButton cmdDate 
               Enabled         =   0   'False
               Height          =   285
               Index           =   0
               Left            =   11040
               Picture         =   "frm_AP_RMA_Entry.frx":2D19
               Style           =   1  'Graphical
               TabIndex        =   116
               ToolTipText     =   "Get Order Date"
               Top             =   480
               Width           =   375
            End
            Begin VB.CommandButton cmdbankAccount 
               Height          =   285
               Left            =   6360
               Picture         =   "frm_AP_RMA_Entry.frx":3023
               Style           =   1  'Graphical
               TabIndex        =   7
               ToolTipText     =   "Get Bank Account"
               Top             =   480
               Width           =   375
            End
            Begin VB.TextBox txtfields 
               DataField       =   "AP PO Check Acct ID"
               Height          =   285
               Index           =   35
               Left            =   4800
               TabIndex        =   6
               Top             =   480
               Width           =   1575
            End
            Begin VB.TextBox txtfields 
               Alignment       =   2  'Center
               DataField       =   "AP PO Check Number"
               Enabled         =   0   'False
               Height          =   285
               Index           =   34
               Left            =   7320
               TabIndex        =   42
               Top             =   480
               Width           =   1695
            End
            Begin VB.CommandButton cmdDate 
               Height          =   285
               Index           =   1
               Left            =   3840
               Picture         =   "frm_AP_RMA_Entry.frx":332D
               Style           =   1  'Graphical
               TabIndex        =   5
               ToolTipText     =   "Get Due Date"
               Top             =   480
               Width           =   375
            End
            Begin VB.TextBox txtfields 
               DataField       =   "AP PO Due Date"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "MM/dd/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   3
               EndProperty
               Height          =   285
               Index           =   6
               Left            =   2520
               Locked          =   -1  'True
               TabIndex        =   37
               Top             =   480
               Width           =   1335
            End
            Begin VB.TextBox txtfields 
               DataField       =   "AP PO Date Requested"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "MM/dd/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   3
               EndProperty
               Height          =   285
               Index           =   20
               Left            =   120
               Locked          =   -1  'True
               TabIndex        =   36
               Top             =   480
               Width           =   1455
            End
            Begin VB.CommandButton cmdDate 
               Height          =   285
               Index           =   20
               Left            =   1560
               Picture         =   "frm_AP_RMA_Entry.frx":3637
               Style           =   1  'Graphical
               TabIndex        =   4
               ToolTipText     =   "Get Ship Date"
               Top             =   480
               Width           =   375
            End
            Begin VB.Label lblfields 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Refund Check Date"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   7
               Left            =   9600
               TabIndex        =   118
               Top             =   240
               Width           =   1815
            End
            Begin VB.Label lblfields 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Bank Account"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   35
               Left            =   4800
               TabIndex        =   109
               Top             =   240
               Width           =   1935
            End
            Begin VB.Label lblfields 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Refund Check No."
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   34
               Left            =   7320
               TabIndex        =   108
               Top             =   240
               Width           =   1695
            End
            Begin VB.Label lblfields 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Due Date"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   6
               Left            =   2520
               TabIndex        =   107
               Top             =   240
               Width           =   1695
            End
            Begin VB.Label lblfields 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "RMA Date"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   20
               Left            =   120
               TabIndex        =   106
               Top             =   240
               Width           =   1815
            End
         End
         Begin MSDataGridLib.DataGrid grdDataGrid 
            Height          =   825
            Left            =   0
            TabIndex        =   124
            Top             =   3960
            Width           =   11520
            _ExtentX        =   20320
            _ExtentY        =   1455
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   11594218
            HeadLines       =   1
            RowHeight       =   15
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   11
            BeginProperty Column00 
               DataField       =   "AP POD Item ID"
               Caption         =   "Item ID"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "AP POD Description"
               Caption         =   "Item Desc."
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column02 
               DataField       =   "AP POD Qty"
               Caption         =   "Order Qty"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column03 
               DataField       =   "AP POD Total Qty Received"
               Caption         =   "Rec. Qty"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column04 
               DataField       =   "AP POD Units"
               Caption         =   "Units"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   """$""#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column05 
               DataField       =   "AP POD Unit Cost"
               Caption         =   "Unit Cost"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   """$""#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column06 
               DataField       =   "AP POD Date Received"
               Caption         =   "Date Rec."
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "M/d/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column07 
               DataField       =   "AP POD Vendor Item ID"
               Caption         =   "Vendor Item Id"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   "0.000E+00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column08 
               DataField       =   "AP POD Item Total"
               Caption         =   "Total"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   """$""#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column09 
               DataField       =   "AP POD Posting Account"
               Caption         =   "GL Account"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column10 
               DataField       =   "AP POD Project ID"
               Caption         =   "Project"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
                  Button          =   -1  'True
                  ColumnWidth     =   1094.74
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1709.858
               EndProperty
               BeginProperty Column02 
                  Alignment       =   2
                  ColumnWidth     =   599.811
               EndProperty
               BeginProperty Column03 
                  Alignment       =   2
                  ColumnWidth     =   599.811
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   524.976
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   1005.165
               EndProperty
               BeginProperty Column06 
                  Button          =   -1  'True
                  ColumnWidth     =   1005.165
               EndProperty
               BeginProperty Column07 
                  ColumnWidth     =   794.835
               EndProperty
               BeginProperty Column08 
                  ColumnWidth     =   1200.189
               EndProperty
               BeginProperty Column09 
                  Button          =   -1  'True
                  ColumnWidth     =   1305.071
               EndProperty
               BeginProperty Column10 
                  Button          =   -1  'True
                  ColumnWidth     =   1200.189
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame frAdvance 
         Height          =   8355
         Left            =   11640
         TabIndex        =   58
         Top             =   0
         Width           =   3375
         Begin VB.PictureBox picStatBox 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   120
            ScaleHeight     =   300
            ScaleWidth      =   3135
            TabIndex        =   128
            Top             =   7980
            Width           =   3135
            Begin VB.CommandButton cmdNext 
               Height          =   300
               Left            =   2300
               Picture         =   "frm_AP_RMA_Entry.frx":3941
               Style           =   1  'Graphical
               TabIndex        =   132
               ToolTipText     =   "Move Forward"
               Top             =   0
               UseMaskColor    =   -1  'True
               Width           =   345
            End
            Begin VB.CommandButton cmdLast 
               Height          =   300
               Left            =   2640
               Picture         =   "frm_AP_RMA_Entry.frx":3C83
               Style           =   1  'Graphical
               TabIndex        =   131
               ToolTipText     =   "End Of Record"
               Top             =   0
               UseMaskColor    =   -1  'True
               Width           =   345
            End
            Begin VB.CommandButton cmdPrevious 
               Height          =   300
               Left            =   460
               Picture         =   "frm_AP_RMA_Entry.frx":3FC5
               Style           =   1  'Graphical
               TabIndex        =   130
               ToolTipText     =   "Move Previous"
               Top             =   0
               UseMaskColor    =   -1  'True
               Width           =   345
            End
            Begin VB.CommandButton cmdFirst 
               Height          =   300
               Left            =   120
               Picture         =   "frm_AP_RMA_Entry.frx":4307
               Style           =   1  'Graphical
               TabIndex        =   129
               ToolTipText     =   "Beginning of the Data"
               Top             =   0
               UseMaskColor    =   -1  'True
               Width           =   345
            End
            Begin VB.Label lblStatus 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Record"
               Height          =   285
               Left            =   810
               TabIndex        =   133
               Top             =   0
               Width           =   1515
            End
         End
         Begin VB.Frame Frame3 
            Height          =   1455
            Left            =   120
            TabIndex        =   111
            Top             =   120
            Width           =   3135
            Begin VB.CommandButton cmdUnPostedDoc 
               Caption         =   "Open"
               Height          =   405
               Left            =   840
               Picture         =   "frm_AP_RMA_Entry.frx":4649
               TabIndex        =   56
               ToolTipText     =   "Load Unposted Document"
               Top             =   900
               Width           =   735
            End
            Begin VB.CommandButton cmdPostDoc 
               Caption         =   "Posted"
               Height          =   405
               Left            =   120
               Picture         =   "frm_AP_RMA_Entry.frx":4AF8
               TabIndex        =   55
               ToolTipText     =   "Load Posted Document"
               Top             =   900
               Width           =   735
            End
            Begin VB.CommandButton Command1 
               Height          =   495
               Left            =   2400
               Picture         =   "frm_AP_RMA_Entry.frx":5091
               Style           =   1  'Graphical
               TabIndex        =   54
               ToolTipText     =   "Search All Record"
               Top             =   240
               Width           =   615
            End
            Begin VB.TextBox txtfields 
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "MM/dd/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   3
               EndProperty
               Height          =   285
               Index           =   0
               Left            =   120
               TabIndex        =   52
               Top             =   480
               Width           =   1455
            End
            Begin VB.CommandButton cmdSearch 
               Height          =   495
               Left            =   1680
               Picture         =   "frm_AP_RMA_Entry.frx":539B
               Style           =   1  'Graphical
               TabIndex        =   53
               ToolTipText     =   "Search Current Record"
               Top             =   240
               Width           =   615
            End
            Begin VB.Label lblfields 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Doc No"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   18
               Left            =   120
               TabIndex        =   112
               Top             =   240
               Width           =   1455
            End
         End
         Begin VB.Frame frButton 
            Height          =   1815
            Left            =   120
            TabIndex        =   59
            Top             =   6120
            Width           =   3135
            Begin VB.CommandButton cmdCreateInvoice 
               Caption         =   "P&ost"
               Height          =   780
               Left            =   2040
               Picture         =   "frm_AP_RMA_Entry.frx":56A5
               Style           =   1  'Graphical
               TabIndex        =   44
               ToolTipText     =   "Create Invoice"
               Top             =   240
               Width           =   975
            End
            Begin VB.CommandButton cmdApprove 
               Caption         =   "Appro&ved"
               Height          =   780
               Left            =   1080
               Picture         =   "frm_AP_RMA_Entry.frx":59AF
               Style           =   1  'Graphical
               TabIndex        =   122
               ToolTipText     =   "Approved Current Document"
               Top             =   240
               Width           =   975
            End
            Begin VB.CommandButton cmdPrint 
               Caption         =   "&Print"
               Height          =   780
               Left            =   120
               Picture         =   "frm_AP_RMA_Entry.frx":5CB9
               Style           =   1  'Graphical
               TabIndex        =   43
               ToolTipText     =   "Print Sales Order"
               Top             =   240
               Width           =   975
            End
            Begin VB.CommandButton cmdCancel 
               Caption         =   "Canc&el"
               Height          =   300
               Left            =   2040
               TabIndex        =   50
               ToolTipText     =   "Cancell Current Process"
               Top             =   1440
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.CommandButton cmdClose 
               Caption         =   "&Close"
               Height          =   300
               Left            =   1080
               TabIndex        =   49
               ToolTipText     =   "Close Order Transaction"
               Top             =   1440
               Width           =   975
            End
            Begin VB.CommandButton cmdRefresh 
               Caption         =   "&Refresh"
               Height          =   300
               Left            =   120
               TabIndex        =   48
               ToolTipText     =   "Refresh All"
               Top             =   1440
               Width           =   975
            End
            Begin VB.CommandButton cmdDelete 
               Caption         =   "&Delete"
               Height          =   300
               Left            =   2040
               TabIndex        =   47
               ToolTipText     =   "Delete Current Record"
               Top             =   1080
               Width           =   975
            End
            Begin VB.CommandButton cmdUpdate 
               Caption         =   "&Update"
               Height          =   300
               Left            =   1080
               TabIndex        =   46
               ToolTipText     =   "Update Current Record"
               Top             =   1080
               Width           =   975
            End
            Begin VB.CommandButton cmdAdd 
               Caption         =   "&Add"
               Height          =   300
               Left            =   120
               TabIndex        =   45
               ToolTipText     =   "Adding New Data"
               Top             =   1080
               Width           =   975
            End
         End
         Begin VB.PictureBox PcData 
            BorderStyle     =   0  'None
            Height          =   2535
            Left            =   120
            ScaleHeight     =   2535
            ScaleWidth      =   3135
            TabIndex        =   60
            Top             =   3600
            Width           =   3135
            Begin VB.CommandButton cmdCreditLimit 
               Height          =   285
               Left            =   1200
               Picture         =   "frm_AP_RMA_Entry.frx":5FC3
               Style           =   1  'Graphical
               TabIndex        =   120
               Top             =   2160
               Width           =   375
            End
            Begin VB.TextBox txtfields 
               DataField       =   "AP PO Subtotal"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   """$""#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   2
               EndProperty
               Height          =   285
               Index           =   24
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   65
               Top             =   0
               Width           =   1935
            End
            Begin VB.TextBox txtfields 
               DataField       =   "AP PO Discount Amt"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   """$""#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   2
               EndProperty
               Height          =   285
               Index           =   25
               Left            =   1680
               Locked          =   -1  'True
               TabIndex        =   64
               Top             =   360
               Width           =   1335
            End
            Begin VB.TextBox txtfields 
               DataField       =   "AP PO Misc Charges"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   """$""#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   2
               EndProperty
               Height          =   285
               Index           =   26
               Left            =   1680
               Locked          =   -1  'True
               TabIndex        =   63
               Top             =   720
               Width           =   1335
            End
            Begin VB.TextBox txtfields 
               DataField       =   "AP PO Shipping"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   """$""#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   2
               EndProperty
               Height          =   285
               Index           =   27
               Left            =   1080
               TabIndex        =   39
               Top             =   1080
               Width           =   1935
            End
            Begin VB.TextBox txtfields 
               DataField       =   "AP PO Discount Percent"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               Height          =   285
               Index           =   28
               Left            =   1080
               TabIndex        =   38
               Top             =   360
               Width           =   495
            End
            Begin VB.TextBox txtfields 
               DataField       =   "AP PO Misc Percent"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               Height          =   285
               Index           =   29
               Left            =   1080
               TabIndex        =   40
               Top             =   720
               Width           =   495
            End
            Begin VB.TextBox txtfields 
               DataField       =   "AP PO Total Amount"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   """$""#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   2
               EndProperty
               Height          =   285
               Index           =   30
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   62
               Top             =   1440
               Width           =   1935
            End
            Begin VB.TextBox txtfields 
               DataField       =   "AP PO Amount Paid"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   """$""#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   2
               EndProperty
               Height          =   285
               Index           =   36
               Left            =   0
               TabIndex        =   41
               Top             =   2160
               Width           =   1215
            End
            Begin VB.TextBox txtfields 
               DataField       =   "AP PO Balance Due"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   """$""#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   2
               EndProperty
               Height          =   285
               Index           =   33
               Left            =   1800
               Locked          =   -1  'True
               TabIndex        =   61
               Top             =   2160
               Width           =   1335
            End
            Begin VB.TextBox txtfields 
               BeginProperty DataFormat 
                  Type            =   5
                  Format          =   "0.00"
                  HaveTrueFalseNull=   1
                  TrueValue       =   "Yes"
                  FalseValue      =   "No"
                  NullValue       =   "N/A"
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   7
               EndProperty
               Height          =   285
               Index           =   23
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   127
               Top             =   1080
               Width           =   495
            End
            Begin VB.Label lblLabels 
               Caption         =   "Shipping:"
               Height          =   255
               Index           =   3
               Left            =   0
               TabIndex        =   121
               Top             =   1080
               Width           =   1095
            End
            Begin VB.Label lblLabels 
               Caption         =   "Sub Total:"
               Height          =   255
               Index           =   24
               Left            =   0
               TabIndex        =   71
               Top             =   0
               Width           =   1095
            End
            Begin VB.Label lblLabels 
               Caption         =   "Discount:"
               Height          =   255
               Index           =   22
               Left            =   0
               TabIndex        =   70
               Top             =   360
               Width           =   1095
            End
            Begin VB.Label lblLabels 
               Caption         =   "Restocking:"
               Height          =   255
               Index           =   23
               Left            =   0
               TabIndex        =   69
               Top             =   720
               Width           =   1095
            End
            Begin VB.Label lblLabels 
               Caption         =   "Total Amount:"
               Height          =   255
               Index           =   25
               Left            =   0
               TabIndex        =   68
               Top             =   1440
               Width           =   1095
            End
            Begin VB.Label lblfields 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Amount Paid"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   36
               Left            =   0
               TabIndex        =   67
               Top             =   1920
               Width           =   1575
            End
            Begin VB.Label lblfields 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Balance Due"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   33
               Left            =   1800
               TabIndex        =   66
               Top             =   1920
               Width           =   1335
            End
         End
         Begin VB.TextBox txtfields 
            DataField       =   "AP PO Description"
            Height          =   1515
            Index           =   32
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   14
            Top             =   1920
            Width           =   3135
         End
         Begin VB.Label lblfields 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Description"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   31
            Left            =   120
            TabIndex        =   72
            Top             =   1680
            Width           =   3135
         End
      End
   End
   Begin VB.Label lblTop 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "RMA"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   110
      Top             =   120
      Width           =   9225
   End
End
Attribute VB_Name = "frm_AP_RMA_Entry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents ADOprimaryrs As ADODB.Recordset
Attribute ADOprimaryrs.VB_VarHelpID = -1
'Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbAddNewFlag As Boolean

Dim db As ADODB.Connection
Dim NewLoad As Boolean
Dim DocType As String
Dim RSstatement As String
Dim VendID0 As String
Dim ShipID0 As String
Dim BankAcct35 As String

Public grdOnAddNew As Boolean

Private Function Datavalidate() As Boolean

  If txtfields(36) > 0 Then
    If txtfields(35) = "" Then
      MsgBox "Please enter a bank account!", , "Error"
      Datavalidate = False
      Exit Function
    End If

    'If cbpurchase(16).Text = "Cash" Then
    'Else
    '  MsgBox "Not a valid back account! This is cash payments", , "Error"
    '  DataValidate = False
      'Exit Function
    'End If
  End If
  
  If CountRecord("[AP POD Item Id]", "[AP Purchase Detail]", db, "[AP POD Document No] = " & txtfields(2)) <= 0 Then
    MsgBox "Must have at least one inventory item!", , "Error"
    Datavalidate = False
    Exit Function
  End If

  'Check for balance due < 0 and check number
  If txtfields(36) > 0 Then
    If txtfields(33) < 0 Then
      MsgBox "Amount paid cannot exceed invoice total!", , "Error"
      Datavalidate = False
      Exit Function
    End If
    If Trim(txtfields(34)) = "" Then
      MsgBox "You must enter a check number!", , "Error"
      Datavalidate = False
      Exit Function
    End If
  End If
  If txtfields(24) = "$0.00" Then
    MsgBox "Cannot approve a transaction with zero amount, your request is cancel", vbInformation, "Error"
      Datavalidate = False
      Exit Function
  End If

  Datavalidate = True
  
  Exit Function

End Function


Private Sub adoPrimaryRS_RecordChangeComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  Select Case adReason
  Case adRsnAddNew
  Case adRsnClose
  Case adRsnDelete
  Case adRsnFirstChange
  Case adRsnMove
  Case adRsnRequery
  Case adRsnResynch
  Case adRsnUndoAddNew
  Case adRsnUndoDelete
  Case adRsnUndoUpdate
  Case adRsnUpdate
  End Select

End Sub

'Private Sub cbPurchase_KeyPress(Index As Integer, KeyAscii As Integer)
'Dim keyResponse As Boolean
'    keyResponse = CtrlValidate(KeyAscii, "")
'    If keyResponse = True Then
'    Else
'       KeyAscii = 0
'    End If
'End Sub

Private Sub loadCombo(LoadTotype As String)
'fill the data into the combobox
Select Case LoadTotype
    'Case "kosong"
    '    ComboInit cbhelp, lblhelp, "SELECT [Form Name] FROM [Help Text]"
'    Case "satu"
'        ComboInit cbPurchase(1), lblfields(1), "SELECT [SYS TAXGRP ID] FROM [SYS Tax Group]", db
'        CheckCombo cbPurchase(1), "[SYS TAXGRP ID]", "[SYS Tax Group]", db, True
'        CalcTotals
    Case "dua"
        ComboInit cbPurchase(2), lblfields(2), "SELECT [LIST SHIP Method] FROM [LIST Shipping Methods]", db
        CheckCombo cbPurchase(2), "[LIST SHIP Method]", "[LIST Shipping Methods]", db, True
        CalcTotals
    Case "tiga"
        ComboInit cbPurchase(3), lblfields(3), "SELECT [LIST PAY Method] FROM [LIST Payment Methods]", db
    Case "lima"
        ComboInit cbPurchase(5), lblfields(5), "SELECT [LIST PAY Description] FROM [LIST Payment Terms]", db
'    Case "limabelas"
'        ComboInit cbPurchase(15), lblfields(15), "SELECT [RECURR TYPE] FROM [RECUR_TYPE]", db
    Case "enambelas"
        ComboInit cbPurchase(16), lblfields(16), "SELECT [LIST PAY Method] FROM [LIST Payment Methods]", db
    Case "semua"
        'ComboInit cbhelp, lblhelp, "SELECT [Form Name] FROM [Help Text]"
'        ComboInit cbPurchase(1), lblfields(1), "SELECT [SYS TAXGRP ID] FROM [SYS Tax Group]", db
        ComboInit cbPurchase(2), lblfields(2), "SELECT [LIST SHIP Method] FROM [LIST Shipping Methods]", db
        ComboInit cbPurchase(5), lblfields(5), "SELECT [LIST PAY Description] FROM [LIST Payment Terms]", db
'        ComboInit cbPurchase(15), lblfields(15), "SELECT [RECURR TYPE] FROM [RECUR_TYPE]", db
        ComboInit cbPurchase(16), lblfields(16), "SELECT [LIST PAY Method] FROM [LIST Payment Methods]", db
End Select
End Sub

Private Sub cbPurchase_LostFocus(Index As Integer)
On Error GoTo exit_EditMode
If ADOprimaryrs.EditMode = adEditInProgress Then
    If IsNull(ADOprimaryrs![AP PO Status]) Or ADOprimaryrs![AP PO Status] <> "Open" Then
        ValidatePower txtfields(12).Text, "Edit", DocType, db
        ADOprimaryrs![AP PO Status] = "Open"
        ADOprimaryrs.Update
        cmdApprove.Picture = fMainForm.imlIcons.ListImages("Locked").Picture
    End If
End If
exit_EditMode:

ShowStatus True
Select Case Index
'Case 1
'   CheckCombo cbPurchase(Index), "[SYS TAXGRP ID]", "[SYS Tax Group]", db, True
'   CalcTotals
Case 2
   CheckCombo cbPurchase(Index), "[LIST SHIP Method]", "[LIST Shipping Methods]", db, True
   CalcTotals
Case 5
   CheckCombo cbPurchase(Index), "[LIST PAY Description]", "[LIST Payment Terms]", db, True
   SetDueDate
   ComboInit cbPurchase(16), lblfields(16), "SELECT [LIST PAY Method] FROM [LIST Payment Methods] WHERE [Payment Terms]='" & cbPurchase(5).Text & "'", db
'Case 15
'   CheckCombo cbPurchase(Index), "[RECURR TYPE]", "[RECUR_TYPE]", db, True
Case 16
   CheckCombo cbPurchase(Index), "[LIST PAY Method]", "[LIST Payment Methods]", db, True
End Select
ShowStatus False
End Sub

Private Sub cmdUpdatedua_Select(Index)
    Select Case Index
'    Case 1
'        LoadCombo "satu"
    Case 2
        loadCombo "dua"
    Case 3
        loadCombo "tiga"
    Case 5
        loadCombo "lima"
'    Case 15
'        LoadCombo "limabelas"
    Case 16
        loadCombo "enambelas"
    End Select
End Sub

Private Sub ChkCompCheck_Click()
    If ChkCompCheck.Value = 1 Then
        ChkCompCheck.Caption = "Computer Check" 'AP PO Computer Check
        txtfields(34).Locked = True
    Else
        ChkCompCheck.Caption = "Manual Check"
        txtfields(34).Locked = False
    End If
    If txtfields(34).Enabled = True Then GetTextColor Me

If NewLoad = True Then Exit Sub
On Error GoTo exit_EditMode
If ADOprimaryrs.EditMode = adEditInProgress Then
    If IsNull(ADOprimaryrs![AP PO Status]) Or ADOprimaryrs![AP PO Status] <> "Open" Then
        ValidatePower txtfields(12).Text, "Edit", DocType, db
        ADOprimaryrs![AP PO Status] = "Open"
        ADOprimaryrs.Update
        cmdApprove.Picture = fMainForm.imlIcons.ListImages("Locked").Picture
    End If
End If
exit_EditMode:
End Sub

Private Sub chkrecurr_Click()
If NewLoad = True Then Exit Sub
On Error GoTo exit_EditMode
If ADOprimaryrs.EditMode = adEditInProgress Then
    If IsNull(ADOprimaryrs![AP PO Status]) Or ADOprimaryrs![AP PO Status] <> "Open" Then
        ValidatePower txtfields(12).Text, "Edit", DocType, db
        ADOprimaryrs![AP PO Status] = "Open"
        ADOprimaryrs.Update
        cmdApprove.Picture = fMainForm.imlIcons.ListImages("Locked").Picture
    End If
End If
exit_EditMode:
End Sub

Private Sub cmdApprove_Click()
Dim Approve As Boolean
If ADOprimaryrs.EditMode = adEditAdd Then Exit Sub
If Datavalidate = False Then Exit Sub
If Not CheckEmpty Then Exit Sub

If ADOprimaryrs![AP PO Status] = "Open" Or IsNull(ADOprimaryrs![AP PO Status]) Then
    Approve = ValidatePower(txtfields(12).Text, "Approve", DocType, db)
    If Approve = True Then
        ADOprimaryrs![AP PO Status] = "Ready"
        ADOprimaryrs.Update
    End If
     MsgBox "This transaction is approved"
Else
     MsgBox "This transaction is already approved"
End If
     cmdApprove.Picture = fMainForm.imlIcons.ListImages("Approved").Picture
End Sub

Private Sub cmdbankAccount_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    cmdBankAccount_Select
    txtfields(35).SetFocus
Else
    'Me.PopupMenu fMainForm.mnuAccount
End If
End Sub

Private Sub cmdBankAccount_Select()
    Dim SQLstatement As String
    Dim No As Integer
    Dim ghead As String
    Dim fhead As String

    No = 1230
    SQLstatement = "select [BANK ACCT ID], [BANK ACCT Name],[BANK ACCT Next Check No]" & _
                    "from [BANK Accounts]"
    ghead = "Account Description"
    fhead = "Account No//Description//Check No"
    
    AllLookup.GetWhichTable No, SQLstatement, ghead, fhead, db
    'AllLookup.Show vbModal
End Sub

Private Sub cmdCreditLimit_Click()
'Dim CurrRequest As Currency
'If mbAddNewFlag Then
'    CurrRequest = 0
'Else
'    CurrRequest = CCur(txtfields(33).Text)
'End If
'If Trim(txtFieldsVendor(0)) <> "" Then
'    CheckCreditLimit CurrRequest, txtFieldsVendor(0), db, True
'End If
End Sub

Private Sub cmdLookupVEND_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    cmdLookupVEND_Select
    VendID0 = txtFieldsVendor(0).Text
    If Trim(txtFieldsVendor(0).Text) <> "" Then GetWEBMAILvendor txtFieldsVendor(0).Text, db, Me
Else
    'Me.PopupMenu mnuCustomer
End If

End Sub

Private Sub cmdLookupVEND_Select()
    If cmdLookupVEND.Visible = False Then
       AllLookup.ToWhichRecord ADOprimaryrs, DocType, "-//Order No//-//-//-//-//-//-//Cust. PO//-//-//-//-//Customer Name"
       'AllLookup.Show vbModal
       GetTextColor Me
    Else
       Dim i As Integer
   AllLookup.GetWhichTable 1300, "Select [AP VEN ID],[AP VEN Name],[AP VEN Address 1]," & _
   "[AP VEN Address 2],[AP VEN City],[AP VEN State],[AP VEN Postal],[AP VEN Country]," & _
   "[AP VEN Phone],[AP VEN Fax],[AP VEN Remit Name],[AP VEN Remit Address 1],[AP VEN Remit Address 2]," & _
   "[AP VEN Remit City],[AP VEN Remit State],[AP VEN Remit Country],[AP VEN Remit Country],[AP VEN Remit Phone],[AP VEN Remit Fax] From " & _
   "[AP Vendor] ", "Vendor Particular", _
   "Vendor ID//Vendor Name//Address 1//Address 2//City//State//Postal//Country", db
       'AllLookup.Show vbModal
        'If VendID0 <> txtFieldsVendor(0).Text And txtFieldsVendor(0).Text <> "" Then
            'CustomerData Me, db, txtFieldsVendor(0).Text, True
            'For i = 0 To 'txtFieldsShip.UBound
            '    'txtFieldsShip(i).Text = ""
            'Next
            'shipToID db, Me
            'cbPurchase_LostFocus 1
            'cbPurchase_LostFocus 5
        'End If
    End If
End Sub

Private Sub cmdPostDoc_Click()
    ClearDatasource
    RSstatement = "SHAPE {select * from [AP Purchase] WHERE [AP PO Document Type]='" & DocType & "' AND [AP PO Posted YN]=True} AS ParentCMD APPEND ({select * from [AP Purchase Detail] } AS ChildCMD RELATE [AP PO Document No] TO [AP POD Document No]) AS ChildCMD"
    OpenDB RSstatement
    picStatBox.Enabled = True
    lblStatus.BackColor = vbWhite
    ProcessDoneMusic "Done"
End Sub

Private Sub cmdPrint_Click()
    If ADOprimaryrs![AP PO Status] = "Open" Then
        MsgBox "This " & DocType & " has not been approved."
        Exit Sub
    End If
    
    Dim frm As New frm_prnPreview
    frm.Record txtfields(12).Text, Me.Name, DocType
    frm.Show
End Sub

Private Sub cmdSearch_Click()
    If mbAddNewFlag = True Then
        MsgBox "Can't search during adding new data"
        Exit Sub
    End If
If ADOprimaryrs Is Nothing Then
Else
    If ADOprimaryrs.RecordCount = 0 Then Exit Sub
    SearchRECORD ADOprimaryrs, grdDataGrid, txtfields(0).Text, lblfields(18).Caption, "AP PO Ext Document No", "AP PO Ext Document No"
    ProcessDoneMusic "Done"
End If
End Sub

Private Sub cmdUnPostedDoc_Click()
    ClearDatasource
    RSstatement = "SHAPE {select * from [AP Purchase] WHERE [AP PO Document Type]='" & DocType & "' AND [AP PO Posted YN]=False} AS ParentCMD APPEND ({select * from [AP Purchase Detail] } AS ChildCMD RELATE [AP PO Document No] TO [AP POD Document No]) AS ChildCMD"
    OpenDB RSstatement
    picStatBox.Enabled = True
    lblStatus.BackColor = vbWhite
    ProcessDoneMusic "Done"
End Sub

Private Sub cmdCreateInvoice_Click()

  'On Error GoTo cmdPost_Click_Error
    If ADOprimaryrs![AP PO Status] = "Open" Then
        MsgBox "This " & DocType & " has not been approved."
        Exit Sub
    End If
    If Datavalidate = False Then Exit Sub
    If Not CheckEmpty Then Exit Sub
  
  cmdUpdate_Click
  
  'Post this RMA to the general ledger
  Dim Success%
  Dim Response%
  Dim SQLstatement As String
  'Dim VendorID$
  'Dim BankID$
  'Dim rsAPPaymentHeader As ADODB.Recordset
  'Dim SetVoid%

  'SetVoid% = False
  
  If chkrecurr.Value = 1 And txtfields(36) > 0 Then
    MsgBox "Cannot post with amount paid on a recurring entry!" & vbCr & "Transaction NOT Posted.", vbInformation, "Error"
    Exit Sub
  End If
  
  'Set rsAPPaymentHeader = New ADODB.Recordset
  'rsAPPaymentHeader.Open "SELECT * FROM [AP Payment Header]", db, adOpenStatic, adLockOptimistic, adCmdText
  'rsAPPaymentHeader.Index = "BankKey"

  'Force record save
  'DoCmd.RunMacro "Save Record"

  ShowStatus True

  'Success% = CalcTotals()
  'If Success% = False Then
  '  MsgBox "An error occurred posting the transaction!", , "Error"
  '  DoCmd.Hourglass False
  '  Exit Sub
  'End If
  
  db.BeginTrans
    'Select Case Me![AP PO Document Type]
    'Case "PO"
    '  DoCmd.RunMacro "Save Record"
    '  DoCmd.Hourglass False
    '  DBEngine.Workspaces(0).CommitTrans
    '  Exit Sub
    'Case "Receiving"
    '  Success% = PostReceiving(CLng(txtFields(2)), True)
    'Case "Voucher"
    '  Success% = PostVoucher(CLng(Me![AP PO Document No]), True)
    'Case "Credit Memo"
    '  Success% = PostPOCreditMemo(CLng(Me![AP PO Document No]), True)
    'Case "RMA"
      Success% = PostRMA(CLng(txtfields(2)), True, db)
    'End Select
    
       If Success% = True Then
      'Do some check processing
       If CCur(txtfields(36).Text) > 0 Then
            If ChkCompCheck.Value = 1 Then
              'Print the check
              Response% = MsgBox("Make sure a check is in the printer and press OK.", vbOKCancel, "Printing Check")
                If Response% = vbCancel Then
                    db.RollbackTrans
                    MsgBox "Transaction process cancelled "
                Else
                  'CheckNo$ = txtfields(34)
                  'BankID$ = txtfields(35)
                  'VendorID$ = txtFieldsVendor(0)
                  PrintCheckLocal
                  Response% = MsgBox("Did the check print correctly?", vbYesNo, "Confirmation")
                    If Response% = vbNo Then
                        db.RollbackTrans
                        MsgBox "Transaction NOT Posted. Check Number " & txtfields(34) & " will be voided", vbInformation, "Information"
                        
                        Dim PayType As String
                        If cbPurchase(16).Text = "Cash" Or cbPurchase(3).Text = "Company Check" Then '<<<------
                          PayType = "Payment"
                        Else
                          PayType = "Charge"
                        End If
                        SQLstatement = "INSERT INTO [AP Payment Header]"
                        SQLstatement = SQLstatement & " ([AP PAY Type],[AP PAY Check No],[AP PAY Vendor No],[AP PAY Transaction Date],[AP PAY Amount],[AP PAY UnApplied Amount],[AP PAY Bank Account],[AP PAY Status],[AP PAY Posted YN],[AP PAY Void],[AP PAY Cleared])"
                        SQLstatement = SQLstatement & " VALUES ('" & PayType & "','" & txtfields(34) & "','" & txtFieldsVendor(0) & "',#" & txtfields(7) & "#," & txtfields(36) & ",0," & txtfields(35) & ",'Posted',True,True,False)"
                        db.Execute SQLstatement
                    Else
                        db.CommitTrans
                        MsgBox "Transaction Posted."
                        ADOprimaryrs![AP PO Posted YN] = True
                        cmdUpdate_Click  'update the data to the database
                    End If
                End If
            Else
                db.CommitTrans
                MsgBox "Transaction Posted. Please write a check Number " & txtfields(34)
                ADOprimaryrs![AP PO Posted YN] = True
                cmdUpdate_Click  'update the data to the database
            End If
       Else
                db.CommitTrans
                MsgBox "Transaction Posted."
                ADOprimaryrs![AP PO Posted YN] = True
                cmdUpdate_Click  'update the data to the database
       End If
    Else
        db.RollbackTrans
        MsgBox "An Error occured during process. Transaction NOT Posted."
    End If
    
  ButtEnabled False
  GetTextColor Me
  ShowStatus False
    
    'If Success% = True Then
      'Do some check processing
    '   If txtfields(36) > 0 Then
    '        If ChkCompCheck.Value = 1 Then
              'Print the check
    '          response% = MsgBox("Make sure a check is in the printer and press OK.", vbOKCancel, "Printing Check")
    '            If response% = vbCancel Then
    '              Success% = False
    '            Else
    '              CheckNo$ = txtfields(34)
    '              BankID$ = txtfields(35)
    '              VendorID$ = txtFieldsVendor(0)
    '              Call PrintCheckLocal
    '              response% = MsgBox("Did the check print correctly?", vbYesNo, "Confirmation")
    '              If response% = vbNo Then
                  'Mark the check void
    '              SetVoid% = True
    '              Success% = False
    '              End If
    '            End If
    '        End If
    '      End If
    '    End If

    '    If Success% = False Then
    '      db.RollbackTrans
    '      MsgBox "Transaction NOT Posted."
    '      If SetVoid% = True Then
    '        rsAPPaymentHeader.AddNew
    '          If cbPurchase(3).Text = "Cash" Or cbPurchase(3).Text = "Company Check" Then '<<<------
    '            rsAPPaymentHeader("AP PAY Type") = "Payment"
    '          Else
    '            rsAPPaymentHeader("AP PAY Type") = "Charge"
    '          End If
    '          rsAPPaymentHeader("AP PAY Check No") = txtfields(34)
    '          rsAPPaymentHeader("AP PAY Vendor No") = txtFieldsVendor
    '          rsAPPaymentHeader("AP PAY Transaction Date") = txtfields(7)
    '          rsAPPaymentHeader("AP PAY Amount") = txtfields(36)
    '          rsAPPaymentHeader("AP PAY UnApplied Amount") = 0
    '          rsAPPaymentHeader("AP PAY Bank Account") = txtfields(35)
    '          rsAPPaymentHeader("AP PAY Status") = "Posted"
    '          rsAPPaymentHeader("AP PAY Posted YN") = True
    '          rsAPPaymentHeader("AP PAY Void") = True
    '          rsAPPaymentHeader("AP PAY Cleared") = False
    '        rsAPPaymentHeader.Update
    '      End If
    ' Else
    '  db.CommitTrans
    '  MsgBox "Transaction Posted."
    '  ADOprimaryrs![AP PO Posted YN] = True
    '  cmdUpdate_Click  'update the data to the database
      'DoCmd.GoToRecord A_FORM, "Purchase Transactions", A_NEWREC
      'DoCmd.GoToControl "AP PO Vendor ID"
    'End If
  'ButtEnabled False
  'ShowStatus False
  
  Exit Sub
  
cmdPost_Click_Error:
  Call ErrorLog("Purchase Transactions", "cmdPost_Click", Now, Err.Number, Err.Description, True, db)
  Resume Next

End Sub
Private Sub PrintCheckLocal()
Dim SQLstatement As String
  'On Error GoTo PrintCheckLocal_Error
  
  'Create a new workspace
  'Dim MyWorkspace As Workspace
  'Set MyWorkspace = DBEngine.CreateWorkspace("Alt Workspace", "Admin", "")
  'DBEngine.Workspaces.Append MyWorkspace

  'Set dbTemp = DBEngine.Workspaces(1).OpenDatabase(db2.Name)

  ShowStatus True
  'gMessage$ = "Formatting Check " & Me![AP PO Check Number]
  'DoCmd.OpenForm "Message"
  'Forms("Message").Refresh
  'DoEvents
  
  db.Execute "DELETE * FROM [Print Check Work]"  '<<<--------create a temporary table
            
  SQLstatement = "INSERT INTO [Print Check Work]"
  SQLstatement = SQLstatement & " ([Vendor ID],[Check Number],[Total Amount]," & _
  "[Transaction Date],[Order],[Visible],[Reference #],[Invoice Amt],[Invoice #]," & _
  "[Invoice Date],[Amount Paid],[Discount],[Net Amt])"
  SQLstatement = SQLstatement & " VALUES ('" & txtFieldsVendor(0) & "','" & txtfields(34) & _
  "'," & txtfields(36) & ",#" & txtfields(7) & "#,1,True,'" & txtfields(12) & "'," & _
  txtfields(30) & ",'" & txtfields(1) & "',#" & txtfields(7) & "#," & txtfields(36) & _
  "," & txtfields(25) & "'," & txtfields(36) & ")"
  db.Execute SQLstatement

  CheckNumberCHQ "READ", db, txtfields(35), txtfields(34)
  
  'Dim rsWork As ADODB.Recordset
  'Set rsWork = New ADODB.Recordset
  'use SQL command --- FASTER
  'rsWork.Open "[Print Check Work]", db, adOpenStatic, adLockOptimistic, adCmdTable
  'With adoPrimaryRS
  'rsWork.AddNew
  '  rsWork("Vendor ID") = txtFieldsVendor(0)
  '  rsWork("Check Number") = txtfields(34)
  '  rsWork("Total Amount") = txtfields(36)
  '  rsWork("Transaction Date") = txtfields(7)
  '  rsWork("Order") = 1
  '  rsWork("Visible") = True
  '  rsWork("Reference #") = txtfields(12)
  '  rsWork("Invoice Amt") = txtfields(30)
  '  rsWork("Invoice #") = txtfields(1)
  '  rsWork("Invoice Date") = txtfields(7)
  '  rsWork("Amount Paid") = txtfields(36)
  '  rsWork("Discount") = txtfields(25)
  '  rsWork("Net Amt") = txtfields(36)
  'rsWork.Update

  'rsWork.Close
  'Set rsWork = Nothing
  'Increment Next Check No
  
  'Dim rsBank As ADODB.Recordset
  'Set rsBank = New ADODB.Recordset
  '<<<------------use SQL select statement
  'rsBank.Open "[Bank Accounts]", db, adOpenStatic, adLockOptimistic, adCmdTable
  'rsBank.Index = "PrimaryKey"
  'rsBank.MoveFirst
  'rsBank.Find "[BANK ACCT ID]='" & txtfields(35) & "'"
  'rsBank.Edit
  '  rsBank("BANK ACCT Next Check No") = txtfields(34) + 1 'Val(rsBank("BANK ACCT Next Check No") + 1)
    'Me![Next Check No] = Val(Me![AP PO Check Number]) + 1 'Val(rsBank("BANK ACCT Next Check No") + 1)
  'rsBank.Update
  
  
  'rsBank.Close
  'dbTemp.Close
  
  'DoCmd.Close A_FORM, "Message"

  'DoCmd.OpenReport "rpt - Check", A_NORMAL, , "[Check Number] = '" & Me![AP PO Check Number] & "'"

  Exit Sub
PrintCheckLocal_Error:
  Call ErrorLog("Purchase Transactions", "PrintCheckLocal", Now, Err.Number, Err.Description, True, db)
  Resume Next

End Sub

Private Sub cmdDate_Click(Index As Integer)
Select Case Index
Case 0
    Menu_Calendar.WhoCallMe True, 1002
    If txtfields(7).Text <> "" Then ADOprimaryrs![AP PO Check Date] = txtfields(7).Text
Case 1
    Dim iResponse As Integer
    iResponse = MsgBox("Due Date were set automaticly... Are sure you want to change it?", vbYesNo, "Due Date")
    If iResponse = vbNo Then Exit Sub
    Menu_Calendar.WhoCallMe True, 1001
    If txtfields(6).Text <> "" Then ADOprimaryrs![AP PO Due Date] = txtfields(6).Text
    'txtfields(6).SetFocus
Case 20
    Menu_Calendar.WhoCallMe True, 1000
    If txtfields(20).Text <> "" Then ADOprimaryrs![AP PO Date Requested] = txtfields(20).Text
    'txtfields(20).SetFocus
End Select
    'Menu_Calendar.Show vbModal

End Sub

Private Sub cmdInvoiceBackOrder_Click()
    cmdCreateInvoice_Click
End Sub

Private Sub cmdSmallBig_Click()
If cmdSmallBig.Caption = "<<" Then
     picPrimary.Height = 5775
     picPrimary.Width = 11600
     cmdSmallBig.Caption = ">>"
Else
     picPrimary.Height = frAdvance.Height + 100
     picPrimary.Width = 15050
     cmdSmallBig.Caption = "<<"
End If
    Form_Resize
    'Me.Height = picPrimary.Height + 900
    'Me.Width = picPrimary.Width + 200
    
End Sub

Private Sub cmdUpdatedua_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    cmdUpdatedua_Select (Index)
Else
    'Me.PopupMenu mnuCombo
End If
End Sub

Private Sub Command1_Click()
ShowStatus True
    If CheckDocument("select * from [AP Purchase] WHERE [AP PO Document Type]='" & DocType & "' AND [AP PO Ext Document No]='" & txtfields(0) & "'", db, False) = False Then
        Dim Response As Integer
            Response = MsgBox("Search found, Would you like to see it?", vbYesNo, "Information")
            If Response = vbYes Then
                ShowStatus True
                ClearDatasource
                RSstatement = "SHAPE {select * from [AP Purchase] WHERE [AP PO Document Type]='" & DocType & "' AND [AP PO Ext Document No]='" & txtfields(0) & "'} AS ParentCMD APPEND ({select * from [AP Purchase Detail] } AS ChildCMD RELATE [AP PO Document No] TO [AP POD Document No]) AS ChildCMD"
                OpenDB RSstatement
                picStatBox.Enabled = False
                lblStatus.BackColor = vbRed
            End If
            ProcessDoneMusic "Done"
    Else
        MsgBox "Search not found", vbInformation, "Information"
    End If
ShowStatus False
End Sub

Private Sub Form_Load()
'On Error GoTo FormErr
DocType = "RMA"
  Set db = New ADODB.Connection
  db.CursorLocation = adUseClient
  db.Provider = "MSDataShape"
  db.Open "Data " & gblADOProvider
     
     VendID0 = ""
     
     Me.Height = 6600
     Me.Width = 11475
     Me.Top = 0
     Me.Left = 0
    
    RSstatement = "SHAPE {select * from [AP Purchase] WHERE [AP PO Document Type]='" & DocType & "'} AS ParentCMD APPEND ({select * from [AP Purchase Detail] } AS ChildCMD RELATE [AP PO Document No] TO [AP POD Document No]) AS ChildCMD"
    OpenDB RSstatement
    'set the datagrid button to true
    'grdDataGrid.Columns(0).Button = True
    'grdDataGrid.Columns(6).Button = True
    'grdDataGrid.Columns(9).Button = True
    'grdDataGrid.Columns(10).Button = True
    
    grdDataGrid.AllowAddNew = True
    grdDataGrid.AllowDelete = True
    
    'cmdNext.Left = lblStatus.Width + 700
    'cmdLast.Left = cmdNext.Left + 340
    cmdSmallBig_Click
    
  GetTextColor Me
Exit Sub
FormErr:
  MsgBox Err.Description
  ShowStatus False
End Sub

Private Sub OpenDB(SQLstatement As String, Optional NewData As Boolean)
NewLoad = True
ShowStatus True
  
  'If ADOprimaryrs Is Nothing Then
  'Else
  '  ADOprimaryrs.CancelUpdate
  '  ADOprimaryrs.Close
  '  Set ADOprimaryrs = Nothing
  'End If
  
  Set ADOprimaryrs = New ADODB.Recordset
  ADOprimaryrs.Open SQLstatement, db, adOpenKeyset, adLockOptimistic, adCmdText
  With ADOprimaryrs
    If NewData = True Then
        ADOprimaryrs.Find "[AP PO Ext Document No]='" & DocType & AppLoginName & "'"
      If Not .EOF Then
        ADOprimaryrs![AP PO Ext Document No] = AppLoginName & Format(Now, "MMdd") & Right(Format(![AP PO Document No] + 6000, "0000"), 4)
        ADOprimaryrs![AP PO Status] = "Open"
        ADOprimaryrs.Update
      Else
        .MoveFirst
      End If
    End If
  End With
 Dim Ctrl As Control
 For Each Ctrl In Me.Controls
    If TypeOf Ctrl Is TextBox Or TypeOf Ctrl Is CheckBox Or TypeOf Ctrl Is ComboBox Then
        If Ctrl.DataField <> "" Then
           Set Ctrl.DataSource = ADOprimaryrs
           If TypeOf Ctrl Is TextBox And Ctrl.DataField <> "" Then
              If ADOprimaryrs("" & Ctrl.DataField & "").Type = 202 Then Ctrl.MaxLength = ADOprimaryrs("" & Ctrl.DataField & "").DefinedSize
           End If
        End If
    End If
 Next
 loadCombo "semua"
  
  If CheckNewDB(ADOprimaryrs, "RMA Entry") = True Then
    cmdAdd_Click
  Else
    Set grdDataGrid.DataSource = ADOprimaryrs("ChildCMD").UnderlyingValue
  End If
      
 grdOnAddNew = False
 ShowStatus False
 NewLoad = False
End Sub

Private Sub Form_Resize()
  'On Error Resume Next
  'This will resize the grid when the form is resized
  'grdDataGrid.Width = Me.ScaleWidth
  If fMainForm.WindowState = 1 Then Exit Sub
  If Me.WindowState = 0 Then
    Me.Height = picPrimary.Height + 850
    Me.Width = picPrimary.Width + 160
    'cmdSmallBig.Visible = True
    'cmdSmallBig_Click
  ElseIf Me.WindowState = 2 Then
    'cmdSmallBig.Visible = False
    GoTo SkipResize
  Else
    Exit Sub
  End If
  
  'If Me.Height > 6600 Then
  '   Me.Height = 9780
  'Else
  '   Me.Height = 6600
  'End If
  'If Me.Width <= 11670 Then
  '   Me.Width = 11670
  'Else
  '   Me.Width = 15300
  'End If
  'debug.print Me.Height
    'Display the proper direction
    'If Me.Width <= 11655 And Me.Height <= 5775 Then
    '    cmdSmallBig.Caption = ">>"
    'ElseIf Me.Width > 11655 And Me.Height > 5775 Then
    '    cmdSmallBig.Caption = "<<"
    'End If
SkipResize:
    picPrimary.Left = (Me.Width - picPrimary.Width - 100) / 2
    picPrimary.Top = (Me.Height - picPrimary.Height) / 2 + 100
    lblTop.Left = picPrimary.Left
    lblTop.Width = picPrimary.Width
    
    pcMajor.Height = picPrimary.ScaleHeight '- picStatBox.Height
    grdDataGrid.Height = pcMajor.ScaleHeight - grdDataGrid.Top - frThird.Height - 50
    grdDataGrid.ZOrder 0
    frThird.Top = grdDataGrid.Top + grdDataGrid.Height - 50
    'frAdvance.Height = picPrimary.ScaleHeight - frAdvance.Top - 50 '- picStatBox.Height
    picStatBox.Top = frAdvance.Height - picStatBox.Height - 50
    frButton.Top = frAdvance.Height - frButton.Height - picStatBox.Height - 100
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If mbAddNewFlag Then Exit Sub

  Select Case KeyCode
    Case vbKeyEscape
      cmdClose_Click
    Case vbKeyEnd
      If Shift = vbCtrlMask And txtfields(36).Enabled = True Then
        ADOprimaryrs![AP PO Amount Paid] = "$0.00"
        txtFields_LostFocus 36
        calculateALL
      Else
          cmdLast_Click
      End If
    Case vbKeyHome
      If Shift = vbCtrlMask And txtfields(36).Enabled = True Then
        ADOprimaryrs![AP PO Amount Paid] = txtfields(30)
        txtFields_LostFocus 36
        calculateALL
      Else
        cmdFirst_Click
      End If
    Case vbKeyUp, vbKeyPageUp
      If Shift = vbCtrlMask Then
        cmdFirst_Click
      Else
        cmdPrevious_Click
      End If
    Case vbKeyDown, vbKeyPageDown
      If Shift = vbCtrlMask Then
        cmdLast_Click
      Else
        cmdNext_Click
      End If
  End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo FormErr
    ShowStatus False
    If UnloadForm(ADOprimaryrs) = 0 Then
        db.Close
        Set db = Nothing
    Else
        Cancel = 1
    End If
    Set frm_AP_RMA_Entry = Nothing
Exit Sub
FormErr:
  MsgBox Err.Description
  ShowStatus False
End Sub

Private Sub adoPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
Dim Response As Integer
ShowStatus True
  If ADOprimaryrs.BOF Or ADOprimaryrs.EOF Then GoTo JumpIf
  If ADOprimaryrs![AP PO Posted YN] = True Then
     ButtEnabled False
  Else
     ButtEnabled True
     If IsNull(ADOprimaryrs![AP PO Status]) Or ADOprimaryrs![AP PO Status] = "Open" Then
        cmdApprove.Picture = fMainForm.imlIcons.ListImages("Locked").Picture
     Else
        cmdApprove.Picture = fMainForm.imlIcons.ListImages("Approved").Picture
     End If
     
     If ADOprimaryrs![AP PO Amount Paid] > 0 Then
        txtfields(7).Enabled = True
        txtfields(34).Enabled = True
        cmdDate(0).Enabled = True
        txtfields(35).Enabled = True
        cmdbankAccount.Enabled = True
    Else
        txtfields(7).Enabled = False
        txtfields(34).Enabled = False
        cmdDate(0).Enabled = False
        txtfields(35).Enabled = False
        cmdbankAccount.Enabled = False
     End If
  End If
   If mbAddNewFlag = False Then
        If IsNull(ADOprimaryrs![AP PO Vendor ID]) Then
        Else
            CustomerData Me, db, ADOprimaryrs![AP PO Vendor ID], False
        End If
        txtFieldsVendor(0).Locked = True
        'txtFieldsShip(0).Locked = True
        txtfields(36).Locked = False
        If Trim(ADOprimaryrs![AP PO Vendor ID]) <> "" Then GetWEBMAILvendor ADOprimaryrs![AP PO Vendor ID], db, Me
   Else
        lblmail.Visible = False
        lblweb.Visible = False
        txtFieldsVendor(0).Locked = False
        'txtFieldsShip(0).Locked = False
        txtfields(36).Locked = True
   End If
JumpIf:
  GetTextColor Me
  'This will display the current record position for this recordset
  lblStatus.Caption = "Record: " & CStr(ADOprimaryrs.AbsolutePosition) & " of " & CStr(ADOprimaryrs.RecordCount)
ShowStatus False
End Sub

Private Sub ButtEnabled(SetEnabled As Boolean)
        pcMajor.Enabled = SetEnabled
        PcData.Enabled = SetEnabled
        'frTotal.Enabled = SetEnabled
        imgPosted.Visible = Not SetEnabled
        imgOpen.Visible = SetEnabled
        cmdUpdate.Enabled = SetEnabled
        cmdDelete.Enabled = SetEnabled
        cmdRefresh.Enabled = SetEnabled
        cmdCreditLimit.Enabled = SetEnabled
        If mbAddNewFlag = False Then
            cmdCreateInvoice.Enabled = SetEnabled   'cmdPrint
            cmdPrint.Enabled = True
            cmdApprove.Enabled = SetEnabled
        Else
            cmdApprove.Enabled = False
            cmdPrint.Enabled = False
        End If
 Dim cbCtrl As ComboBox
 For Each cbCtrl In Me.cbPurchase
    cbCtrl.Enabled = SetEnabled
    cmdUpdatedua(cbCtrl.Index).Enabled = SetEnabled
 Next
 
 cmdbankAccount.Enabled = SetEnabled
 cmdDate(1).Enabled = SetEnabled
 cmdDate(20).Enabled = SetEnabled
End Sub

Private Sub adoPrimaryRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This is where you put validation code
  'This event gets called when the following actions occur
  Dim bCancel As Boolean

  Select Case adReason
  Case adRsnAddNew
  Case adRsnClose
  Case adRsnDelete
  Case adRsnFirstChange
  Case adRsnMove
  Case adRsnRequery
  Case adRsnResynch
  Case adRsnUndoAddNew
  Case adRsnUndoDelete
  Case adRsnUndoUpdate
  Case adRsnUpdate
  End Select

  If bCancel Then adStatus = adStatusCancel
End Sub

Private Sub cmdAdd_Click()
'On Error GoTo AddErr
ShowStatus True
  If cmdAdd.Caption = "&Save" Then
     If Not CheckEmpty Then
        ShowStatus False
        Exit Sub
     End If
     With ADOprimaryrs
         mbAddNewFlag = False
         ADOprimaryrs.Update
         'cmdUpdate_Click
         '.MovePrevious
         'grdDataGrid.HoldFields
         'grdDataGrid.ReBind
         'grdDataGrid.RefreshLoadRS SQLstatement
         ClearDatasource
         RSstatement = "SHAPE {select * from [AP Purchase] Where [AP PO Document Type]='" & DocType & "' AND [AP PO Posted YN]=False} AS ParentCMD APPEND ({select * from [AP Purchase Detail] } AS ChildCMD RELATE [AP PO Document No] TO [AP POD Document No]) AS ChildCMD"
         OpenDB RSstatement, True
         'If ADOprimaryrs.EOF Then
         'Else
         '   txtFields(12).SetFocus
            'txtFields(12) = AppLoginName & Format(Now, "MMdd") & Format(txtFields(2), "000")
         'End If
         NewLoad = False
         
     End With
     cmdAdd.Caption = "&Add"
     'cmdLookupShip.Visible = False
     SetButtons True
     cmdCreateInvoice.Enabled = True
     cmdPrint.Enabled = True
Else
    mbAddNewFlag = True
    cmdCreateInvoice.Enabled = False
    cmdPrint.Enabled = False
  With ADOprimaryrs
    If Not (.BOF Or .EOF) Then
      mvBookMark = .Bookmark
    End If
     NewLoad = True
     'cmdLookupShip.Visible = True
    .AddNew
    txtfields(12) = DocType & AppLoginName
    txtfields(3) = DocType
    txtfields(4) = FormatDate(Now)
    'txtFields(7) = txtFields(4)
    SetDueDate
    txtSalesPerson(0) = AppLoginName
    lblStatus.Caption = "Add record"
        Dim i As Integer
        If mbAddNewFlag = True Then
           For i = 24 To 30
             Select Case i
                Case 24, 25, 26, 27, 30, 33, 36
                    txtfields(i) = "$0.00"
                Case 28, 29
                    txtfields(i) = "00.00"
             End Select
           Next
           'txtfields(33) = "$0.00"
           'txtfields(36) = "$0.00"
        End If
    SetButtons False
  End With
End If
  ShowStatus False
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub
Public Sub SetDueDate()
If txtfields(4) = "" Then Exit Sub
If mbAddNewFlag = True Then
    DueDateDay db, cbPurchase(5), txtfields(4), txtfields(6), True
Else
    DueDateDay db, cbPurchase(5), txtfields(4), txtfields(6)
End If
End Sub

Private Sub cmdDelete_Click()
Dim DocNo As String 'picStatBox
'Dim DelStatus As String

DocNo = txtfields(12).Text

'     DelStatus = DataDelete(ADOprimaryrs, Me, True)
     
'     If DelStatus = False Then
'        MsgBox "An error occured while attempting to delete " & DocNo & ", closing the " & DocType
'        Unload Me
'     Else
''        If picStatBox.Enabled = False Then
        ShowStatus True
        ClearDatasource
        db.Execute "DELETE FROM [AP Purchase] WHERE [" & txtfields(12).DataField & "]='" & DocNo & "'"
'        MsgBox lblTop & "[" & DocNo & "] has been deleted. Refreshing the database process will take place after this.", vbInformation, "Information"
        'ADOprimaryrs.Requery
            MsgBox lblTop & "[" & DocNo & "] has been deleted." & vbCr & _
            "Opening Unposted " & DocType & " Form", vbInformation, "Information"
'            cmdUnPostedDoc_Click
        RSstatement = "SHAPE {select * from [AP Purchase] Where [AP PO Document Type]='" & DocType & "' AND [AP PO Posted YN]=False} AS ParentCMD APPEND ({select * from [AP Purchase Detail] } AS ChildCMD RELATE [AP PO Document No] TO [AP POD Document No]) AS ChildCMD"
        OpenDB RSstatement
        
        ShowStatus False
'        Else
'            MsgBox lblTop & "[" & DocNo & "] has been deleted.", vbInformation, "Information"
'        End If
'     End If
End Sub

Private Sub cmdRefresh_Click()
    RefreshButton ADOprimaryrs, grdDataGrid
End Sub

Private Sub cmdCancel_Click()
  ShowStatus True
  SetButtons True
  cmdAdd.Caption = "&Add"
  cmdCreateInvoice.Visible = True
  cmdCancel.Visible = False
  'cmdLookupShip.Visible = False
  mbAddNewFlag = False
  ADOprimaryrs.CancelUpdate
  NewLoad = False
  If ADOprimaryrs.RecordCount > 0 Then
    ADOprimaryrs.MoveLast
  Else
    MsgBox "No data to publish. Exiting " & Me.Caption
    Unload Me
    Exit Sub
  End If
  ADOprimaryrs.Resync adAffectCurrent
  If mvBookMark > 0 Then
    ADOprimaryrs.Bookmark = mvBookMark
  Else
    ADOprimaryrs.MoveFirst
  End If
  ShowStatus False
End Sub

Private Sub cmdUpdate_Click()
'Dim FlagStatus As Boolean
    
  'FlagStatus = False

  Call UpdateButton(ADOprimaryrs, mbAddNewFlag)
  
  'SetButtons FlagStatus
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdFirst_Click()
  'On Error GoTo GoFirstError

  ADOprimaryrs.MoveFirst

  Exit Sub

GoFirstError:
  MsgBox Err.Description
End Sub

Private Sub cmdLast_Click()
  'On Error GoTo GoLastError

  ADOprimaryrs.MoveLast

  Exit Sub

GoLastError:
  MsgBox Err.Description
End Sub

Private Sub cmdNext_Click()
  'On Error GoTo GoNextError

  If Not ADOprimaryrs.EOF Then ADOprimaryrs.MoveNext
  If ADOprimaryrs.EOF And ADOprimaryrs.RecordCount > 0 Then
    ProcessDoneMusic "Done"
     'moved off the end so go back
    ADOprimaryrs.MoveLast
  End If
  'show the current record

  Exit Sub
GoNextError:
  MsgBox Err.Description
End Sub

Private Sub cmdPrevious_Click()
  'On Error GoTo GoPrevError

  If Not ADOprimaryrs.BOF Then ADOprimaryrs.MovePrevious
  If ADOprimaryrs.BOF And ADOprimaryrs.RecordCount > 0 Then
    ProcessDoneMusic "Done"
    'moved off the end so go back
    ADOprimaryrs.MoveFirst
  End If
  'show the current record

  Exit Sub

GoPrevError:
  MsgBox Err.Description
End Sub

Private Sub SetButtons(bVal As Boolean)
    
    If mbAddNewFlag = True Then
        cmdAdd.Caption = "&Save"
        cmdCancel.Visible = True
        cmdCancel.Left = cmdUpdate.Left
        cmdCancel.Top = cmdUpdate.Top
    Else
        cmdAdd.Visible = bVal
        cmdCancel.Visible = False
    End If
        cmdUpdate.Visible = bVal
        cmdDelete.Visible = bVal
        cmdClose.Visible = bVal
        cmdRefresh.Visible = bVal
        cmdNext.Enabled = bVal
        cmdFirst.Enabled = bVal
        cmdLast.Enabled = bVal
        cmdPrevious.Enabled = bVal
End Sub

Private Sub grdDataGrid_Error(ByVal DataError As Integer, Response As Integer)
    If DataGridKnownError(DataError) Then
        Response = 0
    End If
End Sub

Private Sub grdDataGrid_GotFocus()
Dim CreateOrder As Integer
    If mbAddNewFlag = True Then
        'cmdAdd.SetFocus
        CreateOrder = MsgBox("This Request will save the data to the database? Are sure to continue", vbYesNo, "Save Quote")
        If CreateOrder = vbNo Then Exit Sub
        cmdAdd_Click
    End If
End Sub

Private Sub grdDataGrid_LostFocus()
    SendKeys ("{LEFT}")
    'grdDataGrid_AfterColEdit (grdDataGrid.col)
    'ADOprimaryrs.Update
    CalcTotals
End Sub

Private Sub grdDataGrid_OnAddNew()
    grdOnAddNew = True
    'grdDataGrid.Row = 1
End Sub

Private Sub grdDataGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If NewLoad = True Then Exit Sub
    If ADOprimaryrs.BOF Or ADOprimaryrs.EOF Then Exit Sub
        If grdDataGrid.col > 0 And grdDataGrid.Row > -1 Then
            If grdDataGrid.Columns(0) = "" Then
                MsgBox "You must select Item ID first before continue", vbInformation, "Error Selection"
                GoTo Damn_Attempt
            End If
        End If
    CalcTotals
Select Case grdDataGrid.col
  Case 2, 3, 5, 7
     grdDataGrid.AllowUpdate = True
  'Case 3
  '   grdDataGrid.AllowUpdate = True
  'Case 6
  '   grdDataGrid.AllowUpdate = True
  Case Else
     grdDataGrid.AllowUpdate = False
  End Select
Exit Sub
Damn_Attempt:
     grdDataGrid.AllowUpdate = False
     grdDataGrid.col = 0
exit_sub:
End Sub

Private Sub grdDataGrid_AfterColEdit(ByVal ColIndex As Integer)
'On Error Resume Next
  If grdDataGrid.Columns(2).Text = "" Then grdDataGrid.Columns(2).Text = 0
  If grdDataGrid.Columns(3).Text = "" Then grdDataGrid.Columns(3).Text = 0
  If grdDataGrid.Columns(5).Text = "" Then grdDataGrid.Columns(5).Text = 0
  If grdDataGrid.Row = -1 Or grdDataGrid.Columns(0) = "" Then Exit Sub
      SendKeys ("{ENTER}")
  If grdDataGrid.Row > 0 Then
      SendKeys ("{up}")
      SendKeys ("{down}")
  ElseIf grdDataGrid.Row = 0 Then
      SendKeys ("{down}")
      SendKeys ("{up}")
  End If
  CalculateTable
  ADOprimaryrs.UpdateBatch adAffectAll
End Sub

Private Sub grdDataGrid_BeforeDelete(Cancel As Integer)
    Dim DeleteCration As Integer
    
    DeleteCration = MsgBox("Attempting to delete the data. " & vbCr & "Are you sure?", vbYesNo, "Deleting Confirmation")
    If DeleteCration = vbNo Then Cancel = 1

End Sub

Private Sub NewgrdDatagrid()
    NewLoad = True
    NewRowForDataGrid ADOprimaryrs, grdDataGrid, "AP PO Date", txtfields(4).Text
    grdOnAddNew = False
    NewLoad = False
End Sub

Private Sub grdDataGrid_ButtonClick(ByVal ColIndex As Integer)
On Error GoTo Error_ButtClick
If mbAddNewFlag = True Then Exit Sub
If grdDataGrid.Columns(0) <> "" Then grdOnAddNew = False
Select Case ColIndex
Case 0   'select the item from the ITEM_ID
    INV_ITEM
Case 6   'Get the type of account for the selected row
    Menu_Calendar.WhoCallMe True, 1020
    'Menu_Calendar.Show vbModal
Case 9   'Get the type of account for the selected row
    COA_grdDataGrid_Butt
Case 10   'Select the project that have been working on
    Proj_Projects
Case Else
End Select

If grdOnAddNew = True And grdDataGrid.Columns(0) <> "" Then NewgrdDatagrid
grdDataGrid_AfterColEdit 0
Exit Sub
Error_ButtClick:
    MsgBox "Please click the Table box before clicking the button"
End Sub

Private Sub INV_ITEM()
   AllLookup.GetWhichTable 1350, "SELECT [INV ITEM Id], [INV ITEM Description]," & _
   "[INV ITEM Unit],[INV ITEM Price], [INV ITEM Inventory Account], [INV ITEM Qty On Hand], " & _
   "[INV ITEM Qty On Order], [INV ITEM Taxable YN],[INV ITEM Last Cost] FROM [INV Items] " & _
   "WHERE [INV ITEM Inactive YN]=FALSE ", "Product", _
   "Item ID//Item Description//Unit//Price//Sales Account//Qty On Hand//Qty On Order//Taxable//Cost", db
   'AllLookup.Show vbModal

End Sub

Private Sub COA_grdDataGrid_Butt()
   AllLookup.GetWhichTable 1302, "Select [GL COA Account No],[GL COA Account Name]," & _
   "[GL COA Asset Type] From [GL Chart Of Accounts] ", "GL Accounts", _
   "Account No//Account Type//Account Type", db
   'AllLookup.Show vbModal
   
End Sub

Private Sub Proj_Projects()
   AllLookup.GetWhichTable 1304, "Select [PROJ ID],[PROJ Name]," & _
   "[PROJ Description] From [PROJ Projects] ", "Project", _
   "Project ID//Project Name//Description", db
   'AllLookup.Show vbModal
   
End Sub

Private Sub CalculateTable()
Dim i As Integer

    'get the total value for the selected rod
    grdDataGrid.Columns(8).Text = grdDataGrid.Columns(3).Value * grdDataGrid.Columns(5).Text
        
Exit Sub
CalculateTable_ERR:

End Sub

Private Sub CalcSalesTaxPercent()

  ''On Error GoTo CalcSalesTaxPercent_Error

  'Calculate tax percent based on tax group

  'Dim db As Database
  Dim rs1 As ADODB.Recordset
  'Dim rs2 As ADODB.Recordset
  Dim TaxPercent#

  Set rs1 = New ADODB.Recordset
  rs1.Open ("SELECT * FROM [SYS Tax Group Detail] WHERE [SYS TAXGRPD Group ID] = '" & cbPurchase(3) & "'"), db, adOpenStatic, adLockOptimistic
  ''On Error Resume Next
  rs1.MoveFirst
  'If Err <> 0 Then
    'No taxes found for this group
  '  TaxPercent# = 0
  'Else
    Do While Not rs1.EOF
      TaxPercent# = TaxPercent# + LookRecord("[SYS Tax Percent]", "[SYS Tax]", db, "[SYS Tax ID] = '" & rs1("SYS TAXGRPD Tax ID") & "'")
      rs1.MoveNext
    Loop
  'End If

  txtfields(29) = TaxPercent#

  Exit Sub
CalcSalesTaxPercent_Error:
  Call ErrorLog(DocType & " Transactions", "CalcSalesTaxPercent", Now, Err.Number, Err.Description, True, db)
  Resume Next
  
End Sub

Private Sub CalcTotals()
If NewLoad = True Then Exit Sub
Dim Calcrs As ADODB.Recordset

    'If Not frTotal.Enabled Then Exit Sub
    If ADOprimaryrs.EOF = True Or ADOprimaryrs.BOF = True Then Exit Sub
    If mbAddNewFlag = True Then Exit Sub
    Set Calcrs = New ADODB.Recordset
    Calcrs.Open "SELECT [AP POD Item Total] FROM [AP Purchase Detail] WHERE [AP POD Document No]=" & txtfields(2), db, adOpenStatic, adLockOptimistic
    
    Dim ItemTotal As Double
    With Calcrs
        If .RecordCount > 0 Then
            .MoveFirst
            While Not .EOF
                ItemTotal = ItemTotal + Calcrs![AP POD Item Total]
                .MoveNext
            Wend
        Else
            ItemTotal = 0
        End If
    End With
    txtfields(24) = FormatCurr(CCur(ItemTotal))
    If txtfields(24).Enabled = True Then ADOprimaryrs![AP PO Subtotal] = txtfields(24)
    calculateALL
    ADOprimaryrs.Update
    
    Calcrs.Close
    Set Calcrs = Nothing
End Sub

Private Sub calculateALL()

  'Calculate totals for this invoice
        If txtfields(24) = "" Then txtfields(24) = "$0.00"
        If CCur(txtfields(24)) <= 0 Then
            txtfields(25) = "$0.00"
            txtfields(26) = "$0.00"
            txtfields(30) = "$0.00"
            txtfields(33) = "$0.00"
            txtfields(36) = "$0.00"
            Exit Sub
        End If
  ''On Error GoTo CalcTotals_Error
  'Calculate discount
        If Val(txtfields(28)) <> 0 Then
            txtfields(25) = txtfields(24) * (txtfields(28) / 100)
            txtfields(25) = FormatCurr(txtfields(25))
        Else
            txtfields(25) = "$0.00"
        End If
        If Val(txtfields(29)) <> 0 Then
            txtfields(26) = txtfields(24) * (txtfields(29) / 100)
            txtfields(26) = FormatCurr(txtfields(26))
        Else
            txtfields(26) = "$0.00"
        End If

  If txtfields(26) = "" Then
    txtfields(26) = "$0.00"
  End If
  
  'Calculate Total
  txtfields(30) = txtfields(24) - txtfields(25) + txtfields(26) + txtfields(27)

  txtfields(30) = FormatCurr(txtfields(30))
  'txtfields(1) = txtfields(30) - txtfields(5)
  
  If txtfields(36) = "." Or txtfields(36) = "" Then txtfields(36) = "$0.00"
  txtfields(33) = txtfields(30) - txtfields(36)
  
  txtfields(33) = FormatCurr(txtfields(33))
  
  Exit Sub
CalcTotals_Error:
End Sub

Private Function CheckEmpty() As Boolean
 Dim Ctrl As Control
 For Each Ctrl In Me.Controls
   If TypeOf Ctrl Is TextBox Then
    If Ctrl.DataField <> "" Then
      Select Case Ctrl.Index
      Case 0
        If Ctrl.Text = "" Then
            MsgBox "There is an empty data either in Customer, Shipping ID or SalesPerson", vbInformation, "Empty Data"
            CheckEmpty = False
            GoTo Out_Of_Here
        End If
      Case 4
        If Ctrl.Text = "" And Ctrl.Name = "txtFields" Then
            MsgBox "There is an empty data in " & lblfields(Ctrl.Index).Caption, vbInformation, "Empty Data"
            CheckEmpty = False
            GoTo Out_Of_Here
        End If
      Case 20
        If Trim(Ctrl.Text) = "" And mbAddNewFlag = False And LCase(Ctrl.Name) = "txtfields" Then
            MsgBox "There is an empty data in " & lblfields(Ctrl.Index).Caption, vbInformation, "Empty Data"
            CheckEmpty = False
            GoTo Out_Of_Here
        End If
      Case 7, 34, 35
        If txtfields(36) <> "" And Not mbAddNewFlag Then
            If Trim(Ctrl.Text) = "" And txtfields(36) <> "$0.00" Then
                MsgBox "There is an empty data in " & lblfields(Ctrl.Index).Caption, vbInformation, "Empty Data"
                CheckEmpty = False
                GoTo Out_Of_Here
            End If
        End If
      Case 14
        If Trim(Ctrl.Text) = "" And LCase(Ctrl.Name) = "txtfields" Then
            MsgBox Ctrl.Name & "There is an empty data in " & lblLabels(Ctrl.Index).Caption, vbInformation, "Empty Data"
            CheckEmpty = False
            GoTo Out_Of_Here
        End If
      End Select
    End If
   End If
    
   If TypeOf Ctrl Is ComboBox Then
      If Ctrl.Text = "" And Ctrl.Index <> 2 Then
        MsgBox "There is an empty data in " & lblfields(Ctrl.Index).Caption, vbInformation, "Empty Data"
        CheckEmpty = False
        GoTo Out_Of_Here
      End If
   End If
 Next
 'Dim i As Integer
 'If mbAddNewFlag = True Then
 '   For i = 24 To 30
 '       If i <> 29 Then txtfields(i) = 0
 '   Next
 '   txtfields(33) = "$0.00"
 '   txtfields(36) = "$0.00"
 'End If
 CheckEmpty = True
Out_Of_Here:
End Function

Private Sub lblweb_Click()
        fMainForm.callWebPage lblweb.ToolTipText
End Sub

Private Sub txtfields_GotFocus(Index As Integer)
    TxtGotFocus txtfields(Index)
End Sub

Private Sub txtfields_KeyPress(Index As Integer, KeyAscii As Integer)
Dim keyResponse As Boolean
Select Case Index
Case 24, 25, 26, 27, 28, 29, 30, 36
     keyResponse = CtrlValidate(KeyAscii, "0123456789.")
     If keyResponse = True Then
     Else
        KeyAscii = 0
     End If
Case 34, 35
    keyResponse = CtrlValidate(KeyAscii, "0123456789")
    If keyResponse = True Then
    Else
       KeyAscii = 0
    End If
End Select
End Sub

Private Sub txtFields_LostFocus(Index As Integer)
On Error GoTo exit_EditMode
If ADOprimaryrs.EditMode = adEditInProgress Then
    If IsNull(ADOprimaryrs![AP PO Status]) Or ADOprimaryrs![AP PO Status] <> "Open" Then
        ValidatePower txtfields(12).Text, "Edit", DocType, db
        ADOprimaryrs![AP PO Status] = "Open"
        ADOprimaryrs.Update
        cmdApprove.Picture = fMainForm.imlIcons.ListImages("Locked").Picture
    End If
End If
exit_EditMode:

Select Case Index
Case 27
    CalcTotals
    GetTextColor Me
Case 28
    'If Trim(txtfields(28)) = "" Then txtfields(28) = "00.00"
    'db.Execute "UPDATE [AP Customer] SET [AR CUST Discount %] = " & CDbl(txtfields(28)) & " WHERE [AR CUST Customer ID] = '" & txtFieldsVendor(0).Text & "'"
    CalcTotals
    'GetTextColor Me
Case 34
    'CheckNumberCHQ
    If Trim(txtfields(35).Text) = "" Or Trim(txtfields(34).Text) = "" Then
        txtfields(34) = " "
        Exit Sub
    ElseIf Trim(txtfields(34)) <> "" Then
        MsgBox "Please select Bank Account first before writing a check number", vbCritical, "Information"
        txtfields(34) = " "
        Exit Sub
    End If
    'If CheckCheckNumber(txtFields(1).Text, txtFields(5).Text, db, True) = "Found" Then
    '    response% = MsgBox("Check Number is already used. Would you like to open Check Management?", vbYesNo, "Information")
    '    If response% = vbYes Then
    '        frm_Check_Management.OpenPosted txtFields(5).Text
    '    End If
    '    txtFields(5).Text = ""
    '    ShowStatus False
    '    Exit Sub
    'End If
    If CheckNumberCHQ("check", db, txtfields(35).Text, txtfields(34).Text) = "Found" Then
        Dim Response As Integer
        Response% = MsgBox("Check Number is already used. Would you like to open Check Management?", vbYesNo, "Information")
        If Response = vbYes Then
            frm_Check_Management.OpenPosted txtfields(34).Text
            txtfields(34) = " "
        Else
            txtfields(34) = " "
        End If
    End If
Case 35
    If BankAcct35 = txtfields(35) Then Exit Sub
    If Trim(txtfields(35)) = "" Then Exit Sub
    If IsNumeric(txtfields(35).Text) Then
        CheckDocument "SELECT [GL COA Account No] FROM [GL Chart Of Accounts] WHERE [GL COA Account No]='" & txtfields(35).Text & "'", db, False, txtfields(35)
    Else
        MsgBox "Only numeric character is accepted", vbInformation, "Information"
        txtfields(35) = " "
    End If
    BankAcct35 = txtfields(35)
Case 36
    calculateALL
    If Trim(txtfields(36)) = "" Then txtfields(36) = "$0.00"
    If CCur(txtfields(36)) > 0 And CCur(txtfields(33)) >= 0 Then
        txtfields(7).Enabled = True
        txtfields(34).Enabled = True
        cmdDate(0).Enabled = True
        cmdbankAccount.Enabled = True
        txtfields(35).Enabled = True
    ElseIf CCur(txtfields(33)) < 0 Then
        MsgBox "Balance must not be less than Zero", vbInformation, "Information"
        txtfields(36) = txtfields(24)
        txtfields(33) = "$0.00"
        txtfields(7).Enabled = True
        txtfields(34).Enabled = True
        cmdDate(0).Enabled = True
    Else
        ''txtfields(7).Text = " "
        txtfields(34).Text = " "
        txtfields(7).Enabled = False
        txtfields(34).Enabled = False
        cmdDate(0).Enabled = False
        cmdbankAccount.Enabled = False
        txtfields(35).Enabled = False
    End If
    GetTextColor Me
End Select
If Trim(txtfields(Index)) = "" Then
    txtfields(Index) = " "
    Exit Sub
End If
End Sub

Private Sub txtFieldsVendor_GotFocus(Index As Integer)
    TxtGotFocus txtFieldsVendor(Index)
End Sub

Private Sub ClearDatasource()
 Dim Ctrl As Control
 For Each Ctrl In Me.Controls
    If TypeOf Ctrl Is TextBox Or TypeOf Ctrl Is CheckBox Or TypeOf Ctrl Is ComboBox Then
        If Ctrl.DataField <> "" Then
           Set Ctrl.DataSource = Nothing
        End If
    End If
 Next
    Set grdDataGrid.DataSource = Nothing
    ADOprimaryrs.CancelUpdate
    ADOprimaryrs.Close
    Set ADOprimaryrs = Nothing
End Sub

Private Sub txtFieldsVendor_LostFocus(Index As Integer)
If txtFieldsVendor(0).Locked = True Then Exit Sub
Dim i As Integer
If txtFieldsVendor(0).Text = "" Then
    For i = 0 To txtFieldsVendor.UBound
        txtFieldsVendor(i).Text = ""
    Next
ElseIf VendID0 <> txtFieldsVendor(0).Text Then
    If Index = 0 Then
        VendorID "Select [AP VEN ID],[AP VEN Name],[AP VEN Address 1]," & _
        "[AP VEN Address 2],[AP VEN City],[AP VEN State],[AP VEN Postal]," & _
        "[AP VEN Country],[AP VEN Phone],[AP VEN Fax]," & _
        "[AP VEN Remit Name],[AP VEN Remit Address 1],[AP VEN Remit Address 2]," & _
        "[AP VEN Remit City],[AP VEN Remit State],[AP VEN Remit Country],[AP VEN Remit Country]," & _
        "[AP VEN Remit Phone],[AP VEN Remit Fax] From [AP Vendor] WHERE [AP VEN ID]='" & txtFieldsVendor(0).Text & "'", db, Me
    End If
End If
VendID0 = txtFieldsVendor(0).Text
End Sub

Private Sub txtSalesPerson_GotFocus(Index As Integer)
    TxtGotFocus txtSalesPerson(Index)
End Sub
