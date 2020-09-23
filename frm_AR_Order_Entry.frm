VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_AR_Order_Entry 
   Caption         =   "Order Entry"
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
         TabIndex        =   130
         Top             =   360
         Width           =   1935
         Begin VB.CommandButton cmdLookupCust 
            Height          =   285
            Left            =   0
            Picture         =   "frm_AR_Order_Entry.frx":0000
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
            MouseIcon       =   "frm_AR_Order_Entry.frx":030A
            MousePointer    =   99  'Custom
            TabIndex        =   132
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
            MouseIcon       =   "frm_AR_Order_Entry.frx":0614
            MousePointer    =   99  'Custom
            TabIndex        =   131
            Top             =   60
            Visible         =   0   'False
            Width           =   435
         End
      End
      Begin VB.PictureBox picMajorbutton 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   8040
         ScaleHeight     =   375
         ScaleWidth      =   975
         TabIndex        =   58
         Top             =   240
         Width           =   975
         Begin VB.CommandButton cmdSmallBig 
            Caption         =   ">>"
            Height          =   375
            Left            =   450
            TabIndex        =   66
            ToolTipText     =   "Enlarge/Shrink"
            Top             =   0
            Width           =   460
         End
         Begin VB.CommandButton Command2 
            Height          =   375
            Left            =   0
            Picture         =   "frm_AR_Order_Entry.frx":091E
            Style           =   1  'Graphical
            TabIndex        =   138
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
         TabIndex        =   83
         Top             =   0
         Width           =   11535
         Begin VB.Frame frFirst 
            Height          =   3015
            Left            =   0
            TabIndex        =   91
            Top             =   0
            Width           =   11535
            Begin VB.TextBox txtFieldsCust 
               DataField       =   "AR ORDER Customer ID"
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
               TabIndex        =   92
               Top             =   120
               Width           =   2295
               Begin VB.TextBox txtfields 
                  DataField       =   "AR ORDER Invoiced Document"
                  Height          =   285
                  Index           =   1
                  Left            =   1080
                  Locked          =   -1  'True
                  TabIndex        =   136
                  Top             =   1680
                  Width           =   1095
               End
               Begin VB.TextBox txtfields 
                  DataField       =   "AR ORDER PO ID"
                  Height          =   285
                  Index           =   14
                  Left            =   1080
                  TabIndex        =   5
                  Top             =   960
                  Width           =   1095
               End
               Begin VB.TextBox txtSalesPerson 
                  DataField       =   "AR ORDER Salesperson"
                  Height          =   285
                  Index           =   0
                  Left            =   1080
                  Locked          =   -1  'True
                  TabIndex        =   41
                  Top             =   2400
                  Width           =   1095
               End
               Begin VB.TextBox txtfields 
                  DataField       =   "AR ORDER Date"
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
                  TabIndex        =   40
                  Top             =   2040
                  Width           =   1095
               End
               Begin VB.TextBox txtfields 
                  DataField       =   "AR ORDER Ext Document #"
                  Height          =   285
                  Index           =   12
                  Left            =   1080
                  Locked          =   -1  'True
                  TabIndex        =   39
                  Top             =   1320
                  Width           =   1095
               End
               Begin VB.Label lblLabels 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000004&
                  Caption         =   "Invoice No:"
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   1
                  Left            =   120
                  TabIndex        =   137
                  Top             =   1680
                  Width           =   975
               End
               Begin VB.Image imgOpen 
                  Height          =   495
                  Left            =   120
                  Picture         =   "frm_AR_Order_Entry.frx":0C28
                  Top             =   240
                  Visible         =   0   'False
                  Width           =   1980
               End
               Begin VB.Label lblLabels 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000004&
                  Caption         =   "Cust. PO No:"
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   14
                  Left            =   120
                  TabIndex        =   96
                  Top             =   960
                  Width           =   975
               End
               Begin VB.Image imgPosted 
                  Height          =   540
                  Left            =   0
                  Picture         =   "frm_AR_Order_Entry.frx":14E8
                  Top             =   240
                  Width           =   2250
               End
               Begin VB.Label lblLabels 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000000&
                  Caption         =   "Salesperson:"
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   28
                  Left            =   120
                  TabIndex        =   95
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
                  TabIndex        =   94
                  Top             =   2040
                  Width           =   975
               End
               Begin VB.Label lblLabels 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000004&
                  Caption         =   "Order No:"
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   12
                  Left            =   120
                  TabIndex        =   93
                  Top             =   1320
                  Width           =   975
               End
            End
            Begin VB.TextBox txtfields 
               DataField       =   "AR ORDER Document Type"
               Height          =   285
               Index           =   3
               Left            =   9600
               Locked          =   -1  'True
               TabIndex        =   99
               Top             =   1200
               Width           =   1335
            End
            Begin VB.TextBox txtfields 
               DataField       =   "AR ORDER Document #"
               Height          =   285
               Index           =   2
               Left            =   9600
               Locked          =   -1  'True
               TabIndex        =   98
               Top             =   1800
               Width           =   1335
            End
            Begin VB.TextBox txtfields 
               DataField       =   "AR ORDER Taxable Subtotal"
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
               Index           =   59
               Left            =   9600
               TabIndex        =   97
               Top             =   600
               Width           =   1335
            End
            Begin VB.TextBox txtFieldsShip 
               DataField       =   "AR ORDER Shipping Fax"
               Height          =   285
               Index           =   9
               Left            =   7800
               Locked          =   -1  'True
               TabIndex        =   38
               Top             =   2520
               Width           =   1215
            End
            Begin VB.TextBox txtFieldsShip 
               DataField       =   "AR ORDER Shipping Phone"
               Height          =   285
               Index           =   8
               Left            =   5760
               Locked          =   -1  'True
               TabIndex        =   37
               Top             =   2520
               Width           =   1215
            End
            Begin VB.TextBox txtFieldsCust 
               DataField       =   "AR ORDER Billing Fax"
               Height          =   285
               Index           =   9
               Left            =   3240
               Locked          =   -1  'True
               TabIndex        =   29
               Top             =   2520
               Width           =   1215
            End
            Begin VB.TextBox txtFieldsCust 
               DataField       =   "AR ORDER Billing Phone"
               Height          =   285
               Index           =   8
               Left            =   1200
               Locked          =   -1  'True
               TabIndex        =   28
               Top             =   2520
               Width           =   1215
            End
            Begin VB.TextBox txtFieldsShip 
               DataField       =   "AR ORDER Shipping Address 1"
               Height          =   285
               Index           =   2
               Left            =   5760
               Locked          =   -1  'True
               TabIndex        =   31
               Top             =   1080
               Width           =   3255
            End
            Begin VB.TextBox txtFieldsShip 
               DataField       =   "AR ORDER Shipping Customer"
               Height          =   285
               Index           =   1
               Left            =   5760
               Locked          =   -1  'True
               TabIndex        =   30
               Top             =   720
               Width           =   3255
            End
            Begin VB.TextBox txtFieldsShip 
               DataField       =   "AR ORDER Shipping ID"
               Height          =   285
               Index           =   0
               Left            =   5760
               Locked          =   -1  'True
               TabIndex        =   3
               Top             =   360
               Width           =   1455
            End
            Begin VB.TextBox txtFieldsCust 
               DataField       =   "AR ORDER Billing Address 1"
               Height          =   285
               Index           =   2
               Left            =   1200
               Locked          =   -1  'True
               TabIndex        =   22
               Top             =   1080
               Width           =   3255
            End
            Begin VB.TextBox txtFieldsCust 
               DataField       =   "AR ORDER Billing Address 2"
               Height          =   285
               Index           =   3
               Left            =   1200
               Locked          =   -1  'True
               TabIndex        =   23
               Top             =   1440
               Width           =   3255
            End
            Begin VB.TextBox txtFieldsCust 
               DataField       =   "AR ORDER Billing City"
               Height          =   285
               Index           =   4
               Left            =   1200
               Locked          =   -1  'True
               TabIndex        =   24
               Top             =   1800
               Width           =   975
            End
            Begin VB.TextBox txtFieldsCust 
               DataField       =   "AR ORDER Billing Postal"
               Height          =   285
               Index           =   6
               Left            =   3720
               Locked          =   -1  'True
               TabIndex        =   26
               Top             =   1800
               Width           =   735
            End
            Begin VB.TextBox txtFieldsCust 
               DataField       =   "AR ORDER Billing State"
               Height          =   285
               Index           =   5
               Left            =   2760
               Locked          =   -1  'True
               TabIndex        =   25
               Top             =   1800
               Width           =   495
            End
            Begin VB.TextBox txtFieldsCust 
               DataField       =   "AR ORDER Billing Country"
               Height          =   285
               Index           =   7
               Left            =   1200
               Locked          =   -1  'True
               TabIndex        =   27
               Top             =   2160
               Width           =   3255
            End
            Begin VB.TextBox txtFieldsCust 
               DataField       =   "AR ORDER Billing Customer"
               Height          =   285
               Index           =   1
               Left            =   1200
               Locked          =   -1  'True
               TabIndex        =   21
               Top             =   720
               Width           =   3255
            End
            Begin VB.TextBox txtFieldsShip 
               DataField       =   "AR ORDER Shipping Country"
               Height          =   285
               Index           =   7
               Left            =   5760
               Locked          =   -1  'True
               TabIndex        =   36
               Top             =   2160
               Width           =   3255
            End
            Begin VB.TextBox txtFieldsShip 
               DataField       =   "AR ORDER Shipping State"
               Height          =   285
               Index           =   5
               Left            =   7320
               Locked          =   -1  'True
               TabIndex        =   34
               Top             =   1800
               Width           =   495
            End
            Begin VB.TextBox txtFieldsShip 
               DataField       =   "AR ORDER Shipping Postal"
               Height          =   285
               Index           =   6
               Left            =   8280
               Locked          =   -1  'True
               TabIndex        =   35
               Top             =   1800
               Width           =   735
            End
            Begin VB.TextBox txtFieldsShip 
               DataField       =   "AR ORDER Shipping City"
               Height          =   285
               Index           =   4
               Left            =   5760
               Locked          =   -1  'True
               TabIndex        =   33
               Top             =   1800
               Width           =   975
            End
            Begin VB.TextBox txtFieldsShip 
               DataField       =   "AR ORDER Shipping Address 2"
               Height          =   285
               Index           =   3
               Left            =   5760
               Locked          =   -1  'True
               TabIndex        =   32
               Top             =   1440
               Width           =   3255
            End
            Begin VB.TextBox txtFieldsCust 
               Height          =   285
               Index           =   10
               Left            =   1440
               Locked          =   -1  'True
               TabIndex        =   129
               Top             =   360
               Width           =   1215
            End
            Begin VB.CommandButton cmdLookupShip 
               Height          =   285
               Left            =   7200
               Picture         =   "frm_AR_Order_Entry.frx":23F3
               Style           =   1  'Graphical
               TabIndex        =   4
               ToolTipText     =   "Get Shipping Place"
               Top             =   360
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.CommandButton cmdShipBack 
               Height          =   285
               Left            =   7200
               Picture         =   "frm_AR_Order_Entry.frx":26FD
               Style           =   1  'Graphical
               TabIndex        =   149
               ToolTipText     =   "Get Shipping Place"
               Top             =   360
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.Label Label1 
               Caption         =   "Document Type:"
               Height          =   255
               Index           =   4
               Left            =   9600
               TabIndex        =   120
               Top             =   960
               Width           =   1335
            End
            Begin VB.Label Label1 
               Caption         =   "Document #:"
               Height          =   255
               Index           =   3
               Left            =   9600
               TabIndex        =   119
               Top             =   1560
               Width           =   1335
            End
            Begin VB.Label Label1 
               Caption         =   " Taxable Subtotal"
               Height          =   255
               Index           =   0
               Left            =   9600
               TabIndex        =   118
               Top             =   360
               Width           =   1335
            End
            Begin VB.Label lblLabels 
               Caption         =   "Telephone:"
               Height          =   255
               Index           =   17
               Left            =   4800
               TabIndex        =   117
               Top             =   2520
               Width           =   975
            End
            Begin VB.Label lblLabels 
               Caption         =   "Facsimile:"
               Height          =   255
               Index           =   16
               Left            =   7080
               TabIndex        =   116
               Top             =   2520
               Width           =   735
            End
            Begin VB.Label lblLabels 
               Caption         =   "Telephone:"
               Height          =   255
               Index           =   15
               Left            =   240
               TabIndex        =   115
               Top             =   2520
               Width           =   975
            End
            Begin VB.Label lblLabels 
               Caption         =   "Facsimile:"
               Height          =   255
               Index           =   2
               Left            =   2520
               TabIndex        =   114
               Top             =   2520
               Width           =   735
            End
            Begin VB.Label lblLabels 
               Caption         =   "  Zip:"
               Height          =   255
               Index           =   19
               Left            =   7800
               TabIndex        =   113
               Top             =   1800
               Width           =   495
            End
            Begin VB.Label lblLabels 
               Caption         =   "  State:"
               Height          =   255
               Index           =   11
               Left            =   6720
               TabIndex        =   112
               Top             =   1800
               Width           =   615
            End
            Begin VB.Label lblLabels 
               Caption         =   "City:"
               Height          =   255
               Index           =   7
               Left            =   4800
               TabIndex        =   111
               Top             =   1800
               Width           =   975
            End
            Begin VB.Label lblLabels 
               Caption         =   "  Zip:"
               Height          =   255
               Index           =   9
               Left            =   3240
               TabIndex        =   110
               Top             =   1800
               Width           =   495
            End
            Begin VB.Label lblLabels 
               Caption         =   "  State:"
               Height          =   255
               Index           =   5
               Left            =   2160
               TabIndex        =   109
               Top             =   1800
               Width           =   615
            End
            Begin VB.Label lblLabels 
               Caption         =   "Address:"
               Height          =   255
               Index           =   8
               Left            =   240
               TabIndex        =   108
               Top             =   1080
               Width           =   975
            End
            Begin VB.Label lblLabels 
               Caption         =   "City:"
               Height          =   255
               Index           =   10
               Left            =   240
               TabIndex        =   107
               Top             =   1800
               Width           =   975
            End
            Begin VB.Label lblLabels 
               Caption         =   "Country:"
               Height          =   255
               Index           =   13
               Left            =   240
               TabIndex        =   106
               Top             =   2160
               Width           =   975
            End
            Begin VB.Label lblLabels 
               Caption         =   "Name:"
               Height          =   255
               Index           =   6
               Left            =   240
               TabIndex        =   105
               Top             =   720
               Width           =   975
            End
            Begin VB.Label lblLabels 
               Caption         =   "Name:"
               Height          =   255
               Index           =   0
               Left            =   4800
               TabIndex        =   104
               Top             =   720
               Width           =   975
            End
            Begin VB.Label lblLabels 
               Caption         =   "Country:"
               Height          =   255
               Index           =   29
               Left            =   4800
               TabIndex        =   103
               Top             =   2160
               Width           =   975
            End
            Begin VB.Label lblLabels 
               Caption         =   "Address:"
               Height          =   255
               Index           =   21
               Left            =   4800
               TabIndex        =   102
               Top             =   1080
               Width           =   975
            End
            Begin VB.Label lblfields 
               Caption         =   "Customer ID:"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   101
               Top             =   360
               Width           =   1035
            End
            Begin VB.Label lblLabels 
               Caption         =   "Shipping ID:"
               Height          =   255
               Index           =   30
               Left            =   4800
               TabIndex        =   100
               Top             =   360
               Width           =   975
            End
         End
         Begin VB.Frame frThird 
            Height          =   855
            Left            =   0
            TabIndex        =   85
            Top             =   4800
            Width           =   11535
            Begin VB.ComboBox cbPurchase 
               DataField       =   "AR ORDER Payment Method"
               Height          =   315
               Index           =   16
               ItemData        =   "frm_AR_Order_Entry.frx":2A07
               Left            =   2520
               List            =   "frm_AR_Order_Entry.frx":2A09
               TabIndex        =   12
               Text            =   "cbPurchase"
               Top             =   440
               Width           =   1335
            End
            Begin VB.CommandButton cmdUpdatedua 
               Height          =   300
               Index           =   16
               Left            =   3840
               Picture         =   "frm_AR_Order_Entry.frx":2A0B
               Style           =   1  'Graphical
               TabIndex        =   13
               ToolTipText     =   "Refresh Payment Methods"
               Top             =   440
               Width           =   375
            End
            Begin VB.ComboBox cbPurchase 
               DataField       =   "AR ORDER Recur Type"
               Height          =   315
               Index           =   15
               ItemData        =   "frm_AR_Order_Entry.frx":2D15
               Left            =   7320
               List            =   "frm_AR_Order_Entry.frx":2D17
               TabIndex        =   16
               Text            =   "cbPurchase"
               Top             =   440
               Width           =   1335
            End
            Begin VB.CommandButton cmdUpdatedua 
               Height          =   300
               Index           =   15
               Left            =   8640
               Picture         =   "frm_AR_Order_Entry.frx":2D19
               Style           =   1  'Graphical
               TabIndex        =   17
               ToolTipText     =   "Refresh Recurring"
               Top             =   440
               Width           =   375
            End
            Begin VB.CommandButton cmdUpdatedua 
               Height          =   300
               Index           =   1
               Left            =   6360
               Picture         =   "frm_AR_Order_Entry.frx":3023
               Style           =   1  'Graphical
               TabIndex        =   15
               ToolTipText     =   "Refresh Sales Tax Group"
               Top             =   440
               Width           =   375
            End
            Begin VB.CommandButton cmdUpdatedua 
               Height          =   280
               Index           =   5
               Left            =   1560
               Picture         =   "frm_AR_Order_Entry.frx":332D
               Style           =   1  'Graphical
               TabIndex        =   11
               ToolTipText     =   "Refresh Payment Terms"
               Top             =   440
               Width           =   375
            End
            Begin VB.CommandButton cmdUpdatedua 
               Height          =   300
               Index           =   2
               Left            =   11040
               Picture         =   "frm_AR_Order_Entry.frx":3637
               Style           =   1  'Graphical
               TabIndex        =   19
               ToolTipText     =   "Refresh  Ship Via"
               Top             =   440
               Width           =   375
            End
            Begin VB.ComboBox cbPurchase 
               DataField       =   "AR ORDER Tax Group"
               Height          =   315
               Index           =   1
               ItemData        =   "frm_AR_Order_Entry.frx":3941
               Left            =   4800
               List            =   "frm_AR_Order_Entry.frx":3943
               TabIndex        =   14
               Text            =   "cbPurchase"
               Top             =   440
               Width           =   1575
            End
            Begin VB.ComboBox cbPurchase 
               DataField       =   "AR ORDER Payment Terms"
               Height          =   315
               Index           =   5
               ItemData        =   "frm_AR_Order_Entry.frx":3945
               Left            =   120
               List            =   "frm_AR_Order_Entry.frx":3947
               TabIndex        =   10
               Text            =   "cbPurchase"
               Top             =   440
               Width           =   1455
            End
            Begin VB.ComboBox cbPurchase 
               DataField       =   "AR ORDER Shipping Method"
               Height          =   315
               Index           =   2
               Left            =   9600
               TabIndex        =   18
               Text            =   "cbPurchase"
               Top             =   440
               Width           =   1455
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
               TabIndex        =   90
               Top             =   195
               Width           =   1695
            End
            Begin VB.Label lblfields 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Recurring "
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   15
               Left            =   7320
               TabIndex        =   89
               Top             =   200
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
               Left            =   9600
               TabIndex        =   88
               Top             =   200
               Width           =   1815
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
               TabIndex        =   87
               Top             =   200
               Width           =   1815
            End
            Begin VB.Label lblfields 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Sales Tax Group"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   1
               Left            =   4800
               TabIndex        =   86
               Top             =   195
               Width           =   1935
            End
         End
         Begin MSDataGridLib.DataGrid grdDataGrid 
            Height          =   825
            Left            =   0
            TabIndex        =   84
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
            Caption         =   "Order Item"
            ColumnCount     =   11
            BeginProperty Column00 
               DataField       =   "AR ORDERD Item Id"
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
               DataField       =   "AR ORDERD Description"
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
               DataField       =   "AR ORDERD Qty"
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
               DataField       =   "AR ORDERD Qty To Invoice"
               Caption         =   "Invoice Qty"
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
               DataField       =   "AR ORDERD Units"
               Caption         =   "Unit"
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
               DataField       =   "AR ORDERD Unit Price"
               Caption         =   "Unit Price"
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
               DataField       =   "AR ORDERD Discount %"
               Caption         =   "Disc"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0%"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column07 
               DataField       =   "AR ORDERD Tax"
               Caption         =   "Tax"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   5
                  Format          =   "0.000E+00"
                  HaveTrueFalseNull=   1
                  TrueValue       =   "Yes"
                  FalseValue      =   "No"
                  NullValue       =   "NA"
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column08 
               DataField       =   "AR ORDERD Item Total"
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
               DataField       =   "AR ORDERD Posting Account"
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
               DataField       =   "AR ORDERD Project"
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
                  ColumnWidth     =   1620.284
               EndProperty
               BeginProperty Column02 
                  Alignment       =   2
                  ColumnWidth     =   900.284
               EndProperty
               BeginProperty Column03 
                  Alignment       =   2
                  ColumnWidth     =   900.284
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   629.858
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   1005.165
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   555.024
               EndProperty
               BeginProperty Column07 
                  ColumnWidth     =   599.811
               EndProperty
               BeginProperty Column08 
                  ColumnWidth     =   1005.165
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
         Begin VB.Frame frSecond 
            Height          =   855
            Left            =   0
            TabIndex        =   121
            Top             =   3000
            Width           =   11535
            Begin VB.TextBox txtfields 
               DataField       =   "AR ORDER Check Date"
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
               TabIndex        =   134
               Top             =   480
               Width           =   1455
            End
            Begin VB.CommandButton cmdDate 
               Enabled         =   0   'False
               Height          =   285
               Index           =   0
               Left            =   11040
               Picture         =   "frm_AR_Order_Entry.frx":3949
               Style           =   1  'Graphical
               TabIndex        =   133
               ToolTipText     =   "Get Order Date"
               Top             =   480
               Width           =   375
            End
            Begin VB.CommandButton cmdbankAccount 
               Height          =   285
               Left            =   6360
               Picture         =   "frm_AR_Order_Entry.frx":3C53
               Style           =   1  'Graphical
               TabIndex        =   9
               ToolTipText     =   "Get Bank Account"
               Top             =   480
               Width           =   375
            End
            Begin VB.TextBox txtfields 
               DataField       =   "AR ORDER Check Acct ID"
               Height          =   285
               Index           =   35
               Left            =   4800
               TabIndex        =   8
               Top             =   480
               Width           =   1575
            End
            Begin VB.TextBox txtfields 
               Alignment       =   2  'Center
               DataField       =   "AR ORDER Check Number"
               Enabled         =   0   'False
               Height          =   285
               Index           =   34
               Left            =   7320
               TabIndex        =   49
               Top             =   480
               Width           =   1695
            End
            Begin VB.CommandButton cmdDate 
               Height          =   285
               Index           =   1
               Left            =   3840
               Picture         =   "frm_AR_Order_Entry.frx":3F5D
               Style           =   1  'Graphical
               TabIndex        =   7
               ToolTipText     =   "Get Due Date"
               Top             =   480
               Width           =   375
            End
            Begin VB.TextBox txtfields 
               DataField       =   "AR ORDER Due Date"
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
               TabIndex        =   43
               Top             =   480
               Width           =   1335
            End
            Begin VB.TextBox txtfields 
               DataField       =   "AR ORDER Ship Date"
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
               TabIndex        =   42
               Top             =   480
               Width           =   1455
            End
            Begin VB.CommandButton cmdDate 
               Height          =   285
               Index           =   20
               Left            =   1560
               Picture         =   "frm_AR_Order_Entry.frx":4267
               Style           =   1  'Graphical
               TabIndex        =   6
               ToolTipText     =   "Get Ship Date"
               Top             =   480
               Width           =   375
            End
            Begin VB.Label lblfields 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Payment/Check Date"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   7
               Left            =   9600
               TabIndex        =   135
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
               TabIndex        =   125
               Top             =   240
               Width           =   1935
            End
            Begin VB.Label lblfields 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Payment/Check No."
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   34
               Left            =   7320
               TabIndex        =   124
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
               TabIndex        =   123
               Top             =   240
               Width           =   1695
            End
            Begin VB.Label lblfields 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Ship Date"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   20
               Left            =   120
               TabIndex        =   122
               Top             =   240
               Width           =   1815
            End
         End
      End
      Begin VB.Frame frAdvance 
         Height          =   8355
         Left            =   11640
         TabIndex        =   67
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
            TabIndex        =   143
            Top             =   7980
            Width           =   3135
            Begin VB.CommandButton cmdNext 
               Height          =   300
               Left            =   2300
               Picture         =   "frm_AR_Order_Entry.frx":4571
               Style           =   1  'Graphical
               TabIndex        =   147
               ToolTipText     =   "Move Forward"
               Top             =   0
               UseMaskColor    =   -1  'True
               Width           =   345
            End
            Begin VB.CommandButton cmdLast 
               Height          =   300
               Left            =   2640
               Picture         =   "frm_AR_Order_Entry.frx":48B3
               Style           =   1  'Graphical
               TabIndex        =   146
               ToolTipText     =   "End Of Record"
               Top             =   0
               UseMaskColor    =   -1  'True
               Width           =   345
            End
            Begin VB.CommandButton cmdPrevious 
               Height          =   300
               Left            =   460
               Picture         =   "frm_AR_Order_Entry.frx":4BF5
               Style           =   1  'Graphical
               TabIndex        =   145
               ToolTipText     =   "Move Previous"
               Top             =   0
               UseMaskColor    =   -1  'True
               Width           =   345
            End
            Begin VB.CommandButton cmdFirst 
               Height          =   300
               Left            =   120
               Picture         =   "frm_AR_Order_Entry.frx":4F37
               Style           =   1  'Graphical
               TabIndex        =   144
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
               TabIndex        =   148
               Top             =   0
               Width           =   1515
            End
         End
         Begin VB.Frame Frame3 
            Height          =   1455
            Left            =   120
            TabIndex        =   127
            Top             =   120
            Width           =   3135
            Begin VB.CommandButton cmdUnPostedDoc 
               Caption         =   "Open"
               Height          =   405
               Left            =   840
               Picture         =   "frm_AR_Order_Entry.frx":5279
               TabIndex        =   63
               ToolTipText     =   "Load Unposted Document"
               Top             =   900
               Width           =   735
            End
            Begin VB.CommandButton cmdPostDoc 
               Caption         =   "Posted"
               Height          =   405
               Left            =   120
               Picture         =   "frm_AR_Order_Entry.frx":5728
               TabIndex        =   62
               ToolTipText     =   "Load Posted Document"
               Top             =   900
               Width           =   735
            End
            Begin VB.CommandButton Command1 
               Height          =   495
               Left            =   2400
               Picture         =   "frm_AR_Order_Entry.frx":5CC1
               Style           =   1  'Graphical
               TabIndex        =   61
               ToolTipText     =   "Search All Record"
               Top             =   240
               Width           =   615
            End
            Begin VB.OptionButton optOrderType 
               Caption         =   "Sales Order"
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   0
               Left            =   1800
               TabIndex        =   64
               ToolTipText     =   "Load Sales Order"
               Top             =   840
               Value           =   -1  'True
               Width           =   1215
            End
            Begin VB.OptionButton optOrderType 
               Caption         =   "Back Order"
               Height          =   255
               Index           =   1
               Left            =   1800
               TabIndex        =   65
               ToolTipText     =   "Load Back Order"
               Top             =   1080
               Width           =   1215
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
               TabIndex        =   59
               Top             =   480
               Width           =   1455
            End
            Begin VB.CommandButton cmdSearch 
               Height          =   495
               Left            =   1680
               Picture         =   "frm_AR_Order_Entry.frx":5FCB
               Style           =   1  'Graphical
               TabIndex        =   60
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
               Caption         =   "Order No"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   18
               Left            =   120
               TabIndex        =   128
               Top             =   240
               Width           =   1455
            End
         End
         Begin VB.Frame frButton 
            Height          =   1815
            Left            =   120
            TabIndex        =   68
            Top             =   6120
            Width           =   3135
            Begin VB.CommandButton cmdCreateInvoice 
               Caption         =   "&Invoice"
               Height          =   780
               Left            =   2040
               Picture         =   "frm_AR_Order_Entry.frx":62D5
               Style           =   1  'Graphical
               TabIndex        =   51
               ToolTipText     =   "Create Invoice"
               Top             =   240
               Width           =   975
            End
            Begin VB.CommandButton cmdApprove 
               Caption         =   "Appro&ved"
               Height          =   780
               Left            =   1080
               Picture         =   "frm_AR_Order_Entry.frx":65DF
               Style           =   1  'Graphical
               TabIndex        =   142
               ToolTipText     =   "Approved Current Document"
               Top             =   240
               Width           =   975
            End
            Begin VB.CommandButton cmdPrint 
               Caption         =   "&Print"
               Height          =   780
               Left            =   120
               Picture         =   "frm_AR_Order_Entry.frx":68E9
               Style           =   1  'Graphical
               TabIndex        =   50
               ToolTipText     =   "Print Sales Order"
               Top             =   240
               Width           =   975
            End
            Begin VB.CommandButton cmdCancel 
               Caption         =   "Canc&el"
               Height          =   300
               Left            =   2040
               TabIndex        =   57
               ToolTipText     =   "Cancell Current Process"
               Top             =   1440
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.CommandButton cmdClose 
               Caption         =   "&Close"
               Height          =   300
               Left            =   1080
               TabIndex        =   56
               ToolTipText     =   "Close Order Transaction"
               Top             =   1440
               Width           =   975
            End
            Begin VB.CommandButton cmdRefresh 
               Caption         =   "&Refresh"
               Height          =   300
               Left            =   120
               TabIndex        =   55
               ToolTipText     =   "Refresh All"
               Top             =   1440
               Width           =   975
            End
            Begin VB.CommandButton cmdDelete 
               Caption         =   "&Delete"
               Height          =   300
               Left            =   2040
               TabIndex        =   54
               ToolTipText     =   "Delete Current Record"
               Top             =   1080
               Width           =   975
            End
            Begin VB.CommandButton cmdUpdate 
               Caption         =   "&Update"
               Height          =   300
               Left            =   1080
               TabIndex        =   53
               ToolTipText     =   "Update Current Record"
               Top             =   1080
               Width           =   975
            End
            Begin VB.CommandButton cmdAdd 
               Caption         =   "&Add"
               Height          =   300
               Left            =   120
               TabIndex        =   52
               ToolTipText     =   "Adding New Data"
               Top             =   1080
               Width           =   975
            End
            Begin VB.CommandButton cmdInvoiceBackOrder 
               Caption         =   "&Invoice"
               Height          =   780
               Left            =   2160
               Picture         =   "frm_AR_Order_Entry.frx":6BF3
               Style           =   1  'Graphical
               TabIndex        =   69
               ToolTipText     =   "Create Invoice"
               Top             =   240
               Visible         =   0   'False
               Width           =   855
            End
         End
         Begin VB.PictureBox PcData 
            BorderStyle     =   0  'None
            Height          =   2535
            Left            =   120
            ScaleHeight     =   2535
            ScaleWidth      =   3135
            TabIndex        =   70
            Top             =   3600
            Width           =   3135
            Begin VB.TextBox txtfields 
               DataField       =   "AR ORDER Tax Freight"
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
               TabIndex        =   140
               Top             =   1080
               Width           =   495
            End
            Begin VB.CommandButton cmdCreditLimit 
               Height          =   285
               Left            =   1200
               Picture         =   "frm_AR_Order_Entry.frx":7035
               Style           =   1  'Graphical
               TabIndex        =   139
               Top             =   2160
               Width           =   375
            End
            Begin VB.CheckBox chkTaxFreight 
               Alignment       =   1  'Right Justify
               Caption         =   "Shipping Charge:"
               DataField       =   "AR ORDER Tax Freight"
               Height          =   255
               Left            =   1200
               TabIndex        =   47
               Top             =   1080
               Width           =   255
            End
            Begin VB.TextBox txtfields 
               DataField       =   "AR ORDER Subtotal"
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
               TabIndex        =   75
               Top             =   0
               Width           =   1935
            End
            Begin VB.TextBox txtfields 
               DataField       =   "AR ORDER Discount Amount"
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
               TabIndex        =   74
               Top             =   360
               Width           =   1335
            End
            Begin VB.TextBox txtfields 
               DataField       =   "AR ORDER Sales Tax"
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
               TabIndex        =   73
               Top             =   720
               Width           =   1335
            End
            Begin VB.TextBox txtfields 
               DataField       =   "AR ORDER Freight"
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
               Left            =   1680
               Locked          =   -1  'True
               TabIndex        =   45
               Top             =   1080
               Width           =   1335
            End
            Begin VB.TextBox txtfields 
               DataField       =   "AR ORDER Discount Percent"
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
               TabIndex        =   44
               Top             =   360
               Width           =   495
            End
            Begin VB.TextBox txtfields 
               DataField       =   "AR ORDER Tax Percent"
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
               Locked          =   -1  'True
               TabIndex        =   46
               Top             =   720
               Width           =   495
            End
            Begin VB.TextBox txtfields 
               DataField       =   "AR ORDER Total"
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
               TabIndex        =   72
               Top             =   1440
               Width           =   1935
            End
            Begin VB.TextBox txtfields 
               DataField       =   "AR ORDER Amount Paid"
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
               TabIndex        =   48
               Top             =   2160
               Width           =   1215
            End
            Begin VB.TextBox txtfields 
               DataField       =   "AR ORDER Balance Due"
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
               TabIndex        =   71
               Top             =   2160
               Width           =   1335
            End
            Begin VB.Label lblLabels 
               Caption         =   "Shipping:"
               Height          =   255
               Index           =   3
               Left            =   0
               TabIndex        =   141
               Top             =   1080
               Width           =   1095
            End
            Begin VB.Label lblLabels 
               Caption         =   "Sub Total:"
               Height          =   255
               Index           =   24
               Left            =   0
               TabIndex        =   81
               Top             =   0
               Width           =   1095
            End
            Begin VB.Label lblLabels 
               Caption         =   "Discount:"
               Height          =   255
               Index           =   22
               Left            =   0
               TabIndex        =   80
               Top             =   360
               Width           =   1095
            End
            Begin VB.Label lblLabels 
               Caption         =   "Sales Tax:"
               Height          =   255
               Index           =   23
               Left            =   0
               TabIndex        =   79
               Top             =   720
               Width           =   1095
            End
            Begin VB.Label lblLabels 
               Caption         =   "Total Amount:"
               Height          =   255
               Index           =   25
               Left            =   0
               TabIndex        =   78
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
               TabIndex        =   77
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
               TabIndex        =   76
               Top             =   1920
               Width           =   1335
            End
         End
         Begin VB.TextBox txtfields 
            DataField       =   "AR ORDER Description"
            Height          =   1515
            Index           =   32
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   20
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
            TabIndex        =   82
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
      Caption         =   "Sales/Delivery Order"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   126
      Top             =   120
      Width           =   9225
   End
End
Attribute VB_Name = "frm_AR_Order_Entry"
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
Dim CustID0 As String
Dim ShipID0 As String
Dim BankAcct35 As String

Public grdOnAddNew As Boolean

Private Function Datavalidate() As Boolean

'  If Me![AR ORDER Amount Paid] > 0 Then 'Or Me![AR ORDER Document Type] = "Credit Memo" Then
  If txtFields(36) > 0 Then 'Or Me![AR ORDER Document Type] = "Credit Memo" Then
    If txtFields(35) = "" Then
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
  
  If CountRecord("[AR ORDERD Item ID]", "[AR Order Detail]", db, "[AR ORDERD Document #] = " & txtFields(2)) <= 0 Then
    MsgBox "Must have at least one inventory item!", , "Error"
    Datavalidate = False
    Exit Function
  End If

  'Check for balance due < 0 and check number
  If txtFields(36) > 0 Then
    If txtFields(33) < 0 Then
      MsgBox "Amount paid cannot exceed invoice total!", , "Error"
      Datavalidate = False
      Exit Function
    End If
    If Trim(txtFields(34)) = "" Then
      MsgBox "You must enter a check number!", , "Error"
      Datavalidate = False
      Exit Function
    End If
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
    Case "satu"
        ComboInit cbPurchase(1), lblfields(1), "SELECT [SYS TAXGRP ID] FROM [SYS Tax Group]", db
    Case "dua"
        ComboInit cbPurchase(2), lblfields(2), "SELECT [LIST SHIP Method] FROM [LIST Shipping Methods]", db
    Case "tiga"
        ComboInit cbPurchase(3), lblfields(3), "SELECT [LIST PAY Method] FROM [LIST Payment Methods]", db
    Case "lima"
        ComboInit cbPurchase(5), lblfields(5), "SELECT [LIST PAY Description] FROM [LIST Payment Terms]", db
    Case "limabelas"
        ComboInit cbPurchase(15), lblfields(15), "SELECT [RECURR TYPE] FROM [RECUR_TYPE]", db
    Case "enambelas"
        ComboInit cbPurchase(16), lblfields(16), "SELECT [LIST PAY Method] FROM [LIST Payment Methods]", db
    Case "semua"
        'ComboInit cbhelp, lblhelp, "SELECT [Form Name] FROM [Help Text]"
        ComboInit cbPurchase(1), lblfields(1), "SELECT [SYS TAXGRP ID] FROM [SYS Tax Group]", db
        ComboInit cbPurchase(2), lblfields(2), "SELECT [LIST SHIP Method] FROM [LIST Shipping Methods]", db
        ComboInit cbPurchase(5), lblfields(5), "SELECT [LIST PAY Description] FROM [LIST Payment Terms]", db
        ComboInit cbPurchase(15), lblfields(15), "SELECT [RECURR TYPE] FROM [RECUR_TYPE]", db
        ComboInit cbPurchase(16), lblfields(16), "SELECT [LIST PAY Method] FROM [LIST Payment Methods]", db
End Select
End Sub

Private Sub cbPurchase_LostFocus(Index As Integer)
On Error GoTo exit_EditMode
If ADOprimaryrs.EditMode = adEditInProgress Then
    If IsNull(ADOprimaryrs![AR ORDER Status]) Or ADOprimaryrs![AR ORDER Status] <> "Open" Then
        ValidatePower txtFields(12).Text, "Edit", DocType, db
        ADOprimaryrs![AR ORDER Status] = "Open"
        ADOprimaryrs.Update
        cmdApprove.Picture = fMainForm.imlIcons.ListImages("Locked").Picture
    End If
End If
exit_EditMode:

ShowStatus True
Select Case Index
Case 1
   CheckCombo cbPurchase(Index), "[SYS TAXGRP ID]", "[SYS Tax Group]", db, True
   CalcTotals
Case 2
   CheckCombo cbPurchase(Index), "[LIST SHIP Method]", "[LIST Shipping Methods]", db, True
   CalcTotals
Case 5
   CheckCombo cbPurchase(Index), "[LIST PAY Description]", "[LIST Payment Terms]", db, True
   SetDueDate
   ComboInit cbPurchase(16), lblfields(16), "SELECT [LIST PAY Method] FROM [LIST Payment Methods] WHERE [Payment Terms]='" & cbPurchase(5).Text & "'", db
Case 15
   CheckCombo cbPurchase(Index), "[RECURR TYPE]", "[RECUR_TYPE]", db, True
Case 16
   CheckCombo cbPurchase(Index), "[LIST PAY Method]", "[LIST Payment Methods]", db, True
End Select
ShowStatus False
End Sub

Private Sub cmdUpdatedua_Select(Index)
    Select Case Index
    Case 1
        loadCombo "satu"
    Case 2
        loadCombo "dua"
    Case 3
        loadCombo "tiga"
    Case 5
        loadCombo "lima"
    Case 15
        loadCombo "limabelas"
    Case 16
        loadCombo "enambelas"
    End Select
End Sub

Private Sub cmdApprove_Click()
Dim Approve As Boolean
If ADOprimaryrs.EditMode = adEditAdd Then Exit Sub
If Datavalidate = False Then Exit Sub
If Not CheckEmpty Then Exit Sub

If ADOprimaryrs![AR ORDER Status] = "Open" Or IsNull(ADOprimaryrs![AR ORDER Status]) Then
    Approve = ValidatePower(txtFields(12).Text, "Approve", DocType, db)
    If Approve = True Then
        ADOprimaryrs![AR ORDER Status] = "Ready"
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
    txtFields(35).SetFocus
Else
    'Me.PopupMenu fMainForm.mnuAccount
End If
End Sub

Private Sub cmdBankAccount_Select()
    Dim SQLstatement As String
    Dim No As Integer
    Dim ghead As String
    Dim fhead As String

    No = 1220
    SQLstatement = "select [BANK ACCT ID], [BANK ACCT Name]" & _
                    "from [BANK Accounts]"
    ghead = "Account Description"
    fhead = "Account No//Description"
    
    AllLookup.GetWhichTable No, SQLstatement, ghead, fhead, db
    'AllLookup.Show vbModal
    'ADOprimaryrs![AR ORDER Check Acct ID] = txtFields(35).Text
End Sub

Private Sub cmdCreditLimit_Click()
Dim CurrRequest As Currency
If mbAddNewFlag Then
    CurrRequest = 0
Else
    CurrRequest = CCur(txtFields(33).Text)
End If
If Trim(txtFieldsCust(0)) <> "" Then
    CheckCreditLimit CurrRequest, txtFieldsCust(0), db, True
End If
End Sub

Private Sub cmdLookupCust_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    cmdLookupCust_Select
    CustID0 = txtFieldsCust(0).Text
Else
    'Me.PopupMenu mnuCustomer
End If

End Sub

Private Sub cmdLookupCust_Select()
    If Me.cmdLookupShip.Visible = False Then
       'AllLookup.ToWhichRecord adoPrimaryRS, "Quote", "Customer Name//SalesPerson//Quote Value//Order Date", "[AR ORDER Bill To]//[AR ORDER Salesperson]//[AR ORDER Total]//[AR ORDER Date]"
       AllLookup.ToWhichRecord ADOprimaryrs, "Quote", "-//Order No//-//-//-//-//-//-//Cust. PO//-//-//-//-//Customer Name"
       'AllLookup.Show vbModal
       GetTextColor Me
    Else
       Dim i As Integer
       AllLookup.GetWhichTable 1000, "Select [AR CUST Customer ID],[AR CUST Name],[AR CUST Address 1]," & _
       "[AR CUST Address 2],[AR CUST City],[AR CUST State],[AR CUST Postal]," & _
       "[AR CUST Country],[AR CUST Phone],[AR CUST Fax],[AR CUST SalesPerson] FROM [AR Customer] ", "Customer Particular", _
       "Customer ID//Customer Name//Address 1//Address 2//City//State//Postal//Country", db
       'AllLookup.Show vbModal
        If CustID0 <> txtFieldsCust(0).Text And txtFieldsCust(0).Text <> "" Then
            CustomerData Me, db, txtFieldsCust(0).Text, True
            For i = 0 To txtFieldsShip.UBound
                txtFieldsShip(i).Text = ""
            Next
            shipToID db, Me
            cbPurchase_LostFocus 1
            cbPurchase_LostFocus 5
        End If
        End If
End Sub

Private Sub cmdLookupShip_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    cmdLookupShip_Select
    ShipID0 = txtFieldsShip(0)
Else
    'Me.PopupMenu mnuShipping
End If
End Sub

Private Sub cmdLookupShip_Select()
   AllLookup.GetWhichTable 1001, "Select [AR SHIP ID],[AR SHIP Name],[AR SHIP Address 1]," & _
   "[AR SHIP Address 2],[AR SHIP City],[AR SHIP State],[AR SHIP Postal],[AR SHIP Country],[AR SHIP Phone],[AR SHIP Fax] From " & _
   "[AR SHIP to] WHERE [AR SHIP Customer ID]='" & txtFieldsCust(0) & "' ", "Shipping Address", _
   "Shipping ID//Place Name//Address 1//Address 2//City//State//Postal//Country", db
   'AllLookup.Show vbModal
End Sub

Private Sub cmdPostDoc_Click()
    ClearDatasource
    RSstatement = "SHAPE {select * from [AR Order] WHERE [AR ORDER Document Type]='" & DocType & "' AND [AR ORDER Invoiced]=True} AS ParentCMD APPEND ({select * from [AR Order Detail] } AS ChildCMD RELATE [AR ORDER Document #] TO [AR ORDERD Document #]) AS ChildCMD"
    OpenDB RSstatement
    picStatBox.Enabled = True
    lblStatus.BackColor = vbWhite
    ProcessDoneMusic "Done"
End Sub

Private Sub cmdPrint_Click()
    If ADOprimaryrs![AR ORDER Status] = "Open" Then
        MsgBox "This " & DocType & " has not been approved."
        Exit Sub
    End If
    
    Dim frm As New frm_prnPreview
    frm.Record txtFields(12).Text, Me.Name, DocType
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
    SearchRECORD ADOprimaryrs, grdDataGrid, txtFields(0).Text, lblfields(18).Caption, "AR ORDER Ext Document #", "AR ORDER Ext Document #"
    ProcessDoneMusic "Done"
End If
End Sub

Private Sub cmdShipBack_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    cmdLookupShip_Select
    ShipID0 = txtFieldsShip(0)
Else
    'Me.PopupMenu mnuShipping
End If

End Sub

Private Sub cmdUnPostedDoc_Click()

    ClearDatasource
    RSstatement = "SHAPE {select * from [AR Order] WHERE [AR ORDER Document Type]='" & DocType & "' AND [AR ORDER Invoiced]=False} AS ParentCMD APPEND ({select * from [AR Order Detail] } AS ChildCMD RELATE [AR ORDER Document #] TO [AR ORDERD Document #]) AS ChildCMD"
    OpenDB RSstatement
    picStatBox.Enabled = True
    lblStatus.BackColor = vbWhite
    ProcessDoneMusic "Done"
End Sub

Private Sub chkTaxFreight_Click()
   If chkTaxFreight.Value = 0 Then
        chkTaxFreight.Caption = "No Shipping Tax"
   ElseIf chkTaxFreight.Value = 1 Then
        chkTaxFreight.Caption = "Got Shipping Tax"
   End If
End Sub

Private Sub cmdCreateInvoice_Click()
Dim DataB As ADODB.Connection
  
  'On Error GoTo cmdCreateInvoice_Error
    If ADOprimaryrs![AR ORDER Status] = "Open" Then
        MsgBox "This " & DocType & " has not been approved."
        Exit Sub
    End If
    If Datavalidate = False Then Exit Sub
    If Not CheckEmpty Then Exit Sub
  
  'update to the database the current data
  cmdUpdate_Click
  '------------------data validation----------------
  'if there is no item data... no sale order
    If txtFields(24) = 0 Then
        MsgBox "No Item to sell... please select the product to sell", vbExclamation, "Error Invoice Creation"
        Exit Sub
    End If
  If ADOprimaryrs![AR ORDER Invoiced] = True Then '
        MsgBox "Duplicate Invoice is not allowed", vbExclamation, "Error Invoice Creation"
        Exit Sub
  End If
      
  If CheckCreditLimit(txtFields(33).Text, txtFieldsCust(0), db) = False Then Exit Sub
 
  '---------------end of data validation----------------------
  
  ShowStatus True
  Dim iSuccess As Integer
  Dim iResponse As Integer
  iResponse = MsgBox("Are you sure you want to create an invoice from this " & DocType & "?", vbYesNo, "Create Invoice")
  If iResponse = vbNo Then
    ShowStatus False
    Exit Sub
  End If

  
  Set DataB = New ADODB.Connection
  DataB.CursorLocation = adUseServer
  DataB.Open gblADOProvider
   
  'Create and post an invoice based on the current Invoice Qty's for this order
  Dim rsSales As ADODB.Recordset
  Dim rsSalesDetail As ADODB.Recordset
  'Dim rsOrders As ADODB.Recordset
  Dim rsOrderDetail As ADODB.Recordset
  'Dim rsOrderTracking As ADODB.Recordset
    
  Set rsSales = New ADODB.Recordset
  rsSales.Open "Select * from [AR Sales]", DataB, adOpenKeyset, adLockOptimistic, adCmdText
        
  'Now do the line items
  Set rsOrderDetail = New ADODB.Recordset
  rsOrderDetail.Open ("SELECT * FROM [AR ORDER Detail] where [AR ORDERD Qty To Invoice] > 0 " & _
  "AND [AR ORDERD Document #] = " & txtFields(2)), DataB, adOpenKeyset, adLockReadOnly, adCmdText
  'Debug.Print ("SELECT * FROM [AR ORDER Detail] where [AR ORDERD Qty To Invoice] > 0 AND [AR ORDERD Document #] = " & txtfields(2))
  'If rsOrderDetail.RecordCount = 0 Then
  '    MsgBox "There is no quantity to invoice", vbCritical, "Error Creation"
  '    ShowStatus False
  '    Exit Sub
  'End If
  
  If rsOrderDetail.RecordCount <= 0 Then
    MsgBox rsOrderDetail.RecordCount & "  There is nothing to sell (Invoice quantity may empty) " & txtFields(2)
    On Error Resume Next
    ShowStatus False
    GoTo Out_Of_Here
  End If
  
  rsOrderDetail.MoveFirst
  With ADOprimaryrs
  
  'Set rsOrders = New ADODB.Recordset
  'rsOrders.Open ("[AR Order]"), DataB, adOpenStatic, adLockOptimistic
  'rsOrders.Index = "PrimaryKey"
  'rsOrders.MoveFirst
  'rsOrders.Find "[AR ORDER Document #]=" & ![AR ORDER Document #]
  'If rsOrders.EOF Then
    'This is a problem
  'End If
  
  Dim lSalesID As Long
  Dim sExtDocument As String
  
  'Build an Ext Document number for this invoice
  Dim iInvoiceCount As Integer
  iInvoiceCount = CountRecord("[Order ID]", "[AR Order Tracking]", db, "[Order ID] = " & ![AR ORDER Document #])
  
    ' Change mouse pointer to hourglass.
   
FindDocument:
  If iInvoiceCount = 0 Then
    sExtDocument = ![AR ORDER Ext Document #]
  Else
    sExtDocument = ![AR ORDER Ext Document #] & "-" & Trim(CStr(iInvoiceCount))
  End If
  
  'Make sure document does not already exist
  'rsSales.Index = "Ext Document #"
  If rsSales.RecordCount > 0 Then
    rsSales.MoveFirst
    rsSales.Find "[AR SALE Ext Document #]='" & sExtDocument & "'"
    If rsSales.EOF Then
      'This is a good thing
    Else
      iInvoiceCount = iInvoiceCount + 1
      GoTo FindDocument
    End If
  End If
     
  With rsSales
   .AddNew
    
    ![AR SALE Ext Document #] = "[" & sExtDocument & "]"
    ![AR SALE Document Type] = "Invoice"
    ![AR SALE Saved YN] = False
    ![AR SALE Status] = "Open"
    ![AR SALE Posted YN] = False
    ![AR SALE Service YN] = False
    ![AR SALE Date] = ADOprimaryrs![AR ORDER Date]
    ![AR SALE Due Date] = ADOprimaryrs![AR ORDER Due Date]
    '![AR SALE Quote Document #] = ADOprimaryrs![AR ORDER Quote Document]
    If Not (IsNull(Trim(ADOprimaryrs![AR ORDER Quote Document #]))) Then
        ![AR SALE Quote Document #] = CStr(ADOprimaryrs![AR ORDER Quote Document #])
    End If
    ![AR SALE PO ID] = CStr(ADOprimaryrs![AR ORDER PO ID])
    If Not (IsNull(Trim(ADOprimaryrs![AR ORDER Tax Exempt ID]))) Then
    '    ![AR SALE Tax Exempt ID] = " "
    'Else
        ![AR SALE Tax Exempt ID] = CStr(ADOprimaryrs![AR ORDER Tax Exempt ID])
    End If
    '("AR SALE Tax Group") = CStr(adoPrimaryRS("AR ORDER Tax Group"))
    ![AR SALE Tax Group] = CStr(ADOprimaryrs![AR ORDER Tax Group])
    ![AR SALE Customer ID] = CStr(ADOprimaryrs![AR ORDER Customer ID])
    '![AR SALE Bill To] = CStr(ADOprimaryrs![AR ORDER Bill To]) & ""
    'If Not (IsNull(Trim(ADOprimaryrs![AR ORDER Billing Customer]))) Then
    '    ![AR SALE Billing Customer] = " "
    'Else
    ![AR SALE Billing Customer] = CStr(ADOprimaryrs![AR ORDER Billing Customer])
    'End If
    ![AR SALE Billing Address 1] = CStr(ADOprimaryrs![AR ORDER Billing Address 1])
    If Not (IsNull(Trim(ADOprimaryrs![AR ORDER Billing Address 2]))) Then
    '    ![AR SALE Billing Address 2] = " "
    'Else
        ![AR SALE Billing Address 2] = CStr(ADOprimaryrs![AR ORDER Billing Address 2])
    End If
    ![AR SALE Billing City] = CStr(ADOprimaryrs![AR ORDER Billing City])
    ![AR SALE Billing State] = CStr(ADOprimaryrs![AR ORDER Billing State])
    ![AR SALE Billing Postal] = CStr(ADOprimaryrs![AR ORDER Billing Postal])
    ![AR SALE Billing Country] = CStr(ADOprimaryrs![AR ORDER Billing Country])
    If Not (IsNull(Trim(ADOprimaryrs![AR ORDER Billing Phone]))) Then
        ![AR SALE Billing Phone] = CStr(ADOprimaryrs![AR ORDER Billing Phone])
    End If
    If Not (IsNull(Trim(ADOprimaryrs![AR ORDER Billing Fax]))) Then
        ![AR SALE Billing fax] = CStr(ADOprimaryrs![AR ORDER Billing Fax])
    End If
    ![AR SALE Shipping Method] = CStr(ADOprimaryrs![AR ORDER Shipping Method])
    ![AR SALE Ship Date] = ADOprimaryrs![AR ORDER Ship Date]
    ![AR SALE Shipping ID] = CStr(ADOprimaryrs![AR ORDER Shipping ID])
    If Not (IsNull(Trim(ADOprimaryrs![AR ORDER Ship To]))) Then
    '    ![AR SALE Shipping Customer] = " "
    'Else
        ![AR SALE Ship To] = CStr(ADOprimaryrs![AR ORDER Ship To])
    End If
    
    If Not (IsNull(Trim(ADOprimaryrs![AR ORDER Shipping Customer]))) Then
    '    ![AR SALE Shipping Customer] = " "
    'Else
        ![AR SALE Shipping Customer] = CStr(ADOprimaryrs![AR ORDER Shipping Customer])
    End If
    '![AR SALE Shipping Address 1] = CStr(ADOprimaryrs![AR ORDER Shipping Address 1])
    If Not (IsNull(Trim(ADOprimaryrs![AR ORDER Shipping Address 1]))) Then
    '    ![AR SALE Shipping Address 1] = " "
    'Else
        ![AR SALE Shipping Address 1] = CStr(ADOprimaryrs![AR ORDER Shipping Address 1])
    End If
    If Not (IsNull(Trim(ADOprimaryrs![AR ORDER Shipping Address 2]))) Then
    '    ![AR SALE Shipping Address 2] = " "
    'Else
        ![AR SALE Shipping Address 2] = CStr(ADOprimaryrs![AR ORDER Shipping Address 2])
    End If
    If Not (IsNull(Trim(ADOprimaryrs![AR ORDER Shipping City]))) Then
        ![AR SALE Shipping City] = CStr(ADOprimaryrs![AR ORDER Shipping City])
    End If
    If Not (IsNull(Trim(ADOprimaryrs![AR ORDER Shipping State]))) Then
        ![AR SALE Shipping State] = CStr(ADOprimaryrs![AR ORDER Shipping State])
    End If
    If Not (IsNull(Trim(ADOprimaryrs![AR ORDER Shipping Postal]))) Then
    ![AR SALE Shipping Postal] = CStr(ADOprimaryrs![AR ORDER Shipping Postal])
    End If
    If Not (IsNull(Trim(ADOprimaryrs![AR ORDER Shipping Country]))) Then
        ![AR SALE Shipping Country] = CStr(ADOprimaryrs![AR ORDER Shipping Country])
    End If
    If Not (IsNull(Trim(ADOprimaryrs![AR ORDER Shipping Phone]))) Then
        ![AR SALE Shipping Phone] = CStr(ADOprimaryrs![AR ORDER Shipping Phone])
    End If
    If Not (IsNull(Trim(ADOprimaryrs![AR ORDER Shipping Fax]))) Then
        ![AR SALE Shipping Fax] = CStr(ADOprimaryrs![AR ORDER Shipping Fax])
    End If
    ![AR SALE Salesperson] = CStr(ADOprimaryrs![AR ORDER Salesperson])
    ![AR SALE Payment Terms] = CStr(ADOprimaryrs![AR ORDER Payment Terms])
    ![AR SALE Payment Method] = CStr(ADOprimaryrs![AR ORDER Payment Method])
    
    ![AR SALE Subtotal] = ADOprimaryrs![AR ORDER Subtotal]
    ![AR SALE Discount Percent] = ADOprimaryrs![AR ORDER Discount Percent]
    ![AR SALE Discount Amount] = ADOprimaryrs![AR ORDER Discount Amount] '
    ![AR SALE Tax Percent] = ADOprimaryrs![AR ORDER Tax Percent]
    
    ![AR SALE Sales Tax] = ADOprimaryrs![AR ORDER Sales Tax]
    ![AR SALE Freight] = ADOprimaryrs![AR ORDER Freight]
    ![AR SALE Total] = ADOprimaryrs![AR ORDER Total]
    ![AR SALE Amount Paid] = ADOprimaryrs![AR ORDER Amount Paid]
    ![AR SALE Balance Due] = ADOprimaryrs![AR ORDER Balance Due]
    
    ![AR SALE Undistributed Amt] = ADOprimaryrs![AR ORDER Undistributed Amt]
    ![AR SALE Commission] = ADOprimaryrs![AR ORDER Commission]
    ![AR SALE Commissionable Sales] = ADOprimaryrs![AR ORDER Commissionable Sales]
    If Not (IsNull(Trim(ADOprimaryrs![AR ORDER Description]))) Then
    '    ![AR SALE Description] = " "
    'Else
        ![AR SALE Description] = CStr(ADOprimaryrs![AR ORDER Description])
    End If
    If Not (IsNull(Trim(ADOprimaryrs![AR ORDER Message]))) Then
    '    ![AR SALE Message] = " "
    'Else
        ![AR SALE Message] = CStr(ADOprimaryrs![AR ORDER Message])
    End If
    If Not (IsNull(Trim(ADOprimaryrs![AR ORDER Apply To Number]))) Then
    '    ![AR SALE Apply To Number] = " "
    'Else
        ![AR SALE Apply To Number] = CStr(ADOprimaryrs![AR ORDER Apply To Number])
    End If
    If Not (IsNull(Trim(ADOprimaryrs![AR ORDER Check Number]))) Then
    '    ![AR SALE Check Number] = " "
    'Else
        ![AR SALE Check Number] = CStr(ADOprimaryrs![AR ORDER Check Number])
    End If
    If Not (IsNull(Trim(ADOprimaryrs![AR ORDER Check Acct ID]))) Then
    '    ![AR SALE Check Acct ID] = " "
    'Else
        ![AR SALE Check Acct ID] = CStr(ADOprimaryrs![AR ORDER Check Acct ID])
    End If
    ![AR SALE Check Date] = ADOprimaryrs![AR ORDER Check Date]
    ![AR SALE Quote Document #] = ADOprimaryrs![AR ORDER Quote Document #]
    If Not (IsNull(Trim(ADOprimaryrs![AR ORDER Notes]))) Then
    '    ![AR SALE Notes] = " "
    'Else
        ![AR SALE Notes] = CStr(ADOprimaryrs![AR ORDER Notes])
    End If
    If Not (IsNull(Trim(ADOprimaryrs![AR ORDER Department]))) Then
    '    ![AR SALE Department] = " "
    'Else
        ![AR SALE Department] = CStr(ADOprimaryrs![AR ORDER Department])
    End If
    ![AR SALE Taxable Subtotal] = ADOprimaryrs![AR ORDER Taxable Subtotal]
    ![AR SALE Tax Freight] = ADOprimaryrs![AR ORDER Tax Freight]
    ![AR SALE Recurring YN] = False
    ![AR SALE Cleared] = False
    ![AR SALE Recur Type] = "Never"
  .Update
  lSalesID = ![AR SALE Document #]
  End With '----------------
  
  ADOprimaryrs![AR ORDER Invoiced Document] = "[" & sExtDocument & "]"
  ADOprimaryrs.Update
  
  'get the item cost price
  Dim GetCost As ADODB.Recordset
  Set GetCost = New ADODB.Recordset
  GetCost.Open "Select [INV ITEM Id],[INV ITEM New Price] from [INV ITEMS]", db, adOpenKeyset, adLockReadOnly, adCmdText
  
  Set rsSalesDetail = New ADODB.Recordset
  rsSalesDetail.Open "SELECT * FROM [AR Sales Detail] WHERE [AR SALED Document #]=" & lSalesID, DataB, adOpenKeyset, adLockOptimistic, adCmdText
  
  With rsSalesDetail
  
  Do While Not rsOrderDetail.EOF
    rsSalesDetail.AddNew
      ![AR SALED Document #] = lSalesID
      ![AR SALED Qty] = rsOrderDetail![AR ORDERD Qty To Invoice]
      ![AR SALED Date] = rsOrderDetail![AR ORDERD Date]
      ![AR SALED Item ID] = CStr(rsOrderDetail![AR ORDERD Item ID])
      ![AR SALED Units] = CStr(rsOrderDetail![AR ORDERD Units])
      ![AR SALED Discount %] = rsOrderDetail![AR ORDERD Discount %]
      ![AR SALED Tax] = rsOrderDetail![AR ORDERD Tax]
      ![AR SALED Tax Rate] = rsOrderDetail![AR ORDERD Tax Rate]
      ![AR SALED Description] = CStr(rsOrderDetail![AR ORDERD Description])
      If Not (IsNull(Trim(rsOrderDetail![AR ORDERD Row Type]))) Then
      '    ![AR SALED Row Type] = " "
      'Else
          ![AR SALED Row Type] = CStr(rsOrderDetail![AR ORDERD Row Type])
      End If
      GetCost.MoveFirst '
      GetCost.Find "[INV ITEM Id]='" & rsOrderDetail![AR ORDERD Item ID] & "'"
      ![AR SALED Item Cost] = GetCost![INV ITEM New Price]
      ''debug.print GetCost![INV ITEM New Price] & "      " & GetCost![INV ITEM Id] & "        " & rsOrderDetail![AR ORDERD Item ID]
      ![AR SALED Unit Price] = rsOrderDetail![AR ORDERD Unit Price]
      ![AR SALED Item Total] = rsOrderDetail![AR ORDERD Item Total]
      ![AR SALED Posting Account] = CStr(rsOrderDetail![AR ORDERD Posting Account])
      If Not (IsNull(Trim(rsOrderDetail![AR ORDERD Department]))) Then
      '    ![AR SALED Department] = " "
      'Else
          ![AR SALED Department] = CStr(rsOrderDetail![AR ORDERD Department])
      End If
      If Not (IsNull(Trim(rsOrderDetail![AR ORDERD Project]))) Then
      '    ![AR SALED Project] = " "
      'Else
          ![AR SALED Project] = CStr(rsOrderDetail![AR ORDERD Project])
      End If
      If Not (IsNull(Trim(rsOrderDetail![AR ORDERD Account Type]))) Then
      '    ![AR SALED Account Type] = " "
      'Else
          ![AR SALED Account Type] = CStr(rsOrderDetail![AR ORDERD Account Type])
      End If
    .Update
    rsOrderDetail.MoveNext
  Loop
     
  End With '--------------
  Dim SQLstatement As String
      SQLstatement = "INSERT INTO [AR Order Tracking] ([Order ID],[Invoice ID])"
      SQLstatement = SQLstatement & " VALUES (" & CInt(txtFields(2)) & "," & lSalesID & ")"
      DataB.Execute SQLstatement
  
  'Set rsOrderTracking = New ADODB.Recordset
  'rsOrderTracking.Open ("SELECT  [Order ID],[Invoice ID] [AR Order Tracking]"), DataB, adOpenStatic, adLockOptimistic
        
  'rsOrderTracking.AddNew
  '  Dim AddInteger2  As Integer
  '  AddInteger2 = txtFields(2)    'they need string to update --- need to find other way
  '  rsOrderTracking![Order ID] = AddInteger2
  '  rsOrderTracking![Invoice ID] = lSalesID
  'rsOrderTracking.Update
  
  'Write out sales number
  '--------------->>>previous access coding
  'DoCmd.SetWarnings False
  'DoCmd.OpenQuery "Delete Bad Sales"
  'DoCmd.SetWarnings True
  
  Dim SProc As String
  SProc = "DELETE DISTINCTROW [AR Sales].*, [AR Sales].[AR SALE Ext Document #], [AR Sales].[AR SALE Customer ID] FROM [AR Sales] WHERE ((([AR Sales].[AR SALE Ext Document #]) Is Null Or ([AR Sales].[AR SALE Ext Document #])="")) OR ((([AR Sales].[AR SALE Customer ID]) Is Null Or ([AR Sales].[AR SALE Customer ID])=""))"
  '-----------execute command below as good as query "Delete Bad Sales"
  db.Execute SProc

  'Call RedoNumbers

  ShowStatus False
  iSuccess = False
  db.BeginTrans
  iResponse = MsgBox("Invoice created, do you want to post it now?" & vbCr & "Posting the Invoice will take a few seconds", vbYesNo, "Posting")
  If iResponse = vbNo Then GoTo byPassInvoiceCreation
  
  If Datavalidate = False Then GoTo byPassInvoiceCreation
  
  ShowStatus True
  
  'Now post the invoice
  iSuccess = PostInvoice(lSalesID, True, db)
  
byPassInvoiceCreation:
  If iSuccess = True Then
    db.CommitTrans
    MsgBox "Invoice created and posted successfully."
    'Now mark this invoice as posted
    'rsSales.Index = "Ext Document #"
    rsSales.MoveFirst
    rsSales.Find "[AR SALE Ext Document #]='[" & sExtDocument & "]'"
    If rsSales.EOF Then
    Else
      'rsSales.Edit
        rsSales("AR SALE Posted YN") = True
        rsSales.Update
    End If
  Else
    db.RollbackTrans
    MsgBox "Invoice created but not posted."
    ShowStatus False
    'chkInvoiced.Value = 1
    'adoPrimaryRS![AR ORDER Invoiced] = True
    'cmdUpdate_Click
    'Exit Sub
  End If
          
  'chkInvoiced.Value = 1
  ADOprimaryrs![AR ORDER Invoiced] = True
  ADOprimaryrs![AR ORDER Status] = "Posted"
  cmdUpdate_Click
  ButtEnabled False
  GetTextColor Me
  
  Dim iBackOrdered As Integer
  iBackOrdered = False
    ShowStatus True
          
  'Check shipping qty vs order quantity, if there are products left to ship then ask about creating a backorder
  rsOrderDetail.MoveFirst
  Do While Not rsOrderDetail.EOF
    If rsOrderDetail("AR ORDERD Qty") > rsOrderDetail("AR ORDERD Qty To Invoice") Then
      iBackOrdered = True
      Exit Do
    End If
    rsOrderDetail.MoveNext
  Loop
          
  If iBackOrdered = True Then
    iResponse = MsgBox("Would you like to create a backorder for quantities not invoiced?", vbYesNo, "Information")
    If iResponse = vbYes Then
      'Create the backorder
      Dim NewOrderNo As String
      NewOrderNo = CreateBackOrder(![AR ORDER Document #], False)
      If NewOrderNo <> "" Then
        MsgBox "Backorder created successfully.", , "Information"
        ADOprimaryrs![AR ORDER BackOrder Document] = NewOrderNo
        ADOprimaryrs.Update
        '.Requery
        'DoCmd.GoToRecord acForm, Me.Name, acLast
        '-------------->>>>alter the CalcTotals
        'Call CalcTotals
      Else
        MsgBox "An error occurred creating the backorder.", , "Information"
      End If
    Else
    End If
    ButtEnabled False
    ShowStatus False
  End If
End With
Out_Of_Here:

  rsSales.Close
  rsSalesDetail.Close
  'rsOrderTracking.Close
  rsOrderDetail.Close
  DataB.Close
  
  Set DataB = Nothing
  'Set rsTemp = Nothing
  Set rsSales = Nothing
  Set rsSalesDetail = Nothing
  'Set rsOrderTracking = Nothing
  Set rsOrderDetail = Nothing
ShowStatus False
  Exit Sub

cmdCreateInvoice_Error:
  Call ErrorLog("Sales Order", "cmdCreateInvoice_Click", Now, Err.Number, Err.Description, True, db)
  Resume Next

End Sub

Private Sub RedoNumbers()
  'Dim SProc As String
  
  'DoCmd.OpenQuery "qryDeleteOrderNumbers"
  'SProc = "DELETE DISTINCTROW [Order Numbers].* FROM [Order Numbers]"
  'db.Execute SProc

  'DoCmd.OpenQuery "qryRedoOrderNumbers"
  'SProc = "INSERT INTO [Order Numbers] ( [Document ID] ) SELECT DISTINCTROW [AR Order].[AR ORDER Ext Document #] FROM [AR Order]"
'  db.Execute SProc

'-----------previous access coding
 ' 'On Error GoTo RedoOrderNumbers_Error

  'Dim rsNumber As ADODB.Recordset
  'Dim rsSales As ADODB.Recordset

  'xxx 1/7/97  7.2b
  'DBEngine.Workspaces(0).BeginTrans
  
  'DoCmd.SetWarnings False
  'DoCmd.OpenQuery "qryDeleteOrderNumbers"
  
  'xxx 3/26/97 7.3
  'DoCmd.OpenQuery "qryRedoOrderNumbers"
  'DoCmd.SetWarnings True

'  Set rsNumber = db2.OpenRecordset("Invoice Numbers")
'  Set rsSales = db2.OpenRecordset("AR Sales")

'  'On Error Resume Next
'  rsSales.MoveFirst
'  If Err = 0 Then
'    Do While Not rsSales.EOF
'      rsNumber.AddNew
'        rsNumber("Document ID") = rsSales("AR SALE Ext Document #")
'      rsNumber.UPDATE
'      rsSales.MoveNext
'    Loop
'  End If
  
  'xxx 1/7/97  7.2b
'  DBEngine.Workspaces(0).CommitTrans

'  Exit Sub
'RedoOrderNumbers_Error:
'  Call ErrorLog("Order Transactions", "RedoSalesNumbers", Now, Err.Number, Err.Description, True, db)
'  Resume Next
End Sub

Private Sub cmdDate_Click(Index As Integer)
Select Case Index
Case 0
    Menu_Calendar.WhoCallMe True, 1002
    If txtFields(7).Text <> "" Then ADOprimaryrs![AR ORDER Check Date] = txtFields(7).Text
Case 1
    Dim iResponse As Integer
    iResponse = MsgBox("Due Date were set automaticly... Are sure you want to change it?", vbYesNo, "Due Date")
    If iResponse = vbNo Then Exit Sub
    Menu_Calendar.WhoCallMe True, 1001
    If txtFields(6).Text <> "" Then ADOprimaryrs![AR ORDER Due Date] = txtFields(6).Text
    'txtfields(6).SetFocus
Case 20
    Menu_Calendar.WhoCallMe True, 1000
    If txtFields(20).Text <> "" Then ADOprimaryrs![AR ORDER Ship Date] = txtFields(20).Text
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


'Private Sub cbpurchase_Change(Index As Integer)
'Select Case Index
'Case 3
'  If cbPurchase(3) = "" Then
'    txtFields(29) = 0
'  Else
'    CalcSalesTaxPercent
'  End If
'End Select
'End Sub

Private Sub cmdUpdatedua_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    cmdUpdatedua_Select (Index)
Else
    'Me.PopupMenu mnuCombo
End If
End Sub

Private Sub Command1_Click()
ShowStatus True
    If CheckDocument("select * from [AR Order] WHERE [AR ORDER Document Type]='" & DocType & "' AND [AR ORDER Ext Document #]='" & txtFields(0) & "'", db, False) = False Then
        Dim Response As Integer
            Response = MsgBox("Search found, Would you like to see it?", vbYesNo, "Information")
            If Response = vbYes Then
                ShowStatus True
                ClearDatasource
                RSstatement = "SHAPE {select * from [AR Order] WHERE [AR ORDER Document Type]='" & DocType & "' AND [AR ORDER Ext Document #]='" & txtFields(0) & "'} AS ParentCMD APPEND ({select * from [AR Order Detail] } AS ChildCMD RELATE [AR ORDER Document #] TO [AR ORDERD Document #]) AS ChildCMD"
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
DocType = "Order"
  Set db = New ADODB.Connection
  db.CursorLocation = adUseClient
  db.Provider = "MSDataShape"
  db.Open "Data " & gblADOProvider
     
     CustID0 = ""
     
     Me.Height = 6600
     Me.Width = 11475
     Me.Top = 0
     Me.Left = 0
    
    RSstatement = "SHAPE {select * from [AR Order] WHERE [AR ORDER Document Type]='" & DocType & "' AND [AR ORDER Invoiced] = False} AS ParentCMD APPEND ({select * from [AR Order Detail] } AS ChildCMD RELATE [AR ORDER Document #] TO [AR ORDERD Document #]) AS ChildCMD"
    OpenDB RSstatement
    'set the datagrid button to true
    'grdDataGrid.Columns(0).Button = True
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
'[AR ORDER Document #],[AR ORDER Document Type],[AR ORDER Taxable Subtotal],[AR ORDER Customer ID],
'[AR ORDER Billing Customer],[AR ORDER Billing Address 1],[AR ORDER Billing Address 2],
'[AR ORDER Billing City],[AR ORDER Billing State],[AR ORDER Billing Postal],[AR ORDER Billing Country],
'[AR ORDER Billing Phone],[AR ORDER Billing Fax],[AR ORDER Shipping ID],[AR ORDER Shipping Customer],
'[AR ORDER Shipping Address 1],[AR ORDER Shipping Address 2],[AR ORDER Shipping City],[AR ORDER Shipping State],
'[AR ORDER Shipping Postal],[AR ORDER Shipping Country],[AR ORDER Shipping Phone],[AR ORDER Shipping Fax],
'[AR ORDER Ext Document #],[AR ORDER PO ID],[AR ORDER Date],[AR ORDER Salesperson],
'[AR ORDER Ship Date],[AR ORDER Due Date],[AR ORDER Check Acct ID],[AR ORDER Check Number],[AR ORDER Check Date],
'[AR ORDER Payment Terms],[AR ORDER Payment Method],[AR ORDER Tax Group],[AR ORDER Recur Type],[AR ORDER Shipping Method],
'[AR ORDER Subtotal],[AR ORDER Discount Percent],[AR ORDER Discount Amount],[AR ORDER Tax Percent],[AR ORDER Sales Tax],
'[AR ORDER Tax Freight],[AR ORDER Freight],[AR ORDER Total],[AR ORDER Amount Paid],[AR ORDER Balance Due]
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
        ADOprimaryrs.Find "[AR ORDER Ext Document #]='" & DocType & AppLoginName & "'"
      If Not .EOF Then
        ADOprimaryrs![AR ORDER Ext Document #] = AppLoginName & Format(Now, "MMdd") & Right(Format(![AR ORDER Document #], "0000"), 4)
        ADOprimaryrs![AR ORDER Status] = "Open"
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
  
  If CheckNewDB(ADOprimaryrs, "Order Entry") = True Then
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
      If Shift = vbCtrlMask And txtFields(36).Enabled = True Then
        ADOprimaryrs![AR ORDER Amount Paid] = "$0.00" 'AR ORDER Amount Paid
        txtFields_LostFocus 36
        calculateALL
      Else
          cmdLast_Click
      End If
    Case vbKeyHome
      If Shift = vbCtrlMask And txtFields(36).Enabled = True Then
        ADOprimaryrs![AR ORDER Amount Paid] = txtFields(30)
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
    Set frm_AR_Order_Entry = Nothing
Exit Sub
FormErr:
  MsgBox Err.Description
  ShowStatus False
End Sub

Private Sub adoPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
Dim Response As Integer
ShowStatus True
  If ADOprimaryrs.BOF Or ADOprimaryrs.EOF Then GoTo JumpIf
  If ADOprimaryrs![AR ORDER Invoiced] = True Then
     ButtEnabled False
  Else
     ButtEnabled True
     If IsNull(ADOprimaryrs![AR ORDER Status]) Or ADOprimaryrs![AR ORDER Status] = "Open" Then
        cmdApprove.Picture = fMainForm.imlIcons.ListImages("Locked").Picture
     Else
        cmdApprove.Picture = fMainForm.imlIcons.ListImages("Approved").Picture
     End If
          
     If ADOprimaryrs![AR ORDER Amount Paid] > 0 Then
        txtFields(7).Enabled = True
        txtFields(34).Enabled = True
        cmdDate(0).Enabled = True
        txtFields(35).Enabled = True
        cmdbankAccount.Enabled = True
    Else
        txtFields(7).Enabled = False
        txtFields(34).Enabled = False
        cmdDate(0).Enabled = False
        txtFields(35).Enabled = False
        cmdbankAccount.Enabled = False
     End If
  End If
   If mbAddNewFlag = False Then
        If IsNull(ADOprimaryrs![AR ORDER Customer ID]) Then
        Else
            CustomerData Me, db, ADOprimaryrs![AR ORDER Customer ID], False
        End If
        txtFieldsCust(0).Locked = True
        txtFieldsShip(0).Locked = True
        txtFields(36).Enabled = False
   Else
        lblmail.Visible = False
        lblweb.Visible = False
        txtFieldsCust(0).Locked = False
        txtFieldsShip(0).Locked = False
        txtFields(36).Enabled = True
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
        cmdShipBack.Visible = SetEnabled
        If mbAddNewFlag = False Then
            cmdCreateInvoice.Enabled = SetEnabled   'cmdPrint
            cmdPrint.Enabled = True
            cmdApprove.Enabled = SetEnabled
        Else
            cmdPrint.Enabled = False
            cmdApprove.Enabled = False
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
         'If txtFieldsCust(3) = "" Then
         '   ![AR ORDER Billing Address 2] = " "
         'End If
         'If txtFieldsShip(3) = "" Then
         '   ![AR ORDER Shipping Address 2] = " "
         'End If
         cmdUpdate_Click
         '.MovePrevious
         'grdDataGrid.HoldFields
         'grdDataGrid.ReBind
         'grdDataGrid.RefreshLoadRS SQLstatement
         'RSstatement = "SHAPE {select * from [AR Order] WHERE [AR ORDER Document Type]='" & DocType & "' AND [AR ORDER Ext Document #]='" & DocType  & AppLoginName & "' Order by [AR ORDER Document #]} AS ParentCMD APPEND ({select * from [AR Order Detail] } AS ChildCMD RELATE [AR ORDER Document #] TO [AR ORDERD Document #]) AS ChildCMD"
         ClearDatasource
         RSstatement = "SHAPE {select * from [AR Order] WHERE [AR ORDER Document Type]='" & DocType & "' AND [AR ORDER Invoiced]=False} AS ParentCMD APPEND ({select * from [AR Order Detail] } AS ChildCMD RELATE [AR ORDER Document #] TO [AR ORDERD Document #]) AS ChildCMD"
         OpenDB RSstatement, True
         'ADOprimaryrs.Find "[AR ORDER Ext Document #]='" & "TempOrderNo" & AppLoginName & "'"
         'If ADOprimaryrs.EOF Then
         'Else
         '   txtFields(12).SetFocus
            'txtFields(12) = AppLoginName & Format(Now, "MMdd") & Format(txtFields(2), "000")
         'End If
         NewLoad = False
         
     End With
     cmdAdd.Caption = "&Add"
     cmdLookupShip.Visible = False
     SetButtons True
     cmdCreateInvoice.Enabled = True
     cmdPrint.Enabled = True
Else
    If DocType = "Backorder" Then
        MsgBox "Can't create new data on BackOrder", vbInformation, "Information"
        Exit Sub
    End If
    mbAddNewFlag = True
    cmdCreateInvoice.Enabled = False
    cmdPrint.Enabled = False
  With ADOprimaryrs
    If Not (.BOF Or .EOF) Then
      mvBookMark = .Bookmark
    End If
     NewLoad = True
     cmdLookupShip.Visible = True
    .AddNew
    txtFields(12) = DocType & AppLoginName
    txtFields(3) = DocType
    txtFields(59) = "$0.00"
    txtFields(4) = FormatDate(Now)
    'txtFields(7) = txtFields(4)
    SetDueDate
    txtSalesPerson(0) = AppLoginName
    lblStatus.Caption = "Add record"
        Dim i As Integer
        If mbAddNewFlag = True Then
           For i = 24 To 30
             Select Case i
                Case 24, 25, 26, 27, 30, 33, 36
                    txtFields(i) = "$0.00"
                Case 28, 29
                    txtFields(i) = "00.00"
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
If txtFields(4) = "" Then Exit Sub
If mbAddNewFlag = True Then
    DueDateDay db, cbPurchase(5), txtFields(4), txtFields(6), True
Else
    DueDateDay db, cbPurchase(5), txtFields(4), txtFields(6)
End If
End Sub

Private Sub cmdDelete_Click()
Dim DocNo As String 'picStatBox
'Dim DelStatus As String

DocNo = txtFields(12).Text

'     DelStatus = DataDelete(ADOprimaryrs, Me, True)
     
'     If DelStatus = False Then
'        MsgBox "An error occured while attempting to delete " & DocNo & ", closing the " & DocType
'        Unload Me
'     Else
''        If picStatBox.Enabled = False Then
        ShowStatus True
        ClearDatasource
        db.Execute "DELETE FROM [AR Order] WHERE [" & txtFields(12).DataField & "]='" & DocNo & "'"
'        MsgBox lblTop & "[" & DocNo & "] has been deleted. Refreshing the database process will take place after this.", vbInformation, "Information"
        'ADOprimaryrs.Requery
            MsgBox lblTop & "[" & DocNo & "] has been deleted." & vbCr & _
            "Opening Unposted " & DocType & " Form", vbInformation, "Information"
'            cmdUnPostedDoc_Click
        RSstatement = "SHAPE {select * from [AR Order] WHERE [AR ORDER Document Type]='" & DocType & "' AND [AR ORDER Invoiced] = False} AS ParentCMD APPEND ({select * from [AR Order Detail] } AS ChildCMD RELATE [AR ORDER Document #] TO [AR ORDERD Document #]) AS ChildCMD"
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
  cmdLookupShip.Visible = False
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
  Case 2, 3, 6
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
On Error Resume Next
  If grdDataGrid.Columns(2).Text = "" Then grdDataGrid.Columns(2).Text = 0
  If grdDataGrid.Columns(3).Text = "" Then grdDataGrid.Columns(3).Text = 0
  If grdDataGrid.Columns(6).Text = "" Then grdDataGrid.Columns(6).Text = 0
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
    NewRowForDataGrid ADOprimaryrs, grdDataGrid, "AR ORDER Date", txtFields(4).Text
    grdOnAddNew = False
    NewLoad = False
End Sub

Private Sub grdDataGrid_ButtonClick(ByVal ColIndex As Integer)
If mbAddNewFlag = True Then Exit Sub
On Error GoTo Error_ButtClick
If grdDataGrid.Columns(0) <> "" Then grdOnAddNew = False
Select Case ColIndex
Case 0   'select the item from the ITEM_ID
    INV_ITEM
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
   AllLookup.GetWhichTable 1005, "SELECT [INV ITEM Id], [INV ITEM Description]," & _
   "[INV ITEM Unit],[INV ITEM Price], [INV ITEM Sales Account], [INV ITEM Qty On Hand], " & _
   "[INV ITEM Qty On Order], [INV ITEM Taxable YN],[INV ITEM Last Cost] FROM [INV Items] " & _
   "WHERE [INV ITEM Inactive YN]=FALSE ", "Product", _
   "Item ID//Item Description//Unit//Price//Sales Account//Qty On Hand//Qty On Order//Taxable//Cost", db
   'AllLookup.Show vbModal

End Sub

Private Sub COA_grdDataGrid_Butt()
   AllLookup.GetWhichTable 1002, "Select [GL COA Account No],[GL COA Account Name]," & _
   "[GL COA Asset Type] From [GL Chart Of Accounts] ", "GL Accounts", _
   "Account No//Account Type//Account Type", db
   'AllLookup.Show vbModal
   
End Sub

Private Sub Proj_Projects()
   AllLookup.GetWhichTable 1003, "Select [PROJ ID],[PROJ Name]," & _
   "[PROJ Description] From [PROJ Projects] ", "Project", _
   "Project ID//Project Name//Description", db
   'AllLookup.Show vbModal
   
End Sub

Private Sub CalculateTable()
Dim i As Integer

    'get the total value for the selected rod
    grdDataGrid.Columns(8).Text = grdDataGrid.Columns(3).Value * grdDataGrid.Columns(5).Text
    grdDataGrid.Columns(8).Text = grdDataGrid.Columns(8).Value - (grdDataGrid.Columns(8).Value * grdDataGrid.Columns(6).Value)
        
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

  txtFields(29) = TaxPercent#

  Exit Sub
CalcSalesTaxPercent_Error:
  Call ErrorLog(DocType & " Transactions", "CalcSalesTaxPercent", Now, Err.Number, Err.Description, True, db)
  Resume Next
  
End Sub


Private Sub CalcTotals()
If NewLoad = True Then Exit Sub
Dim Calcrs As Recordset
    'If Not frTotal.Enabled Then Exit Sub
    If ADOprimaryrs.EOF = True Or ADOprimaryrs.BOF = True Then Exit Sub
    If mbAddNewFlag = True Then Exit Sub
    grdDataGrid.Row = grdDataGrid.Row
    'SUM all the data in AR ORDERD DETAIL that match the order document id
    Set Calcrs = New ADODB.Recordset
    Calcrs.Open "SELECT [AR ORDERD Item Total] FROM [AR Order Detail] WHERE [AR ORDERD Document #]=" & txtFields(2), db, adOpenStatic, adLockOptimistic
    
    Dim ItemTotal As Double
    With Calcrs
        If .RecordCount > 0 Then
            .MoveFirst
            While Not .EOF
                ItemTotal = ItemTotal + Calcrs![AR ORDERD Item Total]
                .MoveNext
            Wend
        Else
            ItemTotal = 0
        End If
    End With
    txtFields(24) = FormatCurr(CCur(ItemTotal))
    If txtFields(24).Enabled = True Then ADOprimaryrs![AR ORDER Subtotal] = txtFields(24).Text
    calculateALL
    ADOprimaryrs.Update
    
    Calcrs.Close
    Set Calcrs = Nothing
End Sub

Private Sub calculateALL()

  'Calculate totals for this invoice
        If txtFields(24) = "" Then txtFields(24) = "$0.00"
        If CCur(txtFields(24)) <= 0 Then
            txtFields(25) = "$0.00"
            txtFields(26) = "$0.00"
            txtFields(30) = "$0.00"
            txtFields(33) = "$0.00"
            txtFields(36) = "$0.00"
            Exit Sub
        End If
  ''On Error GoTo CalcTotals_Error
  'Calculate discount
        If Val(txtFields(28)) <> 0 Then
            txtFields(25) = txtFields(24) * (txtFields(28) / 100)
            txtFields(25) = FormatCurr(txtFields(25))
        Else
            txtFields(25) = "$0.00"
        End If
        If Val(txtFields(29)) <> 0 Then
            txtFields(26) = txtFields(24) * (txtFields(29) / 100)
            txtFields(26) = FormatCurr(txtFields(26))
        Else
            txtFields(26) = "$0.00"
        End If

  If txtFields(26) = "" Then
    txtFields(26) = "$0.00"
  End If
  
  'Calculate Total
  txtFields(30) = txtFields(24) - txtFields(25) + txtFields(26) + txtFields(27)

  txtFields(30) = FormatCurr(txtFields(30))
  'txtfields(1) = txtfields(30) - txtfields(5)
  
  If txtFields(36) = "." Or txtFields(36) = "" Then txtFields(36) = "$0.00"
  txtFields(33) = txtFields(30) - txtFields(36)
  txtFields(33) = FormatCurr(txtFields(33))
  
  txtFields(59) = txtFields(24)
  If txtFields(23) = "Yes" Then
    txtFields(59) = CCur(txtFields(59)) + CCur(txtFields(27))

  End If
  txtFields(59) = FormatCurr(txtFields(59) - txtFields(25))
  
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
        If txtFields(36) <> "" And Not mbAddNewFlag Then
            If Trim(Ctrl.Text) = "" And txtFields(36) <> "$0.00" Then
                MsgBox "There is an empty data in " & lblfields(Ctrl.Index).Caption, vbInformation, "Empty Data"
                CheckEmpty = False
                GoTo Out_Of_Here
            End If
        End If
      Case 14
        If Trim(Ctrl.Text) = "" And LCase(Ctrl.Name) = "txtfields" Then
            MsgBox "There is an empty data in " & lblLabels(Ctrl.Index).Caption, vbInformation, "Empty Data"
            CheckEmpty = False
            GoTo Out_Of_Here
        End If
      End Select
    End If
   End If
    
   If TypeOf Ctrl Is ComboBox Then
      If Ctrl.Text = "" Then
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


Private Sub optOrderType_Click(Index As Integer)
If mbAddNewFlag = True Then
    MsgBox "Can't change during adding new data"
    Exit Sub
End If '

Dim SetEnabled As Boolean
    Select Case Index
    Case 0 'Order
        ShowStatus True
        DocType = "Order"
        lblTop.Caption = "Sales/Delivery " & DocType
        ClearDatasource
        RSstatement = "SHAPE {select * from [AR Order] WHERE [AR ORDER Document Type]='Order' Order by [AR ORDER Document #]} AS ParentCMD APPEND ({select * from [AR Order Detail] } AS ChildCMD RELATE [AR ORDER Document #] TO [AR ORDERD Document #]) AS ChildCMD"
        OpenDB RSstatement
        ShowStatus True
        SetEnabled = True
        cmdInvoiceBackOrder.Visible = False
        optOrderType(0).ForeColor = &HFF8080
        optOrderType(1).ForeColor = &H0&
    Case 1 'Backorder
        ShowStatus True
        DocType = "Backorder"
        lblTop.Caption = "Sales/Delivery " & DocType
        ClearDatasource
        RSstatement = "SHAPE {select * from [AR Order] WHERE [AR ORDER Document Type]='Backorder' Order by [AR ORDER Document #]} AS ParentCMD APPEND ({select * from [AR Order Detail] } AS ChildCMD RELATE [AR ORDER Document #] TO [AR ORDERD Document #]) AS ChildCMD"
        OpenDB RSstatement
        ShowStatus True
        SetEnabled = False
        cmdInvoiceBackOrder.Visible = True
        optOrderType(1).ForeColor = &HFF8080
        optOrderType(0).ForeColor = &H0&
    End Select
        'frFirst.Enabled = SetEnabled
        'frSecond.Enabled = SetEnabled
        cmdAdd.Enabled = SetEnabled
        'cmdUpdate.Enabled = SetEnabled
        cmdDelete.Enabled = SetEnabled
        'cmdRefresh.Enabled = SetEnabled
        'cmdClose.Enabled = SetEnabled
        cmdCreateInvoice.Visible = SetEnabled
        'GetTextColor Me
        ShowStatus False
    
    picStatBox.Enabled = True
    lblStatus.BackColor = vbWhite
    ProcessDoneMusic "Done"

End Sub

'Private Sub txtFields_Change(Index As Integer)
'Select Case Index
'Case 2
    'If txtfields(12) = "TempOrderNo" & AppLoginName Then
    '    ADOprimaryrs![AR ORDER Ext Document #] = AppLoginName & Format(Now, "MMdd") & Format(txtfields(2), "000")
        'txtFields(12) = AppLoginName & Format(Now, "MMdd") & Format(txtFields(2), "000")
    '    ADOprimaryrs.Update
    'End If
'End Select

'End Sub

Private Sub txtfields_GotFocus(Index As Integer)
    TxtGotFocus txtFields(Index)
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
    If IsNull(ADOprimaryrs![AR ORDER Status]) Or ADOprimaryrs![AR ORDER Status] <> "Open" Then
        ValidatePower txtFields(12).Text, "Edit", DocType, db
        ADOprimaryrs![AR ORDER Status] = "Open"
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
    db.Execute "UPDATE [AR Customer] SET [AR CUST Discount %] = " & CDbl(txtFields(28)) & " WHERE [AR CUST Customer ID] = '" & txtFieldsCust(0).Text & "'"
    CalcTotals
    GetTextColor Me
Case 34
    'CheckNumberCHQ
    If Trim(txtFields(35).Text) = "" Then
        'MsgBox "Please select Bank Account first before writing a check number", vbCritical, "Information"
        txtFields(34) = " "
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
    'If CheckNumberCHQ("check", db, txtfields(35).Text, txtfields(34).Text) = "Found" Then
    '    Dim Response As Integer
    '    Response = MsgBox("This check has already been used." & vbCr & _
    '    "Would you like computer to search for a valid check number?", vbYesNo, "Information")
    '    If Response = vbYes Then
    '    Else
    '        txtfields(34) = " "
    '    End If
    'End If
Case 35
    If BankAcct35 = txtFields(35) Then Exit Sub
    If Trim(txtFields(35)) = "" Then Exit Sub
    If IsNumeric(txtFields(35).Text) Then
        CheckDocument "SELECT [GL COA Account No] FROM [GL Chart Of Accounts] WHERE [GL COA Account No]='" & txtFields(35).Text & "'", db, False, txtFields(35)
    Else
        MsgBox "Only numeric character is accepted", vbInformation, "Information"
        txtFields(35) = " "
    End If
    BankAcct35 = txtFields(35)
Case 36
    calculateALL
    If CCur(txtFields(36)) > 0 And CCur(txtFields(33)) >= 0 Then
        txtFields(7).Enabled = True
        txtFields(34).Enabled = True
        cmdDate(0).Enabled = True
        cmdbankAccount.Enabled = True
        txtFields(35).Enabled = True
    ElseIf CCur(txtFields(33)) < 0 Then
        MsgBox "Balance must not be less than Zero", vbInformation, "Information"
        txtFields(36) = txtFields(24)
        txtFields(33) = "$0.00"
        txtFields(7).Enabled = True
        txtFields(34).Enabled = True
        cmdDate(0).Enabled = True
        cmdbankAccount.Enabled = True
        txtFields(35).Enabled = True
    Else
        'txtfields(7).Text = " "
        txtFields(34).Text = " "
        txtFields(7).Enabled = False
        txtFields(34).Enabled = False
        cmdDate(0).Enabled = False
        cmdbankAccount.Enabled = False
        txtFields(35).Enabled = False
    End If
    GetTextColor Me
End Select
If Trim(txtFields(Index)) = "" Then
    txtFields(Index) = " "
    Exit Sub
End If
End Sub

Private Sub txtFieldsCust_GotFocus(Index As Integer)
    TxtGotFocus txtFieldsCust(Index)
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

Private Sub txtFieldsCust_LostFocus(Index As Integer)
If txtFieldsCust(0).Locked = True Then Exit Sub
Dim i As Integer
If txtFieldsCust(0).Text = "" Then
    For i = 0 To txtFieldsCust.UBound
        txtFieldsCust(i).Text = ""
    Next
ElseIf CustID0 <> txtFieldsCust(0).Text Then
    If Index = 0 Then
        'If CheckDocument("SELECT [AR CUST Customer ID] FROM [AR Customer] WHERE [AR CUST Customer ID]='" & txtFieldsCust(0).Text & "'", db, False, txtFieldsCust(0), "Customer ID") = False Then
        'found
            CustomerID "Select [AR CUST Customer ID],[AR CUST Name],[AR CUST Address 1]," & _
           "[AR CUST Address 2],[AR CUST City],[AR CUST State],[AR CUST Postal]," & _
           "[AR CUST Country],[AR CUST Phone],[AR CUST Fax],[AR CUST SalesPerson] " & _
           "FROM [AR Customer] WHERE [AR CUST Customer ID]='" & txtFieldsCust(0).Text & "'", db, Me
    'MsgBox txtFieldsCust(0).Text
           If txtFieldsCust(0).Text <> "" Then
                CustomerData Me, db, txtFieldsCust(0).Text, True
                cbPurchase_LostFocus 1
                cbPurchase_LostFocus 5
            End If
        'End If
    End If
End If
CustID0 = txtFieldsCust(0).Text
End Sub

Private Sub txtFieldsShip_GotFocus(Index As Integer)
    TxtGotFocus txtFieldsShip(Index)
End Sub

Private Sub txtFieldsShip_LostFocus(Index As Integer)
If txtFieldsShip(0).Locked = True Then Exit Sub
Dim i As Integer
If Index = 0 Then
    'If txtFieldsCust(0).Text = "" Or ShipID0 = txtFieldsShip(0).Text Or txtFieldsShip(0).Text = "" Then
    If txtFieldsCust(0).Text = "" Or txtFieldsShip(0).Text = "" Then
        For i = 0 To txtFieldsShip.UBound
        txtFieldsShip(i).Text = ""
        Next
    ElseIf ShipID0 <> txtFieldsShip(0).Text Then
        shipToID db, Me
        ShipID0 = txtFieldsShip(0)
    End If
End If
End Sub

Private Sub txtSalesPerson_GotFocus(Index As Integer)
    TxtGotFocus txtSalesPerson(Index)
End Sub
